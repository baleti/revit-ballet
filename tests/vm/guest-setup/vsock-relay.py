"""
vsock to SSH relay - uses raw Windows sockets throughout
Avoids Python socket.fromfd() which doesn't work with vsock on Windows
"""
import ctypes
import threading
import sys
import argparse
import os

# Windows socket constants
AF_VSOCK = 40
AF_INET = 2
SOCK_STREAM = 1
INVALID_SOCKET = -1
SOCKET_ERROR = -1

# Configuration
VSOCK_PORT = 3389
SSH_HOST = '127.0.0.1'
SSH_PORT = 3389
VMADDR_CID_ANY = 0xFFFFFFFF
BUFFER_SIZE = 65536  # Increased from 8192 to 64KB for better performance

# Load ws2_32.dll
ws2_32 = ctypes.WinDLL('ws2_32', use_last_error=True)

# Structures
class sockaddr_vm(ctypes.Structure):
    _fields_ = [
        ("svm_family", ctypes.c_ushort),
        ("svm_reserved1", ctypes.c_ushort),
        ("svm_port", ctypes.c_uint),
        ("svm_cid", ctypes.c_uint),
    ]

class sockaddr_in(ctypes.Structure):
    _fields_ = [
        ("sin_family", ctypes.c_short),
        ("sin_port", ctypes.c_ushort),
        ("sin_addr", ctypes.c_byte * 4),
        ("sin_zero", ctypes.c_byte * 8),
    ]

# Winsock functions
socket_func = ws2_32.socket
socket_func.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_int]
socket_func.restype = ctypes.c_int

bind_func = ws2_32.bind
bind_func.argtypes = [ctypes.c_int, ctypes.c_void_p, ctypes.c_int]
bind_func.restype = ctypes.c_int

listen_func = ws2_32.listen
listen_func.argtypes = [ctypes.c_int, ctypes.c_int]
listen_func.restype = ctypes.c_int

accept_func = ws2_32.accept
accept_func.argtypes = [ctypes.c_int, ctypes.c_void_p, ctypes.POINTER(ctypes.c_int)]
accept_func.restype = ctypes.c_int

connect_func = ws2_32.connect
connect_func.argtypes = [ctypes.c_int, ctypes.c_void_p, ctypes.c_int]
connect_func.restype = ctypes.c_int

closesocket_func = ws2_32.closesocket
closesocket_func.argtypes = [ctypes.c_int]
closesocket_func.restype = ctypes.c_int

recv_func = ws2_32.recv
recv_func.argtypes = [ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
recv_func.restype = ctypes.c_int

send_func = ws2_32.send
send_func.argtypes = [ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
send_func.restype = ctypes.c_int

WSAGetLastError = ws2_32.WSAGetLastError
WSAGetLastError.restype = ctypes.c_int

# WSAStartup
class WSADATA(ctypes.Structure):
    _fields_ = [
        ("wVersion", ctypes.c_ushort),
        ("wHighVersion", ctypes.c_ushort),
        ("szDescription", ctypes.c_char * 257),
        ("szSystemStatus", ctypes.c_char * 129),
        ("iMaxSockets", ctypes.c_ushort),
        ("iMaxUdpDg", ctypes.c_ushort),
        ("lpVendorInfo", ctypes.c_char_p),
    ]

WSAStartup = ws2_32.WSAStartup
WSAStartup.argtypes = [ctypes.c_ushort, ctypes.POINTER(WSADATA)]
WSAStartup.restype = ctypes.c_int

WSACleanup = ws2_32.WSACleanup
WSACleanup.restype = ctypes.c_int

def init_winsock():
    """Initialize Winsock"""
    wsa_data = WSADATA()
    result = WSAStartup(0x0202, ctypes.byref(wsa_data))  # Request Winsock 2.2
    if result != 0:
        raise OSError(f"WSAStartup failed: {result}")
    print("[+] Winsock initialized")

def htons(port):
    """Host to network short (byte order)"""
    return ((port & 0xFF) << 8) | ((port >> 8) & 0xFF)

def inet_addr(ip_string):
    """Convert IP string to network byte order"""
    parts = [int(x) for x in ip_string.split('.')]
    return (parts[0]) | (parts[1] << 8) | (parts[2] << 16) | (parts[3] << 24)

def create_vsock_server():
    """Create vsock server socket"""
    
    # Create socket
    sock_fd = socket_func(AF_VSOCK, SOCK_STREAM, 0)
    if sock_fd == INVALID_SOCKET:
        raise OSError(f"socket() failed: {WSAGetLastError()}")
    
    print(f"[+] Created vsock socket (fd={sock_fd})")
    
    # Bind
    addr = sockaddr_vm()
    addr.svm_family = AF_VSOCK
    addr.svm_reserved1 = 0
    addr.svm_port = VSOCK_PORT
    addr.svm_cid = VMADDR_CID_ANY
    
    print(f"[*] Binding to CID {addr.svm_cid:#x}, port {addr.svm_port}")
    
    if bind_func(sock_fd, ctypes.byref(addr), ctypes.sizeof(addr)) == SOCKET_ERROR:
        closesocket_func(sock_fd)
        raise OSError(f"bind() failed: {WSAGetLastError()}")
    
    print(f"[+] Bound successfully")
    
    # Listen
    if listen_func(sock_fd, 5) == SOCKET_ERROR:
        closesocket_func(sock_fd)
        raise OSError(f"listen() failed: {WSAGetLastError()}")
    
    print(f"[+] Listening on vsock port {VSOCK_PORT}")
    
    return sock_fd

def connect_to_ssh():
    """Create TCP connection to SSH server"""
    
    # Create TCP socket
    sock_fd = socket_func(AF_INET, SOCK_STREAM, 0)
    if sock_fd == INVALID_SOCKET:
        raise OSError(f"socket() failed: {WSAGetLastError()}")
    
    # Connect to SSH
    addr = sockaddr_in()
    addr.sin_family = AF_INET
    addr.sin_port = htons(SSH_PORT)
    
    # Convert IP address
    ip_bytes = inet_addr(SSH_HOST)
    addr.sin_addr = (ctypes.c_byte * 4).from_buffer_copy(ip_bytes.to_bytes(4, 'little'))
    
    if connect_func(sock_fd, ctypes.byref(addr), ctypes.sizeof(addr)) == SOCKET_ERROR:
        closesocket_func(sock_fd)
        raise OSError(f"connect() failed: {WSAGetLastError()}")
    
    return sock_fd

def relay_data(source_fd, dest_fd, direction):
    """Relay data between two sockets using raw file descriptors"""
    buffer = ctypes.create_string_buffer(BUFFER_SIZE)
    
    try:
        while True:
            # Receive data
            bytes_received = recv_func(source_fd, ctypes.addressof(buffer), BUFFER_SIZE, 0)
            
            if bytes_received > 0:
                # Send all data
                total_sent = 0
                data_to_send = bytes_received
                
                while total_sent < data_to_send:
                    sent = send_func(dest_fd, 
                                   ctypes.addressof(buffer) + total_sent, 
                                   data_to_send - total_sent, 
                                   0)
                    if sent == SOCKET_ERROR:
                        error = WSAGetLastError()
                        print(f"[-] {direction}: send error {error}")
                        return
                    total_sent += sent
                    
            elif bytes_received == 0:
                # Connection closed gracefully
                print(f"[*] {direction}: Connection closed")
                break
            else:
                # Error
                error = WSAGetLastError()
                print(f"[-] {direction}: recv error {error}")
                break
                    
    except Exception as e:
        print(f"[-] {direction}: Exception {e}")
        import traceback
        traceback.print_exc()

def handle_connection(vsock_fd, client_cid, client_port):
    """Handle vsock connection by relaying to SSH"""
    print(f"\n[+] Connection from CID {client_cid}, port {client_port}")
    
    ssh_fd = None
    try:
        # Connect to SSH
        ssh_fd = connect_to_ssh()
        print(f"[+] Connected to SSH at {SSH_HOST}:{SSH_PORT}")
        print(f"[*] Starting bidirectional relay...")
        
        # Create relay threads
        t1 = threading.Thread(target=relay_data, args=(vsock_fd, ssh_fd, "vsock->SSH"), name="vsock->SSH")
        t2 = threading.Thread(target=relay_data, args=(ssh_fd, vsock_fd, "SSH->vsock"), name="SSH->vsock")
        
        t1.daemon = True
        t2.daemon = True
        t1.start()
        t2.start()
        
        print(f"[*] Relay threads started")
        
        # Wait for both threads
        t1.join()
        t2.join()
        
        print(f"[+] Connection from CID {client_cid} finished")
        
    except Exception as e:
        print(f"[-] Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print(f"[*] Closing connection from CID {client_cid}")
        closesocket_func(vsock_fd)
        if ssh_fd is not None:
            closesocket_func(ssh_fd)

def daemonize():
    """Daemonize the process on Windows using DETACHED_PROCESS"""
    if sys.platform != 'win32':
        print("[-] Daemon mode only supported on Windows")
        return
    
    # On Windows, we'll use subprocess to create a detached process
    import subprocess
    
    # Get the current script path
    script_path = os.path.abspath(sys.argv[0])
    
    # Prepare arguments (remove -f/--foreground if present)
    args = [arg for arg in sys.argv[1:] if arg not in ['-f', '--foreground']]
    
    # Add foreground flag to prevent infinite loop
    args.append('--foreground')
    
    # Create detached process
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen(
        [sys.executable, script_path] + args,
        creationflags=DETACHED_PROCESS,
        close_fds=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    
    print("[+] Started as background process")
    sys.exit(0)

def run_server():
    """Main server loop"""
    print("[*] vsock SSH Relay for Windows")
    print("[*] Using raw Windows sockets")
    print(f"[*] Buffer size: {BUFFER_SIZE} bytes")
    print()
    
    # Initialize Winsock
    init_winsock()
    
    server_fd = None
    try:
        # Create server
        server_fd = create_vsock_server()
        
        print(f"[+] Forwarding vsock port {VSOCK_PORT} -> {SSH_HOST}:{SSH_PORT}")
        print(f"\n[*] Ready! Press Ctrl+C to stop\n")
        
        # Accept loop
        while True:
            try:
                # Accept connection
                client_addr = sockaddr_vm()
                addr_len = ctypes.c_int(ctypes.sizeof(client_addr))
                
                client_fd = accept_func(server_fd, ctypes.byref(client_addr), ctypes.byref(addr_len))
                
                if client_fd == INVALID_SOCKET:
                    print(f"[-] accept() failed: {WSAGetLastError()}")
                    continue
                
                # Handle in thread
                thread = threading.Thread(
                    target=handle_connection, 
                    args=(client_fd, client_addr.svm_cid, client_addr.svm_port)
                )
                thread.daemon = True
                thread.start()
                
            except KeyboardInterrupt:
                print("\n[*] Shutting down...")
                break
            except Exception as e:
                print(f"[-] Error: {e}")
                import traceback
                traceback.print_exc()
                
    except Exception as e:
        print(f"[-] Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if server_fd is not None:
            closesocket_func(server_fd)
        WSACleanup()

def main():
    parser = argparse.ArgumentParser(description='vsock to SSH relay for Windows')
    parser.add_argument('-f', '--foreground', action='store_true',
                        help='Run in foreground (default: run as background process)')
    
    args = parser.parse_args()
    
    if not args.foreground:
        # Daemonize
        daemonize()
    else:
        # Run in foreground
        run_server()

if __name__ == "__main__":
    main()
