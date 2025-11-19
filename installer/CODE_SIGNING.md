# Code Signing and Publisher Verification

## The "Publisher is Unverified" Warning

When Revit loads revit-ballet, you may see a warning: **"Publisher could not be verified. Are you sure you want to run this software?"**

This is different from the "Load Always/Load Once" prompt that our registry entries prevent.

### Two Types of Revit Security Prompts:

1. **"Load Always/Load Once"** (✅ SOLVED by registry entries)
   - Appears for unsigned addins on first load
   - Asks user to choose "Load Always" or "Load Once"
   - **Fixed by**: `HKEY_CURRENT_USER\...\CodeSigning` registry entries
   - Our installer already handles this!

2. **"Publisher is Unverified"** (⚠️ Requires code signing certificate)
   - Appears for DLLs without digital signatures
   - Windows/Revit can't verify the publisher
   - **Only fixed by**: Code signing certificate

## How to Fix "Publisher is Unverified"

### Option 1: Code Signing Certificate (Recommended for Production)

**What you need:**
- Code signing certificate from a trusted Certificate Authority (CA)
- Costs: ~$100-400/year depending on provider

**Popular providers:**
- DigiCert
- Sectigo (formerly Comodo)
- GlobalSign
- SSL.com

**Process:**
1. Purchase code signing certificate
2. Install certificate on your build machine
3. Sign the DLL during build:
   ```bash
   signtool sign /f certificate.pfx /p password /tr http://timestamp.digicert.com /td SHA256 /fd SHA256 revit-ballet.dll
   ```
4. Sign the installer.exe as well

**Benefits:**
- No security warnings
- Professional appearance
- Users can verify publisher identity
- Required for enterprise deployment

### Option 2: Self-Signed Certificate (Testing Only)

**Not recommended** - Creates different warnings and requires users to install your certificate as trusted.

### Option 3: Live with the Warning (Acceptable for Personal Use)

If this is for personal/internal use:
- Users can click "Load" when prompted
- Warning only appears once per Revit session
- Our registry entries prevent repeated prompts
- No cost, but not professional

## Recommendation

**For personal/internal use**: Option 3 (no certificate needed)

**For public distribution**: Option 1 (purchase code signing certificate)

## Adding Code Signing to Build Process

If you get a certificate, add this to build scripts:

```bash
# Build the DLL
dotnet build commands/revit-ballet.csproj /p:RevitYear=2024 -c Release

# Sign the DLL
signtool sign /f "path\to\certificate.pfx" /p "password" \
  /tr http://timestamp.digicert.com /td SHA256 /fd SHA256 \
  commands/bin/2024/revit-ballet.dll

# Build installer (which embeds signed DLL)
dotnet build installer/installer.csproj -c Release

# Sign the installer
signtool sign /f "path\to\certificate.pfx" /p "password" \
  /tr http://timestamp.digicert.com /td SHA256 /fd SHA256 \
  installer/bin/Release/installer.exe
```

**Important**: Always use timestamping (`/tr` flag) so signatures remain valid even after certificate expires.

## Related Files

- `cleanup-registry.ps1` - PowerShell script to remove CodeSigning registry entries (for testing)
