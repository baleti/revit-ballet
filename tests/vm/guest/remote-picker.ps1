# End-to-end check of the Windows remote command picker (remote/RevitBalletRemote), run IN THE INTERACTIVE SESSION (see it.ps1; ssh sessions have no desktop).
# Starts RevitBalletRemote.exe from C:\ci\remote, presses the real hotkey (Win+`), types queries into the real window, reads the status label through UI Automation,
# presses Enter on a command and leaves the evidence in C:\ci\remote-picker.txt: lines "OK <name>" / "FAIL <name>: why". The caller then asks the agent whether the command ran.
param([int]$Port = 23717)
$ErrorActionPreference = "Stop"
trap { Add-Content "C:\ci\remote-picker.txt" ("FAIL exception: " + $_.Exception.Message + " at line " + $_.InvocationInfo.ScriptLineNumber); exit 1 }
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public static class Keys2 {
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
  public static void Chord(byte[] mods, byte key) {
    foreach (var m in mods) keybd_event(m, 0, 0, UIntPtr.Zero);
    keybd_event(key, 0, 0, UIntPtr.Zero); keybd_event(key, 0, 2, UIntPtr.Zero);
    foreach (var m in mods) keybd_event(m, 0, 2, UIntPtr.Zero);
  }
}
"@
$out = "C:\ci\remote-picker.txt"; Remove-Item $out -ErrorAction SilentlyContinue
function Ok($n) { Add-Content $out "OK $n" }
function Fail($n, $why) { Add-Content $out "FAIL ${n}: $why" }
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
function Find-Picker($timeoutMs = 8000) {
  $end = (Get-Date).AddMilliseconds($timeoutMs)
  while ((Get-Date) -lt $end) {
    $w = $AE::RootElement.FindFirst($TS::Children, (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "Revit commands")))
    if ($w) { return $w }
    Start-Sleep -Milliseconds 200
  }
  $null
}
# WinForms exposes AccessibleName as the UIA Name (the query box is named "query"); the status label's Name is its text ("12 of 241 commands ...")
function Query-Box($win) {
  # a native EDIT control does not surface AccessibleName: it is the one element of class WindowsForms10.EDIT*
  foreach ($e in $win.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) { if ($e.Current.ClassName -like "WindowsForms10.EDIT*") { return $e } }
  $null
}
function Send-Text($text) {
  # SendKeys metacharacters + ^ % ~ ( ) { } [ ] are escaped by wrapping them in braces
  $esc = [regex]::Replace($text, '[+^%~(){}\[\]]', '{$0}')
  [System.Windows.Forms.SendKeys]::SendWait($esc)
}
function Set-Query($win, $text) {
  # no SetFocus: the picker must already have the keyboard focus on its query box after the hotkey (that is part of what is tested)
  Start-Sleep -Milliseconds 200
  [System.Windows.Forms.SendKeys]::SendWait("^a"); [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
  if ($text) { Send-Text $text }
  Start-Sleep -Milliseconds 800
}
function Status($win) {
  $all = $win.FindAll($TS::Descendants, (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)))
  foreach ($e in $all) { if ($e.Current.Name -match '^\d+ of \d+ commands') { return $e.Current.Name } }
  ""
}

Get-Process RevitBalletRemote -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500
$exe = "C:\ci\remote\RevitBalletRemote.exe"
Start-Process $exe            # no --show: the picker must come from the hotkey, which is what is under test
Start-Sleep -Seconds 3

# 1. the real hotkey opens the picker
[Keys2]::Chord(@(0x5B), 0xC0)       # LWIN + VK_OEM_3 (`)
$win = Find-Picker
if (-not $win) { Fail "hotkey-opens-picker" "no window 'Revit commands' after Win+``" ; exit 1 } else { Ok "hotkey-opens-picker" }

# 2. rows arrive (cache or live probe), counts follow the query language
Start-Sleep -Seconds 6
$all = Status $win
if ($all -match '^(\d+) of (\d+) commands') { $total = [int]$Matches[2]; if ($total -ge 50 -and [int]$Matches[1] -eq $total) { Ok "rows-loaded ($total)" } else { Fail "rows-loaded" "status=[$all]" } } else { Fail "rows-loaded" "status=[$all]"; exit 1 }
Set-Query $win "deselect"
$s = Status $win; if ($s -match '^([1-9]\d*) of (\d+) commands' -and [int]$Matches[1] -lt $total) { Ok "bare-word-filters ($s)" } else { Fail "bare-word-filters" "status=[$s]" }
Set-Query $win "//input any"
$s = Status $win; if ($s -match '^([1-9]\d*) of (\d+) commands' -and [int]$Matches[1] -lt $total) { Ok "scoped-filter ($s)" } else { Fail "scoped-filter" "status=[$s]" }
Set-Query $win "!deselect"
$s = Status $win; if ($s -match '^(\d+) of (\d+) commands' -and [int]$Matches[1] -gt 0 -and [int]$Matches[1] -lt $total) { Ok "negation ($s)" } else { Fail "negation" "status=[$s]" }
Set-Query $win "//zzz foo"
$s = Status $win; if ($s -match '^0 of') { Ok "unresolved-field-narrows-to-nothing" } else { Fail "unresolved-field" "status=[$s]" }
Set-Query $win "/"
$s = Status $win; if ($s -match "^$total of") { Ok "half-typed-is-inert" } else { Fail "half-typed" "status=[$s]" }

# 3. Esc closes, the hotkey toggles it back
[System.Windows.Forms.SendKeys]::SendWait("{ESC}"); Start-Sleep -Milliseconds 600
if (-not (Find-Picker 1500)) { Ok "escape-closes" } else { Fail "escape-closes" "window still there" }
[Keys2]::Chord(@(0x5B), 0xC0); $win = Find-Picker
if ($win) { Ok "hotkey-reopens" } else { Fail "hotkey-reopens" "no window"; exit 1 }

# 4. run a command: CloseWorksets (no document is open, so Revit answers "No active document."); the caller checks the agent's job list
Start-Sleep -Seconds 3
Set-Query $win "closeworksets"
$s = Status $win; if ($s -match '^1 of') { Ok "one-row-for-closeworksets" } else { Fail "one-row-for-closeworksets" "status=[$s]" }
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}"); Start-Sleep -Seconds 2
if (-not (Find-Picker 1500)) { Ok "enter-hides-picker" } else { Fail "enter-hides-picker" "window still there" }
Add-Content $out "DONE"
