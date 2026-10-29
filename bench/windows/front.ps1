# Excel's window in front, for the scripts of bench/windows that run on the screen (dot-sourced): Windows lets a process
# take the foreground once it sent the last input, so Shift goes down and up first; a window that keeps the foreground
# (Windows's Get Started did) stops the script.
Add-Type -Namespace XlScreen -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, UIntPtr x);
'@
function Front($hwnd) {
  $h = [IntPtr]$hwnd
  for ($i = 0; $i -lt 3; $i++) {
    [XlScreen.Win]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    [XlScreen.Win]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
    [void][XlScreen.Win]::SetForegroundWindow($h)
    Start-Sleep -Milliseconds 400
    if ([XlScreen.Win]::GetForegroundWindow() -eq $h) { return }
  }
  throw "Excel did not come to the front: $([XlScreen.Win]::GetForegroundWindow()) is, not $hwnd"
}
