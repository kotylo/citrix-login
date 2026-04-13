# Helper: reads password from Windows Credential Manager and writes to stdout
Add-Type -Namespace Cred -Name Mgr -MemberDefinition @'
[DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
public static extern bool CredReadW(string target, int type, int reserved, out IntPtr cred);
[DllImport("advapi32.dll")]
public static extern void CredFree(IntPtr cred);
[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct CREDENTIAL {
    public int Flags; public int Type; public string TargetName; public string Comment;
    public long LastWritten; public int CredentialBlobSize; public IntPtr CredentialBlob;
    public int Persist; public int AttributeCount; public IntPtr Attributes;
    public string TargetAlias; public string UserName;
}
'@

$ptr = [IntPtr]::Zero
$ok = [Cred.Mgr]::CredReadW($args[0], 1, 0, [ref]$ptr)
if (-not $ok) {
    Write-Error "Failed to read credential '$($args[0])'"
    exit 1
}
$c = [System.Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [Type][Cred.Mgr+CREDENTIAL])
$pw = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($c.CredentialBlob, $c.CredentialBlobSize / 2)
[Cred.Mgr]::CredFree($ptr)
Write-Host $pw -NoNewline
