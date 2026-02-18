Set-Location "c:\dev\me\mikes-windows-tools"
$msg = "chore: remove leftover _do_commit.ps1 temp script"
$tmp = [System.IO.Path]::GetTempFileName()
[System.IO.File]::WriteAllText($tmp, $msg)
git add -A
git commit -F $tmp
Remove-Item $tmp
