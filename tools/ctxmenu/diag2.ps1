# Quick diag - list all verbs for SystemFileAssociations\image\shell
# and check for Extended (shift-only) verbs and printto-style hidden verbs
$hkcr = [Microsoft.Win32.Registry]::ClassesRoot

Write-Host "=== SystemFileAssociations\image\shell ==="
$k = $hkcr.OpenSubKey('SystemFileAssociations\image\shell')
if ($k) {
    foreach ($n in $k.GetSubKeyNames()) {
        $sk = $k.OpenSubKey($n)
        $mui = $sk.GetValue('MUIVerb')
        $ext = $sk.GetValue('Extended')
        $ld  = $sk.GetValue('LegacyDisable')
        Write-Host "  [$n]  MUI='$mui'  Extended=$(if($null -ne $ext){'YES'}else{'no'})  LegacyDisable=$(if($null -ne $ld){'YES'}else{'no'})"
        $sk.Close()
    }
    $k.Close()
} else { Write-Host "  (not found)" }

Write-Host ""
Write-Host "=== SystemFileAssociations\.jpg\shell ==="
$k = $hkcr.OpenSubKey('SystemFileAssociations\.jpg\shell')
if ($k) {
    foreach ($n in $k.GetSubKeyNames()) {
        $sk = $k.OpenSubKey($n)
        $mui = $sk.GetValue('MUIVerb')
        $ext = $sk.GetValue('Extended')
        Write-Host "  [$n]  MUI='$mui'  Extended=$(if($null -ne $ext){'YES'}else{'no'})"
        $sk.Close()
    }
    $k.Close()
} else { Write-Host "  (not found)" }

Write-Host ""
Write-Host "=== jpegfile\shell ==="
$k = $hkcr.OpenSubKey('jpegfile\shell')
if ($k) {
    foreach ($n in $k.GetSubKeyNames()) {
        $sk = $k.OpenSubKey($n)
        $mui = $sk.GetValue('MUIVerb')
        $ext = $sk.GetValue('Extended')
        Write-Host "  [$n]  MUI='$mui'  Extended=$(if($null -ne $ext){'YES'}else{'no'})"
        $sk.Close()
    }
    $k.Close()
} else { Write-Host "  (not found)" }

Write-Host ""
Write-Host "=== AppXkcg7y8nhftqvy2sxh5grnx46hbybnk39\shell (Windows Photos?) ==="
$k = $hkcr.OpenSubKey('AppXkcg7y8nhftqvy2sxh5grnx46hbybnk39\shell')
if ($k) {
    foreach ($n in $k.GetSubKeyNames()) {
        $sk = $k.OpenSubKey($n)
        $mui = $sk.GetValue('MUIVerb')
        $ext = $sk.GetValue('Extended')
        Write-Host "  [$n]  MUI='$mui'  Extended=$(if($null -ne $ext){'YES'}else{'no'})"
        $sk.Close()
    }
    $k.Close()
} else { Write-Host "  (not found)" }

# Search ALL AppX keys for "rotate" or "Photos" verbs
Write-Host ""
Write-Host "=== AppX keys with rotate/Photos verbs ==="
foreach ($name in $hkcr.GetSubKeyNames() | Where-Object { $_ -like 'AppX*' }) {
    $sk = $hkcr.OpenSubKey("$name\shell")
    if (-not $sk) { continue }
    foreach ($v in $sk.GetSubKeyNames()) {
        if ($v -match 'rotate|photo|rotateimage|RotateCW|RotateCCW' -or
            ($sk.OpenSubKey($v).GetValue('MUIVerb') -match 'rotate|Rotate')) {
            $vk = $sk.OpenSubKey($v)
            $appName = $hkcr.OpenSubKey($name).GetValue('')
            Write-Host "  $name ($appName) -> [$v] MUI='$($vk.GetValue('MUIVerb'))'"
            $vk.Close()
        }
    }
    $sk.Close()
}
