Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = "选择图片文件夹"
if ($dialog.ShowDialog() -ne 'OK') { exit }

$folder = $dialog.SelectedPath
$exts = @('.jpg','.jpeg','.png','.gif','.webp','.bmp')
$images = Get-ChildItem -Path $folder -File | Where-Object { $exts -contains $_.Extension.ToLower() } | Sort-Object Name

if ($images.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show('文件夹中没有图片','提示')
    exit
}

$list = ""
for ($i = 0; $i -lt $images.Count; $i++) {
    $list += ($i+1).ToString() + ". " + $images[$i].Name + "`n"
}

$msg = "共 " + $images.Count + " 张图片：`n`n" + $list + "`n输入排列顺序（如 3,1,2）`n留空 = 保持原序加编号"
$input = [Microsoft.VisualBasic.Interaction]::InputBox($msg, "图片排序", "")

if ($input -eq '') {
    $order = 1..$images.Count
} else {
    $order = $input -split '[,\s]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
}

if ($order.Count -ne $images.Count) {
    [System.Windows.Forms.MessageBox]::Show("编号数量不对，需要 " + $images.Count + " 个","错误")
    exit
}

$pad = $images.Count.ToString().Length
$preview = ""
$ops = @()

for ($i = 0; $i -lt $order.Count; $i++) {
    $src = $images[$order[$i]-1]
    $base = $src.Name
    while ($base.Length -gt 0 -and $base[0] -ge '0' -and $base[0] -le '9') { $base = $base.Substring(1) }
    if ($base[0] -eq '_') { $base = $base.Substring(1) }
    if ($base -eq '') { $base = $src.Name }
    $new = "{0:D$pad}_{1}" -f ($i+1), $base
    $preview += ($i+1).ToString() + ". " + $src.Name + " -> " + $new + "`n"
    $ops += @{ Src = $src.FullName; Base = $base }
}

$c = [System.Windows.Forms.MessageBox]::Show("确认重命名？`n`n" + $preview, "确认", "YesNo")
if ($c -ne 'Yes') { exit }

foreach ($img in $images) {
    $tmp = "__tmp_" + [guid]::NewGuid().ToString("N").Substring(0,8) + "_" + $img.Name
    Rename-Item -Path $img.FullName -NewName $tmp
}

$tmpFiles = Get-ChildItem -Path $folder -File -Filter '__tmp_*' | Sort-Object Name
for ($i = 0; $i -lt $tmpFiles.Count; $i++) {
    $new = "{0:D$pad}_{1}" -f ($i+1), $ops[$i].Base
    Rename-Item -Path $tmpFiles[$i].FullName -NewName $new
}

[System.Windows.Forms.MessageBox]::Show("重命名完成！共 " + $images.Count + " 张","成功")
