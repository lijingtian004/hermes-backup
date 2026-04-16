Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "图片排序工具"
$form.Size = New-Object System.Drawing.Size(900, 650)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(26, 26, 46)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)

# --- Top Panel ---
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Top"
$panel.Height = 55
$panel.BackColor = [System.Drawing.Color]::FromArgb(22, 33, 62)
$panel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 5)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "📂 选择文件夹"
$btnOpen.Size = New-Object System.Drawing.Size(120, 35)
$btnOpen.Location = New-Object System.Drawing.Point(10, 10)
$btnOpen.FlatStyle = "Flat"
$btnOpen.BackColor = [System.Drawing.Color]::FromArgb(233, 69, 96)
$btnOpen.ForeColor = [System.Drawing.Color]::White
$btnOpen.FlatAppearance.BorderSize = 0
$btnOpen.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnRename = New-Object System.Windows.Forms.Button
$btnRename.Text = "✅ 确认重命名"
$btnRename.Size = New-Object System.Drawing.Size(120, 35)
$btnRename.Location = New-Object System.Drawing.Point(140, 10)
$btnRename.FlatStyle = "Flat"
$btnRename.BackColor = [System.Drawing.Color]::FromArgb(39, 174, 96)
$btnRename.ForeColor = [System.Drawing.Color]::White
$btnRename.FlatAppearance.BorderSize = 0
$btnRename.Enabled = $false
$btnRename.Cursor = [System.Windows.Forms.Cursors]::Hand

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "🔄 重置"
$btnReset.Size = New-Object System.Drawing.Size(80, 35)
$btnReset.Location = New-Object System.Drawing.Point(270, 10)
$btnReset.FlatStyle = "Flat"
$btnReset.BackColor = [System.Drawing.Color]::FromArgb(15, 52, 96)
$btnReset.ForeColor = [System.Drawing.Color]::White
$btnReset.FlatAppearance.BorderSize = 0
$btnReset.Enabled = $false
$btnReset.Cursor = [System.Windows.Forms.Cursors]::Hand

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "共 0 张图片"
$lblCount.AutoSize = $true
$lblCount.Location = New-Object System.Drawing.Point(370, 18)
$lblCount.ForeColor = [System.Drawing.Color]::FromArgb(233, 69, 96)

$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = ""
$lblFolder.AutoSize = $true
$lblFolder.Location = New-Object System.Drawing.Point(500, 18)
$lblFolder.ForeColor = [System.Drawing.Color]::Gray

$panel.Controls.AddRange(@($btnOpen, $btnRename, $btnReset, $lblCount, $lblFolder))

# --- ListView ---
$listView = New-Object System.Windows.Forms.ListView
$listView.Dock = "Fill"
$listView.View = "LargeIcon"
$listView.BackColor = [System.Drawing.Color]::FromArgb(26, 26, 46)
$listView.ForeColor = [System.Drawing.Color]::White
$listView.AllowDrop = $false
$listView.MultiSelect = $false
$listView.HideSelection = $false
$listView.LargeImageList = New-Object System.Windows.Forms.ImageList
$listView.LargeImageList.ImageSize = New-Object System.Drawing.Size(160, 160)
$listView.LargeImageList.ColorDepth = "Depth32Bit"

# --- Status Bar ---
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(22, 33, 62)
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "拖拽图片调整顺序，然后点击「确认重命名」"
$statusLabel.ForeColor = [System.Drawing.Color]::Gray
$statusBar.Items.Add($statusLabel) | Out-Null

$form.Controls.Add($listView)
$form.Controls.Add($panel)
$form.Controls.Add($statusBar)

# --- Data ---
$script:images = @()  # array of @{ Name, Path, Thumb }
$script:dragIdx = -1

# --- Thumbnail ---
function Get-Thumbnail($path, $size = 160) {
    try {
        $img = [System.Drawing.Image]::FromFile($path)
        $ratio = [Math]::Min($size / $img.Width, $size / $img.Height)
        $w = [int]($img.Width * $ratio)
        $h = [int]($img.Height * $ratio)
        $thumb = New-Object System.Drawing.Bitmap($size, $size)
        $g = [System.Drawing.Graphics]::FromImage($thumb)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.Clear([System.Drawing.Color]::FromArgb(15, 15, 35))
        $x = ($size - $w) / 2
        $y = ($size - $h) / 2
        $g.DrawImage($img, $x, $y, $w, $h)
        $g.Dispose()
        $img.Dispose()
        return $thumb
    } catch { return $null }
}

# --- Refresh List ---
function Refresh-List {
    $listView.Items.Clear()
    $listView.LargeImageList.Images.Clear()

    for ($i = 0; $i -lt $script:images.Count; $i++) {
        $img = $script:images[$i]
        $thumb = Get-Thumbnail $img.Path
        if ($thumb) {
            $listView.LargeImageList.Images.Add($thumb)
            $thumb.Dispose()
        }
        $item = New-Object System.Windows.Forms.ListViewItem(($i + 1).ToString() + ". " + $img.Name)
        $item.ImageIndex = $i
        $item.Tag = $i
        $listView.Items.Add($item) | Out-Null
    }

    $lblCount.Text = "共 " + $script:images.Count + " 张图片"
    $btnRename.Enabled = $script:images.Count -gt 0
    $btnReset.Enabled = $script:images.Count -gt 0
}

# --- Open Folder ---
$btnOpen.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "选择图片文件夹"
    if ($dialog.ShowDialog() -ne "OK") { return }

    $folder = $dialog.SelectedPath
    $lblFolder.Text = $folder
    $exts = @('.jpg','.jpeg','.png','.gif','.webp','.bmp')

    $files = Get-ChildItem -Path $folder -File | Where-Object {
        $exts -contains $_.Extension.ToLower()
    } | Sort-Object Name

    $script:images = @()
    $statusLabel.Text = "正在加载..."
    $form.Refresh()

    foreach ($f in $files) {
        $script:images += @{ Name = $f.Name; Path = $f.FullName }
    }

    Refresh-List
    $statusLabel.Text = "已加载 " + $files.Count + " 张图片 — 拖拽调整顺序"
})

# --- Drag & Drop reorder ---
$listView.Add_ItemDrag({
    $script:dragIdx = $_.Item.Tag
})

$listView.Add_DragEnter({
    $_.Effect = [System.Windows.Forms.DragDropEffects]::Move
})

$listView.Add_DragOver({
    $target = $listView.GetItemAt($_.X, $_.Y)
    if ($target) {
        $listView.InsertionMark.Index = $target.Index
        $listView.InsertionMark.AppearsAfterItem = ($_.X - $target.Bounds.Left) -gt ($target.Bounds.Width / 2)
    } else {
        $listView.InsertionMark.Index = -1
    }
})

$listView.Add_DragLeave({
    $listView.InsertionMark.Index = -1
})

$listView.Add_DragDrop({
    $target = $listView.GetItemAt($_.X, $_.Y)
    if (-not $target -or $script:dragIdx -lt 0) { return }

    $toIdx = $target.Index
    if ($listView.InsertionMark.AppearsAfterItem) { $toIdx++ }
    if ($script:dragIdx -eq $toIdx -or $script:dragIdx -eq ($toIdx - 1)) { return }

    $item = $script:images[$script:dragIdx]
    $script:images = @($script:images | Where-Object { $_ -ne $item })
    if ($toIdx -gt $script:dragIdx) { $toIdx-- }
    $script:images = @($script:images[0..($toIdx-1)] + $item + $script:images[$toIdx..($script:images.Length-1)]) 2>$null
    if ($toIdx -eq 0) { $script:images = @($item) + $script:images }
    if ($toIdx -ge $script:images.Length) { $script:images = $script:images + $item }

    # Rebuild properly
    $newArr = @()
    $inserted = $false
    for ($i = 0; $i -le $script:images.Length; $i++) {
        if ($i -eq $toIdx -and -not $inserted) {
            $newArr += $item
            $inserted = $true
        }
        if ($i -lt $script:images.Length -and $script:images[$i] -ne $item) {
            $newArr += $script:images[$i]
        }
    }
    if (-not $inserted) { $newArr += $item }

    $script:images = $newArr
    $script:dragIdx = -1
    $listView.InsertionMark.Index = -1
    Refresh-List
})

# --- Reset ---
$btnReset.Add_Click({
    $folder = $lblFolder.Text
    if (-not $folder) { return }
    $exts = @('.jpg','.jpeg','.png','.gif','.webp','.bmp')
    $files = Get-ChildItem -Path $folder -File | Where-Object {
        $exts -contains $_.Extension.ToLower()
    } | Sort-Object Name
    $script:images = @()
    foreach ($f in $files) {
        $script:images += @{ Name = $f.Name; Path = $f.FullName }
    }
    Refresh-List
    $statusLabel.Text = "已重置顺序"
})

# --- Rename ---
$btnRename.Add_Click({
    if ($script:images.Count -eq 0) { return }

    $pad = $script:images.Count.ToString().Length
    $preview = ""
    for ($i = 0; $i -lt $script:images.Count; $i++) {
        $base = $script:images[$i].Name
        while ($base.Length -gt 0 -and $base[0] -ge '0' -and $base[0] -le '9') { $base = $base.Substring(1) }
        if ($base.Length -gt 0 -and $base[0] -eq '_') { $base = $base.Substring(1) }
        if ($base -eq '') { $base = $script:images[$i].Name }
        $new = "{0:D$pad}_{1}" -f ($i+1), $base
        $preview += ($i+1).ToString() + ". " + $script:images[$i].Name + " -> " + $new + "`n"
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "确认重命名 " + $script:images.Count + " 张图片？`n`n" + $preview,
        "确认重命名",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -ne "Yes") { return }

    try {
        # Phase 1: temp names
        foreach ($img in $script:images) {
            $tmp = "__tmp_" + [guid]::NewGuid().ToString("N").Substring(0,8) + "_" + $img.Name
            Rename-Item -Path $img.Path -NewName $tmp
        }

        # Phase 2: final names
        $folder = $lblFolder.Text
        $tmpFiles = Get-ChildItem -Path $folder -File -Filter '__tmp_*' | Sort-Object Name
        for ($i = 0; $i -lt $tmpFiles.Count; $i++) {
            $base = $script:images[$i].Name
            while ($base.Length -gt 0 -and $base[0] -ge '0' -and $base[0] -le '9') { $base = $base.Substring(1) }
            if ($base.Length -gt 0 -and $base[0] -eq '_') { $base = $base.Substring(1) }
            if ($base -eq '') { $base = $script:images[$i].Name }
            $new = "{0:D$pad}_{1}" -f ($i+1), $base
            Rename-Item -Path $tmpFiles[$i].FullName -NewName $new
        }

        # Reload
        $files = Get-ChildItem -Path $folder -File | Where-Object {
            $exts = @('.jpg','.jpeg','.png','.gif','.webp','.bmp')
            $exts -contains $_.Extension.ToLower()
        } | Sort-Object Name
        $script:images = @()
        foreach ($f in $files) {
            $script:images += @{ Name = $f.Name; Path = $f.FullName }
        }
        Refresh-List
        $statusLabel.Text = "✅ 重命名完成！"
        [System.Windows.Forms.MessageBox]::Show("重命名完成！", "成功")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("出错: " + $_.Exception.Message, "错误", "OK", "Error")
    }
})

$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
