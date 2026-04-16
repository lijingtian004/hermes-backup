@echo off
chcp 65001 >nul
echo ========================================
echo   图片排序重命名工具
echo ========================================
echo.
echo 正在启动 PowerShell...
echo.

powershell -ExecutionPolicy Bypass -NoProfile -Command "try { Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName Microsoft.VisualBasic; $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = '选择图片文件夹'; if ($d.ShowDialog() -ne 'OK') { Write-Host '已取消'; exit }; $folder = $d.SelectedPath; Write-Host '文件夹: ' + $folder; $exts = '.jpg','.jpeg','.png','.gif','.webp','.bmp'; $imgs = Get-ChildItem $folder -File | Where { $exts -contains $_.Extension.ToLower() } | Sort Name; Write-Host '找到 ' + $imgs.Count + ' 张图片'; if ($imgs.Count -eq 0) { Write-Host '没有图片！'; pause; exit }; for ($i=0;$i -lt $imgs.Count;$i++) { Write-Host ('  {0}. {1}' -f ($i+1), $imgs[$i].Name) }; Write-Host ''; $inp = Read-Host '输入顺序(如3,1,2) 回车=原序'; if ($inp -eq '') { $ord = 1..$imgs.Count } else { $ord = $inp -split '[,\s]+' | Where {$_ -match '^\d+$'} | % {[int]$_} }; if ($ord.Count -ne $imgs.Count) { Write-Host '数量不对! 需要' + $imgs.Count + '个'; pause; exit }; $pad = $imgs.Count.ToString().Length; Write-Host '预览:'; $ops = @(); for ($i=0;$i -lt $ord.Count;$i++) { $s = $imgs[$ord[$i]-1]; $b = $s.Name; while($b.Length -gt 0 -and $b[0] -ge '0' -and $b[0] -le '9'){$b=$b.Substring(1)}; if($b[0] -eq '_'){$b=$b.Substring(1)}; if($b -eq ''){$b=$s.Name}; $n = '{0:D' + $pad + '}_{1}' -f ($i+1),$b; Write-Host ('  {0}. {1} -> {2}' -f ($i+1),$s.Name,$n); $ops += @{Src=$s.FullName;New=$n} }; $c = Read-Host '确认? (y/n)'; if ($c -ne 'y') { Write-Host '已取消'; pause; exit }; Write-Host '重命名中...'; foreach($img in $imgs){ $t='__tmp_'+[guid]::NewGuid().ToString('N').Substring(0,8)+'_'+$img.Name; Rename-Item $img.FullName $t }; $tmps = Get-ChildItem $folder -File -Filter '__tmp_*' | Sort Name; for($i=0;$i -lt $tmps.Count;$i++){ Rename-Item $tmps[$i].FullName $ops[$i].New }; Write-Host '完成！共 ' + $imgs.Count + ' 张'; pause } catch { Write-Host '错误: ' + $_.Exception.Message; pause }"

echo.
pause
