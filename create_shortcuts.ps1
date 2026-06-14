$desktop = [Environment]::GetFolderPath('Desktop')
$python = (Get-Command python).Source
$src = 'c:\工作文件\claude\doc-md\src'

$ws = New-Object -ComObject WScript.Shell

$sc = $ws.CreateShortcut("$desktop\转换工具.lnk")
$sc.TargetPath = $python
$sc.Arguments = "$src\converter_launcher.py"
$sc.WorkingDirectory = $src
$sc.Description = 'Word-Markdown双向智能转换 - 拖放文件到此图标'
$sc.Save()
Write-Output '转换工具.lnk OK'

$sc2 = $ws.CreateShortcut("$desktop\转换设置.lnk")
$sc2.TargetPath = $python
$sc2.Arguments = "$src\settings_app.py"
$sc2.WorkingDirectory = $src
$sc2.Description = '配置Word-Markdown转换参数'
$sc2.Save()
Write-Output '转换设置.lnk OK'
