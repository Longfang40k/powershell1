Get-ChildItem -Path E:\Shares\Common\ -Include *.* -File -Recurse | foreach { $_.Delete()}
