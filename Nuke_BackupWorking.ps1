# Delete Backup Working folder
$path = "c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Backup Working"
$longPath = "\\?\$path"
$dir = [System.IO.DirectoryInfo]::new($longPath)
$dir.Delete($true)
Write-Host "Backup Working folder deleted!"
