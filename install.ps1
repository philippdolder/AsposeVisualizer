$myDocuments = [environment]::getfolderpath("mydocuments")
$visualStudioFolder = Get-ChildItem $myDocuments -Filter "Visual Studio 2012"

if ($visualStudioFolder -eq $null) {
    Write-Error "Cannot install Aspose Debugger Visualizer. Directory '$myDocuments\Visual Studio 2012' was not found" -ErrorAction Stop
}

Copy-Item "$PSScriptRoot\AsposeVisualizer.dll" -Destination "$($visualStudioFolder.FullName)\Visualizers"
Copy-Item "$PSScriptRoot\Aspose.Words.dll" -Destination "$($visualStudioFolder.FullName)\Visualizers"

Write-Output "Installed Aspose Debugger Visualizer successfully"