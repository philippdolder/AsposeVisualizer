$myDocuments = [environment]::getfolderpath("mydocuments")
$visualStudioFolder = Get-ChildItem $myDocuments -Filter "Visual Studio 2012"

if ($visualStudioFolder -eq $null) {
    Write-Error "Cannot install Aspose Debugger Visualizer. Directory '$myDocuments\Visual Studio 2012' was not found" -ErrorAction Stop
}

Write-Output "Downloading Aspose.Words 13.10.0 from NuGet"
.\NuGet.exe install Aspose.Words -Version 13.10.0 -ExcludeVersion -NoCache -NonInteractive -OutputDirectory .

Copy-Item "$PSScriptRoot\AsposeVisualizer.dll" -Destination "$($visualStudioFolder.FullName)\Visualizers"
Copy-Item "$PSScriptRoot\Aspose.Words\lib\net3.5\Aspose.Words.dll" -Destination "$($visualStudioFolder.FullName)\Visualizers"

Remove-Item -Recurse -Force "$PSScriptRoot\Aspose.Words"

Write-Output "Installed Aspose Debugger Visualizer successfully"