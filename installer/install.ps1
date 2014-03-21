$myDocuments = [environment]::getfolderpath("mydocuments")
$visualStudio2012Folder = Get-ChildItem $myDocuments -Filter "Visual Studio 2012"
$visualStudio2013Folder = Get-ChildItem $myDocuments -Filter "Visual Studio 2013"

if (($visualStudio2012Folder -eq $null) -and ($visualStudio2013Folder -eq $null)) {
	Write-Error "Could not install Aspose Debugger Visualizer. None of the supported Visual Studio Version (2012, 2013) were found" -ErrorAction Stop
}

Write-Output "Downloading Aspose.Words 14.2.0 from NuGet"
.\NuGet.exe install Aspose.Words -Version 14.2.0 -ExcludeVersion -NoCache -NonInteractive -OutputDirectory .

if ($visualStudio2012Folder -ne $null) {
    Write-Output "Installing Aspose Debugger Visualizer for Visual Studio 2012."
	Copy-Item "$PSScriptRoot\AsposeVisualizer.2012.dll" -Destination "$($visualStudio2012Folder.FullName)\Visualizers"
	Copy-Item "$PSScriptRoot\Aspose.Words\lib\net3.5-client\Aspose.Words.dll" -Destination "$($visualStudio2012Folder.FullName)\Visualizers"
	Write-Output "Installed Aspose Debugger Visualizer successfully for Visual Studio 2012"
}

if ($visualStudio2013Folder -ne $null) {
    Write-Output "Installing Aspose Debugger Visualizer for Visual Studio 2013."
	Copy-Item "$PSScriptRoot\AsposeVisualizer.2013.dll" -Destination "$($visualStudio2013Folder.FullName)\Visualizers"
	Copy-Item "$PSScriptRoot\Aspose.Words\lib\net3.5-client\Aspose.Words.dll" -Destination "$($visualStudio2013Folder.FullName)\Visualizers"
	Write-Output "Installed Aspose Debugger Visualizer successfully for Visual Studio 2013"
}

Remove-Item -Recurse -Force "$PSScriptRoot\Aspose.Words"
