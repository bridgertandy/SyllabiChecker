# Source - https://stackoverflow.com/a
# Posted by MFT
# Retrieved 2025-11-20, License - CC BY-SA 3.0

# Modified by Bridger Tandy 2/13/2026

# Load the System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a new FolderBrowserDialog object
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# (Optional) Set the initial directory for the dialog
# $folderBrowser.SelectedPath = "C:\Users\YourUser\Documents"

# (Optional) Set a description for the dialog
$folderBrowser.Description = "Select a directory"

# Create an invisible anchor form that the dialog box is attached to so that it forces the dialog box to be active
$anchor = New-Object System.Windows.Forms.Form
$anchor.TopMost = $true
$anchor.Opacity = 0
$anchor.ShowInTaskbar = $false
$anchor.Show()
$anchor.Activate()

# Open dialog box with the anchor as the owner and capture the output
$result = $folderBrowser.ShowDialog($anchor)

# Check if the output was "OK" and capture directory path
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
	$documents_path = $folderBrowser.SelectedPath
	Write-Host "`nYou selected: $documents_path"
} else {
    Write-Host "`nNo directory was selected."
}    

# Dispose of the dialog object and anchor form to release resources
$folderBrowser.Dispose()
$anchor.Dispose()


$word_app = New-Object -ComObject Word.Application

Write-Host "Saved PDFs:"

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {

    $document = $word_app.Documents.Open($_.FullName)

    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"

    Start-Sleep -Milliseconds 500 
     
    $document.SaveAs([ref] $pdf_filename, [ref] 17)

    Write-Host "$($_.BaseName).pdf"

    $document.Close()
}
Write-Host "`n"
$word_app.Quit()
