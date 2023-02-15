# Define the directory to search
$dir = "CHANGE ME (i.e. C:\Users\username\Desktop\WordDocuments"

# Define the phrase(s) to search for
$phrases = "test phrase 1", "testphrase 2", "example phrase 3"

# Loop through each file in the directory
foreach ($file in Get-ChildItem $dir -Filter *.docx) {
    # Load the Word document
    $doc = New-Object -ComObject Word.Application
    $doc.Visible = $false
    $doc = $doc.Documents.Open($file.FullName)

    # Loop through each phrase and search for it in the document
    foreach ($phrase in $phrases) {
        $count = ($doc.Content.Text | Select-String -AllMatches $phrase).Matches.Count
        if ($count -gt 0) {
            Write-Host "Found '$phrase' in $($file.Name) ($count time(s))"
        }
    }

    # Close the Word document
    $doc.Close()
}

# Output the statistics
$stats = @{}
foreach ($phrase in $phrases) {
    $stats[$phrase] = 0
}
foreach ($file in Get-ChildItem $dir -Filter *.docx) {
    # Load the Word document
    $doc = New-Object -ComObject Word.Application
    $doc.Visible = $false
    $doc = $doc.Documents.Open($file.FullName)

    # Loop through each phrase and search for it in the document
    foreach ($phrase in $phrases) {
        $count = ($doc.Content.Text | Select-String -AllMatches $phrase).Matches.Count
        if ($count -gt 0) {
            $stats[$phrase] += 1
        }
    }

    # Close the Word document
    $doc.Close()
}
Write-Host ""
Write-Host "Statistics:"
foreach ($key in $stats.Keys) {
    Write-Host "$($key): $($stats[$key]) documents"
}

Read-Host "Press Enter to continue..."
