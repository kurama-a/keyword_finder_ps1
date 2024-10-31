# Directory path to scan
$DirectoryPath = "C:\Users\arnau\Desktop\Travail\"
# Keywords to search
$Keywords = @("test", "rien") # Add your keywords here
# Output file
$OutputCsv = "C:\Users\arnau\Desktop\Travail\results.csv"

# Initialize the CSV file
"File Path, Keywords Found" | Out-File -FilePath $OutputCsv

# Function to search for all keywords in text content and return those found
function FindKeywords {
    param (
        [string]$Content
    )
    $foundKeywords = @()
    foreach ($keyword in $Keywords) {
        if ($Content -match $keyword) {
            $foundKeywords += $keyword
        }
    }
    return $foundKeywords
}

# Function to read the content of a Word file
function Read-WordFile {
    param ([string]$filePath)
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $doc = $Word.Documents.Open($filePath)
    $content = $doc.Content.Text
    $doc.Close()
    $Word.Quit()
    return $content
}

# Function to read the content of an Excel file
function Read-ExcelFile {
    param ([string]$filePath)
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $Excel.Workbooks.Open($filePath)
    $content = ""
    foreach ($sheet in $Workbook.Sheets) {
        foreach ($cell in $sheet.UsedRange) {
            $content += $cell.Text + " "
        }
    }
    $Workbook.Close()
    $Excel.Quit()
    return $content
}

# Function to read the content of a PowerPoint file
function Read-PowerPointFile {
    param ([string]$filePath)
    $PowerPoint = New-Object -ComObject PowerPoint.Application
    $PowerPoint.Visible = $false
    $Presentation = $PowerPoint.Presentations.Open($filePath, $false, $true, $false)
    $content = ""
    foreach ($slide in $Presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame) {
                $content += $shape.TextFrame.TextRange.Text + " "
            }
        }
    }
    $Presentation.Close()
    $PowerPoint.Quit()
    return $content
}

# Extensions to process
$extensionsToProcess = @(".docx", ".doc", ".xlsx", ".pptx", ".pdf", ".txt", ".csv", ".odt", ".ods", ".odp", ".xls")

# Get all files to process
$files = Get-ChildItem -Path $DirectoryPath -Recurse | Where-Object { $extensionsToProcess -contains $_.Extension.ToLower() }
$totalFiles = $files.Count
$currentFile = 0

# Process each file
foreach ($file in $files) {
    $currentFile++
    Write-Progress -Activity "Scanning files" -Status "$currentFile out of $totalFiles" -PercentComplete (($currentFile / $totalFiles) * 100)
    
    $filePath = $file.FullName
    $content = ""
    $foundKeywords = @() # Array to store all found keywords

    # Check if keywords are in the file name
    foreach ($keyword in $Keywords) {
        if ($filePath -match $keyword) {
            $foundKeywords += $keyword
        }
    }

    # If no keywords are found in the filename, check the content
    if ($foundKeywords.Count -eq 0) {
        switch ($file.Extension.ToLower()) {
            ".txt" { $content = Get-Content -Path $filePath -ErrorAction SilentlyContinue | Out-String }
            ".csv" { $content = Get-Content -Path $filePath -ErrorAction SilentlyContinue | Out-String }
            ".docx" { $content = Read-WordFile -filePath $filePath }
            ".xlsx" { $content = Read-ExcelFile -filePath $filePath }
            ".pptx" { $content = Read-PowerPointFile -filePath $filePath }
            ".pdf" { $content = "PDF format not supported for content extraction" }
            ".doc" { $content = Read-WordFile -filePath $filePath }
            ".xls" { $content = Read-ExcelFile -filePath $filePath }
            ".odt" { $content = Read-WordFile -filePath $filePath }
            ".ods" { $content = Read-ExcelFile -filePath $filePath }
            ".odp" { $content = Read-PowerPointFile -filePath $filePath }
        }

        # Check for keywords in content if the file can be read
        if ($content -ne "") {
            $foundKeywords += FindKeywords -Content $content
        }
    }

    # Record result if any keywords are found
    if ($foundKeywords.Count -gt 0) {
        "$filePath, $($foundKeywords -join '; ')" | Out-File -FilePath $OutputCsv -Append
    }
}

Write-Output "Scan completed. Results saved in $OutputCsv."
