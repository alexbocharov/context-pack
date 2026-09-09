<#
.SYNOPSIS
    📦 context-pack - A universal CLI tool to pack any repository into a single document.
.PARAMETER Format
    The output format: 'docx' (default, requires MS Word) or 'txt' (lightweight plain text).
.EXAMPLE
    .\pack.ps1 -Format docx
.EXAMPLE
    .\pack.ps1 -Format txt
#>
param (
    [ValidateSet('docx', 'txt')]
    [string]$Format = 'docx'
)

# Comprehensive list of production source code and configuration extensions
$AllowedExtensions = @(
    # .NET & C# / F#
    ".cs", ".cshtml", ".razor", ".fs", ".csproj", ".fsproj", ".sln", ".config",
    # JavaScript / TypeScript & Web
    ".js", ".jsx", ".ts", ".tsx", ".json", ".html", ".css", ".scss", ".sass", ".vue",
    # Python
    ".py", ".ipynb", ".toml",
    # Go / Rust / C / C++
    ".go", ".rs", ".c", ".cpp", ".h", ".hpp",
    # Java / Kotlin / Scala
    ".java", ".kt", ".scala", ".gradle",
    # DevOps / Cloud / Container / Config
    ".yaml", ".yml", ".tf", ".dockerfile", "dockerfile", ".ini", ".env", ".example",
    # Documentation & Data
    ".md", ".txt", ".sql", ".graphql", ".proto"
)

# Comprehensive list of build artifacts, dependencies, caches, and IDE folders to ignore
$ExcludeFolders = @(
    # Generic & Version Control
    ".git", ".svn", ".hf", ".github",
    # IDEs & Editors
    ".vs", ".idea", ".vscode",
    # .NET & Aspire
    "bin", "obj",
    # Node.js / Frontend
    "node_modules", ".next", ".nuxt", "dist", "out", "build",
    # Python
    "__pycache__", ".venv", "venv", "env", ".pytest_cache", ".mypy_cache",
    # Rust / Go / Java
    "target", "vendor", ".gradle", ".m2",
    # Output formats to prevent self-packing loops
    ".docx", ".pdf", ".zip", ".tar.gz"
)

# Get current repository details dynamically
$RepoRoot = $PSScriptRoot
$RepoName = (Split-Path $RepoRoot -Leaf)
$OutputFileName = "Repository_Source_Code.$Format"
$OutputPath = Join-Path $PSScriptRoot $OutputFileName

Write-Host "Packing repository: [$RepoName] into [$Format] format..." -ForegroundColor Cyan
Write-Host "Scanning files..." -ForegroundColor Cyan

# Recursive file scan with exclusions
$Files = Get-ChildItem -Path $RepoRoot -Recurse -File | Where-Object {
    $FilePath = $_.FullName
    $FileName = $_.Name.ToLower()
    $Extension = $_.Extension.ToLower()
    $Include = $false
    
    if ($AllowedExtensions -contains $Extension -or $AllowedExtensions -contains $FileName) {
        $Include = $true
        foreach ($Folder in $ExcludeFolders) {
            if ($FilePath -like "*\$Folder\*") {
                $Include = $false
                break
            }
        }
    }
    return $Include
}

$Count = 0

# ==============================================================================
# MODE 1: MICROSOFT WORD EXPORT (.docx)
# ==============================================================================
if ($Format -eq 'docx') {
    # Clean up zombie Word processes from previous failed runs before starting
    Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
    
    try {
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Doc = $Word.Documents.Add()
        $Selection = $Word.Selection
    } catch {
        Write-Host "Error: Microsoft Word COM object could not be initialized. Please ensure MS Word is installed or use '-Format txt'." -ForegroundColor Red
        Exit
    }

    # Document Header (Using language-agnostic Heading 1 and Normal styles)
    $Selection.Style = "Heading 1"
    $Selection.TypeText("Repository Source Code: $RepoName`n")
    $Selection.Style = "Normal"
    $Selection.TypeText("Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n")
    $Selection.TypeText("Root Path: $RepoRoot`n")
    $Selection.InsertBreak(7) # 7 is the integer value for wdPageBreak, works everywhere

    foreach ($File in $Files) {
        # Skip output file if running repeatedly
        if ($File.FullName -eq $OutputPath) { continue } 
        
        $RelativePath = $File.FullName.Substring($RepoRoot.Length + 1)
        Write-Host "Processing ($(($Count+1))): $RelativePath" -ForegroundColor Yellow
        
        # File Path Header Style
        $Selection.Font.Bold = $true
        $Selection.Font.Size = 11
        $Selection.Font.Name = "Consolas"
        $Selection.TypeText("📄 File: $RelativePath`n")
        $Selection.Font.Bold = $false
        $Selection.TypeText("--------------------------------------------------------------------------------`n")
        
        try {
            $Content = [System.IO.File]::ReadAllText($File.FullName, [System.Text.Encoding]::UTF8)
            $Selection.Font.Size = 8.5
            $Selection.Font.Name = "Consolas"
            $Selection.TypeText($Content + "`n`n")
        } catch {
            $Selection.Font.Italic = $true
            $Selection.TypeText("[Binary file or error reading content]`n`n")
            $Selection.Font.Italic = $false
        }
        
        $Selection.Font.Size = 10
        $Selection.TypeText("================================================================================`n`n")
        $Count++
    }

    try {
        $Doc.SaveAs($OutputPath)
        $Doc.Close()
        $Word.Quit()
    } catch {
        Write-Host "Error saving the document. Please ensure the file is not open in another window." -ForegroundColor Red
    } finally {
        if ($Word) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null }
    }
}

# ==============================================================================
# MODE 2: PLAIN TEXT EXPORT (.txt)
# ==============================================================================
else {
    if (Test-Path $OutputPath) { Remove-Item $OutputPath }
    
    $Header = @"
================================================================================
REPOSITORY: $RepoName
GENERATED ON: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
ROOT PATH: $RepoRoot
================================================================================

"@
    Add-Content -Path $OutputPath -Value $Header -Encoding UTF8

    foreach ($File in $Files) {
        if ($File.FullName -eq $OutputPath) { continue }
        $RelativePath = $File.FullName.Substring($RepoRoot.Length + 1)
        Write-Host "Processing ($(($Count+1))): $RelativePath" -ForegroundColor Yellow
        
        $FileHeader = @"
--------------------------------------------------------------------------------
📄 File: $RelativePath
--------------------------------------------------------------------------------
"@
        Add-Content -Path $OutputPath -Value $FileHeader -Encoding UTF8
        
        try {
            $Content = [System.IO.File]::ReadAllText($File.FullName, [System.Text.Encoding]::UTF8)
            Add-Content -Path $OutputPath -Value $Content -Encoding UTF8
        } catch {
            Add-Content -Path $OutputPath -Value "[Binary file or error reading content]" -Encoding UTF8
        }
        
        Add-Content -Path $OutputPath -Value "`n================================================================================`n`n" -Encoding UTF8
        $Count++
    }
}

Write-Host "Done! Successfully processed $Count files." -ForegroundColor Green
Write-Host "File saved to: $OutputPath" -ForegroundColor Green
