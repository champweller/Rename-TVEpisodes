#!/usr/bin/env pwsh

<#
.SYNOPSIS
    Automatically renames TV show episode files to Jellyfin naming convention using TMDB API

.DESCRIPTION
    This PowerShell Core script processes TV show directories and renames .mkv files to proper
    episode names using The Movie Database (TMDB) API. It works across Windows, Mac, and Linux.
    
    Directory Structure Expected:
    - Base folder contains subdirectories named: TV_SHOW_NAME_S1_D1
    - Each subdirectory contains .mkv files with generic names
    
    Output Format (Jellyfin Compatible):
    - Series Name S01E01.mkv
    - Series Name S01E02.mkv
    - Series Name S01E01-E02.mkv (for multi-episode files)

.PARAMETER BasePath
    The root directory containing TV show subdirectories (e.g., /home/KSW/Videos)

.PARAMETER UseTMDB
    Switch to enable TMDB API integration for accurate episode naming

.PARAMETER DryRun
    Switch to preview what files would be renamed without actually renaming them

.EXAMPLE
    ./Rename-TVEpisodes.ps1 -BasePath "/home/KSW/Videos" -UseTMDB
    
.EXAMPLE
    ./Rename-TVEpisodes.ps1 -BasePath "C:\Videos" -DryRun

.NOTES
    Author: Circuit Savers KSW
    Requires: PowerShell Core 6.0+ for cross-platform compatibility
    API: The Movie Database (TMDB) v3
#>

param(
    # Base directory containing TV show folders - REQUIRED
    [Parameter(Mandatory=$true, HelpMessage="Enter the base path containing TV show directories")]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$BasePath,
    
    # Enable TMDB API integration for episode naming - OPTIONAL
    [Parameter(Mandatory=$false, HelpMessage="Use TMDB API for accurate episode naming")]
    [switch]$UseTMDB,
    
    # Dry run mode - preview changes without executing - OPTIONAL
    [Parameter(Mandatory=$false, HelpMessage="Preview changes without actually renaming files")]
    [switch]$DryRun,

    # Optional destination folder to move renamed files - NEW FEATURE
    [Parameter(Mandatory=$false, HelpMessage="Destination folder to move renamed files to (will be created if it doesn't exist)")]
    [string]$MoveToFolder
)

# =============================================================================
# CONFIGURATION SECTION
# =============================================================================

# TMDB API Configuration
$TMDBConfig = @{
    APIKey = "YOUR_API_KEY"
    ReadAccessToken = "YOUR_READ_ACCESS_TOKEN"
    BaseURL = "https://api.themoviedb.org/3"
}

# File processing configuration
$Config = @{
    # Supported video file extensions
    SupportedExtensions = @('.mkv', '.mp4', '.avi', '.mov')
    # Minimum file size to process (in bytes) - helps skip sample files
    MinFileSize = 50MB
    # Maximum API requests per minute (TMDB limit is 40/10 seconds)
    APIRateLimit = 35
}

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

function Write-ColorOutput {
    <#
    .SYNOPSIS
        Writes colored output to console for better readability
    #>
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    
    if ($Host.UI.SupportsVirtualTerminal -or $PSVersionTable.PSVersion.Major -ge 6) {
        $ColorCodes = @{
            "Red" = "`e[31m"
            "Green" = "`e[32m"
            "Yellow" = "`e[33m"
            "Blue" = "`e[34m"
            "Magenta" = "`e[35m"
            "Cyan" = "`e[36m"
            "White" = "`e[37m"
            "Reset" = "`e[0m"
        }
        Write-Host "$($ColorCodes[$Color])$Message$($ColorCodes['Reset'])"
    } else {
        Write-Host $Message
    }
}

function Test-CrossPlatformPath {
    <#
    .SYNOPSIS
        Validates and normalizes file paths across different operating systems
    #>
    param([string]$Path)
    
    try {
        # Convert to absolute path and normalize separators
        $ResolvedPath = Resolve-Path $Path -ErrorAction Stop
        return $ResolvedPath.Path
    }
    catch {
        Write-ColorOutput "Error: Invalid path '$Path'" -Color "Red"
        return $null
    }
}

function Invoke-TMDBRequest {
    <#
    .SYNOPSIS
        Makes HTTP requests to TMDB API with proper error handling and rate limiting
    #>
    param(
        [string]$Endpoint,
        [hashtable]$Parameters = @{}
    )
    
    # Rate limiting - simple implementation
    Start-Sleep -Milliseconds 100
    
    try {
        # Build query string from parameters
        $QueryString = ""
        if ($Parameters.Count -gt 0) {
            $QueryParts = @()
            foreach ($Key in $Parameters.Keys) {
                $QueryParts += "$Key=$([System.Web.HttpUtility]::UrlEncode($Parameters[$Key]))"
            }
            $QueryString = "?" + ($QueryParts -join "&")
        }
        
        $Uri = "$($TMDBConfig.BaseURL)$Endpoint$QueryString"
        
        # Create headers with authentication
        $Headers = @{
            "Authorization" = "Bearer $($TMDBConfig.ReadAccessToken)"
            "Content-Type" = "application/json"
        }
        
        Write-ColorOutput "Making TMDB API request: $Endpoint" -Color "Cyan"
        
        # Make the HTTP request
        $Response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get -TimeoutSec 30
        return $Response
    }
    catch {
        Write-ColorOutput "TMDB API Error: $($_.Exception.Message)" -Color "Red"
        return $null
    }
}

function Search-TVShowOnTMDB {
    <#
    .SYNOPSIS
        Searches for TV show information on TMDB using series name
    #>
    param([string]$SeriesName)
    
    if (-not $UseTMDB) {
        return $null
    }
    
    Write-ColorOutput "Searching TMDB for: $SeriesName" -Color "Blue"
    
    # Clean up series name for search
    $CleanSeriesName = $SeriesName -replace '[_\-\.]', ' '
    $CleanSeriesName = $CleanSeriesName -replace '\s+', ' '
    $CleanSeriesName = $CleanSeriesName.Trim()
    
    $SearchParams = @{
        "api_key" = $TMDBConfig.APIKey
        "query" = $CleanSeriesName
    }
    
    $SearchResult = Invoke-TMDBRequest -Endpoint "/search/tv" -Parameters $SearchParams
    
    if ($SearchResult -and $SearchResult.results -and $SearchResult.results.Count -gt 0) {
        $TopResult = $SearchResult.results[0]
        Write-ColorOutput "Found TV Show: $($TopResult.name) (ID: $($TopResult.id))" -Color "Green"
        return $TopResult
    } else {
        Write-ColorOutput "No TMDB results found for: $SeriesName" -Color "Yellow"
        return $null
    }
}

function Get-SeasonEpisodesFromTMDB {
    <#
    .SYNOPSIS
        Retrieves episode information for a specific season from TMDB
    #>
    param(
        [int]$ShowID,
        [int]$SeasonNumber
    )
    
    if (-not $UseTMDB) {
        return $null
    }
    
    Write-ColorOutput "Getting episodes for Season $SeasonNumber" -Color "Blue"
    
    $SeasonParams = @{
        "api_key" = $TMDBConfig.APIKey
    }
    
    $SeasonData = Invoke-TMDBRequest -Endpoint "/tv/$ShowID/season/$SeasonNumber" -Parameters $SeasonParams
    
    if ($SeasonData -and $SeasonData.episodes) {
        Write-ColorOutput "Found $($SeasonData.episodes.Count) episodes for Season $SeasonNumber" -Color "Green"
        return $SeasonData.episodes
    } else {
        Write-ColorOutput "No episode data found for Season $SeasonNumber" -Color "Yellow"
        return $null
    }
}

function Parse-DirectoryName {
    <#
    .SYNOPSIS
        Parses TV show directory name to extract series name, season, and disc information
        Expected format: TV_SHOW_NAME_S1_D1
    #>
    param([string]$DirectoryName)
    
    Write-ColorOutput "Parsing directory: $DirectoryName" -Color "Cyan"
    
    # Regex pattern to match: SERIES_NAME_S#_D#
    $Pattern = '^(.+?)_S(\d+)_D(\d+)$'
    
    if ($DirectoryName -match $Pattern) {
        $ParsedInfo = @{
            SeriesName = $Matches[1] -replace '_', ' '  # Replace underscores with spaces
            SeasonNumber = [int]$Matches[2]
            DiscNumber = [int]$Matches[3]
            OriginalName = $DirectoryName
        }
        
        Write-ColorOutput "Parsed - Series: '$($ParsedInfo.SeriesName)', Season: $($ParsedInfo.SeasonNumber), Disc: $($ParsedInfo.DiscNumber)" -Color "Green"
        return $ParsedInfo
    } else {
        Write-ColorOutput "Directory name doesn't match expected pattern: $DirectoryName" -Color "Yellow"
        return $null
    }
}

function Get-VideoFiles {
    <#
    .SYNOPSIS
        Gets all video files from a directory, sorted by name for consistent ordering
    #>
    param([string]$DirectoryPath)
    
    $VideoFiles = Get-ChildItem -Path $DirectoryPath -File | 
                  Where-Object { $_.Extension -in $Config.SupportedExtensions -and $_.Length -gt $Config.MinFileSize } |
                  Sort-Object Name
    
    Write-ColorOutput "Found $($VideoFiles.Count) video files in directory" -Color "Cyan"
    
    return $VideoFiles
}

function Generate-EpisodeName {
    <#
    .SYNOPSIS
        Generates proper episode filename using Jellyfin naming convention
    #>
    param(
        [string]$SeriesName,
        [int]$SeasonNumber,
        [int]$EpisodeNumber,
        [string]$Extension,
        [string]$EpisodeTitle = $null
    )
    
    # Format season and episode with leading zeros
    $SeasonFormatted = "S{0:D2}" -f $SeasonNumber
    $EpisodeFormatted = "E{0:D2}" -f $EpisodeNumber
    
    # Clean series name for filename (remove invalid characters)
    $CleanSeriesName = $SeriesName -replace '[<>:"/\\|?*]', ''
    $CleanSeriesName = $CleanSeriesName.Trim()
    
    # Build filename: "Series Name S01E01.mkv"
    $FileName = "$CleanSeriesName $SeasonFormatted$EpisodeFormatted$Extension"
    
    Write-ColorOutput "Generated filename: $FileName" -Color "Green"
    
    return $FileName
}

function Rename-EpisodeFile {
    <#
    .SYNOPSIS
        Renames a single episode file with proper error handling
    #>
    param(
        [System.IO.FileInfo]$SourceFile,
        [string]$NewFileName,
        [bool]$IsDryRun = $false
    )
    
    $DestinationPath = Join-Path $SourceFile.Directory.FullName $NewFileName
    
    # Check if destination file already exists
    if (Test-Path $DestinationPath) {
        Write-ColorOutput "Warning: Destination file already exists: $NewFileName" -Color "Yellow"
        return $false
    }
    
    if ($IsDryRun) {
        Write-ColorOutput "[DRY RUN] Would rename: '$($SourceFile.Name)' → '$NewFileName'" -Color "Magenta"
        return $true
    } else {
        try {
            Rename-Item -Path $SourceFile.FullName -NewName $NewFileName -Force
            Write-ColorOutput "✓ Renamed: '$($SourceFile.Name)' → '$NewFileName'" -Color "Green"
            return $true
        }
        catch {
            Write-ColorOutput "✗ Failed to rename '$($SourceFile.Name)': $($_.Exception.Message)" -Color "Red"
            return $false
        }
    }
}

# =============================================================================
# MAIN PROCESSING FUNCTION
# =============================================================================

function Calculate-EpisodeNumbers {
    <#
    .SYNOPSIS
        Calculates the starting episode number based on disc number by counting episodes in previous discs
        This handles varying numbers of episodes per disc dynamically
    #>
    param(
        [int]$DiscNumber,
        [int]$SeasonNumber,
        [string]$SeriesName,
        [string]$BasePath
    )
    
    Write-ColorOutput "Calculating episode numbers for Disc $DiscNumber..." -Color "Blue"
    
    # If this is disc 1, episodes start at 1
    if ($DiscNumber -eq 1) {
        Write-ColorOutput "Disc 1 - Episodes start at 1" -Color "Blue"
        return 1
    }
    
    # For discs 2+, count episodes in all previous discs
    $TotalPreviousEpisodes = 0
    
    # Look for all previous disc directories for this season
    for ($PrevDisc = 1; $PrevDisc -lt $DiscNumber; $PrevDisc++) {
        # Construct the expected directory name for previous disc
        $PrevDiscDirName = "$($SeriesName -replace ' ', '_')_S$SeasonNumber" + "_D$PrevDisc"
        $PrevDiscPath = Join-Path $BasePath $PrevDiscDirName
        
        Write-ColorOutput "Checking previous disc directory: $PrevDiscDirName" -Color "Cyan"
        
        if (Test-Path $PrevDiscPath) {
            # Count video files in previous disc
            $PrevDiscFiles = Get-VideoFiles -DirectoryPath $PrevDiscPath
            $PrevDiscEpisodeCount = $PrevDiscFiles.Count
            $TotalPreviousEpisodes += $PrevDiscEpisodeCount
            
            Write-ColorOutput "Disc $PrevDisc contains $PrevDiscEpisodeCount episodes" -Color "Green"
        } else {
            Write-ColorOutput "Warning: Previous disc directory not found: $PrevDiscDirName" -Color "Yellow"
            Write-ColorOutput "Assuming 0 episodes for missing disc $PrevDisc" -Color "Yellow"
        }
    }
    
    $StartingEpisode = $TotalPreviousEpisodes + 1
    Write-ColorOutput "Total episodes in previous discs: $TotalPreviousEpisodes" -Color "Blue"
    Write-ColorOutput "Disc $DiscNumber episodes will start at: $StartingEpisode" -Color "Green"
    
    return $StartingEpisode
}

function Process-TVShowDirectory {
    <#
    .SYNOPSIS
        Main function that processes a single TV show directory
    #>
    param([string]$DirectoryPath)
    
    $DirectoryName = Split-Path $DirectoryPath -Leaf
    Write-ColorOutput "`n" + "="*60 -Color "White"
    Write-ColorOutput "Processing Directory: $DirectoryName" -Color "Yellow"
    Write-ColorOutput "="*60 -Color "White"
    
    # Parse directory name to extract show information
    $ShowInfo = Parse-DirectoryName -DirectoryName $DirectoryName
    if (-not $ShowInfo) {
        Write-ColorOutput "Skipping directory due to naming format mismatch" -Color "Red"
        return
    }
    
    # Get video files from directory
    $VideoFiles = Get-VideoFiles -DirectoryPath $DirectoryPath
    if ($VideoFiles.Count -eq 0) {
        Write-ColorOutput "No video files found in directory" -Color "Yellow"
        return
    }
    
    # Calculate starting episode number based on disc number
    # This dynamically counts episodes in previous discs to handle varying episode counts
    $StartingEpisodeNumber = Calculate-EpisodeNumbers -DiscNumber $ShowInfo.DiscNumber -SeasonNumber $ShowInfo.SeasonNumber -SeriesName $ShowInfo.SeriesName -BasePath (Split-Path $DirectoryPath -Parent)
    
    Write-ColorOutput "This disc contains $($VideoFiles.Count) episodes, starting from episode $StartingEpisodeNumber" -Color "Cyan"
    
    # Get TMDB information if enabled
    $TMDBShow = $null
    $TMDBEpisodes = $null
    
    if ($UseTMDB) {
        $TMDBShow = Search-TVShowOnTMDB -SeriesName $ShowInfo.SeriesName
        if ($TMDBShow) {
            $TMDBEpisodes = Get-SeasonEpisodesFromTMDB -ShowID $TMDBShow.id -SeasonNumber $ShowInfo.SeasonNumber
        }
    }
    
    # Determine series name to use (TMDB name takes precedence)
    $FinalSeriesName = if ($TMDBShow) { $TMDBShow.name } else { $ShowInfo.SeriesName }
    
    # Process each video file
    $FileIndex = 0
    $SuccessCount = 0
    $FailureCount = 0
    
    foreach ($VideoFile in $VideoFiles) {
        Write-ColorOutput "`nProcessing file: $($VideoFile.Name)" -Color "Cyan"
        
        # Calculate actual episode number based on disc and file position
        $EpisodeNumber = $StartingEpisodeNumber + $FileIndex
        
        Write-ColorOutput "Calculated episode number: $EpisodeNumber (Disc $($ShowInfo.DiscNumber), File $($FileIndex + 1))" -Color "Blue"
        
        # Get episode title from TMDB if available
        $EpisodeTitle = $null
        if ($TMDBEpisodes -and $EpisodeNumber -le $TMDBEpisodes.Count) {
            $TMDBEpisode = $TMDBEpisodes[$EpisodeNumber - 1]
            $EpisodeTitle = $TMDBEpisode.name
            Write-ColorOutput "TMDB Episode Title: $EpisodeTitle" -Color "Blue"
        }
        
        # Generate new filename
        $NewFileName = Generate-EpisodeName -SeriesName $FinalSeriesName -SeasonNumber $ShowInfo.SeasonNumber -EpisodeNumber $EpisodeNumber -Extension $VideoFile.Extension -EpisodeTitle $EpisodeTitle
        
        # Rename the file
        $RenameSuccess = Rename-EpisodeFile -SourceFile $VideoFile -NewFileName $NewFileName -IsDryRun $DryRun
        
        if ($RenameSuccess) {
            $SuccessCount++
        } else {
            $FailureCount++
        }
        
        $FileIndex++
    }
    
    # Summary for this directory
    Write-ColorOutput "`nDirectory Summary:" -Color "White"
    Write-ColorOutput "✓ Successfully processed: $SuccessCount files" -Color "Green"
    if ($FailureCount -gt 0) {
        Write-ColorOutput "✗ Failed to process: $FailureCount files" -Color "Red"
    }
}

# =============================================================================
# MAIN SCRIPT EXECUTION
# =============================================================================

function Main {
    <#
    .SYNOPSIS
        Main script entry point with initialization and cleanup
    #>
    
    Write-ColorOutput @"
╔══════════════════════════════════════════════════════════════════════════════╗
║                        TV Show Episode Renamer                              ║
║                     PowerShell Core Cross-Platform                          ║
╚══════════════════════════════════════════════════════════════════════════════╝
"@ -Color "Cyan"
    
    # Validate base path
    $NormalizedBasePath = Test-CrossPlatformPath -Path $BasePath
    if (-not $NormalizedBasePath) {
        Write-ColorOutput "Script terminated due to invalid base path." -Color "Red"
        exit 1
    }
    
    Write-ColorOutput "Base Path: $NormalizedBasePath" -Color "White"
    Write-ColorOutput "TMDB Integration: $(if ($UseTMDB) { 'Enabled' } else { 'Disabled' })" -Color "White"
    Write-ColorOutput "Mode: $(if ($DryRun) { 'DRY RUN (Preview Only)' } else { 'LIVE (Files will be renamed)' })" -Color "White"
    
    if ($DryRun) {
        Write-ColorOutput "`nDRY RUN MODE: No files will actually be renamed!" -Color "Yellow"
    }
    
    # Get all subdirectories that match the expected pattern
    $TVShowDirectories = Get-ChildItem -Path $NormalizedBasePath -Directory | 
                         Where-Object { $_.Name -match '^.+_S\d+_D\d+$' }
    
    if ($TVShowDirectories.Count -eq 0) {
        Write-ColorOutput "No TV show directories found matching the pattern: SERIES_NAME_S#_D#" -Color "Yellow"
        Write-ColorOutput "Please ensure your directories follow this naming convention." -Color "White"
        exit 0
    }
    
    Write-ColorOutput "`nFound $($TVShowDirectories.Count) TV show directories to process:" -Color "White"
    foreach ($Dir in $TVShowDirectories) {
        Write-ColorOutput "  • $($Dir.Name)" -Color "Cyan"
    }
    
    # Confirm before proceeding (unless it's a dry run)
    if (-not $DryRun) {
        Write-ColorOutput "`nThis will rename files in the above directories." -Color "Yellow"
        $Confirmation = Read-Host "Do you want to continue? (y/N)"
        if ($Confirmation -notmatch '^[Yy]$') {
            Write-ColorOutput "Operation cancelled by user." -Color "Yellow"
            exit 0
        }
    }
    
    # Process each TV show directory
    $TotalSuccess = 0
    $TotalFailures = 0
    
    foreach ($Directory in $TVShowDirectories) {
        try {
            Process-TVShowDirectory -DirectoryPath $Directory.FullName
        }
        catch {
            Write-ColorOutput "Critical error processing directory '$($Directory.Name)': $($_.Exception.Message)" -Color "Red"
            $TotalFailures++
        }
    }
    
   
   
# Move files section
    if($MoveToFolder){
        if($DryRun){ 
            Write-Output "Move Option Selected with DryRun option"
            
            $PathTest = Test-Path -Path $MoveToFolder
            if($PathTest -eq $true){
                Write-Output "Directory is present"
                Write-Output "Doing Whatif move"
                get-childitem -Path $BasePath -Recurse -Include "*.mkv", "*.mp4", "*.avi", "*.mov" | 
                     ForEach-Object { Move-Item -Path $_.FullName -Destination $MoveToFolder -whatif}
                
            }
            else{
                Write-Output "Directory is not present"
                New-Item -Path $MoveToFolder -ItemType Directory -WhatIf
                Write-Output "Doing Whatif move"
                get-childitem -Path $BasePath -Recurse -Include "*.mkv", "*.mp4", "*.avi", "*.mov" | 
                     ForEach-Object { Move-Item -Path $_.FullName -Destination $MoveToFolder -whatif}
            }
        }
        else{
            # dry run not selected do actual moving
            Write-Output "Move Option Selected - Performing actual move"
            
            $PathTest = Test-Path -Path $MoveToFolder
            if($PathTest -eq $true){
                Write-Output "Directory is present"
                Write-Output "Doing move"
                get-childitem -Path $BasePath -Recurse -Include "*.mkv", "*.mp4", "*.avi", "*.mov" | 
                     ForEach-Object { Move-Item -Path $_.FullName -Destination $MoveToFolder}
            }
            else{
                Write-Output "Directory is not present, creating it"
                New-Item -Path $MoveToFolder -ItemType Directory
                Write-Output "Doing move"
                get-childitem -Path $BasePath -Recurse -Include "*.mkv", "*.mp4", "*.avi", "*.mov" | 
                     ForEach-Object { Move-Item -Path $_.FullName -Destination $MoveToFolder}
            }
        }
    }
    else {
        Write-Output "Move Option not Selected"
    }

    # Final summary
    Write-ColorOutput "`n" + "="*80 -Color "White"
    Write-ColorOutput "SCRIPT EXECUTION COMPLETE" -Color "White"
    Write-ColorOutput "="*80 -Color "White"
    Write-ColorOutput "Processed $($TVShowDirectories.Count) directories" -Color "White"
    
    if ($DryRun) {
        Write-ColorOutput "DRY RUN completed - no files were actually renamed." -Color "Magenta"
    } else {
        Write-ColorOutput "File renaming operation completed." -Color "Green"
    }
    
    Write-ColorOutput "`nFor support or issues, check the TMDB API documentation:" -Color "White"
    Write-ColorOutput "https://developer.themoviedb.org/reference/intro/getting-started" -Color "Cyan"
} # <-- Only ONE closing brace needed here for the Main function

# =============================================================================
# SCRIPT ENTRY POINT
# =============================================================================

# Only run main function if script is executed directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    Main
}