# TEST comment
# Set encoding to UTF8 to handle Arabic text correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- Configuration ---
$videoSourceDir = "C:\Users\LEGION\Videos"
$titleFilePath  = "C:\Users\LEGION\Desktop\title.txt"
$workDir        = "C:\Users\LEGION\Desktop\إخراج"

# Destination Paths
$destOrigVideo  = "D:\01 - الفيديو"
$destCompVideo  = "D:\02 - مضغوط"
$destAudio      = "D:\03 - صوت"

# Ensure base directories exist
if (-not (Test-Path $workDir)) { New-Item -ItemType Directory -Path $workDir -Force | Out-Null }
if (-not (Test-Path $destOrigVideo)) { New-Item -ItemType Directory -Path $destOrigVideo -Force | Out-Null }
if (-not (Test-Path $destCompVideo)) { New-Item -ItemType Directory -Path $destCompVideo -Force | Out-Null }
if (-not (Test-Path $destAudio)) { New-Item -ItemType Directory -Path $destAudio -Force | Out-Null }

# --- FFmpeg Command Builders ---

function Get-RcLookaheadFrames($fps) {
    $lookahead = if ($fps) { [int][Math]::Round($fps * 2) } else { 64 }
    return [Math]::Max(10, [Math]::Min($lookahead, 120))
}

function Get-AudioEncodeArgs($inputPath, $outputPath, $metaTitle, $albumArtist) {
    return "-y -i `"$inputPath`" -vn -c:a libopus -b:a 17k -map_metadata -1 -metadata title=`"$metaTitle`" -metadata album_artist=`"$albumArtist`" -metadata:s:a:0 title=`"$metaTitle`" `"$outputPath`""
}
function Get-H264480EncodeArgs($inputPath, $outputPath, $lookaheadFrames, $metaTitle) {
    # H.264 480p; CQ 29 keeps size reasonable
    return "-y -i `"$inputPath`" -vf `"scale=-2:480`" -map_metadata -1 -c:v h264_nvenc -preset p3 -tune hq -cq 26 -rc vbr -spatial-aq 1 -temporal-aq 1 -rc-lookahead $lookaheadFrames -b:v 0 -c:a copy -c:s copy -map_chapters 0 -metadata:s:a:0 title=`"$metaTitle`" `"$outputPath`""
}

# --- Helper Functions ---

# Helper to wrap text with RTL directional marks for proper Arabic display.
function Convert-ToRtl {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    $RLE = [char]0x202B  # Right-to-Left Embedding
    $PDF = [char]0x202C  # Pop Directional Formatting

    return "$RLE$Text$PDF"
}

function Show-FileSelector {
    param(
        [string]$InitialDirectory
    )

    # Load the required .NET assembly
    Add-Type -AssemblyName System.Windows.Forms

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select a Video File"
    $openFileDialog.Filter = "Video Files|*.mp4;*.mkv;*.mov;*.avi;*.flv;*.obs|All Files|*.*"

    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $openFileDialog.InitialDirectory = $InitialDirectory
    } else {
        $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    }

    # Optional: Allow selecting multiple files? (Set to $false for single file)
    $openFileDialog.Multiselect = $false

    # Show the dialog. If user clicks OK, return the path.
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }

    return $null
}

# Show RTL-friendly confirmation dialog for title and album artist. Returns $true if OK, $false if Cancel/closed.
function Show-TitleConfirmationDialog {
    param(
        [string]$TitleText,
        [string]$ArtistText
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Confirm Title"
    $form.Size = New-Object System.Drawing.Size(420, 200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    $form.RightToLeftLayout = $true
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $lblTitleCaption = New-Object System.Windows.Forms.Label
    $lblTitleCaption.Text = "العنوان:"
    $lblTitleCaption.Location = New-Object System.Drawing.Point(20, 20)
    $lblTitleCaption.AutoSize = $true
    $lblTitleCaption.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    $form.Controls.Add($lblTitleCaption)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = $TitleText
    $lblTitle.Location = New-Object System.Drawing.Point(20, 42)
    $lblTitle.Size = New-Object System.Drawing.Size(360, 40)
    $lblTitle.AutoSize = $false
    $lblTitle.MaximumSize = New-Object System.Drawing.Size(360, 0)
    $lblTitle.AutoEllipsis = $false
    $lblTitle.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblTitle)

    $lblArtistCaption = New-Object System.Windows.Forms.Label
    $lblArtistCaption.Text = "الفنان:"
    $lblArtistCaption.Location = New-Object System.Drawing.Point(20, 88)
    $lblArtistCaption.AutoSize = $true
    $lblArtistCaption.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    $form.Controls.Add($lblArtistCaption)

    $lblArtist = New-Object System.Windows.Forms.Label
    $lblArtist.Text = $ArtistText
    $lblArtist.Location = New-Object System.Drawing.Point(20, 108)
    $lblArtist.Size = New-Object System.Drawing.Size(360, 24)
    $lblArtist.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
    $form.Controls.Add($lblArtist)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "موافق"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Location = New-Object System.Drawing.Point(220, 140)
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "إلغاء"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Location = New-Object System.Drawing.Point(120, 140)
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    return ($result -eq [System.Windows.Forms.DialogResult]::OK)
}

# Get average FPS for the first video stream using ffprobe.
function Get-VideoFps($filePath) {
    # avg_frame_rate is typically like "30000/1001" or "30/1"
    $rate = & ffprobe -v error -select_streams v:0 -show_entries stream=avg_frame_rate -of default=noprint_wrappers=1:nokey=1 "$filePath" 2>$null
    if (-not $rate) { return $null }

    $rateStr = ($rate | Select-Object -First 1).Trim()
    if (-not $rateStr) { return $null }

    if ($rateStr -match '^(?<num>\d+)\s*/\s*(?<den>\d+)$') {
        $num = [double]$Matches['num']
        $den = [double]$Matches['den']
        if ($den -le 0) { return $null }
        return ($num / $den)
    }

    # Sometimes it's already a decimal
    try { return [double]$rateStr } catch { return $null }
}

# Check whether a file is locked (used to wait for OBS to finish writing).
function Test-IsFileLocked($filePath) {
    try {
        $stream = [System.IO.File]::Open($filePath, 'Open', 'Read', 'None')
        $stream.Close()
        return $false
    } catch {
        return $true
    }
}

# --- Main Workflow ---

# --- Step 1: Wait for Title File (Must have 2 lines) ---
Write-Host "Checking Title file for 2 lines..." -ForegroundColor Cyan
while ($true) {
    if (Test-Path $titleFilePath) {
        # Read file as an array of lines
        $lines = Get-Content $titleFilePath -Encoding UTF8
        # Filter out empty lines just in case
        $validLines = $lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        if ($validLines.Count -ge 2) {
            $cleanTitle  = $validLines[0].Trim() # Line 1: Filename part
            $albumArtist = $validLines[1].Trim() # Line 2: Metadata Artist
            Write-Host "Title file OK. Showing confirmation..." -ForegroundColor Cyan
            $confirmed = Show-TitleConfirmationDialog -TitleText $cleanTitle -ArtistText $albumArtist
            if (-not $confirmed) {
                exit 0
            }
            break # Found both lines and user confirmed, proceed
        }
    }
    Write-Host "Waiting for 2 lines in $titleFilePath..." -ForegroundColor Yellow
    Start-Sleep -Milliseconds 500
}

# --- Step 2: Select Video (Recent or Browse) ---
Write-Host "Looking for recent videos..." -ForegroundColor Cyan
$recentVideos = @(Get-ChildItem -Path $videoSourceDir -File |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 3)

# Build the selection list
$menuItems = $recentVideos
$browseOptionIndex = $menuItems.Count + 1

# If no files found, the "Browse" option becomes option 1
if ($recentVideos.Count -eq 0) {
    Write-Host "No recent videos found in source folder." -ForegroundColor Yellow
    $browseOptionIndex = 1
}

$selectedFilePath = $null

do {
    Write-Host ""
    # Print Recent Files
    for ($i = 0; $i -lt $menuItems.Count; $i++) {
        $v = $menuItems[$i]
        $dt = $v.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        Write-Host "  $($i + 1). $($v.Name) ($dt)" -ForegroundColor White
    }

    # Print Browse Option
    Write-Host "  $browseOptionIndex. [Browse for a file...]" -ForegroundColor Cyan

    $choice = Read-Host "Select option (1-$browseOptionIndex)"

    # Check if user picked a recent file
    $idx = $null
    if ([int]::TryParse($choice, [ref]$idx)) {
        if ($idx -ge 1 -and $idx -le $menuItems.Count) {
            # User selected a recent file
            $latestVideo = $menuItems[$idx - 1]
            $selectedFilePath = $latestVideo.FullName
            break
        }
        elseif ($idx -eq $browseOptionIndex) {
            # User selected "Browse"
            Write-Host "Opening file dialog..." -ForegroundColor DarkGray
            $browsedFile = Show-FileSelector -InitialDirectory $videoSourceDir

            if ($browsedFile) {
                # Create a simple object to mimic the Get-ChildItem object so the rest of the script works
                $latestVideo = Get-Item $browsedFile
                $selectedFilePath = $latestVideo.FullName
                break
            } else {
                Write-Host "No file selected. Please try again." -ForegroundColor Red
            }
        }
        else {
            Write-Host "Invalid number." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Please enter a number." -ForegroundColor Yellow
    }

} while ($true)

Write-Host "Selected: $($latestVideo.Name)" -ForegroundColor Green

# --- Step 3: Wait for OBS to Stop Recording ---
Write-Host "Monitoring: $($latestVideo.Name)" -ForegroundColor Cyan
Write-Host "OBS is currently recording... Waiting." -ForegroundColor Yellow
while (Test-IsFileLocked -filePath $latestVideo.FullName) {
    Start-Sleep -Milliseconds 1000
}

Write-Host "Recording finished! Starting processing..." -ForegroundColor Green

# --- Step 4: Define Names and Paths ---
# Format: "YYYYMMDD Title" (Space separator)
$datePrefix = Get-Date -Format "yyyyMMdd " 
$baseName   = "$datePrefix$cleanTitle"

$audioOutPath   = Join-Path $workDir "$baseName.opus"
$videoOutPath   = Join-Path $workDir "$baseName.mp4"

# --- Step 5: Process Files (NO OVERWRITE) ---

# 5a. Extract Audio
if (-not (Test-Path $audioOutPath)) {
    Write-Host "  > Extracting Audio..."
    $metaTitle = "الشيخ محمد فواز النمر"
    $audioArgs = Get-AudioEncodeArgs -inputPath $latestVideo.FullName -outputPath $audioOutPath -metaTitle $metaTitle -albumArtist $albumArtist
    Write-Host "  > ffmpeg audio args:" -ForegroundColor Magenta
    Write-Host "    ffmpeg $audioArgs"
    $null = Start-Process -FilePath "ffmpeg" -ArgumentList $audioArgs -Wait -NoNewWindow -PassThru
} else {
    Write-Host "  > Audio already exists. Skipping." -ForegroundColor DarkGray
}

# 5b. Re-encode Video (H.264 NVENC 480p)
if (-not (Test-Path $videoOutPath)) {
    Write-Host "  > Re-encoding Video (H.264 NVENC 480p)..."
    $fps = Get-VideoFps -filePath $latestVideo.FullName
    $lookaheadFrames = Get-RcLookaheadFrames -fps $fps
    $metaTitle = "الشيخ محمد فواز النمر"
    $h264Args = Get-H264480EncodeArgs -inputPath $latestVideo.FullName -outputPath $videoOutPath -lookaheadFrames $lookaheadFrames -metaTitle $metaTitle
    Write-Host "  > ffmpeg H.264 480p args:" -ForegroundColor Magenta
    Write-Host "    ffmpeg $h264Args"
    $null = Start-Process -FilePath "ffmpeg" -ArgumentList $h264Args -Wait -NoNewWindow -PassThru
} else {
    Write-Host "  > Compressed video already exists. Skipping." -ForegroundColor DarkGray
}

# --- Step 6: Copy to Destinations (NO OVERWRITE) ---

# 6a. Copy Original to YYYY-MM-DD folder
$folderDateName = Get-Date -Format "yyyy-MM-dd"
$dailyFolder    = Join-Path $destOrigVideo $folderDateName

# Create daily folder if needed
if (-not (Test-Path $dailyFolder)) {
    New-Item -ItemType Directory -Path $dailyFolder -Force | Out-Null
}

Read-Host "Copying to destination... Press Enter to continue..." -ForegroundColor Yellow

$finalOrigPath = Join-Path $dailyFolder "$baseName$($latestVideo.Extension)"

if (-not (Test-Path $finalOrigPath)) {
    Write-Host "  > Copying Original to $dailyFolder..."
    Copy-Item -Path $latestVideo.FullName -Destination $finalOrigPath
} else {
    Write-Host "  > Original file already in destination. Skipping." -ForegroundColor DarkGray
}

# 6b. Copy Audio
$finalAudioPath = Join-Path $destAudio "$baseName.opus"
if (-not (Test-Path $finalAudioPath)) {
    Write-Host "  > Copying Audio..."
    Copy-Item -Path $audioOutPath -Destination $finalAudioPath
} else {
    Write-Host "  > Audio file already in destination. Skipping." -ForegroundColor DarkGray
}

# 6c. Copy Compressed Video
$finalCompPath = Join-Path $destCompVideo "$baseName.mp4"
if (-not (Test-Path $finalCompPath)) {
    Write-Host "  > Copying Compressed Video..."
    Copy-Item -Path $videoOutPath -Destination $finalCompPath
} else {
    Write-Host "  > Compressed video already in destination. Skipping." -ForegroundColor DarkGray
}

# --- Final Step: Exit ---
Write-Host "All tasks completed successfully. Exiting." -ForegroundColor Green
Start-Sleep -Seconds 30