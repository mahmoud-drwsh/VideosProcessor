<#
.SYNOPSIS
    Video processing script. Run as a script: .\main.ps1 (do not dot-source).
#>
param(
    [switch]$debugProgram
)
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

function Get-AudioEncodeArgs($inputPath, $outputPath, $metaTitle, $albumArtist) {
    return "-y -i `"$inputPath`" -vn -c:a libopus -b:a 17k -map_metadata -1 -metadata title=`"$metaTitle`" -metadata album_artist=`"$albumArtist`" -metadata:s:a:0 title=`"$metaTitle`" `"$outputPath`""
}
function Get-H264480EncodeArgs($inputPath, $outputPath, $metaTitle) {
    # H.264 480p via libx264; CRF 23 = good quality, preset medium = balance of speed/size
    return "-y -i `"$inputPath`" -vf `"scale=-2:480`" -map_metadata -1 -c:v libx264 -preset medium -crf 23 -c:a copy -c:s copy -map_chapters 0 -metadata:s:a:0 title=`"$metaTitle`" `"$outputPath`""
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

# Escape text for safe use in HTML (e.g. value="" attribute).
function Escape-HtmlForDialog {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
}

# HTML dialog for title and artist (RTL, editable). Returns [PSCustomObject] with CleanTitle, AlbumArtist on OK; $null on Cancel.
function Show-TitleConfirmationDialog {
    param(
        [string]$TitleText,
        [string]$ArtistText
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $callbackSource = @'
using System;
using System.Runtime.InteropServices;
[ComVisible(true)]
public class DialogResultCallback {
    public static bool? Result;
    public static string TitleResult;
    public static string ArtistResult;
    public void Confirm(bool ok, object titleObj, object artistObj) {
        Result = ok;
        TitleResult = titleObj != null ? titleObj.ToString() : "";
        ArtistResult = artistObj != null ? artistObj.ToString() : "";
    }
}
'@
    try { Add-Type -TypeDefinition $callbackSource } catch { }
    [DialogResultCallback]::Result = $null
    [DialogResultCallback]::TitleResult = ""
    [DialogResultCallback]::ArtistResult = ""

    $titleEscaped = Escape-HtmlForDialog -Text $TitleText
    $artistEscaped = Escape-HtmlForDialog -Text $ArtistText

    $html = @"
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="UTF-8">
<style>
  body { font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 14px; padding: 24px; margin: 0; }
  .row { margin-bottom: 16px; }
  label { display: inline-block; width: 90px; font-weight: bold; }
  input[type="text"] { width: 280px; padding: 8px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
  .buttons { margin-top: 28px; padding-top: 16px; }
  .buttons button { padding: 12px 28px; margin-left: 14px; cursor: pointer; font-size: 14px; border-radius: 6px; border: 1px solid #ccc; }
  .buttons button:first-of-type { margin-left: 0; background: #0078d4; color: #fff; border-color: #0078d4; }
  .buttons button:hover { opacity: 0.9; }
</style>
</head>
<body>
  <div class="row"><label>العنوان:</label><input type="text" id="titleInput" value="$titleEscaped" dir="rtl"></div>
  <div class="row"><label>الفنان:</label><input type="text" id="artistInput" value="$artistEscaped" dir="rtl"></div>
  <div class="buttons">
    <button type="button" onclick="window.external.Confirm(true, document.getElementById('titleInput').value, document.getElementById('artistInput').value)">موافق</button>
    <button type="button" onclick="window.external.Confirm(false, '', '')">إلغاء</button>
  </div>
</body>
</html>
"@

    $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "title_confirm_" + [Guid]::NewGuid().ToString("N") + ".html")
    [System.IO.File]::WriteAllText($tempFile, $html, [System.Text.Encoding]::UTF8)
    try {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Confirm Title"
        $form.Size = New-Object System.Drawing.Size(520, 260)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"

        $browser = New-Object System.Windows.Forms.WebBrowser
        $browser.Dock = [System.Windows.Forms.DockStyle]::Fill
        $browser.ScriptErrorsSuppressed = $true
        $browser.IsWebBrowserContextMenuEnabled = $false
        $form.Controls.Add($browser)

        $callback = New-Object DialogResultCallback
        $browser.ObjectForScripting = $callback

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 150
        $timer.Add_Tick({
            if ([DialogResultCallback]::Result -ne $null) {
                $timer.Stop()
                $form.Close()
            }
        })
        $form.Add_Shown({ $timer.Start() })
        $form.Add_FormClosed({ $timer.Stop() })

        $fileUri = [System.Uri]::new("file:///" + $tempFile.Replace("\", "/").Replace(" ", "%20"))
        $browser.Navigate($fileUri.AbsoluteUri)
        $null = $form.ShowDialog()

        if ([DialogResultCallback]::Result -eq $true) {
            $t = [DialogResultCallback]::TitleResult; if (-not $t) { $t = "" }
            $a = [DialogResultCallback]::ArtistResult; if (-not $a) { $a = "" }
            return [PSCustomObject]@{
                CleanTitle  = $t.Trim()
                AlbumArtist = $a.Trim()
            }
        }
        return $null
    } finally {
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
    }
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
            $result = Show-TitleConfirmationDialog -TitleText $cleanTitle -ArtistText $albumArtist
            if (-not $result) {
                exit 0
            }
            $cleanTitle = $result.CleanTitle
            $albumArtist = $result.AlbumArtist
            break # User confirmed (possibly edited), proceed
        }
    }
    Write-Host "Waiting for 2 lines in $titleFilePath..." -ForegroundColor Yellow
    Start-Sleep -Milliseconds 500
}

# --- Step 2: Select Video (file dialog, opens in user's Videos folder) ---
$videosFolder = Join-Path $env:USERPROFILE "Videos"
if (-not (Test-Path $videosFolder)) { $videosFolder = $videoSourceDir }
$selectedPath = Show-FileSelector -InitialDirectory $videosFolder
if (-not $selectedPath) { exit 0 }
$latestVideo = Get-Item $selectedPath
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

# 5b. Re-encode Video (H.264 libx264 480p)
if (-not (Test-Path $videoOutPath)) {
    Write-Host "  > Re-encoding Video (H.264 libx264 480p)..."
    $metaTitle = "الشيخ محمد فواز النمر"
    $h264Args = Get-H264480EncodeArgs -inputPath $latestVideo.FullName -outputPath $videoOutPath -metaTitle $metaTitle
    Write-Host "  > ffmpeg H.264 480p args:" -ForegroundColor Magenta
    Write-Host "    ffmpeg $h264Args"
    $null = Start-Process -FilePath "ffmpeg" -ArgumentList $h264Args -Wait -NoNewWindow -PassThru
} else {
    Write-Host "  > Compressed video already exists. Skipping." -ForegroundColor DarkGray
}

# --- Step 6: Copy to Destinations (NO OVERWRITE) ---
if (-not $debugProgram) {

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

} else {
    Write-Host "Copy step skipped (debug mode)." -ForegroundColor Cyan
}

# --- Final Step: Exit ---
Write-Host "All tasks completed successfully. Exiting." -ForegroundColor Green
Start-Sleep -Seconds 5