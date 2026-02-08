<#
.SYNOPSIS
    Video processing script. Run as a script: .\main.ps1 (do not dot-source).
#>
param(
    [switch]$debugProgram,
    [switch]$skipAudio,
    [switch]$skipVideo,
    [string]$Title,
    [string]$Artist
)
# Set encoding to UTF8 to handle Arabic text correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- Global variables (paths and config used throughout the script) ---
$videoSourceDir = "C:\Users\LEGION\Videos"
$titleFilePath  = "C:\Users\LEGION\Desktop\title.txt"
$workDir        = "C:\Users\LEGION\Desktop\إخراج"
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
    # -map 0:a:0 = use only first audio stream (video is never decoded); -application voip = tuned for speech/lectures
    return "-y -i `"$inputPath`" -map 0:a:0 -vn -c:a libopus -application voip -b:a 18k -map_metadata -1 -metadata title=`"$metaTitle`" -metadata album_artist=`"$albumArtist`" -metadata:s:a:0 title=`"$metaTitle`" `"$outputPath`""
}
function Get-H265480EncodeArgs($inputPath, $outputPath, $metaTitle) {
    return "-y -fps_mode passthrough -hwaccel cuda -hwaccel_output_format cuda -i `"$inputPath`" -vf `"scale_cuda=-2:480`" -map_metadata -1 -c:v hevc_nvenc -preset p7 -c:a copy -map_chapters 0 -metadata:s:a:0 title=`"$metaTitle`" `"$outputPath`""
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
    }
    else {
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

# Sanitize a string for safe use in file/directory names (remove path-breaking characters).
function Get-SafeFileName {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sanitized = $Text
    foreach ($c in $invalid) {
        $sanitized = $sanitized.Replace([string]$c, " ")
    }
    return ($sanitized -replace '\s+', ' ').Trim()
}

# If title starts with a date (e.g. 2026/02/04, 2026\02\04, 2026-02-04), normalize to YYYYMMDD (no slashes).
function Normalize-LeadingDateInTitle {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    if ($Text -match '^\s*(\d{4})[-/\s\\]+(\d{1,2})[-/\s\\]+(\d{1,2})(\s.*|$)') {
        $y = $Matches[1]
        $m = $Matches[2].PadLeft(2, '0')
        $d = $Matches[3].PadLeft(2, '0')
        $rest = $Matches[4]
        return "$y$m$d$rest"
    }
    return $Text
}

# Escape text for safe use in HTML (e.g. value="" attribute).
function Escape-HtmlForDialog {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
}

# HTML dialog for title and artist (RTL, editable). Returns [PSCustomObject] with CleanTitle, AlbumArtist, SkipAudio, SkipVideo on OK; $null on Cancel.
function Show-TitleConfirmationDialog {
    param(
        [string]$TitleText,
        [string]$ArtistText,
        [switch]$InitialSkipAudio,
        [switch]$InitialSkipVideo
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $skipAudioChecked = if ($InitialSkipAudio) { " checked" } else { "" }
    $skipVideoChecked = if ($InitialSkipVideo) { " checked" } else { "" }

    $callbackSource = @'
using System;
using System.Runtime.InteropServices;
[ComVisible(true)]
public class DialogResultCallbackV2 {
    public static bool? Result;
    public static string TitleResult;
    public static string ArtistResult;
    public static bool SkipAudioResult;
    public static bool SkipVideoResult;
    public void Confirm(bool ok, object titleObj, object artistObj, bool skipAudio, bool skipVideo) {
        Result = ok;
        TitleResult = titleObj != null ? titleObj.ToString() : "";
        ArtistResult = artistObj != null ? artistObj.ToString() : "";
        SkipAudioResult = skipAudio;
        SkipVideoResult = skipVideo;
    }
}
'@
    try { Add-Type -TypeDefinition $callbackSource } catch { }
    [DialogResultCallbackV2]::Result = $null
    [DialogResultCallbackV2]::TitleResult = ""
    [DialogResultCallbackV2]::ArtistResult = ""
    [DialogResultCallbackV2]::SkipAudioResult = $false
    [DialogResultCallbackV2]::SkipVideoResult = $false

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
  input[type="checkbox"] { margin-left: 8px; vertical-align: middle; }
  .buttons { margin-top: 28px; padding-top: 16px; }
  .buttons button { padding: 12px 28px; margin-left: 14px; cursor: pointer; font-size: 14px; border-radius: 6px; border: 1px solid #ccc; }
  .buttons button:first-of-type { margin-left: 0; background: #0078d4; color: #fff; border-color: #0078d4; }
  .buttons button:hover { opacity: 0.9; }
</style>
</head>
<body>
  <div class="row"><label>العنوان:</label><input type="text" id="titleInput" value="$titleEscaped" dir="rtl"></div>
  <div class="row"><label>الفنان:</label><input type="text" id="artistInput" value="$artistEscaped" dir="rtl"></div>
  <div class="row"><label>تخطي الصوت:</label><input type="checkbox" id="skipAudio"$skipAudioChecked></div>
  <div class="row"><label>تخطي الفيديو:</label><input type="checkbox" id="skipVideo"$skipVideoChecked></div>
  <div class="buttons">
    <button type="button" onclick="window.external.Confirm(true, document.getElementById('titleInput').value, document.getElementById('artistInput').value, document.getElementById('skipAudio').checked, document.getElementById('skipVideo').checked)">موافق</button>
    <button type="button" onclick="window.external.Confirm(false, '', '', false, false)">إلغاء</button>
  </div>
<script>
window.initSelectOnFirstClick = function() {
  var firstTime = { titleInput: true, artistInput: true };
  function doSelect(el, id) {
    if (firstTime[id]) {
      firstTime[id] = false;
      el.select();
    }
  }
  function add(id) {
    var el = document.getElementById(id);
    if (!el) return;
    el.onfocus = function() { doSelect(el, id); };
    el.onmouseup = function() { doSelect(el, id); };
  }
  add('titleInput');
  add('artistInput');
};
</script>
</body>
</html>
"@

    $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "title_confirm_" + [Guid]::NewGuid().ToString("N") + ".html")
    [System.IO.File]::WriteAllText($tempFile, $html, [System.Text.Encoding]::UTF8)
    try {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Confirm Title"
        $form.Size = New-Object System.Drawing.Size(520, 320)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"

        $browser = New-Object System.Windows.Forms.WebBrowser
        $browser.Dock = [System.Windows.Forms.DockStyle]::Fill
        $browser.ScriptErrorsSuppressed = $true
        $browser.IsWebBrowserContextMenuEnabled = $false
        $form.Controls.Add($browser)

        $callback = New-Object DialogResultCallbackV2
        $browser.ObjectForScripting = $callback

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 150
        $timer.Add_Tick({
                if ([DialogResultCallbackV2]::Result -ne $null) {
                    $timer.Stop()
                    $form.Close()
                }
            })
        $form.Add_Shown({ $timer.Start() })
        $form.Add_FormClosed({ $timer.Stop() })

        $browser.Add_DocumentCompleted({
            if ($browser.Document -and $browser.ReadyState -eq [System.Windows.Forms.WebBrowserReadyState]::Complete) {
                try {
                    $browser.Document.InvokeScript("initSelectOnFirstClick")
                } catch { }
            }
        })

        $fileUri = [System.Uri]::new("file:///" + $tempFile.Replace("\", "/").Replace(" ", "%20"))
        $browser.Navigate($fileUri.AbsoluteUri)
        $null = $form.ShowDialog()

        if ([DialogResultCallbackV2]::Result -eq $true) {
            $t = [DialogResultCallbackV2]::TitleResult; if (-not $t) { $t = "" }
            $a = [DialogResultCallbackV2]::ArtistResult; if (-not $a) { $a = "" }
            return [PSCustomObject]@{
                CleanTitle  = $t.Trim()
                AlbumArtist = $a.Trim()
                SkipAudio   = [DialogResultCallbackV2]::SkipAudioResult
                SkipVideo   = [DialogResultCallbackV2]::SkipVideoResult
            }
        }
        return $null
    }
    finally {
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
    }
    catch {
        return $true
    }
}

# --- Main Workflow ---

# --- Step 1: Resolve Title and Artist (parameters or title.txt, no UI) ---
Write-Host "Resolving title and artist..." -ForegroundColor Cyan

$cleanTitle = $null
$albumArtist = $null

if ($Title -and $Artist) {
    # Prefer explicit parameters when provided
    $cleanTitle = $Title.Trim()
    $albumArtist = $Artist.Trim()
}
elseif (Test-Path $titleFilePath) {
    # Fallback: use existing title.txt file (2 non-empty lines: title, artist)
    $lines = Get-Content $titleFilePath -Encoding UTF8
    $validLines = $lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($validLines.Count -ge 2) {
        $cleanTitle = $validLines[0].Trim() # Line 1: Filename part
        $albumArtist = $validLines[1].Trim() # Line 2: Metadata Artist
    }
}

if (-not $cleanTitle -or -not $albumArtist) {
    Write-Host "Title and Artist must be provided via -Title/-Artist parameters or in title.txt (first two non-empty lines)." -ForegroundColor Red
    exit 1
}

# Normalize leading date to YYYYMMDD (remove slashes) for consistency
$cleanTitle = Normalize-LeadingDateInTitle -Text $cleanTitle

# Effective skip flags are now purely from CLI switches
$doSkipAudio = [bool]$skipAudio
$doSkipVideo = [bool]$skipVideo

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
# Sanitize title so backslashes etc. don't break paths.
$safeTitle = Get-SafeFileName -Text $cleanTitle
# If title already starts with a date (e.g. 20260204, 2026-02-04, 2026\02\04), don't add our own.
$titleStartsWithDate = $cleanTitle -match '^\s*(\d{8}|\d{4}[-/\s\\]?\d{2}[-/\s\\]?\d{2})'
$baseName = if ($titleStartsWithDate) { $safeTitle } else { (Get-Date -Format "yyyyMMdd ") + $safeTitle }

$audioOutPath = Join-Path $workDir "$baseName.opus"
$videoOutPath = Join-Path $workDir "$baseName.mp4"

# Paths for original copy (used before encoding)
$folderDateName  = Get-Date -Format "yyyy-MM-dd"
$dailyFolder     = Join-Path $destOrigVideo $folderDateName
$finalOrigPath   = Join-Path $dailyFolder "$baseName$($latestVideo.Extension)"
$workDirOrigPath = Join-Path $workDir "$baseName$($latestVideo.Extension)"

# --- Step 4b: Copy Original First (to output folder + final folder, before re-encode) ---
if (-not $debugProgram) {
    if (-not (Test-Path $dailyFolder)) {
        New-Item -ItemType Directory -Path $dailyFolder -Force | Out-Null
    }
    if (-not (Test-Path $workDirOrigPath)) {
        Write-Host "  > Copying Original to work folder ($workDir)..."
        Copy-Item -Path $latestVideo.FullName -Destination $workDirOrigPath
    } else {
        Write-Host "  > Original already in work folder. Skipping." -ForegroundColor DarkGray
    }
    if (-not (Test-Path $finalOrigPath)) {
        Write-Host "  > Copying Original to $dailyFolder..."
        Copy-Item -Path $latestVideo.FullName -Destination $finalOrigPath
    } else {
        Write-Host "  > Original already in final folder. Skipping." -ForegroundColor DarkGray
    }
}

# --- Step 5: Process Files (NO OVERWRITE) ---

# 5a. Extract Audio (skipped if user chose skip audio)
if ($doSkipAudio) {
    Write-Host "  > Audio step skipped (user choice)." -ForegroundColor DarkGray
} elseif (-not (Test-Path $audioOutPath)) {
    Write-Host "  > Extracting Audio..."
    $metaTitle = "الشيخ محمد فواز النمر"
    $audioArgs = Get-AudioEncodeArgs -inputPath $latestVideo.FullName -outputPath $audioOutPath -metaTitle $metaTitle -albumArtist $albumArtist
    Write-Host "  > ffmpeg audio args:" -ForegroundColor Magenta
    Write-Host "    ffmpeg $audioArgs"
    $null = Start-Process -FilePath "ffmpeg" -ArgumentList $audioArgs -Wait -NoNewWindow -PassThru
} else {
    Write-Host "  > Audio already exists. Skipping." -ForegroundColor DarkGray
}

# 5b. Re-encode Video (H.265 hevc_nvenc 480p) (skipped if user chose skip video)
if ($doSkipVideo) {
    Write-Host "  > Video re-encode skipped (user choice)." -ForegroundColor DarkGray
} elseif (-not (Test-Path $videoOutPath)) {
    Write-Host "  > Re-encoding Video (H.265 hevc_nvenc 480p)..."
    $metaTitle = "الشيخ محمد فواز النمر"
    $h265Args = Get-H265480EncodeArgs -inputPath $latestVideo.FullName -outputPath $videoOutPath -metaTitle $metaTitle
    Write-Host "  > ffmpeg H.265 480p args:" -ForegroundColor Magenta
    Write-Host "    ffmpeg $h265Args"
    $null = Start-Process -FilePath "ffmpeg" -ArgumentList $h265Args -Wait -NoNewWindow -PassThru
} else {
    Write-Host "  > Compressed video already exists. Skipping." -ForegroundColor DarkGray
}

# --- Step 6: Copy to Destinations (NO OVERWRITE) ---
if (-not $debugProgram) {

    # 6a. Original already copied in Step 4b (to work folder + final folder).

    # 6b. Copy Audio (skipped if user chose skip audio)
    if ($doSkipAudio) {
        Write-Host "  > Audio copy skipped (user choice)." -ForegroundColor DarkGray
    } else {
        $finalAudioPath = Join-Path $destAudio "$baseName.opus"
        if (-not (Test-Path $finalAudioPath)) {
            Write-Host "  > Copying Audio..."
            Copy-Item -Path $audioOutPath -Destination $finalAudioPath
        } else {
            Write-Host "  > Audio file already in destination. Skipping." -ForegroundColor DarkGray
        }
    }

    # 6c. Copy Compressed Video (skipped if user chose skip video)
    if ($doSkipVideo) {
        Write-Host "  > Compressed video copy skipped (user choice)." -ForegroundColor DarkGray
    } else {
        $finalCompPath = Join-Path $destCompVideo "$baseName.mp4"
        if (-not (Test-Path $finalCompPath)) {
            Write-Host "  > Copying Compressed Video..."
            Copy-Item -Path $videoOutPath -Destination $finalCompPath
        } else {
            Write-Host "  > Compressed video already in destination. Skipping." -ForegroundColor DarkGray
        }
    }

}
else {
    Write-Host "Copy step skipped (debug mode)." -ForegroundColor Cyan
}

# --- Final Step: Exit ---
Write-Host "All tasks completed successfully. Exiting." -ForegroundColor Green
Start-Sleep -Seconds 1