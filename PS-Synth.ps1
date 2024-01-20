# https://stackoverflow.com/questions/203890/creating-sine-or-square-wave-in-c-sharp
# https://www.electronics-tutorials.ws/waveforms/waveforms.html

# Add-Type -AssemblyName System.Windows.Forms # Required for keystroke capture
# Get key name: [System.Windows.Forms.Keys]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown').VirtualKeyCode

if (-not $MyInvocation.MyCommand.Path) { Write-Host "`t Save the code in ps1 file, then run it!" -ForegroundColor Red ; return $null }

$Host.UI.RawUI.WindowTitle = "PS-Synth"
Add-Type -AssemblyName WindowsBase # Required for keystroke  capture for new algorithm
Add-Type -AssemblyName PresentationCore # Required for MediaPlayer class & new capture algorithm

Function DrawKeyboard($first, $last) {
    # Note: The length of black key names must match the length of white ones. The size of both arrays should be exacly the maximum index loop below.
    # Note: Some of these names are fake, but they are required in order to assigned right names according to current index in the loop. 
    $WhiteName = @('Q', 'W','E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P', '[', ']', 'Z', 'X', 'C', 'V', 'B', 'N', 'M', '<', '>', '?', " ")
    $BlackName = @('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '-', '+', 'A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', ':', """")

    Write-Host (" " * ($last - $first) * 8 + " ") -BackgroundColor DarkGray

    for ($j=0; $j -lt 5; $j++){
        for ($i = $first; $i -lt $last; $i++){
            $NoBlack = @(0, 3, 7, 10, 14, 17, 21, 24, 28, 31, 35, 39, 43)
            if ($NoBlack -contains $i) { $IsBlack = $false }
            else { $IsBlack = $true }

            $BigWhite = @(1, 8, 15, 22, 29, 36, 43)
            if ($BigWhite -contains $i) { $white = 5 }
            else { $white = 4 }
    
            Write-Host (" " * [int](-not $IsBlack)) -BackgroundColor DarkGray -NoNewline
            Write-Host ("  $($BlackName[$i])  " * [int]($IsBlack)) -BackgroundColor Black -ForegroundColor Red -NoNewline
            Write-Host (" " * $white) -BackgroundColor White -NoNewline
        }
        Write-Host " " -BackgroundColor DarkGray
    }

    for ($i=0; $i -lt 7; $i++){
        for ($j = $first; $j -lt $last; $j++){
            Write-Host " " -BackgroundColor DarkGray -NoNewline
            Write-Host ("   $($WhiteName[$j])   ") -BackgroundColor White -ForegroundColor Black -NoNewline
        }
        Write-Host " " -BackgroundColor DarkGray
    }

    Write-Host (" " * ($last - $first) * 8 + " ") -BackgroundColor DarkGray
}

### Synth formula ###
# Sine Wave = Sin(2 * pi * f * t)

### Constants ###
# t depends on samples per second. In this case it's 44100. That means (1/44100) seconds interval per unit (step)
# The time increases by multiplying the step value: ((1/44100) * step)

Function Sine($frequency, $TimeStep, $samples, $volume, $fade) {
    $sine = 2 * [System.Math]::PI * $frequency * $TimeStep

    for ($step = 0; $step -lt $samples; $step++) {
        $fadeout = [System.Math]::Pow($fade, - $TimeStep * $step)
        
        ### Reduce sound presure (voltage) N times for the Nth overtone
        ### Reference http://www.sengpielaudio.com/calculator-levelchange.htm

        $wave = [System.Math]::Sin($sine * $step) + 
                [System.Math]::Sin(2 * $sine * $step) / 3.16227766016838 +  # [System.Math]::Pow(10, [System.Math]::Log(2, 2) / 2)
                [System.Math]::Sin(3 * $sine * $step) / 6.20127870518404 +  # [System.Math]::Pow(10, [System.Math]::Log(3, 2) / 2)
                [System.Math]::Sin(4 * $sine * $step) / 10 +                # [System.Math]::Pow(10, [System.Math]::Log(4, 2) / 2)
                [System.Math]::Sin(5 * $sine * $step) / 14.4865192364027 +  # [System.Math]::Pow(10, [System.Math]::Log(5, 2) / 2)
                [System.Math]::Sin(6 * $sine * $step) / 19.6101651138814    # [System.Math]::Pow(10, [System.Math]::Log(6, 2) / 2)
        
        $Script:mWriter.Write([int16]($volume * $wave * $wave * $wave * $fadeout))
    }
}

Function Octave($i) {
    $Script:tone += $i
    
    ### Verify whether the current octave goes out of boundary.
    if (($Script:tone + 37) -gt $Script:LastTone) { $Script:tone -= $i }
    if ($Script:tone -lt 0) { $Script:tone -= $i }

    $Script:key = 0..37 | ForEach-Object {$_ + $Script:tone}

    $host.UI.RawUI.CursorPosition = $Script:CursOct
    $oct = $Script:tone / 12

    $text = " Tones Range: From " + $OctNames[$oct] + " to " + $OctNames[$oct + 3] + " " * 10
    Write-Host "$text `n" -ForegroundColor Cyan
}

Function Play($i, $name) {
    $Script:CanPlay[$i] = $false
    $Script:MediaPlayer[$i].Position = [TimeSpan]::Zero
    $name += [string][System.Math]::Truncate($i / 12)
    
    $frequency = $Script:Frequencies[$i]
    if (-not $Script:LastFreq) { $Script:LastFreq = $frequency }

    ### Find the difference in cents between current and previous tone.
    $pitch = 1731.23404906676 * [System.Math]::Log($frequency / $Script:LastFreq)
    $Script:LastFreq = $frequency

    $host.UI.RawUI.CursorPosition = $Script:CursPlay

    Write-Host " Tone      : $name                    " -ForegroundColor Green
    Write-Host " Interval  : $pitch cents                    " -ForegroundColor Yellow
    Write-Host " Frequency : $frequency Hz                    " -ForegroundColor Red
}

Function Stop($i) {
    if (-not $Script:sustain) { $Script:MediaPlayer[$i].Position = [TimeSpan]::MaxValue }
    $Script:CanPlay[$i] = $true
}

Function Refresh() {
    Clear-Host

    if ($Host.Name -ne "ConsoleHost") { DrawKeyboard 0 23 }
    else {
        if ([System.Console]::WindowWidth -lt 105) { [System.Console]::WindowWidth = 105 }
        if ([System.Console]::WindowHeight -lt 41) { [System.Console]::WindowHeight = 41 }
        
        if ([System.Console]::WindowWidth -ge 185) { DrawKeyboard 0 23 }
        else {
            DrawKeyboard 0 10
            Write-Host ""
            DrawKeyboard 10 23
        }
    }

    Write-Host "`n To show up 1 row keyboard lower the font size, maximize the console and refresh. " -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host " Esc" -BackgroundColor DarkGray -ForegroundColor Yellow -NoNewline
    Write-Host " - Refresh the interface " -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host " CapsLock" -BackgroundColor DarkGray -ForegroundColor Yellow -NoNewline
    Write-Host " - Turn it on/off to sustain/release the sound " -BackgroundColor DarkGray -ForegroundColor Black
    Write-Host " ←" -BackgroundColor DarkGray -ForegroundColor Yellow -NoNewline
    Write-Host " - Octave down | " -BackgroundColor DarkGray -ForegroundColor Black -NoNewline
    Write-Host " →" -BackgroundColor DarkGray -ForegroundColor Yellow -NoNewline
    Write-Host " - Octave up `n" -BackgroundColor DarkGray -ForegroundColor Black

    $Script:CursOct = $host.UI.RawUI.CursorPosition ### Console coordinates to output current octaves range
    Octave 0 ### Don't shift octave, just show current playing range
    $Script:CursPlay = $host.UI.RawUI.CursorPosition ### Console coordinates to output playing note info
}

Clear-Host

Write-Host "`n Samples are generated in real-time depending on input values. Console-based graphic interface will show up when they are ready. `n" -ForegroundColor Cyan

$A4 = Try { [double](Read-Host -Prompt " Type A4 frequency between 415 and 466 or press ENTER to load default value 440.`n A4") } Catch {440}
if ($A4 -lt 415 -or $A4 -gt 466) { $A4 = 440 }

$Duration = Try { [double](Read-Host -Prompt "`n Type sound duration in seconds between 4 and 20 or press ENTER to load default value 6.`n Duration") } Catch {6}
if ($Duration -lt 4 -or $Duration -gt 20) { $Duration = 6 }

$tpo = Try { [System.Math]::Truncate((Read-Host -Prompt "`n Type number of tones per octave between 5 and 144 or press ENTER to load default value 12.`n TET/EDO")) } Catch {12}
if ($tpo -lt 5 -or $tpo -gt 144) { $tpo = 12 }

#### Build RIFF / WAV header ####

[int]$msDuration = $Duration * 1000
[int]$volume = 26000 # [uint16]::MaxValue -shr 2
[int]$formatChunkSize = 16;
[int]$headerSize = 8;
[int16]$formatType = 1;
[int16]$tracks = 1;
[int]$BitRate = 44100;
[double]$TimeStep = 1 / $BitRate
[int16]$bitsPerSample = 16;
[int16]$frameSize = [System.Math]::Truncate($tracks * ($bitsPerSample + 7) / 8)
[int]$bytesPerSecond = $BitRate * $frameSize
[int]$waveSize = 4
[int]$Samples = $BitRate * $msDuration / 1000
[int]$dataChunkSize = $samples * $frameSize
[int]$fileSize = $waveSize + $headerSize + $formatChunkSize + $headerSize + $dataChunkSize

[byte[]]$RIFF = @()
$RIFF += [System.Text.Encoding]::ASCII.GetBytes("RIFF")
$RIFF += [System.BitConverter]::GetBytes($fileSize)
$RIFF += [System.Text.Encoding]::ASCII.GetBytes("WAVE")
$RIFF += [System.Text.Encoding]::ASCII.GetBytes("fmt ")
$RIFF += [System.BitConverter]::GetBytes($formatChunkSize)
$RIFF += [System.BitConverter]::GetBytes($formatType)
$RIFF += [System.BitConverter]::GetBytes($tracks)
$RIFF += [System.BitConverter]::GetBytes($BitRate)
$RIFF += [System.BitConverter]::GetBytes($bytesPerSecond)
$RIFF += [System.BitConverter]::GetBytes($frameSize)
$RIFF += [System.BitConverter]::GetBytes($bitsPerSample)
$RIFF += [System.Text.Encoding]::ASCII.GetBytes("data")
$RIFF += [System.BitConverter]::GetBytes($dataChunkSize)

#### End RIFF header ####

Write-Host "`n A4 = $A4 Hz. Duration = $msDuration miliseconds. Temperament: $tpo`-TET/EDO" -ForegroundColor Yellow
$Host.UI.RawUI.WindowTitle = "PS-Synth | A4 = $A4 Hz ; Duration = $msDuration ms ; $tpo`-TET/EDO"

$Am1 = $A4 / 32 # A-1 frequency
$Ratios = @()

for ($i = 0; $i -lt $tpo ; $i++) {
    $Ratios += [System.Math]::Pow(2, $i/$tpo)
}

$Frequencies = @()

for ($i = 0; $i -lt 10; $i++) { # Find frequencies within 10 octaves
    $om = [System.Math]::Pow(2, $i) # Octave multiplier

    foreach ($ratio in $Ratios) {
        $Frequencies += $Am1 * $ratio * $om
    }
}

# Index 0 is "A-1" (lower than sub-contra). We need to start the octave from "C0" (sub-contra).
# Last tone is C#, 4 steps from A

$Frequencies = $Frequencies[3..($Frequencies.Count - $tpo + 4)]

<#
    Fade out function: f(t) = e^(-t)
    Let's call "e" a fade out  coefficient. It should be between 1 and 3
    fade = 1 # Means no fade out
    fade = 2.71828 # (Euler's constant) is standard fade out
    If fade > 2.71828 the sound fades out quicker than usual.
    In terms of my experience value about 2, fades out the sound more naturally for grand piano
    
    Voltage level should drop down to 0.0005 at the end in order to have decent fade out
    If f(t) = 0005 then: 
        0.0005 = fade^(-t)
        fade = -t root of 0.0005
        fade = 0.0005^(-1/t)
#>

$fade = [System.Math]::Pow(0.0005, -1000 / $msDuration)

$ScriptPath = (Get-Item $MyInvocation.MyCommand.Path).DirectoryName
$SamplesPath = Join-Path -Path $ScriptPath -ChildPath "Samples\A4=$($A4)Hz;Duration=$($msDuration)ms;$($tpo)-TET"
$null = [System.IO.Directory]::CreateDirectory($SamplesPath)
$files = [System.IO.Directory]::GetFiles($SamplesPath, "id????=?????.??????????????Hz.wav", [System.IO.SearchOption]::TopDirectoryOnly) | Sort-Object

if ($files.Count -ge $Frequencies.Count) {
    Write-Host "`n`t Samples with the above settings are already present." -ForegroundColor Cyan
}

else {
    Write-Host "`n Processing $($Frequencies.Count) samples . . . " -NoNewline -ForegroundColor Yellow

    for ($i = 0; $i -lt $Frequencies.Count; $i++) {
        Write-Host "$($i + 1), " -ForegroundColor Cyan -NoNewline

        $mStream = New-Object IO.MemoryStream
        $mWriter = [System.IO.BinaryWriter]::new($mStream)
        $mWriter.Write($RIFF)

        Sine $Frequencies[$i] $TimeStep $Samples $volume $fade
        $mWriter.Close()

        $id = "{0:0000}" -f ($i + 1)
        $sid =  "{0:00000.00000000000000}" -f $Frequencies[$i]
        $sp = "$($SamplesPath)\id$id`=$sid`Hz.wav"

        $fStream = [System.IO.FileStream]::new($sp, [System.IO.FileMode]::Create)
        $fWriter = [System.IO.BinaryWriter]::new($fStream)
        $fWriter.Write($mStream.ToArray())

        $mStream.Close()
        $mWriter.Close()
        $fStream.Close()
        $fWriter.Close()
    }

    Write-Host "- Finished" -ForegroundColor Yellow
}

$files = [System.IO.Directory]::GetFiles($SamplesPath, "id????=?????.??????????????Hz.wav", [System.IO.SearchOption]::TopDirectoryOnly) | Sort-Object
Write-Host "`n Loading $($files.Count) samples . . . " -NoNewline -ForegroundColor Green
$MediaPlayer = @()

for($i = 0; $i -lt $files.Count; $i++) {
    $MediaPlayer += New-Object System.Windows.Media.Mediaplayer
    $MediaPlayer[$i].Open($files[$i])
    $MediaPlayer[$i].Position = [TimeSpan]::MaxValue
    $MediaPlayer[$i].Volume = 1
    $MediaPlayer[$i].Play()

    ### Wait until current sample is loaded. If too many files are loaded simultaneously the .NET class hungs and audio driver crashes. 
    while (-not $MediaPlayer[$i].HasAudio) { Start-Sleep -Milliseconds 50 }
    Write-Host "$($i + 1), " -NoNewline
}

# Number of octaves
$no = [System.Math]::Truncate($Frequencies.Count / 12) + 1
$OctNames = "C0 (Sub-contra)", "C1 (Contra)", "C2 (Great)", "C3 (Small)", "C4 (1st)", "C5 (2nd)", "C6 (3rd)"
for ($i = 7; $i -lt $no; $i++) { $OctNames += "C$i ($($i - 3)th)" }

# Start playing from an octave where the frequency is at least 120 Hz.
$tone = 0
while ($Frequencies[$tone] -lt 120) { $tone += 12 }
$LastTone = $Frequencies.Count - 1

Refresh

[bool[]]$CanPlay = $Frequencies | ForEach-Object { $true }
[bool[]]$Released = 1..5 | ForEach-Object { $true } # Determines whether functional keys are pressed or released.

while ($true) {
    $sustain = [bool][Windows.Input.Keyboard]::GetKeyStates([System.Windows.Input.Key]::CapsLock)

    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Escape) -and $Released[0])      { Refresh ; $Released[0] = $false }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Escape)   -and -not $Released[0]) { $Released[0] = $true }

    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Left) -and $Released[1])        { $Released[1] = $false ;  Octave -12 }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Left)   -and -not $Released[1])   { $Released[1] = $true }

    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Right) -and $Released[3])       { $Released[3] = $false ; Octave +12 }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Right)   -and -not $Released[3])  { $Released[3] = $true }

    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Q)               -and $CanPlay[$key[0]])  { Play $key[0]  'C'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D2)              -and $CanPlay[$key[1]])  { Play $key[1]  'C#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::W)               -and $CanPlay[$key[2]])  { Play $key[2]  'D'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D3)              -and $CanPlay[$key[3]])  { Play $key[3]  'D#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::E)               -and $CanPlay[$key[4]])  { Play $key[4]  'E'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::R)               -and $CanPlay[$key[5]])  { Play $key[5]  'F'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D5)              -and $CanPlay[$key[6]])  { Play $key[6]  'F#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::T)               -and $CanPlay[$key[7]])  { Play $key[7]  'G'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D6)              -and $CanPlay[$key[8]])  { Play $key[8]  'G#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Y)               -and $CanPlay[$key[9]])  { Play $key[9]  'A'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D7)              -and $CanPlay[$key[10]]) { Play $key[10] 'A#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::U)               -and $CanPlay[$key[11]]) { Play $key[11] 'B'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::I)               -and $CanPlay[$key[12]]) { Play $key[12] 'C'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D9)              -and $CanPlay[$key[13]]) { Play $key[13] 'C#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::O)               -and $CanPlay[$key[14]]) { Play $key[14] 'D'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::D0)              -and $CanPlay[$key[15]]) { Play $key[15] 'D#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::P)               -and $CanPlay[$key[16]]) { Play $key[16] 'E'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::OemOpenBrackets) -and $CanPlay[$key[17]]) { Play $key[17] 'F'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::OemPlus)         -and $CanPlay[$key[18]]) { Play $key[18] 'F#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Oem6)            -and $CanPlay[$key[19]]) { Play $key[19] 'G'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::A)               -and $CanPlay[$key[20]]) { Play $key[20] 'G#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::Z)               -and $CanPlay[$key[21]]) { Play $key[21] 'A'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::S)               -and $CanPlay[$key[22]]) { Play $key[22] 'A#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::X)               -and $CanPlay[$key[23]]) { Play $key[23] 'B'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::C)               -and $CanPlay[$key[24]]) { Play $key[24] 'C'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::F)               -and $CanPlay[$key[25]]) { Play $key[25] 'C#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::V)               -and $CanPlay[$key[26]]) { Play $key[26] 'D'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::G)               -and $CanPlay[$key[27]]) { Play $key[27] 'D#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::B)               -and $CanPlay[$key[28]]) { Play $key[28] 'E'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::N)               -and $CanPlay[$key[29]]) { Play $key[29] 'F'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::J)               -and $CanPlay[$key[30]]) { Play $key[30] 'F#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::M)               -and $CanPlay[$key[31]]) { Play $key[31] 'G'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::K)               -and $CanPlay[$key[32]]) { Play $key[32] 'G#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::OemComma)        -and $CanPlay[$key[33]]) { Play $key[33] 'A'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::L)               -and $CanPlay[$key[34]]) { Play $key[34] 'A#' }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::OemPeriod)       -and $CanPlay[$key[35]]) { Play $key[35] 'B'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::OemQuestion)     -and $CanPlay[$key[36]]) { Play $key[36] 'C'  }
    if ([Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::OemQuotes)       -and $CanPlay[$key[37]]) { Play $key[37] 'C#' }

    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Q)               -and -not $CanPlay[$key[0]])  { Stop $key[0]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D2)              -and -not $CanPlay[$key[1]])  { Stop $key[1]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::W)               -and -not $CanPlay[$key[2]])  { Stop $key[2]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D3)              -and -not $CanPlay[$key[3]])  { Stop $key[3]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::E)               -and -not $CanPlay[$key[4]])  { Stop $key[4]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::R)               -and -not $CanPlay[$key[5]])  { Stop $key[5]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D5)              -and -not $CanPlay[$key[6]])  { Stop $key[6]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::T)               -and -not $CanPlay[$key[7]])  { Stop $key[7]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D6)              -and -not $CanPlay[$key[8]])  { Stop $key[8]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Y)               -and -not $CanPlay[$key[9]])  { Stop $key[9]  }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D7)              -and -not $CanPlay[$key[10]]) { Stop $key[10] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::U)               -and -not $CanPlay[$key[11]]) { Stop $key[11] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::I)               -and -not $CanPlay[$key[12]]) { Stop $key[12] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D9)              -and -not $CanPlay[$key[13]]) { Stop $key[13] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::O)               -and -not $CanPlay[$key[14]]) { Stop $key[14] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::D0)              -and -not $CanPlay[$key[15]]) { Stop $key[15] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::P)               -and -not $CanPlay[$key[16]]) { Stop $key[16] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::OemOpenBrackets) -and -not $CanPlay[$key[17]]) { Stop $key[17] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::OemPlus)         -and -not $CanPlay[$key[18]]) { Stop $key[18] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Oem6)            -and -not $CanPlay[$key[19]]) { Stop $key[19] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::A)               -and -not $CanPlay[$key[20]]) { Stop $key[20] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::Z)               -and -not $CanPlay[$key[21]]) { Stop $key[21] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::S)               -and -not $CanPlay[$key[22]]) { Stop $key[22] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::X)               -and -not $CanPlay[$key[23]]) { Stop $key[23] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::C)               -and -not $CanPlay[$key[24]]) { Stop $key[24] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::F)               -and -not $CanPlay[$key[25]]) { Stop $key[25] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::V)               -and -not $CanPlay[$key[26]]) { Stop $key[26] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::G)               -and -not $CanPlay[$key[27]]) { Stop $key[27] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::B)               -and -not $CanPlay[$key[28]]) { Stop $key[28] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::N)               -and -not $CanPlay[$key[29]]) { Stop $key[29] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::J)               -and -not $CanPlay[$key[30]]) { Stop $key[30] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::M)               -and -not $CanPlay[$key[31]]) { Stop $key[31] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::K)               -and -not $CanPlay[$key[32]]) { Stop $key[32] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::OemComma)        -and -not $CanPlay[$key[33]]) { Stop $key[33] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::L)               -and -not $CanPlay[$key[34]]) { Stop $key[34] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::OemPeriod)       -and -not $CanPlay[$key[35]]) { Stop $key[35] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::OemQuestion)     -and -not $CanPlay[$key[36]]) { Stop $key[36] }
    if ([Windows.Input.Keyboard]::IsKeyUp([System.Windows.Input.Key]::OemQuotes)       -and -not $CanPlay[$key[37]]) { Stop $key[37] }
}
