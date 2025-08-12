#/////////////////////////////////////////////////////////////////////////////////////////////
cls
$Name="EVTX_TIME_RANGE"
$version="0.4"
$Creation_Date = "07:24 12/08/2025"

#Delete unused method
#--------------------
$Creation_How = "Full ChatGPT+Gemini+Humain Brain"

#Logo Exemple
#------------
$Logo = "________  ________`n| EVTX |  | EVTX |`n|------|=>|------|`n|      |  |      |`n|______|  |______|`n"

#Todo
#----
$Todo="To Do"
$Todo+="0.4_Correct Vertical Time Line"
$Todo+="0.2_Add Vertical Time Line"
$Todo+="0.1_Correct Picture creation"

#Label
#-----
if (($Name.Length) -gt ($version.Length)){
	$fior1=$Name.Length+10;
	$fior2=($name.length)-($name.length)
	$fior3=($name.length)-($version.length)
} else {
	$fior1=$version.length+10;
	$fior2=($version.length)-($name.length)
	$fior3=(($version.length)-($version.length))
}

$fior1result="";for ($j=1; $j -le $fior1; $j++) { $fior1result+="#" }
$fior2result="";for ($j=1; $j -le $fior2; $j++) { $fior2result+=" " }
$fior3result="";for ($j=1; $j -le $fior3; $j++) { $fior3result+=" " }

write-host $fior1result
write-host "####"$Name$fior2result" ####"
write-host "####"$version$fior3result" ####"
write-host $fior1result
get-date -displayHint Time
write-host $Logo

sleep
cls
#/////////////////////////////////////////////////////////////////////////////////////////////

# ─── CONFIGURATION ────────────────────────────────────────────────
$inputFolder = ".\IN"
$outputCsv = ".\OUT\EVTX_Time_Range.csv"
$outputImage = ".\OUT\EVTX_Time_Range.png"

Add-Type -AssemblyName System.Drawing

# ─── RÉCUPÉRATION DES DONNÉES ─────────────────────────────────────
$results = @()

foreach ($file in Get-ChildItem -Path $inputFolder -Filter *.evtx) {
    try {
        $first = Get-WinEvent -Path $file.FullName -MaxEvents 1 -oldest | Sort-Object TimeCreated | Select-Object -First 1
        $last  = Get-WinEvent -Path $file.FullName -MaxEvents 1 | Sort-Object TimeCreated -Descending | Select-Object -First 1

        write-host -fore green "FIRST : $($first.TimeCreated)"
        write-host -fore green "LAST  : $($last.TimeCreated)"

        if ($first -and $last) {
            $results += [PSCustomObject]@{
                NomFichier        = $file.Name
                PremierEvenement  = $first.TimeCreated
                DernierEvenement  = $last.TimeCreated
            }
        }
    } catch {
        Write-Warning "Erreur avec $($file.FullName) : $_"
    }
}

# ─── EXPORT CSV ───────────────────────────────────────────────────
$results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
Write-Host "`n CSV exporté vers $outputCsv"

# ─── GÉNÉRATION IMAGE ─────────────────────────────────────────────
$imgWidth = 1200
$imgHeight = 50 + ($results.Count * 40)
$bmp = New-Object System.Drawing.Bitmap $imgWidth, $imgHeight
$gfx = [System.Drawing.Graphics]::FromImage($bmp)
$gfx.SmoothingMode = "AntiAlias"
$font = New-Object System.Drawing.Font "Arial", 10
$brush = [System.Drawing.Brushes]::Black
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::SkyBlue), 10

# Fond blanc
$gfx.Clear([System.Drawing.Color]::White)

# Dates min/max
$minDate = ($results | Measure-Object -Property PremierEvenement -Minimum).Minimum
$maxDate = ($results | Measure-Object -Property DernierEvenement -Maximum).Maximum
$timespan = ($maxDate - $minDate).TotalSeconds

# Fonction conversion date → X
function Get-X([datetime]$date) {
    $offset = ($date - $minDate).TotalSeconds
    return [int](($offset / $timespan) * ($imgWidth - 200)) + 100
}

# Axe horizontal
$gfx.DrawLine([System.Drawing.Pens]::Gray, 100, 30, $imgWidth - 100, 30)
$gfx.DrawString($minDate.ToString("yyyy-MM-dd HH:mm"), $fon
