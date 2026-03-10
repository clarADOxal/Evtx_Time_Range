#/////////////////////////////////////////////////////////////////////////////////////////////
cls
$Name="EVTX_TIME_RANGE"
$version="0.49"
$Creation_Date = "19:44 10/03/2026"

$Todo+="export en Xlsx"
$Todo+="export en HTML"
$Todo+="présentation en OGV"
$Todo+="Lecture optimisée par blocs"
$Todo+="Lecture multi-thread"
$Todo+="Date au format yyyy-mm-dd meme en csv"
$Todo+="OK - Retrait des guillemets pour etre utilisé avec TimeLineExplorer"
$Todo+="Utilisable avec EVTX et CSV"
$Todo+="Afficher en console la toute premire date et la dernière"


#Logo Exemple
$Logo = "________  ________`n| EVTX |  | EVTX |`n|------|=>|------|`n|      |  |      |`n|______|  |______|`n"

#Label
if (($Name.Length) -gt ($version.Length)){
    $fior1=$Name.Length+10;
    $fior2=($Name.Length)-($Name.Length)
    $fior3=($Name.Length)-($version.Length)
} else {
    $fior1=$version.Length+10;
    $fior2=($version.Length)-($Name.Length)
    $fior3=(($version.Length)-($version.Length))
}

$fior1result="";for ($j=1; $j -le $fior1; $j++) { $fior1result+="#" }
$fior2result="";for ($j=1; $j -le $fior2; $j++) { $fior2result+=" " }
$fior3result="";for ($j=1; $j -le $fior3; $j++) { $fior3result+=" " }



write-host $fior1result
write-host "####"$Name$fior2result" ####"
write-host "####"$version$fior3result" ####"
write-host $fior1result
get-date -displayHint Time
$PSVER = $PSVersionTable.PsVersion.Major


#--- Check PowerShell Version
write-host "Powershell Version in use : " -nonewline
write-host -fore red $PSVER
write-host $Logo


sleep 5
if ($PSVER -lt "5"){Write-host -fore red "Upgrade your PowerShell";sleep 5;exit}


# --- CONFIGURATION ------------------------------------------------
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$inputFolder = ".\uploads\auto\C%3A\Windows\System32\winevt\Logs"
$outputFolder = ".\OUT"

$totalevtx=(get-childitem $inputFolder *.evtx).Count
$emptycount = 0

if (-not (Test-Path $inputFolder)) { New-Item -Path $inputFolder -ItemType Directory | Out-Null; Write-Host -ForegroundColor Red "Folder IN Created" }
if (-not (Test-Path $outputFolder)) { New-Item -Path $outputFolder -ItemType Directory | Out-Null; Write-Host -ForegroundColor Red "Folder OUT Created" }

$outputFolderFull = (Resolve-Path $outputFolder).Path
$outputCsv   = Join-Path -Path $outputFolderFull -ChildPath ("EVTX_Time_Range_" + $timestamp + ".csv")
$outputCsv4TimeLineExpl   = Join-Path -Path $outputFolderFull -ChildPath ("EVTX_Time_Range_" + $timestamp + "_ForTimeLineExplorer.csv")
$outputImage = Join-Path -Path $outputFolderFull -ChildPath ("EVTX_Time_Range_" + $timestamp + ".png")


Add-Type -AssemblyName System.Drawing

# --- RÉCUPÉRATION DES DONNÉES -------------------------------------
$results = @()
foreach ($file in Get-ChildItem -Path $inputFolder -Filter *.evtx) {
    try {
        $first = Get-WinEvent -Path $file.FullName -MaxEvents 1 -Oldest -ErrorAction SilentlyContinue
        $last  = Get-WinEvent -Path $file.FullName -MaxEvents 1 -ErrorAction SilentlyContinue

	$reader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader($file.FullName, [System.Diagnostics.Eventing.Reader.PathType]::FilePath)
	$count = 0
	while ($reader.ReadEvent()) { $count++ }
	$eventCount = $count


	# --- NOUVELLES DONNÉES À AJOUTER ---
#	$eventCount = (Get-WinEvent -Path $file.FullName -ErrorAction SilentlyContinue).Count
#	$eventCount = [int](wevtutil qe "$($file.FullName)" /c:1 /rd:true 2>$null)
#	$eventCount = [int](wevtutil qe "$($file.FullName)" /lf:true /c:0 2>$null)


	$fileSizeKB = [math]::Round(($file.Length / 1KB), 2)
	$daysDiff   = [math]::Round((($last.TimeCreated - $first.TimeCreated).TotalDays), 2)
#	$daysDiff = [math]::Floor(($last.TimeCreated - $first.TimeCreated).TotalDays)
	# -----------------------------------


        if (-not $first -or -not $last) { Write-Warning "$($file.Name) ne contient aucun événement."; $emptycount++;continue }

        Write-Host -ForegroundColor Green "File : $($file.Name)"
#	$results += [PSCustomObject]@{ NomFichier = $file.Name; PremierEvenement = $first.TimeCreated; DernierEvenement = $last.TimeCreated }

	$results += [PSCustomObject]@{
		NomFichier        = $file.Name
		PremierEvenement  = $first.TimeCreated
		DernierEvenement  = $last.TimeCreated
		NombreEvenements  = $eventCount      # <<< ajouté
		PoidsFichierKo    = $fileSizeKB      # <<< ajouté
		#DureeJours        = $daysDiff       # <<< ajouté
		DureeJours = [int][math]::Floor(($last.TimeCreated - $first.TimeCreated).TotalDays)

}

    } catch { Write-Warning "Erreur avec $($file.FullName) : $_" }
}

# --- STATISTIQUES GLOBALES (placer ici) --------------------------
if ($results.Count -gt 0) {
    $globalMin  = ($results | Measure-Object -Property PremierEvenement -Minimum).Minimum
    $globalMax  = ($results | Measure-Object -Property DernierEvenement -Maximum).Maximum
    $globalDays = [math]::Floor(($globalMax - $globalMin).TotalDays)

    $totalEvents  = ($results | Measure-Object -Property NombreEvenements -Sum).Sum
    $totalSizeMB  = [math]::Round((($results | Measure-Object -Property PoidsFichierKo -Sum).Sum / 1024),2)

}

# --- EXPORT CSV ---------------------------------------------------
if ($results.Count -eq 0) { 
    Write-Warning "Aucun fichier .evtx valide trouvé. Aucun CSV ne sera généré." 
}
else {

$csvData = $results |
Select-Object `
NomFichier,
@{Name="PremierEvenement";Expression={$_.PremierEvenement.ToString("yyyy-MM-dd HH:mm:ss")}},
@{Name="DernierEvenement";Expression={$_.DernierEvenement.ToString("yyyy-MM-dd HH:mm:ss")}},
NombreEvenements,
PoidsFichierKo,
DureeJours

# CSV normal
$csvData | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

# CSV pour TimelineExplorer
$csvData | Export-Csv -Path $outputCsv4TimeLineExpl -NoTypeInformation -Encoding UTF8
(Get-Content $outputCsv4TimeLineExpl) -replace '"' , '' | Set-Content $outputCsv4TimeLineExpl



write-host "============================================================"
$validcount = $totalevtx - $emptycount

Write-Host ("EVTX total:{0} | analysés:{1} | vides:{2} | events:{3} | période:{4} jours | volume:{5} MB" -f `
        $totalevtx, $results.Count, $emptycount, $totalEvents, $globalDays, $totalSizeMB)


Write-Host "Première trace globale :" $globalMin
Write-Host "Dernière trace globale :" $globalMax


}

# --- GÉNÉRATION IMAGE ---------------------------------------------
if ($results.Count -eq 0) {
    Write-Warning "Aucun événement valide trouvé. L'image ne sera pas générée."
} else {
    $imgWidth  = 1200
    $imgHeight = 50 + ($results.Count * 40)

    # Supprime l'image existante si nécessaire
    if (Test-Path $outputImage) {
        Remove-Item $outputImage -Force
    }

    # Création du bitmap et des objets graphiques
    $bmp  = New-Object System.Drawing.Bitmap $imgWidth, $imgHeight
    $gfx  = [System.Drawing.Graphics]::FromImage($bmp)
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

    # Fonction conversion date ? X
    function Get-X([datetime]$date) {
        $offset = ($date - $minDate).TotalSeconds
        return [int](($offset / $timespan) * ($imgWidth - 200)) + 100
    }

    # Axe horizontal
    $gfx.DrawLine([System.Drawing.Pens]::Gray, 100, 30, $imgWidth - 100, 30)
    $gfx.DrawString($minDate.ToString("yyyy-MM-dd HH:mm"), $font, $brush, 100, 10)
    $gfx.DrawString($maxDate.ToString("yyyy-MM-dd HH:mm"), $font, $brush, $imgWidth - 200, 10)

    # Frise pour chaque fichier
    $i = 0
    foreach ($item in $results) {
        $y = 60 + ($i * 40)
        $xStart = Get-X $item.PremierEvenement
        $xEnd   = Get-X $item.DernierEvenement

        # Ligne représentant la période du fichier
        $gfx.DrawLine($pen, $xStart, $y, $xEnd, $y)

        # Nom du fichier à gauche
        $gfx.DrawString($item.NomFichier, $font, $brush, 10, $y - 7)

        $i++
    }

    # Sauvegarde image via FileStream pour éviter les erreurs GDI+
    try {
        $fs = [System.IO.File]::Open($outputImage, [System.IO.FileMode]::Create)
        $bmp.Save($fs, [System.Drawing.Imaging.ImageFormat]::Png)
        $fs.Close()
	Write-host "============================================================"
        Write-Host "Frise enregistrée dans :"
	Write-host $outputImage
    } catch {
        Write-Warning "Impossible de sauvegarder l'image : $_"
    }

    # Libération des ressources graphiques
    $gfx.Dispose()
    $bmp.Dispose()
}