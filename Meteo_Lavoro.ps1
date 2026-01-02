# IMPORTANTE: Salvare questo file con codifica "UTF-8 with BOM"
<#
.SYNOPSIS
Script PowerShell AVANZATO e GRATUITO per monitoraggio meteo (previsioni ogni 3 ore) e qualitÃ  dell'aria,
con generazione di report Word e spiegazione stato.

.DESCRIPTION
- Utilizza API 5 day / 3 hour forecast per previsioni a intervalli di 3 ore.
- Spiegazione del motivo per stato ATTENZIONE/CRITICO.
- Report Word limitato alle ore 07-19 per stare in una pagina (considerando slot di 3 ore).
- Correzioni per formato data, apertura file, output COM.

.NOTES
Autore    : Gemini (Evoluzione per l'utente)
Versione  : 2.14 (API 3-Ore, Adattamenti Struttura Dati)
Data      : 2025-05-07
#>

param(
    [Parameter(Mandatory=$false)] [string]$ApiKey = "10cd392e7b7ac7e0d18694484f80977d", # <<< MODIFICA QUI!!!
    [Parameter(Mandatory=$false)] [string]$City = "Siracusa,IT",
    [Parameter(Mandatory=$false)] [string]$Language = "it",
    [Parameter(Mandatory=$false)] [double]$ThresholdRainMm = 5.0, # Ora si riferisce a mm in 3 ore
    [Parameter(Mandatory=$false)] [double]$ThresholdWindKmh = 40.0,
    [Parameter(Mandatory=$false)] [double]$ThresholdTempLowC = 5.0,
    [Parameter(Mandatory=$false)] [double]$ThresholdTempHighC = 30.0
)

# --- Inizializzazione Tipi ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- DEFINIZIONE FUNZIONI AUSILIARIE ---

function Add-ColoredText {
    param( [System.Windows.Forms.RichTextBox]$Control, [string]$Text, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black, [bool]$Bold = $false, [bool]$AddNewLine = $false )
    if ($null -eq $Control -or $Control.IsDisposed -or -not $Control.IsHandleCreated) { return }
    try { $Control.SelectionStart = $Control.TextLength; $Control.SelectionLength = 0; $Control.SelectionColor = $Color; $currentFontStyle = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }; try { $Control.SelectionFont = New-Object System.Drawing.Font($Control.Font.FontFamily, $Control.Font.Size, $currentFontStyle) -ErrorAction Stop } catch {}; $textToAppend = $Text; if ($AddNewLine) { $textToAppend += "`r`n" }; $Control.AppendText($textToAppend); $Control.SelectionColor = $Control.ForeColor; $Control.SelectionFont = $Control.Font } catch { Write-Warning "Errore Add-ColoredText: $($_.Exception.Message)" }
}

function Convert-WindDirection {
    param($degrees); [string]$directionName = ""; switch ($degrees) {{$_ -ge 348.75 -or $_ -lt 11.25}{ $directionName="Nord" };{$_ -ge 11.25 -and $_ -lt 33.75}{ $directionName="Nord-Nord-Est" };{$_ -ge 33.75 -and $_ -lt 56.25}{ $directionName="Nord-Est" };{$_ -ge 56.25 -and $_ -lt 78.75}{ $directionName="Est-Nord-Est" };{$_ -ge 78.75 -and $_ -lt 101.25}{ $directionName="Est" };{$_ -ge 101.25 -and $_ -lt 123.75}{ $directionName="Est-Sud-Est" };{$_ -ge 123.75 -and $_ -lt 146.25}{ $directionName="Sud-Est" };{$_ -ge 146.25 -and $_ -lt 168.75}{ $directionName="Sud-Sud-Est" };{$_ -ge 168.75 -and $_ -lt 191.25}{ $directionName="Sud" };{$_ -ge 191.25 -and $_ -lt 213.75}{ $directionName="Sud-Sud-Ovest" };{$_ -ge 213.75 -and $_ -lt 236.25}{ $directionName="Sud-Ovest" };{$_ -ge 236.25 -and $_ -lt 258.75}{ $directionName="Ovest-Sud-Ovest" };{$_ -ge 258.75 -and $_ -lt 281.25}{ $directionName="Ovest" };{$_ -ge 281.25 -and $_ -lt 303.75}{ $directionName="Ovest-Nord-Ovest" };{$_ -ge 303.75 -and $_ -lt 326.25}{ $directionName="Nord-Ovest" };{$_ -ge 326.25 -and $_ -lt 348.75}{ $directionName="Nord-Nord-Ovest" };default{ $directionName="$degrees° (?)" }}; $directionNameUpper = $directionName.ToUpper(); $lastDashIndex = $directionNameUpper.LastIndexOf('-'); if ($lastDashIndex -gt 0 -and $directionNameUpper.IndexOf('-') -ne $lastDashIndex){ $part1 = $directionNameUpper.Substring(0, $lastDashIndex); $part2 = $directionNameUpper.Substring($lastDashIndex + 1); return "$part1/$part2" } else { return $directionNameUpper }
}

function Get-WeatherIcon {
      param($weatherMain, $description); $weatherMainLower = $weatherMain.ToLower(); $descriptionLower = $description.ToLower(); $weatherIconSymbol = switch -Wildcard ($weatherMainLower) {"clear"{"☀️"};"clouds"{if($descriptionLower -like "*pochi*" -or $descriptionLower -like "*poche*"){"🌤️"}elseif($descriptionLower -like "*coperto*"){"☁️"}else{"🌥️"}};"rain"{if($descriptionLower -like "*forte*"){"🌧️"}else{"🌦️"}};"drizzle"{"💧"};"thunderstorm"{"⛈️"};"snow"{"❄️"};"mist"{"🌫️"};"fog"{"🌫️"};"smoke"{"💨"};"haze"{"🌫️"};"dust"{"🌬️"};"sand"{"🌬️"};"ash"{"🌋"};"squall"{"🌬️"};"tornado"{"🌪️"};default{"❓"}}; return $weatherIconSymbol
}

function Get-AqiInfo {
    param($aqi); $color = [System.Drawing.Color]::Black; $text = ""; switch ($aqi) { 1{$text="Buona";$color=[System.Drawing.Color]::Green} 2{$text="Discreta";$color=[System.Drawing.Color]::LimeGreen} 3{$text="Moderata";$color=[System.Drawing.Color]::Orange} 4{$text="Scadente";$color=[System.Drawing.Color]::Red} 5{$text="Pessima";$color=[System.Drawing.Color]::Purple} default{$text="N/D";$color=[System.Drawing.Color]::Gray} }; return @{ Text = $text; Color = $color }
}

function Set-ButtonPanelLayout {
    param($panel, $btn1, $btn2); try { $panelWidth=[int]$panel.ClientSize.Width; $buttonWidth=[int]$btn1.Width; $buttonHeight=[int]$btn1.Height; $spacing=20; $totalButtonsWidth=($buttonWidth*2)+$spacing; $buttonY=([int]$panel.ClientSize.Height-$buttonHeight)/2; $startX=5; if($panelWidth -gt $totalButtonsWidth){$startX=($panelWidth-$totalButtonsWidth)/2}; $btn1.Location=New-Object System.Drawing.Point([int]$startX,[int]$buttonY); $btn2.Location=New-Object System.Drawing.Point([int]($startX+$buttonWidth+$spacing),[int]$buttonY) } catch { Write-Warning "AVVISO Posizionamento Pulsanti: $($_.Exception.Message)" }
}

function New-WordReport {
    param(
        [Parameter(Mandatory=$true)] $ReportData
    )
    Write-Host "[Word Report] Funzione avviata."
    $reportPath = "C:\Users\Coemi\Desktop\REPORT METEO"
    $safeCityName = ($ReportData.CityName -replace '[\\/:*?"<>|]', '').ToUpper()
    $baseFileName = "REPORT_METEO_3H_$($safeCityName)_$((Get-Date).ToString('dd_MM_yyyy_HH_mm'))"
    $fullFilePath = Join-Path -Path $reportPath -ChildPath ($baseFileName + ".docx")

    if (-not (Test-Path -Path $reportPath -PathType Container)) { Write-Host "[Word Report] Creo directory: $reportPath"; try { New-Item -ItemType Directory -Path $reportPath -Force -EA Stop|Out-Null } catch { Write-Warning "[Word Report] Impossibile creare directory '$reportPath'. Errore: $($_.Exception.Message)"; return $false } }

    $wdAlignParagraphCenter=1; $wdAlignParagraphLeft=0; $wdColorBlack=0; $wdColorRed=255; $wdColorOrange=46079; $wdColorGreen=32768; $wdColorBlue=16711680
    $wdFormatDocumentDefault=16; $wdStory=6; $wdAutoFitContent=1; $wdStyleHeading1=-2; $wdStyleHeading2=-3; $wdStyleNormal=-1
    $fsTitle = 16; $fsDate = 10; $fsH1 = 14; $fsH2 = 12; $fsNormal = 10; $fsTable = 9; $fsStatus = 11; $fsAlert = 10

    $wordApp = $null; $wordDoc = $null; $wordSelection = $null
    try {
        Write-Host "[Word Report] Avvio MS Word COM..."; $wordApp = New-Object -ComObject Word.Application; $wordApp.Visible = $false;
        $wordDoc = $wordApp.Documents.Add(); $wordSelection = $wordApp.Selection;

        $docPageSetup = $wordDoc.PageSetup; $docPageSetup.LeftMargin=$wordApp.InchesToPoints(0.6); $docPageSetup.RightMargin=$wordApp.InchesToPoints(0.6); $docPageSetup.TopMargin=$wordApp.InchesToPoints(0.7); $docPageSetup.BottomMargin=$wordApp.InchesToPoints(0.7)
        $wordSelection.ParagraphFormat.SpaceBefore = 0; $wordSelection.ParagraphFormat.SpaceAfter = 3;

        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleHeading1) } catch { $wordSelection.Style = "Heading 1"}
        $wordSelection.ParagraphFormat.Alignment = $wdAlignParagraphCenter; $wordSelection.Font.Size = $fsTitle; $wordSelection.Font.Bold = $true
        [void]$wordSelection.TypeText("Report Meteo (3 Ore) - $($ReportData.CityName), $($ReportData.Country)")
        [void]$wordSelection.TypeParagraph(); $wordSelection.Font.Size = $fsDate; $wordSelection.Font.Bold = $false
        [void]$wordSelection.TypeText("Generato il: $((Get-Date).ToString('dddd dd MMMM yyyy HH:mm:ss'))"); [void]$wordSelection.TypeParagraph(); [void]$wordSelection.TypeParagraph()

        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleHeading1) } catch { $wordSelection.Style = "Heading 1"}; $wordSelection.Font.Size = $fsH1; $wordSelection.Font.Bold = $true; $wordSelection.ParagraphFormat.Alignment = $wdAlignParagraphLeft; $wordSelection.TypeText("Riepilogo Giornaliero"); [void]$wordSelection.TypeParagraph()
        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleNormal) } catch { $wordSelection.Style = "Normal"}; $wordSelection.Font.Size = $fsNormal; $wordSelection.Font.Bold = $false; $wordSelection.ParagraphFormat.SpaceAfter = 2
        $minT=if($null -ne $ReportData.MinTemp){$ReportData.MinTemp.ToString('F1')}else{"N/D"}; $maxT=if($null -ne $ReportData.MaxTemp){$ReportData.MaxTemp.ToString('F1')}else{"N/D"}
        $minH=if($null -ne $ReportData.MinHumidity){$ReportData.MinHumidity.ToString('F0')}else{"N/D"}; $maxH=if($null -ne $ReportData.MaxHumidity){$ReportData.MaxHumidity.ToString('F0')}else{"N/D"}; $avgH=if($null -ne $ReportData.AvgHumidity){$ReportData.AvgHumidity.ToString('F0')}else{"N/D"}
        $minP=if($null -ne $ReportData.MinPressure){$ReportData.MinPressure.ToString('F0')}else{"N/D"}; $maxP=if($null -ne $ReportData.MaxPressure){$ReportData.MaxPressure.ToString('F0')}else{"N/D"}; $avgP=if($null -ne $ReportData.AvgPressure){$ReportData.AvgPressure.ToString('F0')}else{"N/D"}
        $minV=if($null -ne $ReportData.MinVisibility){($ReportData.MinVisibility/1000).ToString('F1')}else{"N/D"}; $maxV=if($null -ne $ReportData.MaxVisibility){($ReportData.MaxVisibility/1000).ToString('F1')}else{"N/D"}
        [void]$wordSelection.TypeText("• Temperatura: Min ~$minT°C / Max ~$maxT°C"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("• Umidità: Min ~$minH% / Max ~$maxH% (Media: ~$avgH%)"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("• Pressione: Min ~$minP hPa / Max ~$maxP hPa (Media: ~$avgP hPa)"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("• Visibilità: Min ~$minV km / Max ~$maxV km"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("• Condizioni: Mattina ~$($ReportData.MorningDesc) / Pomeriggio ~$($ReportData.AfternoonDesc)"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("• Prob. Max Precip: $($ReportData.MaxPop.ToString('F0'))%"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("• Alba/Tramonto: $($ReportData.Sunrise) / $($ReportData.Sunset)"); [void]$wordSelection.TypeParagraph(); [void]$wordSelection.TypeParagraph()

        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleHeading2) } catch { $wordSelection.Style = "Heading 2"}
        $wordSelection.Font.Size = $fsH2; $wordSelection.Font.Bold = $true; $wordSelection.TypeText("Soglie di Allarme"); [void]$wordSelection.TypeParagraph()
        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleNormal) } catch { $wordSelection.Style = "Normal"}
        $wordSelection.Font.Size = $fsNormal - 1; $wordSelection.Font.Bold = $false; $wordSelection.ParagraphFormat.SpaceAfter = 0
        [void]$wordSelection.TypeText("  - Pioggia > $($ReportData.Thresholds.RainMm.ToString('F1')) mm/3h"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("  - Vento > $($ReportData.Thresholds.WindKmh.ToString('F1')) km/h"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("  - Temp. < $($ReportData.Thresholds.TempLowC.ToString('F1')) °C"); [void]$wordSelection.TypeParagraph()
        [void]$wordSelection.TypeText("  - Temp. > $($ReportData.Thresholds.TempHighC.ToString('F1')) °C"); [void]$wordSelection.TypeParagraph(); [void]$wordSelection.TypeParagraph()

        if ($ReportData.AqiData) {
            try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleHeading1) } catch { $wordSelection.Style = "Heading 1"}; $wordSelection.Font.Size = $fsH1; $wordSelection.Font.Bold = $true; $wordSelection.TypeText("Qualità dell'Aria"); [void]$wordSelection.TypeParagraph()
            try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleNormal) } catch { $wordSelection.Style = "Normal"}; $wordSelection.Font.Size = $fsNormal; $wordSelection.Font.Bold = $false; $wordSelection.ParagraphFormat.SpaceAfter = 3
            $aqiText = "$($ReportData.AqiData.main.aqi) - $($ReportData.AqiInfo.Text)"; [void]$wordSelection.TypeText("  • Indice AQI: $aqiText"); [void]$wordSelection.TypeParagraph()
            $wordSelection.Font.Size = $fsNormal -1 ; [void]$wordSelection.TypeText("    PM₂.₅: $($ReportData.AqiData.components.pm2_5)|PM₁₀: $($ReportData.AqiData.components.pm10)|O₃: $($ReportData.AqiData.components.o3)|NO₂: $($ReportData.AqiData.components.no2) (µg/m³)"); [void]$wordSelection.TypeParagraph(); [void]$wordSelection.TypeParagraph()
        }

        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleHeading1) } catch { $wordSelection.Style = "Heading 1"}
        $wordSelection.Font.Size = $fsH1; $wordSelection.Font.Bold = $true; $wordSelection.TypeText("Dettaglio Previsioni (07-19, intervalli 3h)"); [void]$wordSelection.TypeParagraph()
        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleNormal) } catch { $wordSelection.Style = "Normal"}
        $wordSelection.Font.Bold = $false; $wordSelection.Font.Size = $fsTable; $wordSelection.ParagraphFormat.SpaceAfter = 0

        $forecastsForTable = $ReportData.Forecasts | Where-Object {
            $hour = ([datetimeoffset]::FromUnixTimeSeconds($_.dt)).DateTime.ToLocalTime().Hour
            $hour -ge 7 -and $hour -le 19 # Questo filtro ora agirÃ  su slot di 3 ore
        }

        if ($forecastsForTable -and $forecastsForTable.Count -gt 0) {
            $tableRange = $wordSelection.Range; $numRows = $forecastsForTable.Count + 1; $numCols = 5; $meteoTable = $null
            try {
                $meteoTable = $wordDoc.Tables.Add($tableRange, $numRows, $numCols);
                $meteoTable.Borders.Enable = $true; $meteoTable.Rows[1].Range.Font.Bold = $true; $meteoTable.Rows[1].Shading.BackgroundPatternColor = -587137025
                $meteoTable.Cell(1, 1).Range.Text = "Ora"; $meteoTable.Cell(1, 2).Range.Text = "Condizione"; $meteoTable.Cell(1, 3).Range.Text = "T(P)°C"
                $meteoTable.Cell(1, 4).Range.Text = "Vento"; $meteoTable.Cell(1, 5).Range.Text = "Prec(mm/3h)" # Modificato
                $rowIdx = 2
                foreach ($item in $forecastsForTable) {
                    $timeStr = ([datetimeoffset]::FromUnixTimeSeconds($item.dt)).DateTime.ToLocalTime().ToString("HH:mm")
                    $desc = if ($item.weather -and $item.weather.Count -gt 0) {$item.weather[0].description.Substring(0,1).ToUpper()+$item.weather[0].description.Substring(1)} else {"N/D"}
                    $tempStr = if ($item.main) {"$($item.main.temp.ToString('F1'))($($item.main.feels_like.ToString('F1')))"} else {"N/D"}
                    $windStr = if ($item.wind) {"$($($item.wind.speed * 3.6).ToString('F1'))km/h $(Convert-WindDirection -degrees $item.wind.deg)"} else {"N/D"}
                    
                    $precipAmount = 0; $precipType="P" # Default a Pioggia
                    if ($item.rain -and $item.rain.PSObject.Properties.Contains("3h")) { $precipAmount = $item.rain."3h" }
                    elseif ($item.snow -and $item.snow.PSObject.Properties.Contains("3h")) { $precipAmount = $item.snow."3h"; $precipType="N" }
                    $precipStr = if ($precipAmount -gt 0) {"${precipType}:$($precipAmount.ToString('F1'))"} else {"-"}

                    $meteoTable.Cell($rowIdx, 1).Range.Text = $timeStr; $meteoTable.Cell($rowIdx, 2).Range.Text = $desc
                    $meteoTable.Cell($rowIdx, 3).Range.Text = $tempStr; $meteoTable.Cell($rowIdx, 4).Range.Text = $windStr
                    $meteoTable.Cell($rowIdx, 5).Range.Text = $precipStr
                    
                    if ($item.main.temp -lt $ReportData.Thresholds.TempLowC) { $meteoTable.Cell($rowIdx, 3).Range.Font.Color = $wdColorBlue }
                    if ($item.main.temp -gt $ReportData.Thresholds.TempHighC) { $meteoTable.Cell($rowIdx, 3).Range.Font.Color = $wdColorRed }
                    if (($item.wind.speed * 3.6) -gt $ReportData.Thresholds.WindKmh) { $meteoTable.Cell($rowIdx, 4).Range.Font.Color = $wdColorOrange }
                    if ($precipAmount -gt $ReportData.Thresholds.RainMm) { $meteoTable.Cell($rowIdx, 5).Range.Font.Color = $wdColorRed } # Confronta mm/3h con soglia
                    $rowIdx++
                }
                [void]$meteoTable.AutoFitBehavior($wdAutoFitContent);
            } catch { Write-Warning "[Word Report] Errore tabella: $($_.Exception.Message)"; throw }
            [void]$wordSelection.EndKey($wdStory); [void]$wordSelection.TypeParagraph(); [void]$wordSelection.TypeParagraph()
        } else { [void]$wordSelection.TypeText("Nessuna previsione disponibile per l'intervallo 07-19."); [void]$wordSelection.TypeParagraph(); [void]$wordSelection.TypeParagraph() }

        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleHeading1) } catch { $wordSelection.Style = "Heading 1"}
        $wordSelection.Font.Size = $fsH1; $wordSelection.Font.Bold = $true; [void]$wordSelection.TypeText("Stato Operatività"); [void]$wordSelection.TypeParagraph()
        try { $wordSelection.Style = $wordApp.ActiveDocument.Styles.Item([ref]$wdStyleNormal) } catch { $wordSelection.Style = "Normal"}
        $wordSelection.Font.Bold = $true; $wordSelection.Font.Size = $fsStatus; $statusColor = $wdColorBlack
        switch($ReportData.OperativityStatus){ "OK"{$statusColor=$wdColorGreen};"ATTENZIONE"{$statusColor=$wdColorOrange};"CRITICO"{$statusColor=$wdColorRed};"DATI PARZIALI"{$statusColor=$wdColorOrange};"ERRORE DATI"{$statusColor=$wdColorRed} }
        $statusReason = ""
        if ($ReportData.OperativityStatus -ne "OK" -and $ReportData.OperativityStatus -ne "ERRORE DATI" -and $ReportData.OperativityStatus -ne "DATI PARZIALI") {
            $reasonsList = @(); if ($ReportData.CausedByRain) { $reasonsList += "Precipitazioni" }; if ($ReportData.CausedByWind) { $reasonsList += "Vento" }; if ($ReportData.CausedByTemp) { $reasonsList += "Temperatura" }
            if ($reasonsList.Count -gt 0) { $statusReason = " (Causa: $($reasonsList -join ', '))" }
        }
        $wordSelection.Font.Color = $statusColor; [void]$wordSelection.TypeText("Stato Generale: $($ReportData.OperativityStatus)$statusReason"); [void]$wordSelection.TypeParagraph()
        $wordSelection.Font.Color = $wdColorBlack; $wordSelection.Font.Bold = $false; $wordSelection.Font.Size = $fsAlert

        if ($ReportData.Alerts -and $ReportData.Alerts.Count -gt 0) {
            $wordSelection.Font.Bold = $true; [void]$wordSelection.TypeText("Avvisi Dettagliati:"); [void]$wordSelection.TypeParagraph(); $wordSelection.Font.Bold = $false
            foreach ($msg in $ReportData.Alerts) {
                $alertColor = if ($msg -like "*Intensa*" -or $msg -like "*CRITICO*") { $wdColorRed } else { $wdColorOrange }
                $wordSelection.Font.Color = $alertColor; [void]$wordSelection.TypeText("  • $msg"); [void]$wordSelection.TypeParagraph()
            }
            $wordSelection.Font.Color = $wdColorBlack
        } elseif ($ReportData.OperativityStatus -eq "OK") { $wordSelection.Font.Color = $wdColorGreen; [void]$wordSelection.TypeText("Nessun avviso meteo significativo rilevato."); [void]$wordSelection.TypeParagraph(); $wordSelection.Font.Color = $wdColorBlack }
        
        [void]$wordDoc.SaveAs2($fullFilePath, $wdFormatDocumentDefault);
        return $true
    } catch { Write-Warning "[Word Report] Errore COM: $($_.Exception.ToString())"; return $false }
    finally {
        if ($null -ne $wordDoc)    { try { [void]$wordDoc.Close([ref]$false) } catch { Write-Warning "[Word Report] Errore chiusura doc."} }
        if ($null -ne $wordApp)    { try { [void]$wordApp.Quit() } catch { Write-Warning "[Word Report] Errore chiusura app Word."} }
        if ($null -ne $wordSelection) { $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordSelection) }
        if ($null -ne $wordDoc)       { $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDoc) }
        if ($null -ne $wordApp)       { $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) }
        $wordSelection = $null; $wordDoc = $null; $wordApp = $null; [gc]::Collect(); [gc]::WaitForPendingFinalizers();
    }
}

# --- Setup Finestra e Controlli ---
Write-Host "Setup Finestra..."
if (-not (Test-Connection "openweathermap.org" -Count 1 -Quiet)) { Write-Warning "Connessione OpenWeatherMap fallita."; exit 1} Write-Host "Connessione OpenWeatherMap OK."
$form = New-Object System.Windows.Forms.Form; $form.Text = "Super Meteo Lavoro ($City) - Controllo Ore $((Get-Date).ToString('HH:mm'))"
$form.Size = New-Object System.Drawing.Size(750, 650); $form.MinimumSize = New-Object System.Drawing.Size(700, 500)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi; $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$richTextBox = New-Object System.Windows.Forms.RichTextBox; $richTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill; $richTextBox.ReadOnly = $true
$richTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $richTextBox.BackColor = [System.Drawing.Color]::White
$richTextBox.Font = New-Object System.Drawing.Font("Cascadia Code", 10); if ($richTextBox.Font.Name -ne "Cascadia Code") { $richTextBox.Font = New-Object System.Drawing.Font("Segoe UI Emoji", 10) }; if ($richTextBox.Font.Name -ne "Segoe UI Emoji") { $richTextBox.Font = New-Object System.Drawing.Font("Consolas", 10) }
$buttonPanel = New-Object System.Windows.Forms.Panel; $buttonPanel.Height = 50; $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "OK"; $okButton.Size = New-Object System.Drawing.Size(100, 30); $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.AcceptButton = $okButton
$openReportButton = New-Object System.Windows.Forms.Button; $openReportButton.Name = "openReportButton"; $openReportButton.Text = "Apri Report"; $openReportButton.Size = New-Object System.Drawing.Size(100, 30)
$buttonPanel.Controls.Add($okButton); $buttonPanel.Controls.Add($openReportButton)
$form.Controls.Add($richTextBox); $form.Controls.Add($buttonPanel)
Write-Host "Controlli Finestra Creati."

$form_Load = {
    $formInLoad = $this
    $formInLoad.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $rtb = $formInLoad.Controls | Where-Object {$_.GetType().Name -eq 'RichTextBox'} | Select-Object -First 1
    if ($null -eq $rtb) { Write-Warning "[Load] RichTextBox non trovata!"; $formInLoad.Cursor = [System.Windows.Forms.Cursors]::Default; return }
    $rtb.Text = "Avvio Super Meteo Lavoro (3 Ore)...`r`nRecupero coordinate e dati API..."

    $script:alertMessages = [System.Collections.Generic.List[string]]::new(); $script:operativityStatus = "OK";
    $script:CausedByRain = $false; $script:CausedByWind = $false; $script:CausedByTemp = $false
    $lat = $null; $lon = $null; $script:forecastData = $null # Modificato da oneCallData
    $script:sunrise = "N/D"; $script:sunset = "N/D"; $script:cityName = "N/D"; $script:country = ""
    $script:minTempForDay = $null; $script:maxTempForDay = $null; $script:morningDesc = "N/D"; $script:afternoonDesc = "N/D"; $script:maxPop = 0
    $script:aqiData = $null; $script:aqiInfo = $null; $script:todayForecasts = $null
    $script:minHumidity = $null; $script:maxHumidity = $null; $script:avgHumidity = $null; $script:minPressure = $null; $script:maxPressure = $null; $script:avgPressure = $null; $script:minVisibility = $null; $script:maxVisibility = $null
    $cGreen=[System.Drawing.Color]::Green; $cOrange=[System.Drawing.Color]::DarkOrange; $cRed=[System.Drawing.Color]::Red; $cOrangeRed=[System.Drawing.Color]::OrangeRed; $cBlack=[System.Drawing.Color]::Black; $cBlue=[System.Drawing.Color]::Blue; $cBlueViolet=[System.Drawing.Color]::BlueViolet; $cGray=[System.Drawing.Color]::Gray

    # 1. Geocoding
    $geoApiUrl = "http://api.openweathermap.org/geo/1.0/direct?q=$([uri]::EscapeDataString($City))&limit=1&appid=$ApiKey"
    try {
        Write-Host "[Load] Contatto API Geocoding..."; Add-ColoredText -Control $rtb -Text "`nRecupero coordinate..." -Color $cGray -AddNewLine $true
        $geoData = Invoke-RestMethod -Uri $geoApiUrl -Method Get -ErrorAction Stop
        if ($geoData -and $geoData.Count -gt 0) {
            $lat=$geoData[0].lat; $lon=$geoData[0].lon; $script:cityName=$geoData[0].name; $script:country=$geoData[0].country;
            Add-ColoredText -Control $rtb -Text "Coordinate OK." -Color $cGreen -AddNewLine $true
        } else { throw "Città '$City' non trovata o API Geocoding fallita." }
    } catch { Write-Warning "[Load] ERRORE Geocoding: $($_.Exception.Message)"; $rtb.Clear(); Add-ColoredText -Control $rtb -Text "!!! ERRORE Geocoding !!!`r`nCittà '$City' non trovata?`n$($_.Exception.Message)" -Color $cRed -Bold $true -AddNewLine $true; $script:operativityStatus = "ERRORE DATI"; }

    # 2. 5 day / 3 hour Forecast API (se abbiamo Lat/Lon)
    if ($null -ne $lat -and $null -ne $lon -and $script:operativityStatus -ne "ERRORE DATI") {
        $forecastApiUrl = "http://api.openweathermap.org/data/2.5/forecast?lat=$lat&lon=$lon&appid=$ApiKey&units=metric&lang=$Language"
        try {
            Write-Host "[Load] Contatto API Previsioni 3 Ore..."; Add-ColoredText -Control $rtb -Text "Recupero previsioni (intervalli 3 ore)..." -Color $cGray -AddNewLine $true
            $script:forecastData = Invoke-RestMethod -Uri $forecastApiUrl -Method Get -ErrorAction Stop
            if ($null -eq $script:forecastData -or $null -eq $script:forecastData.list) { throw "Errore API Previsioni o dati mancanti." }
            Add-ColoredText -Control $rtb -Text "Previsioni (3 Ore) OK." -Color $cGreen -AddNewLine $true
            # Estrai Alba/Tramonto
            if ($script:forecastData.city) {
                if ($script:forecastData.city.sunrise) { $script:sunrise = ([datetimeoffset]::FromUnixTimeSeconds($script:forecastData.city.sunrise)).DateTime.ToLocalTime().ToString("HH:mm") }
                if ($script:forecastData.city.sunset)  { $script:sunset  = ([datetimeoffset]::FromUnixTimeSeconds($script:forecastData.city.sunset)).DateTime.ToLocalTime().ToString("HH:mm") }
            }
        } catch { Write-Warning "[Load] ERRORE API Previsioni: $($_.Exception.Message)"; Add-ColoredText -Control $rtb -Text "!!! ERRORE API PREVISIONI (3 Ore) !!!`r`n$($_.Exception.Message)" -Color $cRed -Bold $true -AddNewLine $true; $script:operativityStatus = "ERRORE DATI" }
    }

    # 3. AQI API (solo se abbiamo Lat/Lon)
    if ($null -ne $lat -and $null -ne $lon -and $script:operativityStatus -ne "ERRORE DATI") {
        $apiUrlAqi = "http://api.openweathermap.org/data/2.5/air_pollution?lat=$lat&lon=$lon&appid=$ApiKey"
        try {
            Write-Host "[Load] Contatto API AQI..."; Add-ColoredText -Control $rtb -Text "Recupero AQI..." -Color $cGray -AddNewLine $true;
            $aqiRawData = Invoke-RestMethod -Uri $apiUrlAqi -Method Get -ErrorAction Stop;
            if ($aqiRawData -and $aqiRawData.list -and $aqiRawData.list.Count -gt 0) {
                $script:aqiData = $aqiRawData.list[0]; $aqiValue = $script:aqiData.main.aqi; $script:aqiInfo = Get-AqiInfo -aqi $aqiValue;
                Add-ColoredText -Control $rtb -Text "Dati AQI (AQI=$aqiValue) OK." -Color $script:aqiInfo.Color -AddNewLine $true
            } else { throw "Dati AQI non validi." }
        } catch { Write-Warning "[Load] ERRORE API AQI: $($_.Exception.Message)"; Add-ColoredText -Control $rtb -Text "!!! ERRORE API AQI !!!`r`n$($_.Exception.Message)" -Color $cOrangeRed -Bold $true -AddNewLine $true } # Non impostare ERRORE DATI per AQI fallito
    }

    # --- Popolamento Finale RichTextBox ---
    $rtb.Clear()
    Add-ColoredText -Control $rtb -Text "--- Dettaglio Meteo (3 Ore) per $script:cityName, $script:country ---`r`n" -Color ([System.Drawing.Color]::DarkBlue) -Bold $true
    Add-ColoredText -Control $rtb -Text "$((Get-Date).ToString('dddd dd MMMM yyyy'))`r`n" -AddNewLine $false
    Add-ColoredText -Control $rtb -Text "Alba: $script:sunrise 🌅  Tramonto: $script:sunset 🌇`r`n" -AddNewLine $true
    Add-ColoredText -Control $rtb -Text (("-" * 75) + "`r`n")

    if ($script:aqiData -and $script:aqiInfo) {
        Add-ColoredText -Control $rtb -Text "💨 Qualità Aria: "; Add-ColoredText -Control $rtb -Text "$($script:aqiData.main.aqi) - $($script:aqiInfo.Text)" -Color $script:aqiInfo.Color -Bold $true -AddNewLine $true
        Add-ColoredText -Control $rtb -Text "    PM₂.₅: $($script:aqiData.components.pm2_5)|PM₁₀: $($script:aqiData.components.pm10)|O₃: $($script:aqiData.components.o3)|NO₂: $($script:aqiData.components.no2) (µg/m³)`r`n" -AddNewLine $true
        Add-ColoredText -Control $rtb -Text (("-" * 75) + "`r`n")
    }

    if ($script:forecastData -and $script:forecastData.list -and $script:operativityStatus -ne "ERRORE DATI") {
        $today = (Get-Date).Date
        $script:todayForecasts = $script:forecastData.list | Where-Object {
            $ts = $_.dt
            if ($null -eq $ts) { return $false }
            $forecastDate = ([datetimeoffset]::FromUnixTimeSeconds($ts)).DateTime.ToLocalTime()
            $forecastDate.Date -eq $today -and $forecastDate.Hour -ge 6 # Prendi gli slot da oggi dalle 6 in poi
        }

        if ($script:todayForecasts -and $script:todayForecasts.Count -gt 0) {
            $validForecasts = $script:todayForecasts | Where-Object { $null -ne $_.main }
            if ($validForecasts) {
                $script:minTempForDay = ($validForecasts | Measure-Object -Property {$_.main.temp_min} -Minimum -EA SilentlyContinue).Minimum
                $script:maxTempForDay = ($validForecasts | Measure-Object -Property {$_.main.temp_max} -Maximum -EA SilentlyContinue).Maximum
                $script:minHumidity = ($validForecasts | Where-Object { $null -ne $_.main.humidity } | Measure-Object -Property {$_.main.humidity} -Minimum -EA SilentlyContinue).Minimum
                $script:maxHumidity = ($validForecasts | Where-Object { $null -ne $_.main.humidity } | Measure-Object -Property {$_.main.humidity} -Maximum -EA SilentlyContinue).Maximum
                $script:avgHumidity = ($validForecasts | Where-Object { $null -ne $_.main.humidity } | Measure-Object -Property {$_.main.humidity} -Average -EA SilentlyContinue).Average
                $script:minPressure = ($validForecasts | Where-Object { $null -ne $_.main.pressure } | Measure-Object -Property {$_.main.pressure} -Minimum -EA SilentlyContinue).Minimum
                $script:maxPressure = ($validForecasts | Where-Object { $null -ne $_.main.pressure } | Measure-Object -Property {$_.main.pressure} -Maximum -EA SilentlyContinue).Maximum
                $script:avgPressure = ($validForecasts | Where-Object { $null -ne $_.main.pressure } | Measure-Object -Property {$_.main.pressure} -Average -EA SilentlyContinue).Average
                # Visibility: API /data/2.5/forecast fornisce 'visibility' direttamente nell'elemento della lista
                $visValues = $validForecasts | Where-Object { $null -ne $_.visibility } | Select-Object -ExpandProperty visibility
                if($visValues) {
                    $script:minVisibility = ($visValues | Measure-Object -Minimum -EA SilentlyContinue).Minimum
                    $script:maxVisibility = ($visValues | Measure-Object -Maximum -EA SilentlyContinue).Maximum
                }
            }
            $maxPopVal = ($script:todayForecasts | Where-Object {$null -ne $_.pop} | Measure-Object -Property pop -Maximum -EA SilentlyContinue).Maximum; $script:maxPop = if ($maxPopVal){[int]($maxPopVal*100)}else{0}
            $mF = $script:todayForecasts|Where-Object{$null -ne $_.dt -and ([datetimeoffset]::FromUnixTimeSeconds($_.dt)).DateTime.ToLocalTime().Hour -lt 14}|Select-Object -First 1
            $aF = $script:todayForecasts|Where-Object{$null -ne $_.dt -and ([datetimeoffset]::FromUnixTimeSeconds($_.dt)).DateTime.ToLocalTime().Hour -ge 14}|Select-Object -First 1
            $script:morningDesc = if($mF -and $mF.weather){$mF.weather[0].description.Substring(0,1).ToUpper()+$mF.weather[0].description.Substring(1)}else{"N/D"}; $script:afternoonDesc = if($aF -and $aF.weather){$aF.weather[0].description.Substring(0,1).ToUpper()+$aF.weather[0].description.Substring(1)}else{"N/D"}

            $minTStr = if ($null -ne $script:minTempForDay) { $script:minTempForDay.ToString('F1') } else { "N/D" }; $maxTStr = if ($null -ne $script:maxTempForDay) { $script:maxTempForDay.ToString('F1') } else { "N/D" }
            Add-ColoredText -Control $rtb -Text "📊 Riepilogo Giornaliero:`r`n" -Bold $true; Add-ColoredText -Control $rtb -Text "    🌡️ Temp: Min ~$minTStr°C / Max ~$maxTStr°C`r`n"
            Add-ColoredText -Control $rtb -Text "    🌦️ Cond: Mattina ~$($script:morningDesc) / Pomeriggio ~$($script:afternoonDesc)`r`n"; Add-ColoredText -Control $rtb -Text "    💧 Prob. Max Precip: $($script:maxPop.ToString('F0'))%`r`n" -AddNewLine $true
            Add-ColoredText -Control $rtb -Text (("-" * 18) + " Dettaglio Previsioni (3 Ore) " + ("-" * 18) + "`r`n`r`n")

            foreach ($item in $script:todayForecasts) {
                if (-not ($item.main -and $item.weather -and $item.weather.Count -gt 0 -and $item.wind)) { continue }
                $dateTime = ([datetimeoffset]::FromUnixTimeSeconds($item.dt)).DateTime.ToLocalTime(); $timeStr = $dateTime.ToString("HH:mm")
                $temp = $item.main.temp; $feelsLike = $item.main.feels_like; $description = $item.weather[0].description.Substring(0,1).ToUpper()+$item.weather[0].description.Substring(1)
                $humidity = $item.main.humidity; $pressure = $item.main.pressure; $clouds = $item.clouds.all
                $windSpeed_mps = $item.wind.speed; $windSpeed_kmh = ($windSpeed_mps * 3.6); $windDirectionDeg = $item.wind.deg; $windDirectionStr = Convert-WindDirection -degrees $windDirectionDeg
                $pop = ($item.pop * 100); $weatherIconDisplay = Get-WeatherIcon -weatherMain $item.weather[0].main -description $description
                $tempColor = $cBlack; $windColor = $cBlack; $rainColor = $cBlack; $alertPrefix = ""

                if ($temp -lt $ThresholdTempLowC) { $tempColor=$cBlue; $script:alertMessages.Add("[$timeStr] Temp. Bassa: $($temp.ToString('F1'))°C"); if ($script:operativityStatus -ne "CRITICO") {$script:operativityStatus="ATTENZIONE"; $script:CausedByTemp=$true} }
                if ($temp -gt $ThresholdTempHighC) { $tempColor=$cRed; $script:alertMessages.Add("[$timeStr] Temp. Alta: $($temp.ToString('F1'))°C"); if ($script:operativityStatus -ne "CRITICO") {$script:operativityStatus="ATTENZIONE"; $script:CausedByTemp=$true} }
                if ($windSpeed_kmh -gt $ThresholdWindKmh) { $windColor=$cOrange; $script:alertMessages.Add("[$timeStr] Vento Forte: $($windSpeed_kmh.ToString('F1')) km/h da $windDirectionStr"); if ($script:operativityStatus -ne "CRITICO") {$script:operativityStatus="ATTENZIONE"; $script:CausedByWind=$true} }
                
                $rainAmount = 0;
                if ($item.rain -and $item.rain.PSObject.Properties.Contains("3h")) { $rainAmount = $item.rain."3h" }
                elseif ($item.snow -and $item.snow.PSObject.Properties.Contains("3h")) { $rainAmount = $item.snow."3h" }

                if ($rainAmount -gt 0) {
                    $rainColor=$cBlueViolet;
                    if ($rainAmount -gt $ThresholdRainMm) { # $ThresholdRainMm ora è per mm/3h
                        $rainColor=$cRed; $alertPrefix="⚠️ "; $script:alertMessages.Add("[$timeStr] Precipitazioni Int. (3h): $($rainAmount.ToString('F1')) mm"); $script:operativityStatus="CRITICO"; $script:CausedByRain=$true
                    } elseif ($script:operativityStatus -ne "CRITICO") { $script:operativityStatus="ATTENZIONE"; $script:CausedByRain=$true }
                }

                Add-ColoredText -Control $rtb -Text "$alertPrefix" -Color $rainColor -Bold $true; Add-ColoredText -Control $rtb -Text "[$timeStr] $weatherIconDisplay $description`r`n" -Bold $true
                Add-ColoredText -Control $rtb -Text "    🌡️ Temp: "; Add-ColoredText -Control $rtb -Text "$($temp.ToString('F1'))°C" -Color $tempColor; Add-ColoredText -Control $rtb -Text " (Perc.: $($feelsLike.ToString('F1'))°C)`r`n"
                Add-ColoredText -Control $rtb -Text "    💧 Umidità: $humidity%|Press: $pressure hPa|Nuvole: $clouds%|P.Precip: $($pop.ToString('F0'))%`r`n"
                Add-ColoredText -Control $rtb -Text "    🌬️ Vento: "; Add-ColoredText -Control $rtb -Text "$($windSpeed_kmh.ToString('F1')) km/h" -Color $windColor; Add-ColoredText -Control $rtb -Text " ($($windSpeed_mps.ToString('F1')) m/s) da $windDirectionStr ($windDirectionDeg°)`r`n"
                if ($rainAmount -gt 0) {
                    $precipTypeDesc = if ($item.snow -and $item.snow.PSObject.Properties.Contains("3h")){"Neve"}else{"Pioggia"}
                    Add-ColoredText -Control $rtb -Text "    🌧️ $precipTypeDesc (3h): "; Add-ColoredText -Control $rtb -Text "$($rainAmount.ToString('F1')) mm" -Color $rainColor -AddNewLine $true
                }
                Add-ColoredText -Control $rtb -Text "`r`n"
            }
        } else { Add-ColoredText -Control $rtb -Text "Nessuna previsione dettagliata per oggi (intervalli 3 ore).`r`n" -Color $cOrangeRed -AddNewLine $true; if ($script:operativityStatus -eq "OK") {$script:operativityStatus="DATI PARZIALI"} }
    } elseif ($script:operativityStatus -ne "ERRORE DATI") { Add-ColoredText -Control $rtb -Text "Impossibile elaborare previsioni (3 ore).`r`n" -Color $cRed -Bold $true -AddNewLine $true }

    switch($script:operativityStatus){ "OK"{$script:operativityColor=$cGreen};"ATTENZIONE"{$script:operativityColor=$cOrange};"CRITICO"{$script:operativityColor=$cRed};"DATI PARZIALI"{$script:operativityColor=$cOrangeRed};"ERRORE DATI"{$script:operativityColor=$cRed}; default{$script:operativityColor=$cBlack} }
    $statusReason = ""
    if ($script:operativityStatus -ne "OK" -and $script:operativityStatus -ne "ERRORE DATI" -and $script:operativityStatus -ne "DATI PARZIALI") {
        $reasonsList = @(); if ($script:CausedByRain) { $reasonsList += "Precipitazioni" }; if ($script:CausedByWind) { $reasonsList += "Vento" }; if ($script:CausedByTemp) { $reasonsList += "Temperatura" }
        if ($reasonsList.Count -gt 0) { $statusReason = " (Causa: $($reasonsList -join ', '))" }
    }
    Add-ColoredText -Control $rtb -Text (("-" * 25) + " Stato Operatività " + ("-" * 25) + "`r`n")
    Add-ColoredText -Control $rtb -Text "Stato Generale: "; Add-ColoredText -Control $rtb -Text "$script:operativityStatus$statusReason" -Color $script:operativityColor -Bold $true -AddNewLine $true
    if ($script:alertMessages.Count -gt 0) { Add-ColoredText -Control $rtb -Text "Avvisi Dettagliati:`r`n" -Bold $true; foreach ($msg in $script:alertMessages) { $alertC=if($msg-like"*Intensa*"){$cRed}else{$cOrange}; Add-ColoredText -Control $rtb -Text " - $msg`r`n" -Color $alertC } } elseif ($script:operativityStatus -eq "OK") { Add-ColoredText -Control $rtb -Text "Nessun avviso significativo rilevato.`r`n" -Color $cGreen }
    Add-ColoredText -Control $rtb -Text (("-" * 75) + "`r`n")
    Add-ColoredText -Control $rtb -Text "Script v2.14 (3 Ore). Dati da OpenWeatherMap.`r`n" -Color $cGray -AddNewLine $true

    Set-ButtonPanelLayout -panel $buttonPanel -btn1 $okButton -btn2 $openReportButton
    $rtb.SelectionStart = 0; $rtb.ScrollToCaret()
    $this.Cursor = [System.Windows.Forms.Cursors]::Default
    Write-Host "[Load] Popolamento RichTextBox completato."
}

$form.Add_Load($form_Load)

$openReportButton.add_Click({
    $thisButton = $this
    $thisButton.Enabled = $false; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Write-Host "Generazione report Word (3 Ore) richiesta..."

    if (($null -eq $script:todayForecasts -or $script:todayForecasts.Count -eq 0) -and $script:operativityStatus -ne "ERRORE DATI") {
        [System.Windows.Forms.MessageBox]::Show($form, "Dati meteo (3 ore) non disponibili o errore API. Impossibile generare il report.", "Dati Mancanti", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); $thisButton.Enabled = $true; $form.Cursor = [System.Windows.Forms.Cursors]::Default; return
    }

    $reportInputData = @{
        CityName=$script:cityName; Country=$script:country; Sunrise=$script:sunrise; Sunset=$script:sunset
        MinTemp=$script:minTempForDay; MaxTemp=$script:maxTempForDay; MinHumidity=$script:minHumidity; MaxHumidity=$script:maxHumidity; AvgHumidity=$script:avgHumidity
        MinPressure=$script:minPressure; MaxPressure=$script:maxPressure; AvgPressure=$script:avgPressure; MinVisibility=$script:minVisibility; MaxVisibility=$script:maxVisibility
        MorningDesc=$script:morningDesc; AfternoonDesc=$script:afternoonDesc; MaxPop=$script:maxPop; AqiData=$script:aqiData; AqiInfo=$script:aqiInfo
        Forecasts=$script:todayForecasts; Alerts=$script:alertMessages; OperativityStatus=$script:operativityStatus
        Thresholds=@{ RainMm=$ThresholdRainMm; WindKmh=$ThresholdWindKmh; TempLowC=$ThresholdTempLowC; TempHighC=$ThresholdTempHighC }
        CausedByRain=$script:CausedByRain; CausedByWind=$script:CausedByWind; CausedByTemp=$script:CausedByTemp
    }
    $generatedFilePath = $null; $reportSuccess = $false
    try {
        Write-Host "[Click] Controllo Word..."; try { $testWord = New-Object -ComObject Word.Application -EA Stop; $testWord.Quit(); Remove-Variable testWord } catch { throw "MS Word non installato/accessibile." }
        Write-Host "[Click] Chiamata a New-WordReport..."; $reportSuccess = New-WordReport -ReportData $reportInputData -ErrorAction Stop
        
        if (-not $reportSuccess) { throw "La funzione New-WordReport ha indicato un fallimento." }

        $safeCityName = ($script:cityName -replace '[\\/:*?"<>|]', '').ToUpper()
        $expectedFileName = "REPORT_METEO_3H_$($safeCityName)_$((Get-Date).ToString('dd_MM_yyyy_HH_mm')).docx" # Nome file aggiornato
        $reportPath = "C:\Users\Coemi\Desktop\REPORT METEO"
        $generatedFilePath = Join-Path -Path $reportPath -ChildPath $expectedFileName
        Write-Host "[Click] Percorso file atteso: '$generatedFilePath'"

    } catch { Write-Warning "[Click] Errore Generazione/Validazione: $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show($form, "Errore generazione report:`n$($_.Exception.Message)", "Errore Report", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); $generatedFilePath = $null }
    finally { $thisButton.Enabled = $true; $form.Cursor = [System.Windows.Forms.Cursors]::Default }

    if ($reportSuccess -and -not [string]::IsNullOrWhiteSpace($generatedFilePath)) {
        Write-Host "[Click] Verifica esistenza file: '$generatedFilePath'"
        $resolvedPath = $null; try { $resolvedPath = Resolve-Path -Path $generatedFilePath -ErrorAction Stop } catch { Write-Warning "[Click] Impossibile risolvere '$generatedFilePath': $($_.Exception.Message)" }

        if ($null -ne $resolvedPath) {
            Write-Host "[Click] Tentativo Invoke-Item: '$($resolvedPath.Path)'"
            try { Invoke-Item -Path $resolvedPath.Path -ErrorAction Stop }
            catch { [System.Windows.Forms.MessageBox]::Show($form, "Report generato ma impossibile aprirlo.`nFile: '$($resolvedPath.Path)'`nErrore: $($_.Exception.Message)", "Apertura Fallita", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)}
        } else { [System.Windows.Forms.MessageBox]::Show($form, "Report generato ma il percorso '$generatedFilePath' non è valido o il file non è stato trovato.", "Errore Percorso/File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    } else { Write-Host "[Click] Generazione report fallita o percorso non valido, apertura saltata." }
})

$buttonPanel.add_Resize({ Set-ButtonPanelLayout -panel $buttonPanel -btn1 $okButton -btn2 $openReportButton })

# --- Visualizzazione Form ---
$richTextBox.SelectionStart = 0; $richTextBox.ScrollToCaret()
Write-Host "Visualizzazione Finestra Super Meteo (3 Ore)..."
$form.TopMost = $true;
$form.ShowDialog() | Out-Null
$form.Dispose(); Write-Host "Script Super Meteo (3 Ore) terminato."