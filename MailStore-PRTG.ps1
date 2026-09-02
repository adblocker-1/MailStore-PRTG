<#
.SYNOPSIS
    PRTG EXE/Script Advanced Sensor: prueft MailStore Archivierungs-/Export-Profile
    und MailStore Jobs. Komplett eigenstaendig, ohne MS.PS.Lib.

.DESCRIPTION
    Spricht die MailStore Administration API (HTTPS, Port 8463) direkt an.
    Es wird KEIN zusaetzliches PowerShell-Modul benoetigt - eine Datei kopieren, fertig.

    Holt alle Profil-Ausfuehrungen (GetWorkerResults) und alle Job-Ausfuehrungen
    (GetJobResults) mit jeweils EINEM API-Call und wertet sie aus.

    Es gibt zwei getrennte Zeitfenster:
      -LookbackMinutes (Standard 30) - kurzes Fenster fuer Fehlerzustaende.
          Ein einmaliger Fehler faellt nach 30 Minuten wieder aus dem Sensor heraus,
          der Sensor wird also von allein wieder gruen.
      -OverdueHours (Standard 24)    - langes Fenster fuer die Frage, ob ein
          automatisches Profil komplett stehen geblieben ist. Dieser Zustand darf
          NICHT nach 30 Minuten verschwinden, sonst bleibt er unbemerkt.

.NOTES
    Ablage:  ...\Custom Sensors\EXEXML\MailStore-PRTG.ps1
    Sensor:  "EXE/Script Advanced"
    Voraussetzung: Administration API in der MailStore Dienstkonfiguration aktiviert,
                   verwendeter Benutzer ist Administrator mit Login-Privileg "api".

.EXAMPLE
    -User "%windowsuser" -Password "%windowspassword" -MailStoreServer "%host" -LookbackMinutes 30
#>

[CmdletBinding()]
param(
    # MailStore Administrator (Login-Privileg "api" erforderlich)
    [Parameter(Mandatory = $true)]
    [string]$User,

    [Parameter(Mandatory = $true)]
    [string]$Password,

    [string]$MailStoreServer = 'localhost',

    [int]$Port = 8463,

    # Zeitfenster in MINUTEN, in dem Fehlerergebnisse gezaehlt werden.
    # Kurz halten, damit der Sensor nach einem einmaligen Fehler wieder gruen wird.
    [int]$LookbackMinutes = 30,

    # Zeitfenster in STUNDEN fuer "Profil ohne Ausfuehrung". 0 = Pruefung deaktiviert.
    [int]$OverdueHours = 24,

    # Zeitzone fuer die API. $Local = Zeitzone des MailStore-Servers
    [string]$TimeZoneId = '$Local',

    # Standard ist: Zertifikatsfehler ignorieren (MailStore nutzt meist ein self-signed Zert)
    [switch]$RequireValidCertificate,

    # Nur Profile mit serverseitiger Automatik zaehlen
    [switch]$OnlyAutomaticProfiles,

    # Kommaseparierte Profil-IDs, die ignoriert werden sollen, z.B. "4,7"
    [string]$ExcludeProfileIds = '',

    # Job-Auswertung abschalten
    [switch]$NoJobs,

    [int]$TimeoutSec = 45,

    # Rohdaten als JSON nach %TEMP% schreiben (zur Feldnamen-Diagnose)
    [switch]$DumpRaw
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'
$InvariantCulture      = [System.Globalization.CultureInfo]::InvariantCulture

# ===========================================================================
# API-Client (ersetzt MS.PS.Lib)
# ===========================================================================

$script:BaseUri    = ('https://{0}:{1}' -f $MailStoreServer, $Port)
$script:TimeoutSec = $TimeoutSec
$script:AuthHeader = @{
    Authorization = 'Basic ' + [Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes(('{0}:{1}' -f $User, $Password)))
}
$script:ExtraWebParams = @{}

function Enable-CertificateBypass {
    # In Windows PowerShell 5.1 darf hier KEIN Scriptblock stehen:
    #     [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    # Der TLS-Stack ruft den Callback auf einem eigenen Thread ohne Runspace auf.
    # Der Scriptblock kann dort nicht laufen, die Validierung wirft, und .NET meldet
    # das nach aussen als "Die zugrunde liegende Verbindung wurde geschlossen:
    # Unerwarteter Fehler beim Senden" - sieht aus wie TLS, ist aber der Callback.
    # Deshalb ein kompilierter Delegat, der ohne Runspace auskommt.
    if (-not ('MailStoreCertBypass' -as [type])) {
        try {
            Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

public static class MailStoreCertBypass
{
    public static void Enable()
    {
        ServicePointManager.ServerCertificateValidationCallback =
            delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
            {
                return true;
            };
    }
}
'@
        }
        catch {
            throw ("Zertifikatspruefung konnte nicht deaktiviert werden: {0}. " +
                   "Alternative: gueltiges Zertifikat am MailStore hinterlegen und " +
                   "den Sensor mit -RequireValidCertificate betreiben." -f $_.Exception.Message)
        }
    }
    [MailStoreCertBypass]::Enable()
}

function Initialize-MailStoreTransport {
    # TLS erzwingen - aeltere .NET-Defaults verhandeln sonst nur SSL3/TLS1.0.
    #
    # TLS 1.3 hier bewusst NICHT setzen. Der Enum-Wert "Tls13" existiert ab
    # .NET Framework 4.8, aber Windows Server vor 2022 kann TLS 1.3 ueber Schannel
    # nicht aushandeln; der Handshake bricht dann sofort ab.
    $protocols = [System.Net.SecurityProtocolType]::Tls12
    foreach ($name in @('Tls11', 'Tls')) {
        if ([Enum]::GetNames([System.Net.SecurityProtocolType]) -contains $name) {
            $protocols = $protocols -bor [System.Net.SecurityProtocolType]::$name
        }
    }
    [System.Net.ServicePointManager]::SecurityProtocol = $protocols

    if (-not $RequireValidCertificate) {
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            # PowerShell 7+ ignoriert den Callback, nutzt stattdessen den Parameter
            $script:ExtraWebParams['SkipCertificateCheck'] = $true
        } else {
            Enable-CertificateBypass
        }
    }
}

function Invoke-MailStoreRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [hashtable]$Body
    )

    $webParams = @{
        Uri                = ('{0}/{1}' -f $script:BaseUri, $Path)
        Method             = 'Post'
        Headers            = $script:AuthHeader
        UseBasicParsing    = $true
        TimeoutSec         = $script:TimeoutSec
        MaximumRedirection = 0
        ErrorAction        = 'Stop'
    }
    foreach ($k in $script:ExtraWebParams.Keys) { $webParams[$k] = $script:ExtraWebParams[$k] }
    if ($Body -and $Body.Count -gt 0) { $webParams['Body'] = $Body }

    try {
        $response = Invoke-WebRequest @webParams
    }
    catch {
        $webResponse = $null
        if ($_.Exception.PSObject.Properties.Name -contains 'Response') { $webResponse = $_.Exception.Response }

        if ($null -eq $webResponse) {
            # Transportfehler: DNS, Port, Firewall oder TLS-Handshake.
            # Die innerste Exception nennt meist den echten Grund.
            $detail = $_.Exception.Message
            $inner  = $_.Exception.InnerException
            while ($inner) {
                $detail = '{0} -> {1}' -f $detail, $inner.Message
                $inner  = $inner.InnerException
            }
            throw ("Verbindung zu {0} nicht moeglich: {1} (Pruefen: Administration API aktiviert, Port {2} erreichbar, TLS 1.2 am Server aktiv)" -f $script:BaseUri, $detail, $Port)
        }

        $statusCode = 0
        try { $statusCode = [int]$webResponse.StatusCode } catch { }

        switch ($statusCode) {
            401 { throw 'Anmeldung abgelehnt (HTTP 401). Benutzer, Passwort oder das Login-Privileg "api" pruefen.' }
            403 { throw 'Zugriff verweigert (HTTP 403). Der Benutzer besitzt keine Administrator-Rechte.' }
            404 { throw ("API-Endpunkt nicht gefunden (HTTP 404): {0}" -f $Path) }
            default {
                throw ("HTTP {0} beim Aufruf von '{1}': {2}" -f $statusCode, $Path, $_.Exception.Message)
            }
        }
    }

    # MailStore liefert die Antworten mit UTF-8-BOM aus, das ConvertFrom-Json nicht vertraegt.
    $text = $null
    if ($response.RawContentStream) {
        $null = $response.RawContentStream.Seek(0, [System.IO.SeekOrigin]::Begin)
        $reader = New-Object System.IO.StreamReader(
            $response.RawContentStream, [System.Text.Encoding]::UTF8, $true)
        $text = $reader.ReadToEnd()
        $reader.Dispose()
    } else {
        $text = [string]$response.Content
    }

    if ($null -eq $text) { return $null }
    $text = $text.TrimStart([char]0xFEFF).Trim()
    if ($text -eq '') { return $null }

    try {
        return ($text | ConvertFrom-Json)
    }
    catch {
        $preview = $text.Substring(0, [Math]::Min(200, $text.Length))
        throw ("Antwort von '{0}' ist kein gueltiges JSON: {1}" -f $Path, $preview)
    }
}

function Invoke-MailStoreApiCall {
    param(
        [Parameter(Mandatory = $true)][string]$Function,
        [hashtable]$Parameters = @{}
    )

    $body = @{}
    foreach ($key in $Parameters.Keys) {
        if ($null -ne $Parameters[$key] -and "$($Parameters[$key])" -ne '') {
            $body[$key] = "$($Parameters[$key])"
        }
    }

    $result = Invoke-MailStoreRequest -Path ('api/invoke/{0}' -f $Function) -Body $body

    # Lang laufende Aufrufe: solange pollen, bis der Server einen Endstatus meldet.
    $guard = 0
    while ($result -and "$($result.statusCode)" -eq 'running' -and $guard -lt 60) {
        $guard++
        $result = Invoke-MailStoreRequest -Path 'api/get-status' -Body @{
            token                  = "$($result.token)"
            lastKnownStatusVersion = "$($result.statusVersion)"
            millisecondsTimeout    = '5000'
        }
    }

    return $result
}

# ===========================================================================
# Hilfsfunktionen
# ===========================================================================

function ConvertTo-XmlText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Format-Number {
    param($Value, [switch]$AsFloat)
    if ($AsFloat) { return ([double]$Value).ToString('0.##', $InvariantCulture) }
    return ([int]$Value).ToString($InvariantCulture)
}

function New-PrtgChannel {
    param(
        [string]$Name,
        $Value,
        [string]$CustomUnit = '',
        [switch]$AsFloat,
        $LimitMaxError,
        $LimitMaxWarning,
        [string]$LimitErrorMsg = ''
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('  <result>')
    [void]$sb.AppendLine(('    <channel>{0}</channel>' -f (ConvertTo-XmlText $Name)))
    [void]$sb.AppendLine(('    <value>{0}</value>' -f (Format-Number -Value $Value -AsFloat:$AsFloat)))
    [void]$sb.AppendLine('    <unit>Custom</unit>')
    if ($CustomUnit) {
        [void]$sb.AppendLine(('    <customUnit>{0}</customUnit>' -f (ConvertTo-XmlText $CustomUnit)))
    }
    if ($AsFloat) { [void]$sb.AppendLine('    <float>1</float>') }
    else          { [void]$sb.AppendLine('    <float>0</float>') }
    [void]$sb.AppendLine('    <showChart>1</showChart>')
    [void]$sb.AppendLine('    <showTable>1</showTable>')
    [void]$sb.AppendLine('    <mode>Absolute</mode>')

    if ($null -ne $LimitMaxError -or $null -ne $LimitMaxWarning) {
        [void]$sb.AppendLine('    <limitmode>1</limitmode>')
        if ($null -ne $LimitMaxError) {
            [void]$sb.AppendLine(('    <limitmaxerror>{0}</limitmaxerror>' -f (Format-Number $LimitMaxError)))
        }
        if ($null -ne $LimitMaxWarning) {
            [void]$sb.AppendLine(('    <limitmaxwarning>{0}</limitmaxwarning>' -f (Format-Number $LimitMaxWarning)))
        }
        if ($LimitErrorMsg) {
            [void]$sb.AppendLine(('    <limiterrormsg>{0}</limiterrormsg>' -f (ConvertTo-XmlText $LimitErrorMsg)))
        }
    }
    [void]$sb.AppendLine('  </result>')
    $sb.ToString()
}

function Write-PrtgFatal {
    param([string]$Message)
    $t = ConvertTo-XmlText ($Message -replace '\s+', ' ')
    if ($t.Length -gt 900) { $t = $t.Substring(0, 900) + '...' }
    Write-Output "<prtg>`n  <error>1</error>`n  <text>$t</text>`n</prtg>"
    exit 0
}

# Liest die erste vorhandene Eigenschaft aus einem Objekt (Feldnamen variieren je MailStore-Version)
function Get-Prop {
    param($Object, [string[]]$Names)
    if ($null -eq $Object) { return $null }
    $available = @($Object.PSObject.Properties.Name)
    foreach ($n in $Names) {
        $match = $available | Where-Object { $_ -eq $n } | Select-Object -First 1
        if (-not $match) {
            $match = $available | Where-Object { $_ -ieq $n } | Select-Object -First 1
        }
        if ($match) {
            $v = $Object.$match
            if ($null -ne $v) { return $v }
        }
    }
    return $null
}

function ConvertTo-DateTimeOrNull {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $parsed = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::None
    if ([datetime]::TryParse("$Value", $InvariantCulture, $styles, [ref]$parsed)) { return $parsed }
    return $null
}

function Assert-MSResponse {
    param($Response, [string]$CallName)
    if ($null -eq $Response) { throw "Keine Antwort auf API-Aufruf '$CallName'." }
    $status = Get-Prop $Response @('statusCode', 'StatusCode')
    if ($status -and "$status" -ne 'succeeded') {
        $msg = Get-Prop (Get-Prop $Response @('error')) @('message')
        throw ("API-Aufruf '{0}' lieferte Status '{1}'{2}" -f $CallName, $status, $(if ($msg) { ": $msg" } else { '.' }))
    }
}

# ===========================================================================
# Hauptteil
# ===========================================================================

$channels    = New-Object System.Collections.Generic.List[string]
$problemList = New-Object System.Collections.Generic.List[string]

try {
    Initialize-MailStoreTransport

    # --- Welche API-Funktionen kennt dieser Server? ------------------------
    $metadata  = Invoke-MailStoreRequest -Path 'api/get-metadata'
    $supported = @($metadata | ForEach-Object { "$($_.name)" })
    if ($supported.Count -eq 0) {
        throw "Der Server hat keine API-Funktionsliste geliefert. Ist die Administration API aktiviert?"
    }

    # --- Zeitfenster -------------------------------------------------------
    # Zwei getrennte Fenster:
    #   countMinutes  - kurz. Hier werden Fehlerstatus gezaehlt, damit ein einmaliger
    #                   Fehler nach kurzer Zeit wieder aus dem Sensor verschwindet.
    #   overdueHours  - lang. Nur fuer die Frage "ist ein Profil komplett stehen geblieben".
    $countMinutes   = [Math]::Max(1, [Math]::Abs($LookbackMinutes))
    $overdueMinutes = [Math]::Abs($OverdueHours) * 60
    $fetchMinutes   = [Math]::Max($countMinutes, $overdueMinutes)

    $now         = Get-Date
    $fetchFrom   = $now.AddMinutes(-$fetchMinutes)
    $countFrom   = $now.AddMinutes(-$countMinutes)
    $overdueFrom = $now.AddMinutes(-$overdueMinutes)
    # 5 Minuten Puffer nach vorne: gleicht kleine Zeitdifferenzen Probe <-> MailStore-Server aus
    $fetchTo     = $now.AddMinutes(5)

    $fromString = $fetchFrom.ToString('yyyy-MM-ddTHH:mm:ss', $InvariantCulture)
    $toString   = $fetchTo.ToString('yyyy-MM-ddTHH:mm:ss', $InvariantCulture)

    # --- Profile -----------------------------------------------------------
    $excluded = @()
    if ($ExcludeProfileIds) {
        $excluded = @($ExcludeProfileIds -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $profileResponse = Invoke-MailStoreApiCall -Function 'GetProfiles' -Parameters @{ raw = 'true' }
    Assert-MSResponse $profileResponse 'GetProfiles'
    $allProfiles = @($profileResponse.result)

    $profileNames = @{}
    foreach ($p in $allProfiles) {
        $profileKey = "$(Get-Prop $p @('id'))"
        if ($profileKey) {
            $profileName = Get-Prop $p @('name', 'displayName', 'description')
            if (-not $profileName) { $profileName = "Profil $profileKey" }
            $profileNames[$profileKey] = "$profileName"
        }
    }

    $automaticProfiles = @($allProfiles | Where-Object {
        $_.serverSideExecution -and $_.serverSideExecution.automatic -eq $true -and
        $excluded -notcontains "$(Get-Prop $_ @('id'))"
    })
    $automaticIds = @($automaticProfiles | ForEach-Object { "$(Get-Prop $_ @('id'))" })

    # --- Profil-Ausfuehrungen ---------------------------------------------
    # Ein Call ohne profileID liefert alle Ergebnisse des Zeitfensters.
    $workerResponse = Invoke-MailStoreApiCall -Function 'GetWorkerResults' -Parameters @{
        fromIncluding = $fromString
        toExcluding   = $toString
        timeZoneID    = $TimeZoneId   # Achtung: hier grosses "ID"
    }
    Assert-MSResponse $workerResponse 'GetWorkerResults'
    $workerResults = @($workerResponse.result)

    # --- Auswertung --------------------------------------------------------
    $succeeded = 0; $withWarnings = 0; $withErrors = 0; $failed = 0; $aborted = 0; $unknownState = 0
    $lastRunTimes   = New-Object 'System.Collections.Generic.List[datetime]'
    $seenProfileIds = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($entry in $workerResults) {
        $profileId = "$(Get-Prop $entry @('profileID', 'profileId', 'profileid'))"
        if ($excluded -contains $profileId) { continue }
        if ($OnlyAutomaticProfiles -and ($automaticIds -notcontains $profileId)) { continue }

        $endTime = ConvertTo-DateTimeOrNull (Get-Prop $entry @('endTime', 'end', 'startTime', 'start'))
        if ($endTime) {
            $lastRunTimes.Add($endTime)
            if ($endTime -ge $overdueFrom -and $profileId) { [void]$seenProfileIds.Add($profileId) }
            # Ergebnisse ausserhalb des Zaehlfensters nicht mitzaehlen
            if ($endTime -lt $countFrom) { continue }
        } elseif ($profileId) {
            [void]$seenProfileIds.Add($profileId)
        }

        $displayName = if ($profileId -and $profileNames.ContainsKey($profileId)) { $profileNames[$profileId] } else { "Profil $profileId" }
        $state = "$(Get-Prop $entry @('result', 'status', 'state'))".ToLowerInvariant()

        switch ($state) {
            'succeeded'             { $succeeded++ }
            'completedwithwarnings' { $withWarnings++; $problemList.Add("$displayName (Warnungen)") }
            'completedwitherrors'   { $withErrors++;   $problemList.Add("$displayName (Fehler)") }
            'failed'                { $failed++;       $problemList.Add("$displayName (fehlgeschlagen)") }
            'cancelled'             { $aborted++;      $problemList.Add("$displayName (abgebrochen)") }
            'canceled'              { $aborted++;      $problemList.Add("$displayName (abgebrochen)") }
            'disconnected'          { $aborted++;      $problemList.Add("$displayName (Verbindung verloren)") }
            'threadabort'           { $aborted++;      $problemList.Add("$displayName (Thread-Abbruch)") }
            default {
                $unknownState++
                $problemList.Add("$displayName (unbekannter Status: $state)")
            }
        }
    }

    # --- Profile ohne Ausfuehrung -----------------------------------------
    $overdueProfiles = @()
    if ($OverdueHours -gt 0) {
        $overdueProfiles = @($automaticIds | Where-Object { -not $seenProfileIds.Contains($_) })
        foreach ($id in $overdueProfiles) {
            $n = if ($profileNames.ContainsKey($id)) { $profileNames[$id] } else { "Profil $id" }
            $problemList.Add("$n (keine Ausfuehrung seit $OverdueHours h)")
        }
    }

    # --- Jobs --------------------------------------------------------------
    $jobsSucceeded = 0
    $jobsFailed    = 0
    $jobsEnabled   = 0
    $jobResults    = @()

    $jobsAvailable = (-not $NoJobs) -and ($supported -contains 'GetJobs') -and ($supported -contains 'GetJobResults')
    if ($jobsAvailable) {
        $jobsResponse = Invoke-MailStoreApiCall -Function 'GetJobs'
        Assert-MSResponse $jobsResponse 'GetJobs'
        $allJobs = @($jobsResponse.result)

        $jobNames = @{}
        foreach ($j in $allJobs) {
            $jid = "$(Get-Prop $j @('id'))"
            if ($jid) {
                $jn = Get-Prop $j @('name', 'action')
                $jobNames[$jid] = "$(if ($jn) { $jn } else { "Job $jid" })"
            }
            if ((Get-Prop $j @('enabled')) -ne $false) { $jobsEnabled++ }
        }

        $jobResultResponse = Invoke-MailStoreApiCall -Function 'GetJobResults' -Parameters @{
            fromIncluding = $fromString
            toExcluding   = $toString
            timeZoneId    = $TimeZoneId   # Achtung: hier kleines "d"
        }
        Assert-MSResponse $jobResultResponse 'GetJobResults'
        $jobResults = @($jobResultResponse.result)

        foreach ($jr in $jobResults) {
            $endTime = ConvertTo-DateTimeOrNull (Get-Prop $jr @('endTime', 'end', 'startTime', 'start'))
            if ($endTime) {
                $lastRunTimes.Add($endTime)
                if ($endTime -lt $countFrom) { continue }
            }
            $jid   = "$(Get-Prop $jr @('jobId', 'jobID', 'id'))"
            $jname = if ($jobNames.ContainsKey($jid)) { $jobNames[$jid] } else { "Job $jid" }
            $state = "$(Get-Prop $jr @('result', 'status', 'state'))".ToLowerInvariant()

            if ($state -eq 'succeeded') {
                $jobsSucceeded++
            } else {
                $jobsFailed++
                $problemList.Add("$jname (Job: $state)")
            }
        }
    }

    # --- Letzte Ausfuehrung ------------------------------------------------
    $hoursSinceLastRun = [Math]::Round($fetchMinutes / 60, 2)
    if ($lastRunTimes.Count -gt 0) {
        $newest = ($lastRunTimes | Sort-Object -Descending | Select-Object -First 1)
        $delta  = ($now - $newest).TotalHours
        if ($delta -lt 0) { $delta = 0 }
        $hoursSinceLastRun = [Math]::Round($delta, 2)
    }

    # --- Rohdaten-Dump (Diagnose) -----------------------------------------
    if ($DumpRaw) {
        $dumpPath = Join-Path $env:TEMP ("mailstore_prtg_dump_{0}.json" -f ($MailStoreServer -replace '[^\w\.-]', '_'))
        [PSCustomObject]@{
            supportedFunctions = $supported
            profiles           = $allProfiles
            workerResults      = $workerResults
            jobResults         = $jobResults
        } | ConvertTo-Json -Depth 8 | Out-File -FilePath $dumpPath -Encoding UTF8 -Force
        $problemList.Add("Dump: $dumpPath")
    }

    # --- Kanaele -----------------------------------------------------------
    $channels.Add((New-PrtgChannel -Name 'Archivierung erfolgreich'      -Value $succeeded    -CustomUnit 'Laeufe'))
    $channels.Add((New-PrtgChannel -Name 'Archivierung mit Warnungen'    -Value $withWarnings -CustomUnit 'Laeufe' -LimitMaxWarning 0 -LimitErrorMsg 'Profilausfuehrung mit Warnungen'))
    $channels.Add((New-PrtgChannel -Name 'Archivierung mit Fehlern'      -Value $withErrors   -CustomUnit 'Laeufe' -LimitMaxError 0 -LimitErrorMsg 'Profilausfuehrung mit Fehlern abgeschlossen'))
    $channels.Add((New-PrtgChannel -Name 'Archivierung fehlgeschlagen'   -Value $failed       -CustomUnit 'Laeufe' -LimitMaxError 0 -LimitErrorMsg 'Profilausfuehrung fehlgeschlagen'))
    $channels.Add((New-PrtgChannel -Name 'Archivierung abgebrochen'      -Value $aborted      -CustomUnit 'Laeufe' -LimitMaxError 0 -LimitErrorMsg 'Profilausfuehrung abgebrochen oder Verbindung verloren'))
    $channels.Add((New-PrtgChannel -Name 'Archivierung unbekannt'        -Value $unknownState -CustomUnit 'Laeufe' -LimitMaxWarning 0 -LimitErrorMsg 'Unbekannter Ergebnisstatus - Skript pruefen'))
    if ($OverdueHours -gt 0) {
        $channels.Add((New-PrtgChannel -Name 'Archivierung ohne Ausfuehrung' -Value $overdueProfiles.Count -CustomUnit 'Archivierung' -LimitMaxError 0 -LimitErrorMsg ("Automatisches Profil ist seit {0} h nicht mehr gelaufen" -f $OverdueHours)))
        $channels.Add((New-PrtgChannel -Name 'Letzte Ausfuehrung vor'   -Value $hoursSinceLastRun -CustomUnit 'Stunden' -AsFloat -LimitMaxError $OverdueHours -LimitErrorMsg 'Es wurde ueberhaupt nichts mehr ausgefuehrt'))
    } else {
        $channels.Add((New-PrtgChannel -Name 'Letzte Ausfuehrung vor'   -Value $hoursSinceLastRun -CustomUnit 'Stunden' -AsFloat))
    }
    $channels.Add((New-PrtgChannel -Name 'Automatische Archivierung'     -Value $automaticIds.Count -CustomUnit 'Archivierung'))

    if ($jobsAvailable) {
        $channels.Add((New-PrtgChannel -Name 'Geplante Jobs erfolgreich'  -Value $jobsSucceeded -CustomUnit 'Laeufe'))
        $channels.Add((New-PrtgChannel -Name 'Geplante Jobs fehlerhaft'   -Value $jobsFailed    -CustomUnit 'Laeufe' -LimitMaxError 0 -LimitErrorMsg 'MailStore Job nicht erfolgreich'))
        $channels.Add((New-PrtgChannel -Name 'Geplante Jobs konfiguriert' -Value $jobsEnabled   -CustomUnit 'Geplante Jobs'))
    }

    # --- Sensormeldung -----------------------------------------------------
    if ($problemList.Count -gt 0) {
        $text = ($problemList | Select-Object -Unique) -join '; '
    } else {
        $text = "Letzte $countMinutes Min. ohne Fehler ($succeeded Archivierung, $jobsSucceeded Geplante Jobs erfolgreich)"
    }
    if ($text.Length -gt 1800) { $text = $text.Substring(0, 1800) + '...' }

    $output = New-Object System.Text.StringBuilder
    [void]$output.AppendLine('<prtg>')
    foreach ($c in $channels) { [void]$output.Append($c) }
    [void]$output.AppendLine(('  <text>{0}</text>' -f (ConvertTo-XmlText $text)))
    [void]$output.AppendLine('</prtg>')

    Write-Output $output.ToString()
    exit 0
}
catch {
    Write-PrtgFatal ("MailStore-Pruefung fehlgeschlagen: {0}" -f $_.Exception.Message)
}