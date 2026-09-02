# MailStore-PRTG

PRTG-Sensor (`EXE/Script Advanced`), der die **Archivierungs-/Export-Profile** und die
**geplanten Jobs** eines MailStore Servers überwacht.

Das Skript spricht die MailStore **Administration API** (HTTPS, Port 8463) direkt an.
Es wird **kein** zusätzliches PowerShell-Modul benötigt (kein `MS.PS.Lib`) –
eine einzige Datei auf die PRTG-Probe kopieren, Sensor anlegen, fertig.

---

## Inhalt

| Datei                | Zweck                                       |
|----------------------|---------------------------------------------|
| `MailStore-PRTG.ps1` | Das Sensor-Skript                           |
| `README.md`          | Diese Anleitung                             |

---

## Was der Sensor prüft

Mit jeweils **einem** API-Call werden alle Profil-Ausführungen (`GetWorkerResults`) und
alle Job-Ausführungen (`GetJobResults`) geholt und ausgewertet. Der Sensor arbeitet mit
**zwei getrennten Zeitfenstern**:

| Fenster            | Parameter          | Standard | Zweck                                                                                                                   |
|--------------------|--------------------|----------|-------------------------------------------------------------------------------------------------------------------------|
| Kurzes Fenster     | `-LookbackMinutes` | 30 Min.  | Zählt Fehlerzustände. Ein einmaliger Fehler fällt nach 30 Minuten wieder heraus – der Sensor wird **von allein wieder grün**. |
| Langes Fenster     | `-OverdueHours`    | 24 Std.  | Erkennt Profile, die komplett stehen geblieben sind. Dieser Zustand darf **nicht** nach 30 Minuten verschwinden, sonst bleibt er unbemerkt. |

Genau diese Trennung ist der Kern des Skripts: kurzfristige Fehler quittieren sich selbst,
ein dauerhaft nicht mehr laufendes Profil bleibt sichtbar.

---

## Voraussetzungen

**Auf dem MailStore Server**

1. **Administration API aktiviert**
   MailStore Server Service Configuration → *Netzwerkeinstellungen* → *Administration API*
   aktivieren (Standard-Port **8463**).
2. **Benutzer mit API-Recht**
   Ein MailStore-Benutzer mit der Rolle *Administrator* und dem Login-Privileg **`api`**.
   Empfehlung: ein eigener, nur für das Monitoring angelegter Benutzer.
3. **Port 8463/TCP** von der PRTG-Probe aus erreichbar (Firewall).

**Auf der PRTG-Probe**

- PRTG Network Monitor (Sensortyp *EXE/Script Advanced*).
- Windows PowerShell **5.1** oder PowerShell **7+** (beides wird unterstützt).
- TLS 1.2 aktiv (das Skript erzwingt TLS 1.2/1.1/1.0 selbst).

---

## Installation

1. `MailStore-PRTG.ps1` auf die **Probe** kopieren, auf der der Sensor laufen soll:

   ```
   C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\EXEXML\MailStore-PRTG.ps1
   ```

   > Bei mehreren Proben bzw. im Failover-Cluster muss die Datei auf **jede** Probe
   > und jeden Cluster-Knoten kopiert werden.

2. Datei-Sperre entfernen (falls aus dem Web/per Mail übertragen):

   ```powershell
   Unblock-File 'C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\EXEXML\MailStore-PRTG.ps1'
   ```

3. Ausführungsrichtlinie setzen (einmalig pro Probe, als Administrator):

   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
   ```

   Läuft auf der Probe noch eine 32-Bit-PowerShell mit, die Richtlinie dort ebenfalls setzen:
   `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`

---

## Sensor in PRTG anlegen

1. Auf dem Gerät (dem MailStore-Server) → **Sensor hinzufügen** → **EXE/Script Advanced**.
2. Einstellungen:

   | Feld                          | Wert                                                                                             |
   |-------------------------------|--------------------------------------------------------------------------------------------------|
   | **EXE/Script**                | `MailStore-PRTG.ps1`                                                                              |
   | **Parameter**                 | `-User "%windowsuser" -Password "%windowspassword" -MailStoreServer "%host"`                       |
   | **Environment**               | *Default environment*                                                                              |
   | **Security Context**          | *Use security context of probe service* (die Anmeldung erfolgt über die Parameter, nicht über Windows) |
   | **Mutex Name**                | leer                                                                                               |
   | **Timeout (Sek.)**            | 60 (muss über `-TimeoutSec` liegen, Standard 45)                                                   |
   | **Result Handling**           | *Discard result* (zur Fehlersuche vorübergehend *Write result to disk*)                            |
   | **Scanning Interval**         | 5 – 10 Minuten                                                                                     |

3. Speichern. Beim ersten Durchlauf legt PRTG alle Kanäle automatisch an.

**Zu den Platzhaltern:** `%host`, `%windowsuser` und `%windowspassword` füllt PRTG aus den
Geräte-Einstellungen. Wer die MailStore-Zugangsdaten nicht als Windows-Anmeldedaten des
Geräts hinterlegen möchte, trägt sie direkt ein:

```
-User "prtg-monitor" -Password "GeheimesKennwort" -MailStoreServer "mailstore.firma.local"
```

Enthält das Passwort Leerzeichen oder Sonderzeichen, unbedingt in **doppelte
Anführungszeichen** setzen.

---

## Parameter

| Parameter                  | Typ      | Standard    | Bedeutung                                                                                       |
|----------------------------|----------|-------------|--------------------------------------------------------------------------------------------------|
| `-User`                    | string   | –           | **Pflicht.** MailStore-Administrator mit Login-Privileg `api`.                                     |
| `-Password`                | string   | –           | **Pflicht.** Passwort dieses Benutzers.                                                            |
| `-MailStoreServer`         | string   | `localhost` | Hostname/IP des MailStore Servers.                                                                 |
| `-Port`                    | int      | `8463`      | Port der Administration API.                                                                       |
| `-LookbackMinutes`         | int      | `30`        | Kurzes Fenster (Minuten), in dem Fehlerergebnisse gezählt werden.                                  |
| `-OverdueHours`            | int      | `24`        | Langes Fenster (Stunden) für „Profil ohne Ausführung“. `0` deaktiviert diese Prüfung.              |
| `-TimeZoneId`              | string   | `$Local`    | Zeitzone für die API. `$Local` = Zeitzone des MailStore-Servers.                                    |
| `-RequireValidCertificate` | switch   | aus         | Erzwingt ein gültiges Zertifikat. Ohne diesen Schalter werden Zertifikatsfehler ignoriert (MailStore nutzt meist ein selbstsigniertes Zertifikat). |
| `-OnlyAutomaticProfiles`   | switch   | aus         | Wertet nur Profile mit serverseitiger Automatik aus; manuell gestartete Läufe werden ignoriert.     |
| `-ExcludeProfileIds`       | string   | leer        | Kommaseparierte Profil-IDs, die komplett ignoriert werden, z. B. `"4,7"`.                          |
| `-NoJobs`                  | switch   | aus         | Job-Auswertung abschalten (nur Archivierungsprofile prüfen).                                       |
| `-TimeoutSec`              | int      | `45`        | Timeout je HTTP-Aufruf.                                                                            |
| `-DumpRaw`                 | switch   | aus         | Schreibt die API-Rohdaten als JSON nach `%TEMP%` – nur zur Diagnose.                               |

---

## Kanäle

| Kanal                            | Einheit         | Grenzwert            | Bedeutung                                                            |
|----------------------------------|-----------------|----------------------|------------------------------------------------------------------------|
| Archivierung erfolgreich         | Laeufe          | –                    | Erfolgreiche Profilausführungen im kurzen Fenster                      |
| Archivierung mit Warnungen       | Laeufe          | **Warnung** ab 1     | `completedWithWarnings`                                                |
| Archivierung mit Fehlern         | Laeufe          | **Fehler** ab 1      | `completedWithErrors`                                                  |
| Archivierung fehlgeschlagen      | Laeufe          | **Fehler** ab 1      | `failed`                                                               |
| Archivierung abgebrochen         | Laeufe          | **Fehler** ab 1      | `cancelled`, `disconnected`, `threadAbort`                             |
| Archivierung unbekannt           | Laeufe          | **Warnung** ab 1     | Status, den das Skript nicht kennt (→ mit `-DumpRaw` prüfen)           |
| Archivierung ohne Ausführung¹    | Archivierung    | **Fehler** ab 1      | Automatische Profile ohne Lauf innerhalb von `-OverdueHours`           |
| Letzte Ausführung vor            | Stunden         | Fehler > `OverdueHours`¹ | Alter der jüngsten Profil-/Job-Ausführung                          |
| Automatische Archivierung        | Archivierung    | –                    | Anzahl der Profile mit serverseitiger Automatik (ohne Ausschlüsse)     |
| Geplante Jobs erfolgreich²       | Laeufe          | –                    | Erfolgreiche Job-Ausführungen im kurzen Fenster                        |
| Geplante Jobs fehlerhaft²        | Laeufe          | **Fehler** ab 1      | Jobs mit einem Ergebnis ungleich `succeeded`                           |
| Geplante Jobs konfiguriert²      | Geplante Jobs   | –                    | Anzahl der aktivierten Jobs                                            |

¹ nur wenn `-OverdueHours` größer 0 ist.
² nur wenn der Server `GetJobs`/`GetJobResults` unterstützt und `-NoJobs` nicht gesetzt ist.

Die **Sensormeldung** listet die betroffenen Profil- bzw. Job-Namen im Klartext auf,
z. B. `Exchange Postfaecher (Fehler); Journal-Export (keine Ausfuehrung seit 24 h)`.
Ohne Befund steht dort `Letzte 30 Min. ohne Fehler (…)`.

Bei einem Verbindungs- oder Anmeldefehler liefert das Skript ein PRTG-Fehlerergebnis
(`<error>1</error>`) mit Klartextmeldung statt eines Absturzes – der Sensor wird rot und
nennt den Grund.

---

## Beispiele

Standardüberwachung, MailStore auf demselben Gerät wie in PRTG hinterlegt:

```
-User "%windowsuser" -Password "%windowspassword" -MailStoreServer "%host"
```

Nur automatische Profile, engeres Fehlerfenster, Testprofile 4 und 7 ausgenommen:

```
-User "prtg" -Password "geheim" -MailStoreServer "mailstore.firma.local" -LookbackMinutes 15 -OnlyAutomaticProfiles -ExcludeProfileIds "4,7"
```

Nur Archivierungsprofile, keine Jobs, „stehen geblieben“ erst nach 48 Stunden:

```
-User "prtg" -Password "geheim" -MailStoreServer "10.0.0.20" -NoJobs -OverdueHours 48
```

Gültiges Zertifikat erzwingen (empfohlen, sobald am MailStore ein offizielles
Zertifikat hinterlegt ist):

```
-User "prtg" -Password "geheim" -MailStoreServer "mailstore.firma.local" -RequireValidCertificate
```

Abweichender Port und Prüfung „Profil ohne Ausführung“ abgeschaltet:

```
-User "prtg" -Password "geheim" -MailStoreServer "mailstore.firma.local" -Port 8465 -OverdueHours 0
```

---

## Manueller Test auf der Probe

Vor dem Anlegen des Sensors lohnt sich ein Testlauf direkt auf der Probe:

```powershell
cd 'C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\EXEXML'
.\MailStore-PRTG.ps1 -User 'prtg' -Password 'geheim' -MailStoreServer 'mailstore.firma.local'
```

Erwartet wird ein XML-Block, der mit `<prtg>` beginnt und mit `</prtg>` endet.
Wichtig: Es darf **keine** weitere Ausgabe vor `<prtg>` stehen, sonst verwirft PRTG das Ergebnis.

---

## Fehlersuche

| Meldung im Sensor                                             | Ursache                                                      | Lösung                                                                                      |
|---------------------------------------------------------------|--------------------------------------------------------------|---------------------------------------------------------------------------------------------|
| `Anmeldung abgelehnt (HTTP 401)`                              | Benutzer/Passwort falsch oder Login-Privileg `api` fehlt      | Zugangsdaten prüfen, in MailStore beim Benutzer das Privileg **`api`** setzen                |
| `Zugriff verweigert (HTTP 403)`                               | Benutzer ist kein Administrator                               | Rolle *Administrator* zuweisen                                                               |
| `API-Endpunkt nicht gefunden (HTTP 404)`                      | Ältere MailStore-Version kennt die Funktion nicht             | MailStore aktualisieren oder `-NoJobs` verwenden                                             |
| `Verbindung zu https://… nicht moeglich`                      | DNS, Port, Firewall oder TLS-Handshake                        | Administration API aktiviert? Port 8463 offen? TLS 1.2 am Server aktiv?                      |
| `Der Server hat keine API-Funktionsliste geliefert`           | Administration API nicht aktiv, falscher Port/Dienst          | Service Configuration prüfen                                                                 |
| `Antwort von '…' ist kein gueltiges JSON`                     | Der Port beantwortet etwas anderes (z. B. Weboberfläche)      | Port der Administration API prüfen (`-Port`)                                                 |
| `Zertifikatspruefung konnte nicht deaktiviert werden`         | `Add-Type` auf der Probe blockiert                            | Gültiges Zertifikat am MailStore hinterlegen und mit `-RequireValidCertificate` betreiben    |
| Kanal `Archivierung unbekannt` steht auf 1 oder höher         | Ergebnisstatus, den das Skript nicht kennt                    | Einmal mit `-DumpRaw` laufen lassen, JSON aus `%TEMP%` prüfen und den Status ergänzen        |
| Sensor meldet `keine Ausfuehrung seit 24 h`, obwohl Profile laufen | Profile laufen seltener als 24 h oder werden manuell gestartet | `-OverdueHours` erhöhen, betroffene Profile über `-ExcludeProfileIds` ausnehmen              |
| PRTG meldet „Response not well-formed: (XML)“                 | Zusätzliche Ausgabe vor dem XML (z. B. Profilskript der Probe) | Skript manuell aufrufen und prüfen, was zusätzlich ausgegeben wird                           |

Zur Diagnose in PRTG zusätzlich **Result Handling → Write result to disk** aktivieren; die
Rohantwort landet dann unter
`C:\ProgramData\Paessler\PRTG Network Monitor\Logs\sensors\`.

`-DumpRaw` schreibt eine Datei `mailstore_prtg_dump_<server>.json` nach `%TEMP%` des
Kontos, unter dem der Probe-Dienst läuft (bei `LocalSystem`:
`C:\Windows\Temp` bzw. `C:\Windows\SysWOW64\config\systemprofile\AppData\Local\Temp`).

---

## Sicherheitshinweise

- Für das Monitoring einen **eigenen MailStore-Benutzer** mit Administrator-Rolle und
  Login-Privileg `api` anlegen, kein persönliches Konto verwenden.
- Das Passwort steht im Parameterfeld des Sensors. Wo möglich `%windowspassword`
  verwenden, damit es nur in den Geräte-Anmeldedaten von PRTG liegt.
- Standardmäßig werden Zertifikatsfehler ignoriert, weil MailStore ab Werk ein
  selbstsigniertes Zertifikat verwendet. Sobald ein vertrauenswürdiges Zertifikat
  hinterlegt ist, den Sensor mit `-RequireValidCertificate` betreiben.
- Der Sensor liest ausschließlich (`GetProfiles`, `GetWorkerResults`, `GetJobs`,
  `GetJobResults`) – es werden keine Profile, Jobs oder Archivdaten verändert.

---

## Funktionsweise

1. `api/get-metadata` – ermittelt, welche API-Funktionen der Server kennt
   (davon hängt ab, ob die Job-Kanäle angelegt werden).
2. `GetProfiles` – Namen der Profile und die Kennzeichnung „serverseitige Automatik“.
3. `GetWorkerResults` – alle Profilausführungen im Abfragefenster, ein einziger Aufruf.
4. `GetJobs` / `GetJobResults` – Jobs und deren Ausführungen (entfällt bei `-NoJobs`).

Lang laufende Aufrufe werden über `api/get-status` gepollt, bis der Server einen
Endstatus meldet. Das Abfragefenster ist immer das größere der beiden Zeitfenster;
gezählt wird anschließend getrennt nach kurzem und langem Fenster.
Beim Abrufen wird 5 Minuten in die Zukunft gepuffert, um kleine Zeitdifferenzen
zwischen Probe und MailStore-Server auszugleichen.

---

## Kompatibilität

- Windows PowerShell 5.1 und PowerShell 7+
  (unter 5.1 wird die Zertifikatsprüfung über einen kompilierten Delegaten abgeschaltet,
  unter 7+ über `-SkipCertificateCheck`).
- MailStore Server mit aktivierter Administration API. Die Job-Kanäle erscheinen nur auf
  Versionen, die `GetJobs`/`GetJobResults` unterstützen.
- Feldnamen der API werden versionstolerant gelesen (z. B. `profileID`/`profileId`).
