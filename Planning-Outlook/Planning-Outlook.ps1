<#
.SYNOPSIS
    Application Planning CPDL : du fichier Excel au calendrier Outlook partagé, en une seule étape.

.DESCRIPTION
    Ce script démarre un petit serveur web local (accessible uniquement depuis ce poste)
    puis ouvre la page Planning-Outlook.html dans le navigateur.

    La page lit le planning Excel, contrôle les incohérences, puis demande au script
    de créer les rendez-vous directement dans le calendrier partagé via l'API COM Outlook :
      - titre, horaires, lieu et catégorie Outlook ;
      - collaborateur ajouté en participant « Obligatoire » ;
      - envoi de l'invitation facultatif ;
      - détection des doublons (même titre, même début).

    La liste des collaborateurs, les horaires et les catégories se modifient depuis la page.
    Ils sont enregistrés dans config.json, à côté du script.

    Sécurité :
      - écoute sur http://localhost uniquement (aucun accès depuis le réseau) ;
      - chaque appel doit présenter un jeton aléatoire, différent à chaque lancement ;
      - aucune donnée ne quitte le poste (seule la librairie de lecture Excel est
        téléchargée depuis cdnjs.cloudflare.com par le navigateur).

.PARAMETER Port
    Port local d'écoute. Par défaut : 8765.

.PARAMETER FichierConfig
    Chemin du fichier de configuration. Par défaut : config.json à côté du script.

.PARAMETER SansNavigateur
    N'ouvre pas le navigateur automatiquement (l'adresse est affichée dans la console).

.EXAMPLE
    .\Planning-Outlook.ps1

.EXAMPLE
    .\Planning-Outlook.ps1 -Port 9000

.NOTES
    Version   : 2.0
    Auteur    : DSI EFS – Tours
    Date      : Octobre 2026

    Prérequis :
      - Windows PowerShell 5.1 (ou PowerShell 7) ;
      - Outlook 2021 ou Microsoft 365 installé, configuré et ouvert sur le poste ;
      - droits d'édition sur le calendrier partagé cible.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = 'Port local d''écoute.')]
    [ValidateRange(1024, 65535)]
    [int]$Port = 8765,

    [Parameter(Mandatory = $false, HelpMessage = 'Chemin du fichier de configuration JSON.')]
    [string]$FichierConfig = (Join-Path $PSScriptRoot 'config.json'),

    [Parameter(Mandatory = $false, HelpMessage = 'Ne pas ouvrir le navigateur automatiquement.')]
    [switch]$SansNavigateur
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ──────────────────────────────────────────────────────────────────────────────
# Constantes Outlook COM
# ──────────────────────────────────────────────────────────────────────────────
$OL_APPOINTMENT_ITEM = 1   # olAppointmentItem (aussi DefaultItemType d'un dossier Calendrier)
$OL_MEETING          = 1   # olMeeting : le rendez-vous devient une réunion avec participants
$OL_REQUIRED         = 1   # olRequired : participant « Obligatoire »
$OL_BUSY             = 2   # olBusy

$FichierHtml = Join-Path $PSScriptRoot 'Planning-Outlook.html'
$Utf8SansBom = New-Object System.Text.UTF8Encoding($false)

$script:Outlook   = $null
$script:Namespace = $null
$script:Arret     = $false

# ──────────────────────────────────────────────────────────────────────────────
# Fonctions Outlook
# ──────────────────────────────────────────────────────────────────────────────

function Get-OutlookNamespace {
    <#
    .SYNOPSIS
        Renvoie l'espace de noms MAPI (connexion à Outlook au premier appel).
    #>
    if ($null -eq $script:Namespace) {
        try {
            $script:Outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        }
        catch {
            try {
                $script:Outlook = New-Object -ComObject Outlook.Application
            }
            catch {
                throw "Impossible de démarrer Outlook. Vérifiez qu'il est installé et correctement configuré."
            }
        }
        $script:Namespace = $script:Outlook.GetNamespace('MAPI')
    }
    return $script:Namespace
}

function Get-CalendriersRecursif {
    <#
    .SYNOPSIS
        Parcourt récursivement un dossier Outlook et collecte les sous-dossiers de type Calendrier.
    #>
    param(
        [Parameter(Mandatory)] [object]$Dossier,
        [Parameter(Mandatory)] [System.Collections.Generic.List[object]]$Liste
    )

    if ($Dossier.DefaultItemType -eq $OL_APPOINTMENT_ITEM) {
        $Liste.Add($Dossier)
    }

    foreach ($sousDossier in $Dossier.Folders) {
        try {
            Get-CalendriersRecursif -Dossier $sousDossier -Liste $Liste
        }
        catch {
            # Sous-dossier inaccessible (droits insuffisants) : on l'ignore
        }
    }
}

function Get-TousLesCalendriers {
    <#
    .SYNOPSIS
        Renvoie tous les dossiers Calendrier accessibles depuis Outlook.
    #>
    $namespace = Get-OutlookNamespace
    $liste     = [System.Collections.Generic.List[object]]::new()

    foreach ($store in $namespace.Stores) {
        try {
            Get-CalendriersRecursif -Dossier $store.GetRootFolder() -Liste $liste
        }
        catch {
            # Certains stores peuvent être inaccessibles (boîte déconnectée, etc.)
        }
    }

    return , $liste
}

function Find-Calendrier {
    <#
    .SYNOPSIS
        Retrouve un calendrier par son FolderPath complet.
    #>
    param([Parameter(Mandatory)] [string]$FolderPath)

    foreach ($cal in (Get-TousLesCalendriers)) {
        if ($cal.FolderPath -eq $FolderPath) { return $cal }
    }
    throw "Calendrier introuvable : $FolderPath"
}

function ConvertFrom-DateIso {
    <#
    .SYNOPSIS
        Convertit une date ISO locale (2026-10-20T09:00:00) en [datetime], sans tenir compte du fuseau.
    #>
    param([Parameter(Mandatory)] [object]$Valeur)

    if ($Valeur -is [datetime]) {
        return [datetime]::SpecifyKind($Valeur, [DateTimeKind]::Unspecified)
    }
    return [datetime]::ParseExact([string]$Valeur, 'yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-CleEvenement {
    param([string]$Sujet, [datetime]$Debut)
    return '{0}|{1:yyyyMMddHHmm}' -f $Sujet.Trim().ToLowerInvariant(), $Debut
}

function Get-ClesExistantes {
    <#
    .SYNOPSIS
        Renvoie les clés (titre + début) des rendez-vous déjà présents entre deux dates.

    .NOTES
        Restrict attend des dates au format de la langue d'Outlook : on utilise le format
        court du poste ('g'). En cas de refus, on parcourt le calendrier complet (plus lent).
    #>
    param(
        [Parameter(Mandatory)] [object]$Calendrier,
        [Parameter(Mandatory)] [datetime]$Min,
        [Parameter(Mandatory)] [datetime]$Max
    )

    $cles  = [System.Collections.Generic.HashSet[string]]::new()
    $debut = $Min.Date
    $fin   = $Max.Date.AddDays(1)

    try {
        $filtre   = "[Start] >= '{0}' AND [Start] < '{1}'" -f $debut.ToString('g'), $fin.ToString('g')
        $elements = $Calendrier.Items.Restrict($filtre)
    }
    catch {
        Write-Host '  Filtre Outlook refusé : parcours complet du calendrier.' -ForegroundColor DarkYellow
        $elements = $Calendrier.Items
    }

    foreach ($element in $elements) {
        try {
            $debutElement = [datetime]$element.Start
            if ($debutElement -ge $debut -and $debutElement -lt $fin) {
                [void]$cles.Add((Get-CleEvenement -Sujet $element.Subject -Debut $debutElement))
            }
        }
        catch {
            # Élément sans date exploitable (note, élément corrompu) : ignoré
        }
    }

    return , $cles
}

function Add-RendezVous {
    <#
    .SYNOPSIS
        Crée un rendez-vous dans le calendrier, avec ses participants obligatoires.
    #>
    param(
        [Parameter(Mandatory)] [object]$Calendrier,
        [Parameter(Mandatory)] [object]$Evt,
        [Parameter(Mandatory)] [datetime]$Debut,
        [Parameter(Mandatory)] [datetime]$Fin,
        [bool]$Envoyer
    )

    $avertissement = $null
    $rdv = $Calendrier.Items.Add($OL_APPOINTMENT_ITEM)

    $rdv.Subject     = [string]$Evt.sujet
    $rdv.Start       = $Debut
    $rdv.End         = $Fin
    $rdv.Location    = [string]$Evt.lieu
    $rdv.BusyStatus  = $OL_BUSY
    $rdv.ReminderSet = $false
    if ($Evt.categorie) { $rdv.Categories = [string]$Evt.categorie }

    $participants = @($Evt.participants)
    if ($participants.Count -gt 0) {
        # Réunion : chaque collaborateur est ajouté dans la zone « Obligatoire »
        $rdv.MeetingStatus = $OL_MEETING
        foreach ($p in $participants) {
            $destinataire      = $rdv.Recipients.Add([string]$p.email)
            $destinataire.Type = $OL_REQUIRED
        }
        if (-not $rdv.Recipients.ResolveAll()) {
            $nonResolus = foreach ($d in $rdv.Recipients) { if (-not $d.Resolved) { $d.Name } }
            $avertissement = "Adresse non reconnue par Outlook : $($nonResolus -join ', ')"
        }
        $rdv.ResponseRequested = $false
    }

    $rdv.Save()

    if ($Envoyer -and $participants.Count -gt 0) {
        $rdv.Send()
    }

    return $avertissement
}

function Invoke-Import {
    <#
    .SYNOPSIS
        Traite une demande d'import envoyée par la page (un lot d'événements).
    #>
    param([Parameter(Mandatory)] [object]$Demande)

    $calendrier = Find-Calendrier -FolderPath ([string]$Demande.calendrier)
    $evenements = @($Demande.evenements)
    $resultats  = [System.Collections.Generic.List[object]]::new()

    if ($evenements.Count -eq 0) { return @{ resultats = $resultats } }

    # Conversion des dates (format ISO local envoyé par la page : 2026-10-20T09:00:00).
    # PowerShell 7 les convertit déjà en [datetime] lors du ConvertFrom-Json, pas PowerShell 5.1.
    foreach ($evt in $evenements) {
        $evt | Add-Member -NotePropertyName dateDebut -NotePropertyValue (ConvertFrom-DateIso $evt.debut) -Force
        $evt | Add-Member -NotePropertyName dateFin   -NotePropertyValue (ConvertFrom-DateIso $evt.fin)   -Force
        $evt.debut = $evt.dateDebut.ToString('yyyy-MM-ddTHH:mm:ss')
    }

    $existants = $null
    if ($Demande.ignorerDoublons) {
        $dates     = @($evenements | ForEach-Object { $_.dateDebut } | Sort-Object)
        $existants = Get-ClesExistantes -Calendrier $calendrier -Min $dates[0] -Max $dates[-1]
    }

    foreach ($evt in $evenements) {
        $resultat = [ordered]@{
            sujet         = [string]$evt.sujet
            debut         = [string]$evt.debut
            statut        = ''
            message       = $null
            avertissement = $null
        }
        $cle = Get-CleEvenement -Sujet $evt.sujet -Debut $evt.dateDebut

        try {
            if ($null -ne $existants -and $existants.Contains($cle)) {
                $resultat.statut  = 'doublon'
                $resultat.message = 'déjà présent'
                Write-Host "  [DOUBLON] $($evt.sujet) ($($evt.debut))" -ForegroundColor DarkGray
            }
            elseif ($Demande.simulation) {
                $resultat.statut = 'simule'
                Write-Host "  [SIM]     $($evt.sujet) ($($evt.debut))" -ForegroundColor DarkCyan
            }
            else {
                $resultat.avertissement = Add-RendezVous -Calendrier $calendrier -Evt $evt `
                    -Debut $evt.dateDebut -Fin $evt.dateFin -Envoyer ([bool]$Demande.envoyer)
                $resultat.statut = 'cree'
                if ($null -ne $existants) { [void]$existants.Add($cle) }
                Write-Host "  [OK]      $($evt.sujet) ($($evt.debut))" -ForegroundColor Green
            }
        }
        catch {
            $resultat.statut  = 'erreur'
            $resultat.message = $_.Exception.Message
            Write-Host "  [ERREUR]  $($evt.sujet) : $($_.Exception.Message)" -ForegroundColor Red
        }

        $resultats.Add($resultat)
    }

    return @{ resultats = $resultats }
}

# ──────────────────────────────────────────────────────────────────────────────
# Fonctions serveur HTTP
# ──────────────────────────────────────────────────────────────────────────────

function Send-Reponse {
    param(
        [Parameter(Mandatory)] [System.Net.HttpListenerContext]$Contexte,
        [int]$Code = 200,
        [string]$Contenu = '',
        [string]$Type = 'application/json; charset=utf-8'
    )

    $octets = $Utf8SansBom.GetBytes($Contenu)
    $reponse = $Contexte.Response
    $reponse.StatusCode      = $Code
    $reponse.ContentType     = $Type
    $reponse.ContentLength64 = $octets.Length
    $reponse.Headers['Cache-Control']          = 'no-store'
    $reponse.Headers['X-Content-Type-Options'] = 'nosniff'
    $reponse.OutputStream.Write($octets, 0, $octets.Length)
    $reponse.OutputStream.Close()
}

function Send-Json {
    param(
        [Parameter(Mandatory)] [System.Net.HttpListenerContext]$Contexte,
        [Parameter(Mandatory)] [object]$Objet,
        [int]$Code = 200
    )
    Send-Reponse -Contexte $Contexte -Code $Code -Contenu (ConvertTo-Json -InputObject $Objet -Depth 6 -Compress)
}

function Read-Corps {
    param([Parameter(Mandatory)] [System.Net.HttpListenerRequest]$Requete)
    $lecteur = New-Object System.IO.StreamReader($Requete.InputStream, [Text.Encoding]::UTF8)
    try { return $lecteur.ReadToEnd() } finally { $lecteur.Close() }
}

function Invoke-Requete {
    <#
    .SYNOPSIS
        Aiguille une requête HTTP vers le bon traitement.
    #>
    param(
        [Parameter(Mandatory)] [System.Net.HttpListenerContext]$Contexte,
        [Parameter(Mandatory)] [string]$Jeton
    )

    $requete = $Contexte.Request
    $chemin  = $requete.Url.AbsolutePath
    $methode = $requete.HttpMethod

    # Protection contre le « DNS rebinding » : seul l'hôte localhost est accepté
    if ($requete.UserHostName -notin @("localhost:$Port", "127.0.0.1:$Port")) {
        Send-Json -Contexte $Contexte -Code 403 -Objet @{ erreur = 'Hôte non autorisé.' }
        return
    }

    # Page principale : le jeton de session est injecté dans le HTML
    if ($methode -eq 'GET' -and $chemin -eq '/') {
        $html = [IO.File]::ReadAllText($FichierHtml, [Text.Encoding]::UTF8).Replace('__JETON_SESSION__', $Jeton)
        Send-Reponse -Contexte $Contexte -Contenu $html -Type 'text/html; charset=utf-8'
        return
    }

    if (-not $chemin.StartsWith('/api/')) {
        Send-Json -Contexte $Contexte -Code 404 -Objet @{ erreur = 'Ressource inconnue.' }
        return
    }

    # Toutes les API exigent le jeton de session
    if ($requete.Headers['X-Jeton'] -ne $Jeton) {
        Send-Json -Contexte $Contexte -Code 401 -Objet @{ erreur = 'Jeton de session invalide. Rechargez la page depuis le script.' }
        return
    }

    try {
        switch ("$methode $chemin") {
            'GET /api/config' {
                $contenu = if (Test-Path -Path $FichierConfig -PathType Leaf) {
                    [IO.File]::ReadAllText($FichierConfig, [Text.Encoding]::UTF8)
                } else { '{}' }
                Send-Reponse -Contexte $Contexte -Contenu $contenu
            }
            'POST /api/config' {
                $corps = Read-Corps -Requete $requete
                $null  = $corps | ConvertFrom-Json   # validation : lève une erreur si JSON invalide
                [IO.File]::WriteAllText($FichierConfig, $corps, $Utf8SansBom)
                Write-Host "  Configuration enregistrée : $FichierConfig" -ForegroundColor Green
                Send-Json -Contexte $Contexte -Objet @{ ok = $true }
            }
            'GET /api/calendriers' {
                $liste = foreach ($cal in (Get-TousLesCalendriers)) {
                    @{ nom = [string]$cal.Name; folderPath = [string]$cal.FolderPath }
                }
                Send-Json -Contexte $Contexte -Objet @{ calendriers = @($liste) }
            }
            'POST /api/importer' {
                $demande = Read-Corps -Requete $requete | ConvertFrom-Json
                $mode    = if ($demande.simulation) { 'Simulation' } else { 'Import' }
                Write-Host "$mode de $(@($demande.evenements).Count) événement(s) dans $($demande.calendrier)" -ForegroundColor Cyan
                Send-Json -Contexte $Contexte -Objet (Invoke-Import -Demande $demande)
            }
            'POST /api/quitter' {
                Send-Json -Contexte $Contexte -Objet @{ ok = $true }
                $script:Arret = $true
            }
            default {
                Send-Json -Contexte $Contexte -Code 404 -Objet @{ erreur = 'Action inconnue.' }
            }
        }
    }
    catch {
        Write-Host "  [ERREUR] $chemin : $($_.Exception.Message)" -ForegroundColor Red
        try { Send-Json -Contexte $Contexte -Code 500 -Objet @{ erreur = $_.Exception.Message } } catch { }
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# Point d'entrée principal
# ──────────────────────────────────────────────────────────────────────────────

Write-Host ''
Write-Host '╔══════════════════════════════════════════════════════╗' -ForegroundColor DarkCyan
Write-Host '║   Planning CPDL  →  Calendrier Outlook partagé v2.0  ║' -ForegroundColor DarkCyan
Write-Host '║   DSI EFS – Direction des Systèmes d''Information     ║' -ForegroundColor DarkCyan
Write-Host '╚══════════════════════════════════════════════════════╝' -ForegroundColor DarkCyan
Write-Host ''

if (-not (Test-Path -Path $FichierHtml -PathType Leaf)) {
    Write-Error "Page introuvable : $FichierHtml (elle doit être dans le même dossier que le script)."
    exit 1
}

$jeton    = [guid]::NewGuid().ToString('N')
$adresse  = "http://localhost:$Port/"
$ecouteur = New-Object System.Net.HttpListener
$ecouteur.Prefixes.Add($adresse)

try {
    $ecouteur.Start()
}
catch {
    Write-Error "Impossible d'écouter sur $adresse ($($_.Exception.Message)). Le port est peut-être déjà utilisé : relancez avec -Port 9000 par exemple."
    exit 1
}

Write-Host "Application disponible sur : $adresse" -ForegroundColor Green
Write-Host "Configuration              : $FichierConfig" -ForegroundColor Gray
Write-Host 'Laissez cette fenêtre ouverte pendant l''utilisation. Ctrl+C pour arrêter.' -ForegroundColor Yellow
Write-Host ''

if (-not $SansNavigateur) {
    Start-Process $adresse
}

try {
    while ($ecouteur.IsListening -and -not $script:Arret) {
        $attente = $ecouteur.BeginGetContext($null, $null)
        # Attente par tranches de 500 ms pour que Ctrl+C reste pris en compte
        while (-not $attente.AsyncWaitHandle.WaitOne(500)) { }
        $contexte = $ecouteur.EndGetContext($attente)
        Invoke-Requete -Contexte $contexte -Jeton $jeton
    }
}
finally {
    $ecouteur.Stop()
    $ecouteur.Close()
    Write-Host ''
    Write-Host 'Serveur arrêté.' -ForegroundColor Gray
}
