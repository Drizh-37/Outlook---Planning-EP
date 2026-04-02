<#
.SYNOPSIS
    Import un fichier ICS dans un calendrier Outlook partagé.

.DESCRIPTION
    Ce script parse un fichier .ICS et importe les événements VEVENT dans
    le calendrier Outlook désigné via l'API COM. Il prend en charge le
    ciblage par FolderPath complet ou par nom simple, le listing des
    calendriers disponibles et un mode simulation (dry run).

.PARAMETER CheminICS
    Chemin complet vers le fichier .ics à importer.

.PARAMETER CheminCalendrier
    FolderPath complet du calendrier cible (ex. \\dupont@efs.sante.fr\Calendrier).
    Prioritaire sur -NomCalendrier.

.PARAMETER NomCalendrier
    Nom simple du calendrier cible.
    Déconseillé si plusieurs calendriers partagent le même nom.

.PARAMETER ListerCalendriers
    Affiche tous les calendriers disponibles avec leur FolderPath, puis quitte.

.PARAMETER Simulation
    Parcourt le fichier ICS et affiche les événements sans les créer dans Outlook.

.EXAMPLE
    .\Import-ICS-CalendrierPartage.ps1 -CheminICS 'C:\temp\dummy.ics' -ListerCalendriers

.EXAMPLE
    .\Import-ICS-CalendrierPartage.ps1 -CheminICS 'C:\temp\planning.ics' `
        -CheminCalendrier '\\equipe-tours@efs.sante.fr\Calendrier' -Simulation

.EXAMPLE
    .\Import-ICS-CalendrierPartage.ps1 -CheminICS 'C:\temp\planning.ics' `
        -CheminCalendrier '\\equipe-tours@efs.sante.fr\Calendrier'

.NOTES
    Version   : 1.1
    Auteur    : DSI EFS – Tours
    Date      : Avril 2026

    Prérequis :
      - Outlook 2021 ou Microsoft 365 installé et configuré sur le poste.
      - Droits d'édition sur le calendrier partagé cible.
      - Politique d'exécution PowerShell : RemoteSigned minimum.
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

    Limites connues :
      - Les événements récurrents (RRULE) ne sont pas interprétés (F-10).
      - Les fuseaux horaires (TZID) sont traités de façon simplifiée (F-11 partiel).
      - Aucune déduplication par UID ICS (F-09).
      - Aucun rollback possible après import (F-08).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'Chemin complet vers le fichier .ics à importer.')]
    [string]$CheminICS,

    [Parameter(Mandatory = $false, HelpMessage = 'FolderPath complet du calendrier cible (prioritaire sur -NomCalendrier).')]
    [string]$CheminCalendrier,

    [Parameter(Mandatory = $false, HelpMessage = 'Nom simple du calendrier cible (fallback).')]
    [string]$NomCalendrier,

    [Parameter(Mandatory = $false, HelpMessage = 'Affiche tous les FolderPath disponibles, puis quitte.')]
    [switch]$ListerCalendriers,

    [Parameter(Mandatory = $false, HelpMessage = 'Affiche les événements sans les créer dans Outlook.')]
    [switch]$Simulation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ──────────────────────────────────────────────────────────────────────────────
# Constantes Outlook COM
# ──────────────────────────────────────────────────────────────────────────────
$OL_FOLDER_CALENDAR  = 9   # olFolderCalendar
$OL_APPOINTMENT_ITEM = 1   # olAppointmentItem

# ──────────────────────────────────────────────────────────────────────────────
# Fonctions utilitaires
# ──────────────────────────────────────────────────────────────────────────────

function Get-OutlookInstance {
    <#
    .SYNOPSIS
        Renvoie une instance COM Outlook (existante ou nouvelle).
    #>
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    }
    catch {
        try {
            return New-Object -ComObject Outlook.Application
        }
        catch {
            Write-Error "Impossible de démarrer Outlook. Vérifiez qu'il est installé et correctement configuré."
            exit 1
        }
    }
}

function Get-CalendriersRecursif {
    <#
    .SYNOPSIS
        Parcourt récursivement un dossier Outlook et collecte les sous-dossiers de type Calendrier.
    #>
    param(
        [Parameter(Mandatory)] [object]$Dossier,
        [Parameter(Mandatory)] [System.Collections.Generic.List[PSCustomObject]]$Liste
    )

    if ($Dossier.DefaultItemType -eq $OL_FOLDER_CALENDAR) {
        $Liste.Add([PSCustomObject]@{
            Nom        = $Dossier.Name
            FolderPath = $Dossier.FolderPath
            Dossier    = $Dossier
        })
    }

    foreach ($sousDossier in $Dossier.Folders) {
        Get-CalendriersRecursif -Dossier $sousDossier -Liste $Liste
    }
}

function Get-TousLesCalendriers {
    <#
    .SYNOPSIS
        Renvoie la liste de tous les calendriers accessibles depuis Outlook.
    #>
    param([Parameter(Mandatory)] [object]$Namespace)

    $liste = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($store in $Namespace.Stores) {
        try {
            Get-CalendriersRecursif -Dossier $store.GetRootFolder() -Liste $liste
        }
        catch {
            # Certains stores peuvent être inaccessibles (boîte déconnectée, etc.)
        }
    }

    return $liste
}

function ConvertFrom-ICSDate {
    <#
    .SYNOPSIS
        Convertit une chaîne de date au format ICS (RFC 5545) en objet [datetime].

    .NOTES
        Traitement simplifié (F-11 partiel) :
        - Le suffixe Z (UTC) est supprimé ; la date est interprétée comme heure locale.
        - Le paramètre TZID est ignoré : risque de décalage lors des changements d'heure.
    #>
    param([string]$DateICS)

    if ([string]::IsNullOrWhiteSpace($DateICS)) { return $null }

    # Supprimer le suffixe UTC 'Z' (traitement simplifié)
    $DateICS = $DateICS -replace 'Z$', ''

    # Format date + heure : 20240415T083000
    if ($DateICS -match '^\d{8}T\d{6}$') {
        return [datetime]::ParseExact($DateICS, 'yyyyMMddTHHmmss', $null)
    }

    # Format date seule (événement sur la journée) : 20240415
    if ($DateICS -match '^\d{8}$') {
        return [datetime]::ParseExact($DateICS, 'yyyyMMdd', $null)
    }

    Write-Warning "Format de date ICS non reconnu : '$DateICS'"
    return $null
}

function ConvertFrom-ICSText {
    <#
    .SYNOPSIS
        Décode les séquences d'échappement ICS dans une valeur texte.
    #>
    param([string]$Texte)

    if ([string]::IsNullOrEmpty($Texte)) { return $Texte }

    return $Texte `
        -replace '\\n',  "`n" `
        -replace '\\N',  "`n" `
        -replace '\\,',  ',' `
        -replace '\\;',  ';' `
        -replace '\\\\', '\'
}

function Read-ICSFile {
    <#
    .SYNOPSIS
        Lit un fichier ICS et retourne la liste des événements VEVENT sous forme de hashtables.

    .NOTES
        Gère le line folding RFC 5545 (continuation de ligne avec espace ou tabulation).
        Les propriétés avec paramètres (ex. DTSTART;TZID=Europe/Paris) sont supportées.
    #>
    param([Parameter(Mandatory)] [string]$Chemin)

    $contenu = Get-Content -Path $Chemin -Encoding UTF8 -Raw

    # Dépliage des lignes (RFC 5545 §3.1 : continuation = CRLF suivi d'un espace ou d'une tabulation)
    $contenu = $contenu -replace "\r\n[ `t]", ''
    $contenu = $contenu -replace "\n[ `t]",   ''

    $evenements = [System.Collections.Generic.List[hashtable]]::new()
    $lignes     = $contenu -split "\r?\n"

    $dansVEvent      = $false
    $evenementCourant = $null

    foreach ($ligne in $lignes) {
        switch ($ligne) {
            'BEGIN:VEVENT' {
                $dansVEvent       = $true
                $evenementCourant = [hashtable]::new()
                break
            }
            'END:VEVENT' {
                if ($dansVEvent -and $null -ne $evenementCourant) {
                    $evenements.Add($evenementCourant)
                }
                $dansVEvent       = $false
                $evenementCourant = $null
                break
            }
            default {
                if (-not $dansVEvent) { break }

                # Propriété avec paramètres : NOM;PARAM=val:valeur
                # Propriété simple         : NOM:valeur
                if ($ligne -match '^([A-Z\-]+)(?:;([^:]*))?:(.*)$') {
                    $nomPropriete = $Matches[1].ToUpper()
                    $valeur       = $Matches[3]

                    # Stocker la valeur (la dernière occurrence l'emporte pour les doublons simples)
                    $evenementCourant[$nomPropriete] = $valeur
                }
                break
            }
        }
    }

    return $evenements
}

function Show-EvenementSimulation {
    <#
    .SYNOPSIS
        Affiche les détails d'un événement en mode simulation.
    #>
    param(
        [Parameter(Mandatory)] [hashtable]$Evt,
        [Parameter(Mandatory)] [int]$Index
    )

    $sujet  = if ($Evt['SUMMARY'])  { ConvertFrom-ICSText $Evt['SUMMARY'] } else { '(sans titre)' }
    $debut  = ConvertFrom-ICSDate ($Evt['DTSTART'])
    $fin    = ConvertFrom-ICSDate ($Evt['DTEND'])
    $lieu   = if ($Evt['LOCATION'])    { ConvertFrom-ICSText $Evt['LOCATION'] }    else { '' }
    $descr  = if ($Evt['DESCRIPTION']) { ConvertFrom-ICSText $Evt['DESCRIPTION'] } else { '' }

    Write-Host ''
    Write-Host "  [$Index] $sujet" -ForegroundColor Cyan

    $debutStr = if ($debut) { $debut.ToString('dd/MM/yyyy HH:mm') } else { 'N/A' }
    $finStr   = if ($fin)   { $fin.ToString('dd/MM/yyyy HH:mm') }   else { 'N/A' }

    Write-Host "       Début       : $debutStr"
    Write-Host "       Fin         : $finStr"

    if ($lieu)  { Write-Host "       Lieu        : $lieu" }
    if ($descr) {
        $descrCourte = if ($descr.Length -gt 100) { $descr.Substring(0, 100) + '…' } else { $descr }
        Write-Host "       Description : $descrCourte"
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# Point d'entrée principal
# ──────────────────────────────────────────────────────────────────────────────

Write-Host ''
Write-Host '╔══════════════════════════════════════════════════════╗' -ForegroundColor DarkCyan
Write-Host '║   Import ICS  →  Calendrier Outlook partagé  v1.1   ║' -ForegroundColor DarkCyan
Write-Host '║   DSI EFS – Direction des Systèmes d''Information    ║' -ForegroundColor DarkCyan
Write-Host '╚══════════════════════════════════════════════════════╝' -ForegroundColor DarkCyan
Write-Host ''

# ── Validation du fichier ICS ────────────────────────────────────────────────
if (-not $ListerCalendriers) {
    if (-not (Test-Path -Path $CheminICS -PathType Leaf)) {
        Write-Error "Fichier ICS introuvable : $CheminICS"
        exit 1
    }

    if ([IO.Path]::GetExtension($CheminICS).ToLower() -ne '.ics') {
        Write-Warning "Le fichier spécifié n'a pas l'extension .ics : $CheminICS"
    }
}

# ── Connexion à Outlook ──────────────────────────────────────────────────────
Write-Host 'Connexion à Outlook...' -ForegroundColor Gray
$outlook   = Get-OutlookInstance
$namespace = $outlook.GetNamespace('MAPI')
Write-Host '  Connecté.' -ForegroundColor Green
Write-Host ''

# ── F-04 : Listing des calendriers ──────────────────────────────────────────
if ($ListerCalendriers) {
    Write-Host 'Calendriers Outlook disponibles :' -ForegroundColor Yellow
    Write-Host ('─' * 68) -ForegroundColor DarkGray

    $tousCalendriers = Get-TousLesCalendriers -Namespace $namespace

    if ($tousCalendriers.Count -eq 0) {
        Write-Host '  Aucun calendrier trouvé.' -ForegroundColor Red
    }
    else {
        foreach ($cal in $tousCalendriers) {
            Write-Host '  Nom        : ' -NoNewline -ForegroundColor Gray
            Write-Host $cal.Nom           -ForegroundColor White
            Write-Host '  FolderPath : ' -NoNewline -ForegroundColor Gray
            Write-Host $cal.FolderPath    -ForegroundColor Green
            Write-Host ('  ' + '─' * 64) -ForegroundColor DarkGray
        }
        Write-Host ''
        Write-Host "  Total : $($tousCalendriers.Count) calendrier(s) trouvé(s)." -ForegroundColor Yellow
    }

    Write-Host ''
    exit 0
}

# ── Résolution du calendrier cible ───────────────────────────────────────────
$calendrierCible = $null

if ($CheminCalendrier) {
    # F-02 : Ciblage par FolderPath (mode prioritaire)
    Write-Host "Recherche du calendrier par FolderPath : $CheminCalendrier" -ForegroundColor Gray

    $tousCalendriers = Get-TousLesCalendriers -Namespace $namespace
    $correspondances = @($tousCalendriers | Where-Object { $_.FolderPath -eq $CheminCalendrier })

    if ($correspondances.Count -eq 0) {
        Write-Error "Aucun calendrier trouvé pour le chemin : $CheminCalendrier"
        Write-Host '  Conseil : utilisez -ListerCalendriers pour afficher les FolderPath disponibles.' -ForegroundColor Yellow
        exit 1
    }

    $calendrierCible = $correspondances[0].Dossier
    Write-Host "  Calendrier cible : $($calendrierCible.Name)" -ForegroundColor Green
    Write-Host "  FolderPath       : $($calendrierCible.FolderPath)" -ForegroundColor Green
}
elseif ($NomCalendrier) {
    # F-03 : Ciblage par nom simple (fallback)
    Write-Host "Recherche du calendrier par nom : $NomCalendrier" -ForegroundColor Gray

    $tousCalendriers = Get-TousLesCalendriers -Namespace $namespace
    $correspondances = @($tousCalendriers | Where-Object { $_.Nom -eq $NomCalendrier })

    if ($correspondances.Count -eq 0) {
        Write-Error "Aucun calendrier trouvé avec le nom : $NomCalendrier"
        Write-Host '  Conseil : utilisez -ListerCalendriers pour afficher les noms disponibles.' -ForegroundColor Yellow
        exit 1
    }

    # F-07 : Avertissement en cas de doublons de nom de calendrier
    if ($correspondances.Count -gt 1) {
        Write-Host ''
        Write-Host "  ⚠  ATTENTION : $($correspondances.Count) calendriers portent le nom '$NomCalendrier' :" -ForegroundColor Yellow
        foreach ($cal in $correspondances) {
            Write-Host "       - $($cal.FolderPath)" -ForegroundColor Yellow
        }
        Write-Host "  → Le premier résultat sera utilisé. Préférez -CheminCalendrier pour garantir l'unicité." -ForegroundColor Yellow
        Write-Host ''
    }

    $calendrierCible = $correspondances[0].Dossier
    Write-Host "  Calendrier cible : $($calendrierCible.Name)" -ForegroundColor Green
    Write-Host "  FolderPath       : $($calendrierCible.FolderPath)" -ForegroundColor Green
}
else {
    Write-Error "Aucun calendrier cible spécifié. Utilisez -CheminCalendrier (recommandé) ou -NomCalendrier."
    Write-Host '  Conseil : utilisez -ListerCalendriers pour afficher les calendriers disponibles.' -ForegroundColor Yellow
    exit 1
}

Write-Host ''

# ── Parsing du fichier ICS ───────────────────────────────────────────────────
Write-Host "Lecture du fichier ICS : $CheminICS" -ForegroundColor Gray

try {
    $evenements = Read-ICSFile -Chemin $CheminICS
}
catch {
    Write-Error "Erreur lors de la lecture du fichier ICS : $_"
    exit 1
}

if ($evenements.Count -eq 0) {
    Write-Warning "Aucun événement VEVENT trouvé dans le fichier ICS."
    exit 0
}

Write-Host "  $($evenements.Count) événement(s) VEVENT trouvé(s)." -ForegroundColor Cyan
Write-Host ''

# ── F-05 : Mode simulation ───────────────────────────────────────────────────
if ($Simulation) {
    Write-Host ('═' * 56) -ForegroundColor Yellow
    Write-Host '  MODE SIMULATION – Aucun événement ne sera créé dans Outlook' -ForegroundColor Yellow
    Write-Host ('═' * 56) -ForegroundColor Yellow

    $index = 1
    foreach ($evt in $evenements) {
        Show-EvenementSimulation -Evt $evt -Index $index
        $index++
    }

    Write-Host ''
    Write-Host ('═' * 56) -ForegroundColor Yellow
    Write-Host "  $($evenements.Count) événement(s) seraient importés dans :" -ForegroundColor Yellow
    Write-Host "  $($calendrierCible.FolderPath)" -ForegroundColor Green
    Write-Host ('═' * 56) -ForegroundColor Yellow
    Write-Host ''
    exit 0
}

# ── F-01 : Import réel ───────────────────────────────────────────────────────
Write-Host 'Démarrage de l''import...' -ForegroundColor Cyan
Write-Host ('─' * 56) -ForegroundColor DarkGray

$nbImportes = 0
$nbErreurs  = 0

foreach ($evt in $evenements) {
    $sujet = if ($evt['SUMMARY']) { ConvertFrom-ICSText $evt['SUMMARY'] } else { '(sans titre)' }

    try {
        # Création du rendez-vous via l'API COM Outlook
        $rdv = $calendrierCible.Items.Add($OL_APPOINTMENT_ITEM)

        # Sujet (F-01)
        $rdv.Subject = $sujet

        # Dates (F-01)
        $debut = ConvertFrom-ICSDate ($evt['DTSTART'])
        $fin   = ConvertFrom-ICSDate ($evt['DTEND'])

        if ($null -ne $debut) { $rdv.Start = $debut }
        if ($null -ne $fin)   { $rdv.End   = $fin   }

        # Lieu (F-01)
        if ($evt['LOCATION']) {
            $rdv.Location = ConvertFrom-ICSText $evt['LOCATION']
        }

        # Description (F-01)
        if ($evt['DESCRIPTION']) {
            $rdv.Body = ConvertFrom-ICSText $evt['DESCRIPTION']
        }

        $rdv.Save()

        Write-Host "  [OK]     $sujet" -ForegroundColor Green
        $nbImportes++
    }
    catch {
        Write-Host "  [ERREUR] $sujet" -ForegroundColor Red
        Write-Host "           $_" -ForegroundColor DarkRed
        $nbErreurs++
    }
}

# ── F-06 : Résumé de fin d'exécution ────────────────────────────────────────
Write-Host ''
Write-Host ('═' * 56) -ForegroundColor Cyan
Write-Host '  RÉSUMÉ D''EXÉCUTION' -ForegroundColor Cyan
Write-Host ('═' * 56) -ForegroundColor Cyan
Write-Host "  Calendrier cible                 : $($calendrierCible.FolderPath)" -ForegroundColor White
Write-Host "  Événements importés avec succès  : $nbImportes" -ForegroundColor Green

if ($nbErreurs -gt 0) {
    Write-Host "  Événements en erreur             : $nbErreurs" -ForegroundColor Red
}
else {
    Write-Host "  Événements en erreur             : $nbErreurs" -ForegroundColor Gray
}

Write-Host ('═' * 56) -ForegroundColor Cyan
Write-Host ''
