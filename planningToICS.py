#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Planning EP → ICS Outlook
=========================
Convertit un planning Excel (colonnes PR / EXPL) en fichier .ics
avec invitations Outlook : destinataire Obligatoire, sans option de
réponse, description vide et catégories couleur Outlook.

Dépendances : pip install openpyxl
Python      : 3.8+
"""

from __future__ import annotations

import os
import re
import uuid
from datetime import date, datetime
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk
from tkinter import ttk

try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
    import openpyxl


# ─────────────────────────────────────────────────────────────────────────────
# Référentiel des collaborateurs
# ─────────────────────────────────────────────────────────────────────────────

COLLABORATEURS: dict[str, dict] = {
    # clé       : fragment du nom de famille présent dans la cellule Excel
    'BECAVIN':  {
        'nom_complet': 'Florent Becavin',
        'email':       'florent.becavin@efs.sante.fr',
        'site':        'NTS',
        'code':        'FBE',
    },
    'KOIBICH':  {
        'nom_complet': 'Sophien Koibich',
        'email':       'sophien.koibich@efs.sante.fr',
        'site':        'NTS',
        'code':        'SKO',
    },
    'ROUHIER':  {
        'nom_complet': 'Quentin Rouhier',
        'email':       'quentin.rouhier@efs.sante.fr',
        'site':        'NTS',
        'code':        'QRO',
    },
    'PERRIN':   {
        'nom_complet': 'Philippe Perrin',
        'email':       'philippe.perrin@efs.sante.fr',
        'site':        'TOU',
        'code':        'PPE',
    },
    'COUSIN':   {
        'nom_complet': 'Sylvain Cousin',
        'email':       'sylvain.cousin@efs.sante.fr',
        'site':        'TOU',
        'code':        'SCO',
    },
    'VENIN':    {
        'nom_complet': 'Hugo Venin',
        'email':       'hugo.venin@efs.sante.fr',
        'site':        'TOU',
        'code':        'HVE',
    },
    'CREZE':    {
        'nom_complet': 'Benoit Creze',
        'email':       'benoit.creze@efs.sante.fr',
        'site':        'ANG',
        'code':        'BCR',
    },
    'BODARD':   {
        'nom_complet': 'Emmanuel de Bodard',
        'email':       'emmanuel.DEBODARDDELAJACOPIERE@efs.sante.fr',
        'site':        'ANG',
        'code':        'EBO',
    },
}

# Intitulé des événements PR selon le site du collaborateur
SITE_LABELS: dict[str, str] = {
    'NTS': 'PROXI NTS',
    'TOU': 'PROXI TOU',
    'ANG': 'PROXI ANG',
}

# Catégories Outlook (noms par défaut Outlook en français)
CATEGORIES: dict[str, str] = {
    'EXPL': 'Catégorie verte',
    'ANG':  'Catégorie violette',
    'TOU':  'Catégorie bleue',
    'NTS':  'Catégorie orange',
}

HEURE_DEBUT = 9    # 09:00
HEURE_FIN   = 16   # 16:00
TIMEZONE    = 'Europe/Paris'


# ─────────────────────────────────────────────────────────────────────────────
# Utilitaires ICS (RFC 5545)
# ─────────────────────────────────────────────────────────────────────────────

def _escape(text: str) -> str:
    """Échappe les caractères spéciaux ICS."""
    text = text.replace('\\', '\\\\')
    text = text.replace(';',  '\\;')
    text = text.replace(',',  '\\,')
    text = text.replace('\n', '\\n')
    return text


def _fold(line: str) -> str:
    """
    Plie les lignes à 75 octets maximum (RFC 5545 §3.1).
    La continuation utilise CRLF + espace.
    """
    encoded = line.encode('utf-8')
    if len(encoded) <= 75:
        return line + '\r\n'

    result = ''
    while len(encoded) > 75:
        chunk = encoded[:75]
        # Reculer pour ne pas couper au milieu d'un caractère multi-octets
        while chunk and (chunk[-1] & 0xC0) == 0x80:
            chunk = chunk[:-1]
        result  += chunk.decode('utf-8') + '\r\n '
        encoded  = encoded[len(chunk):]
    result += encoded.decode('utf-8') + '\r\n'
    return result


def _fmt_dt(dt: datetime) -> str:
    return dt.strftime('%Y%m%dT%H%M%S')


# ─────────────────────────────────────────────────────────────────────────────
# Génération ICS
# ─────────────────────────────────────────────────────────────────────────────

_VTIMEZONE_PARIS = (
    'BEGIN:VTIMEZONE\r\n'
    'TZID:Europe/Paris\r\n'
    'BEGIN:STANDARD\r\n'
    'TZOFFSETFROM:+0200\r\n'
    'TZOFFSETTO:+0100\r\n'
    'TZNAME:CET\r\n'
    'DTSTART:19701025T030000\r\n'
    'RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10\r\n'
    'END:STANDARD\r\n'
    'BEGIN:DAYLIGHT\r\n'
    'TZOFFSETFROM:+0100\r\n'
    'TZOFFSETTO:+0200\r\n'
    'TZNAME:CEST\r\n'
    'DTSTART:19700329T020000\r\n'
    'RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3\r\n'
    'END:DAYLIGHT\r\n'
    'END:VTIMEZONE\r\n'
)


def _build_vevent(
    summary:         str,
    dtstart:         datetime,
    dtend:           datetime,
    organizer_name:  str,
    organizer_email: str,
    attendee_name:   str,
    attendee_email:  str,
    category:        str,
) -> str:
    """Retourne un bloc VEVENT complet (chaîne ICS pliée)."""
    dtstamp = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    uid     = str(uuid.uuid4())

    lines = [
        'BEGIN:VEVENT',
        f'UID:{uid}',
        f'DTSTAMP:{dtstamp}',
        f'DTSTART;TZID={TIMEZONE}:{_fmt_dt(dtstart)}',
        f'DTEND;TZID={TIMEZONE}:{_fmt_dt(dtend)}',
        f'SUMMARY:{_escape(summary)}',
        # Description vide (demande utilisateur)
        'DESCRIPTION:',
        f'CATEGORIES:{_escape(category)}',
        # Organisateur
        f'ORGANIZER;CN="{_escape(organizer_name)}":mailto:{organizer_email}',
        # Invité Obligatoire — RSVP=FALSE : pas d'options de réponse
        (
            f'ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;'
            f'PARTSTAT=NEEDS-ACTION;RSVP=FALSE;'
            f'CN="{_escape(attendee_name)}":mailto:{attendee_email}'
        ),
        # Propriétés Microsoft pour désactiver les options de réponse
        'X-MICROSOFT-CDO-BUSYSTATUS:BUSY',
        'X-MICROSOFT-DISALLOW-COUNTER:TRUE',
        'X-MICROSOFT-CDO-REPLYTIME:0',
        'STATUS:CONFIRMED',
        'END:VEVENT',
    ]
    return ''.join(_fold(l) for l in lines)


def generer_ics(
    evenements:      list[dict],
    organizer_name:  str,
    organizer_email: str,
) -> str:
    """Assemble le fichier ICS complet."""
    header = (
        'BEGIN:VCALENDAR\r\n'
        'VERSION:2.0\r\n'
        'PRODID:-//EFS DSI//Planning EP//FR\r\n'
        'CALSCALE:GREGORIAN\r\n'
        'METHOD:REQUEST\r\n'
        'X-WR-CALNAME:Planning EP\r\n'
        f'X-WR-TIMEZONE:{TIMEZONE}\r\n'
    )
    body = ''
    for evt in evenements:
        body += _build_vevent(
            summary         = evt['summary'],
            dtstart         = evt['dtstart'],
            dtend           = evt['dtend'],
            organizer_name  = organizer_name,
            organizer_email = organizer_email,
            attendee_name   = evt['attendee_name'],
            attendee_email  = evt['attendee_email'],
            category        = evt['category'],
        )
    return header + _VTIMEZONE_PARIS + body + 'END:VCALENDAR\r\n'


# ─────────────────────────────────────────────────────────────────────────────
# Parsing du fichier Excel
# ─────────────────────────────────────────────────────────────────────────────

_MOIS_FR: dict[str, int] = {
    'jan': 1, 'janv': 1,
    'fév': 2, 'fevr': 2, 'fev': 2,
    'mar': 3, 'mars': 3,
    'avr': 4, 'avril': 4,
    'mai': 5,
    'juin': 6, 'jui': 6,
    'juil': 7, 'juill': 7,
    'aoû': 8, 'aout': 8, 'août': 8,
    'sep': 9, 'sept': 9,
    'oct': 10,
    'nov': 11,
    'déc': 12, 'dec': 12,
}


def _parse_date_cell(valeur) -> date | None:
    """Convertit une valeur de cellule Excel en objet date."""
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    if not isinstance(valeur, str):
        return None

    s = valeur.strip()

    # "13-avr", "13-avr.", "13 avr", "13 avril"
    m = re.match(r'^(\d{1,2})[- ]([a-zéûàèù]+)\.?$', s, re.IGNORECASE)
    if m:
        jour     = int(m.group(1))
        mois_str = m.group(2).lower()
        for abbr, num in _MOIS_FR.items():
            if mois_str.startswith(abbr):
                annee = datetime.now().year
                try:
                    return date(annee, num, jour)
                except ValueError:
                    pass

    # Formats numériques : dd/mm/yyyy, yyyy-mm-dd, dd/mm/yy
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _find_header_row(ws, max_scan: int = 10) -> tuple[int, int]:
    """
    Détecte automatiquement la ligne d'en-tête et la première colonne de dates.
    Retourne (numéro_ligne, colonne_première_date) ; indexés en 1.
    """
    for row in range(1, min(max_scan + 1, ws.max_row + 1)):
        for col in range(1, ws.max_column + 1):
            if _parse_date_cell(ws.cell(row=row, column=col).value) is not None:
                return row, col
    return 1, 4  # fallback


def _find_name_column(ws, header_row: int, date_col_start: int) -> int:
    """
    Cherche la colonne 'Nom Personne' dans la ligne d'en-tête.
    Retourne l'index (1-based).
    """
    keywords = ['nom', 'personne', 'collaborateur', 'agent', 'prénom', 'prenom']
    for col in range(1, date_col_start):
        val = ws.cell(row=header_row, column=col).value
        if val and isinstance(val, str):
            if any(k in val.lower() for k in keywords):
                return col
    # Fallback : on cherche la dernière colonne non-date avant date_col_start
    return max(1, date_col_start - 1)


def _identify_collab(cell_value: str) -> dict | None:
    """Identifie un collaborateur à partir du contenu d'une cellule Nom."""
    upper = str(cell_value).upper()
    for key, data in COLLABORATEURS.items():
        if key in upper:
            return data
    return None


def parser_excel(chemin: str, ligne_header: int = 0) -> tuple[list[dict], list[str]]:
    """
    Lit le fichier Excel et retourne (liste_événements, liste_avertissements).

    Si ligne_header == 0, la ligne d'en-tête est détectée automatiquement.
    """
    wb = openpyxl.load_workbook(chemin, data_only=True)
    ws = wb.active

    avertissements: list[str] = []
    evenements:     list[dict] = []

    # Détection automatique ou utilisation de la valeur fournie
    if ligne_header == 0:
        header_row, col_dates_start = _find_header_row(ws)
    else:
        header_row      = ligne_header
        col_dates_start = 1
        # Chercher la première colonne de dates sur la ligne spécifiée
        for col in range(1, ws.max_column + 1):
            if _parse_date_cell(ws.cell(row=header_row, column=col).value) is not None:
                col_dates_start = col
                break

    col_nom = _find_name_column(ws, header_row, col_dates_start)

    # Collecte des dates de l'en-tête
    dates_par_col: dict[int, date] = {}
    for col in range(col_dates_start, ws.max_column + 1):
        d = _parse_date_cell(ws.cell(row=header_row, column=col).value)
        if d:
            dates_par_col[col] = d

    if not dates_par_col:
        avertissements.append(
            "Impossible de détecter les dates dans la ligne d'en-tête. "
            f"Vérifiez que la ligne {header_row} contient bien des dates."
        )
        return evenements, avertissements

    # Parcours des lignes de données
    for row in range(header_row + 1, ws.max_row + 1):
        nom_val = ws.cell(row=row, column=col_nom).value
        if not nom_val:
            continue

        collab = _identify_collab(str(nom_val))
        if not collab:
            continue

        for col, dt in dates_par_col.items():
            cell_val = ws.cell(row=row, column=col).value
            if not cell_val:
                continue

            token = str(cell_val).strip().upper()

            if token == 'PR':
                site     = collab['site']
                label    = SITE_LABELS.get(site, f'PROXI {site}')
                summary  = f"{label} ({collab['code']})"
                category = CATEGORIES.get(site, '')

            elif token == 'EXPL':
                summary  = f"Exploitation ({collab['code']})"
                category = CATEGORIES['EXPL']

            else:
                continue  # Ignorer FB90, BB90, PS93, etc.

            dtstart = datetime(dt.year, dt.month, dt.day, HEURE_DEBUT, 0, 0)
            dtend   = datetime(dt.year, dt.month, dt.day, HEURE_FIN,   0, 0)

            evenements.append({
                'summary':        summary,
                'dtstart':        dtstart,
                'dtend':          dtend,
                'attendee_name':  collab['nom_complet'],
                'attendee_email': collab['email'],
                'category':       category,
            })

    if not evenements:
        avertissements.append(
            "Aucun événement PR ou EXPL trouvé. "
            "Vérifiez que les cellules contiennent exactement 'PR' ou 'EXPL' "
            f"et que la ligne d'en-tête est bien la ligne {header_row}."
        )

    return evenements, avertissements


# ─────────────────────────────────────────────────────────────────────────────
# Interface graphique
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self) -> None:
        super().__init__()
        self.title('Planning EP → ICS Outlook')
        self.geometry('760x620')
        self.minsize(640, 520)
        self.configure(bg='#f0f0f0')
        self._build_ui()

    # ── Construction de l'interface ──────────────────────────────────────────

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TLabel',      background='#f0f0f0', font=('Segoe UI', 9))
        style.configure('TButton',     font=('Segoe UI', 9))
        style.configure('TLabelframe', background='#f0f0f0')
        style.configure('TLabelframe.Label', font=('Segoe UI', 9, 'bold'))
        style.configure('TFrame',      background='#f0f0f0')
        style.configure('Title.TLabel',
                        font=('Segoe UI', 13, 'bold'), foreground='#003366',
                        background='#f0f0f0')
        style.configure('Sub.TLabel',
                        foreground='#555555', background='#f0f0f0',
                        font=('Segoe UI', 9))
        style.configure('Generate.TButton',
                        font=('Segoe UI', 10, 'bold'))

        # ── Bandeau titre ────────────────────────────────────────────────────
        header = ttk.Frame(self, padding=(16, 10, 16, 6))
        header.pack(fill='x')
        ttk.Label(header, text='Planning EP  →  ICS Outlook', style='Title.TLabel').pack(anchor='w')
        ttk.Label(header,
                  text='Lit un planning Excel (colonnes PR / EXPL) et génère un fichier .ics '
                       'avec invitations Outlook.',
                  style='Sub.TLabel').pack(anchor='w')
        ttk.Separator(self).pack(fill='x', padx=0)

        # ── Corps principal ──────────────────────────────────────────────────
        body = ttk.Frame(self, padding=(16, 10))
        body.pack(fill='both', expand=True)

        # Fichier Excel
        lf = ttk.LabelFrame(body, text='Fichier Excel source', padding=(8, 6))
        lf.pack(fill='x', pady=(0, 8))

        self.var_excel = tk.StringVar()
        ttk.Entry(lf, textvariable=self.var_excel).pack(
            side='left', expand=True, fill='x', padx=(0, 6))
        ttk.Button(lf, text='Parcourir…', command=self._pick_excel).pack(side='left')

        # Paramètres Excel
        lf2 = ttk.LabelFrame(body, text='Paramètres Excel', padding=(8, 6))
        lf2.pack(fill='x', pady=(0, 8))

        ttk.Label(lf2, text="Numéro de ligne d'en-tête (0 = auto) :").grid(
            row=0, column=0, sticky='w')
        self.var_header_row = tk.IntVar(value=0)
        ttk.Spinbox(lf2, from_=0, to=30, textvariable=self.var_header_row,
                    width=5).grid(row=0, column=1, sticky='w', padx=(8, 0))
        ttk.Label(lf2, text='  (0 = détection automatique)',
                  style='Sub.TLabel').grid(row=0, column=2, sticky='w', padx=(4, 0))

        # Organisateur
        lf3 = ttk.LabelFrame(body, text='Organisateur  (expéditeur des invitations)', padding=(8, 6))
        lf3.pack(fill='x', pady=(0, 8))
        lf3.columnconfigure(1, weight=1)
        lf3.columnconfigure(3, weight=1)

        ttk.Label(lf3, text='Nom :').grid(row=0, column=0, sticky='w')
        self.var_org_name = tk.StringVar(value='Service Informatique EFS')
        ttk.Entry(lf3, textvariable=self.var_org_name).grid(
            row=0, column=1, sticky='ew', padx=(6, 20))

        ttk.Label(lf3, text='E-mail :').grid(row=0, column=2, sticky='w')
        self.var_org_email = tk.StringVar()
        ttk.Entry(lf3, textvariable=self.var_org_email).grid(
            row=0, column=3, sticky='ew', padx=(6, 0))

        ttk.Label(lf3,
                  text="L'e-mail doit correspondre à votre compte Outlook pour que les invitations soient envoyées.",
                  style='Sub.TLabel').grid(row=1, column=0, columnspan=4, sticky='w', pady=(4, 0))

        # Fichier de sortie
        lf4 = ttk.LabelFrame(body, text='Fichier ICS de sortie', padding=(8, 6))
        lf4.pack(fill='x', pady=(0, 8))

        self.var_output = tk.StringVar()
        ttk.Entry(lf4, textvariable=self.var_output).pack(
            side='left', expand=True, fill='x', padx=(0, 6))
        ttk.Button(lf4, text='Parcourir…', command=self._pick_output).pack(side='left')

        # Bouton Générer
        btn_row = ttk.Frame(body)
        btn_row.pack(fill='x', pady=(0, 6))
        ttk.Button(btn_row, text='Générer le fichier ICS',
                   style='Generate.TButton',
                   command=self._generate).pack(side='right')

        # Journal
        lf5 = ttk.LabelFrame(body, text='Journal', padding=(8, 6))
        lf5.pack(fill='both', expand=True)

        self.log = scrolledtext.ScrolledText(
            lf5, height=8, state='disabled',
            font=('Consolas', 8),
            bg='#1e1e1e', fg='#d4d4d4', insertbackground='white',
        )
        self.log.pack(fill='both', expand=True)
        self.log.tag_configure('ok',   foreground='#6dbf67')
        self.log.tag_configure('warn', foreground='#f0c040')
        self.log.tag_configure('err',  foreground='#f44747')
        self.log.tag_configure('info', foreground='#9cdcfe')
        self.log.tag_configure('dim',  foreground='#808080')

    # ── Helpers UI ───────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = '') -> None:
        self.log.configure(state='normal')
        self.log.insert('end', msg + '\n', tag)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _pick_excel(self) -> None:
        path = filedialog.askopenfilename(
            title='Sélectionner le fichier Excel',
            filetypes=[('Fichiers Excel', '*.xlsx *.xlsm *.xls'), ('Tous', '*.*')],
        )
        if path:
            self.var_excel.set(path)
            if not self.var_output.get():
                self.var_output.set(os.path.splitext(path)[0] + '.ics')

    def _pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title='Enregistrer le fichier ICS',
            defaultextension='.ics',
            filetypes=[('Fichiers ICS', '*.ics'), ('Tous', '*.*')],
        )
        if path:
            self.var_output.set(path)

    # ── Action principale ────────────────────────────────────────────────────

    def _generate(self) -> None:
        excel_path  = self.var_excel.get().strip()
        output_path = self.var_output.get().strip()
        org_name    = self.var_org_name.get().strip()
        org_email   = self.var_org_email.get().strip()
        header_row  = self.var_header_row.get()

        # Validation
        if not excel_path:
            messagebox.showerror('Champ manquant', 'Veuillez sélectionner un fichier Excel.')
            return
        if not os.path.isfile(excel_path):
            messagebox.showerror('Fichier introuvable', f'Fichier Excel introuvable :\n{excel_path}')
            return
        if not output_path:
            messagebox.showerror('Champ manquant', 'Veuillez indiquer un fichier ICS de sortie.')
            return
        if not org_email:
            messagebox.showerror('Champ manquant',
                                 "Veuillez renseigner l'e-mail de l'organisateur.\n"
                                 'Il doit correspondre à votre compte Outlook.')
            return

        self._log('─' * 64, 'dim')
        self._log(f'Lecture : {excel_path}', 'info')

        # Parsing Excel
        try:
            evenements, avertissements = parser_excel(excel_path, header_row)
        except Exception as exc:
            self._log(f'[ERREUR] Lecture Excel : {exc}', 'err')
            messagebox.showerror('Erreur', f'Impossible de lire le fichier Excel :\n{exc}')
            return

        for msg in avertissements:
            self._log(f'[AVERT] {msg}', 'warn')

        if not evenements:
            messagebox.showwarning('Aucun événement',
                                   'Aucun événement PR ou EXPL trouvé dans le fichier Excel.\n\n'
                                   'Vérifiez :\n'
                                   '• que les cellules contiennent exactement "PR" ou "EXPL"\n'
                                   '• le numéro de ligne d\'en-tête (ou laissez 0 pour auto)')
            return

        self._log(f'{len(evenements)} événement(s) trouvé(s) :', 'ok')

        # Résumé par type
        compteur: dict[str, int] = {}
        for e in evenements:
            key = e['summary'].split(' (')[0]
            compteur[key] = compteur.get(key, 0) + 1
        for k in sorted(compteur):
            self._log(f'   {compteur[k]:3d}  ×  {k}', 'info')

        # Génération ICS
        try:
            ics_content = generer_ics(evenements, org_name, org_email)
        except Exception as exc:
            self._log(f'[ERREUR] Génération ICS : {exc}', 'err')
            messagebox.showerror('Erreur', f'Erreur lors de la génération ICS :\n{exc}')
            return

        # Écriture
        try:
            with open(output_path, 'w', encoding='utf-8', newline='') as fh:
                fh.write(ics_content)
        except Exception as exc:
            self._log(f'[ERREUR] Écriture fichier : {exc}', 'err')
            messagebox.showerror('Erreur', f"Impossible d'écrire le fichier ICS :\n{exc}")
            return

        self._log(f'[OK] Fichier ICS enregistré : {output_path}', 'ok')
        messagebox.showinfo(
            'Succès',
            f'Fichier ICS généré avec succès !\n\n'
            f'{output_path}\n\n'
            f'{len(evenements)} événement(s) créé(s).\n\n'
            f'Ouvrez le fichier .ics avec Outlook pour importer\n'
            f'les événements et envoyer les invitations.',
        )


# ─────────────────────────────────────────────────────────────────────────────
# Point d'entrée
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = App()
    app.mainloop()
