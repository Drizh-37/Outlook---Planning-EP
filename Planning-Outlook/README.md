# Planning CPDL : du fichier Excel au calendrier Outlook partagé

Une seule application remplace le duo « page HTML + script d'import ICS ».
Elle lit le planning Excel, signale les incohérences, puis crée directement
les rendez-vous dans le calendrier partagé, collaborateur compris dans la zone
**Obligatoire**.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `Lancer-Planning-Outlook.cmd` | Double-clic pour démarrer l'application |
| `Planning-Outlook.ps1` | Petit serveur local qui affiche la page et pilote Outlook |
| `Planning-Outlook.html` | L'interface (lecture Excel, contrôles, configuration) |
| `config.json` | Créé au premier enregistrement : équipe, horaires, catégories, calendrier |

Les trois premiers fichiers doivent rester dans le même dossier.

## Utilisation

1. Ouvrir Outlook.
2. Double-cliquer sur `Lancer-Planning-Outlook.cmd`.
   Une fenêtre PowerShell s'ouvre (à laisser ouverte), puis le navigateur affiche l'application.
3. Étape 1 : déposer le fichier Excel (onglet dont le nom commence par « Planning »).
4. Étape 2 : vérifier l'aperçu et les incohérences.
5. Étape 3 : le calendrier `\\Cpdl.ServiceInformatique\Calendrier` est déjà renseigné. Cliquer sur **Vérifier**, puis **Simuler**, puis **Injecter dans Outlook**.

Options de l'étape 3 :

- **Ignorer les événements déjà présents** (coché par défaut) : un rendez-vous de même titre,
  même jour et même heure n'est pas recréé. On peut donc relancer un import sans doublon.
- **Envoyer l'invitation aux collaborateurs** (décoché par défaut) : sans cette case,
  le collaborateur figure en « Obligatoire » mais ne reçoit aucun mail.

Seul ce calendrier est ouvert : l'application ne parcourt pas les autres. Un autre chemin peut être saisi dans le champ ; il est mémorisé pour la fois suivante.

## Configurer l'équipe

Dans l'encart « Codes traités », ouvrir **Modifier la configuration** :

- ajouter, modifier ou supprimer un collaborateur (nom tel qu'écrit dans Excel, site, email) ;
- renseigner les variantes d'écriture du nom, séparées par `;` (ex. `KOIBITCH Sophien; KOIBICH Sophie`) ;
- ajuster les horaires et le nom exact des catégories Outlook ;
- **Enregistrer** : le fichier `config.json` est écrit à côté du script.

Les boutons **Exporter / Importer (JSON)** permettent de partager la même configuration
entre plusieurs postes.

## Mode autonome

Ouverte directement (double-clic sur le fichier HTML), la page fonctionne sans PowerShell :
export ICS uniquement, configuration gardée dans le navigateur. Le fichier ICS contient
désormais le participant obligatoire et la catégorie ; le script
`Import-ICS-CalendrierPartage.ps1` (v1.2, à la racine du dépôt) sait les reprendre.

## Sécurité

- Le serveur n'écoute que `http://localhost` : aucun autre poste ne peut s'y connecter.
- Chaque lancement génère un jeton aléatoire ; toute requête sans ce jeton est refusée.
- Les données du planning restent sur le poste. Seule la librairie de lecture Excel
  (SheetJS) est chargée depuis `cdnjs.cloudflare.com`.
- Aucun droit administrateur n'est nécessaire.

## En cas de souci

| Symptôme | Piste |
|---|---|
| La fenêtre se ferme aussitôt | Stratégie d'exécution bloquée par GPO : lancer `powershell -ExecutionPolicy Bypass -File Planning-Outlook.ps1` ou faire signer le script |
| « Port déjà utilisé » | `Lancer-Planning-Outlook.cmd -Port 9000` |
| « Calendrier introuvable » | Vérifier que la boîte Cpdl.ServiceInformatique apparaît dans Outlook (ou que vous avez les droits sur son calendrier) et que le chemin est exact |
| Dates sur la mauvaise année | Saisir l'année du premier jour dans « Année de la première colonne » (étape 1) ; le passage décembre → janvier est automatique |
| Alerte Outlook « Un programme tente d'accéder… » | Antivirus non reconnu par Outlook : accepter, ou voir avec l'équipe sécurité |
| Catégorie affichée en blanc | Le nom saisi ne correspond pas exactement à une catégorie existante dans Outlook |
| « Adresse non reconnue par Outlook » | Corriger l'email du collaborateur dans la configuration |
