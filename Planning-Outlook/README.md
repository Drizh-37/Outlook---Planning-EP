# Planning CPDL : du fichier Excel au calendrier Outlook partagé

Une seule application remplace le duo « page HTML + script d'import ICS ».
Elle lit le planning Excel, signale les incohérences, puis crée directement
les rendez-vous dans le calendrier partagé, collaborateur compris dans la zone
**Obligatoire**.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `Lancer-Planning-Outlook.cmd` | Double-clic pour démarrer l'application |
| `Planning-Outlook.ps1` | Moteur invisible : affiche la page dans sa fenêtre et pilote Outlook |
| `Planning-Outlook.html` | L'interface (lecture Excel, contrôles, configuration) |
| `config.json` | Créé au premier enregistrement : équipe, codes, horaires, catégories, calendrier |
| `Planning-Outlook.log` | Journal du dernier lancement (utile en cas de souci) |

Les trois premiers fichiers doivent rester dans le même dossier.

## Utilisation

1. Ouvrir Outlook.
2. Double-cliquer sur `Lancer-Planning-Outlook.cmd`.
   L'application s'ouvre dans sa propre fenêtre (Edge en mode application, sans onglets
   ni barre d'adresse). Aucune console PowerShell n'est affichée.
3. Étape 1 : déposer le fichier Excel (onglet dont le nom commence par « Planning »).
4. Étape 2 : vérifier l'aperçu et les incohérences.
5. Étape 3 : le calendrier `\\Cpdl.ServiceInformatique\Calendrier` est déjà renseigné. Cliquer sur **Vérifier**, puis **Simuler**, puis **Injecter dans Outlook**.

Options de l'étape 3 :

- **Ignorer les événements déjà présents** (coché par défaut) : un rendez-vous de même titre,
  même jour et même heure n'est pas recréé. On peut donc relancer un import sans doublon.
- **Envoyer l'invitation aux collaborateurs** (décoché par défaut) : sans cette case,
  le collaborateur figure en « Obligatoire » mais ne reçoit aucun mail.

Seul ce calendrier est ouvert : l'application ne parcourt pas les autres. Un autre chemin peut être saisi dans le champ ; il est mémorisé pour la fois suivante.

Pour quitter : fermer la fenêtre, ou cliquer sur **Fermer l'application**. Le moteur s'arrête
de lui-même quelques secondes après la fermeture (3 minutes au plus).
Relancer l'application alors qu'elle tourne déjà rouvre simplement une fenêtre.

## Configurer l'équipe

Dans l'encart « Codes traités », ouvrir **Modifier la configuration** :

- ajouter, modifier ou supprimer un collaborateur (nom tel qu'écrit dans Excel, site, email) ;
- renseigner les variantes d'écriture du nom, séparées par `;` (ex. `KOIBITCH Sophien; KOIBICH Sophie`) ;
- gérer les **codes reconnus** dans les cellules (voir ci-dessous) ;
- ajuster les horaires par défaut et le nom exact des catégories Outlook ;
- **Enregistrer** : le fichier `config.json` est écrit à côté du script.

Les boutons **Exporter / Importer (JSON)** permettent de partager la même configuration
entre plusieurs postes.

## Configurer les codes

Chaque code reconnu dans le planning se décrit sur une ligne :

| Champ | Rôle | Exemple |
|---|---|---|
| Code | Valeur saisie dans la cellule Excel (majuscules ou minuscules) | `EXPL` |
| Variantes | Autres écritures du même code, séparées par `;` | `PRCP` pour `PR` |
| Libellé | Début du titre du rendez-vous | `Exploitation` → « Exploitation — Sophien » |
| Mode | **Par site** : catégorie et lieu suivent le site du collaborateur (titre « Proxi NTS — Sophien »). **Fixe** : catégorie et lieu propres au code | `PR` par site, `EXPL` fixe |
| Catégorie, Lieu | Pour un code fixe uniquement | `Catégorie verte`, `EFS` |
| Début, Fin | Horaires propres au code ; vides = horaires par défaut | `08:00` à `18:00` |
| 1/jour | Signale une incohérence si deux collaborateurs ont ce code le même jour (par site pour un code « par site ») | coché pour `PR` et `EXPL` |

Une cellule peut contenir plusieurs codes séparés par `/` (ex. `PR/EXPL`) : un rendez-vous est créé pour chacun.
Les valeurs inconnues (PS93, BF93…) sont ignorées.

Attention : changer le libellé d'un code change le titre des rendez-vous. La détection des doublons
se fait sur le titre ; les rendez-vous déjà importés avec l'ancien libellé ne seront plus reconnus.

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
| Rien ne s'ouvre | Lancer `Lancer-Planning-Outlook.cmd /console` pour voir les messages, ou lire `Planning-Outlook.log`. Si l'exécution des scripts est bloquée par GPO, faire signer le script |
| « Port déjà utilisé » | `Lancer-Planning-Outlook.cmd -Port 9000` |
| « Calendrier introuvable » | Vérifier que la boîte Cpdl.ServiceInformatique apparaît dans Outlook (ou que vous avez les droits sur son calendrier) et que le chemin est exact |
| Dates sur la mauvaise année | Saisir l'année du premier jour dans « Année de la première colonne » (étape 1) ; le passage décembre → janvier est automatique |
| Alerte Outlook « Un programme tente d'accéder… » | Antivirus non reconnu par Outlook : accepter, ou voir avec l'équipe sécurité |
| Catégorie affichée en blanc | Le nom saisi ne correspond pas exactement à une catégorie existante dans Outlook |
| « Adresse non reconnue par Outlook » | Corriger l'email du collaborateur dans la configuration |
