# Spécifications Fonctionnelles, Techniques et Guide des Écrans — SyncContactsOG

- **Version courante** : `1.0.11`
- **Date de mise à jour** : `2026-11-22`
- **Plateforme cible** : Windows 10 / Windows 11 (x64)
- **Technologies** : Python 3.12, CustomTkinter, PyWin32 (MAPI/COM), Microsoft Graph API, Google People API, SQLite, OpenPyXL, PyInstaller

---

## 1. Objectif global & Périmètre

### 1.1 Finalité
**SyncContactsOG** est une application Windows portable destinée à synchroniser, unifier et archiver de manière fiable les carnets d'adresses entre **Microsoft Outlook** (client lourd classique ou compte cloud Microsoft 365) et **Google Contacts**.

### 1.2 Périmètre
- **Inclus (*In-Scope*)** :
  - Synchronisation unidirectionnelle et bidirectionnelle des contacts personnels.
  - Conservation de la fidélité des données (noms composés, `FileAs`, téléphones multiples, adresses postales, notes riches RTF, catégories/libellés et photos de profil haute définition).
  - Archivage local consultable (base SQLite avec photos binaires + classeur Excel `.xlsx`).
  - Exécution interactive (interface graphique) et en arrière-plan (zone de notification Windows / démarrage automatique).
- **Exclus (*Out-of-Scope*)** :
  - Synchronisation des calendriers, tâches ou courriels.
  - Synchronisation multi-comptes simultanés au sein d'une même instance.

---

## 2. Acteurs & Modes d'utilisation

| Mode d'exécution | Déclenchement | Comportement attendu |
| :--- | :--- | :--- |
| **Interactif (GUI)** | Lancement manuel de `SyncContactsOG.exe` | Affiche la fenêtre principale, permet le paramétrage, la simulation (*Dry-Run*) et le lancement manuel. |
| **Arrière-plan (Tray)** | Fermeture fenêtre (si option active) ou argument `--minimized` | Réduit l'application dans la zone de notification Windows et exécute les synchronisations planifiées à intervalle régulier. |

---

## 3. Besoins fonctionnels

- **BF-01 — Connectivité double source** :
  - *Outlook* : Support du connecteur local **Outlook Classique (COM/MAPI)** et du connecteur cloud **Microsoft 365 (Graph API OAuth2)**.
  - *Google* : Connexion via **Google People API (OAuth2)** avec renouvellement automatique des jetons.
- **BF-02 — Modes de synchronisation** :
  - `OUTLOOK_TO_GOOGLE_ONLY` : Propagation unidirectionnelle d'Outlook vers Google (mode validé en production).
  - `GOOGLE_TO_OUTLOOK_ONLY` : Propagation unidirectionnelle de Google vers Outlook.
  - `MERGE_OUTLOOK_WINS` : Synchronisation bidirectionnelle avec priorité Outlook en cas de conflit.
  - `MERGE_GOOGLE_WINS` : Synchronisation bidirectionnelle avec priorité Google en cas de conflit.
  - `MERGE_PROMPT` : Synchronisation bidirectionnelle avec boîte de dialogue interactive à chaque conflit.
- **BF-03 — Préservation des champs avancés** :
  - Normalisation et formatage optionnel du champ `FileAs` (`Name SURNAME` ou `SURNAME, Name`).
  - Protection des notes Outlook au format riche (RTF) contre l'écrasement par du texte brut Google (sauf forçage explicite).
  - Synchronisation bidirectionnelle des photos de contact (comparaison par empreinte SHA-256).
  - Correspondance bidirectionnelle entre les catégories Outlook et les libellés (groupes) Google Contacts.
- **BF-04 — Sécurité des suppressions & Simulation** :
  - Mode **Dry-Run** permettant d'exécuter l'analyse complète et le journal prévisionnel sans aucune écriture.
  - Blocage de sécurité automatique si le nombre de suppressions dépasse le seuil configuré (`max_deletions_threshold`) ou si une source retourne 0 contact alors que la base de liaison n'est pas vide.
  - Sauvegarde préalable des contacts supprimés dans `backups/`.
- **BF-05 — Archivage local résilient (`.db` + `.xlsx`)** :
  - Création automatique dans `<dossier_exe>\contacts\` si les droits d'écriture sont présents.
  - Repli interactif via sélecteur de dossier Windows si le répertoire de l'exécutable est en lecture seule.
  - Surveillance de l'espace disque libre avant écriture : en cas d'espace insuffisant, renommage automatique des fichiers existants avec le suffixe `_NOT_UP_TO_DATE`.

---

## 4. Hypothèses fonctionnelles et techniques

### 4.1 Hypothèses fonctionnelles
- Chaque contact synchronisé est apparié de manière persistante via le couple `(outlook_entry_id, google_resource_name)`.
- Une synchronisation incrémentale (*Delta*) est utilisée par défaut ; une synchronisation complète (*Full Sync*) est déclenchée manuellement ou automatiquement si la dernière *Full Sync* date de plus de 7 jours.
- L'utilisateur dispose d'un profil Outlook configuré localement (pour le mode COM) ou d'un accès réseau HTTPS vers `graph.microsoft.com` et `people.googleapis.com`.

### 4.2 Hypothèses techniques
- **Architecture portable sans pollution** : L'exécutable `SyncContactsOG.exe` est autonome (`--onefile`). Aucun fichier technique n'est écrit dans le répertoire de l'exécutable.
- **Isolation des données techniques** : L'état de synchronisation (`sync_state.db`), les journaux (`sync_details.log`) et les sauvegardes (`backups/`) résident dans `%LOCALAPPDATA%\SyncContactsOG\`.
- **Sécurité des identifiants** : Les secrets OAuth (`credentials.json`) sont embarqués dans le binaire à la compilation, et les jetons d'accès utilisateur sont chiffrés dans le **Gestionnaire d'informations d'identification Windows** (*Windows Credential Manager*). Les préférences sont stockées dans le Registre Windows (`HKCU\Software\OutlookGoogleSyncPro`).
- **Instance unique** : Un verrou système Windows (`Local\SyncContactsOG_SingleInstance_Mutex`) interdit l'exécution simultanée de plusieurs instances ou versions afin d'écarter tout risque de verrouillage SQLite ou de désynchronisation.

---

## 5. Authentification Google (`credentials.json` & Jeton OAuth2)

### 5.1 Rôle des deux éléments d'authentification
1. **Fichier client OAuth (`credentials.json`)** :
   - Contient l'identifiant client (`client_id`) et le secret d'application (`client_secret`) déclarés dans Google Cloud Console pour l'accès à la **Google People API** (scope `https://www.googleapis.com/auth/contacts`).
   - **Version standard (`SyncContactsOG.exe`)** : `credentials.json` est déjà embarqué de manière chiffrée à l'intérieur de l'exécutable. L'utilisateur n'a aucun fichier `.json` à copier ou configurer.
   - **Version sans clé (`SyncContactsOG_NoKey.exe`)** : Destinée à un usage avec un projet Google Cloud personnel ; l'application invite à sélectionner un fichier `credentials.json` (type *Application de bureau / Desktop App*) lors de la première connexion, puis le mémorise de façon sécurisée.
2. **Jeton d'accès et de rafraîchissement (`token`)** :
   - Généré automatiquement après le consentement de l'utilisateur dans son navigateur web lors du clic sur `Tester la connexion` ou `Synchroniser maintenant`.
   - **Stockage sécurisé** : Le jeton OAuth2 (incluant le `refresh_token`) est enregistré de manière chiffrée dans le **Gestionnaire d'informations d'identification Windows** (*Windows Credential Manager*, entrée dédiée à `SyncContactsOG` et au compte Google connecté). Aucun fichier `token.json` en clair n'est laissé sur le disque.
   - **Renouvellement transparent** : Tant que l'accès n'est pas révoqué dans les paramètres de sécurité du compte Google, l'application renouvelle automatiquement le jeton expiré en arrière-plan sans redemander de connexion.

### 5.2 Déroulement de la fenêtre de connexion Google
1. Saisir l'adresse Gmail ou Google Workspace dans le champ **Compte Google**, puis cliquer sur **`Tester la connexion`** (ou lancer directement une synchronisation).
2. Si aucun jeton valide n'est présent pour cette adresse, l'application ouvre automatiquement votre **navigateur web par défaut** sur la page officielle de connexion Google (`accounts.google.com`) en écoutant sur un port local temporaire (`localhost`).
3. Dans le navigateur : choisir le compte Google correspondant, valider l'écran d'autorisation d'accès aux contacts (*Google People API*), puis fermer l'onglet dès que le message de confirmation s'affiche.
4. De retour dans **SyncContactsOG** : le champ e-mail est verrouillé en lecture seule pour éviter toute erreur de saisie accidentelle et le statut passe à **`✔ Connecté`**.
5. **Changement ou réinitialisation de compte** : Cliquer sur le bouton **`✕`** situé à droite du champ e-mail Google pour déverrouiller le champ, purger la session courante et associer un autre compte Google.

### 5.3 Chaîne de récupération et validation pré-compilation (`build.ps1`)
Lors de la génération d'une release standard (`SyncContactsOG.exe`), `build.ps1` recherche et valide `credentials.json` en ligne auprès des serveurs OAuth2 de Google (`validate_creds.py`) selon l'ordre de priorité suivant, en ignorant automatiquement tout fichier invalide ou révoqué :
1. `credentials.json` placé à côté de l'exécutable principal (`$ProjectDir\credentials.json`).
2. Sauvegarde locale (`C:\Temp\credentials_backup.json`).
3. Extraction depuis l'exécutable principal existant (`SyncContactsOG.exe`).
4. Extraction depuis les releases précédentes (`releases\SyncContactsOG_v*.exe`, de la plus récente à la plus ancienne).

---

## 6. Architecture & Flux d'exécution

1. **Contrôle préalable** : Vérification du Mutex d'instance unique, validation de l'adresse e-mail et contrôle des droits d'écriture du dossier d'archive si l'archivage local est activé.
2. **Collecte & Canonisation** : Lecture des contacts Outlook et Google, conversion de chaque fiche en une structure canonique normalisée et calcul de son empreinte SHA-256.
3. **Calcul du différentiel (Delta / Full)** : Comparaison des empreintes actuelles avec celles enregistrées dans `sync_state.db` (`contact_map`) pour identifier les créations, mises à jour, conflits et suppressions.
4. **Application & Journalisation** : Exécution des appels API (avec *backoff* exponentiel en cas de limite de débit HTTP 429) et affichage temps réel dans la console de droite.
5. **Post-traitement** : Mise à jour transactionnelle de `sync_state.db`, vérification de l'espace disque libre et génération atomique de `contacts_archive.db` et `contacts_archive.xlsx`.

---

## 7. Spécifications des Écrans & Guide d'Utilisation Détaillé

### 7.1 Vue d'ensemble de l'interface principale
L'interface est structurée en deux zones fixes côte à côte, surmontées d'une barre de titre et d'en-tête :
- **Barre d'en-tête** : Affiche le titre, le numéro de version (`vX.Y.Z`), le sélecteur de langue (`FR` / `EN`) et le basculeur de thème (`Dark` / `Light`). Tout changement est sauvegardé instantanément.
- **Colonne de gauche (Panneau de configuration et d'action)** : Regroupe 5 cartes fonctionnelles (*Comptes & Profils*, *Mode de Synchronisation*, *Options & Sécurité*, *Intégration Windows*, *Contacts exclusifs*).
- **Colonne de droite (Console de journalisation)** : Affiche en temps réel chaque étape et opération effectuée.

### 7.2 Détail par carte et écran secondaire

#### A. Carte 1 — Comptes & Profils
- **Compte Outlook** : Liste déroulante détectant automatiquement les profils Outlook locaux.
- **Sélecteur de connecteur (Boutons radio)** :
  - *Outlook Classique COM* : Lit directement le client Outlook installé sur le PC et affiche le nombre de contacts détectés entre parenthèses.
  - *Microsoft 365 Graph API* : Se connecte au cloud Microsoft 365 via navigateur (OAuth2) et affiche le compteur de contacts cloud.
- **Compte Google** : Champ texte pour l'adresse Gmail/Google Workspace (verrouillé en lecture seule une fois connecté, avec bouton `✕` pour réinitialiser le compte).
- **Bouton `Tester la connexion`** : Déclenche le flux OAuth2 décrit en section 5 et affiche `✔ Connecté`.
- **Synchro auto (heures)** : Sélecteur numérique (`0` à `168`). À `0`, la synchronisation est uniquement manuelle ; au-delà, une synchronisation automatique s'exécute toutes les $N$ heures.

#### B. Carte 2 — Mode de Synchronisation
- **Liste des 5 modes (Boutons radio avec pilule de statut à droite)** :
  - Chaque ligne affiche le nom du mode, un sous-titre explicatif et une pilule uniforme (`✔ Testé` ou `⚠ Non testé`) munie d'une infobulle au survol.
- **Bouton principal `Synchroniser maintenant`** : Lance le processus de synchronisation dans un thread dédié (désactivé pendant l'exécution pour éviter tout double déclenchement).

#### C. Carte 3 — Options & Sécurité
- **Utiliser le champ FileAs** : Active la synchronisation du nom d'affichage Outlook (`FileAs`).
- **Forcer le format FileAs** : Permet d'imposer le style `Name SURNAME` ou `SURNAME, Name` sur les contacts synchronisés.
- **Forcer l'écrasement des notes RTF** : Désactivé par défaut (et grisé en mode *Outlook → Google*) afin de préserver la mise en forme riche des notes Outlook.
- **Synchroniser les photos** : Active la comparaison et le transfert bidirectionnel des photos de profil.
- **Synchroniser les suppressions** : Répercute la suppression d'un contact sur l'autre plateforme (protégé par le compteur **Seuil max suppressions** qui bloque toute suppression massive accidentelle).
- **Confirmation visuelle avant suppression** : Ouvre une fenêtre de validation individuelle avant chaque suppression.
- **Forcer une synchro complète** : Ignore le cache incrémental et compare l'intégralité des champs de tous les contacts.
- **Archive locale (`.db` + `.xlsx`)** : Génère et met à jour la base SQLite avec photos et le classeur Excel dans `contacts\`.
- **Mode Simulation (Dry-Run)** : Exécute l'analyse complète et affiche le bilan prévisionnel dans le journal **sans modifier aucune donnée** dans Outlook ni Google.

#### D. Cartes 4a & 4b — Intégration Windows et Contacts exclusifs
- **Réduire dans la zone de notification** : Cliquer sur la croix `✕` de la fenêtre masque l'application dans le *System Tray* (icône près de l'horloge Windows) au lieu de la quitter. Un double-clic sur l'icône restaure la fenêtre ; un clic droit donne accès à *Ouvrir*, *Synchroniser maintenant* et *Quitter*.
- **Lancer au démarrage de Windows** : Inscrit l'application dans `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` avec l'argument `--minimized` pour démarrer silencieusement dans la zone de notification.
- **Contacts exclusifs & Force Mirror** : Affiche le décompte des contacts présents uniquement d'un côté. En mode unidirectionnel, l'option *Force Mirror* permet d'aligner strictement la cible sur la source en supprimant les contacts excédentaires de la cible.

#### E. Panneau droit — Console de journalisation
- **Barre de recherche et filtre** : Permet de filtrer instantanément les lignes du journal par mot-clé (nom d'un contact, erreur, action) ou par niveau (`Information`, `Warning`, `Error`).

#### F. Fenêtres modales (Écrans contextuels)
1. **Fenêtre de résolution de conflit (`MERGE_PROMPT`)** : S'affiche lorsqu'un même contact a été modifié des deux côtés depuis la dernière synchronisation. Présente côte à côte les valeurs Outlook et Google et propose de conserver *Outlook*, *Google* ou *Ignorer*.
2. **Fenêtre de confirmation de suppression** : Affiche le détail du contact sur le point d'être supprimé et demande confirmation (*Oui* / *Non*).
3. **Sélecteur de dossier d'archive (Lecture seule)** : Si `SyncContactsOG.exe` est lancé depuis un dossier sans droit d'écriture et que l'archive locale est cochée, une boîte de dialogue propose de choisir un dossier alternatif accessible en écriture.

### 7.3 Parcours d'utilisation types
1. **Premier lancement et validation sans risque** :
   - Lancer `SyncContactsOG.exe` $\rightarrow$ vérifier l'adresse Outlook détectée $\rightarrow$ saisir l'adresse Google et cliquer sur *Tester la connexion* $\rightarrow$ sélectionner *Outlook → Google (Unidirectionnel)* $\rightarrow$ cocher *Mode Simulation (Dry-Run)* $\rightarrow$ cliquer sur *Synchroniser maintenant* et contrôler le bilan dans le journal de droite.
2. **Synchronisation réelle et automatisation** :
   - Décocher *Mode Simulation (Dry-Run)* $\rightarrow$ cliquer sur *Synchroniser maintenant* $\rightarrow$ régler *Synchro auto (heures)* sur `4` (par exemple) $\rightarrow$ activer *Réduire dans la zone de notification* et *Lancer au démarrage de Windows*.
3. **Consultation des archives locales** :
   - Ouvrir le sous-dossier `contacts\` situé à côté de l'exécutable pour consulter `contacts_archive.xlsx` dans Excel ou interroger `contacts_archive.db`.

---

## 8. Exigences non fonctionnelles & Sécurité

- **Résilience API** : Gestion automatique des erreurs transitoires réseau et quotas Google People API avec temporisation progressive.
- **Intégrité transactionnelle** : Écriture de l'archive SQLite dans un fichier temporaire avant déplacement atomique (`shutil.move`) pour éviter toute corruption en cas d'interruption.
- **Ergonomie bilingue** : Interface et infobulles dynamiques en Français (`FR`) et Anglais (`EN`), thèmes Sombre et Clair.

---

## 9. Matrice de validation des modes (État courant)

| Mode de synchronisation | Code interne | Statut de validation | Badge UI |
| :--- | :--- | :--- | :--- |
| Outlook → Google (Unidirectionnel) | `OUTLOOK_TO_GOOGLE_ONLY` | Validé et testé en conditions réelles | `✔ Testé` |
| Google → Outlook (Unidirectionnel) | `GOOGLE_TO_OUTLOOK_ONLY` | En attente de validation terrain | `⚠ Non testé` |
| Bidirectionnel (Outlook prioritaire) | `MERGE_OUTLOOK_WINS` | En attente de validation terrain | `⚠ Non testé` |
| Bidirectionnel (Google prioritaire) | `MERGE_GOOGLE_WINS` | En attente de validation terrain | `⚠ Non testé` |
| Bidirectionnel (Demander à chaque conflit) | `MERGE_PROMPT` | En attente de validation terrain | `⚠ Non testé` |