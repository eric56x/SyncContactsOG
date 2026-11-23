# Changelog - SyncContactsOG

All notable changes to **SyncContactsOG** are documented in this file.

## [1.1.1] - 2026-11-23
### Fixed
- **Unified Portable Archive Folder**: Replaced legacy `SyncContactsOG_Archive` references with `.\contacts\` (storing `contacts.db` and `contacts.xlsx`).
- **Documentation & Release Standardization**: Unified `CHANGELOG.md` history, lowercase `docs/` assets (Markdown + HTML), and GitHub Release packaging with SHA-256 checksums.

## [1.1.0] - 2026-11-23
### Added
- **Google People API v1 Engine**: Native OAuth 2.0 flow (`https://www.googleapis.com/auth/contacts`) with automatic token refresh (`.\data\google_token.json`).
- **Guided OAuth Setup Dialog**: Live `credentials.json` validation, People API connectivity verification, blue alert highlighting on failed diagnostics, and step-by-step Google Cloud Console guide.
- **Portable Storage Layout**: Isolated `.\data\` (OAuth credentials & token) and `.\contacts\` (SQLite/Excel backups) relative to the executable.
- **Bilingual UI (EN / FR) & Compact Layout**: Dynamic `Configure Google Connection` / `Configurer la Connexion Google` button with status indicator (`✓` / `✗`) and maximized log panel height.

## [1.0.11] - 2026-11-22
### Corrigé
- Calcul et alignement exact de la largeur de **Synchronization Mode** (card_mode) sur la largeur totale de la colonne inférieure (card_opts + win_stack), alignant son bord droit au pixel près avec **Light / Dark**, **Windows Integration** et **Exclusive contacts**.
## [1.0.10] - 2026-11-22
### Corrigé
- Suppression du widget parasite sous l'option **Simulation Mode** en bas à gauche.
- Alignement horizontal exact du bord inférieur des trois conteneurs inférieurs (**Options & Data Integrity**, **Exclusive contacts** et la zone noire des logs à droite).
## [1.0.9] - 2026-11-22
### Corrigé
- Ajustement de la hauteur (34 px) et de la marge inférieure de la ligne Auto-Sync dans **Accounts & Profiles** afin que [ - ] [ 6 ] [ + ] s'affiche sans aucune coupure DPI.
- Extension verticale (expand=True, ill="both", weight=1) des cartes inférieures gauches (**Options & Data Integrity** et **Exclusive contacts**) pour aligner leur bord inférieur sur la zone noire des logs à droite.
## [1.0.8] - 2026-11-22
### Corrigé
- Suppression de la hauteur fixe sur la ligne Auto-Sync et ajout de marge basse symétrique dans les deux cartes supérieures pour afficher intégralement [ - ] [ 6 ] [ + ] sans aucune coupure.
- Étirement vertical (weight=1, sticky="nsew") de la rangée inférieure gauche (**Options & Data Integrity** et **Exclusive contacts**) afin que leur bord inférieur soit aligné avec le bas de la zone noire des logs à droite.
## [1.0.7] - 2026-11-22
### Corrigé
- Augmentation symétrique de l'espacement inférieur dans les boîtes **Accounts & Profiles** et **Synchronization Mode** afin que le sélecteur [ - ] [ 6 ] [ + ] ne soit plus rogné en bas.
## [1.0.6] - 2026-11-22
### Corrigé
- Suppression du doublon de version dans le grand titre d'en-tête (SyncContactsOG uniquement dans l'en-tête, SyncContactsOG — v1.0.6 conservé dans la barre de titre Windows).
- Reconstruction en grille 3 colonnes (102x28 px, grid_propagate(False)) du sélecteur numérique Auto-Sync afin d'afficher simultanément [ - ], le champ d'heures et [ + ].
## [1.0.5] - 2026-11-22
### Corrigé
- Affichage permanent de la version (SyncContactsOG — v1.0.5 dans la barre de titre Windows et SyncContactsOG   v1.0.5 dans l'en-tête principal, persistant lors du basculement FR/EN).
- Suppression de la contrainte de largeur fixe sur le conteneur du sélecteur Auto-Sync afin d'afficher en entier [ - ] [ 6 ] [ + ].
## [1.0.4] - 2026-11-22
### Modifié
- Ergonomie du sélecteur d'heures Auto-Sync : le champ numérique est désormais encadré par le bouton **-** à gauche et le bouton **+** à droite ([ - ] [ 6 ] [ + ]).
## [1.0.3] - 2026-11-22
### Corrigé
- Alignement horizontal côte à côte (- et +, 26x26 px) des boutons de réglage de l'intervalle Auto-Sync pour éviter le chevauchement vertical.
## [1.0.2] - 2026-11-22
### Corrigé
- Affichage effectif de la version (SyncContactsOG — v1.0.2 dans la barre de titre et badge bleu 1.0.2 à côté du titre principal).
- Remplacement des micro-flèches comprimées (▲/▼) du réglage d'intervalle Auto-Sync par des boutons + et - lisibles.
## [1.0.1] - 2026-11-21
### Ajouté
- Boutons **Tout sélectionner** / **Tout désélectionner** (FR/EN) dans la fenêtre *Contacts présents sur un seul compte* à côté du bouton *Supprimer la sélection*.
- Affichage visible de la version (1.0.1) dans la barre de titre de la fenêtre et dans l'en-tête principal à côté du titre SyncContactsOG.
## [1.0.0] - 2026-11-21
### Version Majeure — Première Release de Production
- **Architecture & Build** : séparation stricte acine / contacts_archive / doc / eleases / source, avec génération systématique des deux exécutables (SyncContactsOG_v1.0.0.exe avec identifiants OAuth2 intégrés et SyncContactsOG_v1.0.0_NoCreds.exe sans identifiants).
- **Documentation & Blueprint** : génération automatique des 4 guides HTML (SPECIFICATIONS_FR.html, SPECIFICATIONS_EN.html, CHANGELOG.html, BEST_PRACTICES.html) dans doc\ et archivage complet dans eleases\.
- **Interface & Fiabilité** : affichage dynamique de la version (SyncContactsOG — v1.0.0) dans le titre et en pied de page, libellés bilingues FR/EN harmonisés, barre de progression strictement monotone et auto-réparation SQLite (PRAGMA integrity_check) sur contacts_archive.db.
## [0.2.12] - 2026-11-21
### Ajouté & Corrigé
- Génération systématique de deux exécutables par version : SyncContactsOG_vX.Y.Z.exe (avec credentials.json intégré) et SyncContactsOG_vX.Y.Z_NoCreds.exe (sans credentials.json).
- Affichage dynamique de la version (__version__) dans le titre de la fenêtre et en discret badge en bas à droite de l'interface.
- Arrêt silencieux de l'instance active via Get-Process dans uild.ps1 (sans erreur lorsque l'application est déjà fermée).
## [0.2.11] - 2026-11-21
### Corrigé
- Mise à jour de la traduction française du commutateur d'archivage local : Conserver une Archive Locale (.db avec photos + .xlsx) (avec conservation de l'info-bulle au survol).
## [0.2.10] - 2026-11-21
### Modifié
- Clarification du libellé de l'option d'archivage local dans l'interface : Keep a Local Archive (.db with photos + .xlsx) (avec conservation de l'info-bulle explicative au survol).
## [0.2.9] - 2026-11-21
### Corrigé
- Vérification d'intégrité SQLite (PRAGMA integrity_check) avant toute mise à jour incrémentale de contacts_archive.db, avec reconstruction complète automatique en cas de fichier endommagé (database disk image is malformed).
## [0.2.8] - 2026-11-21
### Corrigé
- Alignement de la barre de progression globale lors de la génération de l'archive locale post-synchronisation (95 % – 99 % — Étape 5/5 au lieu d'un retour visuel à 20 % — Étape 2/5).
- Correction de l'interpolation dynamique des compteurs de sous-étape lors de la lecture COM Outlook et de la normalisation Google.
## [0.2.6] - 2026-11-21
### Modifié
- Restauration de l'agencement équilibré des panneaux de configuration et de la console.
- Positionnement des boutons Nettoyer directement sous chaque ligne respective (Contacts exclusifs Outlook et Contacts exclusifs Google) avec une largeur confortable (115 px) pour éviter toute troncature.
## [0.2.5] - 2026-11-21
### Modifié
- Harmonisation de la liste déroulante du format Classer sous (Prénom NOM / NOM, Prénom) avec le style neutre discret du bouton Tester la connexion.
- Élargissement du bouton Nettoyer (section *Contacts exclusifs*) pour afficher intégralement le libellé sans troncature.
## [0.2.4] - 2026-11-21
### Modifié
- Atténuation visuelle du bouton Tester la connexion et de la liste déroulante du format Classer sous (Prénom NOM / NOM, Prénom) avec un style neutre discret adapté aux thèmes Clair et Sombre, afin de réserver le bleu vif à l'action principale Lancer maintenant.
## [0.2.3] - 2026-11-21
### Modifié & Ajouté
- Harmonisation visuelle de la pilule ✔ Testé au même format gris neutre et largeur fixe (84px) que ⚠ Non testé.
- Ajout de SPECIFICATIONS.md (spécifications fonctionnelles, techniques et guide détaillé des écrans) et génération automatique de la version HTML autonome à chaque release.
## [0.2.2] - 2026-11-21
### Modifié & Ajouté
- Ajout de pilules de statut de largeur fixe (84px) à droite de chaque mode de synchronisation : ✔ Testé (blanc) pour Outlook → Google et ⚠ Non testé (niveaux de gris) avec infobulle d'avertissement au survol pour les autres modes.
- Refonte visuelle des boutons radio (modes de synchronisation et connecteurs Outlook COM / Microsoft 365 Graph API) avec anneau fin clair et point central blanc.
## [0.2.1] - 2026-11-21
### Ajouté
- Protection d'instance unique via Mutex Windows (Local\SyncContactsOG_SingleInstance_Mutex) empêchant toute exécution simultanée de plusieurs versions ou fenêtres de l'application.
## [0.2.0] - 2026-11-21
### Modifié & Ajouté
- Architecture portable propre : déplacement des fichiers techniques (sync_state.db, sync_details.log, ackups/) vers %LOCALAPPDATA%\SyncContactsOG\.
- Stockage automatique de l'archive locale (contacts_archive.db + .xlsx) dans un sous-dossier dédié SyncContactsOG_Archive\ à côté de l'exécutable.
- Repli interactif avec sélecteur de dossier si le répertoire de l'exécutable est en lecture seule.
- Détection d'espace disque insuffisant avec renommage automatique de l'archive existante en *_NOT_UP_TO_DATE.
## [0.1.0] - 2026-11-21
### Ajouté
- Synchronisation bidirectionnelle et unidirectionnelle Outlook (COM / Graph API) <-> Google Contacts.
- Stockage sécurisé du jeton OAuth et de la clé client JSON via Windows Credential Manager (DPAPI).
- Assistant visuel en 4 étapes (Projet, Branding/Audience, Client OAuth, Importation) avec sélecteur de compte Google anti-redirection SSO M365.
- Interface bilingue FR/EN, thème Clair/Sombre, mappage dynamique du format « Classer sous » (`Prénom NOM` / `NOM, Prénom`), et indicateur System Tray orange pour les contacts exclusifs.
