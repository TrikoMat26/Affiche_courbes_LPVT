# Outil de Suivi et Analyse VPRT

Cet outil (script PowerShell `AfficheCourbeVPRT.ps1`) permet d'analyser, de visualiser et de suivre les rapports de test VPRT sous forme de graphiques détaillés, avec une interface graphique WinForms complète.

## Fonctionnalités Principales
- **Chargement des rapports SEQ-02** : Parcours d'un répertoire avec extraction automatique des données de test brutes.
- **Visualisation Graphique** : Affichage des courbes de mesure de précision (Voies U, V, W) adaptées dynamiquement à la taille de la fenêtre.
- **Mode Comparaison / Multi-Rapports** : Permet de superposer plusieurs relevés dans le temps ou par voie pour analyser la dérive et repérer des erreurs de composants ou de calibration.

## Nouveautés Récentes

### 1. Suivi des valeurs de résistances (U, V, W)
L'interface permet désormais de taguer chaque rapport de test avec les résistances (U, V, W) montées sur la carte lors de l'exécution du rapport.
- **Saisie simplifiée** : En cliquant sur la colonne `Résistances (U,V,W)` du tableau de rapports, une fenêtre s'ouvre pour attribuer les valeurs pertinentes.
- **Persistance** : Ces résistances sont stockées dans le fichier `resistances.json` rattaché à chaque nom de rapport.
- **Mise en évidence sur graphique** : Les modifications de résistances entre deux rapports consécutifs sont signalées lors du tracé sur les courbes.

### 2. Bibliothèque des Codes CMS ("SMD Markings")
Le bouton `🔗 Editer Codes CMS` ouvre un gestionnaire indépendant (non-bloquant) listant la correspondance entre les **Valeurs de résistance normalisées** (ex: 332Ω) et le marquage imprimé sur les composants type CMS (ex: `51A`).
- **Gestion JSON** : Les correspondances sont totalement éditables et sauvegardées instantanément dans le fichier `smd_markings.json`.
- **Lien avec la saisie** : Les listes de sélection de la saisie des résistances s'appuient dynamiquement sur ces valeurs (pour forcer une normalisation de la série).

### 3. Affichage visuel renforcé des relevés (Points uniques)
- **Tolérance graphique** : Si un rapport dispose d'un enregistrement prématuré où la courbe d'erreur ne possède qu'**un seul point valide** mesuré (par ex. arrêt anticipé à la roue codeuse), ce point *reste affiché* sous forme de cercle plein sur les graphiques, contrairement au comportement natif des lignes `Line` qui ne dessinent rien en dessous de 2 points.
