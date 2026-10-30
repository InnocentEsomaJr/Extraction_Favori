# SNIS RDC - Dashboard Performance

Application Streamlit pour analyser la performance SNIS (DHIS2) avec:
- complétude
- promptitude
- comparaison entre zones de santé
- filtrage hiérarchique `Pays` / `Province` / `Zone de santé`
- analyse des violations de règles de validation
- export commenté en **Excel / PowerPoint**

Le script principal est: `dashbord.py`.

## 1. Fonctionnalités

Le dashboard est organisé en 5 onglets:

1. **Base de données**
- Affichage des données brutes (sans `Organisation unit ID`)
- Données des zones de santé du périmètre sélectionné

2. **Complétude**
- Calcul `Reports_Actual`, `Reports_Attendu`, `Complétude_Globale (%)`
- Mise en forme couleur selon seuils
- Graphique de classement des zones

3. **Promptitude**
- Calcul `Promptitude_Globale (%)`
- Score du nombre de datasets avec promptitude `>= 95%`

4. **Analyse Comparative**
- Tableau comparatif complétude/promptitude
- Quadrant de performance
- Top 5 complétude / Flop 5 promptitude
- Tableau fusionné (zone filtrée): indicateurs dataset reporting + actual

5. **Éléments de catégorisation**
- Violations de règles par zone de santé
- Colonnes M-1 / M:
  - `Règles violées (M-1)`
  - `Règles corrigées (M-1 -> M)`
  - `Règles violées (M)`
- `Ratio / 100 rapports` calculé par:
  - `(Règles violées (M) / Reports_Actual) * 100`
- `Score de qualité`
- `Score de qualité /20`, placé juste après `Score de qualité` dans le tableau et l'export Excel.
  Le barème s'applique au score de qualité (moyenne de la complétude et de la promptitude) :
  - moins de 50 % : **0**
  - de 50 à moins de 60 % : **5**
  - de 60 à moins de 80 % : **10**
  - de 80 à moins de 95 % : **15**
  - 95 % et plus : **20**
- **Catégorisation issue du modèle Excel** : toutes les colonnes précédentes sont conservées,
  puis le tableau affiche le barème DPS/ZS, le nombre de datasets à complétude `>=95 %`,
  les anciennes violations restantes, les cinq notes détaillées et leur somme
  `Score catégorisation /100`. Le score de qualité historique et sa note `/20` sont des
  repères distincts et ne sont pas ajoutés au total.

| Critère | Barème actif sur 20 |
| --- | --- |
| Complétude et promptitude | `<50 : 0` ; `50–<60 : 5` ; `60–<80 : 10` ; `80–<95 : 15` ; `>=95 : 20` |
| Datasets à complétude >=95 % (DPS) | `<6 : 0` ; `6–7 : 5` ; `8–9 : 10` ; `10–11 : 15` ; `>=12 : 20` |
| Datasets à complétude >=95 % (ZS) | `<5 : 0` ; `5 : 5` ; `6–7 : 10` ; `8–9 : 15` ; `>=10 : 20` |
| Violations pour 100 rapports | `0–5 : 20` ; `>5–10 : 15` ; `>10–15 : 10` ; `>15–20 : 5` ; `>20 : 0` |
| Corrections de M-1 | Aucune ancienne violation restante : `20` ; au moins une restante : `0` |

Le mode **Automatique** déduit DPS/ZS du niveau DHIS2 de chaque unité. Le sélecteur
permet de choisir explicitement DPS ou ZS si la hiérarchie est indisponible.
Les règles sont consultables dans l'expandeur « Consulter les barèmes de catégorisation ».
Les trois ambiguïtés du modèle ont été clarifiées : promptitude `>=95 % = 20`,
ratio exactement `20 = 5`, et corrections notées sur les violations **restantes**.
Les notes saisies manuellement dans les exemples du modèle ne remplacent pas les barèmes.

Chaque dataset est compté une seule fois pour l'unité analysée. Une ligne directe de
l'unité prime sur ses descendants ; à défaut, son taux est recalculé à partir des
rapports réels et attendus des descendants. Sans données suffisantes, le compte reste
indisponible. Une erreur DHIS2, l'absence de rapports ou une composante manquante laisse
le total vide, avec un motif dans `Statut catégorisation`.

Les résultats restent en session lors d'un changement de barème ou de l'ouverture
de l'aperçu. Changer de favori, période, groupe de règles ou périmètre impose une
nouvelle analyse pour éviter un résultat périmé, tout comme un changement des données
sources. La déconnexion efface ce résultat.
Dans l'onglet de catégorisation, les calendriers **Date de début** et **Date de fin**
définissent l'intervalle M. La référence M-1 reste le mois entier précédant celui de
la date de début. Chaque ligne renvoyée par DHIS2 est conservée et comptée.

Le bloc **Unité d’organisation parent** permet de choisir le pays, puis une province
et éventuellement une zone de santé. Les parents possibles restent les niveaux pays,
province et ZS ; leurs centres et postes apparaissent dans le détail des violations.
Les dates et le parent sont initialisés à partir de la barre latérale et peuvent ensuite
être ajustés pour cette analyse. Ces filtres pilotent sa synthèse de catégorisation et
son détail ; les quatre premiers onglets suivent les filtres de la barre latérale.

Les requêtes de validation utilisent les dates exactes, y compris un même jour ou un
intervalle qui traverse deux années. DHIS2 renvoie les violations des périodes qui
chevauchent ces dates. Les rapports, la complétude, la promptitude et les datasets
sont recalculés pour le même parent et les mois couverts ; les données de déclaration
mensuelles restent celles de mois entiers, sans prorata journalier. Une date manquante
ou un début après la fin désactive le lancement et retire les anciens résultats de règles.

Pour consulter les résultats de catégorisation :

- **Rechercher une zone** ignore les accents, la casse et les espaces supplémentaires.
- **Afficher les résultats** permet de garder les scores complets, les scores incomplets
  ou les scores inférieurs à 80, sans modifier les notes.
- **Trier les résultats** propose l'ordre initial, le score croissant/décroissant ou le nom.
- Le nombre de lignes affichées est indiqué. La recherche, le filtre de score et le tri
  concernent uniquement le tableau à l'écran : les indicateurs, l'aperçu et l'export Excel
  gardent l'analyse complète du périmètre géographique sélectionné.

L'analyse affiche sa progression par zone et conserve sa date de réalisation, en heure
de Kinshasa, lors des changements de vue. L'option **Toutes les règles (Global)** reste
disponible même si aucun groupe de règles n'est configuré. Si la liste des groupes est
indisponible, un message le précise sans bloquer cette option globale.

Le tableau **Détail des règles violées**, sous la synthèse de catégorisation, présente :

- la zone de santé et l'unité d'organisation concernée (centre de santé, poste, etc.) ;
- la période, l'importance (`HIGH`, `MEDIUM`, `LOW`) et le libellé de la règle ;
- les valeurs gauche/droite et l'opérateur de comparaison ;
- les instructions, les expressions et les combinaisons disponibles dans les détails.

Les lignes suivent le parent pays/province/zone, les calendriers et le groupe de règles
choisis dans cet onglet.
Une recherche et un filtre d'importance affinent ce tableau. L'expandeur
« Consulter le détail d'une violation » affiche les explications complètes de la ligne.
Le CSV télécharge les lignes affichées ; la feuille Excel `Details_Violations` et
l'onglet d'aperçu « Règles détaillées » gardent tous les détails du périmètre analysé.
La feuille `Contexte_Validation` indique les dates appliquées, le parent et son ID,
le groupe de règles, les mois des indicateurs, la référence M-1 et le barème utilisé.
Les lignes conservent l'ordre, les périodes, les valeurs et les libellés renvoyés par DHIS2,
y compris les règles dont la description est vide. Une violation présente sur deux
périodes conserve deux lignes et compte deux fois dans la synthèse. Les signatures
sans période servent uniquement à comparer les corrections entre M-1 et M.

Les analyses M et M-1 sont calculées à la demande par `POST /api/dataAnalysis/validationRules`,
avec un corps JSON contenant `ou`, `startDate`, `endDate`, `persist: false`,
`notification: false` et éventuellement `vrg`, comme l'application officielle DHIS2.
Une seule requête couvre le parent sélectionné et l'intervalle complet ; la requête
M-1 utilise le même parent et le mois de référence entier. Les unités descendantes
sont sélectionnées par DHIS2. Les chemins hiérarchiques ne sont pas obligatoires
dans la réponse ; les métadonnées servent à répartir les lignes par zone de santé.
Si ce rattachement manque, les lignes restent affichées et les comptes par zone
restent indisponibles. Un ancien serveur répondant HTTP 404, 405 ou 501 au POST
est interrogé en GET avec les mêmes paramètres. Un refus 401/403 reste une erreur.
Ces analyses ne modifient pas les données DHIS2
et ne déclenchent pas de notifications. Le nombre de règles corrigées correspond aux
signatures de M-1 absentes des résultats M calculés lors de la même analyse.
Chaque clic sur **Lancer l'analyse des violations** effectue de nouvelles requêtes
DHIS2. Les résultats restent en session pendant la consultation, sans nouvel appel
lors d'une recherche, d'un tri ou d'un export. **Actualiser les données** recharge
aussi les indicateurs de déclaration et les métadonnées.
L'expandeur **Vérifier la requête envoyée à DHIS2** présente le compte, le parent,
les dates, le groupe et le nombre de lignes reçues pour M et M-1. Les erreurs
indiquent la méthode et le statut HTTP, sans afficher les identifiants de connexion.
Références : [appel dans l'application officielle](https://github.com/dhis2/data-quality-app/blob/54a3bf8ed25e051fbea00ceef715df81812b3e62/src/pages/ValidationRulesAnalysis/ValidationRulesAnalysis.jsx),
[conversion des résultats serveur](https://github.com/dhis2/data-quality-app/blob/54a3bf8ed25e051fbea00ceef715df81812b3e62/src/pages/ValidationRulesAnalysis/convert-element-from-api-response.js).

Aucune limite personnalisée n'est envoyée par défaut, comme dans l'application DHIS2.
Une réponse atteignant 500 lignes déclenche un avertissement prudent de limite possible,
selon le seuil de l'application officielle ; le serveur peut appliquer sa propre limite.
Les lignes reçues restent consultables, mais le décompte et les scores concernés restent
indisponibles. Un groupe de règles plus précis peut réduire le nombre de résultats.
Si les libellés ou instructions ne sont pas accessibles, les identifiants disponibles
restent visibles avec un message explicatif ; aucune valeur manquante n'est remplacée par zéro.

Le filtrage s'applique à tous les tableaux, graphiques, exports et règles de validation :

- **Pays** : toutes les zones de santé du pays sélectionné.
- **Province** : toutes les zones de la province choisie ; le filtre « Zone de santé »
  permet ensuite de garder toutes les zones ou d'en sélectionner une.
- **Zone de santé** : choix de la province puis d'une seule zone.

Le catalogue DHIS2 des niveaux 1/2/3 alimente les sélecteurs, indépendamment du niveau
enregistré dans le favori. Les données sont demandées au niveau des zones de santé,
même si le favori d'origine affiche des provinces. La sélection utilise les IDs :
les zones homonymes sont distinguées par province ou identifiant. Un changement de
province réinitialise une zone qui n'appartient plus au périmètre. Une province ou un
pays ne sont jamais utilisés comme ligne de zone de santé quand la hiérarchie manque.

Le chargement reprend les éléments et filtres explicites du favori via l'API Analytics,
avec une dimension organisationnelle bornée au pays, à la province ou à la zone.
Les périodes sont agrégées par DHIS2 avant le calcul des taux. Les filtres complexes
non représentables sont signalés, sans retirer silencieusement une restriction du favori.
Référence : [dimensions et niveaux de l'API DHIS2](https://docs.dhis2.org/en/develop/using-the-api/dhis-core-version-242/analytics.html).

Le bouton **Actualiser les données**, dans la barre latérale, recharge catalogue,
données et groupes pour la session en cours. Il invalide son ancienne catégorisation
et ferme l'aperçu : l'analyse des violations doit être relancée. Cette actualisation
ne purge pas le cache de données des autres sessions. Une sélection de mois inversée
interrompt le chargement et retire les anciens résultats jusqu'à correction.

## 2. Export de rapport

Depuis la **sidebar**:
- Choix du type de téléchargement: `Excel`, `PowerPoint`
- Bouton de téléchargement dynamique selon le type choisi
- Bouton `Visualiser le rapport` pour consulter l'aperçu complet sans télécharger

Le rapport exporté inclut:
- tableaux principaux
- commentaires automatiques (lecture des graphiques et tableaux)
- conservation des colorations conditionnelles du dashboard dans les tableaux exportés
- catégorisation complète dans `Resultats_Regles` et légende dans `Baremes_Categorisation`
- détail des violations de la période M dans `Details_Violations`
- dates, unité parent et référence de l'analyse dans `Contexte_Validation`
- (PowerPoint) images des graphiques et tableaux
- feuilles/slides dédiées: base de données, rapports détaillés (réels/attendus), performance finale, promptitude, comparatif, top/flop, résultats des règles

## 3. Prérequis

- Python 3.10+ recommandé
- Accès DHIS2 valide (URL, utilisateur, mot de passe)

## 4. Installation

Depuis le dossier `Extraction_Favori`:

```powershell
python -m pip install -r requirements.txt
```

## 5. Configuration DHIS2

Créer le fichier `Extraction_Favori/.streamlit/secrets.toml`:

```toml
DHIS2_URL = "https://votre-instance-dhis2"
# Optionnel (connexion lente):
# DHIS2_TIMEOUT_CONNECT = 10
# DHIS2_TIMEOUT_READ = 120
# DHIS2_HTTP_RETRIES = 2
```

Le projet lit `DHIS2_URL` via `st.secrets`.
Le **nom d'utilisateur** et le **mot de passe** sont saisis par chaque utilisateur dans la barre latérale (`🔐 Connexion DHIS2`).

## 6. Lancement

```powershell
streamlit run dashbord.py
```

Ensuite:
1. Choisir le favori DHIS2 (ou un ID personnalisé)
2. Choisir la période (année + mois début/fin)
3. Choisir le pays, une province ou une zone de santé
4. Consulter les onglets et exporter le rapport

## 7. Structure du dossier

```text
Extraction_Favori/
|- dashbord.py
|- categorisation.py
|- categorisation_view.py
|- reporting_metrics.py
|- geography.py
|- scoped_analytics.py
|- validation_api.py
|- validation_details.py
|- validation_filters.py
|- validation_scope.py
|- tests/
|- requirements.txt
|- .streamlit/
|  |- secrets.toml
|- .gitignore
```

Vérification des calculs, de l'export et des interactions Streamlit sans accès DHIS2 :

```powershell
python -B -m unittest discover -s tests
```

## 8. Dépannage rapide

- **PowerPoint ne génère pas les images**
  - vérifier `kaleido` et `matplotlib` dans l’environnement

- **Erreur de connexion DHIS2**
  - vérifier `DHIS2_URL` dans `secrets.toml`
  - vérifier les identifiants saisis dans `🔐 Connexion DHIS2`

- **Pas de données**
  - vérifier l’ID du favori DHIS2
  - vérifier la période sélectionnée
  - vérifier les droits utilisateur DHIS2

## 9. Dépendances utilisées

`requirements.txt`:
- `streamlit`
- `pandas`
- `numpy`
- `plotly`
- `requests`
- `xlsxwriter`
- `python-docx`
- `python-pptx`
- `matplotlib`
- `kaleido`
