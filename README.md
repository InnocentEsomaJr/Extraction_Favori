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
- `Moyenne complétude / promptitude (%)`
- `Note moyenne des taux /20`, repère conservé à droite du tableau.
  Le barème des taux s'applique séparément à la complétude et à la promptitude :
  - moins de 50 % : **0**
  - de 50 à moins de 60 % : **5**
  - de 60 à moins de 80 % : **10**
  - de 80 à moins de 95 % : **15**
  - 95 % et plus : **20**
- **Catégorisation issue du modèle Excel** : toutes les colonnes précédentes sont conservées,
  puis le tableau affiche le barème DPS/ZS, le nombre de datasets à complétude `>=95 %`,
  les anciennes violations restantes, les cinq notes détaillées et leur somme
  `Score qualité /100`. Ce total prend les valeurs 0, 5, 10, …, 95, 100.
  La moyenne des deux taux et sa note `/20` restent des repères distincts.

| Critère | Barème actif sur 20 |
| --- | --- |
| Complétude et promptitude | `<50 : 0` ; `50–<60 : 5` ; `60–<80 : 10` ; `80–<95 : 15` ; `>=95 : 20` |
| Datasets à complétude >=95 % (DPS) | `<6 : 0` ; `6–7 : 5` ; `8–9 : 10` ; `10–11 : 15` ; `>=12 : 20` |
| Datasets à complétude >=95 % (ZS) | `<5 : 0` ; `5 : 5` ; `6–7 : 10` ; `8–9 : 15` ; `>=10 : 20` |
| Violations pour 100 rapports | `0–5 : 20` ; `>5–10 : 15` ; `>10–15 : 10` ; `>15–20 : 5` ; `>20 : 0` |
| Violations de M-1 | Aucune violation en M-1 : `20` ; au moins une violation en M-1 : `0`, même corrigée en M |

Le mode **Automatique** déduit DPS/ZS du niveau DHIS2 de chaque unité. Le sélecteur
permet de choisir explicitement DPS ou ZS si la hiérarchie est indisponible.
Les règles sont consultables dans l'expandeur « Consulter les barèmes de catégorisation ».
Les trois ambiguïtés du modèle ont été clarifiées : promptitude `>=95 % = 20`,
ratio exactement `20 = 5`, et note M-1 selon la loi du tout ou rien sur le nombre
de violations **présentes en M-1**, indépendamment de leur disparition en M.
Les notes saisies manuellement dans les exemples du modèle ne remplacent pas les barèmes.

Chaque dataset est compté une seule fois pour l'unité analysée. Une ligne directe de
l'unité prime sur ses descendants ; à défaut, son taux est recalculé à partir des
rapports réels et attendus des descendants. Sans données suffisantes, le compte reste
indisponible. Une erreur DHIS2, l'absence de rapports ou une composante manquante laisse
le total vide, avec un motif dans `Statut catégorisation`.

Les résultats restent en session lors d'un changement de barème ou de l'ouverture
de l'aperçu. Changer de favori, période, groupe de règles ou périmètre impose une
nouvelle analyse pour éviter un résultat périmé. L'actualisation des données et la
déconnexion effacent aussi le résultat et la progression sauvegardée. Une mise à jour
des seuls indicateurs de déclaration ne supprime pas les réponses de validation en cours.
Dans l'onglet de catégorisation, les calendriers **Date de début** et **Date de fin**
définissent l'intervalle M. La référence M-1 reste le mois entier précédant celui de
la date de début. Chaque ligne renvoyée par DHIS2 est conservée et comptée.

Le **Tableau de score qualité** apparaît automatiquement, avant toute analyse des
violations. Il suit les filtres pays, province, zone et dates de cet onglet.
Une province affiche ses zones, le pays affiche ses DPS par défaut (avec le choix
**Détail du tableau qualité → Zones de santé** pour afficher toutes les ZS) et
une zone sélectionnée affiche uniquement ses propres indicateurs. Les DPS utilisent
les métriques Analytics de niveau 2 et le barème DPS ; les zones utilisent le niveau 3
et le barème ZS. Les jointures
utilisent les IDs DHIS2, y compris en présence de noms identiques.

Pour les mêmes mois, le tableau reprend en priorité les chiffres déjà calculés
dans Complétude, Promptitude et Analyse Comparative : taux, rapports réels et
nombre de datasets à complétude ≥95 %. Une zone peut être sélectionnée dans la
province ou le pays déjà chargé, sans relancer une seconde requête. Un dataset
inutilisé dont les cellules sont vides ne supprime plus les chiffres de cette synthèse.
Les ID exacts et un rattachement non ambigu sont obligatoires ; un zéro créé en
l'absence de toute donnée brute n'est pas réutilisé comme mesure.

Si les mois ou les zones demandés ne sont pas disponibles dans cette synthèse,
Analytics DHIS2 complète seulement les métriques manquantes, pour le périmètre
et la période choisis. Si le favori ne demande que les taux, cette requête complète
les cinq métriques nécessaires :
rapports réels, attendus, reçus à temps, taux de complétude et taux de promptitude.
Les indicateurs génériques ne sont pas assimilés à ces comptes. Les trois comptes
servent aux taux pondérés ; la moyenne historique de complétude et promptitude,
avec sa note /20, reste distincte du score qualité /100. Les autres onglets gardent
la sélection du favori. Les datasets entièrement vides pour une unité sont ignorés
pour cette unité ; une valeur manquante dans un dataset actif reste indisponible.

Les données de déclaration existantes sont réutilisées seulement si leurs unités,
leurs métriques et leurs mois correspondent à la sélection. Sinon,
une requête qualité dédiée complète la source. Un favori sans dataset de reporting,
une réponse manquante ou des comptes incomplets sont signalés dans
**Voir les données manquantes du tableau qualité**, sans fabriquer de valeurs.
L'écran affiche « — » pour les données indisponibles ; les exports restent numériques.
Les notes de qualité, complétude, promptitude et datasets restent consultables si
une analyse des violations échoue. Les colonnes de violations et le total /100
attendent leurs analyses respectives ; changer de périmètre retire les anciens
résultats des règles sans retirer les nouveaux indicateurs de déclaration.

Le tableau aligne chaque mesure avec sa note, dans l'ordre : complétude globale,
promptitude globale, nombre de datasets ≥95 %, violations pour 100 rapports et
corrections de M-1, puis le **Score qualité /100**, nommé de même dans les exports.
La moyenne des deux taux est conservée à droite sous son nom explicite.

Les trois boutons utilisent les mêmes filtres pays, province, zone et dates.
Les analyses sont réparties par zone de santé pour éviter une requête volumineuse
sur toute une province. Au niveau pays, les résultats sont regroupés par DPS
(ou par ZS selon le détail choisi).

**Calculer tous les scores (M et M-1)** exécute les deux périodes en une action.
Une réponse vide réussie donne zéro violation ; une erreur
affecte seulement l'unité concernée. Les autres résultats et les indicateurs de
déclaration restent affichés. Aucun échec ne déclenche de relance automatique.
Les erreurs détaillées et les requêtes par unité restent consultables.

Les réponses sont sauvegardées **après chaque zone**, puis le tableau et les
détails sont affichés avant de poursuivre la zone suivante. Une notification
indique la fin du traitement ou son interruption ; les compteurs distinguent
les unités traitées, indisponibles et en attente. Ces résultats restent dans la
session, jusqu'à déconnexion, actualisation des données ou changement de filtres.

En cas d'interruption, **Reprendre les zones restantes** conserve les réponses
déjà reçues. Une requête interrompue sans réponse est signalée comme indisponible
et n'est pas renvoyée automatiquement. Les zones non encore interrogées sont
reprises ; elles restent sans décompte avant réception, jamais à zéro par défaut.
Les trois boutons d'origine sont conservés. Pendant une analyse active, ils sont
désactivés afin de ne pas mélanger deux périmètres ; relancer M lance une nouvelle
analyse et retire l'ancien contrôle M-1.

Exemple Aba : 100 % de complétude, 100 % de promptitude, 9 datasets ≥95 %, 5
violations sur 134 rapports et 2 violations en M-1 donnent
20 + 20 + 15 + 20 + 0 = **75/100**. Le ratio exact (3,73/100) sert au barème.
Même si les 2 anciennes violations sont corrigées en M, la note M-1 reste 0.
Seule une analyse M-1 réussie sans aucune violation donne 20 points.

La sélection **Datasets évalués pour le score qualité** est alimentée par les
métadonnées DHIS2. Elle reprend les datasets du favori, retire le service tertiaire
et ajoute **eSurveillance Hebdomadaire-Edition 2025**, identifié par son nom puis son UID DHIS2.
Si ce nom est absent ou ambigu, un message demande de vérifier la sélection ;
aucun autre dataset de surveillance n'est substitué automatiquement.
Aucun identifiant ni résultat d'exemple n'est imposé.
Les cinq métriques de reporting de chaque dataset choisi alimentent les taux,
les rapports et le compte des datasets à 95 %. Les datasets mensuels et hebdomadaires
font l'objet de requêtes séparées, avec les mêmes filtres d'unités et les mêmes
filtres supplémentaires du favori. eSurveillance est interrogé sur les semaines ISO,
puis rattaché par UID aux indicateurs mensuels. Les rapports réels, attendus et à
temps sont additionnés avant de recalculer les taux globaux ; eSurveillance compte
une seule fois parmi les datasets à 95 %, quel que soit le nombre de semaines.
Ces requêtes utilisent les [métriques de reporting Analytics de DHIS2](https://docs.dhis2.org/en/develop/using-the-api/dhis-core-version-242/analytics.html).

**Périodes et semaines choisies :** la barre latérale propose une période fixe
mensuelle, trimestrielle, semestrielle, annuelle, hebdomadaire ou journalière, ainsi qu'un intervalle
de ces périodes sur une ou plusieurs années. Les trimestres, semestres et années sont
développés en mois entiers pour les datasets mensuels. Les deux bornes affichées sont
**incluses** : un trimestre se termine le dernier jour de son troisième mois, et le
trimestre suivant commence le lendemain. Les dates de validation sont initialisées
sur les bornes exactes demandées et restent modifiables.

Une semaine ISO reste entière, du lundi au dimanche, et est rattachée au **mois de
son dimanche de fin**. W36–W39 de 2026 couvre réellement le 31 août–27 septembre,
mais n'ajoute pas août aux datasets mensuels : seul septembre est agrégé. W40 se
termine le 4 octobre et appartient à octobre. La référence M-1 suit le mois précédant
le premier mois de reporting retenu. Les mois agrégés, les dates demandées et la
couverture réelle des semaines sont affichés et exportés séparément.

Les semaines de surveillance se choisissent librement ou par semaine de début et de
fin. Par défaut, la sélection couvre les semaines dont le dimanche est compris dans
l'intervalle inclusif. L'option « Semaines entièrement comprises » exige que le lundi
et le dimanche soient tous deux dans cet intervalle. La sélection
peut être modifiée et ne comporte aucune limite de quatre ou cinq semaines. Aucun
report automatique de cinquième semaine n'est appliqué. Un semestre peut ainsi utiliser
toutes ses semaines, ou seulement celles que l'utilisateur a retenues.
Une semaine ajoutée manuellement hors des bornes reste visible et signalée ; aucun
rapport hebdomadaire ni mensuel n'est découpé ou calculé au prorata des jours.
L'export ajoute les feuilles `Periodes_Surveillance` et `Reporting_Surveillance`.

Une surveillance indisponible ne devient pas un dataset inactif : le score de l'unité
reste incomplet, tandis que les autres unités restent affichées. La source combinée
mensuelle et hebdomadaire alimente désormais **tous les tableaux** : complétude,
promptitude, analyse comparative et qualité. Les anciennes colonnes de reporting du
favori sont remplacées avant les calculs, afin de ne compter aucun dataset deux fois ;
ses autres indicateurs restent disponibles. La qualité réutilise cette synthèse
seulement pour les mêmes mois, datasets et semaines. Une période qualité indépendante
dispose de son propre choix de semaines. Changer les semaines recalcule les indicateurs
sans relancer les requêtes M/M-1 déjà reçues.

Les deux actions existantes restent séparées :

1. **Lancer l'analyse des violations** interroge uniquement la période filtrée M,
   pour une zone sélectionnée, toutes les zones d'une province ou le pays.
   Le détail, le compteur M, le ratio et la note des violations sont disponibles
   sans requête M-1. Les colonnes M-1, corrections et le total /100 restent vides,
   avec le statut « contrôle des corrections non effectué ».
2. **Contrôler les corrections du mois précédent** devient disponible après une
   analyse M, même partiellement indisponible. Ce bouton interroge M-1 pour les
   mêmes zones et le compare aux résultats M déjà affichés, sans relancer M.
   Il affiche le détail de M-1 et complète les corrections et le total
   si les données nécessaires sont disponibles. Un échec M-1 conserve le détail,
   les compteurs et la note des violations M ; le contrôle peut être réessayé.

**Période des violations affichées** permet de consulter M ou M-1 sans nouvelle
requête. Le détail se filtre par pays, province, zone, importance et recherche.
M-1 indique pour chaque ligne **Corrigée**, **Non corrigée** ou **Non vérifiable**
si M est indisponible pour l'unité. Le CSV contient la période et les filtres
affichés ; l'Excel conserve M et M-1 dans deux feuilles distinctes, avec un aperçu
pour chaque période.

Relancer M efface le contrôle précédent. Changer les dates, le parent ou le groupe
de règles impose une nouvelle analyse M avant un nouveau contrôle.

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
le groupe de règles, les mois des indicateurs, la référence M-1, le barème utilisé,
l'état du contrôle et les horodatages distincts de l'analyse M et du contrôle.
Les lignes conservent l'ordre, les périodes, les valeurs et les libellés renvoyés par DHIS2,
y compris les règles dont la description est vide. Une violation présente sur deux
périodes conserve deux lignes et compte deux fois dans la synthèse. Les signatures
sans période servent uniquement à comparer les corrections entre M-1 et M.

Les analyses M et M-1 sont calculées à la demande par `POST /api/dataAnalysis/validationRules`,
avec un corps JSON contenant `ou`, `startDate`, `endDate`, `persist: false`,
`notification: false` et éventuellement `vrg`, comme l'application officielle DHIS2.
Une seule requête couvre le parent sélectionné et l'intervalle complet ; la requête
M-1, lancée uniquement par le second bouton, utilise le même parent et le mois de référence entier. Les unités descendantes
sont sélectionnées par DHIS2. Les chemins hiérarchiques ne sont pas obligatoires
dans la réponse ; les métadonnées servent à répartir les lignes par zone de santé.
Si ce rattachement manque, les lignes restent affichées et les comptes par zone
restent indisponibles. Un ancien serveur répondant HTTP 404, 405 ou 501 au POST
est interrogé en GET avec les mêmes paramètres. Un refus 401/403 reste une erreur.
Ces analyses ne modifient pas les données DHIS2
et ne déclenchent pas de notifications. Le nombre de règles corrigées correspond aux
signatures de M-1 absentes de l'analyse M affichée.
Chaque clic sur **Lancer l'analyse des violations** effectue une nouvelle requête
DHIS2 pour M. Chaque clic sur le contrôle recharge uniquement M-1. Les résultats restent en session pendant la consultation, sans nouvel appel
lors d'une recherche, d'un tri ou d'un export. **Actualiser les données** recharge
aussi les indicateurs de déclaration et les métadonnées.
L'expandeur **Vérifier la requête envoyée à DHIS2** présente le compte, le parent,
les dates, le groupe et le nombre de lignes reçues pour M, puis M-1 après le contrôle. Les erreurs
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
# DHIS2_ANALYTICS_TIMEOUT_READ = 300
# DHIS2_VALIDATION_TIMEOUT_READ = 900
# DHIS2_HTTP_RETRIES = 2
# DHIS2_AUTH_TIMEOUT_CONNECT = 10
# DHIS2_AUTH_TIMEOUT_READ = 60
# DHIS2_AUTH_RETRIES = 1
```

Le projet lit `DHIS2_URL` via `st.secrets`.
Le **nom d'utilisateur** et le **mot de passe** sont saisis par chaque utilisateur dans la barre latérale (`🔐 Connexion DHIS2`).

L'analyse des règles peut attendre le calcul DHIS2 plus longtemps que les autres
requêtes. `DHIS2_VALIDATION_TIMEOUT_READ` règle son délai de lecture, à 900 secondes
(15 minutes) au minimum. Une ancienne configuration à 600 est ainsi portée à 900.
Sans configuration spécifique, un délai général supérieur à 900 reste conservé. Ce réglage
s'applique uniquement à `dataAnalysis/validationRules`, en POST et en GET de repli.
Les indicateurs utilisent `DHIS2_ANALYTICS_TIMEOUT_READ` (300 secondes par défaut,
ou le délai général s'il est plus grand). Les appels de métadonnées gardent
`DHIS2_TIMEOUT_READ` (90 secondes par défaut).
La vérification des identifiants sur `/api/me` utilise séparément
`DHIS2_AUTH_TIMEOUT_CONNECT` (10 secondes) et `DHIS2_AUTH_TIMEOUT_READ` (60 secondes).
Elle fait au plus une seconde tentative lors d'un délai de lecture dépassé,
d'une interruption réseau ou d'un HTTP 502/503/504. `DHIS2_AUTH_RETRIES = 0`
désactive cette seconde tentative. Les refus 401/403 et les erreurs de certificat
ne sont pas relancés. Le message de connexion distingue lenteur, refus et réseau.
Dans Streamlit Cloud, ces options se règlent dans **App settings → Secrets**.
Le délai de lecture mesure l'attente de données, pas une durée totale garantie.
Une coupure du serveur ou d'un relais peut arriver plus tôt.

En cas de délai dépassé, le diagnostic indique **Indisponible** et laisse le nombre
de lignes vide. Un vrai résultat sans violation indique **Terminé**, avec zéro ligne.
La durée de la requête et son délai de lecture sont affichés dans le diagnostic.
L'application ne relance pas automatiquement un POST qui a dépassé le délai :
DHIS2 peut encore être en train de calculer l'analyse.

Le chargement des indicateurs Analytics est aussi réparti **par unité explicite**,
avec tous les datasets du favori pour cette unité. Le filtre de période et les
filtres supplémentaires restent identiques ; aucun taux n'est moyenné entre lots.
Une province de 13 zones produit ainsi 13 petites requêtes, avec une progression,
au lieu d'une requête unique pour toutes ses zones. Ces appels Analytics ne sont
pas relancés automatiquement après un timeout.

Les réponses Analytics réussies restent en cache une heure, par compte, révision,
périmètre et paramètres. La synthèse de chargement, y compris une erreur ou une
réponse partielle, ne reste en cache que 60 secondes ; **Actualiser les données**
change immédiatement la révision. Une zone échouée conserve des cellules vides,
ses scores ne sont pas calculés sur des données partielles et les autres zones
restent affichées. Les trois boutons de validation restent accessibles même si
les indicateurs Analytics sont indisponibles. Les erreurs réseau sont résumées
sans afficher l'URL technique ni les identifiants de connexion.

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
|- quality_score.py
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
