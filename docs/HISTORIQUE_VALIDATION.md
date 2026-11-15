# Historique des extractions de règles de validation

Le panneau **Historique des règles de validation**, dans l'onglet **Éléments de
catégorisation**, conserve les violations réellement reçues de DHIS2 à une date
d'extraction. Il utilise les dates, le groupe de règles et les zones de l'analyse.
Il ne modifie aucune donnée dans DHIS2.

## Configuration dans Streamlit Cloud

1. Dans GitHub, ouvrir **Settings → Developer settings → Personal access tokens
   → Fine-grained tokens → Generate new token**.
2. Choisir le propriétaire du dépôt, une expiration adaptée, puis **Only select
   repositories** et le dépôt destiné aux archives. Par défaut :
   `InnocentEsomaJr/Extraction_Favori`.
3. Dans **Repository permissions**, autoriser **Contents: Read and write**.
   Les permissions Metadata en lecture sont ajoutées par GitHub. Le propriétaire
   du jeton doit lui-même pouvoir écrire dans le dépôt.
4. Dans l'application Streamlit, ouvrir **Manage app → menu ⋮ → Settings → Secrets**
   et ajouter les clés suivantes, en conservant les secrets DHIS2 existants :

```toml
GITHUB_HISTORY_TOKEN = "REMPLACER_PAR_VOTRE_JETON"
GITHUB_HISTORY_REPOSITORY = "InnocentEsomaJr/Extraction_Favori"
GITHUB_HISTORY_BRANCH = "validation-history"
GITHUB_HISTORY_DIRECTORY = "validation-history"
```

5. Enregistrer les Secrets, puis recharger l'application. Le bouton de sauvegarde
   devient disponible après une analyse M ayant reçu au moins une réponse de zone,
   y compris une réponse vide confirmant l'absence de violations.

Ne placer le vrai jeton ni dans le code, ni dans un commit, ni dans une conversation.
En local, les mêmes clés peuvent être définies dans `.streamlit/secrets.toml`,
qui reste exclu de Git. L'application accepte aussi les variables d'environnement.

La première sauvegarde crée la branche `validation-history` si nécessaire, à partir
de la branche par défaut. Les fichiers d'archive sont ensuite ajoutés à cette
branche ; ils ne sont pas ajoutés à `main` et ne déclenchent pas son redéploiement.
Une règle de protection de branche peut nécessiter une configuration adaptée du
dépôt. Aucun jeton n'est nécessaire pour lire les archives d'un dépôt public.

Les archives contiennent les noms des unités et des règles ainsi que les valeurs
de violation. Elles sont lisibles par les personnes ayant accès au dépôt ; dans
un dépôt public, elles sont publiques. Un dépôt privé distinct peut être choisi
avec `GITHUB_HISTORY_REPOSITORY` pour limiter cette visibilité. Aucun mot de passe,
jeton ni nom de connexion DHIS2 n'est écrit dans le JSON. Des empreintes de
l'instance et du compte permettent de contrôler la compatibilité des extractions.

Références : [créer un jeton à permissions fines](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens),
[permission pour créer un fichier](https://docs.github.com/en/rest/repos/contents#create-or-update-file-contents),
[permission pour créer une branche](https://docs.github.com/en/rest/git/refs#create-a-reference),
[modifier les Secrets Streamlit](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app/app-settings#view-or-update-your-secrets).

## Enregistrer une extraction avant le feedback

1. Sélectionner le pays, la province ou la zone, les dates et le groupe de règles.
2. Cliquer sur **Lancer l'analyse des violations**, ou **Calculer tous les scores
   (M et M-1)**.
3. Ouvrir **Historique des règles de validation**, puis cliquer sur
   **Enregistrer cette extraction dans GitHub**.

Un exemple de nom de fichier est :

```text
septembre_2026-09-01_2026-09-20_<parentUID>__<horodatageUTC>_<identifiant>.json
```

Le préfixe reprend le mois et les dates du filtre ; le parent, l'horodatage et
l'identifiant évitent les collisions entre extractions. Un fichier enregistré
n'est jamais remplacé par l'application. Recliquer pour la même extraction est
sans effet si le contenu est identique. Une nouvelle analyse produit une nouvelle
archive. Chaque JSON garde les lignes reçues et l'état de chaque zone : complet,
en échec, en attente ou tronqué. Une archive partielle ne permet pas de conclure
à la correction des zones incomplètes.

Le bouton **Télécharger cette extraction JSON** fonctionne aussi sans jeton GitHub.
Les archives sont produites à partir des analyses effectuées dans l'application ;
elles ne peuvent pas reconstituer l'état de DHIS2 avant leur enregistrement.
Un fichier est limité à 20 Mio. À 1 000 fichiers dans un dossier, la lecture
signale la limite de l'API GitHub ; choisir un nouveau dossier permet de continuer
les sauvegardes sans supprimer les anciennes archives.

## Contrôler après le feedback

Exemple : une extraction du 1er au 20 septembre est enregistrée le 20 septembre.
Le 30 septembre, le filtre actuel peut couvrir le 20 au 30, ou le mois complet.

1. Choisir le périmètre géographique à examiner dans les filtres de l'analyse.
2. Dans l'historique, cliquer sur **Charger l'historique GitHub**.
3. Choisir le mois et le fichier, puis **Ouvrir cette extraction**.
4. Cliquer sur **Vérifier les corrections de cette extraction**.

Le contrôle interroge les **mêmes dates et le même groupe de règles que le fichier**,
pour les zones communes avec le filtre géographique actuel. Il avance zone par
zone, conserve les réponses reçues et propose **Reprendre le contrôle historique**
après une interruption. Il conserve les résultats actuels M et M-1. Le contrôle
historique et l'analyse principale ne lancent pas de requêtes simultanées.

Le tableau peut être regroupé par **Zones de santé**, **Provinces / DPS** ou
**Pays**. Les filtres actuels limitent les unités comparées : sélectionner une
zone dans une ancienne extraction de province ne compare que cette zone.
Un filtre d'état et un export CSV permettent de préparer le suivi des corrections.

| État | Interprétation |
| --- | --- |
| Toujours violée | La même unité, règle, période et combinaison est encore présente. |
| Corrigée | Elle est absente d'un contrôle complet couvrant l'ancienne période, avec instance, compte et groupe compatibles, unité encore accessible et définition de règle inchangée. |
| Nouvelle | Elle apparaît dans un intervalle déjà couvert par une ancienne extraction complète. |
| Nouvelle dans le filtre actuel | Elle est observée dans le filtre actuel, qui peut couvrir d'autres dates ou zones ; son apparition après l'ancienne extraction n'est pas démontrée. |
| Non vérifiable | La comparaison est incomplète, incompatible ou ne permet pas de confirmer l'absence. |

Comparer seulement le 20–30 à une extraction du 1–20 ne suffit pas à déclarer une
correction. Le contrôle historique relit donc le 1–20. Les métadonnées des unités,
du groupe et des règles sont vérifiées avant de confirmer une disparition ; un
accès invérifiable ou une règle modifiée laisse la ligne non vérifiable. Ces
contrôles ne reconstituent pas tous les changements de droits DHIS2 : une absence
reste un constat fait avec le compte utilisé lors du contrôle.

**Comparer avec → Analyse du filtre actuel** reste possible. Pour confirmer les
disparitions de cette vue, cliquer sur **Vérifier la comparaison avec l'analyse
actuelle**. Si ses dates ne couvrent pas celles de l'archive, les anciennes
violations absentes restent non vérifiables.

Les compteurs de suivi portent sur les identités distinctes de violations ; les
lignes et doublons du serveur restent conservés dans le JSON et le détail. Une
correction observée ne donne ni l'auteur ni la date exacte du changement. Les
contrôles servent à constater l'évolution entre deux extractions, sans modifier
le barème M/M-1 ni les scores qualité existants.

## Résoudre un problème de sauvegarde

- **Bouton désactivé** : vérifier le secret, puis terminer une analyse M ; pendant
  une analyse ou un contrôle, la sauvegarde attend la fin du traitement.
- **Accès refusé** : vérifier l'expiration du jeton, le dépôt choisi, Contents en
  lecture/écriture et les éventuelles règles de l'organisation ou de la branche.
- **Délai réseau dépassé après une sauvegarde** : réessayer le même bouton.
  Si le fichier existe déjà avec le même contenu, il est reconnu sans remplacement.
- **Aucune archive** : vérifier les trois clés dépôt/branche/dossier. La branche et
  le dossier ne sont créés que lors de la première sauvegarde réussie.
- **Non vérifiable** : consulter les précisions de comparaison et l'état complet
  de chaque extraction. Une panne ou une restriction DHIS2 reste une indisponibilité.
