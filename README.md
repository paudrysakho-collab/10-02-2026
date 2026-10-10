# Bases CSV d'envoi — catalogue Journée Dégustation

Génère les fichiers d'envoi par département à partir du classeur de prospection,
des inscriptions et de la liste des désabonnés.

## Utilisation

```bash
python3 build_csv.py                 # cavistes uniquement (périmètre actuel)
python3 build_csv.py --tous          # toutes les catégories du classeur
python3 build_csv.py --onglets Cavistes "Bars à vin"
```

Dépendance : `pip install openpyxl`.

## Sorties (`sorties/`)

| Fichier | Contenu |
|---|---|
| `envoi_85.csv` … `envoi_53.csv` | un fichier par département |
| `envoi_tous_departements.csv` | **tout le monde dans un seul fichier** — c'est celui à charger pour envoyer d'un coup |
| `envoi_tous_departements.xlsx` | même contenu, onglet global + un onglet par département |
| `_rejets.csv` | tout ce qui est écarté, avec le motif |

Quatre colonnes : `prénom`, `mail`, `département`, `lien`. Le lien du catalogue
est déduit du département, donc rien à modifier entre deux envois.

Encodage UTF-8 avec BOM : Excel ouvre les accents correctement.

## Liens catalogue par département

| Dép. | Lien |
|------|------|
| 85 | https://drive.google.com/file/d/1liiQUXAHnWHR_YJdzUNFBg4mQhPLXD1v/view?usp=drive_link |
| 56 | https://drive.google.com/file/d/1b1UhH1wXcGe9clGrCNjk1sdyhFzFf6Aj/view?usp=drive_link |
| 49 | https://drive.google.com/file/d/1R6w7FRgzPpGDzVqzo8ei7JV5A2TCNOxP/view?usp=drive_link |
| 35 | https://drive.google.com/file/d/1fIM1KdH9tLhAYone9MvVKuXof_Ee4IpF/view?usp=drive_link |
| 44 | https://drive.google.com/file/d/15IV5mw8lBlADC-PktLD29ICK9JCEiSmm/view?usp=drive_link |
| 53 | https://drive.google.com/file/d/12eKNUJ32WDHl4ZbJUWjW44NbO9Nr_qxR/view?usp=drive_link |

## Règles appliquées

**Exclusions**
- lignes dont le nom est écrit en **police rouge** dans le classeur → écartées.
  Attention : c'est bien la couleur du texte. Le fond rose pâle est du zébrage
  (269 lignes sur 270 dans l'onglet Cavistes) et le fond bordeaux est l'en-tête.
  Une note négative (« Fermé définitivement », « pas intéressé ») **n'est pas**
  un motif d'exclusion.
- adresses listées dans `desabonnements.csv`
- contacts hors des six départements
- inscrits vignerons (`Invitation = vigneron donc pas besoin`)

**Nettoyage**
- doublons de mail supprimés (première occurrence conservée)
- syntaxe invalide, valeurs parasites (`#ERROR!`, formules), domaines jetables
- existence du domaine vérifiée (MX puis A) par DNS-over-HTTPS, le DNS direct
  étant filtré dans le conteneur

**Colonne prénom**
1. formule déjà validée dans `donnees/modele_relance_lot4.csv`, reprise mot pour mot
2. prénom donné par l'inscrit
3. prénom saisi à la main dans le classeur — un seul mot ou « X et Y ». La colonne
   `Gérant(s)` (et la colonne `Prénom` de l'onglet Restaurants) contient des noms
   d'officiers tirés du registre des entreprises : jamais utilisée.
4. à défaut, génération « à toute l'équipe de/du/de la/des/de l' <établissement> »

## `desabonnements.csv`

Liste de blocage à enrichir après chaque campagne.

| colonne | rôle |
|---|---|
| `mail` | adresse concernée |
| `motif` | ce qui a été demandé |
| `date` | date du message |
| `remplace_par` | nouvelle adresse en cas de changement — le contact est alors **conservé**, avec l'adresse corrigée |

Source : dossier « RETIRER MAILLING » de la boîte `contact@agencescio.com`.
Cette boîte n'est pas le compte connecté à la session, la liste est donc tenue
à la main pour l'instant.
