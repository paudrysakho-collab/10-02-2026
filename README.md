# Bases CSV cavistes par département

Génère une base CSV par département à partir de la base de données et du
fichier d'inscriptions.

## Fichiers produits

`sorties/cavistes_<dep>.csv` pour 85, 56, 49, 35, 44, 53 — 4 colonnes :

| prénom | mail | département | lien |
|--------|------|-------------|------|

plus `sorties/_rejets.csv` qui liste tout ce qui a été écarté, avec le motif
(à relire avant envoi).

## Liens catalogue par département

| Dép. | Lien |
|------|------|
| 85 | https://drive.google.com/file/d/1liiQUXAHnWHR_YJdzUNFBg4mQhPLXD1v/view?usp=drive_link |
| 56 | https://drive.google.com/file/d/1b1UhH1wXcGe9clGrCNjk1sdyhFzFf6Aj/view?usp=drive_link |
| 49 | https://drive.google.com/file/d/1R6w7FRgzPpGDzVqzo8ei7JV5A2TCNOxP/view?usp=drive_link |
| 35 | https://drive.google.com/file/d/1fIM1KdH9tLhAYone9MvVKuXof_Ee4IpF/view?usp=drive_link |
| 44 | https://drive.google.com/file/d/15IV5mw8lBlADC-PktLD29ICK9JCEiSmm/view?usp=drive_link |
| 53 | https://drive.google.com/file/d/12eKNUJ32WDHl4ZbJUWjW44NbO9Nr_qxR/view?usp=drive_link |

## Utilisation

```bash
pip install openpyxl
python3 build_csv.py base.xlsx inscriptions.xlsx -o sorties/
```

## Règles appliquées

- lignes **surlignées en rouge** dans la base de données : exclues
  (rouge vif comme rouge pâle, sur n'importe quelle cellule de la ligne)
- inscriptions : ajoutées
- département : lu dans la colonne dédiée, sinon déduit du code postal
- doublons de mail : supprimés (première occurrence conservée)
- nettoyage des mails : syntaxe invalide, domaine inexistant (vérif. MX puis A)
  et domaines jetables écartés → `_rejets.csv`
- les contacts hors des six départements sont écartés et listés dans `_rejets.csv`
