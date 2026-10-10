#!/usr/bin/env python3
"""Construit les bases CSV par departement a partir de la base de donnees
et du fichier d'inscriptions.

Sortie : un CSV par departement (85, 56, 49, 35, 44, 53) avec 4 colonnes
prenom, mail, departement, lien

Regles :
  - les lignes surlignees en rouge dans la base de donnees sont exclues
  - les inscriptions sont ajoutees
  - doublons de mail supprimes (le premier vu gagne)
  - mails invalides (syntaxe) ou dont le domaine n'existe pas : exclus,
    et listes dans _rejets.csv pour controle

Usage :
  python3 build_csv.py base.xlsx inscriptions.xlsx -o sorties/
"""

import argparse
import csv
import json
import os
import re
import subprocess
import sys
import unicodedata
from collections import OrderedDict

from openpyxl import load_workbook

# ---------------------------------------------------------------- liens
# Ordre fourni par le client : 85, 56, 49, 35, 44, 53
LIENS = OrderedDict([
    ("85", "https://drive.google.com/file/d/1liiQUXAHnWHR_YJdzUNFBg4mQhPLXD1v/view?usp=drive_link"),
    ("56", "https://drive.google.com/file/d/1b1UhH1wXcGe9clGrCNjk1sdyhFzFf6Aj/view?usp=drive_link"),
    ("49", "https://drive.google.com/file/d/1R6w7FRgzPpGDzVqzo8ei7JV5A2TCNOxP/view?usp=drive_link"),
    ("35", "https://drive.google.com/file/d/1fIM1KdH9tLhAYone9MvVKuXof_Ee4IpF/view?usp=drive_link"),
    ("44", "https://drive.google.com/file/d/15IV5mw8lBlADC-PktLD29ICK9JCEiSmm/view?usp=drive_link"),
    ("53", "https://drive.google.com/file/d/12eKNUJ32WDHl4ZbJUWjW44NbO9Nr_qxR/view?usp=drive_link"),
])

DEPARTEMENTS = list(LIENS)

# ------------------------------------------------------- detection colonnes
ALIAS = {
    "prenom": ("prenom", "prenoms", "firstname", "first name", "contact", "nom du contact"),
    "nom": ("nom", "nom de famille", "lastname", "last name", "enseigne", "societe",
            "raison sociale", "nom caviste", "cave", "magasin"),
    "mail": ("mail", "mails", "email", "e mail", "e-mail", "adresse mail",
             "adresse email", "courriel"),
    "departement": ("departement", "dept", "dep", "departement n", "no departement"),
    "cp": ("cp", "code postal", "codepostal", "postal", "zip"),
    "ville": ("ville", "commune", "localite", "city"),
    "adresse": ("adresse", "adresse postale", "address", "rue"),
}


def strip_accents(txt):
    txt = unicodedata.normalize("NFKD", str(txt))
    return "".join(c for c in txt if not unicodedata.combining(c))


def norm(txt):
    if txt is None:
        return ""
    txt = strip_accents(txt).lower().strip()
    return re.sub(r"[^a-z0-9 ]+", " ", txt).strip()


def map_headers(row):
    """row : liste de valeurs d'en-tete -> {role: index}"""
    found = {}
    for idx, cell in enumerate(row):
        key = norm(cell)
        if not key:
            continue
        key = re.sub(r"\s+", " ", key)
        for role, aliases in ALIAS.items():
            if role in found:
                continue
            if key in aliases or any(key.startswith(a + " ") for a in aliases):
                found[role] = idx
    return found


def find_header_row(ws, max_scan=15):
    """Cherche la premiere ligne qui ressemble a un en-tete (contient un mail)."""
    best = (None, {})
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        mapping = map_headers(row)
        if "mail" in mapping and len(mapping) >= len(best[1]):
            best = (r, mapping)
            if "prenom" in mapping or "nom" in mapping:
                break
    return best


# -------------------------------------------------------- surlignage rouge
def is_reddish(hex_argb):
    if not hex_argb or not isinstance(hex_argb, str):
        return False
    h = hex_argb[-6:]
    if len(h) != 6:
        return False
    try:
        r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return False
    # rouge franc ou rouge pale (type surlignage Excel / Sheets)
    return r >= 120 and r - max(g, b) >= 40


def row_is_red(cells):
    """Vrai si au moins une cellule de la ligne est surlignee en rouge."""
    for c in cells:
        fill = c.fill
        if fill is None or fill.patternType in (None, "none"):
            continue
        for color in (fill.fgColor, fill.start_color):
            if color is None:
                continue
            if color.type == "rgb" and is_reddish(color.rgb):
                return True
            if color.type == "indexed" and color.indexed in (10, 3):  # rouge palette
                return True
            if color.type == "theme" and getattr(color, "theme", None) == 10:
                return True
    return False


# ---------------------------------------------------------- lecture fichier
def read_sheet(path, skip_red):
    wb = load_workbook(path, data_only=True)
    out, red_skipped = [], 0
    for ws in wb.worksheets:
        header_row, mapping = find_header_row(ws)
        if not header_row:
            print(f"  [!] {os.path.basename(path)} / onglet '{ws.title}' : "
                  f"pas de colonne mail trouvee, onglet ignore", file=sys.stderr)
            continue
        print(f"  onglet '{ws.title}' : en-tete ligne {header_row}, "
              f"colonnes {sorted(mapping)}")
        for cells in ws.iter_rows(min_row=header_row + 1):
            if skip_red and row_is_red(cells):
                red_skipped += 1
                continue
            vals = [c.value for c in cells]

            def get(role):
                i = mapping.get(role)
                if i is None or i >= len(vals) or vals[i] is None:
                    return ""
                return str(vals[i]).strip()

            rec = {role: get(role) for role in ALIAS}
            if not rec["mail"]:
                continue
            rec["_source"] = f"{os.path.basename(path)}:{ws.title}"
            out.append(rec)
    return out, red_skipped


# ------------------------------------------------------------- departement
def departement_of(rec):
    dep = re.sub(r"\D", "", rec.get("departement", ""))
    if len(dep) == 1:
        dep = "0" + dep
    if dep[:2] in LIENS:
        return dep[:2]
    for champ in ("cp", "adresse", "ville"):
        m = re.search(r"\b(\d{5})\b", rec.get(champ, ""))
        if m and m.group(1)[:2] in LIENS:
            return m.group(1)[:2]
    return None


# ------------------------------------------------------------ mails
MAIL_RE = re.compile(r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+@([A-Za-z0-9-]+\.)+[A-Za-z]{2,}$")

JETABLES = {
    "yopmail.com", "mailinator.com", "guerrillamail.com", "10minutemail.com",
    "trashmail.com", "tempmail.com", "jetable.org", "yopmail.fr",
}

_dns_cache = {}


def domaine_existe(domaine):
    """MX, sinon A : via DNS-over-HTTPS (le DNS direct est filtre ici)."""
    if domaine in _dns_cache:
        return _dns_cache[domaine]
    ok = False
    for rtype in ("MX", "A"):
        try:
            res = subprocess.run(
                ["curl", "-s", "--max-time", "12",
                 f"https://dns.google/resolve?name={domaine}&type={rtype}"],
                capture_output=True, text=True, timeout=20)
            data = json.loads(res.stdout or "{}")
        except Exception:
            data = {}
        if data.get("Status") == 0 and data.get("Answer"):
            ok = True
            break
        if data.get("Status") == 3:  # NXDOMAIN : inutile de tester le A
            if rtype == "MX" and data.get("Authority"):
                continue
            ok = False
            break
    _dns_cache[domaine] = ok
    return ok


def nettoyer_mail(brut):
    """-> (mail_propre, motif_de_rejet | None)"""
    mail = str(brut).strip().strip("<>,;").replace(" ", "")
    mail = mail.split(",")[0].split(";")[0]
    if not mail:
        return None, "vide"
    mail = mail.lower()
    if not MAIL_RE.match(mail):
        return None, "syntaxe invalide"
    domaine = mail.split("@")[1]
    if domaine in JETABLES:
        return None, "domaine jetable"
    if not domaine_existe(domaine):
        return None, f"domaine inexistant ({domaine})"
    return mail, None


def prenom_de(rec):
    p = rec.get("prenom", "").strip()
    if p:
        return p.split()[0].capitalize() if p.isupper() else p
    return ""


# ------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("base", help="base de donnees (.xlsx) - lignes rouges exclues")
    ap.add_argument("inscriptions", nargs="?", help="fichier des inscriptions (.xlsx)")
    ap.add_argument("-o", "--out", default="sorties", help="dossier de sortie")
    ap.add_argument("--no-dns", action="store_true",
                    help="ne pas verifier l'existence des domaines")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)

    print(f"Lecture base : {args.base}")
    lignes, red = read_sheet(args.base, skip_red=True)
    print(f"  {len(lignes)} lignes retenues, {red} lignes rouges exclues")

    if args.inscriptions:
        print(f"Lecture inscriptions : {args.inscriptions}")
        insc, _ = read_sheet(args.inscriptions, skip_red=False)
        print(f"  {len(insc)} lignes")
        lignes += insc

    if args.no_dns:
        global domaine_existe
        domaine_existe = lambda d: True  # noqa: E731

    par_dep = {d: [] for d in DEPARTEMENTS}
    rejets, vus = [], set()

    for rec in lignes:
        mail, motif = nettoyer_mail(rec["mail"])
        if motif:
            rejets.append([rec["mail"], rec.get("prenom", ""), rec.get("nom", ""), motif, rec["_source"]])
            continue
        if mail in vus:
            continue
        dep = departement_of(rec)
        if dep is None:
            rejets.append([mail, rec.get("prenom", ""), rec.get("nom", ""),
                           "departement hors cible ou absent", rec["_source"]])
            continue
        vus.add(mail)
        par_dep[dep].append([prenom_de(rec), mail, dep, LIENS[dep]])

    total = 0
    for dep in DEPARTEMENTS:
        chemin = os.path.join(args.out, f"cavistes_{dep}.csv")
        with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["prénom", "mail", "département", "lien"])
            w.writerows(par_dep[dep])
        total += len(par_dep[dep])
        print(f"{chemin} : {len(par_dep[dep])} contacts")

    chemin = os.path.join(args.out, "_rejets.csv")
    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["mail", "prénom", "nom", "motif", "source"])
        w.writerows(rejets)
    print(f"{chemin} : {len(rejets)} lignes ecartees")
    print(f"\nTotal : {total} contacts dans {len(DEPARTEMENTS)} fichiers")


if __name__ == "__main__":
    main()
