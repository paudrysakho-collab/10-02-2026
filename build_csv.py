#!/usr/bin/env python3
"""Construit les bases CSV d'envoi du catalogue, par département.

Sources : le classeur de prospection, le fichier d'inscriptions, l'ancien CSV
d'envoi (modèle des formules d'appel) et la liste des désabonnés.

Sortie : 4 colonnes  prénom, mail, département, lien
  - un CSV par département (85, 56, 49, 35, 44, 53)
  - un CSV global contenant tout le monde
  - le même contenu en .xlsx (un onglet global + un onglet par département)
  - _rejets.csv : tout ce qui est écarté, avec le motif

Usage :
  python3 build_csv.py                      # cavistes uniquement (défaut)
  python3 build_csv.py --tous               # toutes les catégories
  python3 build_csv.py --onglets Cavistes "Bars à vin"
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
from concurrent.futures import ThreadPoolExecutor

from openpyxl import Workbook, load_workbook

# --------------------------------------------------------------------- liens
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

ROUGE = "FFFF0000"          # police rouge = ligne à ne pas contacter
ONGLETS_DEFAUT = ["Cavistes"]

MAIL_RE = re.compile(r"[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}")
MAIL_STRICT = re.compile(r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}$")
DEPT_PAREN = re.compile(r"\((\d{2})\)")
CP_RE = re.compile(r"\b(\d{5})\b")

JETABLES = {"yopmail.com", "yopmail.fr", "mailinator.com", "guerrillamail.com",
            "10minutemail.com", "trashmail.com", "tempmail.com", "jetable.org"}


# ------------------------------------------------------------------ utilitaires
def strip_accents(txt):
    return "".join(c for c in unicodedata.normalize("NFKD", str(txt))
                   if not unicodedata.combining(c))


def norm(txt):
    if txt is None:
        return ""
    return re.sub(r"[^a-z0-9 ]+", " ", strip_accents(txt).lower()).strip()


def clean_txt(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


# ------------------------------------------------- détection des colonnes
ALIAS = {
    "nom": ("nom", "enseigne", "etablissement", "raison sociale"),
    "mail": ("email", "e mail", "mail", "adresse mail", "courriel"),
    "mail2": ("email 2", "e mail 2", "mail 2"),
    "dept": ("dept", "departement"),
    "cp": ("cp", "code postal"),
    "ville": ("ville", "commune"),
}


def map_headers(row):
    found = {}
    for idx, cell in enumerate(row):
        key = re.sub(r"\s+", " ", norm(cell))
        if not key:
            continue
        for role, aliases in ALIAS.items():
            if role not in found and key in aliases:
                found[role] = idx
    return found


def find_header_row(ws, max_scan=6):
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        m = map_headers(row)
        if "mail" in m and "nom" in m:
            return r, m
    return None, {}


# ------------------------------------------------------------ police rouge
def ligne_rouge(cells):
    for c in cells:
        col = c.font.color if c.font else None
        if col is not None and col.type == "rgb" and col.rgb == ROUGE and c.value is not None:
            return True
    return False


# ------------------------------------------------- prénom saisi à la main
MOTS_NOTE = {
    "envoyer", "liste", "message", "messagerie", "laisser", "laisse", "transmet",
    "patrons", "pas", "veut", "catalogue", "relancer", "rappeller", "appeller",
    "interresse", "interesse", "ferme", "definitivement", "fournisseur",
    "contrat", "exclusif", "deja", "grossiste", "redirige", "disponible",
    "regarder", "moyen", "deplacement", "effectif", "dispo", "mail", "envoye",
    "salon", "appel", "commercial", "marche", "fruit", "none",
}


def prenom_manuel(valeur, nom_etab=""):
    """N'accepte qu'un vrai prénom : un seul mot, ou « X et Y »."""
    txt = clean_txt(valeur)
    if not txt or "@" in txt or any(ch.isdigit() for ch in txt) or "&" in txt:
        return ""
    if not (2 < len(txt) < 26):
        return ""
    parties = re.split(r"\s+et\s+", txt, flags=re.I)
    if not 1 <= len(parties) <= 2:
        return ""
    nom_norm = norm(nom_etab)
    for p in parties:
        p = p.strip()
        if not re.fullmatch(r"[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'-]{1,19}", p):
            return ""
        if norm(p) in MOTS_NOTE or norm(p) in nom_norm:
            return ""
    return txt


# ------------------------------------------------------- formule d'appel
PETITS_MOTS = {"de", "du", "des", "la", "le", "les", "à", "a", "au", "aux", "et",
               "en", "sur", "d", "l"}
MASCULIN = {"chateau", "domaine", "manoir", "logis", "clos", "moulin", "relais",
            "pressoir", "chai", "cellier", "comptoir", "cafe", "bar", "restaurant",
            "bistrot", "caveau", "palais", "petit", "grand", "vieux", "verger",
            "marche", "pub", "salon", "coin", "jardin", "puits", "puy", "bistro"}
FEMININ = {"cave", "maison", "table", "ferme", "brasserie", "auberge", "villa",
           "tour", "cour", "terrasse", "creperie", "vinotheque", "boutique",
           "cabane", "halle", "grange", "bodega", "distillerie", "taverne"}
PLURIEL = {"les", "des", "deux", "trois", "quatre"}


def titre(nom):
    """Remet en casse normale un nom écrit tout en majuscules."""
    lettres = [c for c in nom if c.isalpha()]
    if lettres and sum(c.isupper() for c in lettres) / len(lettres) > 0.7:
        mots = []
        for i, mot in enumerate(nom.split()):
            bas = mot.lower()
            mots.append(bas if i and bas in PETITS_MOTS else bas.capitalize())
        return " ".join(mots)
    return nom


# apostrophe oubliée dans la base : « Latelier du Vin » pour « L'atelier du Vin »
ELIDES_SANS_APOSTROPHE = {
    "latelier", "lescale", "lentrepot", "lepicerie", "linstant", "loenotheque",
    "lunivers", "lalambic", "lardoise", "lessentiel", "lannexe", "letape",
    "letiquette", "lembuscade", "larome", "lambassade", "lauberge", "lecurie",
}


def maj1(txt):
    return txt[:1].upper() + txt[1:] if txt else txt


def formule(nom_etab):
    """« à toute l'équipe du Château de la Rairie », avec les élisions."""
    nom = titre(clean_txt(nom_etab).strip(" -–,;"))
    nom = re.sub(r"^\(sans nom\)$", "", nom, flags=re.I)
    if not nom:
        return "à toute l'équipe"
    base = "à toute l'équipe"
    bas = nom.lower()
    if bas.startswith("les "):
        return f"{base} des {maj1(nom[4:])}"
    if bas.startswith("le "):
        return f"{base} du {maj1(nom[3:])}"
    if bas.startswith("la "):
        return f"{base} de la {maj1(nom[3:])}"
    if bas.startswith(("l'", "l’")):
        return f"{base} de l'{maj1(nom[2:])}"
    mots = nom.split()
    if norm(mots[0]) in ELIDES_SANS_APOSTROPHE:
        reste = " ".join([maj1(mots[0][1:])] + mots[1:])
        return f"{base} de l'{reste}"
    premier = norm(nom.split()[0])
    if not premier:
        return f"{base} de {nom}"
    if premier in PLURIEL or nom[0].isdigit():
        return f"{base} des {nom}" if premier in PLURIEL else f"{base} de {nom}"
    if premier[0] in "aeiouy" or (premier[0] == "h" and len(premier) > 1):
        return f"{base} de l'{nom}"
    if premier in MASCULIN:
        return f"{base} du {nom}"
    if premier in FEMININ:
        return f"{base} de la {nom}"
    return f"{base} de {nom}"


# ------------------------------------------------------------- mails
_dns_cache = {}


def domaine_vivant(domaine):
    if domaine in _dns_cache:
        return _dns_cache[domaine]
    ok = False
    for rtype in ("MX", "A"):
        try:
            res = subprocess.run(
                ["curl", "-s", "--max-time", "12",
                 f"https://dns.google/resolve?name={domaine}&type={rtype}"],
                capture_output=True, text=True, timeout=25)
            data = json.loads(res.stdout or "{}")
        except Exception:
            data = {}
        if data.get("Status") == 0 and data.get("Answer"):
            ok = True
            break
        if data.get("Status") == 3 and not data.get("Authority"):
            break
    _dns_cache[domaine] = ok
    return ok


def normaliser_mail(brut):
    """-> (mail, motif_de_rejet | None)"""
    txt = clean_txt(brut)
    if not txt or "#ERROR" in txt.upper() or txt.startswith("="):
        return None, "valeur parasite"
    # on cherche d'abord tel quel : l'espace sert de délimiteur naturel et
    # évite de coller le mail au texte qui le suit dans la même cellule
    trouve = MAIL_RE.search(txt) or MAIL_RE.search(txt.replace(" ", ""))
    if not trouve:
        return None, "syntaxe invalide"
    mail = trouve.group(0).lower().strip(".,;")
    if not MAIL_STRICT.match(mail):
        return None, "syntaxe invalide"
    if mail.split("@")[1] in JETABLES:
        return None, "domaine jetable"
    return mail, None


# --------------------------------------------------------- lecture classeur
def lire_classeur(chemin, onglets):
    """-> (contacts, rejets)"""
    wb = load_workbook(chemin)
    contacts, rejets = [], []
    for ws in wb.worksheets:
        if ws.title == "Synthèse" or (onglets and ws.title not in onglets):
            continue
        hdr, m = find_header_row(ws)
        premiere = (hdr + 1) if hdr else 1
        for cells in ws.iter_rows(min_row=premiere):
            vals = [c.value for c in cells]
            if not any(v not in (None, "") for v in vals):
                continue
            textes = [clean_txt(v) for v in vals if v is not None]
            texte = " ".join(textes)
            # chaque cellule est examinée séparément : concaténer la ligne
            # collerait le mail au contenu de la cellule suivante
            mails_ligne = [m for t in textes for m in MAIL_RE.findall(t)]
            if not mails_ligne:
                continue

            def col(role):
                i = m.get(role)
                return clean_txt(vals[i]) if i is not None and i < len(vals) else ""

            nom = col("nom") or clean_txt(vals[0])
            mails = [col("mail"), col("mail2")] if m else []
            mails = [x for x in mails if x]
            if not mails:
                mails = mails_ligne

            dep = ""
            md = DEPT_PAREN.search(col("dept") or texte)
            if md:
                dep = md.group(1)
            else:
                brut = re.sub(r"\D", "", col("dept"))
                if len(brut) >= 2 and brut[:2] in LIENS:
                    dep = brut[:2]
                else:
                    mc = CP_RE.search(col("cp") or texte)
                    if mc:
                        dep = mc.group(1)[:2]

            # prénom saisi à la main : uniquement sur l'onglet Cavistes,
            # jamais la colonne Gérant(s) qui contient des noms de registre
            prenom = ""
            if not m:
                for idx in (6, 7):
                    if idx < len(vals):
                        prenom = prenom_manuel(vals[idx], nom)
                        if prenom:
                            break

            rouge = ligne_rouge(cells)
            for brut in mails:
                contacts.append({"nom": nom, "mail_brut": brut, "dept": dep,
                                 "prenom": prenom, "rouge": rouge,
                                 "source": f"{os.path.basename(chemin)}:{ws.title}"})
    return contacts, rejets


def lire_inscriptions(chemin, activites=None):
    wb = load_workbook(chemin, data_only=True)
    ws = wb.active
    hdr = [norm(c.value) for c in ws[1]]

    def idx(*frags):
        for i, h in enumerate(hdr):
            if all(f in h for f in frags):
                return i
        return None

    i_prenom = idx("prenom")
    i_mail1 = idx("adresse", "mail", "souhaitez")
    i_mail2 = idx("adresse", "mail")
    i_cp = idx("code postal")
    i_act = idx("activite")
    i_inv = idx("invitation")
    i_soc = idx("societe")

    contacts, ignores = [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i):
            return clean_txt(row[i]) if i is not None and i < len(row) else ""

        mail = g(i_mail1) or g(i_mail2)
        if not mail:
            continue
        inv = g(i_inv).lower()
        act = g(i_act)
        if "vigneron" in inv:
            ignores.append([mail, g(i_prenom), g(i_soc), "vigneron, pas besoin du catalogue",
                            "inscriptions"])
            continue
        if activites and norm(act) not in {norm(a) for a in activites}:
            ignores.append([mail, g(i_prenom), g(i_soc),
                            f"activité hors périmètre ({act or 'non renseignée'})", "inscriptions"])
            continue
        dep = ""
        mc = CP_RE.search(re.sub(r"\.0$", "", g(i_cp)))
        if mc:
            dep = mc.group(1)[:2]
        contacts.append({"nom": g(i_soc), "mail_brut": mail, "dept": dep,
                         "prenom": g(i_prenom), "rouge": False, "source": "inscriptions"})
    return contacts, ignores


def lire_modele(chemin):
    ref = {}
    if not chemin or not os.path.exists(chemin):
        return ref
    with open(chemin, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            mail = clean_txt(r.get("Email")).lower()
            prenom = clean_txt(r.get("Prenom"))
            if mail and prenom:
                ref[mail] = prenom
    return ref


def lire_desabonnements(chemin):
    bloques, remplace = {}, {}
    if not os.path.exists(chemin):
        return bloques, remplace
    with open(chemin, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            mail = clean_txt(r.get("mail")).lower()
            if not mail:
                continue
            nouveau = clean_txt(r.get("remplace_par")).lower()
            if nouveau:
                remplace[mail] = nouveau
            else:
                bloques[mail] = clean_txt(r.get("motif")) or "désabonné"
    return bloques, remplace


# ------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default="donnees/base.xlsx")
    ap.add_argument("--inscriptions", default="donnees/inscriptions.xlsx")
    ap.add_argument("--modele", default="donnees/modele_relance_lot4.csv")
    ap.add_argument("--desabos", default="desabonnements.csv")
    ap.add_argument("-o", "--out", default="sorties")
    ap.add_argument("--onglets", nargs="*", default=None,
                    help="onglets du classeur à traiter (défaut : Cavistes)")
    ap.add_argument("--tous", action="store_true", help="tous les onglets et toutes les activités")
    ap.add_argument("--no-dns", action="store_true", help="ne pas vérifier les domaines")
    args = ap.parse_args()

    onglets = None if args.tous else (args.onglets or ONGLETS_DEFAUT)
    activites = None if args.tous else ["Caviste"]
    os.makedirs(args.out, exist_ok=True)

    modele = lire_modele(args.modele)
    bloques, remplace = lire_desabonnements(args.desabos)
    print(f"Modèle : {len(modele)} formules d'appel reprises")
    print(f"Désabonnés : {len(bloques)} bloqués, {len(remplace)} adresse(s) corrigée(s)")
    print(f"Périmètre : {'tout le classeur' if onglets is None else ', '.join(onglets)}"
          f" | inscriptions : {'toutes activités' if activites is None else ', '.join(activites)}")

    base, _ = lire_classeur(args.base, onglets)
    insc, ignores = lire_inscriptions(args.inscriptions, activites)
    print(f"Lignes lues : base={len(base)}  inscriptions={len(insc)}")

    rejets = [list(x) for x in ignores]
    etape1 = []
    for c in base + insc:
        if c["rouge"]:
            rejets.append([c["mail_brut"], c["prenom"], c["nom"], "nom en rouge", c["source"]])
            continue
        mail, motif = normaliser_mail(c["mail_brut"])
        if motif:
            rejets.append([c["mail_brut"], c["prenom"], c["nom"], motif, c["source"]])
            continue
        mail = remplace.get(mail, mail)
        if mail in bloques:
            rejets.append([mail, c["prenom"], c["nom"], bloques[mail], c["source"]])
            continue
        if c["dept"] not in LIENS:
            rejets.append([mail, c["prenom"], c["nom"],
                           f"hors zone ({c['dept'] or 'département absent'})", c["source"]])
            continue
        c["mail"] = mail
        etape1.append(c)

    # vérification des domaines
    if not args.no_dns:
        domaines = sorted({c["mail"].split("@")[1] for c in etape1})
        print(f"Vérification de {len(domaines)} domaines (MX/A)...", flush=True)
        with ThreadPoolExecutor(max_workers=16) as ex:
            list(ex.map(domaine_vivant, domaines))
        morts = {d for d in domaines if not _dns_cache.get(d)}
        print(f"  {len(morts)} domaine(s) injoignable(s)")
    else:
        morts = set()

    par_dep = {d: [] for d in DEPARTEMENTS}
    vus = {}
    for c in etape1:
        mail = c["mail"]
        if mail.split("@")[1] in morts:
            rejets.append([mail, c["prenom"], c["nom"],
                           f"domaine inexistant ({mail.split('@')[1]})", c["source"]])
            continue
        prenom = modele.get(mail) or c["prenom"] or formule(c["nom"])
        if prenom.isupper():          # « BENOIT » saisi en majuscules
            prenom = " ".join(p.capitalize() for p in prenom.split())
        if mail in vus:
            # une occurrence plus précise (vrai prénom) améliore la formule
            if not vus[mail][0].lower().startswith("à toute") and \
               prenom.lower().startswith("à toute"):
                pass
            elif vus[mail][0].lower().startswith("à toute") and \
                    not prenom.lower().startswith("à toute"):
                vus[mail][0] = prenom
            rejets.append([mail, prenom, c["nom"], "doublon", c["source"]])
            continue
        ligne = [prenom, mail, c["dept"], LIENS[c["dept"]]]
        vus[mail] = ligne
        par_dep[c["dept"]].append(ligne)

    # ------------------------------------------------------------- écriture
    def ecrire(chemin, lignes):
        with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["prénom", "mail", "département", "lien"])
            w.writerows(lignes)

    total = 0
    print()
    for dep in DEPARTEMENTS:
        chemin = os.path.join(args.out, f"envoi_{dep}.csv")
        ecrire(chemin, par_dep[dep])
        total += len(par_dep[dep])
        print(f"  {chemin:36s} {len(par_dep[dep]):5d} contacts")

    toutes = [l for dep in DEPARTEMENTS for l in par_dep[dep]]
    global_csv = os.path.join(args.out, "envoi_tous_departements.csv")
    ecrire(global_csv, toutes)
    print(f"  {global_csv:36s} {len(toutes):5d} contacts")

    wb = Workbook()
    ws = wb.active
    ws.title = "Tous"
    ws.append(["prénom", "mail", "département", "lien"])
    for l in toutes:
        ws.append(l)
    for dep in DEPARTEMENTS:
        wsd = wb.create_sheet(f"Dept {dep}")
        wsd.append(["prénom", "mail", "département", "lien"])
        for l in par_dep[dep]:
            wsd.append(l)
    for sheet in wb.worksheets:
        sheet.column_dimensions["A"].width = 46
        sheet.column_dimensions["B"].width = 38
        sheet.column_dimensions["C"].width = 13
        sheet.column_dimensions["D"].width = 70
        sheet.freeze_panes = "A2"
    xlsx = os.path.join(args.out, "envoi_tous_departements.xlsx")
    wb.save(xlsx)
    print(f"  {xlsx:36s} {len(toutes):5d} contacts")

    chemin = os.path.join(args.out, "_rejets.csv")
    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["mail", "prénom", "établissement", "motif", "source"])
        w.writerows(rejets)
    print(f"  {chemin:36s} {len(rejets):5d} lignes écartées")

    from collections import Counter
    print("\nMotifs d'écartement :")
    for motif, n in Counter(r[3].split(" (")[0] for r in rejets).most_common():
        print(f"  {n:5d}  {motif}")
    print(f"\nTOTAL : {total} contacts")


if __name__ == "__main__":
    main()
