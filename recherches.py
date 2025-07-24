import json
import re
import os
from unidecode import unidecode


# Chemins d'accès aux fichiers
chemin_fichier = r"C:\Users\JessicaPoles\OneDrive - Timao\Bureau\Données essentielles\decp-2024.json"
chemin_resultats = r"C:\Users\JessicaPoles\OneDrive - Timao\Bureau\Données essentielles\resultats_recherche.json"


# Flag de débogage : on affiche le pattern généré une seule fois
DEBUG = True
expression_debug_affichee = False


# ------------------------------------------------------------------------
#       FONCTIONS DE NORMALISATION ET CONSTRUCTION DE REGEX
# ------------------------------------------------------------------------
def normaliser(texte):
    """Normalise un texte en supprimant les accents et en le mettant en minuscules."""
    return unidecode(texte).lower()


def construire_regex_variantes(variantes):
    """
    Construit une regex pour matcher l'une des variantes comme mot isolé.
    Exemple : pour variantes = ["video", "vidéo"], renvoie : (?i)\b(?:video|vidéo)\b
    """
    if not variantes:
        return r"$^"  # Regex qui ne matche jamais
    escaped = [re.escape(v) for v in variantes]
    return r"(?i)\b(?:%s)\b" % "|".join(escaped)


def construire_regex_groupe(mots):
    """
    Construit une regex pour un groupe de mots à rechercher dans l'ordre strict.
    Seuls les espaces ou ponctuations (caractères non alphanumériques) sont autorisés entre les mots.
    Par exemple, pour mots = ["intelligence", "artificielle"], renvoie :
       (?i)\bintelligence\b[^\w]+\bartificielle\b
    """
    blocs = []
    for mot in mots:
        blocs.append(r"\b" + re.escape(mot) + r"\b")
    # [^\w]+ : une ou plusieurs caractères non alphanumériques (espace ou ponctuation)
    pattern = r"(?i)" + r"[^\w]+".join(blocs)
    return pattern


# ------------------------------------------------------------------------
#        CONSTRUCTION DE L'EXPRESSION PYTHON DE RECHERCHE
# ------------------------------------------------------------------------
def construire_expression(expression, texte_normalise, original_texte):
    """
    Transforme l'expression logique saisie par l'utilisateur en une expression Python.
    - Remplace " ET ", " OU ", " SANS " par les opérateurs logiques Python.
    - Gère les parenthèses.
    - Gère la recherche exacte (termes entre guillemets) et les groupes de mots.
    """
    # Remplacer les opérateurs logiques (avec les espaces autour)
    expr = expression.replace(" ET ", " and ").replace(" OU ", " or ").replace(" SANS ", " and not ")


    # Découper l'expression en tokens (on sépare sur espaces et opérateurs)
    tokens = re.split(r'(\s+|\(|\)|and|or|not)', expr)
    tokens = [tok for tok in tokens if tok.strip() != ""]


    expression_modifiee = []
    i = 0
    while i < len(tokens):
        token = tokens[i].strip()
        # Si le token est un opérateur ou une parenthèse, le conserver tel quel
        if token in {"and", "or", "not", "(", ")"}:
            expression_modifiee.append(token)
            i += 1
            continue


        # Recherche exacte : si le token est entre guillemets, on le traite sans extension
        if token.startswith('"') and token.endswith('"') and len(token) > 1:
            terme_exact = token.strip('"')
            # Construction du pattern exact en utilisant json.dumps pour échapper correctement
            pattern = r"\b" + re.escape(terme_exact) + r"\b"
            expression_modifiee.append(f'bool(re.search({json.dumps(pattern)}, original_texte))')
            i += 1
            continue


        # Sinon, on cherche à constituer un groupe de mots : on rassemble tous les tokens consécutifs
        groupe = [token]
        i += 1
        while i < len(tokens):
            suivant = tokens[i].strip()
            if suivant in {"and", "or", "not", "(", ")"}:
                break
            groupe.append(suivant)
            i += 1


        if len(groupe) == 1:
            # Un seul mot : recherche simple
            mot = normaliser(groupe[0])
            pattern = construire_regex_variantes([mot])
            expression_modifiee.append(f'bool(re.search({json.dumps(pattern)}, original_texte))')
        else:
            # Groupe de mots : on impose l'ordre strict
            groupe_norm = [normaliser(m) for m in groupe]
            pattern = construire_regex_groupe(groupe_norm)
            expression_modifiee.append(f'bool(re.search({json.dumps(pattern)}, original_texte))')


    return " ".join(expression_modifiee)


# ------------------------------------------------------------------------
#                 FONCTIONS DE RECHERCHE
# ------------------------------------------------------------------------
def recherche_avancee(expression, texte):
    """
    Évalue l'expression logique sur le texte donné.
    Construit l'expression Python via construire_expression() et l'évalue.
    """
    original_texte = texte
    texte_normalise = normaliser(texte)
    expression_python = construire_expression(expression, texte_normalise, original_texte)
   
    global expression_debug_affichee
    if DEBUG and not expression_debug_affichee:
        print("Expression Python évaluée :", expression_python)
        expression_debug_affichee = True


    try:
        return eval(expression_python, {"re": re}, {"texte_normalise": texte_normalise, "original_texte": original_texte})
    except Exception as e:
        print("Erreur lors de l'évaluation de l'expression :", e)
        return False


def rechercher_par_expression(fichier_json, expression):
    """
    Parcourt le dataset et applique recherche_avancee() sur chaque "objet" de marché.
    Retourne la liste des marchés correspondants.
    """
    resultats = []
    with open(fichier_json, "r", encoding="utf-8") as fichier:
        donnees = json.load(fichier)


    marches = donnees.get("marches", {}).get("marche", [])
    print("Nombre total de marchés chargés :", len(marches))


    for marche in marches:
        texte_objet = marche.get("objet", "")
        if recherche_avancee(expression, texte_objet):
            print(f"Ligne trouvée : {texte_objet}")
            resultats.append(marche)


    return resultats


# ------------------------------------------------------------------------
#                             MAIN
# ------------------------------------------------------------------------
if __name__ == "__main__":
    expression_logique = input("Entrez votre expression logique : ")
    resultats_positifs = rechercher_par_expression(chemin_fichier, expression_logique)
    print(f"\nNombre de résultats trouvés : {len(resultats_positifs)}")
    if resultats_positifs:
        with open(chemin_resultats, "w", encoding="utf-8") as fichier:
            json.dump(resultats_positifs, fichier, ensure_ascii=False, indent=4)
        print(f"Les résultats ont été sauvegardés dans : {chemin_resultats}")
    else:
        print("Aucun résultat trouvé, aucun fichier n'a été créé.")










import os
import json
import pandas as pd


# Paramètres des dossiers
input_folder = "Mini fichiers"  # Dossier contenant les fichiers JSON découpés
output_folder = "Excels"        # Dossier pour les fichiers Excel générés


# Création du dossier de sortie s'il n'existe pas
if not os.path.exists(output_folder):
    os.makedirs(output_folder)


# Conversion des fichiers JSON en Excel
def convert_json_to_excel():
    try:
        # Parcourir tous les fichiers JSON dans le dossier d'entrée
        for file_name in os.listdir(input_folder):
            if file_name.endswith(".json"):
                input_path = os.path.join(input_folder, file_name)
                output_path = os.path.join(output_folder, file_name.replace(".json", ".xlsx"))


                # Charger le fichier JSON
                print(f"Conversion en cours : {file_name}")
                with open(input_path, "r", encoding="utf-8") as json_file:
                    data = json.load(json_file)


                # Convertir en DataFrame Pandas
                if isinstance(data, list):  # Le fichier doit contenir une liste d'objets
                    df = pd.DataFrame(data)
                    # Exporter en Excel
                    df.to_excel(output_path, index=False)
                    print(f"Fichier Excel créé : {output_path}")
                else:
                    print(f"⚠️ Le fichier {file_name} ne contient pas une liste d'objets, ignoré.")


        print("Conversion terminée avec succès !")


    except Exception as e:
        print(f"Une erreur est survenue : {e}")


# Exécuter la fonction
if __name__ == "__main__":
    convert_json_to_excel()