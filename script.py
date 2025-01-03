import os
import re
import fitz
from urllib.parse import urlparse

#Fonction de récupération de la liste des fichiers PDF
def list_pdfs(directory):
    pdf_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    return pdf_files

#Avoir le chemin de OneDrive
def get_onedrive_root():
    onedrive_path = os.environ.get('OneDrive') or os.environ.get('ONEDRIVE')
    if onedrive_path:
        return onedrive_path
    else:
        raise FileNotFoundError("OneDrive n'est pas configuré ou la variable d'environnement est manquante.")

# Fonction pour convertir un chemin local en URL SharePoint
def local_path_to_sharepoint_url(local_path):
    try:
        relative_path = local_path.split(folder_path_by_user_path, 1)[-1]
        sharepoint_url = f"{SHAREPOINT_BASE_URL}/{relative_path.replace(os.sep, '/').lstrip('/')}"
        return sharepoint_url
    except IndexError:
        raise ValueError("Le chemin local ne contient pas le dossier racine attendu.")


#Fonction de recherche de la phrase contenant le nom des fichiers pdf et ajout du lien
def search_and_link_phrase(pdf_path, sentence, link_target):
    try:
        document = fitz.open(pdf_path)
        last_page = document[-1]
        text_instances = last_page.search_for(sentence)

        if text_instances:
            for inst in text_instances:
                link_uri = local_path_to_sharepoint_url(link_target)
                last_page.insert_link({
                    "from": inst,
                    "uri": link_uri,
                    "kind": fitz.LINK_URI
                })
        document.save(pdf_path, incremental=True, encryption=0)
        return True
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
    return False

#Trouver le nom du site SharePoint
def extract_sharepoint_site_name(sharepoint_url):
    try:
        path_parts = urlparse(sharepoint_url).path.split("/")
        site_name_index = path_parts.index("sites") + 1
        return path_parts[site_name_index]
    except (ValueError, IndexError):
        raise ValueError("L'URL SharePoint est invalide ou ne contient pas '/sites/'.")

# Fonction pour extraire les phrases commençant par DSOP_Methode
def extract_dsop_phrases_from_last_page(pdf_path):
    try:
        document = fitz.open(pdf_path)
        last_page = document[-1]
        text = last_page.get_text("text")

        pattern = r'\bDSOP_Methode\s?.*'
        matches = re.findall(pattern, text, re.IGNORECASE)

        matches = [match.strip() for match in matches]
        return matches
    except Exception as e:
        print(f"Erreur lors de l'analyse de {pdf_path}: {e}")
        return []

#Gestion des fichiers non trouvé
def log_missing_file(source_file, missing_file, log_file):
    log_entry = f"{missing_file} => {source_file}\n"
    if not os.path.exists(log_file):
        with open(log_file, "w") as f:
            f.write(log_entry)
    else:
        with open(log_file, "r") as f:
            if log_entry in f.read():
                return
        with open(log_file, "a") as f:
            f.write(log_entry)

#Supprime les espaces des noms de fichiers
def normalize_name(name):
    return name.replace(" ", "").lower()


#Trouver le fichier local sur OneDrive
def find_folder_in_onedrive(site_name, onedrive_root):
    try:
        normalized_site_name = normalize_name(site_name)

        print("Dossiers trouvés dans OneDrive :")
        print(os.listdir(onedrive_root))

        for folder in os.listdir(onedrive_root):
            folder_path = os.path.join(onedrive_root, folder)
            if os.path.isdir(folder_path) and normalize_name(folder).startswith(normalized_site_name):
                return folder
        raise FileNotFoundError(f"Aucun dossier correspondant à '{site_name}' trouvé dans OneDrive.")
    except FileNotFoundError as e:
        raise e


#Constante de chemin d'accès à l'utilisateur et du dossier de recherche des fichiers PDF
user_path = get_onedrive_root()
SHAREPOINT_BASE_URL = input("Entrez l'URL de base de SharePoint jusqu'à Documents Partagés : ").strip()
site_name = extract_sharepoint_site_name(SHAREPOINT_BASE_URL)
print(f"Nom du site SharePoint extrait : {site_name}")
folder_path_by_user_path = find_folder_in_onedrive(site_name, user_path)
print(f"Dossier OneDrive correspondant : {folder_path_by_user_path}")
log_file = os.path.join(user_path, folder_path_by_user_path, "missing_links.txt")




#Recherche des fichiers PDF
directory = os.path.join(user_path, folder_path_by_user_path)
pdf_files = list_pdfs(directory)
dsop_data = {}

for pdf in pdf_files:
    dsop_phrases = extract_dsop_phrases_from_last_page(pdf)
    dsop_data[os.path.basename(os.path.basename(pdf).split('.')[0])] = dsop_phrases
    for phrase in dsop_phrases:
        found = False
        normalized_phrase = normalize_name(phrase)
        for pd in pdf_files:
            normalized_file = normalize_name(os.path.basename(pd))
            if normalized_phrase + ".pdf" == normalized_file:
                search_and_link_phrase(pdf, phrase, pd)
                print(f"Phrase correspondante trouvée : {phrase}")
                found = True
        if not found:
            log_missing_file(phrase + ".pdf", os.path.basename(pdf), log_file)
            print(f"Phrase introuvable dans les noms de fichiers : {normalized_phrase}.pdf, {pdf}")
