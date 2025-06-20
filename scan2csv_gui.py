import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import csv
import PyPDF2
import anthropic
from dotenv import load_dotenv
import pdfplumber
import pytesseract
from PIL import Image
import platform
import subprocess

# -----------------------------
# Ce script permet d'extraire des informations d'entreprises depuis un PDF et de les exporter en CSV
# Il utilise l'OCR si besoin, l'API Claude d'Anthropic pour structurer les données, et une interface Tkinter
# -----------------------------

# Chargement des variables d'environnement (clé API Anthropic)
load_dotenv()
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
if not ANTHROPIC_API_KEY:
    raise RuntimeError("La clé API Anthropic n'est pas définie dans le fichier .env.")

# Détection automatique du chemin de Tesseract selon l'OS
# (nécessaire pour l'OCR avec pytesseract)
def configurer_tesseract():
    system = platform.system()
    if system == "Windows":
        tesseract_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
        if os.path.exists(tesseract_path):
            return tesseract_path
    elif system == "Darwin":  # macOS
        try:
            result = subprocess.run(['which', 'tesseract'], capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip()
        except:
            pass
    elif system == "Linux":
        try:
            result = subprocess.run(['which', 'tesseract'], capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip()
        except:
            pass
    return None

tesseract_path = configurer_tesseract()
if tesseract_path:
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# -----------------------------
# Extraction du texte du PDF (avec fallback OCR si besoin)
# -----------------------------
def extraire_texte_pdf(chemin_pdf):
    """
    Extrait le texte d'un fichier PDF.
    1. Tente d'abord avec PyPDF2 (texte natif)
    2. Si rien n'est extrait, utilise l'OCR (pdfplumber + pytesseract)
    """
    try:
        texte = ""
        with open(chemin_pdf, 'rb') as f:
            lecteur = PyPDF2.PdfReader(f)
            for page in lecteur.pages:
                texte += page.extract_text() or ""
        if texte.strip():
            return texte
        # Si rien trouvé, fallback OCR
        texte_ocr = ""
        with pdfplumber.open(chemin_pdf) as pdf:
            for page in pdf.pages:
                image = page.to_image(resolution=300)
                pil_image = image.original.convert("RGB")
                texte_ocr += pytesseract.image_to_string(pil_image, lang="fra+eng")
        if not texte_ocr.strip():
            raise ValueError("Aucun texte extrait du PDF, même après OCR.")
        return texte_ocr
    except Exception as e:
        raise RuntimeError(f"Erreur lors de l'extraction du PDF : {e}")

# -----------------------------
# Extraction des entreprises via l'API Anthropic (Claude)
# -----------------------------
def extraire_infos_avec_anthropic(texte):
    """
    Appelle l'API Anthropic pour extraire toutes les entreprises du document.
    Découpe le texte en morceaux si besoin (pour éviter la troncature)
    Fusionne et déduplique les résultats (nom + adresse)
    """
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        # Découpage du texte en chunks si trop long
        max_chunk_size = 15000  # Limite de taille pour chaque requête
        chunks = []
        if len(texte) > max_chunk_size:
            words = texte.split()
            current_chunk = ""
            for word in words:
                if len(current_chunk + " " + word) < max_chunk_size:
                    current_chunk += " " + word if current_chunk else word
                else:
                    if current_chunk:
                        chunks.append(current_chunk.strip())
                    current_chunk = word
            if current_chunk:
                chunks.append(current_chunk.strip())
        else:
            chunks = [texte]
        all_entreprises = []
        # Traitement de chaque chunk séparément
        for i, chunk in enumerate(chunks):
            print(f"Traitement du chunk {i+1}/{len(chunks)}...")
            # Prompt très directif pour Claude
            prompt = (
                "Voici le texte extrait d'un document PDF contenant une liste d'entreprises. "
                "IMPORTANT : Tu DOIS extraire TOUTES les entreprises présentes dans ce texte, sans exception.\n\n"
                "Pour chaque entreprise, extrait :\n"
                "- nom_entreprise : le nom de l'entreprise\n"
                "- nom_contact : le nom du contact principal\n"
                "- telephone : le numéro de téléphone\n"
                "- adresse : l'adresse complète de l'entreprise\n"
                "- url : le site web de l'entreprise\n"
                "- resume : un résumé de 2-3 phrases de l'activité\n\n"
                "RÈGLES STRICTES :\n"
                "1. Tu DOIS extraire TOUTES les entreprises du texte\n"
                "2. Réponds UNIQUEMENT avec le JSON, sans texte avant ou après\n"
                "3. Si une information est manquante, utilise une chaîne vide \"\"\n"
                "4. Assure-toi que le JSON est complet et valide\n\n"
                "Format JSON attendu :\n"
                "[\n"
                '  {"nom_entreprise": "...", "nom_contact": "...", "telephone": "...", "adresse": "...", "url": "...", "resume": "..."},\n'
                "  ...\n"
                "]\n\n"
                f"Texte du document (chunk {i+1}/{len(chunks)}) :\n{chunk}\n"
            )
            # Appel à l'API Anthropic
            response = client.messages.create(
                model="claude-3-opus-20240229",
                max_tokens=4096,  # Limite du modèle
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            import json
            import re
            # Nettoyage de la réponse : suppression des balises de code éventuelles
            reponse = response.content[0].text.strip()
            reponse = re.sub(r'```json|```', '', reponse, flags=re.IGNORECASE).strip()
            print(f'--- Réponse Claude brute (chunk {i+1}) ---')
            print(reponse)
            print('----------------------------')
            # Extraction du JSON dans la réponse
            json_str = None
            start = reponse.find('[')
            end = reponse.rfind(']')
            if start != -1 and end != -1 and end > start:
                json_str = reponse[start:end+1]
            if not json_str:
                json_pattern = r'\[.*?\]'
                matches = re.findall(json_pattern, reponse, re.DOTALL)
                if matches:
                    json_str = max(matches, key=len)
            if not json_str:
                try:
                    json.loads(reponse)
                    json_str = reponse
                except:
                    pass
            if not json_str and reponse.count('[') > 0:
                start = reponse.find('[')
                if start != -1:
                    partial_json = reponse[start:]
                    open_braces = partial_json.count('{')
                    close_braces = partial_json.count('}')
                    open_brackets = partial_json.count('[')
                    close_brackets = partial_json.count(']')
                    missing_braces = open_braces - close_braces
                    missing_brackets = open_brackets - close_brackets
                    if missing_braces > 0:
                        partial_json += '}' * missing_braces
                    if missing_brackets > 0:
                        partial_json += ']' * missing_brackets
                    json_str = partial_json
            if json_str:
                json_str = json_str.replace(': null', ': ""')
                json_str = json_str.replace(':null', ': ""')
                try:
                    entreprises_chunk = json.loads(json_str)
                except json.JSONDecodeError as e:
                    print(f"Erreur JSON: {e}")
                    print(f"JSON à parser: {json_str}")
                    try:
                        json_str = re.sub(r'[^\x00-\x7F]+', '', json_str)
                        json_str = re.sub(r',\s*}', '}', json_str)
                        json_str = re.sub(r',\s*]', ']', json_str)
                        entreprises_chunk = json.loads(json_str)
                    except:
                        print(f"JSON invalide et non réparable pour le chunk {i+1}")
                        continue
                # Vérification des clés pour chaque entreprise
                cles = ["nom_entreprise", "nom_contact", "telephone", "adresse", "url", "resume"]
                for infos in entreprises_chunk:
                    for cle in cles:
                        if cle not in infos:
                            infos[cle] = ""
                all_entreprises.extend(entreprises_chunk)
                print(f"Entreprises extraites du chunk {i+1} : {len(entreprises_chunk)}")
            else:
                print(f"Aucune entreprise trouvée dans le chunk {i+1}")
        # Déduplication sur (nom, adresse)
        entreprises_uniques = []
        cles_vues = set()
        for entreprise in all_entreprises:
            cle = (
                entreprise.get("nom_entreprise", "").strip().upper(),
                entreprise.get("adresse", "").strip().upper()
            )
            if cle not in cles_vues:
                entreprises_uniques.append(entreprise)
                cles_vues.add(cle)
        print(f"Nombre total d'entreprises extraites (après déduplication nom+adresse) : {len(entreprises_uniques)}")
        return entreprises_uniques
    except Exception as e:
        raise RuntimeError(f"Erreur lors de l'appel à l'API Anthropic : {e}")

# -----------------------------
# Export CSV
# -----------------------------
def exporter_csv(infos_liste, chemin_csv):
    """
    Exporte la liste des entreprises extraites dans un fichier CSV.
    """
    try:
        with open(chemin_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=["nom_entreprise", "nom_contact", "telephone", "adresse", "url", "resume"])
            writer.writeheader()
            for infos in infos_liste:
                writer.writerow(infos)
    except Exception as e:
        raise RuntimeError(f"Erreur lors de l'export CSV : {e}")

# -----------------------------
# Interface graphique principale (Tkinter)
# -----------------------------
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Scan2CSV - Extraction PDF vers CSV (Claude)")
        self.geometry("700x500")
        self.resizable(False, False)
        self.chemin_pdf = None
        self.infos_liste = None
        self.chemin_csv = None
        self.creer_widgets()

    def creer_widgets(self):
        # Zone de sélection du PDF
        frame_select = ttk.Frame(self)
        frame_select.pack(pady=20)
        self.label_pdf = ttk.Label(frame_select, text="Aucun fichier PDF sélectionné.")
        self.label_pdf.pack(side=tk.LEFT, padx=10)
        btn_select = ttk.Button(frame_select, text="Sélectionner un PDF", command=self.selectionner_pdf)
        btn_select.pack(side=tk.LEFT)
        # Bouton de traitement
        self.btn_traiter = ttk.Button(self, text="Lancer le traitement", command=self.lancer_traitement, state=tk.DISABLED)
        self.btn_traiter.pack(pady=10)
        # Zone d'aperçu des informations extraites
        self.frame_apercu = ttk.LabelFrame(self, text="Aperçu des informations extraites")
        self.frame_apercu.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.text_apercu = tk.Text(self.frame_apercu, height=15, wrap=tk.WORD, state=tk.DISABLED)
        self.text_apercu.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # Bouton de téléchargement CSV
        self.btn_csv = ttk.Button(self, text="Télécharger le CSV", command=self.telecharger_csv, state=tk.DISABLED)
        self.btn_csv.pack(pady=10)

    def selectionner_pdf(self):
        # Ouvre une boîte de dialogue pour sélectionner un PDF
        chemin = filedialog.askopenfilename(
            title="Sélectionner un fichier PDF",
            filetypes=[("Fichiers PDF", "*.pdf")]
        )
        if chemin:
            self.chemin_pdf = chemin
            self.label_pdf.config(text=os.path.basename(chemin))
            self.btn_traiter.config(state=tk.NORMAL)
            self.text_apercu.config(state=tk.NORMAL)
            self.text_apercu.delete(1.0, tk.END)
            self.text_apercu.config(state=tk.DISABLED)
            self.btn_csv.config(state=tk.DISABLED)
            self.infos_liste = None
            self.chemin_csv = None

    def lancer_traitement(self):
        # Lance l'extraction et l'appel API
        if not self.chemin_pdf:
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier PDF.")
            return
        try:
            self.text_apercu.config(state=tk.NORMAL)
            self.text_apercu.delete(1.0, tk.END)
            self.text_apercu.insert(tk.END, "Extraction du texte du PDF...\n")
            self.text_apercu.update()
            texte = extraire_texte_pdf(self.chemin_pdf)
            self.text_apercu.insert(tk.END, "Appel à l'API Anthropic...\n")
            self.text_apercu.update()
            infos_liste = extraire_infos_avec_anthropic(texte)
            self.infos_liste = infos_liste
            # Affichage de toutes les entreprises dans l'aperçu
            self.text_apercu.delete(1.0, tk.END)
            if not infos_liste:
                self.text_apercu.insert(tk.END, "Aucune entreprise trouvée.\n")
            else:
                for idx, infos in enumerate(infos_liste, 1):
                    self.text_apercu.insert(tk.END, f"Entreprise {idx} :\n")
                    for cle, val in infos.items():
                        self.text_apercu.insert(tk.END, f"  {cle} : {val}\n")
                    self.text_apercu.insert(tk.END, "\n")
            self.text_apercu.config(state=tk.DISABLED)
            self.btn_csv.config(state=tk.NORMAL)
        except Exception as e:
            self.text_apercu.config(state=tk.NORMAL)
            self.text_apercu.insert(tk.END, f"\nErreur : {e}\n")
            self.text_apercu.config(state=tk.DISABLED)
            messagebox.showerror("Erreur", str(e))
            self.btn_csv.config(state=tk.DISABLED)

    def telecharger_csv(self):
        # Ouvre une boîte de dialogue pour enregistrer le CSV
        if not self.infos_liste:
            messagebox.showerror("Erreur", "Aucune information à exporter.")
            return
        chemin = filedialog.asksaveasfilename(
            title="Enregistrer le CSV",
            defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv")]
        )
        if chemin:
            try:
                exporter_csv(self.infos_liste, chemin)
                self.chemin_csv = chemin
                messagebox.showinfo("Succès", f"Fichier CSV enregistré : {chemin}")
            except Exception as e:
                messagebox.showerror("Erreur", str(e))

# -----------------------------
# Lancement de l'application
# -----------------------------
if __name__ == "__main__":
    app = Application()
    app.mainloop() 