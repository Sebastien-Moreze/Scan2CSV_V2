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

# -----------------------------
# Dépendances nécessaires :
# pip install PyPDF2 anthropic python-dotenv pdfplumber pytesseract pillow
# Installer Tesseract-OCR : https://github.com/tesseract-ocr/tesseract
# Par défaut sur Windows : C:\\Program Files\\Tesseract-OCR\\tesseract.exe
# -----------------------------

# -----------------------------
# Ce script charge la clé API Anthropic depuis un fichier .env
# Exemple de fichier .env à placer dans le même dossier :
# ANTHROPIC_API_KEY=VOTRE_CLE_API_ICI
# -----------------------------
load_dotenv()
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
if not ANTHROPIC_API_KEY:
    raise RuntimeError("La clé API Anthropic n'est pas définie dans le fichier .env.")

# Si besoin, préciser le chemin de tesseract (décommenter et adapter si erreur)
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# -----------------------------
# Fonction : Extraction du texte du PDF (avec fallback OCR)
# -----------------------------
def extraire_texte_pdf(chemin_pdf):
    """Extrait le texte d'un fichier PDF. Si échec, tente l'OCR avec Tesseract."""
    try:
        # 1. Essai avec PyPDF2
        texte = ""
        with open(chemin_pdf, 'rb') as f:
            lecteur = PyPDF2.PdfReader(f)
            for page in lecteur.pages:
                texte += page.extract_text() or ""
        if texte.strip():
            return texte
        # 2. Si rien trouvé, essai OCR avec pdfplumber + pytesseract
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
# Fonction : Appel à l'API Anthropic (Claude) pour plusieurs entreprises
# -----------------------------
def extraire_infos_avec_anthropic(texte):
    """Appelle l'API Anthropic pour extraire toutes les entreprises du document."""
    try:
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        prompt = (
            "Voici le texte extrait d'un document PDF. "
            "Merci d'extraire toutes les entreprises présentes dans ce document, avec pour chacune :\n"
            "- nom de l'entreprise\n"
            "- nom du contact\n"
            "- numéro de téléphone\n"
            "- adresse de l'entreprise\n"
            "- URL\n"
            "- résumé de la description (2-3 phrases)\n"
            "\n"
            "Réponds STRICTEMENT et UNIQUEMENT avec la liste JSON, sans aucun texte, commentaire, ni balise de code avant ou après.\n"
            "Format attendu :\n"
            "[\n  {\"nom_entreprise\": ..., \"nom_contact\": ..., \"telephone\": ..., \"adresse\": ..., \"url\": ..., \"resume\": ...},\n  ...\n]\n"
            f"\nTexte :\n{texte}\n"
        )
        response = client.messages.create(
            model="claude-3-opus-20240229",
            max_tokens=2048,
            temperature=0,
            messages=[{"role": "user", "content": prompt}]
        )
        import json
        import re
        # Nettoyage de la réponse : suppression des balises de code éventuelles
        reponse = response.content[0].text.strip()
        reponse = re.sub(r'```json|```', '', reponse, flags=re.IGNORECASE).strip()
        # Affichage debug (optionnel)
        print('--- Réponse Claude brute ---')
        print(reponse)
        print('----------------------------')
        # Cherche la plus grande séquence entre le premier [ et le dernier ]
        start = reponse.find('[')
        end = reponse.rfind(']')
        if start != -1 and end != -1 and end > start:
            json_str = reponse[start:end+1]
            # Remplacer null par "" pour éviter les erreurs CSV
            json_str = json_str.replace(': null', ': ""')
            infos_liste = json.loads(json_str)
        else:
            raise ValueError("Réponse de l'API non conforme (liste JSON non trouvée)")
        # S'assurer que chaque entrée a toutes les clés
        cles = ["nom_entreprise", "nom_contact", "telephone", "adresse", "url", "resume"]
        for infos in infos_liste:
            for cle in cles:
                if cle not in infos:
                    infos[cle] = ""
        return infos_liste
    except Exception as e:
        raise RuntimeError(f"Erreur lors de l'appel à l'API Anthropic : {e}")

# -----------------------------
# Fonction : Export CSV (plusieurs entreprises)
# -----------------------------
def exporter_csv(infos_liste, chemin_csv):
    """Exporte la liste des entreprises extraites dans un fichier CSV."""
    try:
        with open(chemin_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=["nom_entreprise", "nom_contact", "telephone", "adresse", "url", "resume"])
            writer.writeheader()
            for infos in infos_liste:
                writer.writerow(infos)
    except Exception as e:
        raise RuntimeError(f"Erreur lors de l'export CSV : {e}")

# -----------------------------
# Fonction : Interface graphique (Tkinter)
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
        # Sélection du fichier PDF
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