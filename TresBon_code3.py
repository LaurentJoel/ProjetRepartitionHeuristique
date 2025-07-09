import streamlit as st
import pandas as pd
import random
import glob
import os
import uuid
import tempfile
from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import Color, red, blue, green, orange, purple, brown, pink, gray
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from PyPDF2 import PdfReader, PdfWriter, PdfMerger  # Pour fusionner avec la première page existante

# Fonction de nettoyage des fichiers temporaires
def cleanup_temp_files():
    """Nettoie les fichiers temporaires créés lors des uploads"""
    temp_files = glob.glob("temp_*")
    for temp_file in temp_files:
        try:
            os.remove(temp_file)
        except:
            pass  # Ignorer les erreurs de suppression

# Couleurs pour différencier les épreuves
COULEURS_EPREUVES = [
    red, blue, green, orange, purple, brown, pink, gray,
    Color(0.8, 0.2, 0.6), Color(0.2, 0.8, 0.2), Color(0.6, 0.4, 0.8)
]

# --- Fonctions utilitaires pour le dessin PDF ---
def _dessiner_rangee_grille(canvas, x_pos, y_pos, rangee, largeur_cellule, hauteur_cellule, 
                           couleur_par_epreuve, font_name="Helvetica", padding=0.1*cm):
    """
    Dessine une rangée de sièges comme une grille visuelle dans le PDF.
    
    Args:
        canvas: Canvas ReportLab où dessiner
        x_pos, y_pos: Position de départ (haut gauche)
        rangee: Liste de listes représentant la rangée et ses places
        largeur_cellule, hauteur_cellule: Dimensions des cellules
        couleur_par_epreuve: Dictionnaire associant chaque épreuve à une couleur
        font_name: Nom de la police à utiliser
        padding: Espacement interne dans les cellules
    
    Returns:
        y_final: Position Y après la grille (pour continuer le dessin après)
    """    # Numéros de colonnes - optimisés pour plus de colonnes
    canvas.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, 10)  # Police plus grande
    for col_idx in range(len(rangee[0]) if rangee else 0):
        canvas.setFillColor(colors.lightgrey)
        canvas.rect(x_pos + (col_idx * largeur_cellule), y_pos, largeur_cellule, hauteur_cellule, fill=1, stroke=1)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(
            x_pos + (col_idx * largeur_cellule) + (largeur_cellule / 2),
            y_pos + (hauteur_cellule / 2) - 1,  # Centrer mieux
            f"{col_idx+1:02d}"
        )
    
    # Dessiner les lignes avec leurs places
    y_current = y_pos - hauteur_cellule
    
    for ligne_idx, ligne in enumerate(rangee):
        # Numéro de ligne (première colonne) - optimisé
        canvas.setFillColor(colors.lightgrey)
        canvas.rect(x_pos - hauteur_cellule, y_current, hauteur_cellule, hauteur_cellule, fill=1, stroke=1)
        canvas.setFillColor(colors.black)
        canvas.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, 10)  # Police cohérente plus grande
        canvas.drawCentredString(
            x_pos - (hauteur_cellule / 2),
            y_current + (hauteur_cellule / 2) - 1,  # Centrer mieux
            f"{ligne_idx+1:02d}"
        )
        
        # Dessiner chaque place
        for col_idx, place in enumerate(ligne):
            x_cell = x_pos + (col_idx * largeur_cellule)
            
            # Dessiner la cellule
            if place:  # Cellule occupée
                nom_etudiant, matiere = place
                couleur = couleur_par_epreuve.get(matiere, colors.white)
                canvas.setFillColor(couleur)
            else:
                canvas.setFillColor(colors.white)
            
            # Rectangle de la place
            canvas.rect(x_cell, y_current, largeur_cellule, hauteur_cellule, fill=1, stroke=1)
            
            # Texte de la place (nom de l'étudiant)
            if place:
                nom_etudiant, _ = place
                canvas.setFillColor(colors.white)  # Texte blanc sur fond coloré
                canvas.setFont(font_name, 8)  # Police plus grande pour s'adapter aux cellules plus grandes
                
                # Diviser le nom pour l'affichage optimisé dans des cellules plus grandes
                nom_court = nom_etudiant[:18] if nom_etudiant else ""  # Augmenter à 18 caractères max
                parties_nom = nom_court.split(' ')
                
                if len(parties_nom) == 1 or largeur_cellule < 2.2 * cm:
                    # Afficher sur une seule ligne si une seule partie ou cellule encore trop petite
                    canvas.drawCentredString(
                        x_cell + (largeur_cellule / 2),
                        y_current + (hauteur_cellule / 2) - 2,  # Centrer verticalement avec plus d'espace
                        nom_court[:15]  # Permettre plus de caractères
                    )
                else:
                    # Si deux parties (prénom nom), les afficher sur deux lignes
                    canvas.drawCentredString(
                        x_cell + (largeur_cellule / 2),
                        y_current + (hauteur_cellule / 2) + 4,  # Ligne du haut avec plus d'espace
                        parties_nom[0][:10]  # Permettre plus de caractères pour le prénom
                    )
                    canvas.drawCentredString(
                        x_cell + (largeur_cellule / 2),
                        y_current + (hauteur_cellule / 2) - 6,  # Ligne du bas avec plus d'espace
                        parties_nom[1][:10] if len(parties_nom) > 1 else ""  # Permettre plus de caractères pour le nom
                    )
        
        y_current -= hauteur_cellule
    
    return y_current  # Retourner la position Y finale

def _dessiner_porte(canvas, x, y, largeur=2*cm, hauteur=1.5*cm, font_name="Helvetica"):
    """
    Dessine une représentation réaliste d'une porte dans le PDF.
    
    Args:
        canvas: Canvas ReportLab où dessiner
        x, y: Position de départ (bas gauche)
        largeur, hauteur: Dimensions de la porte
        font_name: Nom de la police à utiliser
    """
    # Dessin du cadre de la porte (rectangle extérieur)
    canvas.setStrokeColor(colors.black)
    canvas.setFillColor(colors.lightgrey)
    canvas.rect(x, y, largeur, hauteur, fill=1, stroke=1)
    
    # Dessin de la porte elle-même (rectangle intérieur)
    door_inset = 0.1 * cm
    canvas.setFillColor(colors.darkgrey)
    canvas.rect(x + door_inset, y + door_inset, 
               largeur - (2 * door_inset), hauteur - (2 * door_inset), fill=1, stroke=0)
    
    # Panneau de la porte (lignes de détail)
    panel_inset = 0.3 * cm
    canvas.setStrokeColor(colors.black)
    canvas.setLineWidth(0.5)
    canvas.rect(x + panel_inset, y + panel_inset,
               largeur - (2 * panel_inset), hauteur - (2 * panel_inset), fill=0, stroke=1)
    
    # Poignée de porte
    canvas.setFillColor(colors.black)
    canvas.circle(x + largeur - (0.5 * cm), y + (hauteur / 2), 0.1 * cm, fill=1, stroke=0)
    
    # Trait pour la poignée
    canvas.setLineWidth(1)
    canvas.line(x + largeur - (0.7 * cm), y + (hauteur / 2),
               x + largeur - (0.4 * cm), y + (hauteur / 2))
    
    # Étiquette pour la porte
    canvas.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, 9)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(x + (largeur / 2), y - 0.3 * cm, "PORTE PRINCIPALE")

# --- Définition des structures des salles ---
STRUCTURES_SALLES = {
    "Amphitheatre": {"gauche": (10, 10), "droite": (10, 10)},
    "ISE1-MATH": {"gauche": (7, 4), "droite": (5, 4)},
    "AS1": {"gauche": (5, 4), "droite": (5, 4)},
    "AS2": {"gauche": (5, 2), "milieu": (5, 2), "droite": (5, 2)},
    "AS3": {"gauche": (3, 4), "droite": (3, 4)},
    "ISEL1": {"gauche": (5, 2), "milieu": (5, 2), "droite": (5, 2)},
    "ISEL2": {"gauche": (5, 2), "milieu": (5, 2), "droite": (5, 2)},
    "ISEL3": {"gauche": (5, 2), "milieu": (5, 2), "droite": (5, 2)},
    "ISEECO": {"gauche": (5, 2), "milieu": (5, 2), "droite": (5, 2)},
    "ISEMATH": {"gauche": (7, 4), "droite": (5, 4)},
    "ISE3": {"gauche": (6, 4), "droite": (6, 4)},
    "TSS1": {"gauche": (6, 2), "milieu": (6,2), "droite": (6, 2)}
}

# --- Définition de la classe Salle pour gérer les placements ---
class Salle:
    def __init__(self, nom, structure, porte="gauche"):
        """
        Initialisation d'une salle d'examen.
        Args:
            nom (str): Nom de la salle (ex: AMPHI, ISE1-MATH).
            structure (dict): Structure des rangées (gauche, milieu, droite) avec dimensions (lignes, colonnes).
            porte (str): Position de la porte ("gauche" ou "droite").
        """
        self.nom = nom
        self.porte = porte
        self.structure = structure
        self.rangées = {rangée: [[None] * cols for _ in range(lignes)] 
                        for rangée, (lignes, cols) in structure.items() if rangée in ['gauche', 'milieu', 'droite']}
        self.placements_avec_contraintes_relachees = 0  # Compteur pour les placements avec contraintes relâchées

    def capacite_totale(self):
        """Calcule la capacité totale de la salle."""
        return sum(len(self.rangées[r]) * (len(self.rangées[r][0]) if self.rangées[r] else 0) for r in self.rangées)

    def nombre_etudiants(self):
        """Compte le nombre d'étudiants placés."""
        return sum(1 for rangée in self.rangées for ligne in self.rangées[rangée] for place in ligne if place is not None)

    def nombre_places_vides(self):
        """Calcule le nombre de places vides."""
        return self.capacite_totale() - self.nombre_etudiants()

    def taux_remplissage(self):
        """Calcule le taux de remplissage de la salle."""
        return (self.nombre_etudiants() / self.capacite_totale() * 100) if self.capacite_totale() > 0 else 0

    def place_valide(self, rangée, ligne_idx, col_idx, epreuve):
        """Vérifie si une place est valide pour un étudiant d'une épreuve donnée."""
        if rangée not in self.rangées or self.rangées[rangée][ligne_idx][col_idx] is not None:
            return False
        lignes, colonnes = len(self.rangées[rangée]), len(self.rangées[rangée][0])
        # Vérification des voisins horizontaux
        if col_idx > 0 and self.rangées[rangée][ligne_idx][col_idx-1] and self.rangées[rangée][ligne_idx][col_idx-1][1] == epreuve:
            return False
        if col_idx < colonnes-1 and self.rangées[rangée][ligne_idx][col_idx+1] and self.rangées[rangée][ligne_idx][col_idx+1][1] == epreuve:
            return False
        # Vérification des voisins verticaux
        if ligne_idx > 0 and self.rangées[rangée][ligne_idx-1][col_idx] and self.rangées[rangée][ligne_idx-1][col_idx][1] == epreuve:
            return False
        if ligne_idx < lignes-1 and self.rangées[rangée][ligne_idx+1][col_idx] and self.rangées[rangée][ligne_idx+1][col_idx][1] == epreuve:
            return False
        # Vérification des rangées adjacentes
        for r_adj in self.rangées_adj(rangée):
            if len(self.rangées[r_adj]) > ligne_idx and col_idx < len(self.rangées[r_adj][ligne_idx]):
                if self.rangées[r_adj][ligne_idx][col_idx] and self.rangées[r_adj][ligne_idx][col_idx][1] == epreuve:
                    return False
        return True

    def rangées_adj(self, rangée):
        """Retourne les rangées adjacentes à une rangée donnée."""
        if rangée == 'gauche':
            return ['milieu'] if 'milieu' in self.rangées and self.rangées['milieu'] else ['droite'] if 'droite' in self.rangées and self.rangées['droite'] else []
        elif rangée == 'milieu':
            return [r for r in ['gauche', 'droite'] if r in self.rangées and self.rangées[r]]
        elif rangée == 'droite':
            return ['milieu'] if 'milieu' in self.rangées and self.rangées['milieu'] else ['gauche'] if 'gauche' in self.rangées and self.rangées['gauche'] else []
        return []

    def placer_etudiant(self, nom_etudiant, epreuve):
        """Place un étudiant en remplissant de manière compacte pour éviter les espaces vides."""
        # Initialiser l'attribut si il n'existe pas (pour compatibilité avec les anciens objets)
        if not hasattr(self, 'placements_avec_contraintes_relachees'):
            self.placements_avec_contraintes_relachees = 0
            
        # Essayer d'abord le placement avec contraintes strictes
        if self._placement_compact_sequentiel(nom_etudiant, epreuve):
            return True
        
        # Si échec, essayer le placement sans contrainte d'adjacence de matière
        if self._placement_force_sequentiel(nom_etudiant, epreuve):
            self.placements_avec_contraintes_relachees += 1
            return True
        
        return False
    
    def _placement_force_sequentiel(self, nom_etudiant, epreuve):
        """Placement forcé sans contrainte d'adjacence de matière - garantit le placement si une place libre existe."""
        # Ordre de priorité: gauche d'abord, puis milieu, puis droite
        rangees_ordre = ['gauche', 'milieu', 'droite']
        
        for rangee in rangees_ordre:
            if rangee not in self.rangées or not self.rangées[rangee]:
                continue
            
            # Parcourir ligne par ligne de façon séquentielle
            for li in range(len(self.rangées[rangee])):
                for ci in range(len(self.rangées[rangee][li])):
                    # Vérifier seulement si la position est libre (pas de contrainte de matière)
                    if self.rangées[rangee][li][ci] is None:
                        # Placer l'étudiant
                        self.rangées[rangee][li][ci] = (nom_etudiant, epreuve)
                        return True
        
        return False
    
    def _placement_compact_sequentiel(self, nom_etudiant, epreuve):
        """Algorithme de placement compact séquentiel pour éviter les espaces vides."""
        # Ordre de priorité: gauche d'abord, puis milieu, puis droite
        rangees_ordre = ['gauche', 'milieu', 'droite']
        
        for rangee in rangees_ordre:
            if rangee not in self.rangées or not self.rangées[rangee]:
                continue
            
            # Parcourir ligne par ligne de façon séquentielle
            for li in range(len(self.rangées[rangee])):
                for ci in range(len(self.rangées[rangee][li])):
                    # Vérifier si cette position est libre et valide
                    if self.rangées[rangee][li][ci] is None and self.place_valide(rangee, li, ci, epreuve):
                        # Placer l'étudiant
                        self.rangées[rangee][li][ci] = (nom_etudiant, epreuve)
                        return True
        
        return False
    
    def _backtrack_placement_optimise(self, nom_etudiant, epreuve):
        """Algorithme de backtracking optimisé avec priorité intelligente des rangées."""
        # Ordre de priorité intelligent: gauche, puis milieu, puis droite
        rangees_disponibles = [r for r in ['gauche', 'milieu', 'droite'] if r in self.rangées and self.rangées[r]]
        
        # Vérifier le taux de remplissage pour déterminer la priorité
        rangees_avec_stats = []
        for rangee in rangees_disponibles:
            places_occupees = sum(1 for ligne in self.rangées[rangee] for place in ligne if place is not None)
            places_totales = sum(len(ligne) for ligne in self.rangées[rangee])
            taux_remplissage = places_occupees / places_totales if places_totales > 0 else 0
            rangees_avec_stats.append((rangee, taux_remplissage, places_occupees))
        
        # Stratégie: si gauche est bien remplie (>70%), priorité au milieu
        # Sinon, continuer avec l'ordre normal
        if rangees_avec_stats:
            # Trouver la rangée gauche
            gauche_stats = next((stats for stats in rangees_avec_stats if stats[0] == 'gauche'), None)
            
            if gauche_stats and gauche_stats[1] > 0.7:  # Si gauche est remplie à plus de 70%
                # Priorité au milieu
                ordre_priorite = ['milieu', 'droite', 'gauche']
            else:
                # Ordre normal: gauche d'abord
                ordre_priorite = ['gauche', 'milieu', 'droite']
        else:
            ordre_priorite = ['gauche', 'milieu', 'droite']
        
        # Essayer le placement dans l'ordre de priorité
        for rangee in ordre_priorite:
            if rangee not in self.rangées or not self.rangées[rangee]:
                continue
                
            for li in range(len(self.rangées[rangee])):
                for ci in range(len(self.rangées[rangee][li])):
                    # Vérifier si cette position est valide
                    if self.place_valide(rangee, li, ci, epreuve):
                        # Essayer de placer l'étudiant ici
                        self.rangées[rangee][li][ci] = (nom_etudiant, epreuve)
                        
                        # Vérifier si le placement respecte toutes les contraintes
                        if self._valider_placement_global(rangee, li, ci, epreuve):
                            return True
                        
                        # Si le placement n'est pas optimal, backtrack
                        self.rangées[rangee][li][ci] = None
        
        return False
    
    def _backtrack_placement(self, nom_etudiant, epreuve, rangee_idx, ligne_idx, col_idx):
        """Algorithme de backtracking pour placement optimal."""
        # Ordre de priorité: gauche d'abord, puis milieu, puis droite
        # Cela garantit que le milieu est rempli après la première rangée
        rangees_ordre = ['gauche', 'milieu', 'droite']
        
        # Parcourir toutes les positions à partir de la position courante
        for r_idx in range(rangee_idx, len(rangees_ordre)):
            rangee = rangees_ordre[r_idx]
            if rangee not in self.rangées or not self.rangées[rangee]:
                continue
                
            start_ligne = ligne_idx if r_idx == rangee_idx else 0
            for li in range(start_ligne, len(self.rangées[rangee])):
                start_col = col_idx if (r_idx == rangee_idx and li == ligne_idx) else 0
                for ci in range(start_col, len(self.rangées[rangee][li])):
                    
                    # Vérifier si cette position est valide
                    if self.place_valide(rangee, li, ci, epreuve):
                        # Essayer de placer l'étudiant ici
                        self.rangées[rangee][li][ci] = (nom_etudiant, epreuve)
                        
                        # Vérifier si le placement respecte toutes les contraintes
                        if self._valider_placement_global(rangee, li, ci, epreuve):
                            return True
                        
                        # Si le placement n'est pas optimal, backtrack
                        self.rangées[rangee][li][ci] = None
        
        return False
    
    def _valider_placement_global(self, rangee, ligne_idx, col_idx, epreuve):
        """Validation globale du placement avec heuristiques d'optimisation."""
        # Vérifier que le placement maintient une bonne distribution
        return self._check_distribution_heuristic(rangee, ligne_idx, col_idx, epreuve)
    
    def _check_distribution_heuristic(self, rangee, ligne_idx, col_idx, epreuve):
        """Heuristique pour maintenir une distribution équilibrée."""
        # Compter les étudiants de la même épreuve dans un rayon de 2 cases
        count_same_subject = 0
        radius = 2
        
        for r in ['gauche', 'milieu', 'droite']:
            if r not in self.rangées:
                continue
            for li in range(max(0, ligne_idx - radius), min(len(self.rangées[r]), ligne_idx + radius + 1)):
                for ci in range(max(0, col_idx - radius), min(len(self.rangées[r][li]), col_idx + radius + 1)):
                    if (r != rangee or li != ligne_idx or ci != col_idx):
                        if self.rangées[r][li][ci] and self.rangées[r][li][ci][1] == epreuve:
                            count_same_subject += 1
        
        # Limiter le nombre d'étudiants de même épreuve dans le voisinage
        return count_same_subject <= 2

# --- Génération du PDF avec layout visuel des salles ---
def generer_pdf(salles, semestre, date_epreuve, heure_debut, heure_fin, matieres_par_classe, chemin_pdf="repartition_examens.pdf"):
    """
    Génère un PDF avec un layout visuel des salles d'examen montrant l'arrangement des sièges.
    Chaque salle utilisée aura sa propre page avec un tableau représentant la disposition physique.
    Inclut la première page du fichier '20250130_Répartition_S1N.pdf' comme page de couverture.
    """
    # Nom du fichier temporaire pour la génération des pages de layout
    temp_pdf_path = f"temp_{uuid.uuid4()}.pdf"
    
    # Initialisation du canevas PDF pour les pages de layout
    page_size = landscape(A4)  # Utilisation du format paysage pour plus d'espace horizontal
    c = canvas.Canvas(temp_pdf_path, pagesize=page_size)
    largeur, hauteur = page_size  # En paysage: largeur = 29.7cm, hauteur = 21cm
    marge_gauche = 1.5 * cm  # Réduire les marges pour plus d'espace
    marge_droite = 1.5 * cm
    marge_haut = hauteur - 1.5 * cm  # Réduire la marge haute
    taille_titre = 16
    taille_texte = 10

    # Tentative d'enregistrement de la police Times New Roman
    try:
        pdfmetrics.registerFont(TTFont('TimesNewRoman', 'times.ttf'))
        # Essayer d'enregistrer la variante bold si elle existe
        try:
            pdfmetrics.registerFont(TTFont('TimesNewRoman-Bold', 'timesbd.ttf'))
        except:
            pass  # Ignorer si la police bold n'est pas disponible
        font_name = "TimesNewRoman"
    except:
        font_name = "Helvetica"

    # --- Page de légende des couleurs par matière (seconde page après la couverture) ---
    # Titre centré pour la légende
    c.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, taille_titre)
    c.drawCentredString(largeur / 2, marge_haut, f"LÉGENDE DES MATIÈRES - {semestre}")
    c.setFont(font_name, 12)
    c.drawCentredString(largeur / 2, marge_haut - 1 * cm, f"Date: {date_epreuve} | {heure_debut} - {heure_fin}")
    
    # --- Légende des couleurs par matière ---
    c.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, 14)
    c.drawCentredString(largeur / 2, marge_haut - 2 * cm, "Légende des couleurs par matière:")
    
    # Obtenir les matières uniques pour la légende
    epreuves_uniques = list(set(matieres_par_classe.values()))
    couleur_par_epreuve = {epreuve: COULEURS_EPREUVES[i % len(COULEURS_EPREUVES)] for i, epreuve in enumerate(epreuves_uniques)}
    
    # Calculer la mise en page de la légende - optimisée pour le mode paysage
    # Nombre de colonnes pour la légende
    nb_colonnes = 3  # Plus de colonnes en mode paysage
    nb_epreuves = len(epreuves_uniques)
    items_par_colonne = (nb_epreuves + (nb_colonnes - 1)) // nb_colonnes  # Arrondi au supérieur
    
    # Dessiner les rectangles de couleur avec leur nom de matière
    y_pos = marge_haut - 3.5 * cm  # Position plus haute car moins d'éléments sur la page
    starting_y_pos = y_pos
    
    for i, epreuve in enumerate(epreuves_uniques):
        # Calculer la position X, Y pour une mise en page multi-colonnes
        colonne = i // items_par_colonne
        position_dans_colonne = i % items_par_colonne
        
        x_pos = marge_gauche + colonne * (largeur / 3)  # Diviser l'espace horizontal en 3 colonnes
        y_pos = starting_y_pos - (position_dans_colonne * 0.8 * cm)
        
        couleur = couleur_par_epreuve[epreuve]
        
        # Rectangle de couleur - plus grand en mode paysage
        c.setFillColor(couleur)
        c.rect(x_pos, y_pos, 1.5 * cm, 0.8 * cm, fill=1, stroke=0)
        
        # Nom de l'épreuve - police plus grande et en gras pour meilleure lisibilité
        c.setFillColor(colors.black)
        c.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, 14)  # Police plus grande et gras
        c.drawString(x_pos + 2 * cm, y_pos + 0.3 * cm, f"{epreuve}")
    
    # Texte explicatif en bas de la page de légende
    c.setFont(font_name, 10)
    c.drawCentredString(largeur / 2, 3 * cm, "Les couleurs ci-dessus sont utilisées pour identifier les matières dans les plans de salles.")
    c.drawCentredString(largeur / 2, 2.5 * cm, "Chaque étudiant est placé de façon à éviter que deux étudiants de la même matière soient côte à côte.")
    
    c.showPage()

    # --- Pages pour chaque salle avec layout visuel ---
    for salle in salles:
        if salle.nombre_etudiants() == 0:
            continue  # Ignorer les salles vides
            
        # En-tête de la page pour la salle
        c.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, taille_titre + 2)  # Taille augmentée et gras
        c.drawCentredString(largeur / 2, marge_haut, f"Salle : {salle.nom}")
        
        # Informations sur la salle
        c.setFont(font_name, 12)
        capacite_totale = salle.capacite_totale()
        etudiants_places = salle.nombre_etudiants()
        c.drawString(marge_gauche, marge_haut - 1.5 * cm, f"Capacité: {capacite_totale} places")
        c.drawString(marge_gauche + 6 * cm, marge_haut - 1.5 * cm, f"Occupée: {etudiants_places} étudiants")
        c.drawString(marge_gauche + 12 * cm, marge_haut - 1.5 * cm, f"Taux: {(etudiants_places/capacite_totale)*100:.1f}%")

        y_current = marge_haut - 2.5 * cm  # Commencer plus haut pour plus d'espace pour les tableaux
        
        # Créer le layout visuel pour chaque rangée
        rangee_order = ['gauche', 'milieu', 'droite']
        rangee_titles = {'gauche': 'Rangée de gauche', 'milieu': 'Rangée centrale', 'droite': 'Rangée de droite'}
        
        # Préparer les couleurs pour les épreuves
        epreuves_uniques = list(set(matieres_par_classe.values()))
        couleur_par_epreuve = {epreuve: COULEURS_EPREUVES[i % len(COULEURS_EPREUVES)] for i, epreuve in enumerate(epreuves_uniques)}
        
        # Préparer les rangées à afficher
        rangees_a_afficher = []
        for rangee_nom in rangee_order:
            if rangee_nom in salle.rangées and salle.rangées[rangee_nom]:
                rangee = salle.rangées[rangee_nom]
                if rangee and not all(len(ligne) == 0 for ligne in rangee):
                    rangees_a_afficher.append((rangee_nom, rangee))
        
        # Traiter les rangées en essayant de maximiser l'utilisation de l'espace
        i = 0
        while i < len(rangees_a_afficher):
            # Calculer l'espace disponible sur la page actuelle
            espace_disponible_page = y_current - 5 * cm  # Garder 5cm en bas pour la date et la porte
            
            # Essayer d'ajouter autant de rangées que possible sur cette page
            rangees_pour_cette_page = []
            espace_utilise = 0
            
            for j in range(i, len(rangees_a_afficher)):
                rangee_nom, rangee = rangees_a_afficher[j]
                
                # Calculer les dimensions pour cette rangée
                max_cols = max(len(ligne) for ligne in rangee) if rangee else 0
                if max_cols == 0:
                    continue
                
                # Calculer la largeur des cellules pour s'adapter à la page paysage
                espace_disponible_largeur = largeur - 6*cm  # 6cm de marges totales
                largeur_cellule = min(2.8, espace_disponible_largeur / max_cols) * cm  # Augmenter la taille max des cellules
                hauteur_cellule = 1.2 * cm  # Augmenter la hauteur des cellules
                
                # Calculer l'espace nécessaire pour cette rangée (plus d'espace)
                hauteur_rangee_titre = 1.8 * cm  # Augmenter l'espace du titre
                hauteur_grille = hauteur_cellule * (len(rangee) + 1)  # +1 pour l'en-tête des colonnes
                hauteur_rangee_totale = hauteur_rangee_titre + hauteur_grille + 1.5 * cm  # Augmenter l'espacement
                
                # Vérifier si on peut ajouter cette rangée à la page actuelle
                if espace_utilise + hauteur_rangee_totale <= espace_disponible_page:
                    rangees_pour_cette_page.append((j, rangee_nom, rangee, largeur_cellule, hauteur_cellule, hauteur_rangee_totale))
                    espace_utilise += hauteur_rangee_totale
                else:
                    # Cette rangée ne rentre pas, arrêter l'ajout pour cette page
                    break
            
            # Si aucune rangée ne peut être ajoutée (cas d'une rangée trop grande), forcer l'ajout d'au moins une
            if not rangees_pour_cette_page and i < len(rangees_a_afficher):
                rangee_nom, rangee = rangees_a_afficher[i]
                max_cols = max(len(ligne) for ligne in rangee) if rangee else 0
                if max_cols > 0:
                    espace_disponible_largeur = largeur - 6*cm
                    largeur_cellule = min(2.5, espace_disponible_largeur / max_cols) * cm  # Cellules plus grandes même en cas de force
                    hauteur_cellule = 1.0 * cm  # Hauteur plus grande même en cas de force
                    hauteur_rangee_titre = 1.8 * cm
                    hauteur_grille = hauteur_cellule * (len(rangee) + 1)
                    hauteur_rangee_totale = hauteur_rangee_titre + hauteur_grille + 1.5 * cm
                    rangees_pour_cette_page.append((i, rangee_nom, rangee, largeur_cellule, hauteur_cellule, hauteur_rangee_totale))
            
            # Dessiner toutes les rangées pour cette page
            for idx, rangee_nom, rangee, largeur_cellule, hauteur_cellule, hauteur_rangee_totale in rangees_pour_cette_page:
                # Afficher le titre de la rangée
                c.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, 16)  # Police plus grande pour le titre
                c.drawString(marge_gauche, y_current, rangee_titles[rangee_nom])
                y_current -= 1.8 * cm  # Espacement augmenté entre titre et tableau
                
                # Dessiner la grille de la rangée
                x_pos_start = marge_gauche + 1.0 * cm  # Réduire le décalage
                y_final = _dessiner_rangee_grille(
                    c, x_pos_start, y_current, rangee, largeur_cellule, hauteur_cellule, 
                    couleur_par_epreuve, font_name
                )
                
                # Mettre à jour la position Y pour la prochaine rangée (espacement augmenté)
                y_current = y_final - 1.5 * cm
            
            # Passer aux rangées suivantes (celles qui n'ont pas été traitées sur cette page)
            i += len(rangees_pour_cette_page)
            
            # Si toutes les rangées de cette page ont été traitées et qu'il en reste d'autres
            if i < len(rangees_a_afficher):
                # Vérifier s'il faut une nouvelle page pour les rangées suivantes (critère plus restrictif)
                espace_restant = y_current - 5 * cm
                if espace_restant < 6 * cm:  # Seulement si moins de 6cm restants (au lieu de 8cm)
                    c.showPage()
                    c.setFont(font_name + '-Bold' if font_name == "TimesNewRoman" else font_name, taille_titre + 2)
                    c.drawCentredString(largeur / 2, marge_haut, f"Salle : {salle.nom} (suite)")
                    y_current = marge_haut - 2.5 * cm
        
        # Date et heure en bas de page
        c.setFillColor(colors.black)
        c.setFont(font_name, taille_texte)
        c.drawString(marge_gauche, 2 * cm, f"{date_epreuve}    {heure_debut} - {heure_fin}")
        
        # Utiliser la fonction utilitaire pour dessiner une porte réaliste
        _dessiner_porte(c, largeur - 4 * cm, 2.7 * cm, 2 * cm, 1.5 * cm, font_name)
        
        c.showPage()

    # Finaliser le PDF temporaire contenant les pages de layout
    c.save()
    
    # Fusionner les PDFs: 
    # 1. La première page du fichier "20250130_Répartition_S1N.pdf" comme couverture
    # 2. Les pages de layout générées
    
    merger = PdfMerger()
    
    # 1. Ajouter la première page du fichier couverture
    cover_pdf_path = "20250130_Répartition_S1N.pdf"
    temp_cover_path = None
    
    if os.path.exists(cover_pdf_path):
        try:
            cover_reader = PdfReader(cover_pdf_path)
            if len(cover_reader.pages) > 0:
                cover_writer = PdfWriter()
                # Ajouter uniquement la première page
                cover_writer.add_page(cover_reader.pages[0])
                
                # Enregistrer dans un fichier temporaire
                temp_cover_path = f"temp_cover_{uuid.uuid4()}.pdf"
                with open(temp_cover_path, "wb") as f:
                    cover_writer.write(f)
                
                # Ajouter au merger
                merger.append(temp_cover_path)
        except Exception as e:
            print(f"Erreur lors de l'accès au fichier de couverture: {e}")
            # En cas d'erreur, continuer sans la page de couverture
    else:
        print(f"Fichier de couverture '{cover_pdf_path}' non trouvé.")
    
    # 2. Ajouter les pages de layout générées
    merger.append(temp_pdf_path)
    
    # Sauvegarder le PDF final
    merger.write(chemin_pdf)
    merger.close()
    
    # Nettoyer les fichiers temporaires
    if os.path.exists(temp_pdf_path):
        os.remove(temp_pdf_path)
    
    if temp_cover_path and os.path.exists(temp_cover_path):
        os.remove(temp_cover_path)

# Fonction de réinitialisation de la session
def reset_session_state():
    """Réinitialise complètement l'état de la session après la génération du PDF."""
    keys_to_keep = []  # Vous pouvez ajouter des clés à conserver si nécessaire
    
    # Supprimer toutes les clés de session sauf celles à conserver
    for key in list(st.session_state.keys()):
        if key not in keys_to_keep:
            del st.session_state[key]
    
    # Réinitialiser avec les valeurs par défaut
    st.session_state.config_validated = False
    st.session_state.repartition_completed = False
    st.session_state.classes_choisies = []
    st.session_state.matieres_par_classe = {}
    st.session_state.salles_disponibles = []
    st.session_state.classes_selectionnees = []
    
    # Supprimer les fichiers temporaires
    cleanup_temp_files()

# Initialisation des variables de session
def init_session_state():
    """Initialise les variables de session pour éviter les erreurs."""
    if 'config_validated' not in st.session_state:
        st.session_state.config_validated = False
    if 'repartition_completed' not in st.session_state:
        st.session_state.repartition_completed = False
    if 'classes_choisies' not in st.session_state:
        st.session_state.classes_choisies = []
    if 'matieres_par_classe' not in st.session_state:
        st.session_state.matieres_par_classe = {}
    if 'salles_disponibles' not in st.session_state:
        st.session_state.salles_disponibles = []
    if 'classes_selectionnees' not in st.session_state:
        st.session_state.classes_selectionnees = []

# --- Application Streamlit ---
def main():
    """Application principale Streamlit pour la répartition des étudiants selon les spécifications du projet ISSEA."""
    # Initialisation des variables de session
    init_session_state()
    
    st.set_page_config(
        page_title="Répartition Optimisée des Étudiants aux Examens - ISSEA", 
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # En-tête avec logo et informations institutionnelles
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 20px; background-color: #f0f8ff; border-radius: 10px; margin-bottom: 30px;">
            <h1 style="color: #1e3d59;">📚 ISSEA - Répartition Optimisée des Étudiants</h1>
            <h3 style="color: #2c5aa0;">Institut Sous-régional de Statistique et d'Économie Appliquée</h3>
            <p style="color: #666;">Algorithme: Heuristique vorace + Backtracking avec contraintes strictes</p>
        </div>
        """, unsafe_allow_html=True)

    # Sidebar pour les paramètres principaux
    with st.sidebar:
        st.markdown("## 🎯 Paramètres de Répartition")
        
        # Étape 1: Choix du semestre
        st.markdown("### 📅 Semestre")
        semestre = st.radio(
            "Sélectionnez le semestre", 
            ["Semestre 1", "Semestre 2"], 
            help="Choisissez le semestre concerné par les examens"
        )
        
        # Étape 2: Date et horaires
        st.markdown("### ⏰ Planning")
        date_epreuve = st.date_input("📅 Date de l'épreuve", value=datetime.today())
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            heure_debut = st.time_input("🕐 Début", value=datetime.strptime("08:00", "%H:%M").time())
        with col_h2:
            heure_fin = st.time_input("🕕 Fin", value=datetime.strptime("12:00", "%H:%M").time())

    # Onglets pour organiser l'interface
    tab1, tab2, tab3, tab4 = st.tabs(["📂 Chargement des Données", "🎯 Configuration", "📊 Répartition", "📄 Export PDF"])
    
    with tab1:
        st.markdown("## 📂 Chargement des Fichiers de Données")
        
        st.info("""
        📤 **Instructions de téléchargement:**
        - Téléchargez vos fichiers Excel (.xlsx) en utilisant les boutons ci-dessous
        - Assurez-vous que les fichiers respectent le format attendu
        - Tous les fichiers sont obligatoires pour continuer
        """)
        
        # Colonnes pour les téléchargements
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("### 📋 Fichier des Matières")
            st.markdown(f"*Pour {semestre}*")
            fichier_matieres_upload = st.file_uploader(
                "Télécharger le fichier des matières",
                type=['xlsx', 'xls'],
                key="matieres_uploader",
                help=f"Fichier Excel contenant les matières pour {semestre}"
            )
            
            if fichier_matieres_upload:
                st.success(f"✅ {fichier_matieres_upload.name}")
                # Sauvegarder temporairement le fichier
                with open(f"temp_matieres_{fichier_matieres_upload.name}", "wb") as f:
                    f.write(fichier_matieres_upload.getbuffer())
                fichier_matieres = f"temp_matieres_{fichier_matieres_upload.name}"
                
                # Aperçu rapide du contenu
                try:
                    preview_excel = pd.ExcelFile(fichier_matieres)
                    st.write(f"📊 {len(preview_excel.sheet_names)} classe(s) trouvée(s)")
                except:
                    st.warning("⚠️ Erreur de lecture du fichier")
            else:
                st.warning("⚠️ Fichier manquant")
                fichier_matieres = None
        
        with col2:
            st.markdown("###   Fichier des Étudiants")
            st.markdown(f"*Pour {semestre}*")
            fichier_etudiants_upload = st.file_uploader(
                "Télécharger le fichier des étudiants",
                type=['xlsx', 'xls'],
                key="etudiants_uploader",
                help=f"Fichier Excel contenant la liste des étudiants pour {semestre}. Chaque feuille doit contenir les colonnes 'nom' et 'prenom'."
            )
            
            if fichier_etudiants_upload:
                st.success(f"✅ {fichier_etudiants_upload.name}")
                # Sauvegarder temporairement le fichier
                with open(f"temp_etudiants_{fichier_etudiants_upload.name}", "wb") as f:
                    f.write(fichier_etudiants_upload.getbuffer())
                fichier_etudiants = f"temp_etudiants_{fichier_etudiants_upload.name}"
                
                # Aperçu rapide du contenu et validation des colonnes
                try:
                    preview_excel = pd.ExcelFile(fichier_etudiants)
                    total_etu = 0
                    colonnes_valides = True
                    premiere_classe_testee = False
                    
                    for sheet in preview_excel.sheet_names:
                        df = pd.read_excel(fichier_etudiants, sheet_name=sheet)
                        if not df.empty:
                            total_etu += len(df)
                            
                            # Vérifier les colonnes obligatoires sur la première classe non-vide
                            if not premiere_classe_testee:
                                colonnes_disponibles = [col.lower().strip() for col in df.columns]
                                has_nom = any('nom' in col for col in colonnes_disponibles)
                                has_prenom = any('prenom' in col or 'prénom' in col for col in colonnes_disponibles)
                                
                                if not has_nom or not has_prenom:
                                    colonnes_valides = False
                                    colonnes_manquantes = []
                                    if not has_nom:
                                        colonnes_manquantes.append("nom")
                                    if not has_prenom:
                                        colonnes_manquantes.append("prenom")
                                    st.error(f"❌ Colonnes manquantes: {', '.join(colonnes_manquantes)}")
                                    st.info("📋 Colonnes détectées: " + ", ".join(df.columns.tolist()))
                                
                                premiere_classe_testee = True
                    
                    if colonnes_valides:
                        st.write(f"👥 {total_etu} étudiant(s) dans {len(preview_excel.sheet_names)} classe(s)")
                        st.success("✓ Colonnes 'nom' et 'prenom' détectées")
                    else:
                        st.warning("⚠️ Le fichier ne contient pas les colonnes obligatoires 'nom' et 'prenom'")
                        
                except Exception as e:
                    st.warning(f"⚠️ Erreur de lecture du fichier: {e}")
            else:
                st.warning("⚠️ Fichier manquant")
                fichier_etudiants = None
        
        with col3:
            st.markdown("### 🏫 Fichier des Salles")
            st.markdown("*Configuration des salles*")
            fichier_salles_upload = st.file_uploader(
                "Télécharger le fichier des salles",
                type=['xlsx', 'xls'],
                key="salles_uploader",
                help="Fichier Excel contenant les informations sur les salles d'examen"
            )
            
            if fichier_salles_upload:
                st.success(f"✅ {fichier_salles_upload.name}")
                # Sauvegarder temporairement le fichier
                with open(f"temp_salles_{fichier_salles_upload.name}", "wb") as f:
                    f.write(fichier_salles_upload.getbuffer())
                fichier_salles = f"temp_salles_{fichier_salles_upload.name}"
                
                # Aperçu rapide du contenu
                try:
                    preview_df = pd.read_excel(fichier_salles)
                    salles_count = len(preview_df) if not preview_df.empty else 0
                    st.write(f"🏫 {salles_count} salle(s) trouvée(s)")
                except:
                    st.warning("⚠️ Erreur de lecture du fichier")
            else:
                st.warning("⚠️ Fichier manquant")
                fichier_salles = None

        # Vérification des prérequis avec affichage détaillé
        fichiers_manquants = []
        if not fichier_matieres:
            fichiers_manquants.append("📚 Fichier des matières")
        if not fichier_etudiants:
            fichiers_manquants.append("👥 Fichier des étudiants")
        if not fichier_salles:
            fichiers_manquants.append("🏫 Fichier des salles")
        
        if fichiers_manquants:
            st.warning(f"⚠️ Fichiers manquants: {', '.join(fichiers_manquants)}")
            
            with st.expander("  Aide rapide", expanded=False):
                st.markdown("""
                **Veuillez télécharger les 3 fichiers Excel requis:**
                - 📚 **Matières**: Fichier avec les matières par classe
                - 👥 **Étudiants**: Fichier avec les listes d'étudiants par classe (obligatoire: colonnes 'nom' et 'prenom')
                - 🏫 **Salles**: Fichier avec les noms des salles d'examen
                
                **Format requis:**
                - Chaque fichier doit être au format Excel (.xlsx ou .xls)
                - Le fichier étudiants doit contenir les colonnes 'nom' et 'prenom' dans chaque feuille
                """)
            
            return

        # Affichage du statut global
        st.success("✅ Tous les fichiers ont été téléchargés avec succès!")
        
        # Bouton pour valider les fichiers
        if st.button("🔍 Valider et Analyser les Fichiers", type="primary"):
            try:
                # Test de lecture du fichier matières
                st.info("📚 Validation du fichier des matières...")
                matieres_test = pd.ExcelFile(fichier_matieres)
                
                # Vérification des feuilles
                if len(matieres_test.sheet_names) == 0:
                    st.error("❌ Le fichier des matières ne contient aucune feuille")
                else:
                    # Test de lecture d'une feuille
                    test_sheet = pd.read_excel(fichier_matieres, sheet_name=matieres_test.sheet_names[0])
                    if test_sheet.empty or len(test_sheet.columns) == 0:
                        st.error("❌ Le fichier des matières semble vide ou mal formaté")
                    else:
                        st.success(f"✅ Fichier matières validé ({len(matieres_test.sheet_names)} classes trouvées)")
                
                # Test de lecture du fichier étudiants
                st.info("👥 Validation du fichier des étudiants...")
                etudiants_test = pd.ExcelFile(fichier_etudiants)
                
                # Compter le total d'étudiants et vérifier les colonnes obligatoires
                total_students = 0
                colonnes_valides = True
                classes_avec_erreurs = []
                
                for sheet in etudiants_test.sheet_names:
                    try:
                        df_sheet = pd.read_excel(fichier_etudiants, sheet_name=sheet)
                        if not df_sheet.empty:
                            total_students += len(df_sheet)
                            
                            # Vérifier les colonnes obligatoires pour chaque classe
                            colonnes_disponibles = [col.lower().strip() for col in df_sheet.columns]
                            has_nom = any('nom' in col for col in colonnes_disponibles)
                            has_prenom = any('prenom' in col or 'prénom' in col for col in colonnes_disponibles)
                            
                            if not has_nom or not has_prenom:
                                colonnes_valides = False
                                colonnes_manquantes = []
                                if not has_nom:
                                    colonnes_manquantes.append("nom")
                                if not has_prenom:
                                    colonnes_manquantes.append("prenom")
                                classes_avec_erreurs.append(f"{sheet}: {', '.join(colonnes_manquantes)}")
                    except:
                        pass
                
                if total_students == 0:
                    st.error("❌ Aucun étudiant trouvé dans le fichier")
                elif not colonnes_valides:
                    st.error("❌ Colonnes obligatoires manquantes dans certaines classes:")
                    for erreur in classes_avec_erreurs:
                        st.error(f"  • {erreur}")
                    st.info("💡 Assurez-vous que chaque feuille Excel contient les colonnes 'nom' et 'prenom' (ou 'prénom')")
                else:
                    st.success(f"✅ Fichier étudiants validé ({total_students} étudiants, {len(etudiants_test.sheet_names)} classes)")
                    st.success("✓ Toutes les classes contiennent les colonnes 'nom' et 'prenom'")
                
                # Test de lecture du fichier salles
                st.info("🏫 Validation du fichier des salles...")
                salles_test = pd.read_excel(fichier_salles)
                
                # Vérifier la correspondance avec les structures définies
                col_nom = salles_test.columns[0]
                salles_reconnues = []
                salles_inconnues = []
                
                for salle in salles_test[col_nom]:
                    if str(salle) in STRUCTURES_SALLES:
                        salles_reconnues.append(str(salle))
                    else:
                        salles_inconnues.append(str(salle))
                
                if len(salles_reconnues) == 0:
                    st.error("❌ Aucune salle reconnue dans le fichier")
                    st.write(f"Salles supportées: {', '.join(STRUCTURES_SALLES.keys())}")
                else:
                    st.success(f"✅ Fichier salles validé ({len(salles_reconnues)} salles reconnues)")
                    if salles_inconnues:
                        st.warning(f"⚠️ Salles ignorées: {', '.join(salles_inconnues[:3])}{'...' if len(salles_inconnues) > 3 else ''}")
                
                st.success("🎉 Tous les fichiers ont été validés avec succès!")
                
                # Sauvegarder les fichiers dans la session
                st.session_state.fichier_matieres = fichier_matieres
                st.session_state.fichier_etudiants = fichier_etudiants
                st.session_state.fichier_salles = fichier_salles
                
                # Aperçu des données
                with st.expander(" ️ Aperçu des données chargées", expanded=False):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.write("**Classes (Matières):**")
                        for sheet in matieres_test.sheet_names[:5]:  # Afficher max 5
                            st.write(f"• {sheet}")
                        if len(matieres_test.sheet_names) > 5:
                            st.write(f"... et {len(matieres_test.sheet_names) - 5} autres")
                    
                    with col2:
                        st.write("**Classes (Étudiants):**")
                        for sheet in etudiants_test.sheet_names[:5]:  # Afficher max 5
                            st.write(f"• {sheet}")
                        if len(etudiants_test.sheet_names) > 5:
                            st.write(f"... et {len(etudiants_test.sheet_names) - 5} autres")
                    
                    with col3:
                        st.write("**Salles disponibles:**")
                        col_nom = salles_test.columns[0]
                        for salle in salles_test[col_nom].head(5):
                            st.write(f"• {salle}")
                        if len(salles_test) > 5:
                            st.write(f"... et {len(salles_test) - 5} autres")
                
            except Exception as e:
                st.error(f"❌ Erreur lors de la validation des fichiers: {e}")
                st.write("Vérifiez que vos fichiers respectent le format attendu.")

    with tab2:
        st.markdown("## 🎯 Configuration des Classes et Salles")
        
        # Vérifier que les fichiers sont disponibles dans la session ou localement
        fichier_matieres = st.session_state.get('fichier_matieres', fichier_matieres)
        fichier_etudiants = st.session_state.get('fichier_etudiants', fichier_etudiants)
        fichier_salles = st.session_state.get('fichier_salles', fichier_salles)
        
        if not all([fichier_matieres, fichier_etudiants, fichier_salles]):
            st.error("⚠️ Veuillez d'abord télécharger et valider tous les fichiers dans l'onglet 'Chargement des Données'.")
            return
        
        try:
            # Chargement des fichiers
            matieres_excel = pd.ExcelFile(fichier_matieres)
            excel_etu = pd.ExcelFile(fichier_etudiants)
            df_salles = pd.read_excel(fichier_salles)
            
            st.success("✅ Fichiers chargés avec succès!")
            
            # Configuration des classes d'abord (étape 1)
            st.markdown("### 👨‍🎓 Sélection des Classes pour l'Examen")
            st.info("🎯 **Étape 1:** Sélectionnez d'abord les classes qui participeront à l'examen")
            
            # Debug: Afficher les classes disponibles
            classes_disponibles_debug = excel_etu.sheet_names
            st.write(f"🔍 **Classes détectées dans le fichier étudiants:** {', '.join(classes_disponibles_debug)}")
            
            # Fonction pour calculer le nombre d'étudiants par classe
            def calculer_etudiants_classe(classe_nom):
                try:
                    df_classe = pd.read_excel(fichier_etudiants, sheet_name=classe_nom)
                    return len(df_classe) if not df_classe.empty else 0
                except Exception as e:
                    st.warning(f"⚠️ Erreur lors du calcul des étudiants pour {classe_nom}: {e}")
                    return 0
            
            # Calculer le nombre d'étudiants pour toutes les classes
            classes = excel_etu.sheet_names
            etudiants_par_classe_info = {}
            for classe in classes:
                etudiants_par_classe_info[classe] = calculer_etudiants_classe(classe)
            
            # Interface de sélection progressive des classes
            if 'classes_selectionnees' not in st.session_state:
                st.session_state.classes_selectionnees = []
            
            # Interface pour ajouter/retirer des classes
            col_add, col_remove = st.columns(2)
            
            with col_add:
                st.markdown("**➕ Ajouter une classe:**")
                classes_disponibles = [c for c in classes if c not in st.session_state.classes_selectionnees]
                
                # Ajout libre de classes - pas de limitation de capacité
                if classes_disponibles:
                    classe_a_ajouter = st.selectbox(
                        "Choisir une classe à ajouter",
                        [""] + classes_disponibles,
                        key="classe_add_select"
                    )
                    
                    if classe_a_ajouter and st.button("➕ Ajouter cette classe", key="btn_add"):
                        st.session_state.classes_selectionnees.append(classe_a_ajouter)
                        st.session_state.salles_disponibles = []  # Reset des salles quand on change les classes
                        etudiants_classe = etudiants_par_classe_info[classe_a_ajouter]
                        st.success(f"✅ Classe {classe_a_ajouter} ajoutée ({etudiants_classe} étudiants)")
                        st.rerun()
                else:
                    st.info("Toutes les classes ont été ajoutées")
            
            with col_remove:
                st.markdown("**➖ Retirer une classe:**")
                if st.session_state.classes_selectionnees:
                    classe_a_retirer = st.selectbox(
                        "Choisir une classe à retirer",
                        [""] + st.session_state.classes_selectionnees,
                        key="classe_remove_select"
                    )
                    
                    if classe_a_retirer and st.button("➖ Retirer cette classe", key="btn_remove"):
                        st.session_state.classes_selectionnees.remove(classe_a_retirer)
                        st.session_state.salles_disponibles = []  # Reset des salles quand on change les classes
                        etudiants_classe = etudiants_par_classe_info[classe_a_retirer]
                        st.success(f"✅ Classe {classe_a_retirer} retirée ({etudiants_classe} étudiants)")
                        st.rerun()
                else:
                    st.info("Aucune classe sélectionnée")
            
            # Affichage des classes sélectionnées
            if st.session_state.classes_selectionnees:
                st.markdown("**📋 Classes sélectionnées pour l'examen:**")
                total_etudiants_classes = 0
                for i, classe in enumerate(st.session_state.classes_selectionnees, 1):
                    nb_etudiants = etudiants_par_classe_info[classe]
                    total_etudiants_classes += nb_etudiants
                    st.write(f"{i}. **{classe}** - {nb_etudiants} étudiants")
                
                st.info(f"📊 **Total étudiants sélectionnés: {total_etudiants_classes} étudiants**")
                
                # Bouton pour réinitialiser la sélection des classes si nécessaire
                if st.button("🔄 Réinitialiser la sélection des classes", type="secondary", key="reset_classes"):
                    st.session_state.classes_selectionnees = []
                    st.session_state.salles_disponibles = []  # Reset des salles aussi
                    st.rerun()
            else:
                st.warning("⚠️ Aucune classe sélectionnée pour l'examen.")
                total_etudiants_classes = 0
            
            if not st.session_state.classes_selectionnees:
                st.warning("⚠️ Veuillez sélectionner au moins une classe pour continuer à l'étape suivante.")
                return
            
            # Configuration des salles (étape 2) - seulement si des classes sont sélectionnées
            st.markdown("### 🏫 Sélection des Salles d'Examen")
            st.info(f"🎯 **Étape 2:** Sélectionnez maintenant les salles pour accueillir {total_etudiants_classes} étudiants")
            
            col_nom = next((col for col in df_salles.columns if "nom" in col.lower() or "salle" in col.lower()), df_salles.columns[0])
            
            # Affichage des salles disponibles avec capacités
            st.markdown("**Salles disponibles:**")
            salles_info = []
            for _, row in df_salles.iterrows():
                nom_salle = row[col_nom]
                if nom_salle in STRUCTURES_SALLES:
                    try:
                        structure = STRUCTURES_SALLES[nom_salle]
                        capacite = sum(lignes * cols for lignes, cols in structure.values())
                        salles_info.append((nom_salle, capacite))
                        st.write(f"🏫 **{nom_salle}** - Capacité: {capacite} places")
                    except Exception as e:
                        st.warning(f"🏫 **{nom_salle}** - ⚠️ Erreur de calcul de capacité: {e}")
                else:
                    st.warning(f"🏫 **{nom_salle}** - ⚠️ Structure non définie (ignorée)")
            
            # Tri par capacité décroissante (contrainte du projet)
            salles_info.sort(key=lambda x: x[1], reverse=True)
            
            if not salles_info:
                st.error("❌ Aucune salle avec structure définie trouvée.")
                st.info(f"🔍 **Salles supportées:** {', '.join(STRUCTURES_SALLES.keys())}")
                return
            
            # Laisser l'utilisateur choisir librement les salles (pas de suggestion automatique)
            salles_disponibles = st.multiselect(
                "Salles à utiliser pour les examens", 
                [nom for nom, _ in salles_info],
                default=[],  # Aucune salle sélectionnée par défaut
                help="Sélectionnez les salles que vous souhaitez utiliser pour les examens. Les salles seront utilisées par ordre décroissant de capacité."
            )
            
            if not salles_disponibles:
                st.warning("⚠️ Veuillez sélectionner au moins une salle pour continuer.")
                return
            
            # Validation: si la première salle suffit et que plusieurs salles sont sélectionnées
            premiere_salle_capacite = salles_info[0][1] if salles_info else 0
            if len(salles_disponibles) > 1 and premiere_salle_capacite >= total_etudiants_classes:
                st.warning(f"""
                ⚠️ **Configuration sous-optimale détectée !**
                
                📊 **Analyse de capacité :**
                • Étudiants à placer : **{total_etudiants_classes}**
                • Capacité de {salles_info[0][0]} : **{premiere_salle_capacite}** places
                • Salles supplémentaires sélectionnées : **{len(salles_disponibles) - 1}**
                
                ✅ **Recommandation :** La salle **{salles_info[0][0]}** seule suffit largement !
                
                🎯 **Conseil :** Une configuration optimale utilise le minimum de salles nécessaires pour faciliter la surveillance.
                
                💡 **Vous pouvez continuer avec cette configuration ou ajuster votre sélection.**
                """)
            
            # Stocker les salles sélectionnées dans session_state pour validation croisée
            st.session_state.salles_disponibles = salles_disponibles
            
            # Calculer la capacité totale des salles sélectionnées
            capacite_totale_salles = sum(capacite for nom, capacite in salles_info if nom in salles_disponibles)
            
            # Vérification de la capacité
            if total_etudiants_classes > capacite_totale_salles:
                places_manquantes = total_etudiants_classes - capacite_totale_salles
                st.error(f"""
                ❌ **Capacité insuffisante!**
                
                • Étudiants à placer: {total_etudiants_classes}
                • Capacité des salles: {capacite_totale_salles}
                • Places manquantes: {places_manquantes}
                
                💡 **Solutions:**
                - Ajoutez plus de salles d'examen
                - Retirez des classes de l'examen
                """)
                return
            else:
                taux_utilisation = (total_etudiants_classes / capacite_totale_salles * 100) if capacite_totale_salles > 0 else 0
                st.success(f"✅ **Capacité suffisante!** - Taux d'utilisation: {taux_utilisation:.1f}%")
                st.info(f"📊 **Capacité totale des salles sélectionnées: {capacite_totale_salles} places**")
            
            # Affichage du statut actuel avec barre de progression
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("👥 Étudiants à placer", total_etudiants_classes)
            with col2:
                st.metric("🏫 Capacité disponible", capacite_totale_salles)
            with col3:
                places_restantes = capacite_totale_salles - total_etudiants_classes
                st.metric("📊 Places restantes", places_restantes)
            
            # Barre de progression de la capacité
            if capacite_totale_salles > 0:
                taux_occupation = min(total_etudiants_classes / capacite_totale_salles, 1.0)
                st.progress(taux_occupation)
                if taux_occupation >= 0.9:
                    st.warning(f"⚠️ Capacité presque atteinte ({taux_occupation*100:.1f}%)")
            
            # Bouton pour réinitialiser la sélection des salles si nécessaire
            if st.button("🔄 Réinitialiser la sélection des salles", type="secondary", key="reset_salles"):
                st.session_state.salles_disponibles = []
                st.rerun()
            
            # Utiliser les classes sélectionnées pour la suite
            classes_choisies = st.session_state.classes_selectionnees
            
            # Configuration des matières par classe
            st.markdown("### 📚 Attribution des Matières par Classe")
            matieres_par_classe = {}
            
            for classe in classes_choisies:
                with st.container():
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.write(f"**{classe}**")
                        st.write(f"👥 {etudiants_par_classe_info[classe]} étudiants")
                    with col2:
                        try:
                            df_mat = pd.read_excel(fichier_matieres, sheet_name=classe)
                            liste_matieres = df_mat.iloc[:, 0].dropna().tolist()
                            if liste_matieres:
                                mat = st.selectbox(
                                    f"Matière pour {classe}", 
                                    liste_matieres, 
                                    key=f"mat_{classe}",
                                    help=f"Choisissez l'épreuve que la classe {classe} va composer"
                                )
                                matieres_par_classe[classe] = mat
                            else:
                                st.error(f"❌ Aucune matière trouvée pour la classe {classe}")
                                return
                        except Exception as e:
                            st.error(f"❌ Erreur pour la classe {classe}: {e}")
                            return
            
            # Affichage du résumé de la configuration
            st.markdown("### 📋 Résumé de la Configuration")
            
            # Recalculer le total d'étudiants basé sur les classes sélectionnées
            total_etudiants = total_etudiants_classes
            
            # Affichage des détails par classe
            for classe in classes_choisies:
                nb_etudiants = etudiants_par_classe_info[classe]
                st.write(f"📊 **{classe}**: {nb_etudiants} étudiants - Matière: {matieres_par_classe[classe]}")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("👥 Total Étudiants", total_etudiants)
            with col2:
                st.metric("🏫 Capacité Totale", capacite_totale_salles)
            with col3:
                taux_utilisation = (total_etudiants / capacite_totale_salles * 100) if capacite_totale_salles > 0 else 0
                st.metric("📊 Taux d'Utilisation", f"{taux_utilisation:.1f}%")
            
            # Validation finale
            if total_etudiants > capacite_totale_salles:
                st.error("❌ Configuration invalide! Plus d'étudiants que de places disponibles.")
                st.error("💡 Retirez des classes ou ajoutez plus de salles.")
                return

            # Validation de la configuration
            if st.button("✅ Valider la Configuration", type="primary"):
                st.session_state.config_validated = True
                st.session_state.classes_choisies = classes_choisies
                st.session_state.matieres_par_classe = matieres_par_classe
                st.session_state.salles_disponibles = salles_disponibles
                st.session_state.fichier_etudiants = fichier_etudiants
                st.session_state.total_etudiants = total_etudiants
                st.session_state.capacite_totale = capacite_totale_salles
                st.success("✅ Configuration validée! Passez à l'onglet Répartition.")

        except Exception as e:
            st.error(f"❌ Erreur lors du chargement des fichiers: {e}")
            st.exception(e)
            return

    with tab3:
        if not st.session_state.get('config_validated', False):
            st.warning("⚠️ Veuillez d'abord valider la configuration dans l'onglet précédent.")
            return
            
        st.markdown("## 📊 Algorithme de Répartition Optimisée")
        
        # Récupération des données de session
        classes_choisies = st.session_state.classes_choisies
        matieres_par_classe = st.session_state.matieres_par_classe
        salles_disponibles = st.session_state.salles_disponibles
        fichier_etudiants = st.session_state.fichier_etudiants
        total_etudiants = st.session_state.total_etudiants
        
        # Affichage de l'algorithme utilisé
        with st.expander("🔍 Description de l'Algorithme", expanded=True):
            st.markdown("""
            **Algorithme de Placement Compact par Classe avec Garantie de Placement**
            
            1. **Phase de Tri**: Classes triées par nombre d'étudiants (décroissant)
            2. **Phase de Placement**: Placement classe par classe (remplissage complet avant passage à la suivante)
            3. **Phase de Remplissage**: Placement séquentiel pour éviter les espaces vides
            4. **Contraintes Préférentielles** (appliquées en priorité):
               - Éviter que des étudiants de même matière soient côte à côte
               - Remplissage par ordre décroissant de capacité des salles
               - Placement séquentiel: gauche d'abord, puis milieu, puis droite
               - Évitement des espaces vides (placement compact)
               - Une classe entière est placée avant de passer à la suivante
            5. **Mécanisme de Fallback**: Si les contraintes préférentielles empêchent le placement,
               les contraintes d'adjacence de matière sont relâchées pour **garantir que tous les étudiants soient placés**
            
            ✅ **Garantie**: Tant qu'il y a assez de places dans les salles, **TOUS** les étudiants seront placés.
            """)

        # Bouton pour lancer la répartition
        if st.button("🚀 Lancer la Répartition Automatique", type="primary"):
            with st.spinner("⏳ Répartition en cours..."):
                
                # Étape 1: Chargement des données étudiants
                st.info("📊 Chargement des données des étudiants...")
                etudiants_par_classe = {}
                
                progress_bar = st.progress(0)
                for i, classe in enumerate(classes_choisies):
                    try:
                        df_classe = pd.read_excel(fichier_etudiants, sheet_name=classe)
                        # Prendre la première colonne comme noms des étudiants
                        col_nom = df_classe.columns[0]
                        etudiants_par_classe[classe] = df_classe[col_nom].dropna().tolist()
                        progress_bar.progress((i + 1) / len(classes_choisies))
                    except Exception as e:
                        st.error(f"❌ Erreur lors du chargement de la classe {classe}: {e}")
                        return
                
                progress_bar.empty()
                
                # Étape 2: Création des objets salles
                st.info("🏫 Initialisation des salles...")
                objets_salles = []
                for nom_salle in salles_disponibles:
                    if nom_salle in STRUCTURES_SALLES:
                        structure = STRUCTURES_SALLES[nom_salle]
                        salle = Salle(nom_salle, structure)
                        objets_salles.append(salle)
                
                # Tri des salles par capacité décroissante (contrainte du projet)
                objets_salles.sort(key=lambda s: s.capacite_totale(), reverse=True)
                
                # Étape 3: Calcul des statistiques
                total_places = sum(salle.capacite_totale() for salle in objets_salles)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("👥 Étudiants à placer", total_etudiants)
                with col2:
                    st.metric("🏫 Places disponibles", total_places)
                with col3:
                    utilisation = (total_etudiants / total_places * 100) if total_places > 0 else 0
                    st.metric("📊 Taux d'utilisation", f"{utilisation:.1f}%")
                
                # Étape 4: Placement avec algorithme heuristique et backtracking
                st.info("🔄 Application de l'algorithme de placement...")
                
                # Regrouper les étudiants par classe (pas par matière)
                etudiants_par_classe_ordonnee = {}
                for classe in classes_choisies:
                    matiere = matieres_par_classe[classe]
                    etudiants_par_classe_ordonnee[classe] = []
                    for etu in etudiants_par_classe[classe]:
                        etudiants_par_classe_ordonnee[classe].append((etu, matiere, classe))
                
                # Heuristique: Trier les classes par nombre d'étudiants (décroissant)
                classes_triees = sorted(etudiants_par_classe_ordonnee.keys(), 
                                      key=lambda c: len(etudiants_par_classe_ordonnee[c]), 
                                      reverse=True)
                
                # Mélanger les étudiants dans chaque classe pour éviter les patterns
                for classe in etudiants_par_classe_ordonnee:
                    random.shuffle(etudiants_par_classe_ordonnee[classe])
                
                # Phase de placement classe par classe (remplissage complet)
                non_places = []
                statistiques_placement = {salle.nom: {} for salle in objets_salles}
                tentatives_backtrack = 0
                max_tentatives = 3
                
                progress_bar = st.progress(0)
                total_classes = len(classes_triees)
                
                for classe_idx, classe in enumerate(classes_triees):
                    st.info(f"📚 Placement de la classe {classe} ({len(etudiants_par_classe_ordonnee[classe])} étudiants)")
                    
                    etudiants_classe = etudiants_par_classe_ordonnee[classe]
                    etudiants_non_places_classe = []
                    
                    for idx, (etu, matiere, classe_nom) in enumerate(etudiants_classe):
                        place_trouvee = False
                        
                        # Essayer le placement dans toutes les salles disponibles
                        for salle in objets_salles:
                            if salle.placer_etudiant(etu, matiere):
                                place_trouvee = True
                                statistiques_placement[salle.nom][matiere] = statistiques_placement[salle.nom].get(matiere, 0) + 1
                                break
                        
                        # Si échec de placement pour cet étudiant
                        if not place_trouvee:
                            etudiants_non_places_classe.append((etu, matiere))
                    
                    # Si des étudiants de cette classe n'ont pas pu être placés
                    if etudiants_non_places_classe:
                        non_places.extend(etudiants_non_places_classe)
                        st.warning(f"⚠️ {len(etudiants_non_places_classe)} étudiants de la classe {classe} n'ont pas pu être placés")
                    else:
                        st.success(f"✅ Classe {classe} entièrement placée")
                    
                    # Mise à jour de la barre de progression
                    progress_bar.progress((classe_idx + 1) / total_classes)
                
                progress_bar.empty()
                
                # Stockage des résultats dans la session
                st.session_state.objets_salles = objets_salles
                st.session_state.non_places = non_places
                st.session_state.tentatives_backtrack = tentatives_backtrack
                st.session_state.repartition_completed = True
                
                st.success("✅ Répartition terminée!")
                
        # Affichage des résultats si la répartition est terminée
        if st.session_state.get('repartition_completed', False):
            objets_salles = st.session_state.objets_salles
            non_places = st.session_state.non_places
            tentatives_backtrack = st.session_state.tentatives_backtrack
            
            # Affichage des statistiques de l'algorithme
            if tentatives_backtrack > 0:
                st.info(f"🔄 Algorithme de backtracking utilisé {tentatives_backtrack} fois pour optimiser le placement")

            # Statistiques sur les contraintes relâchées
            total_contraintes_relachees = sum(getattr(salle, 'placements_avec_contraintes_relachees', 0) for salle in objets_salles)
            if total_contraintes_relachees > 0:
                st.info(f"⚡ {total_contraintes_relachees} étudiants placés avec contraintes relâchées (même matière côte à côte autorisée)")
                st.info("💡 Cela garantit que tous les étudiants sont placés, même si l'optimisation parfaite n'est pas possible.")

            # Affichage des statistiques de placement
            st.markdown("### 📊 Statistiques de Placement")
            
            cols = st.columns(len(objets_salles))
            for i, salle in enumerate(objets_salles):
                if salle.nombre_etudiants() > 0:
                    with cols[i % len(cols)]:
                        st.metric(
                            f"🏫 {salle.nom}",
                            f"{salle.nombre_etudiants()}/{salle.capacite_totale()}",
                            f"{salle.taux_remplissage():.1f}%"
                        )

            # Affichage de la répartition visuelle
            st.markdown("### 🗺️ Répartition Visuelle dans les Salles")
            couleurs_hex = ["#FF6B6B", "#4ECDC4", "#45B7D1", "#FFA07A", "#98D8C8", "#F7DC6F", "#BB8FCE", "#85C1E9"]
            epreuves_uniques = list(set(matieres_par_classe.values()))
            
            for salle in objets_salles:
                if salle.nombre_etudiants() < 1:
                    continue
                    
                with st.expander(f"🏫 Salle {salle.nom} ({salle.nombre_etudiants()}/{salle.capacite_totale()} - {salle.taux_remplissage():.1f}%)", expanded=True):
                    for rangée in ['gauche', 'milieu', 'droite']:
                        lignes = salle.rangées.get(rangée, [])
                        if not lignes:
                            continue
                        st.markdown(f"**Rangée {rangée.capitalize()}:**")
                        for i, ligne in enumerate(lignes):
                            if len(ligne) > 0:  # Vérifier que la ligne n'est pas vide
                                cols = st.columns(len(ligne))
                                for j, place in enumerate(ligne):
                                    with cols[j]:
                                        if place:
                                            nom, epreuve = place
                                            # Sécuriser l'accès à l'index des couleurs
                                            try:
                                                idx_couleur = epreuves_uniques.index(epreuve) % len(couleurs_hex)
                                                couleur = couleurs_hex[idx_couleur]
                                            except (ValueError, IndexError):
                                                couleur = "#808080"  # Couleur par défaut (gris)
                                            
                                            st.markdown(
                                                f'<div style="background-color: {couleur}; padding: 8px; border-radius: 5px; color: white; text-align: center; font-size: 11px; margin:  2px;">{nom[:10]}<br><small>({epreuve[:8]})</small></div>', 
                                                unsafe_allow_html=True
                                            )
                                        else:
                                            st.markdown(
                                                '<div style="background-color: #f0f0f0; padding: 8px; border-radius: 5px; text-align: center; font-size: 11px; color: #666; margin: 2px;">Vide</div>', 
                                                unsafe_allow_html=True
                                            )

            # Gestion des non-placés
            if non_places:
                st.error(f"⚠️ **{len(non_places)} étudiants n'ont pas pu être placés:**")
                for etu, mat in non_places:
                    st.write(f"❌ {etu} ({mat})")
                st.info("💡 Suggestion: Ajoutez plus de salles ou réduisez le nombre d'étudiants.")
            else:
                st.success("✅ Tous les étudiants ont été placés avec succès!")

    with tab4:
        if not st.session_state.get('repartition_completed', False):
            st.warning("⚠️ Veuillez d'abord effectuer la répartition dans l'onglet précédent.")
            return
            
        st.markdown("## 📄 Export PDF de la Répartition")
        
        # Récupération des données
        objets_salles = st.session_state.objets_salles
        matieres_par_classe = st.session_state.matieres_par_classe
        classes_choisies = st.session_state.classes_choisies
        
        # Information sur le nouveau format
        st.info("""
        🏫 **Format PDF - Layout Visuel des Salles (Optimisé):**
        - **Orientation paysage** pour une meilleure vue d'ensemble des salles
        - **Page de couverture** importée du fichier 20250130_Répartition_S1N.pdf
        - **Légende des couleurs** par matière sur une page dédiée
        - **Gestion optimisée des pages**: Plusieurs rangées par page quand possible
        - **Grille de placement compacte**: Tableau montrant l'arrangement physique des étudiants  
        - **Rangées séparées**: Gauche, Centre, Droite clairement distinctes
        - **Couleurs par matière**: Chaque épreuve a sa couleur distinctive
        - **Numérotation**: Lignes et colonnes numérotées pour faciliter le placement
        - **Layout adaptatif**: Ajustement automatique des tailles pour maximiser l'utilisation de l'espace
        """)

        # Aperçu des informations qui seront dans le PDF
        with st.expander("📋 Aperçu du contenu du PDF", expanded=True):
            st.write(f"  **Page de couverture**: Première page du fichier 20250130_Répartition_S1N.pdf")
            st.write(f"🎨 **Légende des couleurs**: Une légende pour les {len(set(matieres_par_classe.values()))} matières")
            st.write(f" 📅 **Semestre**: {semestre}")
            st.write(f"📅 **Date**: {date_epreuve.strftime('%d/%m/%Y')}")
            st.write(f"⏰ **Horaires**: {heure_debut.strftime('%H:%M')} - {heure_fin.strftime('%H:%M')}")
            st.write(f"🏫 **Salles utilisées**: {len([s for s in objets_salles if s.nombre_etudiants() > 0])}")
            st.write(f"👥 **Étudiants placés**: {sum(s.nombre_etudiants() for s in objets_salles)}")
            st.write(f"📚 **Classes concernées**: {len(classes_choisies)}")
            
            # Affichage détaillé par salle
            st.markdown("**🏫 Détail par salle:**")
            salles_avec_etudiants = [s for s in objets_salles if s.nombre_etudiants() > 0]
            for salle in salles_avec_etudiants:
                capacite = salle.capacite_totale()
                occupees = salle.nombre_etudiants()
                taux = (occupees/capacite)*100 if capacite > 0 else 0
                
                # Calculer le nombre de rangées pour cette salle
                rangees_utilisees = []
                for rangee_nom in ['gauche', 'milieu', 'droite']:
                    if rangee_nom in salle.rangées and salle.rangées[rangee_nom]:
                        rangee = salle.rangées[rangee_nom]
                        if rangee and not all(len(ligne) == 0 for ligne in rangee):
                            rangees_utilisees.append(rangee_nom)
                
                nb_rangees = len(rangees_utilisees)
                pages_estimees = max(1, (nb_rangees + 2) // 3)  # Estimation: jusqu'à 3 rangées par page
                
                st.write(f"• **{salle.nom}**: {occupees}/{capacite} places ({taux:.1f}%) → {nb_rangees} rangée(s) sur ~{pages_estimees} page(s)")
            
        # Génération et téléchargement du PDF
        if st.button("🏫 Générer le Plan des Salles (PDF)", type="primary"):
            with st.spinner("🏫 Génération du plan des salles en cours..."):
                try:
                    chemin_pdf = f"plan_salles_{semestre.replace(' ', '_')}_{date_epreuve.strftime('%Y%m%d')}.pdf"
                    
                    # Vérifier si le fichier de couverture existe
                    cover_file_exists = os.path.exists("20250130_Répartition_S1N.pdf")
                    
                    generer_pdf(
                        objets_salles,
                        semestre,
                        date_epreuve.strftime("%d/%m/%Y"),
                        heure_debut.strftime("%H:%M"),
                        heure_fin.strftime("%H:%M"),
                        matieres_par_classe,
                        chemin_pdf
                    )
                    
                    if cover_file_exists:
                        st.success(f"✅ Plan des salles généré avec succès, incluant la page de couverture et la légende des couleurs.")
                    else:
                        st.warning(f"⚠️ Plan des salles généré avec succès, mais le fichier de couverture '20250130_Répartition_S1N.pdf' n'a pas été trouvé.")
                    
                    # Bouton de téléchargement
                    with open(chemin_pdf, "rb") as file:
                        st.download_button(
                            label="⬇️ Télécharger le Plan des Salles",
                            data=file,
                            file_name=chemin_pdf,
                            mime="application/pdf",
                            type="primary"
                        )
                        
                    # Statistiques du PDF généré
                    taille_fichier = os.path.getsize(chemin_pdf) / 1024  # KB
                    salles_utilisees = len([s for s in objets_salles if s.nombre_etudiants() > 0])
                    nb_pages_estimees = salles_utilisees + 2  # 1 page couverture + 1 page infos/légende + salles
                    st.info(f"""📊 Détails du PDF:
                    - Taille du fichier: {taille_fichier:.1f} KB
                    - Pages: {nb_pages_estimees} au total
                      • 1 page de couverture (importée de 20250130_Répartition_S1N.pdf)
                      • 1 page d'informations avec légende des couleurs par matière
                      • {salles_utilisees} pages de plans de salles
                    """)
                    
                    # Ajouter un bouton pour démarrer une nouvelle session après le téléchargement
                    if st.button("🔄 Démarrer une nouvelle session", type="primary"):
                        # Réinitialiser l'état de la session
                        reset_session_state()
                        st.success("✅ Session réinitialisée! L'application va redémarrer...")
                        st.rerun()
                
                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération du PDF: {e}")
                    st.exception(e)

if __name__ == "__main__":
    try:
        main()
    finally:
        # Nettoyage des fichiers temporaires à la fin de l'exécution
        cleanup_temp_files()