# Projet Répartition Heuristique ISSEA

Ce projet permet de répartir automatiquement les étudiants dans les salles d'examen de l'ISSEA en respectant des contraintes d'adjacence de matières et d'occupation des salles, avec une interface Streamlit et une exportation PDF visuelle.

## Fonctionnalités principales

- **Répartition automatique** des étudiants dans les salles selon les matières et les contraintes.
- **Algorithme heuristique** avec backtracking pour garantir le placement de tous les étudiants.
- **Ajustement dynamique** : si les contraintes strictes ne suffisent pas, elles sont relâchées pour garantir le placement de tous.
- **Export PDF** : Génération d'un plan de salle visuel, optimisé pour la lisibilité (cellules, polices et espacements adaptés).
- **Interface utilisateur Streamlit** : chargement des fichiers Excel, configuration, visualisation et export.
- **Vidéo démo** incluse ci-dessous.

## Utilisation

1. Lancez l'application Streamlit avec :
   ```bash
   streamlit run TresBon_code3.py
   ```
2. Chargez les fichiers Excel demandés (étudiants, matières, salles).
3. Configurez les classes et salles à utiliser.
4. Visualisez la répartition et exportez le PDF.

## Dépendances principales

- Python 3.8+
- streamlit
- pandas
- reportlab
- PyPDF2

Installez les dépendances avec :

```bash
pip install streamlit pandas reportlab PyPDF2
```

## Démo vidéo

[![Voir la démo](ApercuDocument.mp4)](ApercuDocument.mp4)

Ou ouvrez directement le fichier `ApercuDocument.mp4` dans ce dépôt.

---

**Auteur :** Laurent Joel

Projet ISSEA 2025
