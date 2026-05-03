# 📦 إدارة جرد المكاتب (Office Inventory Management)

Une application Streamlit interactive pour la gestion et le contrôle de l'inventaire des équipements par bureau.

## 🌟 Fonctionnalités

- **Filtrage par Bureau** : Sélectionnez facilement un numéro de bureau pour afficher uniquement ses équipements.
- **Édition Interactive** : Modifiez, ajoutez ou supprimez des équipements directement depuis une interface sous forme de tableau dynamique.
- **Exportation PDF** : Générez des fiches d'inventaire en format PDF avec des paramètres personnalisables (largeur des colonnes, hauteur des lignes, orientation portrait/paysage).
- **Exportation Excel** : Téléchargez les données filtrées et modifiées d'un bureau au format Excel.
- **Sauvegarde Automatique** : Enregistrez toutes vos modifications en un seul clic directement dans le fichier principal `INVENTAIRE EN ORDRE.xlsx`.
- **Avertissements Visuels** : Mise en évidence automatique (en rouge transparent) des anomalies détectées (lorsque la colonne `COMPARAISON` contient le mot 'faux').
- **Optimisation Mobile** : Interface adaptative avec CSS personnalisé pour une expérience de saisie optimale sur smartphone (boutons larges, marges réduites, optimisation du clavier virtuel).

## 🛠️ Prérequis

Assurez-vous d'avoir Python installé ainsi que les dépendances suivantes :

```bash
pip install streamlit pandas openpyxl fpdf
```

## 🚀 Lancement de l'application

1. Placez votre fichier source `INVENTAIRE EN ORDRE.xlsx` dans le même dossier que l'application.
2. Ouvrez un terminal dans le dossier du projet.
3. Lancez l'application avec la commande suivante :

```bash
streamlit run app.py
```
*(Note : Si vous utilisez un environnement spécifique comme Miniconda, la commande pourrait ressembler à `python -m streamlit run app.py`)*

4. L'application s'ouvrira dans votre navigateur web (généralement à l'adresse `http://localhost:8501`). Vous pouvez également y accéder via votre smartphone connecté au même réseau en utilisant l'adresse IP affichée dans le terminal.

## 📱 Utilisation sur Smartphone

L'application intègre des règles spécifiques pour l'affichage sur smartphone :
- L'en-tête et le pied de page par défaut de Streamlit sont cachés pour maximiser l'espace.
- Les boutons prennent toute la largeur pour faciliter la navigation tactile.
- La taille du texte est réduite et la hauteur du tableau de données est plafonnée afin que le clavier virtuel du téléphone n'entrave pas le défilement.
- Certaines colonnes non essentielles (`N° BUREAU`, `DEP`, `OCC`) sont masquées par défaut pendant l'édition.

## ⚙️ Configuration du PDF

Lors de l'exportation PDF, vous pouvez ouvrir la section "إعدادات ملف PDF (خيارات الطباعة)" pour configurer manuellement :
- L'orientation de la page (Paysage ou Portrait).
- La largeur allouée à chaque colonne spécifique (`CODE COMPTBLE`, `OLDCODE2`, `DESIGNATION`, `OBSERVATION`).
- La hauteur de chaque ligne pour compresser l'affichage et insérer plus d'articles sur une seule page.
