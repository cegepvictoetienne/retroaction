# retroaction
Outil de rétroaction pour les travaux et examens. Prend un chiffrier Excel contenant les notes et la grille de correction et génère un fichier pdf par élève pour remettre dans Léa.  

## Installation  

`pip install -r requirements.txt`  

## Utilisation  

`python retroaction.py -i <fichier_retro> -o <dossier_sortie> -s <nom_feuille> -d <denominateur>`  

**-i** : Le chiffrier Excel contenant les rétroactions aux élèves. Chaque élément de la grille d'évaluation est en ligne et chaque élève est une colonne. Relatif au répertoire courant.  
**-o** : Le dossier dans lequel seront créés les pdf et l'archive zip.  Relatif au répertoire courant.  
**-s** : Le nom de la feuille contenant les rétroactions aux élèves.  
**-d** : Le dénominateur de la note de l'évaluation.  
