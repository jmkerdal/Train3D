# Train3D
 Miniature train simulator 3D model HO

Je mets en ligne la version en VB6 qui m'a servi à créer ce projet.

Quelques parties dans le code sont intéressantes à étudier (de mon avis) :
- Le système de création des éléments de rails. J'ai utilisé une matrice de connexion de segment de droites et de courbes. Gère les aiguillages et permet de l’emprunter dans n'importe quel sens.
- La création de textures 2D en mémoire pour afficher les voies à partir de la matrice. L'idée est de partir sur un rendue 3D et de capturer une projection 2D isométrique des couches horizontales. 
- Le système de connexion des voies afin de créer un réseau fermé.
- Le rendue 3D avec les élévations, les tunnels et le placement de caténaires.
- Un système de création de train de wagons qui calcule la position des bogies sur lesquels les caisses sont posées.

L'outil "Wall3D" qui me servais à concevoir des modèles 3D simples avec des textures transparentes. Ce que ne permettait pas les fichiers .x de Microsoft de l'époque.

Contact : jeanmichel.kerdal@free.fr
