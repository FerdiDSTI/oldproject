# old project
Old funny project on vba,
Instructions in french :

1)	Module 1 : Il permet de quitter le Userform. C’est une macro qui s’active lorsqu’on clique sur le bouton « Quitter ».

2)	Module 2 : Il permet de réinitialiser le contenu des cellules. Il faut rentrer les valeurs de la dimension souhaitée afin de nettoyer la zone souhaitée. Il utilise la fonction .clear donc efface tout, c’est à dire le rectangle et les obstacles générés. Ce processus s’active lorsqu’on clique sur « Réinitialiser les données ». 

3)	Module 3 : Il va créer un rectangle dans lequel nous allons travailler. Si l’on crée un rectangle de dimension 20x20, la zone de déplacement de la marche aléatoire sera en réalité 18x18. Pour créer le rectangle, il suffit de donner le nombre de ligne et de colonnes. Nous avons ajouté une fonction de zoom qui adapte l’échelle en fonction de la taille du rectangle, néanmoins si on veut travailler sur des rectangles de grandes dimensions (supérieures à 100x100), cette option de zoom doit être retravaillée. Ce processus s’active lorsqu’on clique sur « Créer un rectangle ». 

4)	Module 4 : Nous avons décidé de garder ce module afin de vous montrer comment nous avions imaginé le processus de la marche aléatoire au départ. Dans un 1er temps, une cellule est sélectionnée aléatoirement grâce à la fonction Rnd vu en cours. Lors de la sélection de cette cellule de départ nous avons eu pour contrainte de nous assurer que cette cellule était bien dans le rectangle. Dans un 2nd temps, nous avons utilisé une boucle For afin de réitérer le processus le nombre indiqué de fois par l’utilisateur. Ce processus aléatoire suit une loi uniforme de probabilité ¼. Nous avons également posé des conditions pour que le déplacement se limite à l’intérieur du rectangle: 
-	Si on est à côté d’un bord, nous passons de 1 chance sur 4 d’aller sur une des cases adjacentes à 1 chance sur 3.
-	Si on est dans un coin, on passe à 1 chance sur 2.
Enfin, nous avons utilisé une fonction permettant de colorer chaque cellule sélectionnée.

5)	Module 5 : Il reprend le concept du module 4 mais l’approche est totalement différente. Le déplacement de la cellule est maintenant contraint par les cases noires délimitant le rectangle et non plus par les coordonnées des bordures. Cette flexibilité nous permet d’intégrer des obstacles à l’intérieur du rectangle.
De plus nous avons remplacé l’ancienne boucle For par une boucle While, boucle qui s’arrêtera une fois le tableau totalement remplit. 
Le processus de sélection de la 1ère case reste le même. 
Plus précisément, ici, au lieu de poser des conditions sur les bords, nous avons pensé qu’il fallait simplement vérifier la validité de la case choisie aléatoirement, c’est à dire vérifier si cette case est déjà coloriée en noir : Si la cellule est de couleur noire alors on revient à la précédente, sinon la cellule est coloriée en rouge. Nous avons ajouté une variable ayant un rôle de compteur d’itération qui nous permet de faire des statistiques simples.
Enfin, nous avons ajouté un Msgbox indiquant le nombre d’itérations nécessaires pour colorier l’intégralité du tableau. Ce module s’active lorsqu’on clique sur le bouton « Marche aléatoire complète » et sur le bouton « Go ! ». C’est le cœur de notre projet.

6)	Module 6 : Il permet d’effacer le contenu du rectangle sans toucher aux obstacles. Il parcourt l’intégralité du rectangle et efface toutes les cases rouges qui sont à l’intérieur. Il s’active en cliquant sur « Effacer le contenu ».

7)	Module 7 : Il fait apparaître des obstacles aléatoirement. Nous avons utilisé la fonction Rnd comme précédemment et ici encore nous avons eu pour contrainte de les faire apparaître à l’intérieur du rectangle. Chaque obstacle mesure 7 cases précisément. Le module choisi une case aléatoirement dans le rectangle en faisant attention de ne pas se trouver coller à un bord, puis va générer des cases noires aléatoirement autour de lui. Une fois qu’il a colorié 7 cases en noires, il s’arrête. Ce module s’active en cliquant sur « Créer un obstacle ».

8)	Module 8 : Il permet d’effacer le contenu des cellules de la feuille et de faire apparaître le UserForm. C’est la fonction Start.

9)	Module 9 : Il permet de faire apparaître un histogramme, nous avons utilisé pour cela la commande « Enregistreur de macro » permettant de coder les actions que nous effectuons.

