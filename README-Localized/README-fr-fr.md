# Complément de rédaction JavaScript pour Word

Découvrez comment créer un complément qui recherche, met en surbrillance et rédige du texte.    

## Table des matières
* [Historique des modifications](#historique-des-modifications)
* [Remarque importante](#remarque-importante)
* [Conditions requises](#conditions-requises)
* [Configurer le projet](#configurer-le-projet)
* [Exécuter le projet](#exécuter-le-projet)
* [Questions et commentaires](#questions-et-commentaires)
* [Ressources supplémentaires](#ressources-supplémentaires)


## Historique des modifications

25 juillet 2016 :
* Version de l’exemple initial.

##Remarque importante

Cet exemple rédige du texte en remplaçant chaque lettre par un astérisque.  La disposition n’est pas conservée telle quelle.  Aucune mesure ni aucun calcul n’est effectué pour déterminer la police, la taille ou la mise en forme de chaque lettre.

## Conditions requises

* Word 2016 pour Windows, build 16.0.6912.1000 ou ultérieur.
* [Node et npm](https://nodejs.org/en/)
* [GIT Bash](https://git-scm.com/downloads) - Vous devez utiliser une version ultérieure, car les builds antérieurs peuvent afficher une erreur lors de la génération de certificats.

## Configurer le projet

À partir de votre environnement Bash, exécutez les commandes suivantes à la racine de ce projet :

1. Dupliquez ce référentiel sur votre ordinateur local.
2. ```npm install``` pour installer toutes les dépendances dans package.json.
3. ```bash gen-cert.sh``` pour créer les certificats nécessaires à l’exécution de cet exemple. 
* Ensuite, dans le référentiel de votre ordinateur local, cliquez deux fois sur ca.crt et sélectionnez **Installer le certificat**. 
* Sélectionnez **Ordinateur local** et choisissez **Suivant** pour continuer. 
* Sélectionnez **Placer tous les certificats dans le magasin suivant**, puis **Parcourir**.  
* Sélectionnez **Autorités de certification racines de confiance** et **OK**. 
* Sélectionnez **Suivant** et **Terminer**. Désormais, votre certificat d’autorité de certification apparaît dans votre magasin de certificats.
4. ```npm start`
            Lancez `` pour démarrer le service de serveur web du nœud.

Vous devez maintenant indiquer à Microsoft Word où trouver le complément.

1. Créez un partage réseau ou [partagez un dossier sur le réseau](https://technet.microsoft.com/fr-fr/library/cc770880.aspx), puis placez-y le fichier manifeste [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml).
3. Lancez Word et ouvrez un document.
4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
5. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
6. Choisissez **Catalogues de compléments approuvés**.
7. Dans le champ **URL du catalogue**, saisissez le chemin réseau pour accéder au partage de dossier qui contient le fichier word-add-in-js-redact-manifest.xml, puis choisissez **Ajouter un catalogue**.
8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.
9. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage de Microsoft Office. Fermez et redémarrez Word.

## Exécuter le projet

1. Ouvrez un document Word.
2. Dans l’onglet **Insertion** de Word 2016, choisissez **Mes compléments**.
3. Sélectionnez l’onglet **DOSSIER PARTAGÉ**.
4. Choisissez **Complément de rédaction Word**, puis sélectionnez **OK**.
5. Si les commandes de complément sont prises en charge par votre version de Word, l’interface utilisateur vous informe que le complément a été chargé.

### Interface utilisateur du ruban

Effectuez les actions suivantes sur le ruban :
* Sélectionnez l’onglet **Révision** et **Afficher le volet Office de rédaction** pour lancer le volet Office.

 > Remarque : Le complément se charge dans un volet Office si les commandes de complément ne sont pas prises en charge par votre version de Word.

### Interface utilisateur de volet Office

Dans le volet Office, vous pouvez :
* Recherchez et mettez en surbrillance du texte dans un document en saisissant un mot dans la zone de texte et en cliquant sur le bouton **Rechercher et mettre en surbrillance**.
  
> REMARQUE :  la recherche fait la distinction entre les minuscules et les majuscules.  Pour annuler une action, appuyez sur les touches Ctrl + Z de votre clavier.

* Rédigez du texte dans un document en saisissant un mot dans la zone de texte, puis en cliquant sur le bouton **Rédiger**.
  
> REMARQUE :  la rédaction fait la distinction entre les minuscules et les majuscules.   

* Rédigez un texte sélectionné dans un document en cliquant sur le bouton **Rédiger le texte sélectionné**.
  
> REMARQUE :  la rédaction fait la distinction entre les minuscules et les majuscules.       
  
### Commandes de compléments

Cet exemple montre comment créer des commandes de complément dans le ruban. Les déclarations de commande de complément se trouvent dans word-add-in-js-redact-manifest.xml. 

## Questions et commentaires

Nous serions ravis de connaître votre opinion sur l’exemple de complément de rédaction Word. Vous pouvez nous faire part de vos suggestions dans la rubrique *Problèmes* de ce référentiel.

Si vous avez des questions générales sur le développement de Microsoft Office 365, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Posez vos questions avec les balises [API] et [office-js].

## Ressources supplémentaires

* [Documentation de complément Office](https://msdn.microsoft.com/fr-fr/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* [Projets de démarrage et exemples de code des API Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Copyright
Copyright (c) 2016 Microsoft Corporation. Tous droits réservés.


