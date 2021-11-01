# Outlook - Enregistrer les courriels sélectionnés au format PDF

![Banner](./banner.svg)

> Macro Outlook pour enregistrer les courriels en tant que fichiers PDF à l'emplacement sélectionné

## Description

Sélectionner un ou plusieurs courriels à partir du client Outlook, cliquer sur un bouton personnalisé du ruban et enregistrer-les dans un dossier spécifique.

Vous pouvez par exemple sélectionner 250 courriels et les enregistrer au format PDF en quelques clics.

## Table des Matières

- [Installation](#installation)
- [Usage](#usage)
- [License](#license)

## Installation

Obtenez une copie du code VBA `module.bas` et copiez-le dans votre client Outlook.

- Appuyez sur `ALT-F11` dans Outlook pour ouvrir la fenêtre `Visual Basic Editor` (alias VBE).
- Créez un nouveau module et copiez/collez le contenu du fichier `module.bas` que vous pouvez trouver dans ce dépôt.
- Fermer le VBE
- Cliquez avec le bouton droit de la souris sur le ruban Outlook pour le personnaliser et y ajouter un nouveau bouton. Attribuer la macro `SaveAsPDFfile` à ce bouton.

Note : Winword doit être installé sur votre ordinateur.

## Usage

1. Sélectionner un ou plusieurs courriels
2. Cliquez sur votre bouton `SaveAsPDFfile`.
3. Quelques fenêtres contextuelles s'affichent, vous demandant par exemple où stocker les courriers électroniques (sous forme de fichiers PDF) et si vous souhaitez ou non supprimer les courriers électroniques une fois qu'ils ont été sauvegardés sous forme de fichiers PDF.
4. C'est tout, attendez un peu et vous obtiendrez vos mails sauvegardés sur votre disque.

![](images/demo.gif)

## License

[MIT](LICENSE)
