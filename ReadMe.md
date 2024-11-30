# Projet de Configuration des Routeurs

## Description
Ce projet est une application qui aide les techniciens réseaux à configurer rapidement les routeurs après avoir réalisé une conception schématique du réseau selon les besoins spécifiques. L'application permet de générer trois fichiers contenant :

- **Configuration des interfaces des routeurs** : Réglages des adresses IP, des sous-réseaux, et autres paramètres des interfaces.
- **Configuration du routage via le protocole RIP** : Mise en place du routage dynamique à l'aide du protocole RIP (Routing Information Protocol).
- **Configuration DHCP** : Paramétrage du serveur DHCP sur les routeurs pour attribuer des adresses IP aux machines automatiquement.

En complément, trois fichiers Excel sont générés, offrant davantage de détails sur le routage et les sous-réseaux.

---

## Fonctionnalités
- **Génération automatique de fichiers texte pour la configuration des routeurs** :
  - `interfaces_config.txt` : Configuration des interfaces des routeurs.
  - `rip_routing_config.txt` : Configuration du routage RIP.
  - `dhcp_config.txt` : Configuration du serveur DHCP.

- **Génération de fichiers Excel pour les détails réseau** :
  - `Table_adressage.xlsx`
  - `Table_de_routage.xlsx`
  - `Table_des_sous_reseaux.xlsx`

- **Personnalisation par l'utilisateur** :
  - Spécification du nombre de réseaux LAN et WAN.
  - Définition du nombre de routeurs à configurer.
  - Sélection des interfaces des routeurs.

- **Interface graphique conviviale (Tkinter)** :
  - Sélection des interfaces réseau pour chaque routeur directement depuis l'application.
  - Fenêtre principale avec navigation par défilement pour accéder aux différentes options.

---

## Installation

### Prérequis
- Python 3.x
- Bibliothèque Tkinter

---

## Utilisation
1. **Configurer le chemin** :
   - Modifiez la variable `base_path` dans le fichier `script.py` pour indiquer le chemin local vers le répertoire `result` où les fichiers seront stockés.

2. **Lancer l'application** :
   - Exécutez le script pour ouvrir l'interface graphique.

3. **Configurer les paramètres** :
   - Indiquez le nombre de réseaux (LANs et WANs) et le nombre de routeurs.
   - Sélectionnez les interfaces réseau pour chaque routeur via l'interface graphique.

4. **Générer la configuration** :
   - Cliquez sur "Générer la configuration" pour produire les fichiers suivants :
     - `interfaces_config.txt` : Configuration des interfaces des routeurs.
     - `rip_routing_config.txt` : Configuration du routage RIP.
     - `dhcp_config.txt` : Configuration du serveur DHCP.

   - Les fichiers Excel suivants seront également générés :
     - `Table_adressage.xlsx`
     - `Table_de_routage.xlsx`
     - `Table_des_sous_reseaux.xlsx`

---

## NB
Pour garantir le bon fonctionnement de l'application :
- **Ne pas supprimer les fichiers Excel**. Ils sont nécessaires pour le traitement des configurations.
- Les fichiers texte peuvent être supprimés si besoin, mais cela pourrait nécessiter une régénération si les configurations sont modifiées.
