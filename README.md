# PC Journal is a simple application made with python that tracks the on/off events of a computer. It logs these events and allows users to export the data to Excel or PDF. The app also includes auto-boot functionality.  
The core idea of this project is to figure out when the PC was shut down, even though we can’t record the shutdown event directly.  
Since Python can’t detect a PC shutting down (the script gets killed), we calculate the shutdown time the next time the PC is turned on, using data saved from the last session.
  
 Features:  
   
✅ Tracks PC ON/OFF sessions

✅ Export logs to Excel

✅  Export logs to PDF

✅  View full log history

✅  Auto-start with Windows (optional)

✅  Session duration tracking in the background

  Try it yourself :  
  1. Clone the project
      git clone https://github.com/DmitriyCh2104/PC_on-off_log.git   
     cd PC_on-off_log      
 2. Install dependencies  
      pip install fpdf pandas pystray pillow pywin32 winshell  
 3. Compile PC_Journal.py in destination folder  
      pyinstaller --onefile PC_Journal.py

# PC Journal est une application simple faite avec python qui suit les événements de démarrage et d'arrêt d'un ordinateur. Elle enregistre ces événements et permet aux utilisateurs d'exporter les données vers Excel ou PDF. L'application intègre également une fonctionnalité de démarrage automatique.
L'idée principale de ce projet est de déterminer quand le PC a été éteint, même si nous ne pouvons pas enregistrer directement cet événement.  
Comme Python ne détecte pas l'arrêt d'un PC (le script est alors interrompu), nous calculons l'heure d'arrêt au prochain démarrage du PC, en utilisant les données enregistrées lors de la dernière session.  
  
Fonctionnalités:
  
✅ Suivi des sessions PC allumées/éteintes

✅ Exportation des journaux vers Excel

✅ Exportation des journaux vers PDF

✅ Affichage de l'historique complet

✅ Démarrage automatique avec Windows (facultatif)

✅ Suivi de la durée des sessions en arrière-plan

Essayez par vous-même :

1. Clonez le projet  
  git clone https://github.com/DmitriyCh2104/PC_on-off_log.git  
  cd PC_on-off_log  
2. Installez les dépendances  
  pip install fpdf pandas pystray pillow pywin32 winshell  
3. Compilez PC_Journal.py dans le dossier de destination  
  pyinstaller --onefile PC_Journal.py  
