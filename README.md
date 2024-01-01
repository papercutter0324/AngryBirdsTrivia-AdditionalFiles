ANGRYBIRDS.SCPT

This file is required for the Angry Birds trivia PPT to work correctly on MacOS computers. However, because of Apple's security policies, the file needs to be installed in a directory that is usually hidden from most users. The good news, though, is that it is very easy to move the file into the correct directory.

INSTALLATION

First, you you need to download AngryBirds.scpt to your Downloads folder. When the download window opens to select where to save it, it might show the filename as only "AngryBirds". This is normal, so you don't need to add the ".scpt" ending. After saving AngryBirds.scpt to your Downloads folder, open up the Terminal and run command shown below. To open the terminal, go into your Launchpad and search for it. By the time you have typed "ter", it should be one of the only options shown.

Command: (You can simple copy and paste the command into the Terminal.)

mv ~/Downloads/AngryBirds.scpt ~/Library/Application\ Scripts/com.microsoft.Powerpoint

ERRORS:

In some rare situations, you might receive an error saying ~/Library/Application Scipts/com.microsoft.Powerpoint or /Users/yourUserName/Library/Application Scripts/com.microsoft.Powerpoint doesn't exist, and you will need to create the directory. To do so, run the following command. Afterward, following the installation step again. (You might need to redownload the file.)

Command: (You can simple copy and paste the command into the Terminal.)

mkdir ~/Library/Application\ Scripts/com.microsoft.Powerpoint

VERIFYING INSTALLATION

If you were able to complete the above steps, you can verify that the file has been correctly installed by opening up the Angry Birds trivia PPT and choosing any of the title screen options. For simplicity, I recommend either "DEFAULTS" or "OPTIONS". If you transition to the next slide without seeing a popup message, then the file has been successfully installed.
