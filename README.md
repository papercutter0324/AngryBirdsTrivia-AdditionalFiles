ANGRYBIRDS.SCPT

This file is required for the Angry Birds trivia PPT to work correctly on MacOS computers. However, because of Apple's security policies, the file needs to be installed in a directory that is usually hidden from most users. The good news, though, is that it is very easy to move the file into the correct directory.

INSTALLATION

First, you you need to download AngryBirds.scpt to your Downloads folder. When the download window opens to select where to save it, it might show the filename as only "AngryBirds". This is normal, so you don't need to add the ".scpt" ending. After saving AngryBirds.scpt to your Downloads folder, open up the Terminal and run command shown below. To open the terminal, go into your Launchpad and search for it. By the time you have typed "ter", it should be one of the only options shown.

Command: (You can simple copy and paste the command into the Terminal.)

mv ~/Downloads/AngryBirds.scpt ~/Library/Application\ Scripts/com.microsoft.Powerpoint

Errors:

In some rare situations, you might receive an error saying ~/Library/Application Scipts/com.microsoft.Powerpoint or /Users/yourUserName/Library/Application Scripts/com.microsoft.Powerpoint doesn't exist, and you will need to create the directory. To do so, run the following command.

Command: (You can simple copy and paste the command into the Terminal.)

mkdir ~/Library/Application\ Scripts/com.microsoft.Powerpoint

VERIFYING INSTALLATION

If you
