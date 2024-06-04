# ANGRYBIRDS.SCPT
--------------------------------------------------------------------------------------------------
This file is required for the Angry Birds trivia PPT to work correctly on MacOS computers. However, because of Apple's security policies, the file needs to be installed in a directory that is usually hidden from most users. The good news, though, is that it is very easy to move the file into the correct directory.

DOWNLOADING & INSTALLATION
--------------------------------------------------------------------------------------------------
Automated Method
Open a Terminal window. You can press Command(⌘) + Spacebar to open Spotlight search and type "Terminal". Then, you can press return to open it. If you prefer to use Finder, you can press Command(⌘) + Shift + U to open the Utilities folder. Double-click on Terminal to open it. With a Terminal window open, copy and paster the following command and press return.

curl -L -o ~/Library/Application\ Scripts/com.microsoft.Powerpoint/AngryBirds.scpt https://github.com/papercutter0324/AngryBirdsTrivia-AdditionalFiles/raw/main/AngryBirds.scpt

Manual Method
You need to manually download AngryBirds.scpt to your Downloads folder. There are three ways to easily do this.
   1 - Right-click on AngryBirds.scpt above, select 'Save As', and save it to your Downloads folder.
   2 - Click on AngryBirds.scpt, which will take you to a new page. On the right, you should see a small button labeled 'Raw'. Click on it and save the file to your Downloads folder.
   3 - On the right on this page is the Releases section. Click on either "Releases" or the most recent release (you should see a green tag and a 'Latest' label next to it). On the new page, click on and download "Source code (zip)". You mau need to click on 'Assets' to see it. Once downloaded, open the zip file and extract AngryBirds.scpt to your Downloads folder.

Next, open a Terminal window (using one of the methods mentioned in the "Automated Method" section) and run the following command.

mv ~/Downloads/AngryBirds.scpt ~/Library/Application\ Scripts/com.microsoft.Powerpoint

ERRORS
--------------------------------------------------------------------------------------------------
In some rare situations, you might receive an error saying ~/Library/Application Scipts/com.microsoft.Powerpoint or /Users/yourUserName/Library/Application Scripts/com.microsoft.Powerpoint doesn't exist, and you will need to create the directory. To do so, run the following command. Afterward, following the installation step again. (You might need to redownload the file.)

Command: (You can simple copy and paste the command into the Terminal.)

mkdir ~/Library/Application\ Scripts/com.microsoft.Powerpoint

VERIFYING INSTALLATION
--------------------------------------------------------------------------------------------------
If you were able to complete the above steps, you can verify that the file has been correctly installed by opening up the Angry Birds trivia PPT and choosing any of the title screen options. For simplicity, I recommend either "DEFAULTS" or "OPTIONS". If you transition to the next slide without seeing a popup message, then the file has been successfully installed.
