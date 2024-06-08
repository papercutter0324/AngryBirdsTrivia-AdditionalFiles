Option Explicit
Option Base 1
Const debugEnabled As Boolean = True

'''''''''''''''''''''''''''''''''''''''''''''''''''
'        GLOBAL CONFIGURATION VARIABLES           '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public numberOfTeams_BorderedArray() As Variant, numberOfTeams_BorderlessArray() As Variant
Public numberOfLotteryTiles_BorderedArray() As Variant, numberOfLotteryTiles_BorderlessArray() As Variant
Public classTrackerDays_SelectedArray() As Variant, classTrackerDays_UnselectedArray() As Variant
Public classTrackerTimes_SelectedArray() As Variant, classTrackerTimes_UnselectedArray() As Variant
Public enabledGrade_SelectedArray() As Variant, enabledGrade_UnselectedArray() As Variant
Public enabledLevel_SelectedArray() As Variant, enabledLevel_UnselectedArray() As Variant
Public enabledBook_SelectedArray() As Variant, enabledBook_UnselectedArray() As Variant
Public enabledUnit_SelectedArray() As Variant, enabledUnit_UnselectedArray() As Variant

Public configSld As Integer, configSld2 As Integer, numberOfTeams As Integer, numberOfLotteryTiles As Integer
Public allowNegatives As Boolean, autoAddPoints As Boolean
Public enableTrivia As Boolean, enableGrammar As Boolean, enableReview As Boolean, enableChallenges As Boolean
Public enabledGrade As String, enabledLevel As String, enabledBook As String, enabledUnit As Integer
Public classDay As String, classTime As String

'''''''''''''''''''''''''''''''''''''''''''''''''''
'           OS & ENVIRONMENT VARIABLES            '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public pathSeparator As String, tempFolder As String, configFolder As String
Dim fileAccessGranted As Boolean
Public osIsWindows As Boolean
Public quotationMark As String

'''''''''''''''''''''''''''''''''''''''''''''''''''
'           GLOBAL GAMEPLAY VARIABLES             '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public alreadySeenSlidesArray() As Integer, previouslySeenQuestionsArray() As String, randomlySelectedLotteryTilesArray() As String
Public variableNamesArray() As String, variableValuesArray() As Variant
Public currentStateNamesArray() As String, currentStateValuesArray() As Variant

Public gameboardSlide As Integer, lotterySlide As Integer, rulesSlide As Integer
Public firstQuestionSlide As Integer, lastQuestionSlide As Integer, firstLotterySlide As Integer, lastLotterySlide As Integer
Public birdsPointsOne As Integer, birdsPointsTwo As Integer, birdsPointsThree As Integer, pigsPointsOne As Integer, pigsPointsTwo As Integer, pigsPointsThree As Integer
Public chosenTile As String, chosenLottery As String, chosenSlide As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''
'         GLOBAL SCOREBOARD VARIABLES             '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public team1Score As Integer, team2Score As Integer, team3Score As Integer, team4Score As Integer, marmosetsScore As Integer
Public swapFirstTeam As String, swapSecondTeam As String, swapInProgress As Boolean
Public doublePoints As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''
'               DEFAULT VARIABLES                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitializeDefaultValues()
    If debugEnabled Then
        Debug.Print vbCrLf & "Initializing default game values."
    End If
    
    quotationMark = """"
    numberOfTeams = 2
    numberOfLotteryTiles = 8
    autoAddPoints = True
    allowNegatives = True
    enableTrivia = True
    enableGrammar = False
    enableReview = False
    enableChallenges = False
    classTime = "NoTime"
    classDay = "NoDay"
    enabledGrade = ""
    enabledLevel = ""
    enabledBook = ""
    enabledUnit = 0
    chosenTile = ""
    chosenLottery = ""
    chosenSlide = 0
    team1Score = 0
    team2Score = 0
    team3Score = 0
    team4Score = 0
    marmosetsScore = 0
    
    ReDim previouslySeenQuestionsArray(1)
    previouslySeenQuestionsArray(1) = ""
    
    If debugEnabled Then
        Debug.Print "   quotationMark: " & quotationMark
        Debug.Print "   numberOfTeams: " & numberOfTeams
        Debug.Print "   numberOfLotteryTiles: " & numberOfLotteryTiles
        Debug.Print "   autoAddPoints: " & autoAddPoints
        Debug.Print "   allowNegatives: " & allowNegatives
        Debug.Print "   enableTrivia: " & enableTrivia
        Debug.Print "   enableGrammar: " & enableGrammar
        Debug.Print "   enableReview: " & enableReview
        Debug.Print "   enableChallenges: " & enableChallenges
        Debug.Print "   classTime: " & classTime
        Debug.Print "   classDay: " & classDay
        Debug.Print "   enabledGrade: " & enabledGrade
        Debug.Print "   enabledLevel: " & enabledLevel
        Debug.Print "   enabledBook: " & enabledBook
        Debug.Print "   enabledUnit: " & enabledUnit
        Debug.Print "   chosenTile: " & chosenTile
        Debug.Print "   chosenLottery: " & chosenLottery
        Debug.Print "   chosenSlide: " & chosenSlide
        Debug.Print "   team1Score: " & team1Score
        Debug.Print "   team2Score: " & team2Score
        Debug.Print "   team3Score: " & team3Score
        Debug.Print "   team4Score: " & team4Score
        Debug.Print "   marmosetsScore: " & marmosetsScore
        
        If IsEmpty(previouslySeenQuestionsArray) Then
            Debug.Print "   previouslySeenQuestionsArray size: Empty"
        Else
            Debug.Print "   previouslySeenQuestionsArray size: " & IIf(UBound(previouslySeenQuestionsArray) <> LBound(previouslySeenQuestionsArray), (UBound(previouslySeenQuestionsArray) - LBound(previouslySeenQuestionsArray)), UBound(previouslySeenQuestionsArray))
            Debug.Print "   previouslySeenQuestionsArray(" & UBound(previouslySeenQuestionsArray) & "): " & previouslySeenQuestionsArray(UBound(previouslySeenQuestionsArray))
        End If
        
        Debug.Print "Default values set."
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'                TITLE SCREEN
'''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TitleScreen(clickedShp As Shape)
    Dim chosenConfigFile As String
    
    #If Mac Then
        CheckForAppleScript
    #End If

    SetOSVariables
    SetFolders
    ResetToDefaultState
    CheckForUpdates
    InitializeDefaultValues

    Select Case clickedShp.Name
        Case "LoadConfig"
            chosenConfigFile = ChooseFileToLoad()
            
            If chosenConfigFile <> "" Then
                #If Mac Then
                    LoadConfigMac "ExistingConfig", chosenConfigFile
                #Else
                    Dim fs As Object, configFileToLoad As Object
                    
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set configFileToLoad = fs.OpenTextFile(chosenConfigFile, 1, False, -1)
                    
                    LoadConfigWindows "ExistingConfig", configFileToLoad
                #End If
            End If
        Case "RestoreGame"
            RestoreGame
        Case "UseDefaults"
            StartGameWithDefaults
        Case "CreateConfig"
            PrepareOptionsMenuStep1
    End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'       SOFTWARE REQUIREMENTS & DEPENDENCIES
'''''''''''''''''''''''''''''''''''''''''''''''''''
#If Mac Then
Private Sub CheckForAppleScript()
        Dim appleScriptPath As String, isScriptInstalled As String, messageToDisplay As String
        Dim userChoice As Integer
        
        appleScriptPath = MacScript("return POSIX path of (path to desktop folder) as string")
        appleScriptPath = Replace(appleScriptPath, "/Desktop", "") & "Library/Application Scripts/com.microsoft.Powerpoint/AngryBirds.scpt"
        
        If debugEnabled Then
            Debug.Print "Verifying that AngryBirds.scpt is correctly installed."
            Debug.Print "Script directory set to: " & appleScriptPath
        End If

        On Error Resume Next
        isScriptInstalled = Dir(appleScriptPath, vbDirectory)
        
        If debugEnabled Then
            Debug.Print IIf(Not (isScriptInstalled = vbNullString), "AngryBirds.scpt found. Now checking if there is a newer version available.", "AngryBirds.scpt is missing. Please follow the instructions for how to install it.")
        End If
        
        If Not (isScriptInstalled = vbNullString) Then
            On Error GoTo 0
            CheckForScriptUpdate
            Exit Sub
        End If
        
        On Error GoTo 0
        messageToDisplay = "Warning! The AngryBirds.scpt file is not correctly installed. The game will not function correctly without it. Would you like to download the latest version now? If no, click 'No' to close the game."
        DisplayMessage messageToDisplay, "AppleScriptNotFound", userChoice
        
        If userChoice = vbYes Then
            ActivePresentation.FollowHyperlink Address:="https://github.com/papercutter0324/applescripts/tree/main"
            
            If debugEnabled Then
                Debug.Print "Opening page to download and install AngryBirds.sctp."
            End If
        End If
        
        ActivePresentation.SlideShowWindow.View.Exit
End Sub
#End If

#If Mac Then
#Else
Private Function CheckForCurl() As Boolean
    Dim objShell As Object, objExec As Object
    Dim output As String, messageToDisplay As String
    Dim checkResult As Boolean
    
    If debugEnabled Then
        Debug.Print "Checking if curl.exe is installed."
    End If
    
    'Create a shell object
    Set objShell = CreateObject("WScript.Shell")
    
    'Attempt to run curl and capture the output
    On Error Resume Next
    Set objExec = objShell.exec("cmd /c curl.exe --version")
    On Error GoTo 0
    
    If objExec Is Nothing Then
        If debugEnabled Then
            Debug.Print "   Not installed. Falling back to .Net."
        End If
        
        Set objExec = Nothing
        Set objShell = Nothing
        
        CheckForCurl = False
    Else
        Do While Not objExec.StdOut.AtEndOfStream
            output = output & objExec.StdOut.ReadLine() & vbCrLf
        Loop
        
        checkResult = ((InStr(output, "curl")) > 0)
        
        If debugEnabled Then
            Debug.Print IIf(checkResult, "   Installed.", "   Not installed. Falling back to .Net.")
        End If
        
        Set objExec = Nothing
        Set objShell = Nothing
        
        CheckForCurl = checkResult
    End If
End Function
#End If

Private Function CheckForExcel() As Boolean
    If debugEnabled Then
        Debug.Print "Verifying that Microsoft Excel is installed."
    End If
    
    #If Mac Then
        CheckForExcel = AppleScriptTask("AngryBirds.scpt", "IsExcelInstalled", "noParam")
    #Else
        Dim xlApp As Object
        
        On Error Resume Next
        Set xlApp = CreateObject("Excel.Application")
        On Error GoTo 0
        
        xlApp.Quit
        
        Set xlApp = Nothing
    
        CheckForExcel = Not (Err.Number <> 0)
    #End If
    
    If debugEnabled Then
        Debug.Print "   Installed: " & CheckForExcel
    End If
End Function

#If Mac Then
#Else
Private Function CheckForDotNet35() As Boolean
        If debugEnabled Then
            Debug.Print "Verifying that Microsoft DotNet 3.5 is installed."
        End If

        CheckForDotNet35 = Not (Dir$(Environ$("systemroot") & "\Microsoft.NET\Framework\v3.5", vbDirectory) = vbNullString)
        
        If debugEnabled Then
            Debug.Print "   Installed: " & CheckForDotNet35
        End If
End Function
#End If

'''''''''''''''''''''''''''''''''''''''''''''''''''
'               INITIAL GAME SETUP
'''''''''''''''''''''''''''''''''''''''''''''''''''
#If Mac Then
Private Sub RequestFileAccess()
        Dim fileNameList As String, filePermissionCandidates() As String
        Dim i As Integer
        
        If debugEnabled Then
            Debug.Print vbCrLf & "Requesting access to game files."
        End If
        
        fileNameList = AppleScriptTask("AngryBirds.scpt", "GetFileList", "configFolder")
        
        If fileNameList = "" Then
            Exit Sub
        End If
        
        filePermissionCandidates = Split(fileNameList, ";")
        
        If Not (UBound(filePermissionCandidates) >= LBound(filePermissionCandidates)) Then
            Exit Sub
        End If
        
        For i = LBound(filePermissionCandidates) To UBound(filePermissionCandidates)
            filePermissionCandidates(i) = configFolder & pathSeparator & filePermissionCandidates(i)
        Next i
        
        fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
        
        If debugEnabled Then
            Debug.Print IIf(fileAccessGranted, "   Access has been granted.", "   Acees has been denied.")
        End If
End Sub
#End If

Private Sub SetOSVariables()
    If debugEnabled Then
        Debug.Print "Setting OS-specific variables."
    End If
    
    #If Mac Then
        osIsWindows = False
        pathSeparator = "/"
    #Else
        osIsWindows = True
        pathSeparator = "\"
    #End If
    
    If debugEnabled Then
        Debug.Print "   Operating System: " & Application.OperatingSystem
        Debug.Print "   File path separator: " & pathSeparator
    End If
End Sub

Private Sub SetFolders()
    If debugEnabled Then
        Debug.Print vbCrLf & "Initializing working directories."
    End If
    
    #If Mac Then
        configFolder = AppleScriptTask("AngryBirds.scpt", "SetConfigDirectory", "NoParam")
        tempFolder = AppleScriptTask("AngryBirds.scpt", "SetTempDirectory", "NoParam")
    
        If configFolder = "" Then
            configFolder = ActivePresentation.Path
        End If
        
        If tempFolder = "" Then
            tempFolder = ActivePresentation.Path
        End If
    #Else
        Dim winDocumentsFolder As String
        Dim objFolder As Object
        
        Set objFolder = CreateObject("WScript.Shell").SpecialFolders
        
        winDocumentsFolder = objFolder("mydocuments")
        configFolder = ConvertOneDriveToLocalPath(winDocumentsFolder & "\DYB_AngryBird_ConfigFiles")
        VerifyFolderExists configFolder, "Config folder", ActivePresentation.Path
        
        If VerifyFileOrFolderExists(Environ$("TEMP")) Then
            tempFolder = Environ$("TEMP")
        ElseIf VerifyFileOrFolderExists(Environ$("TMP")) Then
            tempFolder = Environ$("TMP")
        Else
            tempFolder = ActivePresentation.Path
        End If
    #End If
    
    If debugEnabled Then
        Debug.Print "   Configuration folder:   " & configFolder
        Debug.Print "   Temporary files folder: " & tempFolder
    End If
End Sub

#If Mac Then
#Else
Private Sub VerifyFolderExists(ByVal folderPath As String, ByVal folderName As String, ByVal fallbackPath As String)
    Dim messageToDisplay As String
    Dim fs As Object
        
    If debugEnabled Then
        Debug.Print "   Verfiying that directories exist."
    End If
        
    If VerifyFileOrFolderExists(folderPath) Then
        Exit Sub
    End If
    
    If debugEnabled Then
        Debug.Print "   Folder not found. Attempting to create: " & folderPath
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    fs.CreateFolder folderPath
    On Error GoTo 0

    If VerifyFileOrFolderExists(folderPath) Then
        If debugEnabled Then
            Debug.Print "   Folder creation successful. Using path: " & folderPath
        End If
        
        Exit Sub
    End If
    
    If VerifyFileOrFolderExists(fallbackPath) Then
        folderPath = fallbackPath
        
        If debugEnabled Then
            Debug.Print "   Folder creation unsuccessful. Using fallback path: " & folderPath
        End If
    Else
        messageToDisplay = "Critical error! Could not set a working directory. Game cannot continue."
        DisplayMessage messageToDisplay, "OkOnly"
        
        If debugEnabled Then
            Debug.Print "   Unable to set working directory. Terminating game." & folderPath
        End If
        
        ActivePresentation.SlideShowWindow.View.Exit
    End If
End Sub
#End If

Private Sub ResetToDefaultState()
    Dim sld As Slide, shp As Shape
    Dim numOfTeams As String

    If debugEnabled Then
        Debug.Print vbCrLf & "Resetting PPT to default state."
    End If
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Name = "Slide Category" Then
                Select Case Mid(shp.TextFrame.TextRange.Text, 17)
                    Case "Gameboard"
                        ResetGameboardAndLotteryShapes sld
                        
                        numOfTeams = Right(sld.Shapes("Number of Teams").TextFrame.TextRange.Text, 1)
                        
                        Select Case numOfTeams
                            Case "2"
                                With sld.Shapes
                                    .Item("Team1_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Team2_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Marmoset_Scoreboard").TextFrame.TextRange.Text = 0
                                End With
                            Case "3"
                                With sld.Shapes
                                    .Item("Team1_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Team2_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Team3_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Marmoset_Scoreboard").TextFrame.TextRange.Text = 0
                                End With
                            Case "4"
                                With sld.Shapes
                                    .Item("Team1_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Team2_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Team3_Scoreboard").TextFrame.TextRange.Text = 0
                                    .Item("Team4_Scoreboard").TextFrame.TextRange.Text = 0
                                End With
                        End Select
                    Case "Lottery"
                        ResetGameboardAndLotteryShapes sld
                    Case "Config_1"
                        configSld = sld.SlideIndex
                    Case "Config_2"
                        configSld2 = sld.SlideIndex
                End Select
            ElseIf shp.Name = "ChosenQuestionTile" Then
                shp.Delete
            End If
        Next shp
    Next sld
    
    If debugEnabled Then
        Debug.Print "   Gameboard reset to default state."
        Debug.Print "   Scoreboards reset to 0."
        Debug.Print "   Lottery slide reset to deafult state."
        Debug.Print "   Question IDs removed from trivia slides."
        Debug.Print "   Configuration menu found on slides " & configSld & " and " & configSld2 & "."
    End If
End Sub

Private Sub ResetGameboardAndLotteryShapes(ByVal sld As Slide)
    Dim tline As TimeLine, shp As Shape
    Dim i As Integer
    
    Set tline = sld.TimeLine
    
    For Each shp In sld.Shapes
        shp.Visible = msoTrue
    Next shp
    
    If tline.MainSequence.Count > 0 Then
        For i = tline.MainSequence.Count To 1 Step -1
            tline.MainSequence.Item(i).Delete
        Next i
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'               CHECK FOR UPDATES
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CheckForUpdates()
    Dim currentGameVersion As Long, currentQuestionsVersion As Long, onlineGameVersion As Long, onlineQuestionsVersion As Long
    Dim gameUpdateExists As Boolean, questionsUpdateExists As Boolean
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Checking for game and question updates."
    End If
    
    currentGameVersion = GetCurrentVersionNumber("Game")
    currentQuestionsVersion = GetCurrentVersionNumber("Questions")
    onlineGameVersion = GetOnlineVersionNumber("Game")
    onlineQuestionsVersion = GetOnlineVersionNumber("Questions")
    
    ActivePresentation.Slides(configSld).Shapes("Versions_Game_Update").Visible = (onlineGameVersion > currentGameVersion)
    ActivePresentation.Slides(configSld).Shapes("Versions_Questions_Update").Visible = (onlineQuestionsVersion > currentQuestionsVersion)
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Version Check Results:"
        Debug.Print IIf(onlineGameVersion > currentGameVersion, "   Game update found. Displaying update button.", "   Local game version is up-to-date.")
        Debug.Print IIf(onlineGameVersion > currentGameVersion, "   Questions update found. Displaying update button.", "   Local questions version is up-to-date.")
    End If
End Sub

#If Mac Then
Private Sub CheckForScriptUpdate()
    Dim onlineScriptVersion As Long
    Dim updateResult As Boolean
    
    onlineScriptVersion = GetOnlineVersionNumber("AngryBirds.scpt")
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Checking for AngryBirds.scpt updates."
        Debug.Print "   " & AppleScriptTask("AngryBirds.scpt", "GetScriptVersionNumber", "noParam")
        Debug.Print "   Online version: " & onlineScriptVersion
    End If
    
    If onlineScriptVersion <> 0 Then
        updateResult = AppleScriptTask("AngryBirds.scpt", "GetLatestScriptVersion", onlineScriptVersion)
    End If
    
    If debugEnabled Then
        Debug.Print IIf(updateResult, "   AngryBirds.scpt has been updated to latest version.", "   AngryBirds.scpt is already up-to-date.")
    End If
End Sub
#End If

Private Function GetCurrentVersionNumber(ByVal versionToCheck As String) As Long
    Dim localVersion As Long
    Dim sld As Slide
    
    Set sld = ActivePresentation.Slides(configSld)
    
    On Error Resume Next
    If versionToCheck = "Game" Then
        localVersion = ConvertStringToLong(sld.Shapes("Versions_Game_Version").TextFrame.TextRange.Text, 0, ".")
    ElseIf versionToCheck = "Questions" Then
        localVersion = CLng(sld.Shapes("Versions_Questions_Version").TextFrame.TextRange.Text)
    End If
    On Error GoTo 0
    
    GetCurrentVersionNumber = IIf(localVersion <> 0, localVersion, 0)
    
    If debugEnabled Then
        Debug.Print "   Local " & versionToCheck & " version: " & localVersion
    End If
End Function

Private Function GetOnlineVersionNumber(ByVal versionToCheck As String) As Long
    Dim commitDate As String
    
    #If Mac Then
        On Error Resume Next
        commitDate = AppleScriptTask("AngryBirds.scpt", "CheckForNewVersion", versionToCheck)
        On Error GoTo 0
        
        If commitDate <> "" And commitDate <> "0" Then
            commitDate = JsonConverter.ParseJson(commitDate)(1)("commit")("author")("date")
        End If
    #Else
        Dim fileName As String, urlToQuery As String
        Dim http As Object
        
        fileName = IIf(versionToCheck = "Game", "AngryBirdsPPT.txt", "Questions.xlsx") 'Set which file to check for
        urlToQuery = "https://api.github.com/repos/papercutter0324/AngryBirdsTrivia-AdditionalFiles/commits?path=" & fileName
        
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
        http.Open "GET", urlToQuery, False
        http.setRequestHeader "Accept", "application/json"
        http.send
        
        On Error Resume Next
        commitDate = JsonConverter.ParseJson(http.responseText)(1)("commit")("author")("date")
        On Error GoTo 0
    #End If
    
    If commitDate <> "" Then
        GetOnlineVersionNumber = ConvertStringToLong(commitDate)
    End If
    
    If debugEnabled Then
        Debug.Print "   Online " & versionToCheck & " version: " & GetOnlineVersionNumber
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''
'            PREPARE CONFIG & SETTINGS
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub PrepareOptionsMenuStep1(Optional ByVal loadPreviousConfig As String = "Startup")
    Dim sld As Slide
    Dim borderedShape As String, classDayEnabled As String, classTimeEnabled As String
    Dim enableGrammarOrReview As Boolean
    Dim shapePairs As Variant
    Dim i As Integer
    
    Set sld = ActivePresentation.Slides(configSld)
    
    If loadPreviousConfig = "Startup" Then
        If debugEnabled Then
            Debug.Print vbCrLf & "Beginning setup of the configuration menus."
        End If
        
        CreateConfigButtonArrays
    End If
    
    ' Create a multi-dimension array to track shape pairs and their setting's boolean value
    shapePairs = Array( _
        Array("AllowNegatives_Yes", "AllowNegatives_No", allowNegatives), _
        Array("AutoPoints_Yes", "AutoPoints_No", autoAddPoints), _
        Array("RandomChallenges_Yes", "RandomChallenges_No", enableChallenges), _
        Array("Trivia_Include", "Trivia_Exclude", enableTrivia), _
        Array("Grammar_Include", "Grammar_Exclude", enableGrammar), _
        Array("Review_Include", "Review_Exclude", enableReview))
    
    ' Loop through the shape pairs and set visibility
    For i = LBound(shapePairs) To UBound(shapePairs)
        sld.Shapes(shapePairs(i)(1)).Visible = shapePairs(i)(3)
        sld.Shapes(shapePairs(i)(2)).Visible = Not shapePairs(i)(3)
    Next i
    
    ' Determine if the 'Next' or 'Save' button should be shown
    enableGrammarOrReview = enableGrammar Or enableReview
    sld.Shapes("Next_Config").Visible = enableGrammarOrReview
    sld.Shapes("Save_and_Start").Visible = Not enableGrammarOrReview
    
    ' Ensure these shapes are hidden
    sld.Shapes("Updating Indicator").Visible = msoFalse
    sld.Shapes("Updating Message").Visible = msoFalse
    sld.Shapes("Updating Background").Visible = msoFalse
    
    SetConfigButtonVisibility "NumberOfTeams", sld, LBound(numberOfTeams_BorderedArray), UBound(numberOfTeams_BorderedArray), "NumberOfTeams_" & CStr(numberOfTeams) & "_Border", numberOfTeams_BorderedArray, numberOfTeams_BorderlessArray
    SetConfigButtonVisibility "NumberOfLotteryTiles", sld, LBound(numberOfLotteryTiles_BorderedArray), UBound(numberOfLotteryTiles_BorderedArray), "NumberOfLotteryTiles_" & CStr(numberOfLotteryTiles) & "_Border", numberOfLotteryTiles_BorderedArray, numberOfLotteryTiles_BorderlessArray
    SetConfigButtonVisibility "ClassDay", sld, LBound(classTrackerDays_SelectedArray), UBound(classTrackerDays_SelectedArray), "ClassDay_" & classDay & "_Enabled", classTrackerDays_SelectedArray, classTrackerDays_UnselectedArray
    SetConfigButtonVisibility "ClassTime", sld, LBound(classTrackerTimes_SelectedArray), UBound(classTrackerTimes_SelectedArray), "ClassTime_" & classTime & "_Enabled", classTrackerTimes_SelectedArray, classTrackerTimes_UnselectedArray
    
    PrepareOptionsMenuStep2 loadPreviousConfig
    
    If debugEnabled Then
        Debug.Print "Setup complete. Loading menu." & vbCrLf
    End If
    
    ChangeToNewSlide configSld
    DisplayUpdateMessage
End Sub

Private Sub SetConfigButtonVisibility(ByVal setCase As String, ByVal sld As Slide, ByVal firstElement As Integer, ByVal lastElement As Integer, ByVal selectedShape As String, ByRef borderedArray() As Variant, ByRef borderlessArray() As Variant)
    Dim buttonSection As String
    Dim i As Integer
    
    Select Case setCase
        Case "NumberOfTeams", "NumberOfLotteryTiles"
            For i = firstElement To lastElement
                sld.Shapes(borderedArray(i)).Visible = (borderedArray(i) = selectedShape)
                sld.Shapes(borderlessArray(i)).Visible = Not (borderedArray(i) = selectedShape)
            Next i
        Case "ClassDay", "ClassTime"
            buttonSection = IIf(setCase = "ClassDay", classDay, classTime)
            
            If buttonSection <> "" And buttonSection <> "NoTime" And buttonSection <> "NoDay" Then
                For i = firstElement To lastElement
                    sld.Shapes(borderedArray(i)).Visible = (borderedArray(i) = selectedShape)
                    sld.Shapes(borderlessArray(i)).Visible = Not (borderedArray(i) = selectedShape)
                Next i
            Else
                For i = firstElement To lastElement
                    sld.Shapes(borderedArray(i)).Visible = msoFalse
                    sld.Shapes(borderlessArray(i)).Visible = msoTrue
                Next i
            End If
    End Select
End Sub

Private Sub PrepareOptionsMenuStep2(Optional ByVal resetPoint As String)
    Dim sld As Slide
    Dim i As Integer
    
    Set sld = ActivePresentation.Slides(configSld2)

    Select Case resetPoint
        Case "Startup"
            ResetVisibility sld, enabledGrade_SelectedArray, enabledGrade_UnselectedArray, False, True
            ResetVisibility sld, enabledLevel_SelectedArray, enabledLevel_UnselectedArray, False, False
            ResetVisibility sld, enabledBook_SelectedArray, enabledBook_UnselectedArray, False, False
            ResetVisibility sld, enabledUnit_SelectedArray, enabledUnit_UnselectedArray, False, False
        Case "PreviousConfig"
            If enabledGrade <> "" Then
                SetVisibility sld, enabledGrade_SelectedArray, enabledGrade_UnselectedArray, enabledGrade, "Grade"
            End If

            SetLevelVisibility sld, enabledLevel_SelectedArray, enabledLevel_UnselectedArray, enabledGrade, enabledLevel
            
            If enabledBook <> "" And enabledLevel <> "" Then
                SetBookVisibility sld, enabledBook_SelectedArray, enabledBook_UnselectedArray, enabledLevel, enabledBook
                
                If enabledUnit > 0 Then
                    SetUnitVisibility sld, enabledUnit_SelectedArray, enabledUnit_UnselectedArray, enabledUnit
                Else
                    ResetVisibility sld, enabledUnit_SelectedArray, enabledUnit_UnselectedArray, False, False
                End If
            Else
                ResetVisibility sld, enabledBook_SelectedArray, enabledBook_UnselectedArray, False, False
                ResetVisibility sld, enabledUnit_SelectedArray, enabledUnit_UnselectedArray, False, False
            End If
        Case Else
            If resetPoint = "Grade" Then
                SetVisibility sld, enabledGrade_SelectedArray, enabledGrade_UnselectedArray, enabledGrade, "Grade"
            End If
            
            If resetPoint = "Grade" Or resetPoint = "Level" Then
                SetVisibility sld, enabledLevel_SelectedArray, enabledLevel_UnselectedArray, enabledLevel, "Level"
            End If
            
            If resetPoint = "Grade" Or resetPoint = "Level" Or resetPoint = "Book" Then
                SetVisibility sld, enabledBook_SelectedArray, enabledBook_UnselectedArray, enabledBook, "Book"
                SetUnitVisibility sld, enabledUnit_SelectedArray, enabledUnit_UnselectedArray, enabledUnit
            End If
    End Select
End Sub

Private Sub ResetVisibility(ByRef sld As Slide, ByRef selectedArray() As Variant, ByRef unselectedArray() As Variant, ByVal selectedVisible As Boolean, ByVal unselectedVisible As Boolean)
    Dim i As Integer
    
    For i = LBound(selectedArray) To UBound(selectedArray)
        sld.Shapes(selectedArray(i)).Visible = selectedVisible
        sld.Shapes(unselectedArray(i)).Visible = unselectedVisible
    Next i
End Sub

Private Sub SetVisibility(ByRef sld As Slide, ByRef selectedArray() As Variant, ByRef unselectedArray() As Variant, ByVal enabledValue As String, ByVal prefix As String)
    Dim i As Integer
    
    If enabledValue <> "" Then
        For i = LBound(selectedArray) To UBound(selectedArray)
            sld.Shapes(selectedArray(i)).Visible = (selectedArray(i) = prefix & "_" & enabledValue & "_Enabled")
            sld.Shapes(unselectedArray(i)).Visible = Not (selectedArray(i) = prefix & "_" & enabledValue & "_Enabled")
        Next i
    Else
        For i = LBound(selectedArray) To UBound(selectedArray)
            sld.Shapes(selectedArray(i)).Visible = msoFalse
            sld.Shapes(unselectedArray(i)).Visible = msoFalse
        Next i
    End If
End Sub

Private Sub SetLevelVisibility(ByRef sld As Slide, ByRef selectedArray() As Variant, ByRef unselectedArray() As Variant, ByVal enabledGrade As String, ByVal enabledLevel As String)
    Dim shp As Shape
    Dim i As Integer, caseLevel As String
    
    For i = LBound(unselectedArray) To UBound(unselectedArray)
        Set shp = sld.Shapes(unselectedArray(i))
        caseLevel = unselectedArray(i)
        shp.Visible = IsLevelVisibleForGrade(caseLevel, enabledGrade)
        
        If selectedArray(i) = "Level_" & enabledLevel & "_Enabled" Then
            sld.Shapes(selectedArray(i)).Visible = msoTrue
            sld.Shapes(unselectedArray(i)).Visible = msoFalse
        End If
    Next i
End Sub

Private Sub SetBookVisibility(ByRef sld As Slide, ByRef selectedArray() As Variant, ByRef unselectedArray() As Variant, ByVal enabledLevel As String, ByVal enabledBook As String)
    Dim shp As Shape
    Dim i As Integer, caseBook As String, bookNameArray() As String
    
    For i = LBound(unselectedArray) To UBound(unselectedArray)
        Set shp = sld.Shapes(unselectedArray(i))
        
        bookNameArray = Split(enabledBook_UnselectedArray(i), "_")
        caseBook = bookNameArray(1)
        shp.Visible = IsBookVisibleForLevel(caseBook, enabledLevel)
        
        If selectedArray(i) = "Book_" & enabledBook & "_Enabled" Then
            sld.Shapes(selectedArray(i)).Visible = msoTrue
            sld.Shapes(unselectedArray(i)).Visible = msoFalse
        End If
    Next i
End Sub

Private Sub SetUnitVisibility(ByRef sld As Slide, ByRef selectedArray() As Variant, ByRef unselectedArray() As Variant, ByVal enabledUnit As Integer)
    Dim i As Integer, maxUnitNumber As Integer
    
    maxUnitNumber = SetMaxUnitNumber
    
    If maxUnitNumber = 0 Then
        maxUnitNumber = UBound(selectedArray)
    End If
    
    Select Case True
        Case (enabledUnit > 0)
            For i = LBound(selectedArray) To maxUnitNumber
                sld.Shapes(selectedArray(i)).Visible = (i <= enabledUnit)
                sld.Shapes(unselectedArray(i)).Visible = Not (i <= enabledUnit)
            Next i
            
            If maxUnitNumber < UBound(selectedArray) Then
                For i = maxUnitNumber + 1 To UBound(selectedArray)
                    sld.Shapes(selectedArray(i)).Visible = msoFalse
                    sld.Shapes(unselectedArray(i)).Visible = msoFalse
                Next i
            End If
        Case (enabledUnit < 0)
            For i = LBound(selectedArray) To maxUnitNumber
                sld.Shapes(selectedArray(i)).Visible = msoFalse
                sld.Shapes(unselectedArray(i)).Visible = msoTrue
            Next i
            
            If maxUnitNumber < UBound(selectedArray) Then
                For i = maxUnitNumber + 1 To UBound(selectedArray)
                    sld.Shapes(selectedArray(i)).Visible = msoFalse
                    sld.Shapes(unselectedArray(i)).Visible = msoFalse
                Next i
            End If
        Case Else
            For i = LBound(selectedArray) To UBound(selectedArray)
                sld.Shapes(selectedArray(i)).Visible = msoFalse
                sld.Shapes(unselectedArray(i)).Visible = msoFalse
            Next i
    End Select
End Sub

Private Function SetMaxUnitNumber() As Integer
    Select Case enabledBook
        Case "WC-2.5", "WC-2.6", "WC-3.1", "WC-3.2", "WC-3.3", "WC-3.4", "WC-3.5"
            SetMaxUnitNumber = 0
        Case "Journeys-3.1", "Journeys-3.2"
            SetMaxUnitNumber = 0
        Case "IntoReading-4.1", "IntoReading-4.2"
            SetMaxUnitNumber = 0
        Case "BR-4.3", "BR-5.1", "BR-5.2", "BR-5.3"
            SetMaxUnitNumber = 0
        Case "SL4", "SL5", "SL6", "SL7", "SL8"
            SetMaxUnitNumber = 0
        Case "RD5", "RD6", "RD7", "RD8"
            SetMaxUnitNumber = 0
        Case "ReadUp1", "ReadUp2"
            SetMaxUnitNumber = 20
        Case "TARA"
            SetMaxUnitNumber = 25
        Case "Gravoca1A", "Gravoca1B", "Gravoca1C"
            SetMaxUnitNumber = 0
        Case "Gravoca2A", "Gravoca2B"
            SetMaxUnitNumber = 0
        Case "RS-S1B1", "RS-S1B2", "RS-S1B3", "RS-S1B4", "RS-S1B5", "RS-S1B6", "RS-S1B7"
            SetMaxUnitNumber = 8
        Case "RS-S2B1", "RS-S2B2", "RS-S2B3", "RS-S2B4", "RS-S2B5", "RS-S2B6", "RS-S2B7"
            SetMaxUnitNumber = 8
        Case Else
            SetMaxUnitNumber = -1
    End Select
End Function

Private Sub DisplayUpdateMessage()
    Dim sld As Slide
    Dim messageToDisplay As String

    Set sld = ActivePresentation.Slides(configSld)
    
    If (sld.Shapes("Versions_Game_Update").Visible = msoTrue) Or (sld.Shapes("Versions_Questions_Update").Visible = msoTrue) Then
        messageToDisplay = "There is an update available!" & vbCrLf & "It is highly recommended to download the update before continuing."
        DisplayMessage messageToDisplay, "OkOnly"
    End If
End Sub

Private Sub CreateConfigButtonArrays()
    Dim numberOfTeamsButtons As Integer, numberOfLotteryTilesButtons As Integer, numberOfGradeButtons As Integer, numberOfLevelButtons As Integer
    Dim numberOfBookButtons As Integer, numberOfUnitButtons As Integer, numberOfClassDayButtons As Integer, numberOfClassTimeButtons As Integer
    
    numberOfTeamsButtons = 3
    numberOfLotteryTilesButtons = 7
    numberOfGradeButtons = 5
    numberOfLevelButtons = 24
    numberOfBookButtons = 47
    numberOfUnitButtons = 30
    numberOfClassDayButtons = 9
    numberOfClassTimeButtons = 11

    ReDim numberOfTeams_BorderedArray(1 To numberOfTeamsButtons)
    ReDim numberOfTeams_BorderlessArray(1 To numberOfTeamsButtons)
    ReDim numberOfLotteryTiles_BorderedArray(1 To numberOfLotteryTilesButtons)
    ReDim numberOfLotteryTiles_BorderlessArray(1 To numberOfLotteryTilesButtons)
    ReDim enabledGrade_SelectedArray(1 To numberOfGradeButtons)
    ReDim enabledGrade_UnselectedArray(1 To numberOfGradeButtons)
    ReDim enabledLevel_SelectedArray(1 To numberOfLevelButtons)
    ReDim enabledLevel_UnselectedArray(1 To numberOfLevelButtons)
    ReDim enabledBook_SelectedArray(1 To numberOfLevelButtons)
    ReDim enabledBook_UnselectedArray(1 To numberOfLevelButtons)
    ReDim enabledUnit_SelectedArray(1 To numberOfUnitButtons)
    ReDim enabledUnit_UnselectedArray(1 To numberOfUnitButtons)
    ReDim classTrackerDays_SelectedArray(1 To numberOfClassDayButtons)
    ReDim classTrackerDays_UnselectedArray(1 To numberOfClassDayButtons)
    ReDim classTrackerTimes_SelectedArray(1 To numberOfClassTimeButtons)
    ReDim classTrackerTimes_UnselectedArray(1 To numberOfClassTimeButtons)
    
    numberOfTeams_BorderedArray = Array("NumberOfTeams_2_Border", "NumberOfTeams_3_Border", "NumberOfTeams_4_Border")

    numberOfTeams_BorderlessArray = Array("NumberOfTeams_2_Borderless", "NumberOfTeams_3_Borderless", "NumberOfTeams_4_Borderless")
    
    numberOfLotteryTiles_BorderedArray = Array("NumberOfLotteryTiles_2_Border", "NumberOfLotteryTiles_3_Border", "NumberOfLotteryTiles_4_Border", _
                                               "NumberOfLotteryTiles_5_Border", "NumberOfLotteryTiles_6_Border", "NumberOfLotteryTiles_7_Border", _
                                               "NumberOfLotteryTiles_8_Border")
    
    numberOfLotteryTiles_BorderlessArray = Array("NumberOfLotteryTiles_2_Borderless", "NumberOfLotteryTiles_3_Borderless", "NumberOfLotteryTiles_4_Borderless", _
                                                 "NumberOfLotteryTiles_5_Borderless", "NumberOfLotteryTiles_6_Borderless", "NumberOfLotteryTiles_7_Borderless", _
                                                 "NumberOfLotteryTiles_8_Borderless")

    enabledGrade_SelectedArray = Array("Grade_E4_Enabled", "Grade_E5_Enabled", "Grade_E6_Enabled", "Grade_M1_Enabled", "Grade_M2_Enabled")
    
    enabledGrade_UnselectedArray = Array("Grade_E4_Disabled", "Grade_E5_Disabled", "Grade_E6_Disabled", "Grade_M1_Disabled", "Grade_M2_Disabled")
    
    enabledLevel_SelectedArray = Array("Level_Theseus_Enabled", "Level_Perseus_Enabled", "Level_Odysseus_Enabled", "Level_Hercules_Enabled", _
                                       "Level_Artemis_Enabled", "Level_Hermes_Enabled", "Level_Apollo_Enabled", "Level_Zeus_Enabled", _
                                       "Level_Helios_Enabled", "Level_Poseidon_Enabled", "Level_Gaia_Enabled", "Level_Hera_Enabled", _
                                       "Level_E5-Athena_Enabled", "Level_E6-Song's_Enabled", "Level_M1-Song's_Enabled", "Level_M2-Song's_Enabled", _
                                       "Level_Elephantus_Enabled", "Level_Galaxia_Enabled", "Level_Solis_Enabled", "Level_M1-Major_Enabled", _
                                       "Level_Ursa_Enabled", "Level_Leo_Enabled", "Level_Tigris_Enabled", "Level_M2-Major_Enabled")

    enabledLevel_UnselectedArray = Array("Level_Theseus_Disabled", "Level_Perseus_Disabled", "Level_Odysseus_Disabled", "Level_Hercules_Disabled", _
                                         "Level_Artemis_Disabled", "Level_Hermes_Disabled", "Level_Apollo_Disabled", "Level_Zeus_Disabled", _
                                         "Level_Helios_Disabled", "Level_Poseidon_Disabled", "Level_Gaia_Disabled", "Level_Hera_Disabled", _
                                         "Level_E5-Athena_Disabled", "Level_E6-Song's_Disabled", "Level_M1-Song's_Disabled", "Level_M2-Song's_Disabled", _
                                         "Level_Elephantus_Disabled", "Level_Galaxia_Disabled", "Level_Solis_Disabled", "Level_M1-Major_Disabled", _
                                         "Level_Ursa_Disabled", "Level_Leo_Disabled", "Level_Tigris_Disabled", "Level_M2-Major_Disabled")

    enabledBook_SelectedArray = Array("Book_WC-2.5_Enabled", "Book_WC-2.6_Enabled", "Book_WC-3.1_Enabled", "Book_WC-3.2_Enabled", "Book_WC-3.3_Enabled", _
                                      "Book_WC-3.4_Enabled", "Book_WC-3.5_Enabled", "Book_WC-2.5_Enabled", "Book_Journeys-3.1_Enabled", "Book_Journeys-3.2_Enabled", _
                                      "Book_IntoReading-4.1_Enabled", "Book_IntoReading-4.2_Enabled", "Book_BR-4.3_Enabled", "Book_BR-5.1_Enabled", "Book_BR-5.2_Enabled", _
                                      "Book_BR-5.3_Enabled", "Book_SL4_Enabled", "Book_SL5_Enabled", "Book_SL6_Enabled", "Book_SL7_Enabled", "Book_SL8_Enabled", _
                                      "Book_RD5_Enabled", "Book_RD6_Enabled", "Book_RD7_Enabled", "Book_RD8_Enabled", "Book_ReadUp1_Enabled", "Book_ReadUp2_Enabled", _
                                      "Book_TARA_Enabled", "Book_Gravoca1A_Enabled", "Book_Gravoca1B_Enabled", "Book_Gravoca1C_Enabled", "Book_Gravoca2A_Enabled", _
                                      "Book_Gravoca2B_Enabled", "Book_RS-S1B1_Enabled", "Book_RS-S1B2_Enabled", "Book_RS-S1B3_Enabled", "Book_RS-S1B4_Enabled", _
                                      "Book_RS-S1B5_Enabled", "Book_RS-S1B6_Enabled", "Book_RS-S1B7_Enabled", "Book_RS-S2B1_Enabled", "Book_RS-S2B2_Enabled", _
                                      "Book_RS-S2B3_Enabled", "Book_RS-S2B4_Enabled", "Book_RS-S2B5_Enabled", "Book_RS-S2B6_Enabled", "Book_RS-S2B7_Enabled")
                                       
    enabledBook_UnselectedArray = Array("Book_WC-2.5_Disabled", "Book_WC-2.6_Disabled", "Book_WC-3.1_Disabled", "Book_WC-3.2_Disabled", "Book_WC-3.3_Disabled", _
                                        "Book_WC-3.4_Disabled", "Book_WC-3.5_Disabled", "Book_WC-2.5_Disabled", "Book_Journeys-3.1_Disabled", "Book_Journeys-3.2_Disabled", _
                                        "Book_IntoReading-4.1_Disabled", "Book_IntoReading-4.2_Disabled", "Book_BR-4.3_Disabled", "Book_BR-5.1_Disabled", "Book_BR-5.2_Disabled", _
                                        "Book_BR-5.3_Disabled", "Book_SL4_Disabled", "Book_SL5_Disabled", "Book_SL6_Disabled", "Book_SL7_Disabled", "Book_SL8_Disabled", _
                                        "Book_RD5_Disabled", "Book_RD6_Disabled", "Book_RD7_Disabled", "Book_RD8_Disabled", "Book_ReadUp1_Disabled", "Book_ReadUp2_Disabled", _
                                        "Book_TARA_Disabled", "Book_Gravoca1A_Disabled", "Book_Gravoca1B_Disabled", "Book_Gravoca1C_Disabled", "Book_Gravoca2A_Disabled", _
                                        "Book_Gravoca2B_Disabled", "Book_RS-S1B1_Disabled", "Book_RS-S1B2_Disabled", "Book_RS-S1B3_Disabled", "Book_RS-S1B4_Disabled", _
                                        "Book_RS-S1B5_Disabled", "Book_RS-S1B6_Disabled", "Book_RS-S1B7_Disabled", "Book_RS-S2B1_Disabled", "Book_RS-S2B2_Disabled", _
                                        "Book_RS-S2B3_Disabled", "Book_RS-S2B4_Disabled", "Book_RS-S2B5_Disabled", "Book_RS-S2B6_Disabled", "Book_RS-S2B7_Disabled")

    enabledUnit_SelectedArray = Array("Unit_1_Enabled", "Unit_2_Enabled", "Unit_3_Enabled", "Unit_4_Enabled", "Unit_5_Enabled", "Unit_6_Enabled", "Unit_7_Enabled", _
                                      "Unit_8_Enabled", "Unit_9_Enabled", "Unit_10_Enabled", "Unit_11_Enabled", "Unit_12_Enabled", "Unit_13_Enabled", "Unit_14_Enabled", _
                                      "Unit_15_Enabled", "Unit_16_Enabled", "Unit_17_Enabled", "Unit_18_Enabled", "Unit_19_Enabled", "Unit_20_Enabled", "Unit_21_Enabled", _
                                      "Unit_22_Enabled", "Unit_23_Enabled", "Unit_24_Enabled", "Unit_25_Enabled", "Unit_26_Enabled", "Unit_27_Enabled", "Unit_28_Enabled", _
                                      "Unit_29_Enabled", "Unit_99_Enabled")

    enabledUnit_UnselectedArray = Array("Unit_1_Disabled", "Unit_2_Disabled", "Unit_3_Disabled", "Unit_4_Disabled", "Unit_5_Disabled", "Unit_6_Disabled", "Unit_7_Disabled", _
                                        "Unit_8_Disabled", "Unit_9_Disabled", "Unit_10_Disabled", "Unit_11_Disabled", "Unit_12_Disabled", "Unit_13_Disabled", "Unit_14_Disabled", _
                                        "Unit_15_Disabled", "Unit_16_Disabled", "Unit_17_Disabled", "Unit_18_Disabled", "Unit_19_Disabled", "Unit_20_Disabled", "Unit_21_Disabled", _
                                        "Unit_22_Disabled", "Unit_23_Disabled", "Unit_24_Disabled", "Unit_25_Disabled", "Unit_26_Disabled", "Unit_27_Disabled", "Unit_28_Disabled", _
                                        "Unit_29_Disabled", "Unit_99_Disabled")
                                       
    classTrackerDays_SelectedArray = Array("ClassDay_MonWed_Enabled", "ClassDay_MonFri_Enabled", "ClassDay_WedFri_Enabled", "ClassDay_TuesThurs_Enabled", "ClassDay_Monday_Enabled", _
                                           "ClassDay_Tuesday_Enabled", "ClassDay_Wednesday_Enabled", "ClassDay_Thursday_Enabled", "ClassDay_Friday_Enabled")

    classTrackerDays_UnselectedArray = Array("ClassDay_MonWed_Disabled", "ClassDay_MonFri_Disabled", "ClassDay_WedFri_Disabled", "ClassDay_TuesThurs_Disabled", "ClassDay_Monday_Disabled", _
                                             "ClassDay_Tuesday_Disabled", "ClassDay_Wednesday_Disabled", "ClassDay_Thursday_Disabled", "ClassDay_Friday_Disabled")

    classTrackerTimes_SelectedArray = Array("ClassTime_4PM_Enabled", "ClassTime_430PM_Enabled", "ClassTime_5PM_Enabled", "ClassTime_530PM_Enabled", "ClassTime_6PM_Enabled", _
                                            "ClassTime_630PM_Enabled", "ClassTime_7PM_Enabled", "ClassTime_730PM_Enabled", "ClassTime_8PM_Enabled", "ClassTime_830PM_Enabled", _
                                            "ClassTime_9PM_Enabled")

    classTrackerTimes_UnselectedArray = Array("ClassTime_4PM_Disabled", "ClassTime_430PM_Disabled", "ClassTime_5PM_Disabled", "ClassTime_530PM_Disabled", "ClassTime_6PM_Disabled", _
                                              "ClassTime_630PM_Disabled", "ClassTime_7PM_Disabled", "ClassTime_730PM_Disabled", "ClassTime_8PM_Disabled", "ClassTime_830PM_Disabled", _
                                              "ClassTime_9PM_Disabled")
End Sub

Public Sub UpdateConfiguration(clickedShp As Shape)
    Dim sld As Slide
    Dim shapeName As String, categoryName As String, categorySetting As String, bookName As String, messageToDisplay As String
    Dim shapeNameArray() As String, bookNameArray() As String
    Dim userChoice As Integer, i As Integer
    
    Set sld = ActivePresentation.SlideShowWindow.View.Slide
    
    'Grab the category and value from the clicked shape's name
    shapeName = clickedShp.Name
    shapeNameArray = Split(shapeName, "_")
    categoryName = shapeNameArray(0)
    categorySetting = shapeNameArray(1)
    
    If NoSettingUpdateNeeded(categoryName, categorySetting) Then
        Exit Sub
    End If
    
    If debugEnabled Then
        Select Case shapeName
            Case "Save_and_Start"
                Debug.Print vbCrLf & "Saving settings and beginning game."
            Case "Next_Config"
                Debug.Print "Changing to second options menu."
            Case "Reset_All_Settings"
                
            Case Else
                Debug.Print "Updating: " & categoryName
        End Select
    End If
    
    Select Case categoryName
        Case "NumberOfTeams"
            ToggleArrayConfigButtons sld, categoryName, categorySetting, numberOfTeams, numberOfTeams_BorderedArray, numberOfTeams_BorderlessArray
        Case "NumberOfLotteryTiles"
            ToggleArrayConfigButtons sld, categoryName, categorySetting, numberOfLotteryTiles, numberOfLotteryTiles_BorderedArray, numberOfLotteryTiles_BorderlessArray
        Case "AllowNegatives"
            ToggleBooleanConfigButtons sld, allowNegatives, "AllowNegatives_No", "AllowNegatives_Yes"
        Case "AutoPoints"
            ToggleBooleanConfigButtons sld, autoAddPoints, "AutoPoints_No", "AutoPoints_Yes"
        Case "RandomChallenges"
            ToggleBooleanConfigButtons sld, enableChallenges, "RandomChallenges_No", "RandomChallenges_Yes"
        Case "Trivia"
            If enableTrivia Then
                messageToDisplay = "Are you sure you want to disable all trivia questions?"
                DisplayMessage messageToDisplay, "YesNo", userChoice
                
                If userChoice = vbNo Then
                    Exit Sub
                End If
            End If
            ToggleBooleanConfigButtons sld, enableTrivia, "Trivia_Exclude", "Trivia_Include"
        Case "Grammar"
            ToggleBooleanConfigButtons sld, enableGrammar, "Grammar_Exclude", "Grammar_Include"
        Case "Review"
            ToggleBooleanConfigButtons sld, enableReview, "Review_Exclude", "Review_Include"
        Case "ClassDay", "ClassTime"
            If categoryName = "ClassDay" Then
                ToggleArrayConfigButtons sld, categoryName, categorySetting, classDay, classTrackerDays_SelectedArray, classTrackerDays_UnselectedArray
            Else
                ToggleArrayConfigButtons sld, categoryName, categorySetting, classTime, classTrackerTimes_SelectedArray, classTrackerTimes_UnselectedArray
            End If
            
            If previouslySeenQuestionsArray(1) <> "" Then
                ReDim previouslySeenQuestionsArray(1)
                previouslySeenQuestionsArray(1) = ""
            
                If debugEnabled Then
                    Debug.Print "   Reset list of previously seen questions."
                End If
            End If
            
            CheckForPreviousConfig
        Case "Grade"
            UpdateGrade sld, categorySetting
        Case "Level"
            UpdateLevel sld, categorySetting, shapeName
        Case "Book"
            UpdateBook sld, categorySetting, shapeName
        Case "Unit"
            ToggleArrayConfigButtons sld, categoryName, categorySetting, enabledUnit, enabledUnit_SelectedArray, enabledUnit_UnselectedArray
        Case "Next"
            ValidateAndProceedNext sld
        Case "Save"
            SaveSettings sld
        Case "Reset"
            ResetSettings
    End Select

    If sld.SlideIndex = configSld Then
        sld.Shapes("Next_Config").Visible = enableGrammar Or enableReview
        sld.Shapes("Save_and_Start").Visible = Not (enableGrammar Or enableReview)
    End If
End Sub

Private Function NoSettingUpdateNeeded(ByVal categoryName As String, ByVal categorySetting As String) As Boolean
    Select Case categoryName
        Case "NumberOfTeams"
            NoSettingUpdateNeeded = (categorySetting = numberOfTeams)
        Case "NumberOfLotteryTiles"
            NoSettingUpdateNeeded = (categorySetting = numberOfLotteryTiles)
        Case "ClassDay"
            NoSettingUpdateNeeded = (categorySetting = classDay)
        Case "ClassTime"
            NoSettingUpdateNeeded = (categorySetting = classTime)
        Case "Grade"
            NoSettingUpdateNeeded = (categorySetting = enabledGrade)
        Case "Level"
            NoSettingUpdateNeeded = (categorySetting = enabledLevel)
        Case "Book"
            NoSettingUpdateNeeded = (categorySetting = enabledBook)
    End Select
End Function

Private Sub ToggleArrayConfigButtons(ByVal sld As Slide, ByVal categoryName As String, ByVal settingValue As Variant, ByRef currentValue As Variant, ByVal borderedShapes As Variant, ByVal borderlessShapes As Variant)
    Dim i As Integer, maxUnitNumber As Integer

    Select Case categoryName
        Case "NumberOfTeams", "NumberOfLotteryTiles"
            For i = LBound(borderedShapes) To UBound(borderedShapes)
                sld.Shapes(borderedShapes(i)).Visible = (i + 1 = settingValue)
                sld.Shapes(borderlessShapes(i)).Visible = Not (i + 1 = settingValue)
            Next i
        Case "ClassDay", "ClassTime", "Grade"
            For i = LBound(borderedShapes) To UBound(borderedShapes)
                sld.Shapes(borderedShapes(i)).Visible = (categoryName & "_" & settingValue & "_Disabled" = borderlessShapes(i))
                sld.Shapes(borderlessShapes(i)).Visible = Not (categoryName & "_" & settingValue & "_Disabled" = borderlessShapes(i))
            Next i
        Case "Unit"
            maxUnitNumber = SetMaxUnitNumber
            
            If maxUnitNumber = 0 Then
                maxUnitNumber = UBound(borderedShapes)
            End If
        
            For i = LBound(borderedShapes) To maxUnitNumber
                sld.Shapes(borderedShapes(i)).Visible = (i <= settingValue)
                sld.Shapes(borderlessShapes(i)).Visible = Not (i <= settingValue)
            Next i
            
            If maxUnitNumber < UBound(borderedShapes) Then
                For i = maxUnitNumber + 1 To UBound(borderedShapes)
                    sld.Shapes(borderedShapes(i)).Visible = msoFalse
                    sld.Shapes(borderlessShapes(i)).Visible = msoFalse
                Next i
            End If
    End Select
    
    If debugEnabled Then
        Debug.Print "   Value: " & settingValue
    End If
    
    currentValue = settingValue
End Sub

Private Sub ToggleBooleanConfigButtons(ByVal sld As Slide, ByRef boolVariable As Boolean, ByVal shapeNameNo As String, ByVal shapeNameYes As String)
    boolVariable = Not boolVariable
    sld.Shapes(shapeNameNo).Visible = Not boolVariable
    sld.Shapes(shapeNameYes).Visible = boolVariable
    
    If debugEnabled Then
        Debug.Print "   Value: " & boolVariable
    End If
End Sub

Private Sub UpdateGrade(sld As Slide, categorySetting As String)
    ' For the dependent settings, clear their variables and hide their shapes
    enabledLevel = ""
    enabledBook = ""
    enabledUnit = 0
    PrepareOptionsMenuStep2 "Grade"
    
    ToggleArrayConfigButtons sld, "Grade", categorySetting, enabledGrade, enabledGrade_SelectedArray, enabledGrade_UnselectedArray
    
    UpdateVisibilityBasedOnGrade sld, enabledGrade
End Sub

Private Sub UpdateVisibilityBasedOnGrade(sld As Slide, enabledGrade As String)
    Dim shp_LevelShape As Shape
    Dim caseLevel As String
    Dim i As Integer
    
    For i = LBound(enabledLevel_UnselectedArray) To UBound(enabledLevel_UnselectedArray)
        Set shp_LevelShape = sld.Shapes(enabledLevel_UnselectedArray(i))
        
        caseLevel = enabledLevel_UnselectedArray(i)
        shp_LevelShape.Visible = IsLevelVisibleForGrade(caseLevel, enabledGrade)
    Next i
End Sub

Private Function IsLevelVisibleForGrade(caseLevel As String, enabledGrade As String) As Boolean
    Select Case caseLevel
        Case "Level_Theseus_Disabled", "Level_Perseus_Disabled", "Level_Odysseus_Disabled", "Level_Hercules_Disabled"
            IsLevelVisibleForGrade = (enabledGrade = "E4")
        Case "Level_Artemis_Disabled", "Level_Hermes_Disabled", "Level_Apollo_Disabled", "Level_Zeus_Disabled", "Level_E5-Athena_Disabled"
            IsLevelVisibleForGrade = (enabledGrade = "E5")
        Case "Level_Helios_Disabled", "Level_Poseidon_Disabled", "Level_Gaia_Disabled", "Level_Hera_Disabled", "Level_E6-Song's_Disabled"
            IsLevelVisibleForGrade = (enabledGrade = "E6")
        Case "Level_Elephantus_Disabled", "Level_Galaxia_Disabled", "Level_Solis_Disabled", "Level_M1-Major_Disabled", "Level_M1-Song's_Disabled"
            IsLevelVisibleForGrade = (enabledGrade = "M1")
        Case "Level_Ursa_Disabled", "Level_Leo_Disabled", "Level_Tigris_Disabled", "Level_M2-Major_Disabled", "Level_M2-Song's_Disabled"
            IsLevelVisibleForGrade = (enabledGrade = "M2")
        Case Else
            IsLevelVisibleForGrade = False
    End Select
End Function

Private Sub UpdateLevel(sld As Slide, categorySetting As String, shapeName As String)
    Dim i As Integer
    
    ' For the dependent settings, clear their variables and hide their shapes
    enabledBook = ""
    enabledUnit = 0
    PrepareOptionsMenuStep2 "Level"
    
    EnableClickedLevel sld, shapeName
    
    ToggleArrayConfigButtons sld, "Level", categorySetting, enabledLevel, enabledLevel_SelectedArray, enabledLevel_UnselectedArray
    
    enabledLevel = categorySetting
    
    UpdateBookVisibility sld, enabledLevel
End Sub

Private Sub EnableClickedLevel(sld As Slide, selectedLevel As String)
    Dim levelShape As Shape
    Dim caseLevel As String
    Dim i As Integer
    
    For i = LBound(enabledLevel_UnselectedArray) To UBound(enabledLevel_UnselectedArray)
        Set levelShape = sld.Shapes(enabledLevel_UnselectedArray(i))
        
        caseLevel = enabledLevel_UnselectedArray(i)
        sld.Shapes(enabledLevel_SelectedArray(i)).Visible = (selectedLevel = enabledLevel_UnselectedArray(i))
    
        Select Case caseLevel
            Case "Level_Theseus_Disabled", "Level_Perseus_Disabled", "Level_Odysseus_Disabled", "Level_Hercules_Disabled"
                 levelShape.Visible = (enabledGrade = "E4")
            Case "Level_Artemis_Disabled", "Level_Hermes_Disabled", "Level_Apollo_Disabled", "Level_Zeus_Disabled", "Level_E5-Athena_Disabled"
                 levelShape.Visible = (enabledGrade = "E5")
            Case "Level_Helios_Disabled", "Level_Poseidon_Disabled", "Level_Gaia_Disabled", "Level_Hera_Disabled", "Level_E6-Song's_Disabled"
                 levelShape.Visible = (enabledGrade = "E6")
            Case "Level_Elephantus_Disabled", "Level_Galaxia_Disabled", "Level_Solis_Disabled", "Level_M1-Major_Disabled", "Level_M1-Song's_Disabled"
                 levelShape.Visible = (enabledGrade = "M1")
            Case "Level_Ursa_Disabled", "Level_Leo_Disabled", "Level_Tigris_Disabled", "Level_M2-Major_Disabled", "Level_M2-Song's_Disabled"
                 levelShape.Visible = (enabledGrade = "M2")
        End Select
    Next i
End Sub

Private Sub UpdateBookVisibility(sld As Slide, enabledLevel As String)
    Dim shp_BookShape As Shape
    Dim caseBook As String
    Dim bookNameArray() As String
    Dim i As Integer
    
    For i = LBound(enabledBook_UnselectedArray) To UBound(enabledBook_UnselectedArray)
        Set shp_BookShape = sld.Shapes(enabledBook_UnselectedArray(i))
        
        bookNameArray = Split(enabledBook_UnselectedArray(i), "_")
        caseBook = bookNameArray(1)
        
        shp_BookShape.Visible = IsBookVisibleForLevel(caseBook, enabledLevel)
    Next i
End Sub

Private Function IsBookVisibleForLevel(caseBook As String, enabledLevel As String) As Boolean
    Select Case enabledLevel
         Case "Theseus", "Perseus"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "WC")
        Case "Odysseus"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "Jo")
        Case "Hercules"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "In")
        Case "Artemis", "Hermes"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "BR")
        Case "Apollo", "Zeus"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "SL")
        Case "Helios", "Poseidon", "Gaia"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "RD")
        Case "Hera"
            Select Case "Book_" & caseBook & "_Disabled"
                Case "Book_ReadUp1_Disabled", "Book_ReadUp2_Disabled", "Book_TARA_Disabled"
                    IsBookVisibleForLevel = True
                Case Else
                    IsBookVisibleForLevel = False
            End Select
        Case "Elephantus", "Galaxia"
            Select Case "Book_" & caseBook & "_Disabled"
                Case "Book_Gravoca1A_Disabled", "Book_Gravoca1B_Disabled", "Book_Gravoca1C_Disabled"
                    IsBookVisibleForLevel = True
                Case Else
                    IsBookVisibleForLevel = False
            End Select
        Case "Ursa", "Leo"
            Select Case "Book_" & caseBook & "_Disabled"
                Case "Book_Gravoca2A_Disabled", "Book_Gravoca2B_Disabled"
                    IsBookVisibleForLevel = True
                Case Else
                    IsBookVisibleForLevel = False
            End Select
        Case "Solis", "M1-Major", "M1-Song's", "Tigris", "M2-Major", "M2-Song's"
            IsBookVisibleForLevel = (Left(caseBook, 2) = "RS")
        Case Else
            IsBookVisibleForLevel = False
    End Select
End Function

Private Sub UpdateBook(sld As Slide, categorySetting As String, shapeName As String)
    Dim i As Integer, maxUnitNumber As Integer
    
    ' For the dependent settings, clear their variables and hide their shapes
    enabledUnit = 0
    PrepareOptionsMenuStep2 "Book"
    
    EnableClickedBook sld, shapeName
    
    ToggleArrayConfigButtons sld, "Book", categorySetting, enabledBook, enabledBook_SelectedArray, enabledBook_UnselectedArray
    
    enabledBook = categorySetting
    
    maxUnitNumber = SetMaxUnitNumber
    
    If maxUnitNumber = 0 Then
        maxUnitNumber = UBound(enabledBook_SelectedArray)
    End If
    
    For i = LBound(enabledUnit_SelectedArray) To UBound(enabledUnit_SelectedArray)
        sld.Shapes(enabledUnit_UnselectedArray(i)).Visible = (i <= maxUnitNumber)
    Next i
End Sub

Private Sub EnableClickedBook(sld As Slide, selectedBook As String)
    Dim bookShape As Shape
    Dim bookName As String
    Dim bookNameArray() As String
    Dim i As Integer
    
    For i = LBound(enabledBook_SelectedArray) To UBound(enabledBook_SelectedArray)
        Set bookShape = sld.Shapes(enabledBook_UnselectedArray(i))
        
        bookShape.Visible = msoFalse
        sld.Shapes(enabledBook_SelectedArray(i)).Visible = (selectedBook = enabledBook_UnselectedArray(i))
        
        bookNameArray = Split(enabledBook_UnselectedArray(i), "_")
        bookName = bookNameArray(1)
        
        Select Case enabledLevel
            Case "Theseus", "Perseus"
                bookShape.Visible = (Left(bookName, 2) = "WC")
            Case "Odysseus"
                bookShape.Visible = (Left(bookName, 2) = "Jo")
            Case "Hercules"
                bookShape.Visible = (Left(bookName, 2) = "In")
            Case "Artemis", "Hermes"
                bookShape.Visible = (Left(bookName, 2) = "BR")
            Case "Apollo", "Zeus"
                bookShape.Visible = (Left(bookName, 2) = "SL")
            Case "Helios", "Poseidon", "Gaia"
                bookShape.Visible = (Left(bookName, 2) = "RD")
            Case "Hera"
                Select Case enabledBook_UnselectedArray(i)
                    Case "Book_ReadUp1_Disabled", "Book_ReadUp2_Disabled", "Book_TARA_Disabled"
                        bookShape.Visible = msoTrue
                End Select
            Case "Elephantus", "Galaxia"
                Select Case enabledBook_UnselectedArray(i)
                    Case "Book_Gravoca1A_Disabled", "Book_Gravoca1B_Disabled", "Book_Gravoca1C_Disabled"
                        bookShape.Visible = msoTrue
                End Select
            Case "Ursa", "Leo"
                Select Case enabledBook_UnselectedArray(i)
                    Case "Book_Gravoca2A_Disabled", "Book_Gravoca2B_Disabled"
                        bookShape.Visible = msoTrue
                End Select
            Case "Solis", "M1-Major", "M1-Song's", "Tigris", "M2-Major", "M2-Song's"
                bookShape.Visible = (Left(bookName, 2) = "RS")
        End Select
    Next i
End Sub

Private Sub ValidateAndProceedNext(sld As Slide)
    Dim messageToDisplay As String
    
    If classTime = "NoTime" Then
        messageToDisplay = IIf(classDay = "NoDay", "Please select a class time and day before continuing.", "Please select a class time before continuing.")
        DisplayMessage messageToDisplay, "OkOnly"
    ElseIf classDay = "NoDay" Then
        messageToDisplay = "Please select a class day before continuing."
        DisplayMessage messageToDisplay, "OkOnly"
    Else
        ChangeToNewSlide configSld2
    End If
End Sub

Private Sub SaveSettings(sld As Slide)
    Dim messageToDisplay As String
    
    Select Case sld.SlideIndex
        Case configSld
            If classTime = "NoTime" Or classDay = "NoDay" Then
                messageToDisplay = IIf(classTime = "NoTime", "Please select a class time before continuing.", "Please select a class day before continuing.")
                DisplayMessage messageToDisplay, "OkOnly"
            Else
                SaveConfigAndStart
            End If
        Case configSld2
            If enabledGrade = "" Or enabledLevel = "" Or enabledBook = "" Or enabledUnit = 0 Then
                messageToDisplay = "Please select a GRADE, LEVEL, BOOK, and max UNIT before continuing."
                DisplayMessage messageToDisplay, "OkOnly"
            Else
                SaveConfigAndStart
            End If
    End Select
End Sub

Private Sub SaveConfigAndStart()
    If debugEnabled Then
        Debug.Print vbCrLf & "Saving settings and starting game."
    End If
    
    #If Mac Then
        WriteConfig
        LoadConfigMac "BeginGame"
    #Else
        WriteConfig
        LoadConfigWindows "BeginGame"
    #End If
End Sub

Private Sub ResetSettings()
    Dim messageToDisplay As String
    Dim userChoice As Integer
    
    messageToDisplay = "Are you sure you want to reset all options and start again?"
    DisplayMessage messageToDisplay, "YesNo", userChoice
    
    If userChoice = vbYes Then
        If debugEnabled Then
            Debug.Print vbCrLf & "Resetting all options and returning to first options menu."
        End If
        
        InitializeDefaultValues
        PrepareOptionsMenuStep1
    End If
End Sub

Private Sub CheckForPreviousConfig()
    Dim loadPathResponse As Boolean
    
    If classDay = "NoDay" Or classTime = "NoTime" Then
        Exit Sub
    End If
    
    If debugEnabled Then
        Debug.Print "   Checking for an existing config file for this class."
    End If
   
    loadPathResponse = LoadExistingConfig()
    
    If Not loadPathResponse Then
        Exit Sub
    End If
    
    If debugEnabled Then
        Debug.Print "   Loading existing file."
    End If
    
    #If Mac Then
        Dim chosenConfigFile As String
        chosenConfigFile = configFolder & "/AngryBirds_" & classDay & "_" & classTime & "_Config.txt"
        
        LoadConfigMac "ContinueConfig", chosenConfigFile, "PreviousConfig"
    #Else
        Dim openedFile As Object
        Set openedFile = OpenAFile("Config", "Read", classDay, classTime)
        
        LoadConfigWindows "ContinueConfig", openedFile, "PreviousConfig"
    #End If
End Sub

Private Sub NoFileFound()
    Dim messageToDisplay As String
    Dim userChoice As Integer
    
    messageToDisplay = "File not found. Would you like to continue using default values?"
    DisplayMessage messageToDisplay, "YesNo", userChoice
    
    If userChoice = vbNo Then
        SlideShowWindows(1).View.Exit
    End If
    
    StartGameWithDefaults
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'              INITIALIZE GAME
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub StartGameWithDefaults()
    GetImportantSlideNumbers
    CreateSlideArrays
    CreateBackupArrays
    ChangeToNewSlide rulesSlide
End Sub

Private Sub GetImportantSlideNumbers()
    Dim sld As Slide, shp As Shape
    Dim i As Integer, counter As Integer, sectionIndexNumber As Integer
    Dim sectionStartArray() As Integer, sectionEndArray() As Integer
    
    ReDim sectionStartArray(1 To ActivePresentation.SectionProperties.Count)
    ReDim sectionEndArray(1 To ActivePresentation.SectionProperties.Count)
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Finding slide index numbers for key slides."
    End If
    
    counter = 9
    
    For Each sld In ActivePresentation.Slides
        sectionIndexNumber = sld.sectionIndex
        
        If sectionStartArray(sectionIndexNumber) = 0 Or sld.SlideIndex < sectionStartArray(sectionIndexNumber) Then
            sectionStartArray(sectionIndexNumber) = sld.SlideIndex
        End If
        
        If sld.SlideIndex > sectionEndArray(sectionIndexNumber) Then
            sectionEndArray(sectionIndexNumber) = sld.SlideIndex
        End If
    
        If counter > 0 Then
            For Each shp In sld.Shapes
                If shp.Name = "Slide Category" Then
                    Select Case Mid(shp.TextFrame.TextRange.Text, 17)
                        Case "Rules"
                            rulesSlide = sld.SlideIndex
                            counter = counter - 1
                        Case "Gameboard"
                            If numberOfTeams = Right(sld.Shapes("Number of Teams").TextFrame.TextRange.Text, 1) Then
                                gameboardSlide = sld.SlideIndex
                                counter = counter - 1
                            End If
                        Case "Birds Points"
                            If Right(sld.Shapes("Number of Points").TextFrame.TextRange.Text, 1) = 1 Then
                                birdsPointsOne = sld.SlideIndex
                                counter = counter - 1
                            ElseIf Right(sld.Shapes("Number of Points").TextFrame.TextRange.Text, 1) = 2 Then
                                birdsPointsTwo = sld.SlideIndex
                                counter = counter - 1
                            ElseIf Right(sld.Shapes("Number of Points").TextFrame.TextRange.Text, 1) = 3 Then
                                birdsPointsThree = sld.SlideIndex
                                counter = counter - 1
                            End If
                        Case "Pigs Points"
                            If Right(sld.Shapes("Number of Points").TextFrame.TextRange.Text, 1) = 1 Then
                                pigsPointsOne = sld.SlideIndex
                                counter = counter - 1
                            ElseIf Right(sld.Shapes("Number of Points").TextFrame.TextRange.Text, 1) = 2 Then
                                pigsPointsTwo = sld.SlideIndex
                                counter = counter - 1
                            ElseIf Right(sld.Shapes("Number of Points").TextFrame.TextRange.Text, 1) = 3 Then
                                pigsPointsThree = sld.SlideIndex
                                counter = counter - 1
                            End If
                        Case "Lottery"
                            lotterySlide = sld.SlideIndex
                            counter = counter - 1
                    End Select
                End If
            Next shp
        End If
    Next sld

    For i = 1 To ActivePresentation.SectionProperties.Count
        If ActivePresentation.SectionProperties.Name(i) = "TwoTeams" And numberOfTeams = 2 Then
            firstQuestionSlide = sectionStartArray(i)
            lastQuestionSlide = sectionEndArray(i)
        ElseIf ActivePresentation.SectionProperties.Name(i) = "ThreeTeams" And numberOfTeams = 3 Then
            firstQuestionSlide = sectionStartArray(i)
            lastQuestionSlide = sectionEndArray(i)
        ElseIf ActivePresentation.SectionProperties.Name(i) = "FourTeams" And numberOfTeams = 4 Then
            firstQuestionSlide = sectionStartArray(i)
            lastQuestionSlide = sectionEndArray(i)
        ElseIf ActivePresentation.SectionProperties.Name(i) = "Lottery Events" Then
            firstLotterySlide = sectionStartArray(i)
            lastLotterySlide = sectionEndArray(i)
        End If
    Next i
    
    If debugEnabled Then
        Debug.Print "   Rules slide: " & rulesSlide
        Debug.Print "   Gameboard slide: " & gameboardSlide
        Debug.Print "   Lottery selection slide: " & lotterySlide
        Debug.Print "   Bird points slides: " & birdsPointsOne & ", " & birdsPointsTwo & ", " & birdsPointsThree
        Debug.Print "   Pig points slides: " & pigsPointsOne & ", " & pigsPointsTwo & ", " & pigsPointsThree
        Debug.Print "   Question slides range: " & firstQuestionSlide & " to " & lastQuestionSlide
        Debug.Print "   Lottery slides range: " & firstLotterySlide & " to " & lastLotterySlide
        Debug.Print "Found all required indexes."
    End If
End Sub

Private Sub CreateSlideArrays()
    Dim sld As Slide
    Dim i As Integer, j As Integer, slideCount As Integer, totalAllowedValues As Integer, selectedCount As Integer, unitValue As Integer
    Dim requiredNumberOfQuestions As Integer, finalAvailableCount As Integer, unitTextBoxValue As Integer, debugCount As Integer
    Dim randomValue As String, textCategory As String, textLevel As String, textBook As String, textQuestionID As String, textUnit As String
    Dim levelValuesArray() As String, allowedValuesArray() As Variant
    Dim alreadySeen As Boolean, foundValue As Boolean
    Dim selectionValues As New Collection
    Dim value As Variant
    
    ReDim alreadySeenSlidesArray(1 To ActivePresentation.Slides.Count)
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Randomly selecting lottery tiles."
    End If
    
    allowedValuesArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                               "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14")
    
    If numberOfLotteryTiles > 0 Then
        ReDim randomlySelectedLotteryTilesArray(1 To numberOfLotteryTiles)
        
        totalAllowedValues = UBound(allowedValuesArray)
        selectedCount = 0
        
        Do While selectedCount < numberOfLotteryTiles
            randomValue = allowedValuesArray(Int(RandBetween(1, totalAllowedValues)))
            
            foundValue = False
            For Each value In selectionValues
                If value = randomValue Then
                    foundValue = True
                    Exit For
                End If
            Next value
            
            If Not foundValue Then
                selectionValues.Add randomValue
                selectedCount = selectedCount + 1
                randomlySelectedLotteryTilesArray(selectedCount) = randomValue
            End If
        Loop
    End If
    
    If debugEnabled Then
        Debug.Print "   Selected tiles: " & Join(randomlySelectedLotteryTilesArray, ", ")
    End If
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Creating slide arrays."
    End If
    
    slideCount = ActivePresentation.Slides.Count
    
    For i = 1 To slideCount
        Set sld = ActivePresentation.Slides(i)
        alreadySeen = (sld.SlideShowTransition.Hidden = msoTrue)
        
        If Not alreadySeen Then
            If i >= firstQuestionSlide And i <= lastQuestionSlide Then
                With sld.Shapes
                    textCategory = Mid(.Item("Category").TextFrame2.TextRange.Text, 11)
                    textLevel = Mid(.Item("Level").TextFrame2.TextRange.Text, 8)
                    textBook = Mid(.Item("Book").TextFrame2.TextRange.Text, 7)
                    textQuestionID = Mid(.Item("QuestionID").TextFrame2.TextRange.Text, 13)
                End With
                
                If Not IsEmpty(previouslySeenQuestionsArray) Then
                    For j = LBound(previouslySeenQuestionsArray) To UBound(previouslySeenQuestionsArray)
                        If previouslySeenQuestionsArray(j) = textQuestionID Then
                            alreadySeenSlidesArray(i) = True
                        End If
                    Next j
                End If
                
                If Not alreadySeen Then
                    Select Case textCategory
                        Case "Trivia"
                            alreadySeen = Not enableTrivia
                        Case "Grammar"
                            If enableGrammar And textLevel <> "" Then
                                levelValuesArray = Split(textLevel, ", ")
                                alreadySeen = True
                                For j = LBound(levelValuesArray) To UBound(levelValuesArray)
                                    If levelValuesArray(j) = enabledLevel Then
                                        alreadySeen = False
                                        Exit For
                                    End If
                                Next j
                            End If
                        Case "Review"
                            If enableReview Then
                                With sld.Shapes("Unit").TextFrame2.TextRange
                                    If .Text <> "" Then
                                        textUnit = .Text
                                        unitValue = CInt(Mid(textUnit, 7))
                                        alreadySeen = Not ((textLevel = enabledLevel) And (textBook = enabledBook) And (unitValue <= enabledUnit))
                                    Else
                                        alreadySeen = True
                                    End If
                                End With
                            Else
                                alreadySeen = True
                            End If
                    End Select
                End If
                
                If debugEnabled And Not alreadySeen Then
                    debugCount = debugCount + 1
                End If
            ElseIf i >= firstLotterySlide And i <= lastLotterySlide Then
                alreadySeen = (numberOfTeams < CInt(Mid(sld.Shapes("Needed Number of Teams").TextFrame2.TextRange.Text, 15)))
            Else
                alreadySeen = True
            End If
        End If
        
        alreadySeenSlidesArray(i) = alreadySeen
    Next i
    
    If debugEnabled Then
        Debug.Print "Finished creating arrays."
    End If
    
    'VerifySufficientSlides requiredNumberOfQuestions, finalAvailableCount
    '
    'If requiredNumberOfQuestions > finalAvailableCount Then
    '    If debugEnabled Then
    '        Debug.Print "Sorry, there was an error preparing enough questions to play."
    '    End If
    '
    '    ActivePresentation.SlideShowWindow.View.Exit
    'Else
    '    If debugEnabled Then
    '        Debug.Print "There are sufficient questions available. Continuing game."
    '    End If
    'End If
End Sub

Private Sub VerifySufficientSlides(ByRef requiredNumberOfQuestions As Integer, ByRef finalAvailableCount As Integer)
    Dim i As Integer, availableQuestions As Integer, availableEasy As Integer, availableMedium As Integer, availableHard As Integer
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Verifying sufficient slides are available to play."
    End If
    
    requiredNumberOfQuestions = 40 - numberOfLotteryTiles
    
    For i = firstQuestionSlide To lastQuestionSlide
        If Not alreadySeenSlidesArray(i) Then
            Select Case Trim$(ActivePresentation.Slides(i).Shapes("Difficulty").TextFrame2.TextRange)
                Case "Difficulty: Hard"
                    availableHard = availableHard + 1
                Case "Difficulty: Medium"
                    availableMedium = availableMedium + 1
                Case "Difficulty: Easy"
                    availableEasy = availableEasy + 1
            End Select
        End If
    Next i
    
    If (availableEasy + availableMedium + availableHard) < requiredNumberOfQuestions Then
        DetermineNumberOfQuestionsToReset requiredNumberOfQuestions, availableEasy, availableMedium, availableHard
        
        For i = firstQuestionSlide To lastQuestionSlide
            If alreadySeenSlidesArray(i) = False Then
                finalAvailableCount = finalAvailableCount + 1
            End If
        Next i
    End If
    
    finalAvailableCount = availableQuestions
End Sub

Private Sub DetermineNumberOfQuestionsToReset(ByVal requiredNumberOfQuestions As Integer, ByVal availableEasy As Integer, ByVal availableMedium As Integer, ByVal availableHard As Integer)
    Dim easyToBeFreed As Integer, mediumToBeFreed As Integer, hardToBeFreed As Integer, totalToBeFreed As Integer, availableTotal As Integer
    
    availableTotal = availableEasy + availableMedium + availableHard
    easyToBeFreed = (requiredNumberOfQuestions / 3) - availableEasy
    mediumToBeFreed = (requiredNumberOfQuestions / 3) - availableMedium
    hardToBeFreed = (requiredNumberOfQuestions / 3) - availableHard
    totalToBeFreed = easyToBeFreed + mediumToBeFreed + hardToBeFreed
    
    Do While (requiredNumberOfQuestions - availableTotal) > totalToBeFreed
        If requiredNumberOfQuestions - totalToBeFreed = 1 Then
            mediumToBeFreed = mediumToBeFreed + 1
        ElseIf requiredNumberOfQuestions - totalToBeFreed = 2 Then
            mediumToBeFreed = mediumToBeFreed + 1
            hardToBeFreed = hardToBeFreed + 1
        Else
            easyToBeFreed = easyToBeFreed + 1
            mediumToBeFreed = mediumToBeFreed + 1
            hardToBeFreed = hardToBeFreed + 1
        End If
        
        totalToBeFreed = easyToBeFreed + mediumToBeFreed + hardToBeFreed
    Loop
    
    If totalToBeFreed > 0 Then
        ResetStatusOfPreviouslySeenQuestions easyToBeFreed, mediumToBeFreed, hardToBeFreed
    End If
End Sub

Private Sub ResetStatusOfPreviouslySeenQuestions(Optional ByVal easyToBeFreed As Integer = 0, Optional ByVal mediumToBeFreed As Integer = 0, Optional ByVal hardToBeFreed As Integer = 0)
    Dim randomQuestion As Integer, numberOfQuestionsToReset As Integer, i As Integer, j As Integer
    Dim resetQuestionIDs() As String
    
    numberOfQuestionsToReset = easyToBeFreed + mediumToBeFreed + hardToBeFreed
    
    If numberOfQuestionsToReset = 0 Then
        ResetAllQuestions
        Exit Sub
    End If
    
    ReDim resetQuestionIDs(1 To numberOfQuestionsToReset)
    j = 1
    
    Do
        Do
            randomQuestion = RandBetween(firstQuestionSlide, lastQuestionSlide)
        Loop While Not alreadySeenSlidesArray(randomQuestion)
        
        If randomQuestion > 0 Then
            ActivePresentation.Slides(randomQuestion).SlideShowTransition.Hidden = msoTrue
            
            Select Case ActivePresentation.Slides(randomQuestion).Shapes("Difficulty").TextFrame2.TextRange
                Case "Difficulty: Easy"
                    If ResetQuestionStatus(randomQuestion, easyToBeFreed, resetQuestionIDs) Then
                        numberOfQuestionsToReset = numberOfQuestionsToReset - 1
                    End If
                Case "Difficulty: Medium"
                    If ResetQuestionStatus(randomQuestion, mediumToBeFreed, resetQuestionIDs) Then
                        numberOfQuestionsToReset = numberOfQuestionsToReset - 1
                    End If
                Case "Difficulty: Hard"
                    If ResetQuestionStatus(randomQuestion, hardToBeFreed, resetQuestionIDs) Then
                        numberOfQuestionsToReset = numberOfQuestionsToReset - 1
                    End If
            End Select
        End If
    Loop While numberOfQuestionsToReset > 0
    
    If resetQuestionIDs(1) <> "" Then
        RepopulatePreviouslySeenQuestionsArray resetQuestionIDs()
    End If
End Sub

Private Sub ResetAllQuestions()
    Dim i As Integer
    
    ReDim previouslySeenQuestionsArray(1)
    previouslySeenQuestionsArray(1) = ""
    
    For i = firstQuestionSlide To lastQuestionSlide
        alreadySeenSlidesArray(i) = False
    Next i
End Sub

Private Function ResetQuestionStatus(ByVal randomQuestion As Integer, ByRef requiredCount As Integer, ByRef resetQuestionIDs() As String) As Boolean
    If requiredCount > 0 Then
        requiredCount = requiredCount - 1
        alreadySeenSlidesArray(randomQuestion) = False
        resetQuestionIDs(requiredCount + 1) = "QuestionID: " & ActivePresentation.Slides(randomQuestion).Shapes("QuestionID").TextFrame2.TextRange
        ResetQuestionStatus = True
    End If
End Function

'Something going wrong here. Perhaps to do with resetQuestionIDs?
Private Sub RepopulatePreviouslySeenQuestionsArray(ByRef resetQuestionIDs() As String)
    Dim resultArray() As String
    Dim i As Long, j As Long, k As Long
    Dim found As Boolean
    
    ReDim resultArray(1 To UBound(previouslySeenQuestionsArray))

    k = 1
    
    For i = LBound(previouslySeenQuestionsArray) To UBound(previouslySeenQuestionsArray)
        found = False

        For j = LBound(resetQuestionIDs) To UBound(resetQuestionIDs)
            If previouslySeenQuestionsArray(i) = resetQuestionIDs(j) Then
                found = True
                Exit For
            End If
        Next j

        If Not found Then
            resultArray(k) = previouslySeenQuestionsArray(i)
            k = k + 1
        End If
    Next i

    ReDim Preserve resultArray(1 To k - 1)
    previouslySeenQuestionsArray = resultArray
End Sub

Private Sub CreateBackupArrays()
    Dim i As Integer
    
    ReDim variableNamesArray(1 To 14)
    variableNamesArray(1) = "numberOfTeams"
    variableNamesArray(2) = "numberOfLotteryTiles"
    variableNamesArray(3) = "autoAddPoints"
    variableNamesArray(4) = "allowNegatives"
    variableNamesArray(5) = "enableTrivia"
    variableNamesArray(6) = "enableGrammar"
    variableNamesArray(7) = "enableReview"
    variableNamesArray(8) = "enableChallenges"
    variableNamesArray(9) = "classTime"
    variableNamesArray(10) = "classDay"
    variableNamesArray(11) = "enabledGrade"
    variableNamesArray(12) = "enabledLevel"
    variableNamesArray(13) = "enabledBook"
    variableNamesArray(14) = "enabledUnit"
    
    ReDim currentStateNamesArray(1 To 8)
    currentStateNamesArray(1) = "classDay"
    currentStateNamesArray(2) = "classTime"
    currentStateNamesArray(3) = "numberOfTeams"
    currentStateNamesArray(4) = "team1Score"
    currentStateNamesArray(5) = "team2Score"
    currentStateNamesArray(6) = "team3Score"
    currentStateNamesArray(7) = "team4Score"
    currentStateNamesArray(8) = "marmosetsScore"

    ReDim variableValuesArray(1 To 14)
    variableValuesArray(1) = numberOfTeams
    variableValuesArray(2) = numberOfLotteryTiles
    variableValuesArray(3) = autoAddPoints
    variableValuesArray(4) = allowNegatives
    variableValuesArray(5) = enableTrivia
    variableValuesArray(6) = enableGrammar
    variableValuesArray(7) = enableReview
    variableValuesArray(8) = enableChallenges
    variableValuesArray(9) = classTime
    variableValuesArray(10) = classDay
    variableValuesArray(11) = enabledGrade
    variableValuesArray(12) = enabledLevel
    variableValuesArray(13) = enabledBook
    variableValuesArray(14) = enabledUnit
    
    ReDim currentStateValuesArray(1 To 8)
    currentStateValuesArray(1) = classDay
    currentStateValuesArray(2) = classTime
    currentStateValuesArray(3) = numberOfTeams
    currentStateValuesArray(4) = team1Score
    currentStateValuesArray(5) = team2Score
    currentStateValuesArray(6) = team3Score
    currentStateValuesArray(7) = team4Score
    currentStateValuesArray(8) = marmosetsScore
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'             CONFIG SAVE/LOAD SUBS
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub WriteConfig()
    If debugEnabled Then
        Debug.Print "Beginning to write the config. file."
    End If
    
    #If Mac Then
        Dim configFileName As String, dataToSave As String, writeResults As String
        
        configFileName = "AngryBirds_" & classDay & "_" & classTime & "_Config.txt"
        
        dataToSave = "numberOfTeams = " & CStr(numberOfTeams) & "&" _
                   & "numberOfLotteryTiles = " & CStr(numberOfLotteryTiles) & "&" _
                   & "autoAddPoints = " & CStr(autoAddPoints) & "&" _
                   & "allowNegatives = " & CStr(allowNegatives) & "&" _
                   & "enableTrivia = " & CStr(enableTrivia) & "&" _
                   & "enableGrammar = " & CStr(enableGrammar) & "&" _
                   & "enableReview = " & CStr(enableReview) & "&" _
                   & "enableChallenges = " & CStr(enableChallenges) & "&" _
                   & "classTime = " & classTime & "&" _
                   & "classDay = " & classDay & "&" _
                   & "enabledGrade = " & enabledGrade & "&" _
                   & "enabledLevel = " & enabledLevel & "&" _
                   & "enabledBook = " & enabledBook & "&" _
                   & "enabledUnit = " & CStr(enabledUnit) & "&" _
                   & "previouslySeenQuestionsArray = " & Join(previouslySeenQuestionsArray, ", ")
        
        On Error Resume Next
        writeResults = AppleScriptTask("AngryBirds.scpt", "WriteToFile", configFileName & ";" & dataToSave)
        On Error GoTo 0
        
        If debugEnabled Then
            Debug.Print "Save successful: " & writeResults
        End If
    #Else
        Dim openedFile As Object
        Set openedFile = OpenAFile("Config", "Write")
        
        If openedFile Is Nothing Then
            If debugEnabled Then
                Debug.Print "Critical error writing to file. Terminating game."
            End If
        
            ActivePresentation.SlideShowWindow.View.Exit
        End If

        openedFile.WriteLine "numberOfTeams = " & CStr(numberOfTeams)
        openedFile.WriteLine "numberOfLotteryTiles = " & CStr(numberOfLotteryTiles)
        openedFile.WriteLine "autoAddPoints = " & CStr(autoAddPoints)
        openedFile.WriteLine "allowNegatives = " & CStr(allowNegatives)
        openedFile.WriteLine "enableTrivia = " & CStr(enableTrivia)
        openedFile.WriteLine "enableGrammar = " & CStr(enableGrammar)
        openedFile.WriteLine "enableReview = " & CStr(enableReview)
        openedFile.WriteLine "enableChallenges = " & CStr(enableChallenges)
        openedFile.WriteLine "classTime = " & classTime
        openedFile.WriteLine "classDay = " & classDay
        openedFile.WriteLine "enabledGrade = " & enabledGrade
        openedFile.WriteLine "enabledLevel = " & enabledLevel
        openedFile.WriteLine "enabledBook = " & enabledBook
        openedFile.WriteLine "enabledUnit = " & CStr(enabledUnit)
        openedFile.WriteLine "previouslySeenQuestionsArray = " & Join(previouslySeenQuestionsArray, ", ")
        
        CloseTheFile openedFile
        
        If debugEnabled Then
            Debug.Print "Finished writing temporary config. file."
        End If
        
        MoveSavedFile "Config"
    #End If
End Sub

#If Mac Then
Private Sub LoadConfigMac(ByVal loadingOption As String, Optional ByRef fileToOpen As String, Optional ByVal loadPreviousConfig As String = "")
    Dim configLine As String, lineToPrint As String
    Dim configLineValuesArray() As String
    Dim i As Integer

    If loadingOption <> "BeginGame" Then
        configLine = AppleScriptTask("AngryBirds.scpt", "LoadConfig", fileToOpen)
        configLineValuesArray() = Split(configLine, ";")
                    
        SetVariablesFromLoadedFile configLineValuesArray
    End If
    
    If debugEnabled Then
        lineToPrint = "Previously Seen Questions: " & vbCrLf & "   " & Join(previouslySeenQuestionsArray, vbCrLf & "   ")
        
        If Right(lineToPrint, 2) = ", " Then
            lineToPrint = Left(lineToPrint, Len(lineToPrint) - 2)
        End If
        
        Debug.Print vbCrLf & lineToPrint
    End If
    
    Select Case loadingOption
        Case "ContinueConfig"
            PrepareOptionsMenuStep1 loadPreviousConfig
        Case "ExistingConfig", "BeginGame"
            GetImportantSlideNumbers
            CreateSlideArrays
            CreateBackupArrays
            ChangeToNewSlide rulesSlide
    End Select
End Sub
#End If

#If Mac Then
#Else
Private Sub LoadConfigWindows(ByVal loadingOption As String, Optional ByRef openedFile As Object, Optional ByVal loadPreviousConfig As String = "")
    Dim lineToPrint As String, configLine As String
    Dim configLineValuesArray() As String
    Dim i As Integer
    
    If loadingOption <> "BeginGame" Then
        Do While Not openedFile.AtEndOfStream
            configLine = configLine & openedFile.ReadLine & ";"
        Loop
        CloseTheFile openedFile
        
        configLineValuesArray() = Split(configLine, ";")
        
        SetVariablesFromLoadedFile configLineValuesArray
    End If
    
    If debugEnabled Then
        lineToPrint = "Previously Seen Questions: " & vbCrLf & "   " & Join(previouslySeenQuestionsArray, vbCrLf & "   ")
        
        If Right(lineToPrint, 2) = ", " Then
            lineToPrint = Left(lineToPrint, Len(lineToPrint) - 2)
        End If
        
        Debug.Print vbCrLf & lineToPrint
    End If

    Select Case loadingOption
        Case "ContinueConfig"
            PrepareOptionsMenuStep1 loadPreviousConfig
        Case "ExistingConfig", "BeginGame"
            GetImportantSlideNumbers
            CreateSlideArrays
            CreateBackupArrays
            ChangeToNewSlide rulesSlide
    End Select
End Sub
#End If

Private Sub RestoreGame()
    Dim restoreFilePath As String, inputLine As String, configLines As String, configToOpen As String, optionKey As String, optionValue As String
    Dim configLineValuesArray() As String, questionValuesArray() As String, configValuesArray() As String
    Dim i As Integer, j As Integer, counter As Integer
    Dim fileExists As Boolean
        
    #If Mac Then
        restoreFilePath = AppleScriptTask("AngryBirds.scpt", "SetTempDirectory", "noParam") & "AngryBirds_CurrentState.txt"
        
        fileExists = AppleScriptTask("AngryBirds.scpt", "ExistsFile", restoreFilePath)
        
        If Not fileExists Then
            NoFileFound
            Exit Sub
        End If
        
        configLines = AppleScriptTask("AngryBirds.scpt", "LoadConfig", restoreFilePath)
        configLineValuesArray = Split(configLines, ";")
        
        DetermineClassSettingsToRestore configLineValuesArray
        
        If classDay <> "NoDay" And classTime <> "NoTime" Then
            configToOpen = configFolder & "/AngryBirds_" & classDay & "_" & classTime & "_Config.txt"
            fileExists = AppleScriptTask("AngryBirds.scpt", "ExistsFile", configToOpen)
        
            If Not fileExists Then
                NoFileFound
                Exit Sub
            End If
        
            LoadConfigMac "RestoreGame", configToOpen
        Else
            InitializeDefaultValues
        End If
    #Else
        Dim openedFile As Object
        
        If debugEnabled Then
            Debug.Print "Beginning to load restore data."
        End If
        
        restoreFilePath = tempFolder & "\AngryBirds_CurrentState.txt"
        
        If VerifyFileOrFolderExists(restoreFilePath) Then
            Set openedFile = OpenAFile("CurrentState", "Read")
            
            If openedFile Is Nothing Then
                NoFileFound
                Exit Sub
            End If
        End If
        
        If debugEnabled Then
            Debug.Print "Opened: " & tempFolder & "\AngryBirds_CurrentState.txt"
        End If
        
        Do While Not openedFile.AtEndOfStream
            configLines = configLines & openedFile.ReadLine & ";"
        Loop
        
        CloseTheFile openedFile
        
        configLineValuesArray = Split(configLines, ";")
        DetermineClassSettingsToRestore configLineValuesArray
        
        If classDay <> "NoDay" And classTime <> "NoTime" Then
            Set openedFile = OpenAFile("Config", "Read", classDay, classTime)
            
            If openedFile Is Nothing Then
                NoFileFound
                Exit Sub
            End If
            
            If debugEnabled Then
                Debug.Print "Loading data for: " & classDay & "_" & classTime
            End If
    
            LoadConfigWindows "RestoreGame", openedFile
        Else
            InitializeDefaultValues
        End If
    #End If
    
    SetVariablesFromLoadedFile configLineValuesArray, "Restore"
    CreateBackupArrays
    UpdateScore "All", 0, True
    ChangeToNewSlide gameboardSlide
End Sub

Private Sub DetermineClassSettingsToRestore(ByRef configLineValuesArray() As String)
    Dim optionKey As String, optionValue As String
    Dim configValuesArray() As String
    Dim i As Integer, counter As Integer
    
    For i = 0 To UBound(configLineValuesArray)
        configValuesArray = Split(configLineValuesArray(i), " = ")
            
        optionKey = Trim(configValuesArray(0))
        optionValue = Trim(configValuesArray(1))
    
        Select Case optionKey
            Case "classDay"
                classDay = optionValue
                counter = counter + 1
            Case "classTime"
                classTime = optionValue
                counter = counter + 1
        End Select
        
        If counter = 2 Then
            Exit For
        End If
    Next i
End Sub

Private Sub SetVariablesFromLoadedFile(ByRef configLineValuesArray() As String, Optional ByVal loadType As String = "")
    Dim optionKey As String, optionValue As String
    Dim configValuesArray() As String, questionValuesArray() As String
    Dim i As Integer, j As Integer

    ReDim alreadySeenSlidesArray(1 To ActivePresentation.Slides.Count)
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Loading settings from config file."
    End If
    
    For j = 0 To UBound(configLineValuesArray)
        If configLineValuesArray(j) <> "" Then
            configValuesArray = Split(configLineValuesArray(j), " = ")
        End If
        
        If debugEnabled And configLineValuesArray(j) <> "" Then
            Debug.Print "   " & configLineValuesArray(j)
        End If
        
        optionKey = Trim(configValuesArray(0))
        optionValue = Trim(configValuesArray(1))
    
        Select Case optionKey
            Case "numberOfTeams"
                numberOfTeams = CInt(optionValue)

                If loadType = "Restore" Then
                    GetImportantSlideNumbers

                    For i = 1 To ActivePresentation.Slides.Count
                        alreadySeenSlidesArray(i) = Not ((i >= firstQuestionSlide And i <= lastQuestionSlide) Or (i >= firstLotterySlide And i <= lastLotterySlide))
                        'Set all slideIndex numbers outside of the question and lottery ranges as True as added protection against them being accidentally selected to be shown
                    Next i
                End If
            Case "numberOfLotteryTiles"
                numberOfLotteryTiles = CInt(optionValue)
                
                ReDim randomlySelectedLotteryTilesArray(1 To numberOfLotteryTiles)
            Case "autoAddPoints"
                autoAddPoints = CBool(optionValue)
            Case "allowNegatives"
                allowNegatives = CBool(optionValue)
            Case "enableTrivia"
                enableTrivia = CBool(optionValue)
            Case "enableGrammar"
                enableGrammar = CBool(optionValue)
            Case "enableReview"
                enableReview = CBool(optionValue)
            Case "enableChallenges"
                enableChallenges = CBool(optionValue)
            Case "classTime"
                classTime = optionValue
            Case "classDay"
                classDay = optionValue
            Case "enabledGrade"
                enabledGrade = optionValue
            Case "enabledLevel"
                enabledLevel = optionValue
            Case "enabledBook"
                enabledBook = optionValue
            Case "enabledUnit"
                enabledUnit = CInt(optionValue)
            Case "team1Score"
                team1Score = CInt(optionValue)
            Case "team2Score"
                team2Score = CInt(optionValue)
            Case "team3Score"
                team3Score = IIf(numberOfTeams > 2, CInt(optionValue), 0)
            Case "team4Score"
                team4Score = IIf(numberOfTeams > 3, CInt(optionValue), 0)
            Case "marmosetsScore"
                marmosetsScore = IIf(numberOfTeams < 4, CInt(optionValue), 0)
            Case "alreadySeenSlidesArray"
                If optionValue <> "" Then
                    questionValuesArray = Split(optionValue, ", ")
                    
                    For i = 0 To UBound(questionValuesArray)
                        alreadySeenSlidesArray(questionValuesArray(i)) = True
                    Next i
                End If
            Case "randomlySelectedLotteryTilesArray"
                If optionValue <> "" Then
                    questionValuesArray = Split(optionValue, ", ")
                    
                    ReDim randomlySelectedLotteryTilesArray(1 To UBound(questionValuesArray) + 1)
                    
                    For i = 0 To UBound(questionValuesArray)
                        randomlySelectedLotteryTilesArray(i + 1) = questionValuesArray(i)
                    Next i
                End If
            Case "Hidden Gameboard Tiles"
                If optionValue <> "" Then
                    questionValuesArray = Split(optionValue, ", ")
                    
                    For i = 0 To UBound(questionValuesArray)
                        ActivePresentation.Slides(gameboardSlide).Shapes(questionValuesArray(i)).Visible = msoFalse
                    Next i
                End If
            Case "Hidden Lottery Tiles"
                If optionValue <> "" Then
                    questionValuesArray = Split(optionValue, ", ")
                    
                    For i = 0 To UBound(questionValuesArray)
                        ActivePresentation.Slides(lotterySlide).Shapes(questionValuesArray(i)).Visible = msoFalse
                    Next i
                End If
            Case "previouslySeenQuestionsArray"
                If optionValue <> "" Then
                    questionValuesArray = Split(optionValue, ", ")
                    
                    ReDim previouslySeenQuestionsArray(1 To UBound(questionValuesArray) + 1)
                    
                    For i = 0 To UBound(questionValuesArray)
                        previouslySeenQuestionsArray(i + 1) = questionValuesArray(i)
                    Next i
                End If
        End Select
    Next j
    
    If debugEnabled Then
        Debug.Print "Loading complete."
    End If
End Sub

Private Sub PrepareCurrentStateData()
    Dim outputBuffer As String
    Dim i As Integer
    
    outputBuffer = "" 'Ensure that this value is cleared each time it is called to avoid unexpected behaviour
    
    If (classDay <> "NoDay" And classDay <> "") And (classTime <> "NoTime" And classTime <> "") Then
        If debugEnabled Then
            Debug.Print "Updating class config file with new questionIDs of seen questions."
        End If
        
        WriteConfig
    End If
    
    UpdateCurrentStateVariables
    
    #If Mac Then
        Dim fileName As String, dataToSave As String, paramString As String
        Dim writeResults As Boolean
        
        fileName = "AngryBirds_CurrentState.txt"
        
        If debugEnabled Then
            Debug.Print "Collecting current state data."
        End If
        
        For i = LBound(currentStateNamesArray) To UBound(currentStateNamesArray)
            dataToSave = dataToSave & currentStateNamesArray(i) & " = " & currentStateValuesArray(i) & "&"
        Next i
        
        CollectSaveData "AlreadySeen", "alreadySeenSlidesArray = ", dataToSave
        CollectSaveData "RandomlySelected", "randomlySelectedLotteryTilesArray = ", dataToSave
        CollectSaveData "gameboardSlide", "Hidden Gameboard Tiles = ", dataToSave
        CollectSaveData "lotterySlide", "Hidden Lottery Tiles = ", dataToSave
        
        paramString = fileName & ";" & dataToSave
        
        If Right(paramString, 1) = "&" Then
            paramString = Left(paramString, Len(paramString) - 1) ' Trim the final "&" from the string
        End If
        
        If debugEnabled Then
            Debug.Print "Data collection complete. Attempting to write " & fileName
        End If
        
        On Error Resume Next
        writeResults = AppleScriptTask("AngryBirds.scpt", "WriteToFile", paramString)
        On Error GoTo 0
        
        If debugEnabled Then
            Debug.Print IIf(writeResults, "Writing successful.", "Writing failed.")
        End If
    #Else
        Dim openedFile As Object
        Set openedFile = OpenAFile("CurrentState", "Write")
        
        If openedFile Is Nothing Then
            MsgBox "There was an error creating the config file."
            ActivePresentation.SlideShowWindow.View.Exit
        End If
        
        For i = LBound(currentStateNamesArray) To UBound(currentStateNamesArray)
            openedFile.WriteLine currentStateNamesArray(i) & " = " & currentStateValuesArray(i)
        Next i
        
        WriteDataToFile openedFile, "AlreadySeen", "alreadySeenSlidesArray = "
        WriteDataToFile openedFile, "RandomlySelected", "randomlySelectedLotteryTilesArray = "
        WriteDataToFile openedFile, "gameboardSlide", "Hidden Gameboard Tiles = "
        WriteDataToFile openedFile, "lotterySlide", "Hidden Lottery Tiles = "
        
        CloseTheFile openedFile
    #End If
End Sub

#If Mac Then
Private Sub CollectSaveData(ByVal outputBufferToCreate As String, ByVal outputText As String, ByRef dataToSave As String)
    Dim outputBuffer As String
    
    Select Case outputBufferToCreate
        Case "AlreadySeen", "RandomlySelected"
            outputBuffer = CreateOutputBuffer(outputBufferToCreate)
        Case "gameboardSlide"
            outputBuffer = WriteHiddenShapesInfo(ActivePresentation.Slides(gameboardSlide))
        Case "lotterySlide"
            outputBuffer = WriteHiddenShapesInfo(ActivePresentation.Slides(lotterySlide))
    End Select
        
    If Len(outputBuffer) > 0 Then
        dataToSave = dataToSave & outputText & outputBuffer & "&"
    End If
End Sub
#Else
Private Sub WriteDataToFile(ByRef openedFile As Object, ByVal outputBufferToCreate As String, ByVal outputText As String)
    Dim outputBuffer As String
    
    Select Case outputBufferToCreate
        Case "AlreadySeen", "RandomlySelected"
            outputBuffer = CreateOutputBuffer(outputBufferToCreate)
        Case "gameboardSlide"
            outputBuffer = WriteHiddenShapesInfo(ActivePresentation.Slides(gameboardSlide))
        Case "lotterySlide"
            outputBuffer = WriteHiddenShapesInfo(ActivePresentation.Slides(lotterySlide))
    End Select

    If Len(outputBuffer) > 0 Then
        openedFile.WriteLine outputText & outputBuffer
    End If
End Sub
#End If

Private Function WriteHiddenShapesInfo(ByVal sld As Slide) As String
    Dim shp As Shape
    Dim outputBuffer As String
    Dim existingValuesArray() As String
    Dim i As Integer

    For Each shp In sld.Shapes
        With shp
            If .Visible = msoFalse Then
                If Len(outputBuffer) = 0 Then
                    outputBuffer = .Name & ", "
                ElseIf Len(outputBuffer) > 0 Then
                    existingValuesArray = Split(outputBuffer, ", ")
                
                    If Not IsValuePresentInArray(.Name, existingValuesArray) Then
                        outputBuffer = outputBuffer & .Name & ", "
                    End If
                End If
            End If
        End With
    Next shp
    
    If sld.TimeLine.MainSequence.Count > 0 Then
        For i = sld.TimeLine.MainSequence.Count To 1 Step -1
            With sld.TimeLine.MainSequence.Item(i)
                If .EffectType = msoAnimEffectRandomBars Then
                    If Len(outputBuffer) = 0 Then
                        outputBuffer = .Shape.Name & ", "
                    ElseIf Len(outputBuffer) > 0 Then
                        existingValuesArray = Split(outputBuffer, ", ")
                    
                        If Not IsValuePresentInArray(.Shape.Name, existingValuesArray) Then
                            outputBuffer = outputBuffer & .Shape.Name & ", "
                        End If
                    End If
                End If
            End With
        Next i
    End If
    
    If Len(outputBuffer) > 0 Then
        outputBuffer = Left(outputBuffer, Len(outputBuffer) - 2) ' Trim the final comma and space
        WriteHiddenShapesInfo = outputBuffer
    Else
        WriteHiddenShapesInfo = ""
    End If
End Function

Private Function CreateOutputBuffer(ByVal listType As String) As String
    Dim outputBuffer As String, listSeparator As String
    Dim i As Integer
    
    outputBuffer = "" 'Ensure buffer is empty to prevent unexpected behaviour
    listSeparator = ", "
    
    Select Case listType
        Case "AlreadySeen"
            For i = 1 To ActivePresentation.Slides.Count
                If alreadySeenSlidesArray(i) And ((i >= firstQuestionSlide And i <= lastQuestionSlide) Or (i >= firstLotterySlide And i <= lastLotterySlide)) Then
                    If outputBuffer <> "" Then
                        outputBuffer = outputBuffer & i & listSeparator
                    Else
                        outputBuffer = i & listSeparator
                    End If
                End If
            Next i
        Case "RandomlySelected"
            For i = LBound(randomlySelectedLotteryTilesArray) To UBound(randomlySelectedLotteryTilesArray)
                If outputBuffer <> "" Then
                    outputBuffer = outputBuffer & randomlySelectedLotteryTilesArray(i) & listSeparator
                Else
                    outputBuffer = randomlySelectedLotteryTilesArray(i) & listSeparator
                End If
            Next i
    End Select
    
    If Len(outputBuffer) > 0 Then
        outputBuffer = Left(outputBuffer, Len(outputBuffer) - 2) ' Trim the final comma and space
    End If
    
    CreateOutputBuffer = outputBuffer
End Function

#If Mac Then
#Else
Private Function OpenAFile(ByVal fileType As String, ByVal fileMode As String, Optional ByVal classDay As String = "", Optional ByVal classTime As String = "") As Object
    Dim fs As Object, textFile As Object
    Dim filePath As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Select Case fileType
        Case "CurrentState"
            filePath = tempFolder & "\AngryBirds_" & fileType & ".txt"
        Case "FileHashes"
            filePath = tempFolder & "\FileHashes.txt"
        Case Else
            If classDay = "" Or classTime = "" Then
                filePath = tempFolder & "\AngryBirds_Config_Temp.txt"
            Else
                filePath = configFolder & "\AngryBirds_" & classDay & "_" & classTime & "_Config.txt"
            End If
    End Select
    
    Select Case fileMode
        Case "Read"
            If Not fs.fileExists(filePath) Then
                Set OpenAFile = Nothing
            Else
                Set textFile = fs.OpenTextFile(filePath, 1, False, -1)
            End If
        Case "Write"
            Set textFile = fs.OpenTextFile(filePath, 2, True, -1)
    End Select
    
    If textFile Is Nothing Then
        Set OpenAFile = Nothing
    Else
        Set OpenAFile = textFile
    End If
    
    Set fs = Nothing
End Function
#End If

#If Mac Then
#Else
Private Sub CloseTheFile(ByRef textFile As Object)
    If Not textFile Is Nothing Then
        textFile.Close
        Set textFile = Nothing
    End If
End Sub
#End If

#If Mac Then
#Else
Private Sub MoveSavedFile(ByVal fileType As String)
    Dim tempFilePath As String, newFilePath As String
    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If debugEnabled Then
        Debug.Print "Attempting to move file to the config folder."
    End If
    
    tempFilePath = tempFolder & pathSeparator & "AngryBirds_" & fileType & "_Temp.txt"
    newFilePath = configFolder & pathSeparator & "AngryBirds_" & classDay & "_" & classTime & "_" & fileType & ".txt"
    
    If fs.fileExists(newFilePath) And fs.fileExists(tempFilePath) Then
        If debugEnabled Then
            Debug.Print "Previous version deleted."
        End If
        
        Kill newFilePath
    End If
    
    If fs.fileExists(tempFilePath) Then
        fs.MoveFile tempFilePath, newFilePath
        
        If fs.fileExists(newFilePath) And debugEnabled Then
            Debug.Print "Config. file successfully saved to the config folder."
        End If
    End If
    
    Set fs = Nothing
End Sub
#End If

Private Function ChooseSaveLocation(Optional ByVal versionNumber As String = "") As String
    If debugEnabled Then
        Debug.Print "Selecting save location."
    End If
     
     #If Mac Then
        Dim chosenPath As String
        
        chosenPath = AppleScriptTask("AngryBirds.scpt", "ChooseSaveLocation", "noParam")

        If chosenPath <> "" Then
            ChooseSaveLocation = chosenPath
        End If
    #Else
        Dim saveFileDialog As FileDialog
        Set saveFileDialog = Application.FileDialog(msoFileDialogSaveAs)
        
        With saveFileDialog ' Configure the dialog box
            .Title = "Save File As"
            .FilterIndex = 2
            .InitialFileName = ActivePresentation.Path & "\Angry Birds Trivia v" & versionNumber & ".pptm" ' Set a default file name if needed
            
            If .Show = -1 Then
                ChooseSaveLocation = .SelectedItems(1)
                ChooseSaveLocation = ConvertOneDriveToLocalPath(ChooseSaveLocation)
            Else
                ChooseSaveLocation = "" ' User canceled the dialog
            End If
        End With
    #End If
    
    If debugEnabled Then
        Debug.Print "Save location: " & ChooseSaveLocation
    End If
End Function

Private Function ChooseFileToLoad() As String
    If debugEnabled Then
        Debug.Print "Selecting file to load."
    End If
    
    #If Mac Then
        Dim returnedPath As String
        returnedPath = AppleScriptTask("AngryBirds.scpt", "ChooseFileToOpen", "noParam")
        
        ChooseFileToLoad = returnedPath
    #Else
        Dim configFileDialog As FileDialog
        Set configFileDialog = Application.FileDialog(msoFileDialogFilePicker)
        
        With configFileDialog ' Configure the dialog box
            .Title = "Select a Config File"
            .Filters.Clear
            .Filters.Add "Class Config Files", "*.txt"
            .AllowMultiSelect = False
            .InitialFileName = configFolder
            
            ChooseFileToLoad = IIf(.Show = -1, .SelectedItems(1), "")
        End With
    #End If
    
    Debug.Print IIf(ChooseFileToLoad <> "", "Selected file: " & ChooseFileToLoad, "No file was selected.")
End Function

Private Function LoadExistingConfig() As Boolean
    Dim configToCheckFor As String, messageToDisplay As String
    Dim userChoice As Integer
    
    configToCheckFor = configFolder & pathSeparator & "AngryBirds_" & classDay & "_" & classTime & "_Config.txt"
    messageToDisplay = "A configuration file already existed for this class. Would you like to load it?"
    
    If debugEnabled Then
        Debug.Print "   Searching for: " & configToCheckFor
    End If
    
    #If Mac Then
        Dim doesFileExist As Boolean
        doesFileExist = AppleScriptTask("AngryBirds.scpt", "ExistsFile", configToCheckFor)
        
        If doesFileExist Then
            DisplayMessage messageToDisplay, "YesNo", userChoice
            LoadExistingConfig = (userChoice = vbYes)
        Else
            LoadExistingConfig = False
        End If
    #Else
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        If fs.fileExists(configToCheckFor) Then
            DisplayMessage messageToDisplay, "YesNo", userChoice
            LoadExistingConfig = (userChoice = vbYes)
        Else
            LoadExistingConfig = False
        End If
    #End If
    
    If debugEnabled Then
        Debug.Print IIf(LoadExistingConfig, "   Config file found.", "   Config file either ignored or not found.")
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''
'                 GAMEPLAY SUBS                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeToNewSlide(ByVal newSlideIndex As Integer)
    SlideShowWindows(1).View.GotoSlide newSlideIndex
End Sub

Public Sub ChooseQuestion(clickedShp As Shape)
    Dim tileName As String, messageToDisplay As String
    Dim userChoice As Integer, i As Integer, tbLeft As Integer
    Dim sld As Slide
    
    Const tbHeight As Integer = 105
    Const tbWidth As Integer = 104
    
    tileName = clickedShp.Name
    
    If tileName = "Marmoset (Question)" Then
        messageToDisplay = "Would you like to load a new question?"
        DisplayMessage messageToDisplay, "YesNo", userChoice
        
        If userChoice = vbNo Then
            Exit Sub
        End If
        
        If debugEnabled Then
            Debug.Print vbCrLf & "Selecting a new question to display."
        End If

        AddToPreviouslySeen
    Else
        chosenTile = tileName
        
        If debugEnabled Then
            Debug.Print vbCrLf & "Chosen tile: " & chosenTile
        End If

        If numberOfLotteryTiles > 0 Then
            For i = 1 To numberOfLotteryTiles
                If chosenTile = randomlySelectedLotteryTilesArray(i) Then
                    If debugEnabled Then
                        Debug.Print "Tile was a randomly selected lottery tile. Moving to the lottery slide."
                    End If
                    
                    If chosenLottery <> "" Then
                        RemoveShapes "Animated", "LotteryTile"
                    End If
                    
                    ChangeToNewSlide lotterySlide
                    Exit Sub
                End If
            Next i
        End If
    End If
    
    Do
        chosenSlide = RandBetween(firstQuestionSlide, lastQuestionSlide)
    Loop While alreadySeenSlidesArray(chosenSlide)

    If debugEnabled Then
        Debug.Print "Moving to trivia at slide " & chosenSlide & " and adding questionID '" & chosenTile & "' to slide."
    End If
    
    Set sld = ActivePresentation.Slides(chosenSlide)
    tbLeft = sld.Master.Width - tbWidth

    With sld.Shapes.AddTextbox(msoOrientationHorizontal, tbLeft, 0, tbWidth, tbHeight)
        .Name = "ChosenQuestionTile"
        .TextFrame.TextRange.Text = chosenTile
        .TextFrame.TextRange.Font.Name = "Feast of Flesh BB"
        .TextFrame.TextRange.Font.Size = 96
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame.TextRange.Font.Bold = msoTrue
        .TextFrame.TextRange.Font.Italic = msoTrue
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    End With

    ChangeToNewSlide chosenSlide
    PrepareCurrentStateData
    alreadySeenSlidesArray(chosenSlide) = True
End Sub

Public Sub AwardEggs(clickedShp As Shape)
    Dim shapeName As String, difficultyLevel As String, teamName As String
    Dim shapeNamePartsArray() As String
    Dim pointsToAward As Integer, pointsSlide As Integer
    Dim sldView As SlideShowView
    
    Set sldView = ActivePresentation.SlideShowWindow.View
    
    shapeName = clickedShp.Name

    shapeNamePartsArray = Split(ActivePresentation.Slides(chosenSlide).Shapes("Difficulty").TextFrame2.TextRange.Text, ": ")
    difficultyLevel = shapeNamePartsArray(1)
    shapeNamePartsArray = Split(shapeName, "_")
    teamName = shapeNamePartsArray(0)
    
    Select Case difficultyLevel
        Case "Easy"
            pointsToAward = 1
        Case "Medium"
            pointsToAward = 2
        Case "Hard"
            pointsToAward = 3
    End Select
    
    If doublePoints > 0 Then
        pointsToAward = pointsToAward * 2
        doublePoints = doublePoints - 1
        
        If debugEnabled Then
            Debug.Print vbCrLf & "Point multiplier enabled. Doubling points."
            Debug.Print "Remaining: " & doublePoints
        End If
    End If
    
    If teamName = "Team1" Or teamName = "Team3" Then
        Select Case difficultyLevel
            Case "Easy"
                pointsSlide = birdsPointsOne
            Case "Medium"
                pointsSlide = birdsPointsTwo
            Case "Hard"
                pointsSlide = birdsPointsThree
        End Select
    ElseIf teamName = "Team2" Or teamName = "Team4" Then
        Select Case difficultyLevel
            Case "Easy"
                pointsSlide = pigsPointsOne
            Case "Medium"
                pointsSlide = pigsPointsTwo
            Case "Hard"
                pointsSlide = pigsPointsThree
        End Select
    End If
    
    If autoAddPoints Then
        UpdateScore teamName, pointsToAward
    End If
    
    If debugEnabled And autoAddPoints Then
        Debug.Print vbCrLf & "Awarding " & pointsToAward & " egg(s) to " & teamName
    End If
    
    If ActivePresentation.Slides(chosenSlide).SlideIndex >= firstQuestionSlide And ActivePresentation.Slides(chosenSlide).SlideIndex <= lastQuestionSlide Then
        AddToPreviouslySeen
    End If
    
    If debugEnabled And autoAddPoints Then
        Debug.Print vbCrLf & "Moving to points slide."
    End If
    
    sldView.GotoSlide pointsSlide
End Sub

Public Sub LotteryEvent(clickedShp As Shape)
    chosenLottery = clickedShp.Name

    Do
        chosenSlide = RandBetween(firstLotterySlide, lastLotterySlide)
    Loop While alreadySeenSlidesArray(chosenSlide)

    PrepareCurrentStateData
    
    alreadySeenSlidesArray(chosenSlide) = True

    If ActivePresentation.Slides(chosenSlide).Shapes("Lottery Event").TextFrame2.TextRange.Text = "Lottery Event: Points Multiplier" Then
        doublePoints = 3
    End If
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Chosen lottery tile: " & chosenLottery
        Debug.Print "Moving to lottery slide: " & chosenSlide
    End If

    ChangeToNewSlide chosenSlide
End Sub

Public Sub ReturnToGameboard(clickedShp As Shape)
    Dim sld As Slide
    Dim shp As Shape
    Dim shapeName As String, difficultyLevel As String
    Dim difficultyArray() As String
    
    shapeName = clickedShp.Name

    If shapeName <> "To Gameboard" Then ' "To Gameboard" is found only on the rules slide at the start of the game.
        Set sld = SlideShowWindows(1).View.Slide
        
        If debugEnabled And numberOfTeams < 4 And shapeName = "Return Button" Then
            Debug.Print vbCrLf & "Awarding points to the marmosets."
        End If
        
        If sld.SlideIndex >= firstQuestionSlide And sld.SlideIndex <= lastQuestionSlide Then
            AddToPreviouslySeen
        End If
        
        For Each shp In sld.Shapes
            If shp.Name = "Difficulty" Then
                difficultyArray = Split(sld.Shapes("Difficulty").TextFrame2.TextRange.Text, ": ")
                difficultyLevel = difficultyArray(1)
                Exit For
            End If
        Next shp

        If numberOfTeams < 4 Then
            Select Case difficultyLevel
                Case "Easy"
                    UpdateScore "Marmosets", 1
                Case "Medium"
                    UpdateScore "Marmosets", 2
                Case "Hard"
                    UpdateScore "Marmosets", 3
            End Select
        End If
    
        RemoveShapes "Hide"
        RemoveShapes "Animated", "GameboardTile"
    End If
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Moving to gameboard."
    End If

    UpdateScore "All", 0, True
    
    chosenTile = ""

    ChangeToNewSlide gameboardSlide
End Sub

Private Sub UpdateCurrentStateVariables()
    currentStateValuesArray(1) = classDay
    currentStateValuesArray(2) = classTime
    currentStateValuesArray(3) = numberOfTeams
    currentStateValuesArray(4) = team1Score
    currentStateValuesArray(5) = team2Score
    currentStateValuesArray(6) = team3Score
    currentStateValuesArray(7) = team4Score
    currentStateValuesArray(8) = marmosetsScore
    
    If debugEnabled Then
        Debug.Print "Updating variables to track the game's current state."
        Debug.Print "   classDay: " & classDay
        Debug.Print "   classTime: " & classTime
        Debug.Print "   numberOfTeams: " & numberOfTeams
        Debug.Print "   team1Score: " & team1Score
        Debug.Print "   team2Score: " & team2Score
        Debug.Print "   team3Score: " & team3Score
        Debug.Print "   team4Score: " & team4Score
        Debug.Print "   marmosetsScore: " & marmosetsScore
    End If
End Sub

Private Sub AddToPreviouslySeen()
    Dim sld As Slide
    Dim questionID As String
    Dim newArraySize As Integer
    
    Set sld = SlideShowWindows(1).View.Slide
    
    questionID = Mid(sld.Shapes("QuestionID").TextFrame2.TextRange.Text, 13)
    
    If debugEnabled Then
        Debug.Print "Adding " & questionID & " to the list of previously seen questions."
    End If
    
    If previouslySeenQuestionsArray(1) = "" Then
        previouslySeenQuestionsArray(1) = questionID
    Else
        newArraySize = UBound(previouslySeenQuestionsArray) + 1
        
        ReDim Preserve previouslySeenQuestionsArray(1 To newArraySize)
        previouslySeenQuestionsArray(newArraySize) = questionID
    End If
End Sub

Private Sub RemoveShapes(ByVal removalType As String, Optional ByVal tileType As String = "")
    Dim sld As Slide
    Dim shapeName As String, shapeToRemove As String
    Dim i As Integer

    Select Case removalType
        Case "Hide"
            Set sld = ActivePresentation.Slides(gameboardSlide)
            If sld.TimeLine.MainSequence.Count > 0 Then
                For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                    With sld.TimeLine.MainSequence.Item(i)
                        If .EffectType = msoAnimEffectRandomBars And .Shape.Name <> "" Then
                            shapeName = .Shape.Name
                            .Delete
                            sld.Shapes(shapeName).Visible = msoFalse
                        End If
                    End With
                Next i
            End If
        
            Set sld = ActivePresentation.Slides(lotterySlide)
            If sld.TimeLine.MainSequence.Count > 0 Then
                For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                    With sld.TimeLine.MainSequence.Item(i)
                        If .EffectType = msoAnimEffectRandomBars Then
                            shapeName = .Shape.Name
                            .Delete
                            sld.Shapes(shapeName).Visible = msoFalse
                        End If
                    End With
                Next i
            End If
        Case "Animated"
            If tileType = "GameboardTile" Then
                shapeToRemove = chosenTile
                Set sld = ActivePresentation.Slides(gameboardSlide)
            ElseIf tileType = "LotteryTile" Then
                shapeToRemove = chosenLottery
                Set sld = ActivePresentation.Slides(lotterySlide)
            End If

            With sld.TimeLine.MainSequence.AddEffect(Shape:=sld.Shapes(shapeToRemove), effectId:=msoAnimEffectRandomBars)
                .Timing.Duration = 0.75
                .Timing.TriggerType = msoAnimTriggerWithPrevious
                .Exit = msoTrue
            End With
    End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'               SCOREBOARD SUBS                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Scoreboard(clickedShp As Shape)
    Dim messageToDisplay As String
    Dim tempInt As Integer

    Select Case clickedShp.Name
        Case "team1_p1"
            UpdateScore "Team1", 1
        Case "team1_p2"
            UpdateScore "Team1", 2
        Case "team1_p3"
            UpdateScore "Team1", 3
        Case "team1_s1"
            UpdateScore "Team1", -1
        Case "team1_s2"
            UpdateScore "Team1", -2
        Case "team1_s3"
            UpdateScore "Team1", -3
        Case "team1_half"
            team1Score = Int(team1Score / 2)
            UpdateScore "Team1", 0

        Case "team2_p1"
            UpdateScore "Team2", 1
        Case "team2_p2"
            UpdateScore "Team2", 2
        Case "team2_p3"
            UpdateScore "Team2", 3
        Case "team2_s1"
            UpdateScore "Team2", -1
        Case "team2_s2"
            UpdateScore "Team2", -2
        Case "team2_s3"
            UpdateScore "Team2", -3
        Case "team2_half"
            team2Score = Int(team2Score / 2)
            UpdateScore "Team2", 0

        Case "team3_p1"
            UpdateScore "Team3", 1
        Case "team3_p2"
            UpdateScore "Team3", 2
        Case "team3_p3"
            UpdateScore "Team3", 3
        Case "team3_s1"
            UpdateScore "Team3", -1
        Case "team3_s2"
            UpdateScore "Team3", -2
        Case "team3_s3"
            UpdateScore "Team3", -3
        Case "team3_half"
            team3Score = Int(team3Score / 2)
            UpdateScore "Team3", 0

        Case "team4_p1"
            UpdateScore "Team4", 1
        Case "team4_p2"
            UpdateScore "Team4", 2
        Case "team4_p3"
            UpdateScore "Team4", 3
        Case "team4_s1"
            UpdateScore "Team4", -1
        Case "team4_s2"
            UpdateScore "Team4", -2
        Case "team4_s3"
            UpdateScore "Team4", -3
        Case "team4_half"
            team4Score = Int(team4Score / 2)
            UpdateScore "Team4", 0

        Case "Swap"
            If numberOfTeams = 2 Then
                tempInt = team1Score
                team1Score = team2Score
                team2Score = tempInt
                ActivePresentation.Slides(gameboardSlide).Shapes("Team1_Scoreboard").TextFrame.TextRange.Text = team1Score
                ActivePresentation.Slides(gameboardSlide).Shapes("Team2_Scoreboard").TextFrame.TextRange.Text = team2Score
            Else
                messageToDisplay = "Click on the icons of the two teams to swap points!"
                
                DisplayMessage messageToDisplay, "OkOnly"
                swapInProgress = True
            End If
    End Select
End Sub

Public Sub IconClick(clickedShp As Shape)
    Dim swapTeamName As String
    Dim iconNameArray() As String
    Dim firstTeamPoints As Integer, secondTeamPoints As Integer, tempInt As Integer
    
    iconNameArray = Split(clickedShp.Name, " ")
    swapTeamName = Trim(iconNameArray(0))

    If Not swapInProgress Then
        Exit Sub
    End If
    
    If swapFirstTeam = "" Then
        swapFirstTeam = swapTeamName
    Else
        swapSecondTeam = swapTeamName

        firstTeamPoints = CInt(ActivePresentation.Slides(gameboardSlide).Shapes(swapFirstTeam & "_Scoreboard").TextFrame.TextRange.Text)
        secondTeamPoints = CInt(ActivePresentation.Slides(gameboardSlide).Shapes(swapSecondTeam & "_Scoreboard").TextFrame.TextRange.Text)

        tempInt = firstTeamPoints
        firstTeamPoints = secondTeamPoints
        secondTeamPoints = tempInt

        Select Case swapFirstTeam
            Case "Team1"
                team1Score = firstTeamPoints
            Case "Team2"
                team2Score = firstTeamPoints
            Case "Team3"
                team3Score = firstTeamPoints
            Case "Team4"
                team4Score = firstTeamPoints
        End Select

        Select Case swapSecondTeam
            Case "Team1"
                team1Score = secondTeamPoints
            Case "Team2"
                team2Score = secondTeamPoints
            Case "Team3"
                team3Score = secondTeamPoints
            Case "Team4"
                team4Score = secondTeamPoints
        End Select

        UpdateScore "All", 0

        swapInProgress = False
        swapFirstTeam = ""
        swapSecondTeam = ""
    End If
End Sub

Private Sub UpdateScore(ByVal team As String, ByVal pointsToAward As Integer, Optional ByVal startNewGame As Boolean = False)
    Select Case team
        Case "Team1"
            team1Score = team1Score + pointsToAward
        Case "Team2"
            team2Score = team2Score + pointsToAward
        Case "Team3"
            team3Score = team3Score + pointsToAward
        Case "Team4"
            team4Score = team4Score + pointsToAward
        Case "Marmosets"
            marmosetsScore = marmosetsScore + pointsToAward
        Case "All"
            team1Score = team1Score + pointsToAward
            team2Score = team2Score + pointsToAward
            If numberOfTeams > 2 Then team3Score = team3Score + pointsToAward
            If numberOfTeams > 3 Then team4Score = team4Score + pointsToAward
            If numberOfTeams < 4 Then marmosetsScore = marmosetsScore + pointsToAward
    End Select

    If Not allowNegatives Then
        If team1Score < 0 Then team1Score = 0
        If team2Score < 0 Then team2Score = 0
        If numberOfTeams > 2 And team3Score < 0 Then team3Score = 0
        If numberOfTeams > 3 And team4Score < 0 Then team4Score = 0
        If numberOfTeams < 4 And marmosetsScore < 0 Then marmosetsScore = 0
    End If

    ActivePresentation.Slides(gameboardSlide).Shapes("Team1_Scoreboard").TextFrame.TextRange.Text = team1Score
    ActivePresentation.Slides(gameboardSlide).Shapes("Team2_Scoreboard").TextFrame.TextRange.Text = team2Score
    If numberOfTeams > 2 Then ActivePresentation.Slides(gameboardSlide).Shapes("Team3_Scoreboard").TextFrame.TextRange.Text = team3Score
    If numberOfTeams > 3 Then ActivePresentation.Slides(gameboardSlide).Shapes("Team4_Scoreboard").TextFrame.TextRange.Text = team4Score
    If numberOfTeams < 4 Then ActivePresentation.Slides(gameboardSlide).Shapes("Marmoset_Scoreboard").TextFrame.TextRange.Text = marmosetsScore
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'                    Admin                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub FindFirstAndLastTriviaSlide(ByRef firstQuestion_TwoTeams As Integer, ByRef lastQuestion_FourTeams As Integer)
    Dim sld As Slide
    Dim i As Integer
    Dim sectionStart(), sectionEnd() As Integer
    
    ReDim sectionStart(1 To ActivePresentation.SectionProperties.Count)
    ReDim sectionEnd(1 To ActivePresentation.SectionProperties.Count)

    For Each sld In ActivePresentation.Slides
        If sectionStart(sld.sectionIndex) = 0 Or sld.SlideIndex < sectionStart(sld.sectionIndex) Then
            sectionStart(sld.sectionIndex) = sld.SlideIndex
        End If
        
        If sld.SlideIndex > sectionEnd(sld.sectionIndex) Then
            sectionEnd(sld.sectionIndex) = sld.SlideIndex
        End If
    Next sld
    
    For i = 1 To ActivePresentation.SectionProperties.Count
        If ActivePresentation.SectionProperties.Name(i) = "TwoTeams" Then
            firstQuestion_TwoTeams = sectionStart(i)
        ElseIf ActivePresentation.SectionProperties.Name(i) = "FourTeams" Then
            lastQuestion_FourTeams = sectionEnd(i)
        End If
    Next i
End Sub

Private Sub DisplayMessage(ByVal messageString As String, ByVal messageBoxType As String, Optional ByRef returnedValue As Integer = 0)
    #If Mac Then
        Select Case messageBoxType
            Case "OkOnly"
                returnedValue = AppleScriptTask("AngryBirds.scpt", "OKDialog", messageString)
            Case "OkCancel"
                returnedValue = AppleScriptTask("AngryBirds.scpt", "OkCancelDialog", messageString)
            Case "YesNo"
                returnedValue = AppleScriptTask("AngryBirds.scpt", "YesNoDialog", messageString)
            Case "AppleScriptNotFound"
                returnedValue = MsgBox(messageString, vbYesNo + vbApplicationModal + vbExclamation + vbDefaultButton1, "Warning!")
        End Select
    #Else
        Select Case messageBoxType
            Case "OkOnly"
                MsgBox messageString
            Case "OkCancel"
                returnedValue = MsgBox(messageString, vbOKCancel + vbApplicationModal + vbExclamation + vbDefaultButton1, "Warning!")
            Case "YesNo"
                returnedValue = MsgBox(messageString, vbYesNo + vbApplicationModal + vbExclamation + vbDefaultButton1, "Warning!")
        End Select
    #End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'             DOWNLOAD & VALIDATION
'''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub GetUpdates(clickedShp As Shape)
    Dim sld As Slide
    Dim clickedButton As String, fileType As String, messageToDisplay As String, pathToOpen As String
    Dim useCurl As Boolean, downloadResult As Boolean, excelIsInstalled As Boolean
    Dim userChoice As Integer
    
    If clickedShp.Name = "Versions_Game_Update" Then
        fileType = "Game"
    ElseIf clickedShp.Name = "Versions_Questions_Update" Then
        fileType = "Questions"
    End If
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Beginning update procedure."
        Debug.Print "Update to download: " & fileType & vbCrLf
    End If
    
    #If Win64 Then
        useCurl = CheckForCurl()
        
        If Not useCurl Then
            If Not CheckForDotNet35() Then
                messageToDisplay = "Microsoft .NET Framework 3.5 must be installed in order to update the files. Would you like to download it now?"
                DisplayMessage messageToDisplay, "YesNo", userChoice
            
                If userChoice = vbYes Then
                    ActivePresentation.FollowHyperlink Address:="https://www.microsoft.com/en-ca/download/details.aspx?id=21", NewWindow:=True, AddHistory:=True
                    ActivePresentation.SlideShowWindow.View.Exit
                End If
                Exit Sub
            End If
        End If
    #End If
    
    If fileType = "Game" Then
        messageToDisplay = "Would you like to download the latest version of the PPT?"
        DisplayMessage messageToDisplay, "YesNo", userChoice
    ElseIf fileType = "Questions" Then
        If ActivePresentation.Slides(configSld).Shapes("Versions_Game_Update").Visible Then
            messageToDisplay = "Please update to the most recent version of the PPT before updating the questions."
            DisplayMessage messageToDisplay, "OkOnly"
            Exit Sub
        End If
        
        excelIsInstalled = CheckForExcel()
        
        If Not excelIsInstalled Then
            messageToDisplay = "Sorry, but Excel needs to be installed in order to update the questions. Either install Excel or try downloading the latest version of the PPT."
            DisplayMessage messageToDisplay, "OkOnly"
            Exit Sub
        End If
        
        messageToDisplay = "Would you like to update to the latest set of questions? This process may take a few minutes to complete."
        DisplayMessage messageToDisplay, "YesNo", userChoice
    
        If userChoice = vbNo Then
            Exit Sub
        End If
        
        #If Mac Then
            messageToDisplay = "You may see a window appear requesting permission to access 'Questions.xlsx'. Please grant access for the update to progress and wait for it to complete. If Excel opens, it will close automatically once the update is complete."
        #Else
            messageToDisplay = "Please do not click anything until the process has finished. You will see a message appear once the update is complete."
        #End If
        DisplayMessage messageToDisplay, "OkOnly"
    End If
        
    If userChoice = vbNo Then
        Exit Sub
    End If
    
    Set sld = ActivePresentation.Slides(configSld)
    
    sld.Shapes("Updating Indicator").Visible = msoTrue
    sld.Shapes("Updating Message").Visible = msoTrue
    sld.Shapes("Updating Background").Visible = msoTrue
    
    DownloadNewFile useCurl, fileType, downloadResult, pathToOpen
    
    If Not downloadResult Then
        messageToDisplay = "Failed to download the file."
        DisplayMessage messageToDisplay, "OkOnly"
        
        sld.Shapes("Updating Indicator").Visible = msoFalse
        sld.Shapes("Updating Message").Visible = msoFalse
        sld.Shapes("Updating Background").Visible = msoFalse
        Exit Sub
    End If
    
    If debugEnabled Then
        Debug.Print "Download succussful." & vbCrLf & "Continuing update process."
    End If
    
    If fileType = "Game" Then
        messageToDisplay = "Download successful. Please close this file and open the latest version."
        DisplayMessage messageToDisplay, "OkOnly"
        
        #If Mac Then
            userChoice = AppleScriptTask("AngryBirds.scpt", "OpenContainingFolder", pathToOpen)
        #End If
        
        sld.Shapes("Updating Indicator").Visible = msoFalse
        sld.Shapes("Updating Message").Visible = msoFalse
        sld.Shapes("Updating Background").Visible = msoFalse
        
        ActivePresentation.SlideShowWindow.View.Exit
    ElseIf fileType = "Questions" Then
        RemoveOldQuestionSlides
        ImportFromExcel
        
        sld.Shapes("Versions_Questions_Version").TextFrame2.TextRange.Text = GetOnlineVersionNumber(fileType)
        sld.Shapes("Versions_Questions_Update").Visible = msoFalse
        sld.Shapes("Updating Indicator").Visible = msoFalse
        sld.Shapes("Updating Message").Visible = msoFalse
        sld.Shapes("Updating Background").Visible = msoFalse
        
        ActivePresentation.Save
        
        messageToDisplay = "Questions have been successfully updated and the PPT will now close. Please restart the PPT to continue."
        DisplayMessage messageToDisplay, "OkOnly"
        
        ActivePresentation.SlideShowWindow.View.Exit
    End If
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Update complete. Terminating program."
    End If
End Sub

Private Sub DownloadNewFile(ByVal useCurl As Boolean, ByVal fileTypeToDownload As String, ByRef downloadResult As Boolean, Optional ByRef pathToOpen As String = "")
    Dim fileURL As String, validHash As String, savePath As String, finalPath As String
    Dim versionNumber As Long

    If fileTypeToDownload = "Game" Then
        fileURL = "https://dl.dropboxusercontent.com/scl/fi/obh98ve69z9uyvhv2kyub/AngryBirdsTrivia.pptm?rlkey=og6tx6jacnyzgawzdyyz3zglr&dl=1"
    ElseIf fileTypeToDownload = "Questions" Then
        fileURL = "https://raw.githubusercontent.com/papercutter0324/AngryBirdsTrivia-AdditionalFiles/main/Questions.xlsx"
        'fileURL = "https://www.dropbox.com/scl/fi/jj7kw2b8p67ldzfc7y3k6/Questions.xlsx?rlkey=2tp9pppf1v8aqjcr45y42u4f8&st=jeyefro6&dl=1"
    End If
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Beginning download of " & fileTypeToDownload
    End If
    
    versionNumber = GetOnlineVersionNumber(fileTypeToDownload)
    
    #If Mac Then
        If fileTypeToDownload = "Game" Then
            savePath = tempFolder & "/AngryBirdsTrivia.pptm"
            finalPath = AppleScriptTask("AngryBirds.scpt", "ChooseSaveLocation", versionNumber)
        ElseIf fileTypeToDownload = "Questions" Then
            savePath = tempFolder & "/Questions.xlsx"
        End If
        
        downloadResult = AppleScriptTask("AngryBirds.scpt", "DownloadFile", fileURL & "," & savePath)
        
        If downloadResult Then
            GetCurrentHashes useCurl, fileTypeToDownload, validHash
            downloadResult = GetFileHashMD5(savePath, validHash)
        End If
        
        If fileTypeToDownload = "Game" Then
            downloadResult = AppleScriptTask("AngryBirds.scpt", "MoveFile", savePath & "," & finalPath)
            
            If downloadResult Then
                pathToOpen = finalPath
            End If
        End If
    #Else
        If fileTypeToDownload = "Game" Then
            savePath = tempFolder & "\AngryBirdsTrivia.pptm"
            finalPath = ChooseSaveLocation(versionNumber)
            
            If finalPath = "" Then
                Exit Sub
            End If
        ElseIf fileTypeToDownload = "Questions" Then
            savePath = tempFolder & "\Questions.xlsx"
        End If
        
        downloadResult = DownloadFile(useCurl, fileURL, savePath)
    
        If VerifyFileOrFolderExists(savePath) Then
            GetCurrentHashes useCurl, fileTypeToDownload, validHash
            downloadResult = GetFileHashMD5(savePath, validHash)
        Else
            downloadResult = False
        End If
        
        If fileTypeToDownload = "Game" Then
            Name savePath As finalPath
            downloadResult = VerifyFileOrFolderExists(finalPath)
        End If
    #End If
End Sub

Function DownloadFile(ByVal useCurl As Boolean, ByVal fileURL As String, ByVal savePath As String) As Boolean
    #If Mac Then
        DownloadFile = AppleScriptTask("AngryBirds.scpt", "DownloadFile", fileURL & "," & savePath)
    #Else
        Dim objShell As Object, xmlHTTP As Object, fileStream As Object
        Dim downloadCommand As String
        Dim downloadResult As Boolean
        
        If useCurl Then
            Set objShell = CreateObject("WScript.Shell")
        
            downloadCommand = "cmd /c curl.exe -o """ & savePath & """ """ & fileURL & """"

            downloadResult = (objShell.Run(downloadCommand, 0, True))
        
            Set objShell = Nothing
        Else
            On Error GoTo ErrorHandler
            
            Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            Set fileStream = CreateObject("ADODB.Stream")
            
            xmlHTTP.Open "GET", fileURL, False
            xmlHTTP.send
            
            If xmlHTTP.Status = 200 Then
                fileStream.Open
                fileStream.Type = 1 ' Binary
                fileStream.Write xmlHTTP.responseBody
                fileStream.SaveToFile savePath, 2 ' Overwrite existing
                fileStream.Close
                downloadResult = True
            Else
                If debugEnabled Then
                    Debug.Print "xmlHTTP Error: " & xmlHTTP.Status & " - " & xmlHTTP.statusText
                End If
                
                downloadResult = False
            End If
            
            Set xmlHTTP = Nothing
            Set fileStream = Nothing
            
            DownloadFile = downloadResult
        End If
        
ErrorHandler:
        If Err.Number = 0 Then
            On Error GoTo 0
        Else
            If debugEnabled Then
                Debug.Print "Error Handler: " & Err.Number & " - " & Err.Description
            End If
            downloadResult = False
        End If
    #End If
End Function

Private Sub GetCurrentHashes(ByVal useCurl As Boolean, ByVal fileTypeToDownload As String, ByRef validHash As String)
    Dim fileURL As String, savePath As String, lineToCompare As String
    Dim downloadResult As Boolean
    
    fileURL = "https://raw.githubusercontent.com/papercutter0324/AngryBirdsTrivia-AdditionalFiles/main/FileHashes.txt"
    savePath = tempFolder & pathSeparator & "FileHashes.txt"
    validHash = ""
        
    downloadResult = DownloadFile(useCurl, fileURL, savePath)
        
    If Not VerifyFileOrFolderExists(savePath) Then
        Exit Sub
    End If
        
    Select Case fileTypeToDownload
        Case "Game"
            lineToCompare = "AngryBirds.pptm"
        Case "Questions"
            lineToCompare = "Questions.xlsx"
        Case "AngryBirds.scpt"
            lineToCompare = "AngryBirds.scpt"
        Case Else
            Exit Sub
    End Select
    
    If debugEnabled Then
        Debug.Print vbCrLf & "Checking hashes to verify integrity of downloaded file."
    End If
    
    #If Mac Then
        validHash = AppleScriptTask("AngryBirds.scpt", "GetLatestHashes", savePath & "," & lineToCompare)
    #Else
        Dim hashLine As String, fileName As String, hashValue As String
        Dim hashArray() As String
        Dim stream As Object
        
        Set stream = CreateObject("ADODB.Stream")
        
        stream.Open
        stream.Type = 2 ' Text
        stream.Charset = "UTF-8"
        stream.LoadFromFile savePath
        
        Do While Not stream.EOS
            hashLine = stream.ReadText(-2) ' Read line with UTF-8 encoding
        
            hashArray = Split(hashLine, " = ")
            fileName = Trim(hashArray(0))
            hashValue = Trim(hashArray(1))
            
            If lineToCompare = fileName Then
                validHash = LCase(hashValue)
                Exit Do
            End If
        Loop
    
        stream.Close
    #End If
    
    If debugEnabled Then
        Debug.Print "   Expected hash = " & validHash
    End If
End Sub

Function GetFileHashMD5(ByVal filePath As String, ByVal validHash As String) As Boolean
    #If Mac Then
        GetFileHashMD5 = AppleScriptTask("AngryBirds.scpt", "CompareMD5Hashes", filePath & "," & validHash)
    #Else
        Dim tempFile As String, hashFile As String, hashText As String, hexHash As String
        Dim oShell As Object, fileSystem As Object
    
        tempFile = Environ("TEMP") & "\temphash.txt"
        
        Set oShell = CreateObject("WScript.Shell")
        
        oShell.Run "cmd /c certutil -hashfile """ & filePath & """ MD5 > """ & tempFile & """", 0, True
        
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        
        ' Open the temporary file and read the hash
        If fileSystem.fileExists(tempFile) Then
            With fileSystem.OpenTextFile(tempFile, 1)
                .ReadLine ' Skip the first line
                hexHash = .ReadLine ' Read the hash
                .Close
            End With
            ' Remove the temporary file
            fileSystem.DeleteFile tempFile
        Else
            MsgBox "Error computing MD5 hash"
            GetFileHashMD5 = False
            Exit Function
        End If
        
        Set oShell = Nothing
        Set fileSystem = Nothing
    
        ' Remove any spaces from the hash (CertUtil adds spaces every 2 characters)
        hexHash = Replace(LCase(hexHash), " ", "")
        
        If debugEnabled Then
            Debug.Print "   Returned hash = " & hexHash
        End If
        
        GetFileHashMD5 = (hexHash = validHash)
    #End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''
'         SLIDE UPDATE & FORMATTING ROUTINES
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub FindPosition()
    Dim leftPos As Single, topPos As Single
    Dim myShape As Shape
    
    Set myShape = ActiveWindow.Selection.ShapeRange(1)
    
    leftPos = myShape.Left - ActivePresentation.SlideMaster.Width
    topPos = myShape.Top
    
    If debugEnabled Then
        Debug.Print "Shape position: (" & leftPos & "pt, " & topPos & "pt)"
    End If
End Sub

Private Sub SetProperTriviaSlideFormatting()
    Dim shpLeft As Integer, shpWidth As Integer, shpHeight As Integer, firstQuestion_TwoTeams As Integer, lastQuestion_FourTeams As Integer
    Dim sld As Slide
    
    shpLeft = ActivePresentation.SlideMaster.Width + 20
    shpWidth = GameCode.InchesToPoints(5)
    shpHeight = GameCode.InchesToPoints(0.57)

    FindFirstAndLastTriviaSlide firstQuestion_TwoTeams, lastQuestion_FourTeams
    
    If lastQuestion_FourTeams = 0 Or firstQuestion_TwoTeams = 0 Then
        Exit Sub
    End If
    
    For Each sld In ActivePresentation.Slides
        If sld.SlideNumber >= firstQuestion_TwoTeams And sld.SlideNumber <= lastQuestion_FourTeams Then
            CheckForNotesSection sld, firstQuestion_TwoTeams, lastQuestion_FourTeams
            CheckOffScreenTextboxes sld, shpLeft, shpWidth, shpHeight
        End If
    Next sld

    SlideShowWindows(1).View.Exit
End Sub

Private Function InchesToPoints(ByVal inchesValue As Long) As Integer
    InchesToPoints = inchesValue * 72
End Function

Private Sub CheckForNotesSection(ByVal sld As Slide, ByVal firstQuestion_TwoTeams As Integer, ByVal lastQuestion_FourTeams As Integer)
    Dim notesSectionExists As Boolean
    Dim noteText As String
    
    notesSectionExists = False
    
    For Each shp In sld.NotesPage.Shapes
        With shp
            If .Type = msoPlaceholder And .PlaceholderFormat.Type = ppPlaceholderBody Then
                .Name = "Notes Placeholder 2"
                notesSectionExists = True
                Exit For
            End If
        End With
    Next shp

    If Not notesSectionExists Then
        Set shp = sld.NotesPage.Shapes.AddPlaceholder(ppPlaceholderBody, 0, 0, 432, 324)
        shp.Name = "Notes Placeholder 2"
    End If

    ' Copy the answer into the notes so that the teacher can see it in Presenter View
    noteText = "Answer: " & sld.Shapes("Answer Box").TextFrame2.TextRange.Text
    sld.NotesPage.Shapes("Notes Placeholder 2").TextFrame2.TextRange.Text = noteText
End Sub

Private Sub CheckOffScreenTextboxes(ByVal sld As Slide, ByVal shpLeft As Integer, ByVal shpWidth As Integer, ByVal shpHeight As Integer)
    Dim foundCategory As Boolean, foundLevel As Boolean, foundBook As Boolean, foundUnit As Boolean, foundDifficulty As Boolean, foundQuestionID As Boolean
    Dim currentText As String
    Dim shapesArray As Variant
    
    shapesArray = Array("Category", "Level", "Book", "Unit", "Difficulty", "QuestionID")
    
    For i = LBound(shapesArray) To UBound(shapesArray)
        For Each shp In sld.Shapes
            With shp
                If .Name = shapesArray(i) Then
                    .Locked = msoFalse
                    .Left = shpLeft
    
                    currentText = .TextFrame2.TextRange.Text
    
                    Select Case shapesArray(i)
                        Case "Category"
                            .Top = 0
                            foundCategory = True
                        Case "Level"
                            .Top = 60
                            foundLevel = True
                        Case "Book"
                            .Top = 120
                            foundBook = True
                        Case "Unit"
                            .Top = 180
                            foundUnit = True
                        Case "Difficulty"
                            .Top = 240
                            foundDifficulty = True
                        Case "QuestionID"
                            .Top = 300
                            foundQuestionID = True
                    End Select
    
                    .Width = shpWidth
                    .Height = shpHeight
                    .Fill.ForeColor.RGB = RGB(242, 242, 242)
                    .Line.ForeColor.RGB = RGB(242, 242, 242)
                    .TextFrame2.TextRange.Text = currentText
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignLeft
                    .TextFrame2.TextRange.Font.Size = 28
                    .TextFrame2.TextRange.Font.Name = "Calibri"
                    .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                    .TextFrame2.TextRange.Font.Bold = msoTrue
                    .Locked = msoTrue
                    
                    Exit For
                End If
            End With
        Next shp
    Next i
    
    If Not foundCategory Then
        CreateMissingShape sld, "Category", shpLeft, 0, shpWidth, shpHeight, "Category: "
    End If

    If Not foundLevel Then
        CreateMissingShape sld, "Level", shpLeft, 60, shpWidth, shpHeight, "Level: "
    End If

    If Not foundBook Then
        CreateMissingShape sld, "Book", shpLeft, 120, shpWidth, shpHeight, "Book: "
    End If

    If Not foundUnit Then
        CreateMissingShape sld, "Unit", shpLeft, 180, shpWidth, shpHeight, "Unit: "
    End If

    If Not foundDifficulty Then
        CreateMissingShape sld, "Difficulty", shpLeft, 240, shpWidth, shpHeight, "Difficulty: "
    End If

    If Not foundQuestionID Then
        CreateMissingShape sld, "QuestionID", shpLeft, 300, shpWidth, shpHeight, "QuestionID: "
    End If
End Sub

Private Sub CreateMissingShape(ByVal sld As Slide, ByVal shapeName As String, ByVal shpLeft As Integer, ByVal shpTop As Integer, ByVal shpWidth As Integer, ByVal shpHeight As Integer, ByVal textValue As String)
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, shpLeft, shpTop, shpWidth, shpHeight)

    With shp
        .Name = shapeName
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Line.ForeColor.RGB = RGB(242, 242, 242)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignLeft
        .TextFrame2.TextRange.Font.Size = 28
        .TextFrame2.TextRange.Font.Name = "Calibri"
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Text = textValue
        .Locked = msoTrue
    End With
End Sub

Private Sub RemoveOldQuestionSlides()
    Dim firstQuestion_TwoTeams As Integer, lastQuestion_FourTeams As Integer
    FindFirstAndLastTriviaSlide firstQuestion_TwoTeams, lastQuestion_FourTeams
    
    If lastQuestion_FourTeams = 0 Or firstQuestion_TwoTeams = 0 Then
        Exit Sub
    End If
    
    If debugEnabled Then
        Debug.Print "   Deleting all questions slides from slideIndex " & firstQuestion_TwoTeams & " to " & lastQuestion_FourTeams & "."
    End If
    
    Dim i As Integer, sld As Slide
    For i = lastQuestion_FourTeams To firstQuestion_TwoTeams Step -1
        ActivePresentation.Slides(i).Delete
    Next i
    
    If debugEnabled Then
        Debug.Print "   Deletion complete."
    End If
End Sub

Private Sub ImportFromExcel()
    Dim sld As Slide, shp As Shape
    Dim objSlide As Object, xlApp As Object, xlWorkbook As Object, xlWorksheet As Object
    Dim questionCategory As String, questionLevel As String, questionBook As String, questionUnit As String
    Dim questionDifficulty As String, questionQuestionID As String, questionQuestion As String, questionAnswer As String
    Dim messageToDisplay As String, filePath As String
    Dim i As Integer, j As Integer, rowIndex As Integer, itemCounter As Integer
    Dim templateSlidesArray(1 To 3) As Integer, questionSectionsArray(1 To 3) As Integer
    
    If debugEnabled Then
        Debug.Print "   Beginning importation of new questions."
    End If
    
    filePath = tempFolder & pathSeparator & "Questions.xlsx"
    
    If Not VerifyFileOrFolderExists(filePath) Then
        If debugEnabled Then
            Debug.Print "Questions.xlsx could not be found." & vbCrLf & "Terminating update."
        End If
        
        messageToDisplay = "Sorry, there was an error locating Questions.xlsx. Please try again."
        DisplayMessage messageToDisplay, "OkOnly"
        Exit Sub
    End If

    For i = 1 To ActivePresentation.SectionProperties.Count
        With ActivePresentation.SectionProperties
            Select Case .Name(i)
                Case "Templates"
                    templateSlidesArray(1) = .FirstSlide(i)
                    templateSlidesArray(2) = .FirstSlide(i) + 1
                    templateSlidesArray(3) = .FirstSlide(i) + 2
                    itemCounter = itemCounter + 1
                Case "TwoTeams"
                    questionSectionsArray(1) = i
                    itemCounter = itemCounter + 1
                Case "ThreeTeams"
                    questionSectionsArray(2) = i
                    itemCounter = itemCounter + 1
                Case "FourTeams"
                    questionSectionsArray(3) = i
                    itemCounter = itemCounter + 1
            End Select
        
            If itemCounter = 4 Then
                Exit For
            End If
        End With
    Next i

    ' Create a new instance of Excel
    #If Mac Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Application.Visible = False
    #Else
        Set xlApp = CreateObject("Excel.Application")
    #End If

    ' Open the workbook
    #If Mac Then
        Set xlWorkbook = xlApp.Application.Workbooks.Open(filePath)
    #Else
        Set xlWorkbook = xlApp.Workbooks.Open(filePath)
    #End If

    ' Loop through the worksheets
    For Each xlWorksheet In xlWorkbook.Worksheets
        ' Get the category name
        questionCategory = xlWorksheet.Name
        
        ' Loop through the rows
        With xlWorksheet
            For rowIndex = .Cells(xlWorksheet.Rows.Count, 1).End(-4162).Row To 2 Step -1
            ' Get the values from the worksheet
                questionLevel = .Cells(rowIndex, 1).value
                questionBook = .Cells(rowIndex, 2).value
                questionUnit = .Cells(rowIndex, 3).value
                questionDifficulty = .Cells(rowIndex, 4).value
                questionQuestionID = .Name & Format(rowIndex - 1, "0000")
                questionQuestion = .Cells(rowIndex, 5).value
                questionAnswer = .Cells(rowIndex, 6).value

                ' Loop through the template slides and create new slides based on the category and difficulty
                For i = 1 To 3
                    Set objSlide = ActivePresentation.Slides(templateSlidesArray(i)).Duplicate
    
                    ' Set basic formatting
                    With objSlide.Shapes
                        SetShapeProperties .Item("Category"), "Category: ", questionCategory
                        SetShapeProperties .Item("Level"), "Level: ", questionLevel
                        SetShapeProperties .Item("Book"), "Book: ", questionBook
                        SetShapeProperties .Item("Unit"), "Unit: ", questionUnit
                        SetShapeProperties .Item("Difficulty"), "Difficulty: ", questionDifficulty
                        SetShapeProperties .Item("QuestionID"), "QuestionID: ", questionQuestionID
                        SetShapeProperties .Item("Question Box"), questionQuestion, , 32
                        SetShapeProperties .Item("Answer Box"), questionAnswer, , 32
        
                        'Set specific overrides
                        With .Item("Question Box").TextFrame2
                            .AutoSize = msoAutoSizeTextToFitShape
                            .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                        End With
                        With .Item("Answer Box").TextFrame2
                            .AutoSize = msoAutoSizeTextToFitShape
                            .TextRange.ParagraphFormat.Alignment = ppAlignCenter
                        End With
                    End With
        
                    ' Copy the answer into the notes so that the teacher can see it in Presenter View
                    Set shp = objSlide.NotesPage.Shapes("Notes Placeholder 2")
                    shp.TextFrame2.TextRange.Text = "Answer: " & questionAnswer
    
                    objSlide.SlideShowTransition.Hidden = msoFalse
                    objSlide.MoveToSectionStart questionSectionsArray(i)
                    
                    Set objSlide = Nothing
                Next i
            Next rowIndex
        End With
    Next xlWorksheet

    ' Close the Excel workbook
    xlWorkbook.Close False
    #If Mac Then
        xlApp.Application.Quit
    #Else
        xlApp.Quit
    #End If

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    If debugEnabled Then
        Debug.Print "Importation process complete."
    End If
End Sub

Private Sub SetShapeProperties(ByVal shp As Shape, ByVal defaultText As String, Optional ByVal questionDetails As String = "", Optional ByVal fontSize As Integer = 28)
    Dim textToWrite As String
    If defaultText = "Difficulty: " Then
        'Add a sub to convert from 5 levels to three
        textToWrite = defaultText & questionDetails
    Else
        textToWrite = defaultText & questionDetails
    End If
    
    With shp
        .TextFrame2.TextRange.Text = textToWrite
        .TextFrame2.TextRange.Font.Name = "Calibri"
        .TextFrame2.TextRange.Font.Size = fontSize
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    End With
End Sub

#If Mac Then
#Else
Private Sub FormatWorksheets()
    Dim xlWorkbook As Workbook: Set xlWorkbook = ThisWorkbook
    Dim xlWorksheet As Worksheet

    For Each xlWorksheet In xlWorkbook.Worksheets
         If xlWorksheet.Name <> "Instructions" Then
            With xlWorksheet
                ' Format row 1
                With .Rows(1)
                    .Font.Size = 14
                    .Font.Bold = True
                    .Font.Underline = xlUnderlineStyleSingle
                    .VerticalAlignment = xlBottom
                End With

               ' Format other rows
               With .Rows("2:" & .Rows.Count)
                   .Font.Size = 11
                   .Font.Bold = False
                   .Font.Underline = xlUnderlineStyleNone
                    .VerticalAlignment = xlCenter
                End With

                ' Format columns A, B, C, D
                .Range("A:D").HorizontalAlignment = xlCenter

                ' Format columns E, F
                .Range("E:F").HorizontalAlignment = xlLeft

                ' Apply text format to all rows and columns
                .Cells.NumberFormat = "@"
            End With
        End If
    Next xlWorksheet
End Sub
#End If

#If Mac Then
#Else
Private Sub ExportTriviaToSpreadsheet()
    Dim sld As Slide, shp As Shape
    Dim int_sldIndex As Integer, int_FirstQuestion As Integer, int_LastQuestion As Integer
    Dim str_Question As String, str_Answer As String, str_Category As String, str_Level As String, str_Book As String, str_Unit As String, str_Difficulty As String
    Dim arr_Words() As String, arr_TextboxLines() As String, str_CategoryName As String, str_LevelName As String, str_BookName As String, str_UnitName As String, str_DifficultyName As String

    ' Get slideIndex numbers for int_FirstQuestion and int_LastQuestion
    Dim secIndex As Integer, secStart() As Integer, secEnd() As Integer
    ReDim secStart(1 To ActivePresentation.SectionProperties.Count)
    ReDim secEnd(1 To ActivePresentation.SectionProperties.Count)

    For Each sld In ActivePresentation.Slides
        secIndex = sld.sectionIndex
        If secStart(secIndex) = 0 Or sld.SlideIndex < secStart(secIndex) Then
            secStart(secIndex) = sld.SlideIndex
        End If
        If sld.SlideIndex > secEnd(secIndex) Then
            secEnd(secIndex) = sld.SlideIndex
        End If
    Next sld

    Dim i As Integer
    For i = 1 To ActivePresentation.SectionProperties.Count
        If ActivePresentation.SectionProperties.Name(i) = "TwoTeams" Then
            int_FirstQuestion = secStart(i)
            int_LastQuestion = secEnd(i)
            Exit For
        End If
    Next i

    ' Create a new instance of Excel
    Dim xlApp As Object: Set xlApp = CreateObject("Excel.Application")

    ' Create a new workbook
    Dim xlWorkbook As Object: Set xlWorkbook = xlApp.Workbooks.Add

    ' Loop through the slides
    For int_sldIndex = int_FirstQuestion To int_LastQuestion
        ' Get and isolate the values from the textboxes on the slide
        str_Question = ActivePresentation.Slides(int_sldIndex).Shapes("Question Box").TextFrame2.TextRange.Text
        str_Answer = ActivePresentation.Slides(int_sldIndex).Shapes("Answer Box").TextFrame2.TextRange.Text

        str_Category = ActivePresentation.Slides(int_sldIndex).Shapes("Category").TextFrame2.TextRange.Text
        arr_Words = Split(str_Category, ": ")
        str_Category = arr_Words(1)

        str_Level = ActivePresentation.Slides(int_sldIndex).Shapes("Level").TextFrame2.TextRange.Text
        arr_Words = Split(str_Level, ": ")
        str_Level = arr_Words(1)

        str_Book = ActivePresentation.Slides(int_sldIndex).Shapes("Book").TextFrame2.TextRange.Text
        arr_Words = Split(str_Book, ": ")
        str_Book = arr_Words(1)

        str_Unit = ActivePresentation.Slides(int_sldIndex).Shapes("Unit").TextFrame2.TextRange.Text
        arr_Words = Split(str_Unit, ": ")
        str_Unit = arr_Words(1)

        str_Difficulty = ActivePresentation.Slides(int_sldIndex).Shapes("Difficulty").TextFrame2.TextRange.Text
        arr_Words = Split(str_Difficulty, ": ")
        str_Difficulty = arr_Words(1)

        str_QuestionID = ActivePresentation.Slides(int_sldIndex).Shapes("QuestionID").TextFrame2.TextRange.Text
        arr_Words = Split(str_QuestionID, ": ")
        str_QuestionID = arr_Words(1)

        ' Ensure multi-line questions and answers are saved as multi-line
        arr_TextboxLines = Split(str_Question, vbNewLine)
        str_Question = arr_TextboxLines(0)
        If UBound(arr_TextboxLines) = 1 Then
            str_Question = str_Question & Chr(10) & arr_TextboxLines(1)
        ElseIf UBound(arr_TextboxLines) = 2 Then
            str_Question = str_Question & Chr(10) & arr_TextboxLines(1) & Chr(10) & arr_TextboxLines(2)
        ElseIf UBound(arr_TextboxLines) = 3 Then
            str_Question = str_Question & Chr(10) & arr_TextboxLines(1) & Chr(10) & arr_TextboxLines(2) & Chr(10) & arr_TextboxLines(3)
        End If

        arr_TextboxLines = Split(str_Answer, vbNewLine)
        str_Answer = arr_TextboxLines(0)
        If UBound(arr_TextboxLines) = 1 Then
            str_Answer = str_Answer & Chr(10) & arr_TextboxLines(1)
        ElseIf UBound(arr_TextboxLines) = 2 Then
            str_Answer = str_Answer & Chr(10) & arr_TextboxLines(1) & Chr(10) & arr_TextboxLines(2)
        ElseIf UBound(arr_TextboxLines) = 3 Then
            str_Answer = str_Answer & Chr(10) & arr_TextboxLines(1) & Chr(10) & arr_TextboxLines(2) & Chr(10) & arr_TextboxLines(3)
        End If

        ' Create a new worksheet if necessary
        Dim xlWorksheet As Object
        If Not WorksheetExists(xlWorkbook, str_Category) Then
            Set xlWorksheet = xlWorkbook.Worksheets.Add
            xlWorksheet.Name = str_Category
            xlWorksheet.Cells(1, 1).value = "Level"
            xlWorksheet.Cells(1, 2).value = "Book"
            xlWorksheet.Cells(1, 3).value = "Unit"
            xlWorksheet.Cells(1, 4).value = "Difficulty"
            xlWorksheet.Cells(1, 5).value = "Question"
            xlWorksheet.Cells(1, 6).value = "Answer"
        Else
            Set xlWorksheet = xlWorkbook.Worksheets(str_Category)
        End If

        ' Find the next available row on the worksheet and write the values
        Dim int_NextRow As Integer: int_NextRow = xlWorksheet.Cells(xlWorksheet.Rows.Count, 1).End(Application.xlUp).Row + 1
        xlWorksheet.Cells(int_NextRow, 1).value = str_Level
        xlWorksheet.Cells(int_NextRow, 2).value = str_Book
        xlWorksheet.Cells(int_NextRow, 3).value = str_Unit
        xlWorksheet.Cells(int_NextRow, 4).value = str_Difficulty
        xlWorksheet.Cells(int_NextRow, 5).value = str_Question
        xlWorksheet.Cells(int_NextRow, 6).value = str_Answer

        ' Autofit columns
        xlWorksheet.Columns("A:F").AutoFit
        xlWorksheet.Cells.EntireRow.AutoFit
    Next int_sldIndex

    FormatWorksheets

    'Add in proper code for where to save. This is an admin routine, so it may not need as much care as the game code

    ' Save and close the workbook
    xlWorkbook.SaveAs "G:\Angry Birds 2 WIP\Questions.xlsx"
    xlWorkbook.Close

    ' Quit Excel
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "Saved at G:\Angry Birds 2 WIP\Questions.xlsx"
End Sub
#End If

Private Function WorksheetExists(xlWorkbook As Object, worksheetName As String) As Boolean
    Dim xlSheet As Object
    Dim foundWorksheet As Boolean
    
    For Each xlSheet In xlWorkbook.Worksheets
        If xlSheet.Name = worksheetName Then
            foundWorksheet = True
            Exit For
        End If
    Next xlSheet
   
    If Not foundWorksheet Then
        For Each xlSheet In xlWorkbook.Worksheets
            If xlSheet.Name = "Sheet1" Then
                xlSheet.Name = worksheetName
                foundWorksheet = True
                Exit For
            End If
        Next xlSheet
    End If
    
    WorksheetExists = foundWorksheet
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''
'                     OTHER
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ConvertStringToLong(ByVal stringToConvert As String, Optional ByVal lengthToKeep As Integer = 10, Optional ByVal textSeparator As String = "-") As Long
    If Len(stringToConvert) > 10 Then
        stringToConvert = Left(stringToConvert, lengthToKeep)
    End If
    
    stringToConvert = Replace(stringToConvert, textSeparator, "")
    
    ConvertStringToLong = CLng(stringToConvert)
End Function

Private Function RandBetween(ByVal lowestChoice As Integer, ByVal highestChoice As Integer) As Integer
    Dim i As Integer
    
    Randomize Day(Date) 'Seed Rnd() with a psuedo random number
    i = 10000 * Rnd()
    Randomize i 'Re-seed Rnd() with a more random number
    RandBetween = Int(lowestChoice + (highestChoice - lowestChoice + 1) * Rnd())
End Function

Private Function GetTextPosition(ByVal originalText As String) As Integer
    GetTextPosition = Trim(Mid(originalText, Int(InStr(originalText, ":")) + 1))
End Function

Private Function IsValuePresentInArray(ByVal valueToFind As String, ByRef arrayToSearch() As String) As Boolean
    Dim i As Integer
    
    For i = LBound(arrayToSearch) To UBound(arrayToSearch)
        If arrayToSearch(i) = valueToFind Then
            IsValuePresentInArray = True
            Exit Function
        End If
    Next i
    
    IsValuePresentInArray = False
End Function

#If Win64 Then
Private Function ConvertOneDriveToLocalPath(ByVal folderPath As String) As String
    Dim i As Integer
    
    If Left(folderPath, 23) = "https://d.docs.live.net" Or Left(folderPath, 11) = "OneDrive://" Then
        For i = 1 To 4 ' Strip the OneDrive URI part of the path (everything before the 4th '/')
            folderPath = Mid(folderPath, InStr(folderPath, "/") + 1)
        Next
        
        folderPath = Replace(folderPath, "/", pathSeparator)
        
        ConvertOneDriveToLocalPath = Environ$("OneDrive") & pathSeparator & folderPath
    Else
        folderPath = Replace(folderPath, "/", pathSeparator)
        ConvertOneDriveToLocalPath = folderPath
    End If
End Function
#End If

Private Function VerifyFileOrFolderExists(ByVal pathToCheck As String) As Boolean
    #If Mac Then
        Dim pathExists As Boolean
        pathExists = AppleScriptTask("AngryBirds.scpt", "ExistsFile", pathToCheck)
        
        If Not pathExists Then
            pathExists = AppleScriptTask("AngryBirds.scpt", "ExistsFolder", pathToCheck)
        End If
        
        VerifyFileOrFolderExists = pathExists
    #Else
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        VerifyFileOrFolderExists = (fs.fileExists(pathToCheck) Or fs.FolderExists(pathToCheck))
    #End If
End Function

Private Sub ChangeAssignedMacros()
    Dim sld As Slide
    
    Set sld = ActivePresentation.Slides(8)
    SearchSlideAndReassignMacros sld
    
    Set sld = ActivePresentation.Slides(9)
    SearchSlideAndReassignMacros sld
End Sub

Private Sub SearchSlideAndReassignMacros(sld As Slide)
    Dim shp As Shape
    
    For Each shp In sld.Shapes
        With shp.ActionSettings(ppMouseClick)
            If .Action = ppActionRunMacro Then
                If .Run = "CreateNewConfig" Then
                    .Run = "UpdateConfiguration"
                    Debug.Print "Updated: " & shp.Name
                End If
            End If
        End With
    Next shp
End Sub
