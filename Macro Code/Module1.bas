Attribute VB_Name = "Module1"
Sub EchoRunExecutableAndAddAudio()
    Dim wsh As Object
    Dim exePath As String
    Dim audioPath As String
    Dim currentDir As String
    Dim slide As slide
    Dim shape As shape
    Dim notesText As String
    Dim cmd As String
    Dim process As Object
    Dim fileExists As Boolean
    Dim output As String
    Dim line As String
    Dim slideIndex As Integer
    Dim delayBeforePlayback As Single
    Dim timeout As Single
    Dim slideShowWindow As slideShowWindow

    ' Timing parameters
    delayBeforePlayback = 2000 ' 2 seconds before audio starts playing (in milliseconds)

    ' Get the path of the current PowerPoint presentation
    currentDir = ActivePresentation.Path
    MsgBox "Current directory: " & currentDir

    ' Ensure the folder exists
    If Dir(currentDir & "\Text_To_Speech_Voices", vbDirectory) = "" Then
        MsgBox "Error: Folder 'Text_To_Speech_Voices' not found in " & currentDir
        Exit Sub
    End If

    ' Set the path to your executable file
    exePath = currentDir & "\Text_To_Speech_Voices\Echo_TTS_For_PP_Macro.exe"
    MsgBox "Executable path: " & exePath

    ' Check if the executable file exists
    If Dir(exePath) = "" Then
        MsgBox "Error: Executable file not found at " & exePath
        Exit Sub
    End If

    ' Get the path where the MP3 file will be saved
    slideIndex = Application.ActiveWindow.View.slide.slideIndex
    audioPath = currentDir & "\Slide_" & slideIndex & ".mp3"
    MsgBox "Audio path: " & audioPath

    ' Get the text from the notes field
    Set slide = Application.ActiveWindow.View.slide
    If Not slide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
        MsgBox "Error: No text found in the notes field."
        Exit Sub
    End If
    notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    MsgBox "Text to be sent for Text-to-Speech: " & vbCrLf & notesText

    ' Create a WScript Shell object
    Set wsh = CreateObject("WScript.Shell")

    ' Construct the command to run the executable with the notes text as an argument
    cmd = """" & exePath & """ """ & Replace(notesText, """", """""") & """ """ & audioPath & """"
    MsgBox "Command to run: " & cmd

    ' Run the executable and capture the output
    On Error GoTo ErrHandler
    Set process = wsh.Exec(cmd)

    ' Read the standard output of the process
    output = ""
    Do While Not process.StdOut.AtEndOfStream
        line = process.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop
    MsgBox "Output from executable: " & vbCrLf & output

    ' Wait for the process to complete
    Do While process.Status = 0
        DoEvents
    Loop

    ' Check if the MP3 file exists with a timeout mechanism
    timeout = Timer + 2 ' Wait up to 2 seconds
    Do While Timer < timeout And Dir(audioPath) = ""
        DoEvents
    Loop

    ' Ensure the MP3 file exists
    fileExists = Dir(audioPath) <> ""

    If fileExists Then
        MsgBox "MP3 file found at " & audioPath

        ' Remove any existing audio shapes from the slide
        For Each shape In slide.Shapes
            If shape.Type = msoMedia Then
                shape.Delete
            End If
        Next shape

        ' Add the MP3 file to the slide
        Set shape = slide.Shapes.AddMediaObject2(audioPath, msoFalse, msoTrue, 100, 100)

        ' Set the audio to play automatically and not loop
        With shape.AnimationSettings
            .PlaySettings.PlayOnEntry = msoTrue
            .PlaySettings.HideWhileNotPlaying = msoTrue
            .PlaySettings.LoopUntilStopped = msoFalse ' Prevent looping
        End With

        MsgBox "Audio added!"
    Else
        MsgBox "Error: MP3 file not found at " & audioPath
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error running executable: " & Err.Description

End Sub




Sub RunExecutableAndAddAudioAlloy()
    Dim wsh As Object
    Dim exePath As String
    Dim audioPath As String
    Dim currentDir As String
    Dim slide As slide
    Dim shape As shape
    Dim notesText As String
    Dim cmd As String
    Dim process As Object
    Dim fileExists As Boolean
    Dim output As String
    Dim line As String
    Dim slideIndex As Integer
    Dim delayBeforePlayback As Single
    Dim timeout As Single
    Dim slideShowWindow As slideShowWindow

    ' Timing parameters
    delayBeforePlayback = 2000 ' 2 seconds before audio starts playing (in milliseconds)

    ' Get the path of the current PowerPoint presentation
    currentDir = ActivePresentation.Path

    ' Set the path to your executable file
    exePath = currentDir & "\Text_To_Speech_Voices\Alloy_TTS_For_PP_Macro.exe"

    ' Get the current slide
    Set slide = Application.ActiveWindow.View.slide
    slideIndex = slide.slideIndex

    ' Set the path where the MP3 file will be saved
    audioPath = currentDir & "\Slide_" & slideIndex & ".mp3"

    ' Check if the executable file exists
    If Dir(exePath) = "" Then
        MsgBox "Error: Executable file not found at " & exePath
        Exit Sub
    End If

    ' Get the text from the notes field
    If Not slide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
        MsgBox "Error: No text found in the notes field."
        Exit Sub
    End If
    notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

    ' Display the text to be sent for Text-to-Speech
    MsgBox "Sending the following for Text to Speech: " & vbCrLf & notesText

    ' Create a WScript Shell object
    Set wsh = CreateObject("WScript.Shell")

    ' Construct the command to run the executable with the notes text as an argument
    cmd = """" & exePath & """ """ & Replace(notesText, """", """""") & """ """ & audioPath & """"

    ' Run the executable and capture the output
    On Error GoTo ErrHandler
    Set process = wsh.Exec(cmd)

    ' Read the standard output of the process
    output = ""
    Do While Not process.StdOut.AtEndOfStream
        line = process.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop

    ' Display the output from the executable
    MsgBox "Output from executable: " & vbCrLf & output

    ' Wait for the process to complete
    Do While process.Status = 0
        DoEvents
    Loop

    ' Check if the MP3 file exists with a timeout mechanism
    timeout = Timer + 2 ' Wait up to 10 seconds
    Do While Timer < timeout And Dir(audioPath) = ""
        DoEvents
    Loop

    ' Ensure the MP3 file exists
    fileExists = Dir(audioPath) <> ""

    If fileExists Then
        MsgBox "MP3 file found at " & audioPath

        ' Remove any existing audio shapes from the slide
        For Each shape In slide.Shapes
            If shape.Type = msoMedia Then
                shape.Delete
            End If
        Next shape

        ' Add the MP3 file to the slide
        Set shape = slide.Shapes.AddMediaObject2(audioPath, msoFalse, msoTrue, 100, 100)

        ' Set the audio to play automatically and not loop
        With shape.AnimationSettings
            .PlaySettings.PlayOnEntry = msoTrue
            .PlaySettings.HideWhileNotPlaying = msoTrue
            .PlaySettings.LoopUntilStopped = msoFalse ' Prevent looping
        End With

        MsgBox "Audio added!"
    Else
        MsgBox "Error: MP3 file not found at " & audioPath
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error running executable: " & Err.Description

End Sub

Sub RunExecutableAndAddAudioFable()
    Dim wsh As Object
    Dim exePath As String
    Dim audioPath As String
    Dim currentDir As String
    Dim slide As slide
    Dim shape As shape
    Dim notesText As String
    Dim cmd As String
    Dim process As Object
    Dim fileExists As Boolean
    Dim output As String
    Dim line As String
    Dim slideIndex As Integer
    Dim delayBeforePlayback As Single
    Dim timeout As Single
    Dim slideShowWindow As slideShowWindow

    ' Timing parameters
    delayBeforePlayback = 2000 ' 2 seconds before audio starts playing (in milliseconds)

    ' Get the path of the current PowerPoint presentation
    currentDir = ActivePresentation.Path

    ' Set the path to your executable file
    exePath = currentDir & "\Text_To_Speech_Voices\Fable_TTS_For_PP_Macro.exe"

    ' Get the current slide
    Set slide = Application.ActiveWindow.View.slide
    slideIndex = slide.slideIndex

    ' Set the path where the MP3 file will be saved
    audioPath = currentDir & "\Slide_" & slideIndex & ".mp3"

    ' Check if the executable file exists
    If Dir(exePath) = "" Then
        MsgBox "Error: Executable file not found at " & exePath
        Exit Sub
    End If

    ' Get the text from the notes field
    If Not slide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
        MsgBox "Error: No text found in the notes field."
        Exit Sub
    End If
    notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

    ' Display the text to be sent for Text-to-Speech
    MsgBox "Sending the following for Text to Speech: " & vbCrLf & notesText

    ' Create a WScript Shell object
    Set wsh = CreateObject("WScript.Shell")

    ' Construct the command to run the executable with the notes text as an argument
    cmd = """" & exePath & """ """ & Replace(notesText, """", """""") & """ """ & audioPath & """"

    ' Run the executable and capture the output
    On Error GoTo ErrHandler
    Set process = wsh.Exec(cmd)

    ' Read the standard output of the process
    output = ""
    Do While Not process.StdOut.AtEndOfStream
        line = process.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop

    ' Display the output from the executable
    MsgBox "Output from executable: " & vbCrLf & output

    ' Wait for the process to complete
    Do While process.Status = 0
        DoEvents
    Loop

    ' Check if the MP3 file exists with a timeout mechanism
    timeout = Timer + 2 ' Wait up to 10 seconds
    Do While Timer < timeout And Dir(audioPath) = ""
        DoEvents
    Loop

    ' Ensure the MP3 file exists
    fileExists = Dir(audioPath) <> ""

    If fileExists Then
        MsgBox "MP3 file found at " & audioPath

        ' Remove any existing audio shapes from the slide
        For Each shape In slide.Shapes
            If shape.Type = msoMedia Then
                shape.Delete
            End If
        Next shape

        ' Add the MP3 file to the slide
        Set shape = slide.Shapes.AddMediaObject2(audioPath, msoFalse, msoTrue, 100, 100)

        ' Set the audio to play automatically and not loop
        With shape.AnimationSettings
            .PlaySettings.PlayOnEntry = msoTrue
            .PlaySettings.HideWhileNotPlaying = msoTrue
            .PlaySettings.LoopUntilStopped = msoFalse ' Prevent looping
        End With

        MsgBox "Audio added!"
    Else
        MsgBox "Error: MP3 file not found at " & audioPath
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error running executable: " & Err.Description

End Sub

Sub RunExecutableAndAddAudioOnyx()
    Dim wsh As Object
    Dim exePath As String
    Dim audioPath As String
    Dim currentDir As String
    Dim slide As slide
    Dim shape As shape
    Dim notesText As String
    Dim cmd As String
    Dim process As Object
    Dim fileExists As Boolean
    Dim output As String
    Dim line As String
    Dim slideIndex As Integer
    Dim delayBeforePlayback As Single
    Dim timeout As Single
    Dim slideShowWindow As slideShowWindow

    ' Timing parameters
    delayBeforePlayback = 2000 ' 2 seconds before audio starts playing (in milliseconds)

    ' Get the path of the current PowerPoint presentation
    currentDir = ActivePresentation.Path

    ' Set the path to your executable file
    exePath = currentDir & "\Text_To_Speech_Voices\Onyx_TTS_For_PP_Macro.exe"

    ' Get the current slide
    Set slide = Application.ActiveWindow.View.slide
    slideIndex = slide.slideIndex

    ' Set the path where the MP3 file will be saved
    audioPath = currentDir & "\Slide_" & slideIndex & ".mp3"

    ' Check if the executable file exists
    If Dir(exePath) = "" Then
        MsgBox "Error: Executable file not found at " & exePath
        Exit Sub
    End If

    ' Get the text from the notes field
    If Not slide.NotesPage.Shapes.Placeholders(2).TextFrame.HasText Then
        MsgBox "Error: No text found in the notes field."
        Exit Sub
    End If
    notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

    ' Display the text to be sent for Text-to-Speech
    MsgBox "Sending the following for Text to Speech: " & vbCrLf & notesText

    ' Create a WScript Shell object
    Set wsh = CreateObject("WScript.Shell")

    ' Construct the command to run the executable with the notes text as an argument
    cmd = """" & exePath & """ """ & Replace(notesText, """", """""") & """ """ & audioPath & """"

    ' Run the executable and capture the output
    On Error GoTo ErrHandler
    Set process = wsh.Exec(cmd)

    ' Read the standard output of the process
    output = ""
    Do While Not process.StdOut.AtEndOfStream
        line = process.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop

    ' Display the output from the executable
    MsgBox "Output from executable: " & vbCrLf & output

    ' Wait for the process to complete
    Do While process.Status = 0
        DoEvents
    Loop

    ' Check if the MP3 file exists with a timeout mechanism
    timeout = Timer + 2 ' Wait up to 10 seconds
    Do While Timer < timeout And Dir(audioPath) = ""
        DoEvents
    Loop

    ' Ensure the MP3 file exists
    fileExists = Dir(audioPath) <> ""

    If fileExists Then
        MsgBox "MP3 file found at " & audioPath

        ' Remove any existing audio shapes from the slide
        For Each shape In slide.Shapes
            If shape.Type = msoMedia Then
                shape.Delete
            End If
        Next shape

        ' Add the MP3 file to the slide
        Set shape = slide.Shapes.AddMediaObject2(audioPath, msoFalse, msoTrue, 100, 100)

        ' Set the audio to play automatically and not loop
        With shape.AnimationSettings
            .PlaySettings.PlayOnEntry = msoTrue
            .PlaySettings.HideWhileNotPlaying = msoTrue
            .PlaySettings.LoopUntilStopped = msoFalse ' Prevent looping
        End With

        MsgBox "Audio added!"
    Else
        MsgBox "Error: MP3 file not found at " & audioPath
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error running executable: " & Err.Description

End Sub

Sub SetSlideTransitionAfterAudio()
    Dim slideIndex As Integer
    Dim slideCount As Integer
    Dim shapeIndex As Integer
    Dim shapeCount As Integer
    Dim audioShape As shape
    Dim audioLength As Single
    Dim delayBeforeAudio As Single
    Dim delayBeforeTransition As Single

    delayBeforeAudio = 1 ' Delay before audio starts in seconds
    delayBeforeTransition = 1 ' Delay before slide transition in seconds

    slideCount = ActivePresentation.Slides.Count

    For slideIndex = 1 To slideCount
        shapeCount = ActivePresentation.Slides(slideIndex).Shapes.Count
        For shapeIndex = 1 To shapeCount
            Set audioShape = ActivePresentation.Slides(slideIndex).Shapes(shapeIndex)
            
            ' Check if the shape is a media object
            On Error Resume Next
            If audioShape.Type = msoMedia Then
                ' Check if the media type is sound
                If audioShape.MediaType = ppMediaTypeSound Then
                    On Error GoTo 0
                    ' Assume audio length is available in seconds
                    ' If there's no direct way to get length, you'll need to handle it differently
                    audioLength = 10 ' Default value if length can't be retrieved

                    ' Set the audio to play automatically with a delay
                    With audioShape.AnimationSettings.PlaySettings
                        .PlayOnEntry = msoTrue
                        .HideWhileNotPlaying = msoTrue
                    End With

                    ' Hide the audio shape during the presentation
                    audioShape.Tags.Add "HideInPresentation", "True"

                    ' Show message box with audio length (commented out)
                    ' MsgBox "Slide " & slideIndex & ": Audio length is " & audioLength & " seconds."

                    ' Show message box for audio settings (commented out)
                    ' MsgBox "Slide " & slideIndex & ": Audio is set to play automatically with a " & delayBeforeAudio & " second delay."

                    ' Set the slide transition time to the audio length plus delays
                    With ActivePresentation.Slides(slideIndex).SlideShowTransition
                        .AdvanceOnTime = msoTrue
                        .AdvanceTime = audioLength + delayBeforeTransition
                    End With

                    ' Show message box for transition settings (commented out)
                    ' MsgBox "Slide " & slideIndex & ": Transition is set to occur after " & (audioLength + delayBeforeTransition) & " seconds."

                    ' Add a delay before the audio starts playing
                    audioShape.AnimationSettings.AdvanceTime = delayBeforeAudio
                Else
                    On Error GoTo 0
                    ' MsgBox "Slide " & slideIndex & ": Shape " & shapeIndex & " is a media object but not an audio file." (commented out)
                End If
            Else
                On Error GoTo 0
                ' MsgBox "Slide " & slideIndex & ": Shape " & shapeIndex & " is not a media object." (commented out)
            End If
            On Error GoTo 0
        Next shapeIndex
    Next slideIndex

    ' Final message to indicate the macro has finished (commented out)
    ' MsgBox "Macro completed. All slides processed."
    MsgBox "Macro completed. All slides processed."
End Sub
