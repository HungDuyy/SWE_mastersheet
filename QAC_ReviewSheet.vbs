
Private Sub btn1_Click()
Dim folderpath, workingpath, pythonscriptpath, logpath, outputpath, logrun As String

Dim MyFileSysObj As Object
Set MyFileSysObj = CreateObject("scripting.filesystemobject")


folderpath = TextBox1.Value
If Len(folderpath) > 120 Then
    MsgBox "Please choose input folder path < 120 character, current path is " & Len(inputpatch) & " characters", vbCritical
    Exit Sub
End If

If MyFileSysObj.FolderExists(folderpath) = True Then
Else
    MsgBox "QAC folder path not found !", vbCritical
    Exit Sub
End If

workingpath = Application.ActiveWorkbook.path
logpath = workingpath + "\data\LogsRun.txt"
pythonscriptpath = workingpath + "\data\QAC_Report_Review_Console.exe"

outputpath = Right(folderpath, Len(folderpath) - InStrRev(folderpath, "\"))
outputpath = Left(folderpath, Len(folderpath) - Len(outputpath)) + "QAC_Reviewed_File"
TextBox2.Value = outputpath

'Runing with show console log
'===================================================================
'Dim wsh As Object
'Set wsh = VBA.CreateObject("WScript.Shell")
'wsh.CurrentDirectory = workingpath & "\data\"
'wsh.Run pythonscriptpath & " " & folderpath, vbNormalFocus, True
'wsh.CurrentDirectory = "C:\"
'MsgBox "REVIEW DONE" & vbCrLf, vbInformation
'====================================================================
'Runing with save log
'====================================================================
Dim oWSH As Object
Set oWSH = CreateObject("WScript.Shell")
logrun = logrun & "Running Time: " & Time & vbCrLf
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FileToCreate = FSO.CreateTextFile(logpath)
'run command
oWSH.CurrentDirectory = workingpath & "\data\"
logrun = logrun & oWSH.Exec(pythonscriptpath & " " & folderpath).StdOut.ReadAll()

        'Dim sLine As String
        'Dim oExec, oOutput As Object
        'Set oExec = oWSH.Exec(pythonscriptpath & " " & folderpath)
        'While oExec.Status = WshRunning
           ' Application.Wait (Now + TimeValue("0:00:01"))
        'Wend
        'Set oOutput = oExec.StdOut
        'handle the results as they are written to and read from the StdOut object
        
        'While Not oOutput.atEndOfStream
            'sLine = oOutput.ReadLine
           ' If sLine <> "" Then logrun = logrun & sLine & vbCrLf
        'Wend

oWSH.CurrentDirectory = "C:\"
FileToCreate.Write logrun
FileToCreate.Close
MsgBox "REVIEW DONE" & vbCrLf & "For detail, please check logs at:  " & logpath, vbInformation
'=============================================================================================
End Sub

Private Sub CommandButton1_Click()
Dim sFolder As String
Dim strPath As String
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With

    If sFolder <> "" Then ' if a file was chosen
        ' *********************
        ' put your code in here
        strPath = Application.FileDialog( _
        msoFileDialogFolderPicker).SelectedItems(1)
        TextBox1.Value = strPath
        ' *********************
    End If
End Sub
Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd & "-r " & args)
    oShell.CurrentDirectory = "C:\"
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function
