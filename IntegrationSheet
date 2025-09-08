Private Function getNumOfTask() As Integer
    Dim numOfTask As Integer: numOfTask = Cells(3, 12)
    getNumOfTask = numOfTask
End Function
Private Function getListOfTaskID() As String
    Dim i As Integer
    Dim taskID As String
    
    For i = 1 To getNumOfTask()
        taskID = taskID & Cells(4 - 1 + i, 1) & Space(1)
    Next i
    
    getListOfTaskID = taskID
End Function
Private Function MsgBoxNotification(InputMsg As String) As String
    Dim AnswerYes, taskID, ShellStr As String
    Dim i As Integer
    
    For i = 1 To getNumOfTask()
        taskID = taskID & Cells(4 - 1 + i, 1) & vbNewLine
    Next i
    
    MsgBoxNotification = MsgBox("Do you want to " & InputMsg & " for" & vbNewLine & taskID, vbQuestion + vbYesNo)
End Function
Private Sub clearCurrentTaskList_Click()
    Dim AnswerYes As String
    AnswerYes = MsgBox("Do you want to clear current task list?", vbQuestion + vbYesNo)
    
    If AnswerYes = vbYes Then
        Range("A4:G51").ClearContents
    Else
        Exit Sub
    End If
    
End Sub

Private Sub GetTaskStatus_Click()
    Sheets("Integrator").Select
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim exportissues, taskID, fields, ShellStr, Expath As String
    
    'Get Task Status
    exportissues = "im exportissues "
    fields = " --fields=Type,Status,Summary,desired_Release,Integrated,integration_Comments "
    taskID = getListOfTaskID()
    
    'Define output file
    Expath = Environ("AppData")
    Expath = Left(Expath, InStr(Expath, "AppData") + 7)
    Expath = Expath & "Local\Temp\PTC_Integrator_export.xls"
    
    ShellStr = exportissues & " --exportHeadings --noopenOutputFile --overwriteOutputFile --notrimExcessColumns" _
             & " --outputFile=" & Chr(34) & Expath & Chr(34) _
             & fields _
             & taskID
    wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle, waitOnReturn

    'Update task status
    Sheets("Integrator").Select
    Dim src, des As Workbook
    
    Set src = Workbooks.Open(Expath, True, True)
    Set des = ThisWorkbook
    
    src.Worksheets("Sheet0").Range("A2:F" & getNumOfTask + 1).Copy
    des.Worksheets("Integrator").Range("B4").PasteSpecial xlPasteValues
    
    src.Close False    ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
End Sub
Private Sub tickIntegrated_Click()
    Sheets("Integrator").Select
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    If MsgBoxNotification("tick Integrated") = vbYes Then
        'tick Integrated
        Dim editissues, taskID, tickStatus As String
        
        editissues = "im editissue --field=Integrated="
        tickStatus = Cells(9, 12)
        taskID = getListOfTaskID()
        
        ShellStr = editissues & tickStatus & " " & taskID
        
        wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle, waitOnReturn
    Else
        Exit Sub
    End If
    
    
End Sub
Private Sub updateDesiredRelease_Click()
    Sheets("Integrator").Select
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim targetDesiredRelease As String
    
    targetDesiredRelease = Cells(6, 12)
    If MsgBoxNotification("update Desired Release " & Chr(34) & targetDesiredRelease & Chr(34)) = vbYes Then
        'Update desire release
        Dim editissues, taskID As String
        
        editissues = "im editissue --field=desired_Release="
        
        taskID = getListOfTaskID()
        
        ShellStr = editissues & targetDesiredRelease & " " & taskID
        
        wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle, waitOnReturn
    Else
        Exit Sub
    End If
End Sub

Private Sub updateIntegrationComments_Click()
    Sheets("Integrator").Select
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    If MsgBoxNotification("update Integration Comments") = vbYes Then
        'Update desire release
        Dim editissues, integrationComments, taskID As String
        
        editissues = "im editissue --field=integration_Comments="
        integrationComments = Cells(12, 12)
        taskID = getListOfTaskID()
        
        ShellStr = editissues & Chr(34) & integrationComments & Chr(34) & " " & taskID
        wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle, waitOnReturn
    Else
        Exit Sub
    End If
End Sub


