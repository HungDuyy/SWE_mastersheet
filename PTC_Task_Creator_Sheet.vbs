Private Sub CommandButton1_Click()
    Sheets("Create W_Task").Select
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    Dim PTCpath, field(15), createissue, hostname, port, Wtype, Tasktype, body, Owner, PM As String
    Dim ShellStr As String
    Dim i As Integer
       
    PTCpath = Cells(4, 3)
    PM = Cells(5, 4)
    
    i = 10
    
    Do Until Trim(Cells(i, 2).Value) = ""
    Tasktype = Cells(i, 6)
    Owner = Cells(i, 7)
    Owner = Right(Owner, Len(Owner) - InStr(Owner, "("))
    Owner = Replace(Owner, ")", "")
    
    createissue = "im createissue  -g"
    hostname = " --hostname=" & Chr(34) & Cells(2, 3) & Chr(34)
    port = " --port=" & Chr(34) & Cells(3, 3) & Chr(34)
    Wtype = " --type=" & Chr(34) & "W_Task" & Chr(34)
    field(0) = " --field=" & Chr(34) & "Projekt=" & Cells(i, 3) & Chr(34)
    field(1) = " --field=" & Chr(34) & "Summary=" & Cells(i, 4) & Chr(34)
    field(2) = " --field=" & Chr(34) & "Hella_priority=" & Cells(i, 8) & Chr(34)
    field(3) = " --field=" & Chr(34) & "Start_Date=" & Cells(i, 10) & Chr(34)
    field(4) = " --field=" & Chr(34) & "Planned_Target_Date=" & Cells(i, 11) & Chr(34)
    field(5) = " --field=" & Chr(34) & "Actual_Target_Date=" & Cells(i, 12) & Chr(34)
    field(6) = " --field=" & Chr(34) & "Functional_Safety_relevant=no" & Chr(34)
    field(7) = " --field=" & Chr(34) & "Overall_Responsible=" & PM & Chr(34)
    field(8) = " --field=" & Chr(34) & "Workflow_Responsible=" & Owner & Chr(34)
    field(9) = " --field=" & Chr(34) & "affected_Element=" & Cells(i, 13) & Chr(34)
    field(10) = " --field=" & Chr(34) & "desired_Release=" & Cells(i, 14) & Chr(34)
    field(11) = " --field=" & Chr(34) & "Initial_Planned_Effort=" & Cells(i, 5) & Chr(34)
    field(12) = " --field=" & Chr(34) & "Remaining_Effort_Task=" & Cells(i, 5) & Chr(34)
    field(13) = " --field=" & Chr(34) & "Backward related_Task_" & Tasktype & "=" & Cells(i, 2) & Chr(34)
                                        
      
    body = Cells(i, 9)
    body = Replace(body, Chr(34), "^" & Chr(34))
    body = Replace(body, Chr(39), "^&^#39")
    body = Replace(body, "<", "^<")
    body = Replace(body, ">", "^>")
    body = Replace(body, vbCrLf, "^<br^>")
    body = Replace(body, vbLf, "^<br^>")
        
    body = " --richContentField=Description=" & Chr(39) & body & "^<br^>^<br^>" & Chr(39)
                
                
    
    ShellStr = createissue & hostname & port & Wtype _
                & field(0) & field(1) & field(2) & field(3) _
                & field(4) & field(5) & field(6) & field(7) _
                & field(8) & field(9) & field(10) _
                & field(11) & field(12) & field(13) & body

    If PTCpath <> "" Then
            wsh.CurrentDirectory = PTCpath
    End If
    wsh.Run "cmd.exe /S /C " & ShellStr, windowStyle ', waitOnReturn
    
    i = i + 1
    Loop
    
End Sub

Private Sub CommandButton2_Click()
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1

    Dim PTCpath, Expath, Cr, exportissues, hostname, port, ShellStr As String
    
    Dim c, i, j As Integer
    
    Dim src As Workbook
    Dim Projekt, Summary, Descrip, prio, StartDate, PlannedDate, ActualDate, affectedElement, DesiredRelease As String
    
    Sheets("Create W_Task").Select
    
    PTCpath = Cells(4, 3)
    
    Cr = " "
    c = 10
    Do Until Trim(Cells(c, 2).Value) = ""
    Cr = Cr & Cells(c, 2) & " "
    c = c + 1
    Loop
    
    exportissues = "im exportissues"
    hostname = " --hostname=" & Chr(34) & Cells(2, 3) & Chr(34)
    port = " --port=" & Chr(34) & Cells(3, 3) & Chr(34)
    
    Expath = Environ("AppData")
    Expath = Left(Expath, InStr(Expath, "AppData") + 7)
    Expath = Expath & "Local\Temp\PTC_Cr_export.xls"
    
    ShellStr = exportissues & hostname & port & " --exportHeadings --noopenOutputFile --overwriteOutputFile --notrimExcessColumns" _
    & " --fields=Projekt,Summary,Description,Hella_priority,Start_Date,Planned_Target_Date,Actual_Target_Date,affected_Element,desired_Release" _
    & " --outputFile=" & Chr(34) & Expath & Chr(34) & Cr
    '& " --outputFile=D:\PTC_Cr_export.xls" & Cr
    Cells(13, 14).Value = ShellStr
    If PTCpath <> "" Then
            wsh.CurrentDirectory = PTCpath
    End If
    wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle, waitOnReturn
    
    
   ' Application.Wait (Now + TimeValue("0:00:02"))
    Set src = Workbooks.Open(Expath, True, True)
    i = 2
    j = 10
    
    Do Until Trim(src.Worksheets("sheet0").Cells(i, 3).Value) = ""
    
    Projekt = src.Worksheets("sheet0").Cells(i, 1)
    Summary = src.Worksheets("sheet0").Cells(i, 2)
    Descrip = src.Worksheets("sheet0").Cells(i, 3)
    prio = src.Worksheets("sheet0").Cells(i, 4)
    StartDate = src.Worksheets("sheet0").Cells(i, 5)
    PlannedDate = src.Worksheets("sheet0").Cells(i, 6)
    ActualDate = src.Worksheets("sheet0").Cells(i, 7)
    affectedElement = src.Worksheets("sheet0").Cells(i, 8)
    DesiredRelease = src.Worksheets("sheet0").Cells(i, 9)
    
    Cells(j, 3).Value = Projekt
    Cells(j, 4).Value = Summary
    Cells(j, 5).Value = ""
    Cells(j, 6).Value = ""
    Cells(j, 7).Value = ""
    Cells(j, 8).Value = prio
    Cells(j, 9).Value = Descrip
    Cells(j, 10).Value = StartDate
    Cells(j, 11).Value = PlannedDate
    Cells(j, 12).Value = ActualDate
    Cells(j, 13).Value = affectedElement
    Cells(j, 14).Value = DesiredRelease
    
    i = i + 1
    j = j + 1
    
    Loop
    
    src.Close False    ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
    'Summary = Right(Summary, Len(Summary) - Len(Left(Summary, InStr(Summary, ": "))) - 1)
    Application.ScreenUpdating = True
    
End Sub
