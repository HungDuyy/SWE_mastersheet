Option Explicit


Private Sub GenMeeting_Click()
    Sheets("ReviewMaster").Select
    Dim AnswerYes, SWEphase  As String
    
    SWEphase = Cells(13, 3)
    
    If SWEphase = "1" Or SWEphase = "3a" Or SWEphase = "3b" Or SWEphase = "4" Then
        AnswerYes = MsgBox("Create Team Meeting for SWE." & SWEphase & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Create Meeting")
    Else
           MsgBox ("Wrong SWE Phase (1, 3a, 3b, 4), please check !")
           Exit Sub
    End If
    
    If AnswerYes = vbYes Then
        If (SWEphase = "1") Then
            CreateOutlook ("1")
        ElseIf (SWEphase = "3a") Then
            CreateOutlook ("3a")
        ElseIf (SWEphase = "3b") Then
            CreateOutlook ("3b")
        ElseIf (SWEphase = "4") Then
            CreateOutlook ("4")
        End If
    End If
End Sub

Private Sub GenReviewMaster_Click()
    Sheets("ReviewMaster").Select
    Dim AnswerYes  As String
    Dim SWEphase As String
    Dim checklist As String
    Dim isExist As String
    
    SWEphase = Cells(13, 3)
    
    If (SWEphase = "1") Then
        checklist = Cells(6, 3)
        isExist = Dir(checklist)
    ElseIf (SWEphase = "3a") Then
        checklist = Cells(7, 3)
        isExist = Dir(checklist)
    ElseIf (SWEphase = "3b") Then
        checklist = Cells(8, 3)
        isExist = Dir(checklist)
    Else
        checklist = Cells(9, 3)
        isExist = Dir(checklist)
    End If
    
    
    
    If (SWEphase = "1" Or SWEphase = "3a" Or SWEphase = "3b" Or SWEphase = "4") And Not isExist = "" Then
        AnswerYes = MsgBox("Create Review Master for SWE." & SWEphase & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Create Review Master")
    ElseIf (SWEphase = "1" Or SWEphase = "3a" Or SWEphase = "3b" Or SWEphase = "4") And isExist = "" Then
        AnswerYes = MsgBox("Checklist is not found!" + vbCrLf + "Create review master for SWE." & SWEphase & " without checklist?", vbExclamation + vbYesNo + vbDefaultButton2, "Create Review Master")
    Else
        MsgBox ("Wrong process SWE (1, 3a, 3b, 4), please check !")
        Exit Sub
    End If
    
    If AnswerYes = vbYes Then
        If (SWEphase = "1") Then
            CreateRM ("1")
        ElseIf (SWEphase = "3a") Then
            CreateRM ("3a")
        ElseIf (SWEphase = "3b") Then
            CreateRM ("3b")
        ElseIf (SWEphase = "4") Then
            CreateRM ("4")
        End If
    End If
End Sub
Public Sub CreateOutlook(phase As String)
    Sheets("ReviewMaster").Select
    On Error GoTo Err_Execute
     
    Dim olApp As Outlook.Application
    Dim olAppt As Outlook.AppointmentItem
    Dim blnCreated As Boolean
    Dim olNs As Outlook.Namespace
    Dim CalFolder As Outlook.MAPIFolder
     
    Dim i As Long
     
    On Error Resume Next
    Set olApp = Outlook.Application
     
    If olApp Is Nothing Then
        Set olApp = Outlook.Application
         blnCreated = True
        Err.Clear
    Else
        blnCreated = False
    End If
     
    On Error GoTo 0
     
    Set olNs = olApp.GetNamespace("MAPI")
    Set CalFolder = olNs.GetDefaultFolder(olFolderCalendar)
    If phase = "1" Then
        i = 33
    ElseIf phase = "3a" Then
        i = 34
    ElseIf phase = "3b" Then
        i = 35
    ElseIf phase = "4" Then
        i = 36
    Else
        MsgBox ("Wrong process SWE")
        Exit Sub
    End If
    
    'Do Until Trim(Cells(i, 12).Value) = ""
     
    Set olAppt = CalFolder.Items.Add(olAppointmentItem)
    With olAppt
       .MeetingStatus = olMeeting
    'Define calendar item properties
        .Subject = Cells(i, 7)
        
    ' don't use a location if using a resource
      ' .Location = Cells(i, 12)
        .body = Cells(i, 9)
        '.Categories = Cells(i, 14)
        .Start = Cells(i, 10) + Cells(i, 11)
        .End = Cells(i, 12) + Cells(i, 13)
        .Location = "Microsoft Teams Meeting"
        .BusyStatus = olBusy
        .ReminderMinutesBeforeStart = Cells(i, 14)
        .ReminderSet = True

'## Start Recipient code
' get the recipients
        Dim RequiredAttendee, OptionalAttendee, ResourceAttendee As Outlook.Recipient
        Set RequiredAttendee = .Recipients.Add(Cells(i, 15).Value)
            RequiredAttendee.Type = olRequired
        Set OptionalAttendee = .Recipients.Add(Cells(i, 16).Value)
            OptionalAttendee.Type = olOptional
        'Set ResourceAttendee = .Recipients.Add(Cells(i, 12).Value)
           ' ResourceAttendee.Type = olResource
'## End Recipient code
' For meetings or Group Calendars
' use .Display instead of .Send when testing or if you want to review before sending
        '.Send

        .Display

        
    End With
         
        'i = i + 1
        'Loop
    Set olAppt = Nothing
    Set olApp = Nothing
    
    SendKeys "{F10}", True

    'Switch to ribbon shortcuts

    SendKeys "H", True

    'Hit the Microsoft teams meetings button, requires teams to be installed

    SendKeys "Y1", True
    SendKeys "^~" 'Press CTRL+ENTER ^ means CTRL, ~ means ENTER
    Exit Sub
     
Err_Execute:
    MsgBox "An error occurred - Exporting items to Calendar."
     

End Sub

Public Sub CreateRM(phase As String)
    Sheets("ReviewMaster").Select
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim PTCpath, field(15), createissue, hostname, port, Wtype, body, body_end, rvobject, pmandatory, poptional As String
    Dim ShellStr, temp As String
    Dim Owner(2), Sw3aModuletest(2), Sw3bDev2(2), Sw4Dev(2), Swe1(2), Swe1d(2) As String
    Dim Swe3a(2), Swe3ad(2), Swe3b(2), Swe3bd(2), Swe4(2), Swe4d(2), Swe5(2), Systest(2), PM(2), SwAn(2), SwAr(2), SwArd(2), QD(2), InteTest(2) As String
    Dim checklist As String
    
    Dim strFileExists As String
    
    PTCpath = Cells(1, 3) 'IntegrityClient\bin
    Owner(0) = Cells(16, 3) 'Owner name
    Owner(1) = Cells(16, 4) ' Owner account ID: nguydu18
    Sw3aModuletest(0) = Cells(17, 3) 'name
    Sw3aModuletest(1) = Cells(17, 4) 'account ID: nguydu18

    If Cells(13, 3) = "3a" And (Sw3aModuletest(0) = "" Or Sw3aModuletest(1) = "") Then
        MsgBox ("Process = SWE.3a but not found SW3a Module Test data !")
        Exit Sub
    End If
    Sw3bDev2(0) = Cells(18, 3) 'name
    Sw3bDev2(1) = Cells(18, 4) 'account ID: nguydu18
    If Cells(13, 3) = "3b" And (Sw3bDev2(0) = "" Or Sw3bDev2(1) = "") Then
        MsgBox ("Process = SWE.3b but not found SW3b Dev 2 data!")
        Exit Sub
    End If
    Sw4Dev(0) = Cells(19, 3)
    Sw4Dev(1) = Cells(19, 4)
    
    Swe1(0) = Cells(2, 11) 'name
    Swe1(1) = Cells(2, 12) 'account ID: nguydu18

    Swe1d(0) = Cells(2, 14)
    Swe1d(1) = Cells(2, 15)
    If Swe1d(1) <> "" Then
        Swe1d(1) = ", " & Swe1d(1)
    Else
        Swe1d(0) = "not defined"
        Swe1d(1) = ""
    End If
    
    Swe3a(0) = Cells(3, 11)
    Swe3a(1) = Cells(3, 12)
    Swe3ad(0) = Cells(3, 14)
    Swe3ad(1) = Cells(3, 15)
    If Swe3ad(1) <> "" Then
        Swe3ad(1) = ", " & Swe3ad(1)
    Else
        Swe3ad(0) = "not defined"
        Swe3ad(1) = ""
    End If
    
    Swe3b(0) = Cells(4, 11)
    Swe3b(1) = Cells(4, 12)
    Swe3bd(0) = Cells(4, 14)
    Swe3bd(1) = Cells(4, 15)
    If Swe3bd(1) <> "" Then
        Swe3bd(1) = ", " & Swe3bd(1)
    Else
        Swe3bd(0) = "not defined"
        Swe3bd(1) = ""
    End If
    
    Swe4(0) = Cells(5, 11)
    Swe4(1) = Cells(5, 12)
    Swe4d(0) = Cells(5, 14)
    Swe4d(1) = Cells(5, 15)
    If Swe4d(1) <> "" Then
        Swe4d(1) = ", " & Swe4d(1)
    Else
        Swe4d(0) = "not defined"
        Swe4d(1) = ""
    End If
    
    Swe5(0) = Cells(6, 11)
    Swe5(1) = Cells(6, 12)
    Systest(0) = Cells(7, 11)
    Systest(1) = Cells(7, 12)
    PM(0) = Cells(8, 11)
    PM(1) = Cells(8, 12)
    SwAn(0) = Cells(9, 11)
    SwAn(1) = Cells(9, 12)
    SwAr(0) = Cells(10, 11)
    SwAr(1) = Cells(10, 12)
    SwArd(0) = Cells(10, 14)
    SwArd(1) = Cells(10, 15)
    QD(0) = Cells(11, 11)
    QD(1) = Cells(11, 12)
    InteTest(0) = Cells(12, 11)
    InteTest(1) = Cells(12, 12)
    
    
    createissue = "im createissue  -g"
    hostname = " --hostname=" & Chr(34) & Cells(2, 3) & Chr(34)
    port = " --port=" & Chr(34) & Cells(3, 3) & Chr(34)
    Wtype = " --type=" & Chr(34) & "W_Review Master" & Chr(34)
    field(0) = " --field=" & Chr(34) & "Projekt=" & Cells(22, 3) & Chr(34)
        temp = Cells(21, 3) 'task name
        temp = Replace(temp, Chr(34), Chr(39))
        'temp = Replace(temp, Chr(39), Chr(34))
        'temp = Replace(temp, "<", "^<")
        'temp = Replace(temp, ">", "^>")
    field(1) = " --field=" & Chr(34) & "Summary=" & Cells(4, 3) & "_" & Cells(12, 3) & ": " & temp  'ProjectNam_ReleasePhase : exp: FBD6_BLE_SW_C1:....task name
    field(2) = " --field=" & Chr(34) & "Backward related_Review_Master=" & Cells(15, 3) & Chr(34) 'link task to RM
    field(3) = " --field=" & Chr(34) & "Start_Date=" & Cells(23, 3) & Chr(34)
    field(4) = " --field=" & Chr(34) & "Planned_Target_Date=" & Cells(24, 3) & Chr(34)
    field(5) = " --field=" & Chr(34) & "Actual_Target_Date=" & Cells(25, 3) & Chr(34)
    field(6) = " --field=" & Chr(34) & "Functional_Safety_relevant=" & Cells(5, 3) & Chr(34)
    field(7) = " --field=" & Chr(34) & "Overall_Responsible=" & PM(1) & Chr(34)
    field(8) = " --field=" & Chr(34) & "Workflow_Responsible=" & Owner(1) & Chr(34)
    
    body_end = "^<b^>Review method:^</b^> ^<br^>" _
                            & "Review Meeting ^<br^>^<br^> " _
               & "^<b^>Review Result:^</b^> ^<br^> " _
                            & "a) Next steps: ^<br^>" _
                                & " - complete re-review necessary ^<br^>" _
                                & " - final-review with following participants [person1, person2] on content [chapter 123] necessary ^<br^>" _
                                & " - close findings without further review ^<br^>" _
                                & " - no further activities necessary ^<br^>^<br^>" _
                            & "b) Status ^<br^>" _
                                & " - Release of Task content^<br^>" _
                                & " - Conditional Release of Task content with condition: ...^<br^>" _
                                & " - No Release of Task content: Rework of Task content ^<br^>^<br^>" _
               & "^<b^>Review Findings:^</b^> ^<br^>" _
                            & " - (add small Findings here) ^<br^>" _
                            & " or^<br^>" _
                            & " - see ReviewFindings under tab Relations ^<br^>" _
                            & " or^<br^>" _
                            & "- no Findings ^<br^>^<br^>" & Chr(39)

 '1=========================================================================================================================
    If phase = "1" Then
        checklist = Cells(6, 3) 'review Checklist
        field(1) = field(1) & " - " & Cells(2, 9) & Chr(34) 'Subject extension: Software Requirements
        field(9) = " --field=" & Chr(34) & "Review_Phase=" & Cells(2, 9) & Chr(34) 'Review_Phase=Software Requirements
        field(10) = " --field=" & Chr(34) & "Review_Method=" & Cells(10, 8) & Chr(34)
        field(11) = " --field=" & Chr(34) & "Review_Participants=" & SwAn(1) _
                                                                   & ", " & SwAr(1) _
                                                                   & ", " & SwArd(1) _
                                                                   & ", " & Systest(1) _
                                                                   & ", " & Owner(1) _
                                                                   & ", " & QD(1) _
                                                                   & ", " & PM(1) _
                                                                   & Swe1d(1) _
                                                                   & ", " & Swe1(1) & Chr(34)
                                                                   
        field(12) = " --field=" & Chr(34) & "Review_Checklist_Template=" & Cells(3, 8) & Chr(34)
        field(13) = " --field=" & Chr(34) & "Workproducts=" & Cells(2, 8) & Chr(34)
        field(14) = ""
        
        strFileExists = Dir(checklist)
        If Not strFileExists = "" Then
            field(14) = " --addAttachment=" & Chr(34) & "field=Attachments,path=" & checklist & ",summary=Review checklist" & Chr(34)
        End If
        
        temp = Cells(33, 2) 'review objects contents
        temp = Replace(temp, Chr(34), "^" & Chr(34))
        temp = Replace(temp, Chr(39), "^&^#39")
        temp = Replace(temp, "<", "^<")
        temp = Replace(temp, ">", "^>")
        temp = Replace(temp, vbCrLf, "^<br^>")
        temp = Replace(temp, vbLf, "^<br^>")
        rvobject = "^<b^>Review Objects:^</b^> ^<br^>^<br^>" & temp & "^<br^>^<br^>"
        pmandatory = "^<b^>Invited persons(Mandatory):^</b^>^<br^>" _
                    & "SW Analyst: " & SwAn(0) & "^<br^>" _
                    & "SW Analyst 2: " & Sw3bDev2(0) & "^<br^>" _
                    & "SW Architect: " & SwAr(0) & "^<br^>" _
                    & "SW Architect Deputy: " & SwArd(0) & "^<br^>" _
                    & "System Test Manager -^> SW Test Manager: " & Systest(0) & " (Manager will forward the meeting to tester)^<br^>" _
                    & "SW-Quality: " & QD(0) & "^<br^>^<br^>"
        poptional = "^<b^>Invited persons(Informative/Optional):^</b^>^<br^>" _
                    & "SWE.1 Resp: " & Swe1(0) & "^<br^>" _
                    & "SWE.1 Deputy: " & Swe1d(0) & "^<br^>" _
                    & "PM: " & PM(0) & "^<br^>^<br^>"
        
        body = " --richContentField=" & Chr(39) & "Description=Review Date: " & Cells(14, 3) & "^<br^>^<br^>" _
                & rvobject & pmandatory & poptional _
                & "^<b^>Review Participants:^</b^>^<br^>" _
                    & "See below field: Review Participants^<br^>^<br^> " _
                & "^<b^>Review content:^</b^>" _
                & "Filter for H_ChangeRequestReference: " & Cells(15, 3) & ", (number_of) Requirements^<br^>" _
                & "1) Check of attributes (use the view: Software_Writing_Requirements)^<br^>" _
                & "2) Check of traceability to Sys Requirements and Interfaces^<br^>" _
                & "3) Check of ReqPat Warnings (Set filter for requirements and do check with CTRL+W)^<br^>" _
                & "4) Check if SW-Requirements with none link to System Requirements have set H_Object_Source set to SW-Base-Requirement and if there are really none System Requirements which can be linked or if System Requirements have to be created.^<br^>" _
                & "5) Check the checklist and add checklist to the attachments. The checklist can be found here ^" & Chr(34) & "e:/Projects/CAPE/BMW/FBD6/50_FBD6/10_MGMT/80_REQ_M/20_Plan/20_ProcessDocuments/10_SWE1_SW-Requirements/HF-8347_GE_2022-04-04.docm ^" & Chr(34) & "^<br^>" _
                & "6) Perform check with S.A.U.. The Tool can be found in Doors -^> Central Tools -^> Standard-Attributes-Utility. All SWE.1 checks must be green after the review. ^<br^>^<br^>"
    
        ShellStr = createissue & hostname & port & Wtype _
                & field(0) & field(1) & field(2) & field(3) _
                & field(4) & field(5) & field(6) & field(7) _
                & field(8) & field(9) & field(10) & field(11) _
                & field(12) & field(13) & field(14) & body & body_end
        
        If PTCpath <> "" Then
            wsh.CurrentDirectory = PTCpath
        End If
        wsh.Run "cmd.exe /S /C " & ShellStr, windowStyle ', waitOnReturn
'3a=========================================================================================================================
    ElseIf phase = "3a" Then
            checklist = Cells(7, 3)
            field(1) = field(1) & " - " & Cells(3, 9) & Chr(34)
            field(9) = " --field=" & Chr(34) & "Review_Phase=" & Cells(3, 9) & Chr(34)
            field(10) = " --field=" & Chr(34) & "Review_Method=" & Cells(11, 8) & Chr(34)
            field(11) = " --field=" & Chr(34) & "Review_Participants=" & Swe3a(1) _
                                                                       & Swe3ad(1) _
                                                                       & ", " & SwAn(1) _
                                                                       & ", " & SwAr(1) _
                                                                       & ", " & SwArd(1) _
                                                                       & ", " & Owner(1) _
                                                                       & ", " & Sw3aModuletest(1) _
                                                                       & "," & InteTest(1) _
                                                                       & ", " & QD(1) & Chr(34)
                                                                       
            field(12) = " --field=" & Chr(34) & "Review_Checklist_Template=" & Cells(5, 8) & Chr(34)
            field(13) = " --field=" & Chr(34) & "Workproducts=" & Cells(4, 8) & Chr(34)
            field(14) = ""
        
            strFileExists = Dir(checklist)
            If Not strFileExists = "" Then
                field(14) = " --addAttachment=" & Chr(34) & "field=Attachments,path=" & checklist & ",summary=Review checklist" & Chr(34)
            End If
            
            temp = Cells(34, 2) 'review objects contents
            temp = Replace(temp, Chr(34), "^" & Chr(34))
            temp = Replace(temp, Chr(39), "^&^#39")
            temp = Replace(temp, "<", "^<")
            temp = Replace(temp, ">", "^>")
            temp = Replace(temp, vbCrLf, "^<br^>")
            temp = Replace(temp, vbLf, "^<br^>")
            rvobject = "^<b^>Review Objects:^</b^>^<br^>^<br^>" & temp & "^<br^>^<br^>"
            
            pmandatory = "^<b^>Invited persons(Mandatory):^</b^>^<br^>" _
                        & "SW Analyst: " & SwAn(0) & "^<br^>" _
                        & "SW MUT: " & Sw3aModuletest(0) & "^<br^>" _
                        & "SW Architect: " & SwAr(0) & "^<br^>" _
                        & "SW Architect Deputy: " & SwArd(0) & "^<br^>^<br^>"
                        
            poptional = "^<b^>Invited persons(Informative/Optional):^</b^>^<br^>" _
                        & "SWE.3a Resp.: " & Swe3a(0) & "^<br^>" _
                        & "SWE.3a Deputy: " & Swe3ad(0) & "^<br^>" _
                        & "PM: " & PM(0) & "^<br^>" _
                        & "SW Integration Test: " & InteTest(0) & "^<br^>" _
                        & "SW-Quality: " & QD(0) & "^<br^>^<br^>"
                        
            body = " --richContentField=" & Chr(39) & "Description=Review Date: " & Cells(14, 3) & "^<br^>^<br^>" _
                    & rvobject & pmandatory & poptional _
                    & "^<b^>Review Participants:^</b^>^<br^>" _
                            & "See below field: Review Participants^<br^>^<br^> " _
                    & "^<b^>Review content:^</b^>^<br^>" _
                            & "a) See attached checklist under Attachments and see also guideline: ^" & Chr(34) & "e:/Projects/CAPE/BMW/FBD6/50_FBD6/40_SW/15_Module_Design/FBD6_Common_SWE3a_Detailed_Design_Guideline.docx^" & Chr(34) & " ^<br^>" _
                            & "b) SW design changes in Rhapsody ^<br^>^<br^> "
        
            ShellStr = createissue & hostname & port & Wtype _
                    & field(0) & field(1) & field(2) & field(3) _
                    & field(4) & field(5) & field(6) & field(7) _
                    & field(8) & field(9) & field(10) & field(11) _
                    & field(12) & field(13) & field(14) & body & body_end

            If PTCpath <> "" Then
                wsh.CurrentDirectory = PTCpath
            End If
            wsh.Run "cmd.exe /S /C " & ShellStr, windowStyle ', waitOnReturn
'3b=========================================================================================================================
    ElseIf phase = "3b" Then
            checklist = Cells(8, 3)
            field(1) = field(1) & " - " & Cells(4, 9) & Chr(34)
            field(9) = " --field=" & Chr(34) & "Review_Phase=" & Cells(4, 9) & Chr(34)
            field(10) = " --field=" & Chr(34) & "Review_Method=" & Cells(12, 8) & Chr(34)
            field(11) = " --field=" & Chr(34) & "Review_Participants=" & Swe3b(1) _
                                                                       & Swe3bd(1) _
                                                                       & ", " & SwAn(1) _
                                                                       & ", " & Owner(1) _
                                                                       & ", " & Swe5(1) _
                                                                       & ", " & QD(1) _
                                                                       & ", " & PM(1) _
                                                                       & ", " & Sw3bDev2(1) & Chr(34)
                                                                       
            field(12) = " --field=" & Chr(34) & "Review_Checklist_Template=" & Cells(7, 8) & Chr(34)
            field(13) = " --field=" & Chr(34) & "Workproducts=" & Cells(6, 8) & Chr(34)
            field(14) = " --addAttachment=" & Chr(34) & "field=Attachments,path=" & checklist & ",summary=Review checklist" & Chr(34)
            field(14) = ""
        
            strFileExists = Dir(checklist)
            If Not strFileExists = "" Then
                field(14) = " --addAttachment=" & Chr(34) & "field=Attachments,path=" & checklist & ",summary=Review checklist" & Chr(34)
            End If
            
            temp = GetSWE3bReviewOBJ()
            Cells(35, 2) = temp
            temp = Replace(temp, Chr(34), "^" & Chr(34))
            temp = Replace(temp, Chr(39), "^&^#39")
            temp = Replace(temp, "<", "^<")
            temp = Replace(temp, ">", "^>")
            temp = Replace(temp, vbCrLf, "^<br^>")
            temp = Replace(temp, vbLf, "^<br^>")
            rvobject = "^<b^>Review Objects:^</b^>^<br^>^<br^>" & temp & "^<br^>^<br^>"
            
            pmandatory = "^<b^>Invited persons(Mandatory):^</b^>^<br^>" _
                        & "SW Analyst: " & SwAn(0) & "^<br^>" _
                        & "SW Developer: " & Owner(0) & "^<br^>" _
                        & "SW Developer 2: " & Sw3bDev2(0) & "^<br^>^<br^>"
                        
            poptional = "^<b^>Invited persons(Informative/Optional):^</b^>^<br^>" _
                        & "SWE.3b Resp.: " & Swe3b(0) & "^<br^>" _
                        & "SWE.3b Deputy: " & Swe3bd(0) & "^<br^>" _
                        & "SW Integrator: " & Swe5(0) & "^<br^>" _
                        & "PM: " & PM(0) & "^<br^>" _
                        & "SW-Quality: " & QD(0) & "^<br^>^<br^>"
                        
            body = " --richContentField=" & Chr(39) & "Description=Review Date: " & Cells(14, 3) & "^<br^>^<br^>" _
                    & rvobject & pmandatory & poptional _
                    & "^<b^>Review Participants:^</b^>^<br^>" _
                            & "See below field: Review Participants^<br^>^<br^> " _
                    & "^<b^>Review content:^</b^>^<br^>" _
                            & "See attached checklist under Attachments and see also guideline: ^" & Chr(34) & "e:/Projects/CAPE/BMW/FBD6/50_FBD6/10_MGMT/80_REQ_M/20_Plan/20_ProcessDocuments/30_SWE3_SWConstruction/FBD6_SWE3b_CodeReviewCheckList_HF-8352_GE_2022-06-07.docm^" & Chr(34) & "^<br^>^<br^>"
                    
        
            ShellStr = createissue & hostname & port & Wtype _
                    & field(0) & field(1) & field(2) & field(3) _
                    & field(4) & field(5) & field(6) & field(7) _
                    & field(8) & field(9) & field(10) & field(11) _
                    & field(12) & field(13) & field(14) & body & body_end
                    
            If PTCpath <> "" Then
             wsh.CurrentDirectory = PTCpath
            End If
            wsh.Run "cmd.exe /S /C " & ShellStr, windowStyle ', waitOnReturn
'4=========================================================================================================================
    ElseIf phase = "4" Then
            checklist = Cells(9, 3)
            field(1) = field(1) & " - " & Cells(5, 9) & Chr(34) 'add  extention
            field(9) = " --field=" & Chr(34) & "Review_Phase=SWE.4" & Chr(34)
            field(10) = " --field=" & Chr(34) & "Review_Method=" & Cells(13, 8) & Chr(34)
            field(11) = " --field=" & Chr(34) & "Review_Participants=" & Swe4(1) _
                                                                       & ", " & Sw4Dev(1) _
                                                                       & ", " & SwAn(1) _
                                                                       & Swe4d(1) _
                                                                       & ", " & Owner(1) _
                                                                       & ", " & Swe5(1) _
                                                                       & ", " & QD(1) _
                                                                       & ", " & PM(1) & Chr(34)
                                                                       
            field(12) = " --field=" & Chr(34) & "Review_Checklist_Template=" & Cells(9, 8) & Chr(34)
            field(13) = " --field=" & Chr(34) & "Workproducts=" & Cells(8, 8) & Chr(34)
            field(14) = " --addAttachment=" & Chr(34) & "field=Attachments,path=" & checklist & ",summary=Review checklist" & Chr(34)
            field(14) = ""
        
            strFileExists = Dir(checklist)
            If Not strFileExists = "" Then
                field(14) = " --addAttachment=" & Chr(34) & "field=Attachments,path=" & checklist & ",summary=Review checklist" & Chr(34)
            End If
            
            temp = GetSWE4ReviewOBJ()
            Cells(36, 2) = temp
            temp = Replace(temp, Chr(34), "^" & Chr(34))
            temp = Replace(temp, Chr(39), "^&^#39")
            temp = Replace(temp, "<", "^<")
            temp = Replace(temp, ">", "^>")
            temp = Replace(temp, vbCrLf, "^<br^>")
            temp = Replace(temp, vbLf, "^<br^>")
            rvobject = "^<b^>Review Objects:^</b^>^<br^>^<br^>" & temp & "^<br^>^<br^>"
            
            pmandatory = "^<b^>Invited persons(Mandatory):^</b^>^<br^>" _
                        & "SW Developer: " & Sw4Dev(0) & "^<br^>" _
                        & "SWE.4 Resp.: " & Swe4(0) & "^<br^>" _
                        & "SW-Quality: " & QD(0) & "^<br^>" _
                        & "SW Integrator: " & Swe5(0) & "^<br^>^<br^>"
                        
            poptional = "^<b^>Invited persons(Informative/Optional):^</b^>^<br^>" _
                        & "SWE.4 Deputy: " & Swe4d(0) & "^<br^>" _
                        & "PM: " & PM(0) & "^<br^>^<br^>"
                        
            body = " --richContentField=" & Chr(39) & "Description=Review Date: " & Cells(14, 3) & "^<br^>^<br^>" _
                    & rvobject & pmandatory & poptional _
                    & "^<b^>Review Participants:^</b^>^<br^>" _
                            & "See below field: Review Participants^<br^>^<br^> " _
                    & "^<b^>Review content:^</b^>^<br^>" _
                            & "MUT is based on the: " & Cells(20, 3) & "^<br^>" _
                            & "a) Check of overview report^<br^>" _
                            & "b) See attached checklist ^<br^>^<br^>"
                    
        
            ShellStr = createissue & hostname & port & Wtype _
                    & field(0) & field(1) & field(2) & field(3) _
                    & field(4) & field(5) & field(6) & field(7) _
                    & field(8) & field(9) & field(10) & field(11) _
                    & field(12) & field(13) & field(14) & body & body_end
        If PTCpath <> "" Then
            wsh.CurrentDirectory = PTCpath
        End If
        wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle ', waitOnReturn
    Else
    
        MsgBox ("Wrong phase, please try again")
        
    End If
End Sub
Private Sub import_Click()

    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim PTCpath, Expath, task, phase, exportissues, hostname, port, ShellStr, temp As String
    
    Sheets("ReviewMaster").Select
    
    Cells(14, 3) = Date
    task = Cells(15, 3)
    phase = Cells(13, 3)
    exportissues = "im exportissues"
    hostname = " --hostname=" & Chr(34) & Cells(2, 3) & Chr(34)
    port = " --port=" & Chr(34) & Cells(3, 3) & Chr(34)
    
    Expath = Environ("AppData")
    Expath = Left(Expath, InStr(Expath, "AppData") + 7)
    Expath = Expath & "Local\Temp\PTC_task_export.xls"
    
    ShellStr = exportissues & hostname & port & " --exportHeadings --noopenOutputFile --overwriteOutputFile --notrimExcessColumns" _
    & " --fields=Projekt,Summary,Start_Date,Planned_Target_Date,Actual_Target_Date,desired_Release_Report" _
    & " --outputFile=" & Chr(34) & Expath & Chr(34) & " " & task
    PTCpath = Cells(1, 3)
    If PTCpath <> "" Then
            wsh.CurrentDirectory = PTCpath
    End If
    wsh.Run "cmd.exe /S /C " & ShellStr & "", windowStyle, waitOnReturn
    
   ' Application.Wait (Now + TimeValue("0:00:02"))
    
    Dim src As Workbook
    Dim Projekt, Summary, StartDate, PlannedDate, ActualDate, DesiredRelease As String
    Set src = Workbooks.Open(Expath, True, True)
    Projekt = src.Worksheets("sheet0").Cells(2, 1)
    Summary = src.Worksheets("sheet0").Cells(2, 2)
    StartDate = src.Worksheets("sheet0").Cells(2, 3)
    PlannedDate = src.Worksheets("sheet0").Cells(2, 4)
    ActualDate = src.Worksheets("sheet0").Cells(2, 5)
    DesiredRelease = src.Worksheets("sheet0").Cells(2, 6)
    src.Close False    ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
    Summary = Right(Summary, Len(Summary) - Len(Left(Summary, InStr(Summary, ": "))) - 1)
    
    Cells(21, 3).Value = Summary
    Cells(22, 3).Value = Projekt
    Cells(23, 3).Value = StartDate
    Cells(24, 3).Value = PlannedDate
    Cells(25, 3).Value = ActualDate
    Cells(20, 3).Value = DesiredRelease
       
    If phase = "3a" Or phase = "3b" Then
        temp = GetSWE3aReviewOBJ()
        Cells(34, 2) = temp
        temp = GetSWE3bReviewOBJ()
        Cells(35, 2) = temp
    'ElseIf phase = "3b" Then
       'temp = GetSWE3bReviewOBJ()
        'Cells(35, 2) = temp
    ElseIf phase = "4" Then
        temp = GetSWE4ReviewOBJ()
        Cells(36, 2) = temp
    End If
    Application.ScreenUpdating = True
    
End Sub
Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
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
Public Function GetSWE3aReviewOBJ() As String
    Dim ShellStr As String, getcpinfo As String, Result As String
    Dim Rhapsody As String
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")

    ShellStr = "im viewcp --attributes=no --entryAttributes=member,revision,project " & Cells(15, 3)
    getcpinfo = ShellRun(ShellStr)
    With RE
            .MultiLine = True
            .Global = True
        'Remove Entries:
            .Pattern = ".*Entries:.*\r?\n"
            getcpinfo = .Replace(getcpinfo, "")
            
        'Extract Rhapsody
    
            'Remove all path except sbsx
            .Pattern = "^(?!.*sbsx.*).+$"
            Rhapsody = .Replace(getcpinfo, "")
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            Rhapsody = .Replace(Rhapsody, "")
            .Pattern = "[\r\n]+"
            Rhapsody = .Replace(Rhapsody, vbNewLine)
            
         'Result of Change Package
            
            Result = "Rhapsody: " & vbNewLine _
                                  & Rhapsody
            .Pattern = "\t"
            Result = Replace(Result, "e:/Projects", vbCrLf & "(e:/Projects")
            Result = Replace(Result, "/project.pj", ")" & vbCrLf)
            GetSWE3aReviewOBJ = Result
    End With
End Function
Public Function GetSWE3bReviewOBJ() As String
    Dim ShellStr As String, getcpinfo As String, Result As String
    Dim Rhapsody As String, Source As String, QAC As String
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")

    ShellStr = "im viewcp --attributes=no --entryAttributes=member,revision,project " & Cells(15, 3)
    getcpinfo = ShellRun(ShellStr)
    With RE
            .MultiLine = True
            .Global = True
        'Remove Entries:
            .Pattern = ".*Entries:.*\r?\n"
            getcpinfo = .Replace(getcpinfo, "")
            
        'Extract Rhapsody
    
            'Remove all path except sbsx
            .Pattern = "^(?!.*sbsx.*).+$"
            Rhapsody = .Replace(getcpinfo, "")
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            Rhapsody = .Replace(Rhapsody, "")
            .Pattern = "[\r\n]+"
            Rhapsody = .Replace(Rhapsody, vbNewLine)
        'Extract QAC
        
            'Remove all path except html
            .Pattern = "^(?!.*html.*).+$"
            QAC = .Replace(getcpinfo, "")
            
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            QAC = .Replace(QAC, "")
            
            .Pattern = "[\r\n]+"
            QAC = .Replace(QAC, vbNewLine)

        'Extract Source
        
            'Remove QAC path
            .Pattern = "^(.*html.*).+$"
            Source = .Replace(getcpinfo, "")

            'Remove Rhapsody path
            .Pattern = "^(.*sbsx.*).+$"
            Source = .Replace(Source, "")
            
            'Remove tmb path
            .Pattern = "^(.*tmb.*).+$"
            Source = .Replace(Source, "")
            
            'Remove pdf path
            .Pattern = "^(.*pdf.*).+$"
            Source = .Replace(Source, "")
            
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            Source = .Replace(Source, "")
            
            .Pattern = "[\r\n]+"
            Source = .Replace(Source, vbNewLine)
            
            
         'Result of Change Package
            
            Result = "Rhapsody: " & vbNewLine _
                                  & Rhapsody _
            & "Sources: " & vbNewLine _
                                  & Source _
                        & "QAC: " & vbNewLine _
                                  & QAC
            .Pattern = "\t"
            
            Result = Replace(Result, "e:/Projects", vbCrLf & "(e:/Projects")
            Result = Replace(Result, "/project.pj", ")" & vbCrLf)
            GetSWE3bReviewOBJ = Result
    End With
End Function
Public Function GetSWE4ReviewOBJ() As String
    Dim ShellStr As String, getcpinfo As String, Result As String
    Dim MUT As String, MUT_1 As String, MUT_2 As String
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")

    ShellStr = "im viewcp --attributes=no --entryAttributes=member,revision,project " & Cells(15, 3)
    getcpinfo = ShellRun(ShellStr)
    With RE
            .MultiLine = True
            .Global = True
        'Remove Entries:
            .Pattern = ".*Entries:.*\r?\n"
            getcpinfo = .Replace(getcpinfo, "")
     
        'Extract MUT
        
            'Remove all path except tmb
            .Pattern = "^(?!.*tmb.*).+$"
            MUT_1 = .Replace(getcpinfo, "")

            'Remove all path except PDF
            .Pattern = "^(?!.*pdf.*).+$"
            MUT_2 = .Replace(getcpinfo, "")
            
            MUT = MUT_1 & MUT_2
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            MUT = .Replace(MUT, "")
            
            .Pattern = "[\r\n]+"
            MUT = .Replace(MUT, vbNewLine)
            
         'Result of Change Package
            
            Result = "MUT: " & vbNewLine _
                                  & MUT
            .Pattern = "\t"
            Result = Replace(Result, "e:/Projects", vbCrLf & "(e:/Projects")
            Result = Replace(Result, "/project.pj", ")" & vbCrLf)
            GetSWE4ReviewOBJ = Result
    End With
End Function
Private Sub GetChangePackageDetail_Click()
    Dim ShellStr As String, getcpinfo As String, Result As String
    Dim Rhapsody As String, Source As String, QAC As String, MUT As String, MUT_1 As String, MUT_2 As String
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")

    ShellStr = "im viewcp --attributes=no --entryAttributes=member,revision,project " & Cells(15, 12)
    getcpinfo = ShellRun(ShellStr)
    With RE
            .MultiLine = True
            .Global = True
        'Remove Entries:
            .Pattern = ".*Entries:.*\r?\n"
            getcpinfo = .Replace(getcpinfo, "")
            
        'Extract Rhapsody
    
            'Remove all path except sbsx
            .Pattern = "^(?!.*sbsx.*).+$"
            Rhapsody = .Replace(getcpinfo, "")
            
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            Rhapsody = .Replace(Rhapsody, "")
            
            .Pattern = "[\r\n]+"
            Rhapsody = .Replace(Rhapsody, vbNewLine)

        'Extract QAC
        
            'Remove all path except html
            .Pattern = "^(?!.*html.*).+$"
            QAC = .Replace(getcpinfo, "")
            
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            QAC = .Replace(QAC, "")
            
            .Pattern = "[\r\n]+"
            QAC = .Replace(QAC, vbNewLine)
        
        'Extract MUT
        
            'Remove all path except tmb
            .Pattern = "^(?!.*tmb.*).+$"
            MUT_1 = .Replace(getcpinfo, "")
            
            'Remove all path except PDF
            .Pattern = "^(?!.*pdf.*).+$"
            MUT_2 = .Replace(getcpinfo, "")
            
            MUT = MUT_1 & MUT_2
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            MUT = .Replace(MUT, "")
            
            .Pattern = "[\r\n]+"
            MUT = .Replace(MUT, vbNewLine)

        'Extract Source
        
            'Remove QAC path
            .Pattern = "^(.*html.*).+$"
            Source = .Replace(getcpinfo, "")

            'Remove Rhapsody path
            .Pattern = "^(.*sbsx.*).+$"
            Source = .Replace(Source, "")
            
            'Remove tmb path
            .Pattern = "^(.*tmb.*).+$"
            Source = .Replace(Source, "")
            
            'Remove pdf path
            .Pattern = "^(.*pdf.*).+$"
            Source = .Replace(Source, "")
            
            'Remove empty line
            .Pattern = "^[\t ]*\n"
            Source = .Replace(Source, "")
            
            .Pattern = "[\r\n]+"
            Source = .Replace(Source, vbNewLine)
            
         'Result of Change Package
            
            Result = "Rhapsody: " & vbNewLine _
                                  & Rhapsody _
                                  & vbNewLine _
                                  & vbNewLine _
            & "Sources: " & vbNewLine _
                                  & Source _
                                  & vbNewLine _
                                  & vbNewLine _
                        & "QAC: " & vbNewLine _
                                  & QAC _
                                  & vbNewLine _
                                  & vbNewLine _
                        & "MUT: " & vbNewLine _
                                  & MUT
            .Pattern = "\t"
            Result = .Replace(Result, " __ ")
    End With
    Cells(17, 11) = Result
End Sub

