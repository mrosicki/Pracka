Attribute VB_Name = "BPFR_Module_Admin"
Sub Admin_User_Format()

    Dim LastRow As Long
    Dim wb_name As String
    

    Application.ScreenUpdating = False

    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Administrator_Panel").Cells(Sheets("Administrator_Panel").Rows.Count, "A").End(xlUp).Row
    

    With wb.Sheets("Administrator_Panel").Range("E2:E" & LastRow + 20).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$AI$2:$AI$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With wb.Sheets("Administrator_Panel").Range("C2:C" & LastRow + 20).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$AM$2:$AM$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With wb.Sheets("Administrator_Panel").Range("D2:D" & LastRow + 20).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$AK$2:$AK$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    Application.CutCopyMode = False
    wb.Sheets("Administrator_Panel").Range("B2:E" & LastRow).Copy
    wb.Sheets("Administrator_Panel").Range("F2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    wb.Sheets("Administrator_Panel").Range("J2").Formula = "=IF(F2="""",IF(B2="""",""-"",""New""),IF(AND(B2=F2,C2=G2,D2=H2,E2=I2)=TRUE,""No-Updates"",""Updates""))"
    wb.Sheets("Administrator_Panel").Range("J2").Copy
    wb.Sheets("Administrator_Panel").Range("J3:J" & LastRow + 20).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    wb.Sheets("Administrator_Panel").Select
    wb.Sheets("Administrator_Panel").Range("A2").Select
    Application.CutCopyMode = False

    With wb.Sheets("Administrator_Panel").Range("B2:E" & LastRow + 20).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

    
    wb.Sheets("Administrator_Panel").Range("Q3").Value = wb.Sheets("Control").Range("R2").Value
    wb.Sheets("Administrator_Panel").Range("Q4").Value = wb.Sheets("Control").Range("R3").Value
    wb.Sheets("Administrator_Panel").Range("T3").Value = wb.Sheets("Control").Range("T2").Value
    wb.Sheets("Administrator_Panel").Range("T4").Value = wb.Sheets("Control").Range("U2").Value
    
    Application.ScreenUpdating = True

End Sub


Sub Update_User_List()

    Dim conn As ADODB.Connection
    Dim myRecordset As ADODB.Recordset
    Dim strConn As String
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    Dim Name_User As String
    Dim Key_SPS As String
    Dim Test_Change As Integer
    Dim Test_Validation As Long
    Dim Test_Complete As Long
    Dim Text_Mess As String
    
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    Application.ScreenUpdating = False
    
    '============================= Sub Bar Script Start =============================
    'Declare Sub Level Variables and Objects
    Dim Counter As Long
    Dim SubCounter As Long
    Dim TotalCount As Long


    'Declare the ProgressBar Objects
    Dim Subbar As ProgressBar

    'Initialize a New Instance of the Progressbars
    Set Subbar = New ProgressBar

    'Set all the Properties that need to be set before the
    'ProgresBar is Shown

    With Subbar
        .Title = "Sub Bar"
        .ExcelStatusBar = True
        .StartColour = rgbRed
        .EndColour = rgbGreen
    End With


    '============================= Sub Bar Script End =============================
    
    LastRow = wb.Sheets("Administrator_Panel").Cells(Sheets("Administrator_Panel").Rows.Count, "A").End(xlUp).Row

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = S:\BOM Leverage Database.accdb"
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
     "Data Source = " & strDB
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakrk-fp01\GroupJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
     
    'On Error Resume Next
        'wb.Sheets("Assign_Model_3_Years_Scope").ShowAllData
        'Err.Clear
    
    Test_Change = 0
    
    If wb.Sheets("Administrator_Panel").Cells(3, 13).Value > 0 Then
        Test_Change = 1
    End If
                 
                 
    If Test_Change = 0 Then
        MsgBox "There wasn't found any Updates to be made."
        Exit Sub
    End If
    
    Test_Change = wb.Sheets("Administrator_Panel").Cells(3, 13).Value
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = Test_Change
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
                
    'Name_User = UserName
     
    Set myRecordset = New ADODB.Recordset
    
    On Error GoTo Errr
    For i = 2 To LastRow
        If wb.Sheets("Administrator_Panel").Cells(i, 10) = "Updates" Then
            Key_SPS = wb.Sheets("Administrator_Panel").Cells(i, 6).Value
            'MsgBox Key_SPS
            With myRecordset
               .Open "Select * from User_Access Where User_Name = '" & Key_SPS & "'", _
                  strConn, adOpenKeyset, adUseClient
               .Fields("User_Name").Value = wb.Sheets("Administrator_Panel").Cells(i, 2).Value
               .Fields("User_Groupe").Value = wb.Sheets("Administrator_Panel").Cells(i, 3).Value
               .Fields("User_Domain").Value = wb.Sheets("Administrator_Panel").Cells(i, 4).Value
               .Fields("User_PBU").Value = wb.Sheets("Administrator_Panel").Cells(i, 5).Value
               .Update
               .Close
            End With
            '============================= Sub Bar Script Start =============================
                SubCounter = SubCounter + 1
                Subbar.NextAction "Updating Assignments in BOM Leverage Database... Please wait.", True
            '============================= Sub Bar Script End =============================
        End If
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    Set myRecordset = Nothing
    Set conn = Nothing

    Text_Mess = "All updates successful added to BOM Leverage Database for: [" & Test_Change & "] SPS Line Items."
    
    'MsgBox Test_Change & " updates successful added to BOM Leverage Database."
    MsgBox Text_Mess
    
    Application.ScreenUpdating = True
    Exit Sub
        
Errr:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Global Business Analyst."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = Test_Change
        Subbar.NextAction "Updating Assignments in BOM Leverage Database... Please wait.", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    myRecordset.Close
    
    '--------------
    
    'close the objects
    conn.Close
    
    'destroy the variables
    Set myRecordset = Nothing
    Set conn = Nothing
        

End Sub


Sub Export_User_List()

    'Using ADO to Export data from Excel worksheet (your host application) to an Access Database Table.
    'refer Image 10a to view the existing SalesManager Table in MS Access file "SalesReport.accdb"
    'refer Image 10b for data in Excel worksheet which is exported to Access Database Table.
    'refer Image 10c to view the SalesManager Table in Access file "SalesReport.accdb", after data is exported.
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, LastRow As Long, lFieldCount As Long, FirstRow As Long
    Dim lastID As String
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = ActiveWorkbook.Sheets("Administrator_Panel")
    
    If ws.Cells(4, 13).Value = 0 Then
        MsgBox "There wasn't found any new User to be added."
        Exit Sub
    End If
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    
    '--------------
    'THE CONNECTION OBJECT
    
    'strDBName = "BOM Leverage Database.accdb"
    'strMyPath = ThisWorkbook.Path
    'strDB = "S:\" & strDBName
    'strDB = "S:\" & strDBName
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;"
connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;"
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    Dim strTable As String
    
    strTable = "User_Access"
    adoRecSet.Open Source:=strTable, ActiveConnection:=connDB, CursorType:=adOpenStatic, LockType:=adLockOptimistic
    
    '--------------
    'COPY RECORDS FROM THE EXCEL WORKSHEET:
    'Note: Columns and their order should be the same in both Excel worksheet and in Access database table
    
    lFieldCount = adoRecSet.Fields.Count
    'determine last data row in the worksheet:
    LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    FirstRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
     
    'start copying from second row of worksheet, first row contains field names:
    For i = FirstRow To LastRow
    adoRecSet.AddNew
    For n = 1 To lFieldCount - 1
    adoRecSet.Fields(n).Value = ws.Cells(i, n + 1)
    Next n
    adoRecSet.Update
    Next i
    
    '--------------
    
    'close the objects
    adoRecSet.Close
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing

    MsgBox "Added " & ws.Cells(4, 13).Value & " new Users, into BOM Leverage - User Access Table."

End Sub


Sub BPRW1()

    ThisWorkbook.Sheets("Administrator_Panel").Visible = -1
    

End Sub

Sub BPRW2()

    ThisWorkbook.Sheets("Administrator_Panel").Visible = 2
    

End Sub


Sub Up_Control_Inf()

    
    Dim conn As ADODB.Connection
    Dim myRecordset As ADODB.Recordset
    Dim strConn As String
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    Dim Name_User As String
    Dim Key_SPS As String
    Dim Test_Change As Integer
    Dim Test_Validation As Long
    Dim Test_Complete As Long
    Dim Text_Mess As String
    
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    Application.ScreenUpdating = False
    
    '============================= Sub Bar Script Start =============================
    'Declare Sub Level Variables and Objects
    Dim Counter As Long
    Dim SubCounter As Long
    Dim TotalCount As Long


    'Declare the ProgressBar Objects
    Dim Subbar As ProgressBar

    'Initialize a New Instance of the Progressbars
    Set Subbar = New ProgressBar

    'Set all the Properties that need to be set before the
    'ProgresBar is Shown

    With Subbar
        .Title = "Sub Bar"
        .ExcelStatusBar = True
        .StartColour = rgbRed
        .EndColour = rgbGreen
    End With


    '============================= Sub Bar Script End =============================
    
    'LastRow = wb.Sheets("Administrator_Panel").Cells(Sheets("Administrator_Panel").Rows.Count, "A").End(xlUp).Row

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = S:\BOM Leverage Database.accdb"
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
     "Data Source = " & strDB
    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
     
     
    'On Error Resume Next
        'wb.Sheets("Assign_Model_3_Years_Scope").ShowAllData
        'Err.Clear
    
    Test_Change = 0
    
    If wb.Sheets("Administrator_Panel").Cells(3, 17).Value <> wb.Sheets("Control").Cells(2, 18).Value Then
        Test_Change = 1
    End If
    If wb.Sheets("Administrator_Panel").Cells(4, 17).Value <> wb.Sheets("Control").Cells(3, 18).Value Then
        Test_Change = 1
    End If
                 
                 
    If Test_Change = 0 Then
        MsgBox "There wasn't found any Updates to be made for Start Year or Version - please made any changes prior to executing script."
        Exit Sub
    End If
    
    Test_Change = 2
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = Test_Change
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
                
    'Name_User = UserName
     
    Set myRecordset = New ADODB.Recordset
    
    On Error GoTo Errr
    For i = 2 To 3
        Key_SPS = wb.Sheets("Control").Cells(i, 18).Value
        With myRecordset
           .Open "Select * from Start_Year Where Start_Year = '" & Key_SPS & "'", _
              strConn, adOpenKeyset, adUseClient
           .Fields("Start_Year").Value = wb.Sheets("Administrator_Panel").Cells(i + 1, 17).Value
           .Update
           .Close
        End With
        '============================= Sub Bar Script Start =============================
            SubCounter = SubCounter + 1
            Subbar.NextAction "Updating Assignments in BOM Leverage Database... Please wait.", True
        '============================= Sub Bar Script End =============================
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    Set myRecordset = Nothing
    Set conn = Nothing

    Text_Mess = "All updates successful added to BOM Leverage Database for."
    
    'MsgBox Test_Change & " updates successful added to BOM Leverage Database."
    MsgBox Text_Mess
    
    Application.ScreenUpdating = True
    Exit Sub
        
Errr:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Global Business Analyst."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = Test_Change
        Subbar.NextAction "Updating Assignments in BOM Leverage Database... Please wait.", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    myRecordset.Close
    
    '--------------
    
    'close the objects
    conn.Close
    
    'destroy the variables
    Set myRecordset = Nothing
    Set conn = Nothing
        


End Sub



Sub Up_Act_Window()

    
    Dim conn As ADODB.Connection
    Dim myRecordset As ADODB.Recordset
    Dim strConn As String
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    Dim Name_User As String
    Dim Key_SPS As String
    Dim Test_Change As Integer
    Dim Test_Validation As Long
    Dim Test_Complete As Long
    Dim Text_Mess As String
    Dim Start_Date As Date
    Dim End_Date As Date
    
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    Application.ScreenUpdating = False

    
    'LastRow = wb.Sheets("Administrator_Panel").Cells(Sheets("Administrator_Panel").Rows.Count, "A").End(xlUp).Row

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = S:\BOM Leverage Database.accdb"
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
     "Data Source = " & strDB
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BPFR SPS Database.accdb;" '& strDB
'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BPFR SPS Database.accdb;" '& strDB
strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BPFR SPS Database.accdb;" '& strDB
     
    'On Error Resume Next
        'wb.Sheets("Assign_Model_3_Years_Scope").ShowAllData
        'Err.Clear
    
    Test_Change = 0
    
    If wb.Sheets("Administrator_Panel").Cells(3, 20).Value <> wb.Sheets("Control").Cells(2, 20).Value Then
        Test_Change = 1
    End If
    If wb.Sheets("Administrator_Panel").Cells(4, 20).Value <> wb.Sheets("Control").Cells(2, 21).Value Then
        Test_Change = 1
    End If
                 
                 
    If Test_Change = 0 Then
        MsgBox "There wasn't found any Updates to be made for Start or End Date for BPF Activity - please made any changes prior to executing script."
        Exit Sub
    End If
    
    Test_Change = 2
    
                
    'Name_User = UserName
     
    Set myRecordset = New ADODB.Recordset
    
    On Error GoTo Errr
        Key_SPS = "BPFA"
        'MsgBox Key_SPS
        Start_Date = wb.Sheets("Administrator_Panel").Cells(2, 20).Value
        End_Date = wb.Sheets("Administrator_Panel").Cells(4, 20).Value
        With myRecordset
           .Open "Select * from Time_Window_Activity Where Record_Number = '" & Key_SPS & "'", _
              strConn, adOpenKeyset, adUseClient
           .Fields("Start_Date").Value = Start_Date
           .Fields("End_Date").Value = End_Date
           .Update
           .Close
        End With
    
    Set myRecordset = Nothing
    Set conn = Nothing

    Text_Mess = "All updates successful added to BOM Leverage Database for."
    
    'MsgBox Test_Change & " updates successful added to BOM Leverage Database."
    MsgBox Text_Mess
    
    Application.ScreenUpdating = True
    Exit Sub
        
Errr:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Global Business Analyst."
    
    myRecordset.Close
    
    '--------------
    
    'close the objects
    conn.Close
    
    'destroy the variables
    Set myRecordset = Nothing
    Set conn = Nothing
        


End Sub
