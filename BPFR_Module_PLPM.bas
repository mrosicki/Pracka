Attribute VB_Name = "BPFR_Module_PLPM"
Global BPFR_OoT As Boolean
Public Region_Global As String

Sub BPFR_Load_For_Assigne()
    'Using ADO to Import data from an Access Database Table to an Excel worksheet (your host application).
    'refer Image 9a to view the existing SalesManager Table in MS Access file "SalesReport.accdb".
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, lFieldCount As Long
    Dim rng As Range
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    Dim wb_name As String
    Dim wb As Object
    Dim Region_selected As String
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
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
    
    '--------------
    'THE CONNECTION OBJECT
    
    'strDBName = "BOM Leverage Database.accdb"
    'strMyPath = ThisWorkbook.Path
    'MsgBox strMyPath
    'strDB = "S:\" & strDBName
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Assign_Model_3_Years_Scope")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    'strTable = "SalesManager"
    strTable = "PLPM_View_Assign_3Scope"
    
    'COPY RECORDS FROM SELECTED FIELDS OF A RECORDSET:
    'refer Image 9e to view records copied to Excel worksheet
    
    'copy all records
    strSQL = "SELECT Key_SPS_Data, Model_Number, Model_Description, Project_Number, Deliverable_Name, DLI_Line_Item, SPS_Owner, Plant_Code, PML_Region, PBU, Product_Family, Detailed_Customer_Name,"
    strSQL = strSQL & "Parent_Customer, Volumes_Y_1, Revenue_Y_1, Volumes_Y_2, Revenue_Y_2, Volumes_Y_3, Revenue_Y_3, Status_SPS, SPS_AssignedPLPLPLPM, Date, SPS_AssignedPLPLPLPM FROM PLPM_View_Assign_3Scope WHERE PBU ='" & PBU_User & "'"
    
    Region_Global = MsgBox("Would You like to load SPS Unrecognized End Model Numbers Information - Globally?" & vbNewLine & "[Clicking: 'Yes' = Globally // 'No' = Your former Region]", vbYesNo + vbQuestion, "Globally or Regionally view")
    If Region_Global = vbNo Then
        If Region_User <> "Global" Then
            strSQL = strSQL & "AND PML_Region = '" & Region_User & "'"
            Region_selected = Region_User
        End If
    Else
        Region_selected = "Global"
    End If
    
    
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("A6")
    lFieldCount = adoRecSet.Fields.Count
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = lFieldCount
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
    
    
    
    For i = 0 To lFieldCount - 1
    'copy column names in first row of the worksheet:
    'rng.Offset(0, i).Value = adoRecSet.Fields(i).Name
    
    On Error GoTo Errr
    adoRecSet.MoveFirst
    
    'copy record values starting from second row of the worksheet:
    n = 0
    Do While Not adoRecSet.EOF
    rng.Offset(n, i).Value = adoRecSet.Fields(i).Value
    adoRecSet.MoveNext
    n = n + 1
    Loop
    
    '============================= Sub Bar Script Start =============================
        SubCounter = SubCounter + 1
        Subbar.NextAction "Loading List of SPS Data... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing
    
    Exit Sub

Errr:
    MsgBox "There was not found any value for Your PBU: " & PBU_User & " and Region: " & Region_selected & "."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading List of SPS Data... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing

End Sub

Sub BPFR_Formatting_Assign_Tab()

    Dim LastRow As Long
    Dim CRow As Long
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow2 As Long
    Dim Start_Year As Integer
    Dim n As Integer
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Assign_Model_3_Years_Scope").Cells(Sheets("Assign_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    
    wb.Sheets("Assign_Model_3_Years_Scope").AutoFilterMode = False
    wb.Sheets("Assign_Model_3_Years_Scope").Range("A5:W" & LastRow).AutoFilter
    
    wb.Sheets("Assign_Model_3_Years_Scope").AutoFilter.Sort.SortFields.Add Key _
        :=Range("O5:O" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With wb.Sheets("Assign_Model_3_Years_Scope").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Start_Year = wb.Sheets("Control").Cells(2, 18).Value
    
    For n = 14 To 19
        If n / 2 - Int(n / 2) <> 0 Then
            'not integer
            wb.Sheets("Assign_Model_3_Years_Scope").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Revenue"
        Else
            'Integer
            wb.Sheets("Assign_Model_3_Years_Scope").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Quantity"
        End If
    Next n
    
    wb.Sheets("Assign_Model_3_Years_Scope").Range("N6:N" & LastRow).NumberFormat = "#,##0"
    wb.Sheets("Assign_Model_3_Years_Scope").Range("O6:O" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Assign_Model_3_Years_Scope").Range("P6:P" & LastRow).NumberFormat = "#,##0"
    wb.Sheets("Assign_Model_3_Years_Scope").Range("Q6:Q" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Assign_Model_3_Years_Scope").Range("R6:R" & LastRow).NumberFormat = "#,##0"
    wb.Sheets("Assign_Model_3_Years_Scope").Range("S6:S" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Assign_Model_3_Years_Scope").Range("V6:V" & LastRow).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    wb.Sheets("Assign_Model_3_Years_Scope").Visible = -1
    wb.Sheets("Assign_Model_3_Years_Scope").Select
    For CRow = 6 To LastRow
        If CRow Mod 2 = 0 Then
            With Range(Cells(CRow, 1), Cells(CRow, 22)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With Range(Cells(CRow, 1), Cells(CRow, 22)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        Else
            With Range(Cells(CRow, 1), Cells(CRow, 22)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
            With Range(Cells(CRow, 1), Cells(CRow, 22)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next CRow
    
    LastRow2 = wb.Sheets("Control").Cells(Sheets("Control").Rows.Count, "P").End(xlUp).Row
    
    With wb.Sheets("Assign_Model_3_Years_Scope").Range("U6:U" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$P$2:$P$" & LastRow2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    ActiveWindow.Zoom = 80

    Application.ScreenUpdating = True

End Sub


Sub BPFR_Clear_Assing_3Scope()

    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Assign_Model_3_Years_Scope").Cells(Sheets("Assign_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    
    If LastRow > 5 Then
        wb.Sheets("Assign_Model_3_Years_Scope").Range("A6:W" & LastRow).Clear
        With wb.Sheets("Assign_Model_3_Years_Scope").Range("A6:W" & LastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If

End Sub



Sub BPFR_Update_Assignment()

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
    
    LastRow = wb.Sheets("Assign_Model_3_Years_Scope").Cells(Sheets("Assign_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = S:\BOM Leverage Database.accdb"
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
     "Data Source = " & strDB
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;"
    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
     
    On Error Resume Next
        wb.Sheets("Assign_Model_3_Years_Scope").ShowAllData
        Err.Clear
    
    Test_Change = 0
    For i = 6 To LastRow
        If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 20) <> "Completed" Then
            If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21) <> "TBD" Then
                If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Validation.Value = True Then
                    If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Value <> wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 23).Value Then
                        Test_Change = Test_Change + 1
                    End If
                End If
            End If
        End If
    Next i
                 
    If Test_Change = 0 Then
        MsgBox "You didn't enter any changes in section: 'PLPL or PLPM to Assign' - please make any updates before clicking on 'Update Assignments' Button."
        Exit Sub
    End If
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = Test_Change
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
                
    Name_User = UserName
     
    Set myRecordset = New ADODB.Recordset
    
    On Error GoTo Errr
    For i = 6 To LastRow
        If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 20) <> "Completed" Then
            If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21) <> "TBD" Then
                If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Validation.Value = True Then
                    If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Value <> wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 23).Value Then
                        wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 20).Value = "Assigned"
                        wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 23).Value = wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Value
                        wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 22).Value = Format(BLT_REF_Macros.GMT, "mm/dd/yyyy hh:mm AM/PM")
                        Key_SPS = wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 1).Value
                        With myRecordset
                           .Open "Select * from BPF_SPS_Data_3Scope Where Key_SPS_Data = '" & Key_SPS & "'", _
                              strConn, adOpenKeyset, adUseClient
                           .Fields("Status_SPS").Value = "Assigned"
                           .Fields("SPS_AssignedPLPLPLPM").Value = wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Value
                           .Fields("SPS_UserAssi").Value = Name_User
                           .Fields("Date").Value = wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 22).Value
                           .Update
                           .Close
                        End With
                        '============================= Sub Bar Script Start =============================
                            SubCounter = SubCounter + 1
                            Subbar.NextAction "Updating Assignments in BOM Leverage Database... Please wait.", True
                        '============================= Sub Bar Script End =============================
                    End If
                Else
                    With wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
             End If
            If wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 21).Value <> wb.Sheets("Assign_Model_3_Years_Scope").Cells(i, 23).Value Then
                Test_Complete = Test_Complete + 1
            End If
        End If
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    Set myRecordset = Nothing
    Set conn = Nothing

    Text_Mess = "Updates successful added to BOM Leverage Database for: [" & Test_Change & "] SPS Line Items."

    If Test_Complete > 0 Then
        Text_Mess = Text_Mess & vbNewLine & vbNewLine & "Records not updated with reason:" & vbNewLine & "Change of assignment for 'Completed' SPS Line is not allowed: [" & Test_Complete & "]"
    End If
    
    'MsgBox Test_Change & " updates successful added to BOM Leverage Database."
    MsgBox Text_Mess
    
    Application.ScreenUpdating = True
    Exit Sub
        
Errr:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Pagacz, Dominik."
    
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


Sub BPFR_Create_PBU_List_User()

    Dim i, n As Long
    Dim LastRow As Long
    Dim wb_name As String
    Dim wb As Object
    
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)


    LastRow = wb.Sheets("Control").Cells(Sheets("Control").Rows.Count, "A").End(xlUp).Row
    
    
    n = 2
    For i = 2 To LastRow
        If wb.Sheets("Control").Cells(i, 4).Value = PBU_User Then
            wb.Sheets("Control").Cells(n, 16).Value = wb.Sheets("Control").Cells(i, 1).Value
            n = n + 1
        End If
    Next i

End Sub


Sub BPFR_Load_Start_Year_and_Version()

    'Using ADO to Import data from an Access Database Table to an Excel worksheet (your host application).
    'refer Image 9a to view the existing SalesManager Table in MS Access file "SalesReport.accdb".
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, lFieldCount As Long
    Dim rng As Range
    Dim wb_name As String
    Dim wb As Object
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    
    '--------------
    'THE CONNECTION OBJECT
    
    
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
    
    
    wb.Sheets("Control").Range("R2:R3").Clear
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Control")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'copy all records
    'strSQL = "SELECT User_Name, User_Groupe, User_Domain, User_PBU FROM User_Access"
    strSQL = "SELECT Start_Year FROM Start_Year"
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("R2")
    lFieldCount = adoRecSet.Fields.Count
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = lFieldCount
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
    
    For i = 0 To lFieldCount - 1
    'copy column names in first row of the worksheet:
    'rng.Offset(0, i).Value = adoRecSet.Fields(i).Name
    On Error GoTo Errr
    adoRecSet.MoveFirst
    
    'copy record values starting from second row of the worksheet:
    n = 0
    Do While Not adoRecSet.EOF
    rng.Offset(n, i).Value = adoRecSet.Fields(i).Value
    adoRecSet.MoveNext
    n = n + 1
    Loop
    
    '============================= Sub Bar Script Start =============================
        SubCounter = SubCounter + 1
        Subbar.NextAction "Loading Background Information...", True
    '============================= Sub Bar Script End =============================
    
    Next i

    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing
    Exit Sub

Errr:
    MsgBox "There wasn't found any table named 'Year & Version' - please contact with Administrator."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading Background Information...", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing

End Sub


Sub BPFR_Time_Window()

    'Using ADO to Import data from an Access Database Table to an Excel worksheet (your host application).
    'refer Image 9a to view the existing SalesManager Table in MS Access file "SalesReport.accdb".
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, lFieldCount As Long
    Dim rng As Range
    Dim wb_name As String
    Dim wb As Object
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    
    '--------------
    'THE CONNECTION OBJECT
    
    
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
    
    
    wb.Sheets("Control").Range("T2:U2").Clear
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Control")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'copy all records
    'strSQL = "SELECT User_Name, User_Groupe, User_Domain, User_PBU FROM User_Access"
    strSQL = "SELECT Start_Date, End_Date FROM Time_Window_Activity"
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("T2")
    lFieldCount = adoRecSet.Fields.Count
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = lFieldCount
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
    
    For i = 0 To lFieldCount - 1
    'copy column names in first row of the worksheet:
    'rng.Offset(0, i).Value = adoRecSet.Fields(i).Name
    On Error GoTo Errr
    adoRecSet.MoveFirst
    
    'copy record values starting from second row of the worksheet:
    n = 0
    Do While Not adoRecSet.EOF
    rng.Offset(n, i).Value = adoRecSet.Fields(i).Value
    adoRecSet.MoveNext
    n = n + 1
    Loop
    
    '============================= Sub Bar Script Start =============================
        SubCounter = SubCounter + 1
        Subbar.NextAction "Loading Time Window Information...", True
    '============================= Sub Bar Script End =============================
    
    Next i

    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing
    Exit Sub
Errr:
    MsgBox "There wasn't found any table named 'Start Year & End Year' - please contact with Administrator."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading Time Window Information...", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing


End Sub



Sub Test_If_correct_Date()

    Dim wb_name As String
    Dim wb As Object
    Dim Cur_Date As Double
    Dim Compare_Date As Double
    
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)

    BPFR_OoT = 0

    Cur_Date = Now
    Compare_Date = wb.Sheets("Control").Cells(2, 20).Value


    If Compare_Date > Cur_Date Then
        BPFR_OoT = 1
        Else
        Compare_Date = wb.Sheets("Control").Cells(2, 21).Value
        If Compare_Date < Cur_Date Then
            BPFR_OoT = 1
        End If
    End If

    'MsgBox BPFR_OoT

End Sub




Sub BPFR_Load_For_Assigne_Booked()
    'Using ADO to Import data from an Access Database Table to an Excel worksheet (your host application).
    'refer Image 9a to view the existing SalesManager Table in MS Access file "SalesReport.accdb".
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, lFieldCount As Long
    Dim rng As Range
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    Dim wb_name As String
    Dim wb As Object
    Dim Region_selected As String
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
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
    
    '--------------
    'THE CONNECTION OBJECT
    
    'strDBName = "BOM Leverage Database.accdb"
    'strMyPath = ThisWorkbook.Path
    'MsgBox strMyPath
    'strDB = "S:\" & strDBName
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Assign_Model_Leverage_Data")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    'strTable = "SalesManager"
    strTable = "PLPM_View_Assign_Booked"
    
    'COPY RECORDS FROM SELECTED FIELDS OF A RECORDSET:
    'refer Image 9e to view records copied to Excel worksheet
    
    'copy all records
    strSQL = "SELECT Key_SPS_Data, Model_Number, Model_Description, Project_Number, Deliverable_Name, DLI_Line_Item, SPS_Owner, Plant_Code, PML_Region, PBU, Product_Family, Detailed_Customer_Name,"
    strSQL = strSQL & "Parent_Customer, Volumes_Y_1, Revenue_Y_1, Volumes_Y_2, Revenue_Y_2, Volumes_Y_3, Revenue_Y_3, Volumes_Y_4, Revenue_Y_4, Volumes_Y_5, Revenue_Y_5, Volumes_Y_6, Revenue_Y_6, Volumes_Y_7, Revenue_Y_7, Volumes_Y_8, Revenue_Y_8, Volumes_Y_9, Revenue_Y_9, Volumes_Y_10, Revenue_Y_10,"
    strSQL = strSQL & "Total_Revenue ,Status_SPS, SPS_AssignedPLPLPLPM, Date, SPS_AssignedPLPLPLPM FROM PLPM_View_Assign_Booked WHERE PBU ='" & PBU_User & "'"
    
    'Region_Global = MsgBox("Would You like to load SPS Unrecognized End Model Numbers Information - Globally?" & vbNewLine & "[Clicking: 'Yes' = Globally // 'No' = Your former Region]", vbYesNo + vbQuestion, "Globally or Regionally view")
    If Region_Global = vbNo Then
        If Region_User <> "Global" Then
            strSQL = strSQL & "AND PML_Region = '" & Region_User & "'"
            Region_selected = Region_User
        End If
    Else
        Region_selected = "Global"
    End If
    
    
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("A6")
    lFieldCount = adoRecSet.Fields.Count
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = lFieldCount
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
    
    
    
    For i = 0 To lFieldCount - 1
    'copy column names in first row of the worksheet:
    'rng.Offset(0, i).Value = adoRecSet.Fields(i).Name
    
    On Error GoTo Errr
    adoRecSet.MoveFirst
    
    'copy record values starting from second row of the worksheet:
    n = 0
    Do While Not adoRecSet.EOF
    rng.Offset(n, i).Value = adoRecSet.Fields(i).Value
    adoRecSet.MoveNext
    n = n + 1
    Loop
    
    '============================= Sub Bar Script Start =============================
        SubCounter = SubCounter + 1
        Subbar.NextAction "Loading List of SPS Data... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing
    
    Exit Sub

Errr:
    MsgBox "There was not found any value for Your PBU: " & PBU_User & " and Region: " & Region_selected & "."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading List of SPS Data... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing

End Sub


Sub BPFR_Clear_Assing_Booked()

    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Assign_Model_Leverage_Data").Cells(Sheets("Assign_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    
    If LastRow > 5 Then
        wb.Sheets("Assign_Model_Leverage_Data").Range("A6:AN" & LastRow).Clear
        With wb.Sheets("Assign_Model_Leverage_Data").Range("A6:AN" & LastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If

End Sub


Sub BPFR_Formatting_Assign_Tab_Booked()

    Dim LastRow As Long
    Dim CRow As Long
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow2 As Long
    Dim Start_Year As Integer
    Dim n As Integer
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Assign_Model_Leverage_Data").Cells(Sheets("Assign_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    
    wb.Sheets("Assign_Model_Leverage_Data").AutoFilterMode = False
    wb.Sheets("Assign_Model_Leverage_Data").Range("A5:AL" & LastRow).AutoFilter
    
    wb.Sheets("Assign_Model_Leverage_Data").AutoFilter.Sort.SortFields.Add Key _
        :=Range("AH5:AH" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With wb.Sheets("Assign_Model_Leverage_Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Start_Year = wb.Sheets("Control").Cells(2, 18).Value
    
    For n = 14 To 33
        If n / 2 - Int(n / 2) <> 0 Then
            'not integer
            wb.Sheets("Assign_Model_Leverage_Data").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Revenue"
            wb.Sheets("Assign_Model_Leverage_Data").Range(wb.Sheets("Assign_Model_Leverage_Data").Cells(6, n), wb.Sheets("Assign_Model_Leverage_Data").Cells(LastRow, n)).NumberFormat = "$#,##0"
        Else
            'Integer
            wb.Sheets("Assign_Model_Leverage_Data").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Quantity"
            wb.Sheets("Assign_Model_Leverage_Data").Range(wb.Sheets("Assign_Model_Leverage_Data").Cells(6, n), wb.Sheets("Assign_Model_Leverage_Data").Cells(LastRow, n)).NumberFormat = "#,##0"
        End If
    Next n
    
    wb.Sheets("Assign_Model_Leverage_Data").Range("AH6:AH" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Assign_Model_Leverage_Data").Range("AK6:AK" & LastRow).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    wb.Sheets("Assign_Model_Leverage_Data").Visible = -1
    wb.Sheets("Assign_Model_Leverage_Data").Select
    For CRow = 6 To LastRow
        If CRow Mod 2 = 0 Then
            With Range(Cells(CRow, 1), Cells(CRow, 37)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With Range(Cells(CRow, 1), Cells(CRow, 37)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        Else
            With Range(Cells(CRow, 1), Cells(CRow, 37)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
            With Range(Cells(CRow, 1), Cells(CRow, 37)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next CRow

    LastRow2 = wb.Sheets("Control").Cells(Sheets("Control").Rows.Count, "P").End(xlUp).Row
    
    With wb.Sheets("Assign_Model_Leverage_Data").Range("AJ6:AJ" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$P$2:$P$" & LastRow2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    ActiveWindow.Zoom = 80

    Application.ScreenUpdating = True

End Sub



Sub BPFR_Update_Assignment_Booked()

    Dim conn As ADODB.Connection
    Dim myRecordset As ADODB.Recordset
    Dim strConn As String
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    Dim Name_User As String
    Dim Key_SPS As String
    Dim Test_Change As Integer
    
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
    
    LastRow = wb.Sheets("Assign_Model_Leverage_Data").Cells(Sheets("Assign_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = S:\BOM Leverage Database.accdb"
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
     "Data Source = " & strDB
    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
     
    On Error Resume Next
        wb.Sheets("Assign_Model_Leverage_Data").ShowAllData
        Err.Clear
    
    Test_Change = 0
    For i = 6 To LastRow
        If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 35) <> "Completed" Then
            If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36) <> "TBD" Then
                If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Validation.Value = True Then
                    If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Value <> wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 38).Value Then
                        Test_Change = Test_Change + 1
                    End If
                End If
            End If
        End If
    Next i
                 
    If Test_Change = 0 Then
        MsgBox "You didn't enter any changes in section: 'PLPL or PLPM to Assign' - please make any updates before clicking on 'Update Assignments' Button."
        Exit Sub
    End If
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = Test_Change
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
                
    Name_User = UserName
     
    Set myRecordset = New ADODB.Recordset
    
    On Error GoTo Errr
    For i = 6 To LastRow
        If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 35) <> "Completed" Then
            If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36) <> "TBD" Then
                If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Validation.Value = True Then
                    If wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Value <> wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 38).Value Then
                        wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 35).Value = "Assigned"
                        wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 38).Value = wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Value
                        wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 37).Value = Format(BLT_REF_Macros.GMT, "mm/dd/yyyy hh:mm AM/PM")
                        Key_SPS = wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 1).Value
                        With myRecordset
                           .Open "Select * from BPF_SPS_Data_Booked Where Key_SPS_Data = '" & Key_SPS & "'", _
                              strConn, adOpenKeyset, adUseClient
                           .Fields("Status_SPS").Value = "Assigned"
                           .Fields("SPS_AssignedPLPLPLPM").Value = wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Value
                           .Fields("SPS_UserAssi").Value = Name_User
                           .Fields("Date").Value = wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 37).Value
                           .Update
                           .Close
                        End With
                        '============================= Sub Bar Script Start =============================
                            SubCounter = SubCounter + 1
                            Subbar.NextAction "Updating Assignments in BOM Leverage Database... Please wait.", True
                        '============================= Sub Bar Script End =============================
                    End If
                Else
                    With wb.Sheets("Assign_Model_Leverage_Data").Cells(i, 36).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
             End If
        End If
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    Set myRecordset = Nothing
    Set conn = Nothing

    MsgBox Test_Change & " updates successful added to BOM Leverage Database."
    
    Application.ScreenUpdating = True
    Exit Sub
        
Errr:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Pagacz, Dominik."
    
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


Sub BPFR_Load_Targets()
    'Using ADO to Import data from an Access Database Table to an Excel worksheet (your host application).
    'refer Image 9a to view the existing SalesManager Table in MS Access file "SalesReport.accdb".
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, lFieldCount As Long
    Dim rng As Range
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    Dim wb_name As String
    Dim wb As Object
    Dim Region_selected As String
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
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
    
    '--------------
    'THE CONNECTION OBJECT
    
    'strDBName = "BOM Leverage Database.accdb"
    'strMyPath = ThisWorkbook.Path
    'MsgBox strMyPath
    'strDB = "S:\" & strDBName
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Control")
    
    On Error GoTo Errr
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    'strTable = "SalesManager"
    strTable = "SPS_Coverage_Targets"
    
    'COPY RECORDS FROM SELECTED FIELDS OF A RECORDSET:
    'refer Image 9e to view records copied to Excel worksheet
    
    'copy all records
    strSQL = "SELECT PBU, Region, Target_Year_1, Target_Year_2, Target_Year_3, Leverage FROM SPS_Coverage_Targets WHERE PBU ='" & PBU_User & "'"
    
    'Region_Global = MsgBox("Would You like to load SPS Unrecognized End Model Numbers Information - Globally?" & vbNewLine & "[Clicking: 'Yes' = Globally // 'No' = Your former Region]", vbYesNo + vbQuestion, "Globally or Regionally view")
    If Region_Global = vbNo Then
        strSQL = strSQL & "AND Region = '" & Region_User & "'"
        Region_selected = Region_User
    Else
        strSQL = strSQL & "AND Region = 'Global'"
        Region_selected = "Global"
    End If
    
    
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("X2")
    lFieldCount = adoRecSet.Fields.Count
    
    '============================= Sub Bar Script Start =============================

    'Set the total actions property
    Subbar.TotalActions = lFieldCount
    'Show the Sub bar
    Subbar.ShowBar
    'Move the Second bar below the main Bar
    Subbar.Top = Subbar.Top + Subbar.Height + 10

    '============================= Sub Bar Script End =============================
    
    
    
    For i = 0 To lFieldCount - 1
    'copy column names in first row of the worksheet:
    'rng.Offset(0, i).Value = adoRecSet.Fields(i).Name
    
    adoRecSet.MoveFirst
    
    'copy record values starting from second row of the worksheet:
    n = 0
    Do While Not adoRecSet.EOF
    rng.Offset(n, i).Value = adoRecSet.Fields(i).Value
    adoRecSet.MoveNext
    n = n + 1
    Loop
    
    '============================= Sub Bar Script Start =============================
        SubCounter = SubCounter + 1
        Subbar.NextAction "Loading Targets for Your PBU... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing
    
    Exit Sub

Errr:
    MsgBox "There was not found any value for Your PBU: " & PBU_User & " and Region: " & Region_selected & "."
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading List of SPS Data... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    
    adoRecSet.Close
    
    '--------------
    
    'close the objects
    connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing

End Sub




