Attribute VB_Name = "BPFR_Module_PLPL"
Option Explicit

Public SPS_PLPL_Load As Integer
Public CRepIssue As Long
Public CCB_Issue As Long
Public I_Comments As Long
Public Completed_Load As VbMsgBoxResult


Sub BPFR_Load_For_Assigne_connection()
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
    'Dim Completed_Load As String
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    '============================= Sub Bar Script Start =============================
    'Declare Sub Level Variables and Objects
    Dim Counter As Long
    Dim SubCounter As Long
    Dim TotalCount As Long
    
    'SPS_PLPL_Load = 0
    
    
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

    wb.Sheets("Connect_Model_3_Years_Scope").Visible = -1
    
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
    Set ws = wb.Sheets("Connect_Model_3_Years_Scope")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    'strTable = "SalesManager"
    'strTable = "PLPL_View_Connect_SPS_3Scope"
    
    'COPY RECORDS FROM SELECTED FIELDS OF A RECORDSET:
    'refer Image 9e to view records copied to Excel worksheet
    
    'copy all records
    strSQL = "SELECT Key_SPS_Data, Model_Number, Model_Description, Project_Number, Deliverable_Name, DLI_Line_Item, SPS_Owner, Plant_Code, PML_Region, PBU, Product_Family, Detailed_Customer_Name,"
    strSQL = strSQL & "Parent_Customer, Volumes_Y_1, Revenue_Y_1, Volumes_Y_2, Revenue_Y_2, Volumes_Y_3, Revenue_Y_3, Status_SPS, BPF_Source, Rep_Model_Number, BLT_Key, Comments, Date, BPF_Source, Rep_Model_Number, BLT_Key, Comments, GPL FROM PLPL_View_Connect_SPS_3Scope WHERE SPS_AssignedPLPLPLPM ='" & UserName & "'"
    
    Completed_Load = MsgBox("Would You like to load Completed SPS End Model Numbers?", vbYesNo + vbQuestion, "Load Completed?")
    If Completed_Load = vbYes Then
        strSQL = strSQL & "AND Status_SPS <> 'All'"
    Else
        strSQL = strSQL & "AND Status_SPS = 'Assigned'"
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
    MsgBox "There was not found any assigned SPS Urecognized End Model Number to Your name. Please contact with Your PLPM."
    
    wb.Sheets("Connect_Model_3_Years_Scope").Visible = 0
    
    SPS_PLPL_Load = 1
    'MsgBox SPS_PLPL_Load
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
    
    LastRow = wb.Sheets("Connect_Model_3_Years_Scope").Cells(Sheets("Connect_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    
    wb.Sheets("Connect_Model_3_Years_Scope").AutoFilterMode = False
    wb.Sheets("Connect_Model_3_Years_Scope").Range("A5:AC" & LastRow).AutoFilter
    
    wb.Sheets("Connect_Model_3_Years_Scope").AutoFilter.Sort.SortFields.Add Key _
        :=Range("O5:O" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With wb.Sheets("Connect_Model_3_Years_Scope").AutoFilter.Sort
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
            wb.Sheets("Connect_Model_3_Years_Scope").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Revenue"
        Else
            'Integer
            wb.Sheets("Connect_Model_3_Years_Scope").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Quantity"
        End If
    Next n
    
    wb.Sheets("Connect_Model_3_Years_Scope").Range("N6:N" & LastRow).NumberFormat = "#,##0"
    wb.Sheets("Connect_Model_3_Years_Scope").Range("O6:O" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Connect_Model_3_Years_Scope").Range("P6:P" & LastRow).NumberFormat = "#,##0"
    wb.Sheets("Connect_Model_3_Years_Scope").Range("Q6:Q" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Connect_Model_3_Years_Scope").Range("R6:R" & LastRow).NumberFormat = "#,##0"
    wb.Sheets("Connect_Model_3_Years_Scope").Range("S6:S" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Connect_Model_3_Years_Scope").Range("Y6:Y" & LastRow).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    wb.Sheets("Connect_Model_3_Years_Scope").Visible = -1
    wb.Sheets("Connect_Model_3_Years_Scope").Select
    For CRow = 6 To LastRow
        If CRow Mod 2 = 0 Then
            With Range(Cells(CRow, 1), Cells(CRow, 25)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
     '       With Range(Cells(CRow, 1), Cells(CRow, 25)).Font
            With Range(Cells(CRow, 1), Cells(CRow, 30)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        Else
      '      With Range(Cells(CRow, 1), Cells(CRow, 25)).Interior
            With Range(Cells(CRow, 1), Cells(CRow, 30)).Font
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
       '    With Range(Cells(CRow, 1), Cells(CRow, 25)).Font
            With Range(Cells(CRow, 1), Cells(CRow, 30)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next CRow
    
    With wb.Sheets("Connect_Model_3_Years_Scope").Range("U6:U" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$AF$2:$AF$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    LastRow2 = wb.Sheets("Status_Tab").Cells(wb.Sheets("Status_Tab").Rows.Count, "C").End(xlUp).Row
    wb.Sheets("Status_Tab").Cells(LastRow2 + 1, 1).Value = "-"
    
    With wb.Sheets("Connect_Model_3_Years_Scope").Range("W6:W" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Status_Tab'!$A$6:$A$" & LastRow2 + 1
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    For n = 6 To LastRow
        With wb.Sheets("Connect_Model_3_Years_Scope").Cells(n, 21)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=IF($U$" & n & "=""-"", COUNTIF($V$" & n & ":$W$" & n & ", ""-"")<2)"
             .FormatConditions(1).Interior.ColorIndex = 3 'change for other color when ticked
        End With
        With wb.Sheets("Connect_Model_3_Years_Scope").Cells(n, 22)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=$U$" & n & "=""Rep PN"""
             .FormatConditions(1).Interior.ColorIndex = 45 'change for other color when ticked
        End With
        With wb.Sheets("Connect_Model_3_Years_Scope").Cells(n, 23)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=$U$" & n & "=""Costed BOM"""
             .FormatConditions(1).Interior.ColorIndex = 45 'change for other color when ticked
        End With
        With wb.Sheets("Connect_Model_3_Years_Scope").Cells(n, 24)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=$U" & n & "=""Remove"""
             .FormatConditions(1).Interior.ColorIndex = 45 'change for other color when ticked
        End With
    Next n


    Application.ScreenUpdating = True

End Sub


Sub BPFR_Clear_Assing_3Scope()

    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Connect_Model_3_Years_Scope").Cells(Sheets("Connect_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    
    If LastRow > 5 Then
        wb.Sheets("Connect_Model_3_Years_Scope").Range("A6:AC" & LastRow).Clear
        With wb.Sheets("Connect_Model_3_Years_Scope").Range("A6:AC" & LastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If

End Sub


Sub BPFR_PLPL_Active_View()

    Dim wb_name As String
    Dim wb As Object
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    'MsgBox wb_name
    If wb.Sheets("Connect_Model_3_Years_Scope").Visible = 0 Then
        wb.Sheets("Connect_Model_Leverage_Data").Select
    Else
        wb.Sheets("Connect_Model_3_Years_Scope").Select
    End If
    ActiveWindow.NewWindow
    ActiveWindow.Zoom = 80
    Windows.Arrange ArrangeStyle:=xlHorizontal
    Windows("" & wb_name & ":2").Activate
    Sheets("Status_Tab").Select
    
    
    Application.ScreenUpdating = True
    
End Sub


Sub BPFR_PLPL_Active_Clear_View()

    Dim wb_name As String
    Dim wb As Object
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    Application.ScreenUpdating = False

    wb.Sheets("Connect_Model_3_Years_Scope").Select
    Windows("" & wb_name & ":1").Close
'    wb.Sheets("Status_Tab").Select
'    ActiveWindow.Zoom = 80
'    wb.Sheets("Connect_Model_3_Years_Scope").Select
'    ActiveWindow.Zoom = 80
    ActiveWindow.WindowState = xlMaximized
    
    Application.ScreenUpdating = True

End Sub



Sub BPFR_If_Valid_RepPN()

    Dim i As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    
    LastRow = wb.Sheets("Connect_Model_3_Years_Scope").Cells(Sheets("Connect_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    

    CRepIssue = 0

    For i = 6 To LastRow
        If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Text = "Rep PN" Then
            If IsNumeric(wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22)) Then
                If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22) = 0 Then
                wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).FormatConditions.Delete
                wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).Interior.Color = vbRed
                CRepIssue = CRepIssue + 1
                Else
                    If Len(wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22)) = 8 Then
                    Else
                    wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).FormatConditions.Delete
                    wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).Interior.Color = vbRed
                    CRepIssue = CRepIssue + 1
                    End If
                End If
            Else
                If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22) = "-" Then
                    wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).FormatConditions.Delete
                    wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).Interior.Color = vbRed
                    CRepIssue = CRepIssue + 1
                Else
                    If InStr(1, wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22), "-") Then
                    Else
                        If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22) Like "DK*" Then
                        Else
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).FormatConditions.Delete
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).Interior.Color = vbRed
                        CRepIssue = CRepIssue + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = True

End Sub


Sub BPFR_If_Valid_Costed_BOM()

    Dim i As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    
    LastRow = wb.Sheets("Connect_Model_3_Years_Scope").Cells(Sheets("Connect_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    

    CCB_Issue = 0

    For i = 6 To LastRow
        If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Text = "Costed BOM" Then
            If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).Validation.Value = False Then
                wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).FormatConditions.Delete
                wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).Interior.Color = vbRed
                CCB_Issue = CCB_Issue + 1
            Else
                If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).Value = "-" Then
                    wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).FormatConditions.Delete
                    wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).Interior.Color = vbRed
                    CCB_Issue = CCB_Issue + 1
                End If
            End If
        End If
    Next i

End Sub

Sub BPFR_If_Valid_Comments()

    Dim i As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    
    LastRow = wb.Sheets("Connect_Model_3_Years_Scope").Cells(Sheets("Connect_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row
    

    I_Comments = 0

    For i = 6 To LastRow
        If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Text = "Remove" Then
            If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 24).Value = "-" Then
                wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 24).FormatConditions.Delete
                wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 24).Interior.Color = vbRed
                I_Comments = I_Comments + 1
            End If
        End If
    Next i

End Sub


Sub BPFR_Update_Connection()

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
    Dim i As Long
    Dim n As Integer
    Dim Change_Indicator As Long
    
    
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
    
    LastRow = wb.Sheets("Connect_Model_3_Years_Scope").Cells(Sheets("Connect_Model_3_Years_Scope").Rows.Count, "A").End(xlUp).Row

    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
     
    On Error Resume Next
        wb.Sheets("Connect_Model_3_Years_Scope").ShowAllData
        Err.Clear
    
    Test_Change = 0
    
'    For i = 6 To LastRow
'        If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Value <> "-" Then
'            Test_Change = Test_Change + 1
'        End If
'    Next i

    
    For i = 6 To LastRow
    Change_Indicator = 0
        For n = 21 To 23
            If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, n).Value <> wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, n + 5).Value Then
                Change_Indicator = Change_Indicator + 1
            End If
        Next n
         If Change_Indicator > 0 Then
            Test_Change = Test_Change + 1
        End If
    Next i
    
    
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
            If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Value <> "-" Then
            Change_Indicator = 0
                For n = 21 To 23
                    If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, n).Value <> wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, n + 5).Value Then
                        Change_Indicator = Change_Indicator + 1
                    End If
                Next n
                    If Change_Indicator > 0 Then
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 20).Value = "Completed"
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 25).Value = Format(BLT_REF_Macros.GMT, "mm/dd/yyyy hh:mm AM/PM")
                        Key_SPS = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 1).Value
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 26).Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Value
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 27).Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).Value
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 28).Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).Value
                        wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 29).Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 24).Value
                        With myRecordset
                            .Open "Select * from BPF_SPS_Source Where Key_SPS_Data = '" & Key_SPS & "'", _
                              strConn, adOpenKeyset, adUseClient
                            .Fields("BPF_Source").Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Value
                            .Fields("Rep_Model_Number").Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 22).Value
                            .Fields("BLT_Key").Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 23).Value
                            .Fields("Comments").Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 24).Value
                            .Update
                            .Close
                        End With
                        With myRecordset
                            .Open "Select * from BPF_SPS_Data_3Scope Where Key_SPS_Data = '" & Key_SPS & "'", _
                              strConn, adOpenKeyset, adUseClient
                            .Fields("Status_SPS").Value = "Completed"
                            .Fields("SPS_UserProvi").Value = Name_User
                            .Fields("Date").Value = wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 25).Value
                            .Update
                            .Close
                        End With
                    '============================= Sub Bar Script Start =============================
                        SubCounter = SubCounter + 1
                        Subbar.NextAction "Updating BOM Leverage Database with Your connections... Please wait.", True
                    '============================= Sub Bar Script End =============================
                End If
            End If
        Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    Set myRecordset = Nothing
    Set conn = Nothing

    Text_Mess = "Connections successful added to BOM Leverage Database for: [" & Test_Change & "] SPS Line Items."

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
        Subbar.NextAction "Updating BOM Leverage Database with Your connections... Please wait.", True
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



Sub BPFR_Formatting_Connection_Tab_Booked()

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
    
    LastRow = wb.Sheets("Connect_Model_Leverage_Data").Cells(Sheets("Connect_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    
    wb.Sheets("Connect_Model_Leverage_Data").AutoFilterMode = False
    wb.Sheets("Connect_Model_Leverage_Data").Range("A5:AR" & LastRow).AutoFilter
    
    wb.Sheets("Connect_Model_Leverage_Data").AutoFilter.Sort.SortFields.Add Key _
        :=Range("AH5:AH" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With wb.Sheets("Connect_Model_Leverage_Data").AutoFilter.Sort
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
            wb.Sheets("Connect_Model_Leverage_Data").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Revenue"
            wb.Sheets("Connect_Model_Leverage_Data").Range(wb.Sheets("Connect_Model_Leverage_Data").Cells(6, n), wb.Sheets("Connect_Model_Leverage_Data").Cells(LastRow, n)).NumberFormat = "$#,##0"
        Else
            'Integer
            wb.Sheets("Connect_Model_Leverage_Data").Cells(5, n).Value = Start_Year + Int(n / 2) - 7 & " Quantity"
            wb.Sheets("Connect_Model_Leverage_Data").Range(wb.Sheets("Connect_Model_Leverage_Data").Cells(6, n), wb.Sheets("Connect_Model_Leverage_Data").Cells(LastRow, n)).NumberFormat = "#,##0"
        End If
    Next n
    
    wb.Sheets("Connect_Model_Leverage_Data").Range("AH6:AH" & LastRow).NumberFormat = "$#,##0"
    wb.Sheets("Connect_Model_Leverage_Data").Range("AN6:AN" & LastRow).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    
    wb.Sheets("Connect_Model_Leverage_Data").Visible = -1
    wb.Sheets("Connect_Model_Leverage_Data").Select
    For CRow = 6 To LastRow
        If CRow Mod 2 = 0 Then
            'With Range(Cells(CRow, 1), Cells(CRow, 40)).Interior
             With Range(Cells(CRow, 1), Cells(CRow, 45)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
           ' With Range(Cells(CRow, 1), Cells(CRow, 40)).Font
             With Range(Cells(CRow, 1), Cells(CRow, 45)).Interior
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        Else
           ' With Range(Cells(CRow, 1), Cells(CRow, 40)).Interior
            With Range(Cells(CRow, 1), Cells(CRow, 45)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
          '  With Range(Cells(CRow, 1), Cells(CRow, 40)).Font
          With Range(Cells(CRow, 1), Cells(CRow, 45)).Interior
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next CRow

     With wb.Sheets("Connect_Model_Leverage_Data").Range("AJ6:AJ" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Control'!$AF$2:$AF$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    LastRow2 = wb.Sheets("Status_Tab").Cells(wb.Sheets("Status_Tab").Rows.Count, "C").End(xlUp).Row
    wb.Sheets("Status_Tab").Cells(LastRow2 + 1, 1).Value = "-"
    
    With wb.Sheets("Connect_Model_Leverage_Data").Range("AL6:AL" & LastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Status_Tab'!$A$6:$A$" & LastRow2 + 1
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    For n = 6 To LastRow
        With wb.Sheets("Connect_Model_Leverage_Data").Cells(n, 36)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=IF($AJ$" & n & "=""-"", COUNTIF($AK$" & n & ":$AL$" & n & ", ""-"")<2)"
             .FormatConditions(1).Interior.ColorIndex = 3 'change for other color when ticked
        End With
        With wb.Sheets("Connect_Model_Leverage_Data").Cells(n, 37)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=$AJ$" & n & "=""Rep PN"""
             .FormatConditions(1).Interior.ColorIndex = 45 'change for other color when ticked
        End With
        With wb.Sheets("Connect_Model_Leverage_Data").Cells(n, 38)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=$AJ$" & n & "=""Costed BOM"""
             .FormatConditions(1).Interior.ColorIndex = 45 'change for other color when ticked
        End With
        With wb.Sheets("Connect_Model_Leverage_Data").Cells(n, 39)
             .FormatConditions.Delete
             .FormatConditions.Add Type:=xlExpression, _
                        Formula1:="=$AJ" & n & "=""Remove"""
             .FormatConditions(1).Interior.ColorIndex = 45 'change for other color when ticked
        End With
    Next n


    Application.ScreenUpdating = True

End Sub



Sub BPFR_Load_For_Leverage_Connection()
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
    'Dim Completed_Load As String
    
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
    Set ws = wb.Sheets("Connect_Model_Leverage_Data")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    'strTable = "SalesManager"
    'strTable = "PLPL_View_Connect_SPS_3Scope"
    
    'COPY RECORDS FROM SELECTED FIELDS OF A RECORDSET:
    'refer Image 9e to view records copied to Excel worksheet
    
    wb.Sheets("Connect_Model_Leverage_Data").Visible = -1
    
    'copy all records
    strSQL = "SELECT Key_SPS_Data, Model_Number, Model_Description, Project_Number, Deliverable_Name, DLI_Line_Item, SPS_Owner, Plant_Code, PML_Region, PBU, Product_Family, Detailed_Customer_Name,"
    strSQL = strSQL & "Parent_Customer, Volumes_Y_1, Revenue_Y_1, Volumes_Y_2, Revenue_Y_2, Volumes_Y_3, Revenue_Y_3, Volumes_Y_4, Revenue_Y_4, Volumes_Y_5, Revenue_Y_5, Volumes_Y_6, Revenue_Y_6, Volumes_Y_7, Revenue_Y_7, Volumes_Y_8, Revenue_Y_8, Volumes_Y_9, Revenue_Y_9, Volumes_Y_10, Revenue_Y_10,"
    strSQL = strSQL & "Total_Revenue ,Status_SPS, BPF_Source, Rep_Model_Number, BLT_Key, Comments, Date, BPF_Source, Rep_Model_Number, BLT_Key, Comments, GPL FROM PLPL_View_Connect_SPS_Leverage WHERE SPS_AssignedPLPLPLPM ='" & UserName & "'"
    
    'Completed_Load = MsgBox("Would You like to load Completed SPS End Model Numbers?", vbYesNo + vbQuestion, "Load Completed?")
    If Completed_Load = vbYes Then
        strSQL = strSQL & "AND Status_SPS <> 'All'"
    Else
        strSQL = strSQL & "AND Status_SPS = 'Assigned'"
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
    'MsgBox "There was not found any assigned 'SPS Urecognized End Model Number - Leverage Data' - to Your name. Please contact with Your PLPM."
    
    SPS_PLPL_Load = SPS_PLPL_Load + 2
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading List of SPS Data... Please - wait this may take few minutes.", True
    '============================= Sub Bar Script End =============================
    
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    wb.Sheets("Connect_Model_Leverage_Data").Visible = 0
    
    'adoRecSet.Close
    
    '--------------
    'close the objects
    'connDB.Close
    
    'destroy the variables
    Set adoRecSet = Nothing
    Set connDB = Nothing

End Sub



Sub BPFR_Clear_Connect_Leverage()

    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Connect_Model_Leverage_Data").Cells(Sheets("Connect_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    
    If LastRow > 5 Then
        wb.Sheets("Connect_Model_Leverage_Data").Range("A6:AR" & LastRow).Clear
        With wb.Sheets("Connect_Model_Leverage_Data").Range("A6:AR" & LastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If

End Sub





Sub BPFR_If_Valid_RepPN_Leverage()

    Dim i As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    
    LastRow = wb.Sheets("Connect_Model_Leverage_Data").Cells(Sheets("Connect_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    

    CRepIssue = 0

    For i = 6 To LastRow
        If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 36).Text = "Rep PN" Then
            If IsNumeric(wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37)) Then
                If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37) = 0 Then
                wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).FormatConditions.Delete
                wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).Interior.Color = vbRed
                CRepIssue = CRepIssue + 1
                Else
                    If Len(wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37)) = 8 Then
                    Else
                    wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).FormatConditions.Delete
                    wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).Interior.Color = vbRed
                    CRepIssue = CRepIssue + 1
                    End If
                End If
            Else
                If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37) = "-" Then
                    wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).FormatConditions.Delete
                    wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).Interior.Color = vbRed
                    CRepIssue = CRepIssue + 1
                Else
                    If InStr(1, wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37), "-") Then
                    Else
                        If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37) Like "DK*" Then
                        Else
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).FormatConditions.Delete
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).Interior.Color = vbRed
                        CRepIssue = CRepIssue + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

End Sub


Sub BPFR_If_Valid_Costed_BOM_Leverage()

    Dim i As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    
    LastRow = wb.Sheets("Connect_Model_Leverage_Data").Cells(Sheets("Connect_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    

    CCB_Issue = 0

    For i = 6 To LastRow
        If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 36).Text = "Costed BOM" Then
            If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).Validation.Value = False Then
                wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).FormatConditions.Delete
                wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).Interior.Color = vbRed
                CCB_Issue = CCB_Issue + 1
            Else
                If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).Value = "-" Then
                    wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).FormatConditions.Delete
                    wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).Interior.Color = vbRed
                    CCB_Issue = CCB_Issue + 1
                End If
            End If
        End If
    Next i

End Sub

Sub BPFR_If_Valid_Comments_Leverage()

    Dim i As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    
    LastRow = wb.Sheets("Connect_Model_Leverage_Data").Cells(Sheets("Connect_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row
    

    I_Comments = 0

    For i = 6 To LastRow
        If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 36).Text = "Remove" Then
            If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 39).Value = "-" Then
                wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 39).FormatConditions.Delete
                wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 39).Interior.Color = vbRed
                I_Comments = I_Comments + 1
            End If
        End If
    Next i

End Sub



Sub BPFR_Update_Connection_Leverage()

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
    Dim i As Long
    Dim n As Integer
    Dim Change_Indicator As Long
    
    
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
    
    LastRow = wb.Sheets("Connect_Model_Leverage_Data").Cells(Sheets("Connect_Model_Leverage_Data").Rows.Count, "A").End(xlUp).Row

    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB_SPS
     
    On Error Resume Next
        wb.Sheets("Connect_Model_Leverage_Data").ShowAllData
        Err.Clear
    
    Test_Change = 0
    
'    For i = 6 To LastRow
'        If wb.Sheets("Connect_Model_3_Years_Scope").Cells(i, 21).Value <> "-" Then
'            Test_Change = Test_Change + 1
'        End If
'    Next i

    
    For i = 6 To LastRow
    Change_Indicator = 0
        For n = 36 To 39
            If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, n).Value <> wb.Sheets("Connect_Model_Leverage_Data").Cells(i, n + 5).Value Then
                Change_Indicator = Change_Indicator + 1
            End If
        Next n
         If Change_Indicator > 0 Then
            Test_Change = Test_Change + 1
        End If
    Next i
    
    
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
            If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 36).Value <> "-" Then
            Change_Indicator = 0
                For n = 36 To 39
                    If wb.Sheets("Connect_Model_Leverage_Data").Cells(i, n).Value <> wb.Sheets("Connect_Model_Leverage_Data").Cells(i, n + 5).Value Then
                        Change_Indicator = Change_Indicator + 1
                    End If
                Next n
                    If Change_Indicator > 0 Then
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 35).Value = "Completed"
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 40).Value = Format(BLT_REF_Macros.GMT, "mm/dd/yyyy hh:mm AM/PM")
                        Key_SPS = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 1).Value
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 41).Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 36).Value
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 42).Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).Value
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 43).Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).Value
                        wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 44).Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 39).Value
                        With myRecordset
                            .Open "Select * from BPF_SPS_Source Where Key_SPS_Data = '" & Key_SPS & "'", _
                              strConn, adOpenKeyset, adUseClient
                            .Fields("BPF_Source").Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 36).Value
                            .Fields("Rep_Model_Number").Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 37).Value
                            .Fields("BLT_Key").Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 38).Value
                            .Fields("Comments").Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 39).Value
                            .Update
                            .Close
                        End With
                        With myRecordset
                            .Open "Select * from BPF_SPS_Data_Booked Where Key_SPS_Data = '" & Key_SPS & "'", _
                              strConn, adOpenKeyset, adUseClient
                            .Fields("Status_SPS").Value = "Completed"
                            .Fields("SPS_UserProvi").Value = Name_User
                            .Fields("Date").Value = wb.Sheets("Connect_Model_Leverage_Data").Cells(i, 40).Value
                            .Update
                            .Close
                        End With
                    '============================= Sub Bar Script Start =============================
                        SubCounter = SubCounter + 1
                        Subbar.NextAction "Updating BOM Leverage Database with Your connections... Please wait.", True
                    '============================= Sub Bar Script End =============================
                End If
            End If
        Next i
    
    '============================= Sub Bar Script Start =============================
    Subbar.Terminate
    '============================= Sub Bar Script End =============================
    
    Set myRecordset = Nothing
    Set conn = Nothing

    Text_Mess = "Connections successful added to BOM Leverage Database for: [" & Test_Change & "] SPS Line Items."

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
        Subbar.NextAction "Updating BOM Leverage Database with Your connections... Please wait.", True
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
