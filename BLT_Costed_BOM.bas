Attribute VB_Name = "BLT_Costed_BOM"
Option Explicit
Public CB_NL As Integer

Sub Load_For_Assigne()
    'Using ADO to Import data from an Access Database Table to an Excel worksheet (your host application).
    'refer Image 9a to view the existing SalesManager Table in MS Access file "SalesReport.accdb".
    
    'To use ADO in your VBA project, you must add a reference to the ADO Object Library in Excel (your host application) by clicking Tools-References in VBE, and then choose an appropriate version of Microsoft ActiveX Data Objects x.x Library from the list.
    
    '--------------
    'DIM STATEMENTS
    
    Dim strMyPath As String, strSQL As String
    Dim i As Long, n As Long, lFieldCount As Long
    Dim rng As Range
    Dim strTable As String
    
    'instantiate an ADO object using Dim with the New keyword:
    Dim adoRecSet As New ADODB.Recordset
    Dim connDB As New ADODB.Connection
    Dim wb_name As String
    Dim wb As Object
    Dim Region_Global As Integer
    
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
     
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Status_Tab")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'Opening the table named SalesManager:
    'strTable = "SalesManager"
    strTable = "Status_Table"
    
    'COPY RECORDS FROM SELECTED FIELDS OF A RECORDSET:
    'refer Image 9e to view records copied to Excel worksheet
    
    'copy all records
    strSQL = "SELECT Key_BLT, Replaced, Model_Number, Model_Description, Project_Number, DLI_Number, Plant_Code, PBU, Product_Line, Elect_Eng,"
    strSQL = strSQL & "Mech_Eng, Project_Manager, PML_Region, Cost_2016, Cost_2017, Cost_2018, Cost_2019, Cost_2020, Cost_2021,"
    strSQL = strSQL & "Cost_2022, Cost_2023, Cost_2024, Cost_2025, Cost_2026, Cost_2027, Cost_2028, Cost_2029, Cost_2030,"
    strSQL = strSQL & "Cost_2031, Cost_2032, Cost_2033, Cost_2034, Cost_2035, Status, User_Uploaded, Date_of_Modification FROM Status_Table WHERE Status " ' OR 'Archive'" '& UserName & "'"
    
    If Tool_Function.CheckBox3.Value = True And Tool_Function.CheckBox4.Value = True Then
        strSQL = strSQL & " <> 'ALL' "
    Else
        If Tool_Function.CheckBox3.Value = True Then
            strSQL = strSQL & " <> 'Cnacel'"
        Else
            If Tool_Function.CheckBox4.Value = True Then
                strSQL = strSQL & " <> 'Archive'"
            Else
                strSQL = strSQL & " = 'Active'"
            End If
        End If
    End If
    
    Select Case PLP_Value
    
    Case 1
        wb.Sheets("Status_Tab").Made_Rep.Visible = True
        wb.Sheets("Status_Tab").Update_Status.Visible = True
        wb.Sheets("Status_Tab").ReAssign.Visible = False
        strSQL = strSQL & "AND User_Uploaded = '" & UserName & "'"
    Case 2
        wb.Sheets("Status_Tab").Made_Rep.Visible = False
        wb.Sheets("Status_Tab").Update_Status.Visible = False
        wb.Sheets("Status_Tab").ReAssign.Visible = True
        Region_Global = MsgBox("Would You like to load Costed BOM Information - Globally?" & vbNewLine & "[Clicking: 'Yes' = Globally // 'No' = Regional]", vbYesNo + vbQuestion, "Globally or Regionally view")
        If Region_Global = vbNo Then
            If Region_User <> "Global" Then
                strSQL = strSQL & "AND Region = '" & Region_User & "'"
            End If
        End If
        strSQL = strSQL & "AND PBU = '" & PBU_User & "'"
    End Select
    
    
    
        
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
        Subbar.NextAction "Loading User Information...", True
    '============================= Sub Bar Script End =============================
    
    Next i
    
    'select a column range:
    'Range(ws.Columns(1), ws.Columns(lFieldCount)).AutoFit
    'worksheet columns are deleted because this code is only for demo:
    'Range(ws.Columns(1), ws.Columns(lFieldCount)).Delete
    
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
    MsgBox "There was not found any Costed BOM Information for You - please make sure that You uploaded Your Costed BOM Data via ""Costed BOM Tool""."
    
    CB_NL = 1
    
    '============================= Sub Bar Script Start =============================
        SubCounter = lFieldCount
        Subbar.NextAction "Loading User Information...", True
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


Sub Clear_Status_Tab()

    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Status_Tab").Cells(Sheets("Status_Tab").Rows.Count, "A").End(xlUp).Row
    
    If LastRow > 5 Then
        wb.Sheets("Status_Tab").Range("A6:AJ" & LastRow).Clear
        With wb.Sheets("Status_Tab").Range("A6:AJ" & LastRow).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    End If
    

End Sub

Sub Formatting_Status_Tab()

    Dim LastRow As Long
    Dim CRow As Long
    Dim wb_name As String
    Dim wb As Object
    
    Application.ScreenUpdating = False
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Status_Tab").Cells(Sheets("Status_Tab").Rows.Count, "A").End(xlUp).Row
    
    wb.Sheets("Status_Tab").Range("N6:AG" & LastRow).NumberFormat = "#,##0.00"
    wb.Sheets("Status_Tab").Range("AJ6:AJ" & LastRow).NumberFormat = "mm/dd/yyyy hh:mm"
    
    wb.Sheets("Status_Tab").Visible = -1
    wb.Sheets("Status_Tab").Select
    For CRow = 6 To LastRow
        If CRow Mod 2 = 0 Then
            With Range(Cells(CRow, 1), Cells(CRow, 36)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With Range(Cells(CRow, 1), Cells(CRow, 36)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        Else
            With Range(Cells(CRow, 1), Cells(CRow, 36)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With
            With Range(Cells(CRow, 1), Cells(CRow, 36)).Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
    Next CRow

    Application.ScreenUpdating = True

End Sub



Sub BLT_Update_Record()

    Dim conn As ADODB.Connection
    Dim myRecordset As ADODB.Recordset
    Dim strConn As String
    Dim n As Integer
    Dim Test_of_Variance As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim i As Integer

    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = " & strDB
    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakra-fp01\GroupJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
     
    Set myRecordset = New ADODB.Recordset
    
    'Test if something was changed in User Form
    Test_of_Variance = 0
    
    i = ActiveCell.Row
    
    Test_of_Variance = 0
    
    For n = 3 To 35
        If n = 14 Then
        n = n + 20
        End If
        If Status_BLT.Controls("CBT_" & n) <> wb.Sheets("Status_Tab").Cells(i, n).Value Then
            Test_of_Variance = Test_of_Variance + 1
        End If
    Next n
    
    If Test_of_Variance = 0 Then
        MsgBox "You didn't change anything in Costed BOM Status - User Form! If You don't want to implement any changes please click on Cancel Button."
        Exit Sub
    End If
    Status_BLT.Hide
    'End of Testing
    

    On Error GoTo Errr
    With myRecordset
       .Open "Select * FROM BLT_Main_Table Where Key_BLT = '" & Status_BLT.CBT_Key & "'", _
          strConn, adOpenKeyset, adUseClient
         For n = 3 To 35
            If n = 14 Then
                n = n + 20
            End If
                If Status_BLT.Controls("CBT_" & n) <> wb.Sheets("Status_Tab").Cells(i, n).Value Then
                    Select Case n
                        Case 5
                            .Fields("Project_Number").Value = Status_BLT.CBT_5.Value
                        Case 6
                            .Fields("DLI_Number").Value = Status_BLT.CBT_6.Value
                        Case 7
                            .Fields("Plant_Code").Value = Status_BLT.CBT_7.Value
                        Case 8
                            .Fields("PBU").Value = Status_BLT.CBT_8.Value
                        Case 9
                            .Fields("Product_Line").Value = Status_BLT.CBT_9.Value
                        Case 10
                            .Fields("Elect_Eng").Value = Status_BLT.CBT_10.Value
                        Case 11
                            .Fields("Mech_Eng").Value = Status_BLT.CBT_11.Value
                        Case 12
                            .Fields("Project_Manager").Value = Status_BLT.CBT_12.Value
                        Case 34
                            .Fields("Status").Value = Status_BLT.CBT_34.Value
                        Case 35
                            .Fields("User_Uploaded").Value = Status_BLT.CBT_35.Value
                    End Select
                End If
         Next n
       .Update
       .Close
    End With


    Set myRecordset = Nothing
    Set conn = Nothing

    'Change in "Status_Tab" View information
    For n = 3 To 35
        If n = 14 Then
        n = n + 20
        End If
        If Status_BLT.Controls("CBT_" & n) <> wb.Sheets("Status_Tab").Cells(i, n).Value Then
            wb.Sheets("Status_Tab").Cells(i, n).Value = Status_BLT.Controls("CBT_" & n)
        End If
    Next n


    Exit Sub
    
Errr:
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Pagacz, Dominik."
    
    myRecordset.Close
    
    '--------------
    
    'close the objects
    conn.Close
    
    'destroy the variables
    Set myRecordset = Nothing
    Set conn = Nothing
    
End Sub


Sub BLT_UN_Protect_Sheet()

    Dim wb_name As String
    Dim wb As Object

    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)

    wb.Sheets("Status_Tab").Unprotect Password:="Lockthisup"

End Sub

Sub BLT_Protect_Sheet()

    Dim wb_name As String
    Dim wb As Object

    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)

    wb.Sheets("Status_Tab").Protect Password:="Lockthisup", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    wb.Sheets("Status_Tab").EnableOutlining = True
    wb.Sheets("Status_Tab").EnableSelection = xlNoRestrictions

End Sub


Sub BLT_Active_List_for_Replace()

    Dim wb_name As String
    Dim wb As Object
    Dim i As Long
    Dim section As String
    Dim n As Integer
    Dim LastRow As Long
    Dim ti As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    Application.ScreenUpdating = False
    
    wb.Sheets("Temp_Replace_Tab").Visible = -1
    wb.Sheets("Temp_Replace_Tab").Select
    
    LastRow = wb.Sheets("Status_Tab").Cells(Sheets("Status_Tab").Rows.Count, "A").End(xlUp).Row
    
    ti = 2
    For i = 6 To LastRow
        If wb.Sheets("Status_Tab").Cells(i, 34).Value = "Active" Then
            For n = 1 To 13
                wb.Sheets("Temp_Replace_Tab").Cells(ti, n).Value = wb.Sheets("Status_Tab").Cells(i, n).Value
            Next n
            wb.Sheets("Temp_Replace_Tab").Cells(ti, 14).Value = Format(wb.Sheets("Status_Tab").Cells(i, 36).Value, "mm/dd/yyyy hh:mm")
            ti = ti + 1
        End If
    Next i
    
    wb.Sheets("Status_Tab").Select
    wb.Sheets("Temp_Replace_Tab").Visible = 2
    
    Application.ScreenUpdating = True

End Sub


Sub BLT_Clear_Temp_Replace()

    Dim wb_name As String
    Dim wb As Object
    Dim LastRow As Long
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Temp_Replace_Tab").Cells(Sheets("Temp_Replace_Tab").Rows.Count, "A").End(xlUp).Row
    
    If LastRow > 1 Then
        wb.Sheets("Temp_Replace_Tab").Range("A2:N" & LastRow).Clear
    End If
 

End Sub




Sub BLT_Replace_Update()

    
    Dim conn As ADODB.Connection
    Dim myRecordset As ADODB.Recordset
    Dim strConn As String
    Dim n As Integer
    Dim Test_of_Variance As Integer
    Dim wb_name As String
    Dim wb As Object
    Dim i As Integer
    Dim New_EMN, Rep_EMN As String
    Dim LastRow As Long

    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)

    i = ActiveCell.Row
    New_EMN = wb.Sheets("Status_Tab").Cells(i, 1).Value
    Rep_EMN = BLT_Replace_BOM.ListBox1.Value

'    strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" & _
'     "Data Source = " & strDB
    'strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
    strConn = "Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
    Set myRecordset = New ADODB.Recordset

    On Error GoTo Errr
    With myRecordset
        .Open "Select * FROM BLT_Main_Table Where Key_BLT = '" & Rep_EMN & "'", _
            strConn, adOpenKeyset, adUseClient
        .Fields("Status").Value = "Archive"
        .Update
        .Close
    End With
    
    With myRecordset
        .Open "Select * FROM BLT_Main_Table Where Key_BLT = '" & New_EMN & "'", _
            strConn, adOpenKeyset, adUseClient
        .Fields("Replaced").Value = Rep_EMN
        .Update
        .Close
    End With


    Set myRecordset = Nothing
    Set conn = Nothing

    'Change in "Status_Tab" View information
    
    wb.Sheets("Status_Tab").Cells(i, 2).Value = Rep_EMN
    
    LastRow = wb.Sheets("Status_Tab").Cells(Sheets("Status_Tab").Rows.Count, "A").End(xlUp).Row
    
    For i = 6 To LastRow
        If Rep_EMN = wb.Sheets("Status_Tab").Cells(i, 1).Value Then
            wb.Sheets("Status_Tab").Cells(i, 34).Value = "Archive"
            i = LastRow
        End If
    Next i
    
    Exit Sub
    
Errr:
    MsgBox "Something went wrong - please wait for 5 minutes and if the Error will occure again contact with Pagacz, Dominik."
    
    myRecordset.Close
    
    '--------------
    
    'close the objects
    conn.Close
    
    'destroy the variables
    Set myRecordset = Nothing
    Set conn = Nothing
    

End Sub
