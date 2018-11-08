Attribute VB_Name = "BLT_REF_Macros"
Option Explicit

Public Standard_Time As Date
Dim TInfo As CTime
Global strDB As String
Global strDB_SPS As String
Global AccessTest As Boolean
Public PLP_Value As Integer
Public PBU_User As String
Public Region_User As String
Public Posit_User As String
Public Tool_Version As Integer

Sub Link_to_BOM_Leverage_Database()

    'strDB_SPS = "\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BPFR SPS Database.accdb;"
     strDB_SPS = "\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BPFR SPS Database.accdb;"
    'strDB_SPS = "C:\BPFR_Test_Enviroment\BPFR SPS Database.accdb;"


End Sub


Sub BLT_Access_Test()

AccessTest = True

'  If FolderExists("\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\") Then
    'MsgBox "You have access to the Living Leverage Database."
'    strDB = "\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;"
  If FolderExists("\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\") Then
    'MsgBox "You have access to the Living Leverage Database."
    strDB = "\\plmakrk-fp01\GROUPJ03$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;"
     Else
    AccessTest = False
  End If
End Sub

Private Function FolderExists(ByVal Path As String) As Boolean
  On Error Resume Next
  FolderExists = Dir(Path, vbDirectory) <> ""
End Function


Public Function GMT() As Double
    If TInfo Is Nothing Then
        Set TInfo = New CTime
    Else
        TInfo.Refresh
    End If
    GMT = TInfo.GMT
End Function




Sub TIME_ZONE_FOR_ME()
    
    Standard_Time = GMT
    
    MsgBox Format(Standard_Time, "m/dd/yyyy hh:mm AM/PM")

End Sub



Sub Load_List_of_Users()

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
    
    
    wb.Sheets("Control").Range("A2:D100000").Clear
    wb.Sheets("Control").Range("P2:P100000").Clear
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Control")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'copy all records
    'strSQL = "SELECT User_Name, User_Groupe, User_Domain, User_PBU FROM User_Access"
    strSQL = "SELECT User_Name, User_Groupe, User_Domain, User_PBU FROM User_Access"
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("A2")
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
        Subbar.NextAction "Loading User Information...", True
    '============================= Sub Bar Script End =============================
    
    Next i
    
    Application.ScreenUpdating = False
    
    wb.Sheets("Control").Sort.SortFields.Clear
    wb.Sheets("Control").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wb.Sheets("Control").Sort
        .SetRange Range("A2:D2500")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = True
    
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


Function UserName()
    Dim objInfo
    Dim strLDAP
    Dim strFullName
    
    Set objInfo = CreateObject("ADSystemInfo")
    strLDAP = objInfo.UserName
    Set objInfo = Nothing
    strFullName = GetUserName(strLDAP)
    
    UserName = strFullName

End Function


Function GetUserName(strLDAP)
  Dim objUser
  Dim strName
  Dim arrLDAP
  Dim intIdx
  
  On Error Resume Next
  strName = ""
  Set objUser = GetObject("LDAP://" & strLDAP)
  If Err.Number = 0 Then
  '  strName = objUser.Get("givenName") & Chr(32) & objUser.Get("sn")
    strName = Trim(objUser.Get("sn") & Chr(44) & Chr(32) & objUser.Get("givenName"))
    
'    MsgBox "Get sn: " & objUser.Get("sn")
'    MsgBox "Get givenName: " & objUser.Get("givenName")


    
  End If
  If Err.Number <> 0 Then
    arrLDAP = Split(strLDAP, ",")
    For intIdx = 0 To UBound(arrLDAP)
      If UCase(Left(arrLDAP(intIdx), 3)) = "CN=" Then
        strName = Trim(Mid(arrLDAP(intIdx), 4))
      End If
    Next
  End If
  Set objUser = Nothing
  
  GetUserName = strName
  
End Function


Sub Send_Mail()
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010.
' This example sends the last saved version of the Activeworkbook object .
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
        .To = "Pawel.Zych@delphi.com"
        .CC = ""
        .BCC = ""
        .Subject = "[BLT Access] Request for granting access to BOM Leverage Database Server."
        '.body = "Add body text here" & vbNewLine & signatur
        .HTMLbody = "<html><body><p>Hello,</p><p>Please grant me access to BOM Leverage Database Server.</p>" _
                & "<b>User Name:</b> " & BLT_User_Denied.TextBox1.Value & "</p>" _
                & "<b>Role:</b>" & BLT_User_Denied.ComboBox1.Value & "</p>" _
                & "<b>PML Region:</b>" & BLT_User_Denied.ComboBox2.Value & "</p>" _
                & "<b>PBU:</b>" & BLT_User_Denied.ComboBox3.Value & "</p><br/><br/>" _
                & "<//p>Best Regards,<br/>" & UserName & "</body></html>"
        '.Attachments.Add ActiveWorkbook.FullName
        ' You can add other files by uncommenting the following line.
        '.Attachments.Add ("C:\test.txt")
        ' In place of the following statement, you can use ".Display" to
        ' display the mail.
        .Display
        '.Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub



Sub Check_User_Access()

    Dim i As Integer
    Dim LastRow As Integer
    Dim SaveName As String
    Dim wb_name As String
    Dim wb As Object
    Dim User_Dep As String
    Dim t As Integer
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    LastRow = wb.Sheets("Control").Cells(Rows.Count, "A").End(xlUp).Row
    SaveName = UserName
    
    AccessTest = False
    
    Select Case PLP_Value
    
    Case 1
        User_Dep = "PLP"
        t = 3
    Case 2
        User_Dep = "PLPA"
        t = 4
    Case 3
        User_Dep = "Administrator"
        t = 5
    End Select
        
    For i = 2 To LastRow
        If Sheets("Control").Cells(i, 1).Value = SaveName Then
            If Left(Sheets("Control").Cells(i, 2).Value, t) = User_Dep Then
                AccessTest = True
                PBU_User = Sheets("Control").Cells(i, 4).Value
                Region_User = Sheets("Control").Cells(i, 3).Value
                i = LastRow
            Else
                If Sheets("Control").Cells(i, 2).Value = "Administrator" Then
                    PBU_User = Sheets("Control").Cells(i, 4).Value
                    Region_User = Sheets("Control").Cells(i, 3).Value
                    AccessTest = True
                    i = LastRow
                End If
            End If
        End If
    Next i

End Sub


Sub SetZoom_of_Worksheets()
    
    Dim ws As Worksheet

    For Each ws In Worksheets
        If ws.Name <> "MAIN" Then
            If ws.Visible = -1 Then
                ws.Select
                ActiveWindow.Zoom = 80 '//change as per your requirements
            End If
        End If
    Next ws

End Sub



Sub BLT_Test_Version()

    Dim Excel_Version As String
    Dim Newest_Version As String
    Dim wb_name As String
    Dim wb As Object
    
    Tool_Version = 0
    
    wb_name = ThisWorkbook.Name
    Set wb = Workbooks(wb_name)
    
    Excel_Version = "Ver. " & wb.Sheets("Control").Cells(3, 18).Value
    
    Newest_Version = Tool_Function.BLT_Tool_Version.Caption

    
    If Newest_Version = Excel_Version Then
        Tool_Version = 1
    Else
        Tool_Version = 0
    End If
    

End Sub


Sub IE_Autiomation_Redirect()

    Dim ie As Object
    Set ie = CreateObject("INTERNETEXPLORER.APPLICATION")
    ie.NAVIGATE "http://s03.delphiauto.net/06/ESGSM-2/Direct%20Materials/BLT%20BOM%20Leverage%20Tool/Forms/AllItems.aspx"
    ie.Visible = True

End Sub



Sub Load_List_of_Users_Administrator()

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
    
    
    wb.Sheets("Administrator_Panel").Range("A2:J100000").ClearContents
    'wb.Sheets("Control").Range("P2:P100000").Clear
     
    'Connect to a data source:
    'For pre - MS Access 2007, .mdb files (viz. MS Access 97 up to MS Access 2003), use the Jet provider: "Microsoft.Jet.OLEDB.4.0". For Access 2007 (.accdb database) use the ACE Provider: "Microsoft.ACE.OLEDB.12.0". The ACE Provider can be used for both the Access .mdb & .accdb files.
    connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & strDB
    'connDB.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=\\plkra-fp02\GROUP$\GSM ES Costed BOM\BOMLeverageDatabaseL1\BOMLeverageDatabaseL2\BOMLeverageDatabaseL3\BOMLeverageDatabaseL4\BOM Leverage Database.accdb;" '& strDB
    
    '--------------
    'OPEN RECORDSET, ACCESS RECORDS AND FIELDS
    
    Dim ws As Worksheet
    'set the worksheet:
    Set ws = wb.Sheets("Administrator_Panel")
    
    'Set the ADO Recordset object:
    Set adoRecSet = New ADODB.Recordset
    
    'copy all records
    'strSQL = "SELECT User_Name, User_Groupe, User_Domain, User_PBU FROM User_Access"
    strSQL = "SELECT ID, User_Name, User_Groupe, User_Domain, User_PBU FROM User_Access"
    adoRecSet.Open Source:=strSQL, ActiveConnection:=connDB, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
    
    Set rng = ws.Range("A2")
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
        Subbar.NextAction "Loading User Information...", True
    '============================= Sub Bar Script End =============================
    
    Next i
    
    Application.ScreenUpdating = False
    
    wb.Sheets("Administrator_Panel").Sort.SortFields.Clear
    wb.Sheets("Administrator_Panel").Sort.SortFields.Add Key:=Range("B2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wb.Sheets("Administrator_Panel").Sort
        .SetRange Range("A2:J10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = True
    
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

