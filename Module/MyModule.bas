Attribute VB_Name = "MyModule"
Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Global ConnectString As String
Global UserId As Integer
Global ClientNo As Long
Global systemConfigration  As New Configration

Global OperationEmpStr As String
Global MaintTYpeEmpStr As String
Global PaymentEmpStr As String


Global LoadForm As Boolean
Global ServerName As String
Global UID As String
Global PWD As String
Global DataBase As String
Global DatabaseYear As Integer
Global Const EnglishLayout = 67699721
Global IDBill As Double
Global idCallNo As Double
Global ProgramTitle As String





Global clientId As Double
Global ClientName As String
Global ClientPhoneNBr As String

Global GByanId As Double
Global StrNo As String
Global ByanType As String
Global SelectedStr As String
Global FormOk As Boolean
Global customerName As String
Global customerNumber As Long
Global searchClientIsAllow As Boolean


Sub ColorRow(Row As Integer, Color As Long, FlexGrid As VSFlexGrid)
With FlexGrid
    For i = 1 To .Cols - 2
        .Col = i
        .Row = Row
        .CellBackColor = Color
    Next
End With
End Sub

Function Gettag(UserId As Integer, TagId As Integer) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
    sqlText = "Select * from saipermissions Where UserId = " & UserId & " and TagId=" & TagId
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
        Gettag = True
    Else
        Gettag = False
    End If
Exit Function
ErrorHandler:
Gettag = False
End Function


Function GetFullName(UserId As Integer) As String
On Error GoTo ErrorHandler
Dim rsEmployee  As New ADODB.Recordset
sqlText = "Select FullName From empFullName Where UserId=" & UserId
Set rsEmployee = de.con.Execute(sqlText)
If rsEmployee.RecordCount > 0 Then
    GetFullName = rsEmployee!FullName
Else
    GetFullName = ""
End If

Exit Function
ErrorHandler:
GetFullName = ""
MsgBox Err.Description
End Function


Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Sub MoveCursor(KeyCode As Integer, FlexGrid As VSFlexGrid)
On Error Resume Next
If Not FlexGrid.Visible Then Exit Sub

With FlexGrid
    If KeyCode = vbKeyDown Then
        .Row = .Row + 1
    ElseIf KeyCode = vbKeyUp And .Row <> 1 Then
        .Row = .Row - 1
    End If
If Not .RowIsVisible(.Row) Then
    .TopRow = .Row
End If
.Col = 0
.ColSel = .Cols - 1
End With
End Sub

Sub ChangeToEnglish()
    If GetKeyboardLayout(0) <> EnglishLayout Then ActivateKeyboardLayout 0, 0
End Sub

Sub ChangeToArabic()
    If GetKeyboardLayout(0) = EnglishLayout Then ActivateKeyboardLayout 0, 0
End Sub

Function LikeExpression(Expr As String, Optional MultiLeter As String = "%", Optional OneLetter As String = "_") As String
    Dim X As String
    X = Replace(Expr, "Ç", "*Ç*")
    X = Replace(X, "Ã", "*Ã*")
    X = Replace(X, "Å", "*Å*")
    X = Replace(X, "Â", "*Â*")
    X = Replace(X, "ì", "*ì*")
    X = Replace(X, "í", "*í*")
    
    X = Replace(X, "*Ç*", "[ÇÃÅÂì]")
    X = Replace(X, "*Ã*", "[ÇÃÅÂ]")
    X = Replace(X, "*Å*", "[ÇÃÅÂ]")
    X = Replace(X, "*Â*", "[ÇÃÅÂ]")
    X = Replace(X, "*ì*", "[Çìí]")
    X = Replace(X, "*í*", "[ìí]")
    
    X = Replace(X, MultiLeter, "!" & MultiLeter)
    X = Replace(X, OneLetter, "!" & OneLetter)
    'Replace each space with "_%"
    X = Replace(X, " ", OneLetter & MultiLeter)
    X = Replace(MultiLeter & X & MultiLeter, MultiLeter & MultiLeter, MultiLeter)
    X = "'" & X & "'"
    If InStr(1, X, "!" & MultiLeter) > 0 Or InStr(1, X, "!" & OneLetter) > 0 Then
        X = X & " ESCAPE '!'"
    End If
    LikeExpression = X
End Function


Function ConvertControlDate(ByVal StrDate As String) As String
Dim str1 As String
If IsDate(StrDate) Then
    ConvertControlDate = Right("00" + Mid(StrDate, 4, 2), 2) + "/" + Right("00" + Mid(StrDate, 1, 2), 2) + "/" + Right("0000" + Mid(StrDate, 7, 4), 4)
Else
    ConvertControlDate = "01/01/1900"
End If
End Function

Function ConvertSqlDate(DateStr As String) As String
If IsDate(DateStr) Then
    ConvertSqlDate = Right("00" + Mid(DateStr, 1, 2), 2) + "/" + Right("00" + Mid(DateStr, 4, 2), 2) + "/" + Right("0000" + Mid(DateStr, 7, 4), 4)
Else
    ConvertSqlDate = "__/__/____"
End If
End Function
Function DeleteRow(Grid As VSFlexGrid, Vrow As Integer, Col As Integer, Table As String, Id As String, Optional ByVal Op) As Boolean
On Error GoTo ErrorHandler
With Grid
If IsMissing(Op) Then
    sqlText = "Delete From " & Table & " Where " & Id & " = " & .TextMatrix(Vrow, Col)
    de.con.Execute (sqlText)
ElseIf Op = 2 Then
    sqlText = "Delete From " & Table & " Where " & Id & " = '" & .TextMatrix(Vrow, Col) & "'"
    de.con.Execute (sqlText)
End If
End With
DeleteRow = True
Exit Function
ErrorHandler:
DeleteRow = False
MsgBox (Err.Description)
End Function

Function GetPass() As String
    GetPass = "SAIUserPassword"
End Function

Function ConnectName() As String
    systemConfigration.Password = GetPass()
    ConnectName = "Odbc;Uid=" & systemConfigration.UserId & ";Pwd=" & systemConfigration.Password & ";Dsn=" & systemConfigration.DSN & ";Database=" & systemConfigration.DatabaseName & ";"
End Function


Sub ParseXml()
    Dim doc As New MSXML2.DOMDocument60
    Dim success As Boolean
    doc.async = False
    doc.validateOnParse = True
    
 success = doc.Load(App.Path & "\init.xml")
    If Not success Then
        MsgBox "there is not configration file"
    Else
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Set nodeList = doc.selectNodes("/Application/Keys")
    If Not nodeList Is Nothing Then
    Dim node As MSXML2.IXMLDOMNode
     Dim name As String
     Dim Value As String
    
     For Each node In nodeList
        Select Case node.selectSingleNode("@name").text
         Case "SystemName":
            systemConfigration.SystemName = node.selectSingleNode("@value").text
         Case "ServerName":
            systemConfigration.ServerName = node.selectSingleNode("@value").text
         Case "DatabaseName":
            systemConfigration.DatabaseName = node.selectSingleNode("@value").text
         Case "UserName":
            systemConfigration.UserId = node.selectSingleNode("@value").text
         Case "PasswordMethod":
            systemConfigration.PasswordMethod = node.selectSingleNode("@value").text
         Case "Year":
            systemConfigration.Year = node.selectSingleNode("@value").text
        Case "Version":
            systemConfigration.Version = node.selectSingleNode("@value").text
         Case "DatabaseDestination":
            systemConfigration.DatabaseDestination = node.selectSingleNode("@value").text
         Case "Dsn":
            systemConfigration.DSN = node.selectSingleNode("@value").text
         Case "SystemUserId":
            systemConfigration.SystemUserId = node.selectSingleNode("@value").text
         Case "SystemUserPassword":
            systemConfigration.SystemUserPassword = node.selectSingleNode("@value").text
         End Select
     Next node
    End If
    End If
End Sub

