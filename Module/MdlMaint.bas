Attribute VB_Name = "MdlMaint"

Public ComNo As String
Public Function Reserial() As Boolean
On Error GoTo ErrorHandler
sqlText = "Exec sp_SerByyear"
de.con.Execute (sqlText)
Reserial = True
Exit Function
ErrorHandler:
Reserial = False
MsgBox Err.Description

End Function


Public Sub flexColWidth(mshf As MSHFlexGrid, frm As Form)
    Dim i As Integer
    Dim J As Integer
    Dim X As Integer
    With mshf
        For i = 0 To .Cols - 1
            X = 0
            For J = 0 To .Rows - 1
               If frm.TextWidth(.TextMatrix(J, i)) + 200 > X Then
                    X = frm.TextWidth(.TextMatrix(J, i)) + 200
               End If
            Next J
            .ColWidth(i) = X
        Next i
    End With
End Sub
Public Function TransDateToSql(mDate As String) As String
    If Not IsDate(mDate) Then Exit Function
   TransDateToSql = CStr(Month(mDate)) + "/" + CStr(Day(mDate)) + "/" + CStr(Year(mDate))
    'TransDateToSql = CStr(Day(mDate)) + "/" + CStr(Month(mDate)) + "/" + CStr(Year(mDate))

End Function
Function GetPass() As String
    Dim CnnPass As New ADODB.Connection
    Dim CnnTest As New ADODB.Connection
    Dim CmdTime As New ADODB.Command
    Dim rsCmdTime As New ADODB.Recordset

    Dim StrHashPass As String
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String
    Const StrSymbols As String = "~!@#$%^&*?><~!@#$%^&*?><~!@#"
    Const StrCapLettres As String = "QWERTYUIOPASDFGHJKLZXCVBNMQWERTYUIOQWERT"
    Dim StrSmlLettres As String
    Const StrDigits As String = "0123456789346127589465827346342598069"
    Dim NumDay As Integer, NumMonth As Integer, NumHour As Integer

    
    CnnPass.ConnectionString = "Provider=SQLOLEDB.1;Password=usertime;Persist Security Info=True;User ID=Usertime;Data Source=MAINSERVER;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=HUSAM;Use Encryption for Data=False;Tag with column collation when possible=False"
    CnnPass.Open
    CmdTime.ActiveConnection = CnnPass
    CmdTime.CommandText = " Select GetDate() as CurrTime "


    StrSmlLettres = LCase(StrCapLettres)
    Dim CurrDate As Date
    If rsCmdTime.state = adStateOpen Then rsCmdTime.Close
    Set rsCmdTime = CmdTime.Execute()

    CurrDate = rsCmdTime.fields(0).Value
    NumDay = Day(CurrDate)
    NumMonth = Month(CurrDate)
    NumHour = Hour(CurrDate)

    str1 = Mid(StrSymbols, NumMonth, 4)
    str2 = Mid(StrCapLettres, NumDay, 4)
    str3 = Mid(StrSmlLettres, NumDay + 4, 4)
    str4 = Mid(StrDigits, NumDay, 4)
    str5 = Mid(StrSymbols, NumHour, 4)
    StrHashPass = ""
    For i = 1 To 4
        StrHashPass = StrHashPass & Mid(str1, i, 1)
        StrHashPass = StrHashPass & Mid(str2, i, 1)
        StrHashPass = StrHashPass & Mid(str3, i, 1)
        StrHashPass = StrHashPass & Mid(str4, i, 1)
        StrHashPass = StrHashPass & Mid(str5, i, 1)
    Next i
    CnnPass.Close
    On Error Resume Next
    If CnnTest.state <> adStateClosed Then CnnTest.Close
    CnnTest.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=MAINSERVER;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=HUSAM;Use Encryption for Data=False;Tag with column collation when possible=False"
    CnnTest.Open , "user1", StrHashPass
    CnnTest.Close
    GetPass = StrHashPass

End Function

Public Function ConnectReport() As String
    ConnectReport = "odbc;dsn=DsnHafez2012;uid=" & systemConfigration.UserId & ";pwd=" & systemConfigration.Password
End Function

'Public Function ConnectName() As String
'    ConnectName = "Provider=SQLOLEDB.1;Password=" & GetPass & ";Persist Security Info=True;User ID=user1 ;Initial Catalog=Hafez2012;Data Source=MainServer"
'End Function

