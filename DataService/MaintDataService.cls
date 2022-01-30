VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MaintDataService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Const GasNoForSpecialFamilies = 1
Const GasNoForAllFamilies = 3
Const ExternalReparationDocType = 28
Const QtyIn = 0
Const QtyOut = 1

Dim stockDataService As New stockDataService
Dim validateResult As New ReparationValidate
Dim maintCallResult As New MaintCallValidate
Dim clientValidationResult As New clientValidationResult

Dim reparationWorksCollection As New Collection
Dim reparationPiecesCollection As New Collection

Sub UpdateCallDateAndMntMoveSummaryDate(vdate As String, CallNo As Long)
On Error GoTo ErrorHandler
de.con.BeginTrans

sqlText = "update MntMoveSummary set date = '" & ConvertControlDate(vdate) & "' where callno=" & CallNo
sqlText = sqlText & "update maintcall set calldatetime= '" & ConvertControlDate(vdate) & "' + ' ' + convert(varchar(10),calldatetime,108)  where callno=" & CallNo
de.con.Execute (sqlText)
de.con.CommitTrans

Exit Sub
'02/16/2021
ErrorHandler:
de.con.RollbackTrans
MsgBox Err.Description
End Sub

Public Function IsKeyExsist(ByVal Vcol As FiledColumns, ByVal vKey As Variant) As Boolean
On Error GoTo ErrorHandler
    Dim o As fieldColumn
    Set o = Vcol.Item(vKey)
    If o Is Nothing Then
        IsKeyExsist = False
    Else
        IsKeyExsist = True
    End If
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function

Function GeSearchClients(VTop As Integer, searchFields As FiledColumns, Optional searchFormulas As String) As String
On Error GoTo ErrorHandler

Dim sqlText As String, sqlFilter As String
sqlText = "select top " & VTop & "  AdhamNo , AdhamName , AdhamPhon , MobilePhone , AdhamAdress From AdhamView7"
sqlText = sqlText & " where AdhamNo <> -1"

sqlFilter = ""

If searchFields Is Nothing Then
    'Set rsClients = de.con.Execute(sqlText)
     GeSearchClients = sqlText
    Exit Function
End If
If IsKeyExsist(searchFields, "AdhamName") Then
    If searchFields.Item("AdhamName").isChecked Then
        sqlFilter = sqlFilter & "Or AdhamName like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "adhamPhon") Then
    If searchFields.Item("adhamPhon").isChecked Then
        sqlFilter = sqlFilter & "Or adhamPhon like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "MobilePhone") Then
    If searchFields.Item("MobilePhone").isChecked Then
        sqlFilter = sqlFilter & "Or MobilePhone like" & LikeExpression(searchFormulas)
    End If
End If

If sqlFilter <> "" Then
    sqlFilter = Mid(sqlFilter, 3, Len(sqlFilter))
End If


If Not IsMissing(searchFormulas) And searchFormulas <> "" Then
   sqlText = sqlText & " and  (" & sqlFilter & ")"
End If


GeSearchClients = sqlText
Exit Function
ErrorHandler:
GeSearchClients = ""
End Function


Function GetClients(withReparation As Boolean, VTop As Double, searchFields As FiledColumns, Optional searchFormulas As String) As String
On Error GoTo ErrorHandler

Dim sqlText As String, sqlFilter As String

If VTop = 0 Then
     GetClients = ""
End If
sqlText = "select top " & VTop & "  AdhamNo , AdhamName , AdhamPhon , MobilePhone , AdhamAdress, Zone ,   ZoneName , Notes , defindname "
sqlText = sqlText & " ,countRec = (select COUNT(*) from  maintcall m1  inner join reparation r1 on m1.callno = r1.callno where m1.clino = a1.adhamno ) from adhamview7 a1 left outer join coZone c1 on a1.zone = c1.zoneno "
sqlText = sqlText & " where AdhamNo <> -1"

If Not withReparation Then
    sqlText = sqlText & " and (select COUNT(*) from  maintcall m1  inner join reparation r1 on m1.callno = r1.callno where m1.clino = a1.adhamno ) >0"
End If

sqlFilter = ""

If searchFields Is Nothing Then
    'Set rsClients = de.con.Execute(sqlText)
     GetClients = sqlText
    Exit Function
End If
If IsKeyExsist(searchFields, "AdhamName") Then
    If searchFields.Item("AdhamName").isChecked Then
        sqlFilter = sqlFilter & "Or AdhamName like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "adhamPhon") Then
    If searchFields.Item("adhamPhon").isChecked Then
        sqlFilter = sqlFilter & "Or adhamPhon like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "MobilePhone") Then
    If searchFields.Item("MobilePhone").isChecked Then
        sqlFilter = sqlFilter & "Or MobilePhone like" & LikeExpression(searchFormulas)
    End If
End If


If IsKeyExsist(searchFields, "ZoneName") Then
    If searchFields.Item("ZoneName").isChecked Then
        sqlFilter = sqlFilter & "Or ZoneName like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "defindname") Then
    If searchFields.Item("defindname").isChecked Then
        sqlFilter = sqlFilter & "Or defindname like" & LikeExpression(searchFormulas)
    End If
End If


If IsKeyExsist(searchFields, "AdhamAdress") Then
    If searchFields.Item("AdhamAdress").isChecked Then
        sqlFilter = sqlFilter & "Or AdhamAdress like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "defindname") Then
    If searchFields.Item("defindname").isChecked Then
        sqlFilter = sqlFilter & "Or defindname like" & LikeExpression(searchFormulas)
    End If
End If

If IsKeyExsist(searchFields, "notes") Then
    If searchFields.Item("notes").isChecked Then
        sqlFilter = sqlFilter & "Or notes like" & LikeExpression(searchFormulas)
    End If
End If

If sqlFilter <> "" Then
    sqlFilter = Mid(sqlFilter, 3, Len(sqlFilter))
End If


If Not IsMissing(searchFormulas) And searchFormulas <> "" Then
   sqlText = sqlText & " and  (" & sqlFilter & ")"
End If

'Set rsClients = de.con.Execute(sqlText)


GetClients = sqlText & " Order by AdhamNo desc"
Exit Function
ErrorHandler:
GetClients = ""
End Function

Sub ChangeTheZoneName(ZoneName As String, ZoneNo As Long)
On Error GoTo ErrorHandler
    sqlText = "Update CoZone Set ZoneName='" & ZoneName & "' where zoneNo =" & ZoneNo
    de.con.Execute (sqlText)
Exit Sub
ErrorHandler:

MsgBox Err.Description
End Sub

Function InsertNewZone(ZoneName As String) As Long
On Error GoTo ErrorHandler
Dim rsZoneNo As New ADODB.Recordset
Dim sqlText As String
sqlText = "Insert into CoZone(ZoneName)Values('" & LTrim(RTrim(ZoneName)) & "')"
de.con.Execute (sqlText)
sqlText = "Select @@Identity as ZoneNo"
Set rsZoneNo = de.con.Execute(sqlText)
InsertNewZone = rsZoneNo!ZoneNo
Exit Function
ErrorHandler:
InsertNewZone = -1
MsgBox Err.Description
End Function

Sub TransferOrdersToMaintCall(unPrintedMaintCalls As String)
On Error GoTo ErrorHandler
If unPrintedMaintCalls = "" Then Exit Sub
Dim sqlText As String


sqlText = ""
sqlText = sqlText & " Insert into maintcall ( CallDateTime, ModNo, CliNo, CallDescription, CallState, CallReceiverEmpNo, IsCheck, HostName, CallOrderId )"

sqlText = sqlText & " Select "
sqlText = sqlText & "    convert(varchar(10),dateadd(dd,1,CallDateTime),120)+ ' ' + convert(varchar(20),getdate(),108), "
sqlText = sqlText & "    ModNo, "
sqlText = sqlText & "    CliNo, "
sqlText = sqlText & "    CallDescription, "
sqlText = sqlText & "    CallState, "
sqlText = sqlText & "    CallReceiverEmpNo, "
sqlText = sqlText & "    0, "
sqlText = sqlText & "    HostName, "
sqlText = sqlText & "    CallNo "
sqlText = sqlText & " from maintcallorder where callno in(" & unPrintedMaintCalls & ")"
de.con.Execute (sqlText)

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub


Sub UpdateMaintCallState(unPrintedMaintCalls As String, EnumMaintCallStateRec As EnumMaintCallState)
On Error GoTo ErrorHandler
    If unPrintedMaintCalls = "" Then Exit Sub
    
    If EnumMaintCallStateRec = UnderwayAndPrinted Then
        sqlText = "Update MaintCall Set CallState = " & EnumMaintCallState.UnderwayAndPrinted
        sqlText = sqlText & " Where callNo in (" & unPrintedMaintCalls & ") and isnull(CallState,0) =0"

    ElseIf EnumMaintCallStateRec = 0 Then
        sqlText = "Update MaintCall Set CallState = 0  Where callNo in (" & unPrintedMaintCalls & ")"
        sqlText = sqlText & "and isnull(CallState,0) =" & EnumMaintCallState.UnderwayAndPrinted
    
    End If
    de.con.Execute (sqlText)
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Function GetMaintCallsInfoStatistics() As ADODB.Recordset
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsMaintCallsStatistics As New ADODB.Recordset
     
    sqlText = ""
    sqlText = sqlText & "   Select"
    sqlText = sqlText & "   Count(*)CountAllMaintCalls , "
    sqlText = sqlText & "   Sum(Case When  isnull(m1.callstate,0) & " & EnumMaintCallState.UnderwayAndPrinted & "=" & EnumMaintCallState.UnderwayAndPrinted & "   then 1 else 0 end) CountUnderwayAndPrinted, "
    sqlText = sqlText & "   Sum(Case When  isnull(m1.callstate,0)=0 then 1 else 0 end) CountUnderWay"
    
'    sqlText = sqlText & "   Sum(Case When  m1.callstate & " & EnumMaintCallState.Underway & "=" & EnumMaintCallState.Underway & " and m1.callstate & " & EnumMaintCallState.Printed & "=" & EnumMaintCallState.Printed & "   then 0 else 1 end) CountUnderway , "
'    sqlText = sqlText & "   Sum(Case When  m1.callstate & " & EnumMaintCallState.Printed & "=" & EnumMaintCallState.Printed & " and m1.callstate & " & EnumMaintCallState.Repared & "=" & EnumMaintCallState.Repared & "   then 0 else 1 end) Countprinted  "
'    sqlText = sqlText & "   Sum(Case When  m1.callstate & " & EnumMaintCallState.Repared & "=" & EnumMaintCallState.Repared & " then 1 else 0 end) CountRepared  "
   
    sqlText = sqlText & "   From  "
    sqlText = sqlText & "   maintcall m1 left outer join"
    sqlText = sqlText & "   reparation r1 on m1.callno = r1.callno where r1.CallNo is null"
    
    Set rsMaintCallsStatistics = de.con.Execute(sqlText)
    Set GetMaintCallsInfoStatistics = rsMaintCallsStatistics
    
Exit Function
ErrorHandler:
MsgBox Err.Description

End Function

Function GetMaintCallOrdersInfo(vdate As String, allOrders As Boolean, Optional searchFormula As String = "") As ADODB.Recordset
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsMaintCalls As New ADODB.Recordset
     
    sqlText = ""
    sqlText = sqlText & "select"
    sqlText = sqlText & "   m1.IsCheck Chk , "
    sqlText = sqlText & "   m2.CallNo  , "
    sqlText = sqlText & "   m1.callNo , "
    sqlText = sqlText & "   Convert(varchar(10),m1.CallDateTime,103)CallDate , "
    sqlText = sqlText & "   Convert(varchar(5),m1.CallDateTime,108)CallTime , "
    sqlText = sqlText & "   adhamname , "
    sqlText = sqlText & "   a1.zone , "
    sqlText = sqlText & "   c1.ZoneName , "
    sqlText = sqlText & "   adhamadress , "
    sqlText = sqlText & "   adhamphon , "
    sqlText = sqlText & "   MobilePhone , "
    sqlText = sqlText & "   m1.ModNo , "
    sqlText = sqlText & "   a2.prodfamname , "
    sqlText = sqlText & "   m1.calldescription , "
    sqlText = sqlText & "   a1.Notes , "
    sqlText = sqlText & "   a1.defindname , "
    sqlText = sqlText & "   m1.CallReceiverEmpNo  , "
    sqlText = sqlText & "   e1.FullName  ,"
    sqlText = sqlText & "       countrec ="
    sqlText = sqlText & "       (select count(*)"
    sqlText = sqlText & "       From MaintcallOrder mc where convert(varchar(10),mc.calldatetime,101)  = convert(varchar(10),m1.calldatetime,101) and mc.clino =m1.CliNo"
    sqlText = sqlText & "       )"
     
    sqlText = sqlText & "   From  "
    sqlText = sqlText & "   MaintCallOrder m1 inner join "
    sqlText = sqlText & "   AdhamProductFamily a2 on m1.ModNo = a2.ProdFamNo inner join"
    sqlText = sqlText & "   adhamview7 a1 on m1.clino = a1.adhamno left outer join"
    sqlText = sqlText & "   EmpFullName e1 on m1.CallReceiverEmpNo = e1.EmpNo left outer join"
    sqlText = sqlText & "   CoZone c1 on a1.zone = c1.zoneno left outer join "
    sqlText = sqlText & "  MaintCall m2 on m1.callNo = m2.CallOrderId "
    sqlText = sqlText & " Where m1.callNo is not null "
    
    If IsDate(vdate) Then
            sqlText = sqlText & " and convert(varchar(10),m1.calldatetime,101)='" & ConvertControlDate(vdate) & "'"
    Else
            sqlText = sqlText & " and convert(varchar(10),m1.calldatetime,101)='" & ConvertControlDate(Format(Date, "dd/mm/yyyy")) & "'"
    End If
    
    If allOrders Then
        sqlText = sqlText & " and m2.callNo is not null"
    Else
        sqlText = sqlText & " and m2.callNo is  null"
    End If
    
    
    
    If searchFormula <> "" Then
        sqlText = sqlText & " and  (Convert(varchar(10),m1.CallNo) ='" & searchFormula & "'"
        sqlText = sqlText & " or  Convert(varchar(10),m2.CallNo) ='" & searchFormula & "'"
        sqlText = sqlText & " or m1.calldescription like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or adhamname like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or ZoneName like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or adhamadress like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or adhamphon like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or MobilePhone like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or a1.Notes like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or a1.defindname like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or FullName like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or ProdFamName like " & LikeExpression(searchFormula)
        sqlText = sqlText & ")"
    End If
    
    

    Set rsMaintCallOrders = de.con.Execute(sqlText)
    Set GetMaintCallOrdersInfo = rsMaintCallOrders
    
Exit Function
ErrorHandler:
MsgBox Err.Description

End Function

Function GetMaintCallsInfo(Optional searchFormula As String = "") As ADODB.Recordset
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsMaintCalls As New ADODB.Recordset
     
    sqlText = ""
    sqlText = sqlText & "select"
    If searchFormula <> "" Then
         sqlText = sqlText & " Top 10"
    End If
    sqlText = sqlText & "   IsCheck Chk , "
    sqlText = sqlText & "   m1.CallOrderId , "
    sqlText = sqlText & "   m1.callNo , "
    sqlText = sqlText & "   m1.callstate , "
    sqlText = sqlText & "   Convert(varchar(10),m1.CallDateTime,103)CallDate , "
    sqlText = sqlText & "   Convert(varchar(5),m1.CallDateTime,108)CallTime , "
    sqlText = sqlText & "   adhamname , "
    sqlText = sqlText & "   a1.zone , "
    sqlText = sqlText & "   c1.ZoneName , "
    sqlText = sqlText & "   adhamadress , "
    sqlText = sqlText & "   adhamphon , "
    sqlText = sqlText & "   MobilePhone , "
    sqlText = sqlText & "   m1.ModNo , "
    sqlText = sqlText & "   a2.prodfamname , "
    sqlText = sqlText & "   calldescription , "
    sqlText = sqlText & "   a1.Notes , "
    sqlText = sqlText & "   a1.defindname , "
    sqlText = sqlText & "   CallReceiverEmpNo  , "
    sqlText = sqlText & "   e1.FullName  ,"
    sqlText = sqlText & "       countrec ="
    sqlText = sqlText & "       (select count(*)"
    sqlText = sqlText & "       From Maintcall mc where convert(varchar(10),mc.calldatetime,101)  = convert(varchar(10),m1.calldatetime,101) and mc.clino =m1.CliNo"
    sqlText = sqlText & "       )"
    
    





    
    
    
    sqlText = sqlText & "   From  "
    sqlText = sqlText & "   MaintCall m1 inner join "
    sqlText = sqlText & "   AdhamProductFamily a2 on m1.ModNo = a2.ProdFamNo inner join"
    sqlText = sqlText & "   adhamview7 a1 on m1.clino = a1.adhamno left outer join"
    sqlText = sqlText & "   EmpFullName e1 on m1.CallReceiverEmpNo = e1.EmpNo left outer join"
    sqlText = sqlText & "   CoZone c1 on a1.zone = c1.zoneno left outer join"
    sqlText = sqlText & "   reparation r1 on m1.callno = r1.callno"
    sqlText = sqlText & " Where r1.CallNo is null "
    If searchFormula <> "" Then
        sqlText = sqlText & " and  (Convert(varchar(10),m1.CallNo) ='" & searchFormula & "'"
        sqlText = sqlText & " or calldescription like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or adhamname like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or ZoneName like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or adhamadress like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or adhamphon like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or MobilePhone like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or a1.Notes like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or a1.defindname like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or FullName like " & LikeExpression(searchFormula)
        sqlText = sqlText & " or ProdFamName like " & LikeExpression(searchFormula)
        sqlText = sqlText & ")"
    End If
    
    

    Set rsMaintCalls = de.con.Execute(sqlText)
    Set GetMaintCallsInfo = rsMaintCalls
    
Exit Function
ErrorHandler:
MsgBox Err.Description

End Function

Function GetLastCallNoForTheClient(cliNo As Long, callDate As Date) As Variant
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsLastCallNo As New ADODB.Recordset
sqlText = "select top 1 r1.CallNo from Reparation r1 inner join maintcall m1 on r1.CallNo = m1.CallNo "
sqlText = sqlText & "   where m1.clino=" & cliNo & " and RepDate <= '" & ConvertControlDate(callDate) & "'"
sqlText = sqlText & " order by r1.callno desc"
Set rsLastCallNo = de.con.Execute(sqlText)
If Not rsLastCallNo.EOF Then
    GetLastCallNoForTheClient = rsLastCallNo!CallNo
Else
    GetLastCallNoForTheClient = Null
End If


Exit Function
ErrorHandler:
GetLastCallNoForTheClient = Null
MsgBox Err.Description

End Function

Function GetLastVisitForTheClient(cliNo As Long, callDate As Date) As Variant
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsLastVisit As New ADODB.Recordset
sqlText = "select top 1 RepDate from Reparation r1 inner join maintcall m1 on r1.CallNo = m1.CallNo "
sqlText = sqlText & "    where m1.clino=" & cliNo & " and RepDate <= '" & ConvertControlDate(callDate) & "'"
sqlText = sqlText & " order by r1.callno desc"
Set rsLastVisit = de.con.Execute(sqlText)
If Not rsLastVisit.EOF Then
    GetLastVisitForTheClient = rsLastVisit!RepDate
Else
    GetLastVisitForTheClient = Null
End If


Exit Function
ErrorHandler:
GetLastVisitForTheClient = Null
MsgBox Err.Description

End Function

Function GetClientVisitorCount(cliNo As Long, callDate As Date) As Integer
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsClientVisitorsCount As New Recordset
    sqlText = ""
    sqlText = sqlText & "   Select"
    sqlText = sqlText & "   COUNT(*)CountRec From "
    sqlText = sqlText & "   MaintCall m1  inner join "
    sqlText = sqlText & "   Reparation r1 on m1.callno = r1.callno "
    sqlText = sqlText & "   where   r1.repdate<='" & ConvertControlDate(callDate) & "' And m1.cliNo = " & cliNo
    Set rsClientVisitorsCount = de.con.Execute(sqlText)
    If rsClientVisitorsCount.RecordCount <> 0 Then
        GetClientVisitorCount = rsClientVisitorsCount!CountRec
    Else
         GetClientVisitorCount = 0
    End If
    
Exit Function
ErrorHandler:
GetClientVisitorCount = 0
MsgBox Err.Description
End Function

Function GetFamillyListWithGas() As Integer()
Dim famillyListWithGas(4) As Integer
     famillyListWithGas(0) = 1
     famillyListWithGas(1) = 17
     famillyListWithGas(2) = 16
     famillyListWithGas(3) = 11
    GetFamillyListWithGas = famillyListWithGas
End Function


Public Function GetMaxReparationNo() As Double
    On Error GoTo ErrorHandler
    Dim RsMax As New ADODB.Recordset
    sqlText = "Select isnull(Max(RepNo),0) MaxRepNo From Reparation"
    Set RsMax = de.con.Execute(sqlText)
    GetMaxReparationNo = RsMax!MaxRepNo + 1
    Exit Function
ErrorHandler:
    MsgBox Err.Description
    GetMaxReparationNo = -1
End Function

Function GetModelNo(ItemNo As String) As Integer
On Error GoTo ErrorHandler
    Dim sqlText As String
    Dim rsItem As New ADODB.Recordset
    sqlText = "Select ModNo From AdhamModels Where ItemNo ='" & ItemNo & "'"
    Set rsItem = de.con.Execute(sqlText)
    If Not rsItem.EOF Then
        GetModelNo = rsItem!ModNo
    Else
        GetModelNo = 0
    End If
Exit Function
ErrorHandler:
    MsgBox Err.Description
    GetModelNo = 0
End Function
Function GetMaxProdNo() As Double
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsMaxProdNo As New ADODB.Recordset
sqlText = "select isnull(Max(ProdNo),0) as MaxProdNo from AdhamProducts"
Set rsMaxProdNo = de.con.Execute(sqlText)
GetMaxProdNo = rsMaxProdNo!maxProdNo
Exit Function
ErrorHandler:
GetMaxProdNo = 0
End Function

Function GetCollectionFromSql(reparationInfo As ReparationViewModel, Vindex As Integer) As Collection
On Error GoTo ErrorHandler
    Dim sqlText As String
    Dim rs As New ADODB.Recordset
    Dim items As New Collection
    
    Select Case Vindex
        Case 1:
        sqlText = "Select reptypeNo as item , reptypeNo as vkey From ReparationWorks Where repNo=" & reparationInfo.repNo
        Case 2:
            sqlText = "select pieceNo as item  , pieceNo as vkey  from reparationpieces r1 where RepNo =" & reparationInfo.repNo
    End Select
    Set rs = de.con.Execute(sqlText)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            items.Add rs.fields("item").Value, rs.fields("vkey").Value
            rs.MoveNext
        Loop
    End If
    Set GetCollectionFromSql = items
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
    Set GetCollectionFromSql = items

End Function

'Function GetMaxProdNo() As Double
'
'    Dim rsMaxProdNo As New ADODB.Recordset
'    sqltext = "select isnull(Max(ProdNo),0) as MaxProdNo from AdhamProducts"
'    Set rsMaxProdNo = de.con.Execute(sqltext)
'    GetMaxProdNo = rsMaxProdNo!maxProdNo + 1
'
'End Function

Function GetClientInfoQuery(maintCallInfo As MaintCallViewModel) As String
On Error GoTo ErrorHandler
Dim adhamNo As Long
Dim sqlText As String
sqlText = ""

With maintCallInfo.clientInfo
If .ClientState = UpdateClient Then
    sqlText = "update AdhamView7 set adhamname='" & .AdhamName & "',adhamphon='" & .AdhamPhon & "',MobilePhone='" & _
    .MobilePhone & "',adhamadress='" & .AdhamAdress & "',zone='" & .Zone & "',Notes='" & .Notes & "',Defindname='" & .Defindname & "' Where adhamNo = " & .adhamNo
    
ElseIf .ClientState = NewClient Then
    .adhamNo = GetNewClientNo()
    sqlText = "insert into AdhamView7(Id_ComNo, AdhamNo,  adhamname, adhamphon, MobilePhone, adhamadress, zone,kind ,  Notes , Defindname   )Values(1," & _
    .adhamNo & ",'" & .AdhamName & "','" & .AdhamPhon & "','" & .MobilePhone & "','" & .AdhamAdress & "','" & .Zone & "',2,'" & .Notes & "','" & .Defindname & "')"
    maintCallInfo.cliNo = .adhamNo
End If
End With
    GetClientInfoQuery = sqlText
Exit Function
ErrorHandler:
GetClientInfoQuery = ""
MsgBox Err.Description
End Function


Function GetMaxCallNo() As Long
On Error GoTo ErrorHandler
Dim rsMaxCallNo As New ADODB.Recordset
sqlText = "Select isnull(Max(CallNo),0)MaxCallNo from MaintCall"
Set rsMaxCallNo = de.con.Execute(sqlText)
GetMaxCallNo = rsMaxCallNo!MaxCallNo + 1

Exit Function
ErrorHandler:
MsgBox Err.Description
GetMaxCallNo = -1
End Function

Function GetMaintCallInfoQuery(maintCallInfo As MaintCallViewModel, vState As EnumState, choice As Boolean) As String
On Error GoTo ErrorHandler
Dim sqlText As String
sqlText = ""

With maintCallInfo
    If vState = NewRecord Then
        '.CallNo = GetMaxCallNo()
        If choice Then
            sqlText = sqlText & " insert into MaintCall( "
        Else
           sqlText = sqlText & " insert into maintCallOrder("
        End If
'        sqlText = sqlText & "Id_ComNo,"
'        sqlText = sqlText & "CallState,"
        sqlText = sqlText & "CallDateTime,"
        sqlText = sqlText & "ModNo,"
        sqlText = sqlText & "CliNo,"
        sqlText = sqlText & "CallDescription,"
'        sqlText = sqlText & "Notes,"
'        sqlText = sqlText & "defindname,"
        sqlText = sqlText & "CallReceiverEmpNo,"
        sqlText = sqlText & "CallState) "
        sqlText = sqlText & "Values("
        'sqlText = sqlText & "," & .CallNo
'        sqlText = sqlText & ",0"
        sqlText = sqlText & "'" & Format(.CallDateTime, "mm/dd/yyyy HH:MM") & "'"
        sqlText = sqlText & "," & .ModNo
        sqlText = sqlText & "," & .cliNo
        sqlText = sqlText & ",'" & .CallDEscription & "'"
'        sqlText = sqlText & ",'" & .Notes & "'"
'        sqlText = sqlText & ",'" & .Defindname & "'"
        sqlText = sqlText & "," & .CallReceiverEmpNo
        sqlText = sqlText & ",0"
        sqlText = sqlText & ")"
 '       sqlText = sqlText & "Select @@Identity as LastId "
        
    ElseIf vState = UpdateRecord Then

        sqlText = sqlText & "   Update MaintCall Set "
        sqlText = sqlText & "   ModNo='" & .ModNo & "'"
        sqlText = sqlText & "   ,CliNo='" & .cliNo & "'"
        sqlText = sqlText & "   ,CallDescription='" & .CallDEscription & "'"
'        sqlText = sqlText & "   ,Notes='" & .Notes & "'"
'        sqlText = sqlText & "   ,defindname='" & .Defindname & "'"
'        sqlText = sqlText & "   ,CallReceiverEmpNo=" & .CallReceiverEmpNo
'        sqlText = sqlText & "   ,CallEntryEmpNo=" & .CallEntryEmpNo
        sqlText = sqlText & "   Where CallNo=" & .CallNo
   ElseIf vState = DeleteRecord Then
        sqlText = sqlText & " Delete From MaintCall Where CAllNo=" & .CallNo
    End If
End With
GetMaintCallInfoQuery = sqlText
Exit Function
ErrorHandler:
GetMaintCallInfoQuery = ""
MsgBox Err.Description

End Function


Function GetNewClientNo() As Long
On Error GoTo errorhandkler
Dim rsClient As New ADODB.Recordset
sqlText = "Select isnull(Max(AdhamNo),0) as MaxAdhamNo From AdhamView7"
Set rsClient = de.con.Execute(sqlText)
GetNewClientNo = rsClient!MaxAdhamNo + 1
Exit Function
errorhandkler:
GetNewClientNo = -1
MsgBox Err.Description
End Function


Function GetProductInformationQuery(reparationInfo As ReparationViewModel) As String
Dim sqlText As String
Dim maxProdNo As Double
sqlText = ""
With reparationInfo.AdhamProductViewModel
        .ProdBaCodeNo = GenerateBarcodeNo(reparationInfo.AdhamProductViewModel)
        If (reparationInfo.ReparationState = NewRecord Or reparationInfo.ReparationState = UpdateRecord) And .ModNo <> 0 And (IsNull(.ProdNo) Or .ProdNo = 0) Then
            maxProdNo = GetMaxProdNo
            .ProdNo = maxProdNo + 1
            sqlText = sqlText & " Insert into AdhamProducts "
            sqlText = sqlText & " ( ProdNo , ProdBaCodeNo , ProdDate , ProdPurchaseDate , BarcodeSerNo , GasNo,ModNo) values ( "
            sqlText = sqlText & .ProdNo & " , "
            sqlText = sqlText & IIf(.ProdBaCodeNo & "" = "", " null ", "'" & Trim(.ProdBaCodeNo) & "'")
            
            If Not IsDate(.ProdDate) And IsNull(.ProdDate) Then
                 sqlText = sqlText & ",null"
            Else
                sqlText = sqlText & ",'" & ConvertControlDate(.ProdDate) & "'"
            End If
            
            If Not IsDate(.ProdPurchaseDate) And IsNull(.ProdPurchaseDate) Then
                 sqlText = sqlText & ",null"
            Else
                sqlText = sqlText & ",'" & ConvertControlDate(.ProdPurchaseDate) & "'"
            End If
            
            sqlText = sqlText & IIf(.GasNo & "" = "", ",null ", ",'" & Trim(.GasNo) & "'")
            sqlText = sqlText & IIf(.BarcodeSerialNo & "" = "", " ,null ", ",'" & Trim(.BarcodeSerialNo) & "'")
            
            sqlText = sqlText & "," & .ModNo & ")"
        ElseIf reparationInfo.ReparationState = UpdateRecord And Not IsNull(.ProdNo) Then
            If .ModNo = 0 Or .ProdBaCodeNo = "" Or IsNull(.ProdPurchaseDate) Or Not IsDate(.ProdPurchaseDate) Then
                sqlText = "delete from AdhamProducts where prodNo=" & .ProdNo
            Else
                sqlText = sqlText & " Update AdhamProducts Set  ProdBaCodeNo='" & .ProdBaCodeNo & "',ProdDate='" & ConvertControlDate(.ProdDate) & "'"
                sqlText = sqlText & ",ProdPurchaseDate='" & ConvertControlDate(.ProdPurchaseDate) & "',ModNo=" & .ModNo
                sqlText = sqlText & ",GasNo='" & .GasNo & "'"
                sqlText = sqlText & ",BarcodeSerNo='" & .BarcodeSerialNo & "'"
                sqlText = sqlText & " Where ProdNo=" & .ProdNo
            End If
        ElseIf reparationInfo.ReparationState = DeleteRecord Then
            If Not IsNull(reparationInfo.ProdNo) Then
                sqlText = sqlText & " Delete From AdhamProducts Where ProdNo=" & reparationInfo.ProdNo
            End If
   End If
End With
GetProductInformationQuery = sqlText

End Function

Function GetReparationWorksQuery(reparationInfo As ReparationViewModel) As String

Dim sqlText As String
Dim reparationWorkInfo As ReparationWorkViewModel
sqlText = ""

Set reparationWorksCollection = GetCollectionFromSql(reparationInfo, 1)
For Each reparationWorkInfo In reparationInfo.ReparationWorksViewModel
    With reparationWorkInfo
        If Not Exists(reparationWorksCollection, .RepTypeNo) Then
            sqlText = sqlText & " Insert into ReparationWorks(repno , reptypeno)Values("
            sqlText = sqlText & reparationInfo.repNo & ",'" & .RepTypeNo & "')"
        Else
            reparationWorksCollection.Remove .RepTypeNo
        End If
    End With
Next
' Delete collectom from sql
For deleteWorkState = 1 To reparationWorksCollection.Count
    sqlText = sqlText & " delete from ReparationWorks where repno=" & reparationInfo.repNo & " and reptypeno='"
    sqlText = sqlText & reparationWorksCollection(deleteWorkState) & "'"
Next deleteWorkState

GetReparationWorksQuery = sqlText
End Function

Function GetReparationPiecesQuery(reparationInfo As ReparationViewModel) As String
Dim sqlText As String
Dim reparationPieceInfo As ReparationPieceViewModel
sqlText = ""

    Set reparationPiecesCollection = GetCollectionFromSql(reparationInfo, 2)
        For Each reparationPieceInfo In reparationInfo.ReparationPiecesViewModel
            With reparationPieceInfo
                If Not Exists(reparationPiecesCollection, .stkno) Then
                    Dim ByanId As Double
                    ByanId = stockDataService.GetMaxByanId
                    sqlText = sqlText & " Insert into ReparationPieces(RepNo , PieceNo ,Qty , Price)Values("
                    sqlText = sqlText & reparationInfo.repNo & ",'" & .stkno & "'," & .Qty & "," & .Price & ")"
                    
                    sqlText = sqlText & "Insert Into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType,Empno)Values("
                    sqlText = sqlText & ByanId & "," & .stkId & "," & reparationInfo.Strid & ",getdate () ,"
                    sqlText = sqlText & ExternalReparationDocType & "," & reparationInfo.CallNo & "," & .Qty & "," & QtyOut & "," & empNo & ")"
                Else
                    reparationPiecesCollection.Remove .stkno
                End If
            End With
        Next
    
    'Delete Pieces
    Dim maxByanId As Double
    maxByanId = stockDataService.GetMaxByanId
    For deleteStkNo = 1 To reparationPiecesCollection.Count
    
        sqlText = sqlText & " Insert Into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType,Empno )"
        sqlText = sqlText & " select " & maxByanId & ", c1.id StkId  , n1.Id strid , getdate() , " & ExternalReparationDocType & "  , m1.CallNo ,  r1.Qty , " & QtyIn & "," & empNo & " from"
        sqlText = sqlText & " reparationpieces r1  inner join"
        sqlText = sqlText & " CoStock c1 on r1.pieceno = c1.stkno collate arabic_ci_as inner join"
        sqlText = sqlText & " reparation r2 on r1.RepNo = r2.RepNo inner join"
        sqlText = sqlText & " maintcall m1 on r2.CallNo = m1.callno inner join"
        sqlText = sqlText & " MntMoveSummary m2 on m1.callno = m2.CallNo inner join"
        sqlText = sqlText & " MaintTeam m3 on m2.TeamNo = m3.TeamNo inner join"
        sqlText = sqlText & " namestr n1 on m3.empno = n1.strno"
        sqlText = sqlText & " where PieceNo = '" & reparationPiecesCollection(deleteStkNo) & "' and r1.repno =" & reparationInfo.repNo
        maxByanId = maxByanId + 1
        
        sqlText = sqlText & " delete from ReparationPieces where repno=" & reparationInfo.repNo & " and PieceNo='"
        sqlText = sqlText & reparationPiecesCollection.Item(deleteStkNo) & "'"
        
    Next deleteStkNo
GetReparationPiecesQuery = sqlText
End Function

Function GetProdNo(reparationProductViewModel As AdhamProductViewModel) As Variant

With reparationProductViewModel
        If .ModNo <> 0 _
        And .ProdBaCodeNo <> "" _
        And Not IsNull(.ProdPurchaseDate) _
        And IsDate(.ProdPurchaseDate) _
        And Not IsNull(.ProdNo) Then
            GetProdNo = .ProdNo
        Else
            GetProdNo = Null
        End If
End With

End Function

Function GetFullName(empNo As Integer) As String
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rsFullName As New ADODB.Recordset
sqlText = "Select FullName From _Source .Dbo.FullName Where empno = " & empNo
Set rsFullName = de.con.Execute(sqlText)
If rsFullName.RecordCount > 0 Then
    GetFullName = rsFullName!FullName
Else
    GetFullName = ""
End If
Exit Function
ErrorHandler:
GetFullName = ""
MsgBox Err.Description
End Function

Public Function GetReparationInfoQuery(reparationInfo As ReparationViewModel) As String

Dim sqlText As String
sqlText = ""
With reparationInfo
.empNo = empNo
.FullName = GetFullName(.empNo)
If .ReparationState = NewRecord Then
        .repNo = GetMaxReparationNo
        .ProdNo = GetProdNo(.AdhamProductViewModel)
        sqlText = sqlText & " insert into Reparation(RepNo, CallNo, TeamNo, ProdNo, RepDate, RepTimeBegin, RepTimeEnd, RepPrice, "
        sqlText = sqlText & "Cash,regestdate, Notes,Carried, CliRecever, Description, VoltBefor, VoltAfter,EmpNo)Values("
        sqlText = sqlText & .repNo & "," & .CallNo & "," & .TeamNo & ","
        sqlText = sqlText & IIf(IsNull(.ProdNo), "null", .ProdNo) & ",'" & ConvertControlDate(.RepDate) & "','" & .RepTimeBegin & "','" & .RepTimeEnd & "'," & .RepPrice & ","
        sqlText = sqlText & IIf(IsNull(.Cash), "null", .Cash) & ",'" & ConvertControlDate(.regestdate) & "','" & .Notes & "'," & IIf(.Carried, 1, 0) & ","
        sqlText = sqlText & IIf(IsNull(.CliRecever), "null", .CliRecever) & ",'" & .Description & "'," & .VoltBefor & "," & .VoltAfter & "," & empNo & ")"
        sqlText = sqlText & " update maintCall Set CallStatus=" & IIf(IsNull(.CallStatus), "null", .CallStatus) & "  Where callNo=" & .CallNo
ElseIf .ReparationState = UpdateRecord Then
        .ProdNo = GetProdNo(.AdhamProductViewModel)
        sqlText = " update Reparation set TeamNo=" & .TeamNo & ",ProdNo=" & IIf(IsNull(.ProdNo), "null", .ProdNo) & ",RepDate='" & ConvertControlDate(.RepDate)
        sqlText = sqlText & "',RepTimeBegin='" & .RepTimeBegin & "',RepTimeEnd='" & .RepTimeEnd & "',RepPrice=" & .RepPrice
        sqlText = sqlText & ",Cash=" & IIf(IsNull(.Cash), "null", .Cash) & ",regestdate='" & ConvertControlDate(.regestdate) & "',Notes='" & .Notes & "'"
        sqlText = sqlText & ",Carried=" & IIf(.Carried, 1, 0) & ", CliRecever=" & IIf(IsNull(.CliRecever), "null", .CliRecever) & ", Description='" & .Description
        sqlText = sqlText & "',VoltBefor=" & .VoltBefor & ", VoltAfter=" & .VoltAfter & ",EmpNo=" & empNo
        sqlText = sqlText & " Where RepNo=" & .repNo
        sqlText = sqlText & " update maintCall Set CallStatus=" & IIf(IsNull(.CallStatus), "null", .CallStatus) & "  Where callNo=" & .CallNo
ElseIf .ReparationState = DeleteRecord Then
        Set reparationInfo.ReparationPiecesViewModel = New ReparationPiecesViewModel
        sqlText = sqlText & " delete from reparationworks where repno=" & .repNo
        sqlText = sqlText & GetReparationPiecesQuery(reparationInfo)
        sqlText = sqlText & " delete from reparationPieces where repno=" & .repNo
        sqlText = sqlText & " delete from reparation where repno=" & .repNo
        sqlText = sqlText & " update maintCall Set CallStatus=null Where callNo=" & .CallNo
End If
End With
GetReparationInfoQuery = sqlText
End Function

Function CheckBalanceOfThePircesByTeamNo(reparationInfo As ReparationViewModel) As Boolean

'Check balance  from the ItemNo and Team
    Dim rsStkNoTeamBalance As New ADODB.Recordset
    For deleteStkNo = 1 To reparationPiecesCollection.Count
        sqlText = "select isnull(sum(fnlqnt),0)fnlqnt from MaintTeam m1 inner join stkinf s1 on s1.StrNo = m1.Empno where stkno='" & reparationPiecesCollection(deleteStkNo) & "' and m1.teamno=" & reparationInfo.TeamNo
        Set rsStkNoTeamBalance = de.con.Execute(sqlText)
        If rsStkNoTeamBalance!fnlqnt < 0 Then
            CheckBalanceOfThePircesByTeamNo = False
            Exit Function
        Else
            CheckBalanceOfThePircesByTeamNo = True
            Exit Function
        End If
        
    Next deleteStkNo
    CheckBalanceOfThePircesByTeamNo = True
End Function

Public Function ReparationValidResult(reparationInfo As ReparationViewModel) As Collection
On Error GoTo ErrorHandler


Dim result As Collection
Set result = validateResult.GetReparationValidResult(reparationInfo)
Set ReparationValidResult = result
Exit Function
ErrorHandler:
    result.Add (Err.Description)
    Set ReparationValidResult = result
End Function

Public Function SaveChanges(reparationInfo As ReparationViewModel, Optional vState) As ReparationResult
On Error GoTo ErrorHandler
Dim result As New ReparationResult
Dim ReparationValidationResult As Collection
Dim repTransaction As Double
Dim Item As Variant

Dim sqlText As String


Set ReparationValidationResult = ReparationValidResult(reparationInfo)
If ReparationValidationResult.Count > 0 Then
    result.ReparationResultStatus = False
    For Each Item In ReparationValidationResult
        result.ReparationResultDescription = result.ReparationResultDescription & Item & Chr(13)
    Next
    Set SaveChanges = result
    Exit Function
End If
    
If Not result.ReparationResultStatus Then
    Set SaveChanges = result
    Exit Function
End If

sqlText = ""

sqlText = sqlText & GetProductInformationQuery(reparationInfo)
sqlText = sqlText & GetReparationInfoQuery(reparationInfo)
If reparationInfo.ReparationState <> DeleteRecord Then
    sqlText = sqlText & GetReparationWorksQuery(reparationInfo)
    sqlText = sqlText & GetReparationPiecesQuery(reparationInfo)
End If
If sqlText = "" Then
    Set SaveChanges = result
    Exit Function
End If

    repTransaction = de.con.BeginTrans
    de.con.Execute (sqlText)
    If CheckBalanceOfThePircesByTeamNo(reparationInfo) Then
        If repTransaction > 0 Then de.con.CommitTrans
    Else
        If repTransaction > 0 Then de.con.RollbackTrans
        result.ReparationResultDescription = "the Stock Number :" & reparationPiecesCollection(deleteStkNo) & " not enough"
        result.ReparationResultStatus = False
        Set SaveChanges = result
        Exit Function
    End If
    Set SaveChanges = result
    Exit Function
ErrorHandler:
    result.ReparationResultDescription = Err.Description
    result.ReparationResultStatus = False
    If repTransaction > 0 Then de.con.RollbackTrans
    Set SaveChanges = result
End Function

Function ExecProcedure(unPrintedMaintCalls As String) As Boolean
On Error GoTo ErrorHandler
    sqlText = "sp_GetClientInfo '" & unPrintedMaintCalls & "'"
    de.con.Execute (sqlText)
    ExecProcedure = True
Exit Function
ErrorHandler:
ExecProcedure = False
MsgBox Err.Description
End Function

Public Function MaintCallValidResult(maintCallInfo As MaintCallViewModel) As Collection
On Error GoTo ErrorHandler

Dim result As Collection
Set result = maintCallResult.GetMaintCallValidResult(maintCallInfo)
Set MaintCallValidResult = result
Exit Function
ErrorHandler:
    result.Add (Err.Description)
    Set MaintCallValidResult = result
End Function

Function GetClientValidationResult(clientInfo As AdhamViewModel) As Collection
On Error GoTo ErrorHandler

Dim result As Collection
Set result = clientValidationResult.GetClinetValidationResult(clientInfo)
Set GetClientValidationResult = result
Exit Function
ErrorHandler:
    result.Add (Err.Description)
    Set GetClientValidationResult = result
End Function

Function SaveClient(maintCallInfo As MaintCallViewModel) As maintCallResult

On Error GoTo ErrorHandler
Dim result As New maintCallResult
Dim clientValidationResult As Collection
Set clientValidationResult = GetClientValidationResult(maintCallInfo.clientInfo)


If clientValidationResult.Count > 0 Then
    result.MaintCallResultStatus = False
    For Each Item In clientValidationResult
        result.MaintCallResultDescription = result.MaintCallResultDescription & Item & Chr(13)
    Next
    Set SaveClient = result
    Exit Function
End If
    
If Not result.MaintCallResultStatus Then
    Set SaveClient = result
    Exit Function
End If

sqlText = sqlText & GetClientInfoQuery(maintCallInfo)

If sqlText = "" Then
    Set SaveClient = result
    Exit Function
End If


de.con.Execute (sqlText)
Set SaveClient = result


Exit Function

ErrorHandler:
    result.MaintCallResultDescription = Err.Description
    result.MaintCallResultStatus = False

    Set SaveClient = result

End Function

Function SaveMaintCallChanges(maintCallInfo As MaintCallViewModel, vState As EnumState, choice As Boolean) As maintCallResult
On Error GoTo ErrorHandler
Dim result As New maintCallResult
Dim MaintCallValidationResult As Collection
Dim maintCallTransaction As Double
Dim Item As Variant

Dim sqlText As String

If vState <> DeleteRecord Then
    Set MaintCallValidationResult = MaintCallValidResult(maintCallInfo)
Else
   Set MaintCallValidationResult = New Collection
End If

If MaintCallValidationResult.Count > 0 Then
    result.MaintCallResultStatus = False
    For Each Item In MaintCallValidationResult
        result.MaintCallResultDescription = result.MaintCallResultDescription & Item & Chr(13)
    Next
    Set SaveMaintCallChanges = result
    Exit Function
End If
    
If Not result.MaintCallResultStatus Then
    Set SaveMaintCallChanges = result
    Exit Function
End If

sqlText = ""
sqlText = sqlText & GetClientInfoQuery(maintCallInfo)
sqlText = sqlText & ";" & GetMaintCallInfoQuery(maintCallInfo, vState, choice)

If sqlText = "" Then
    Set SaveMaintCallChanges = result
    Exit Function
End If

maintCallTransaction = de.con.BeginTrans
 de.con.Execute (sqlText)
If maintCallInfo.CallNo = 0 Then
    maintCallInfo.CallNo = getMaxClaimId()
End If
Set SaveMaintCallChanges = result
If maintCallTransaction > 0 Then de.con.CommitTrans

Exit Function

ErrorHandler:
    result.MaintCallResultDescription = Err.Description
    result.MaintCallResultStatus = False
    If maintCallTransaction > 0 Then de.con.RollbackTrans
    Set SaveMaintCallChanges = result
End Function

Sub ChangeIsCheckMaintCallNos(CallNos As String, chkValue As Boolean)
On Error GoTo ErrorHandler
If chkValue Then
    sqlText = "Update MaintCall Set IsCheck=1 where CallNo in (" & CallNos & ")"
Else
    sqlText = "Update MaintCall Set IsCheck=0 where CallNo in (" & CallNos & ")"
End If
de.con.Execute (sqlText)
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub ChangeIsCheckMaintCall(CallNo As Long, chkValue As Boolean)
On Error GoTo ErrorHandler
If chkValue Then
    sqlText = "Update MaintCall Set IsCheck=1 where CallNo=" & CallNo
Else
    sqlText = "Update MaintCall Set IsCheck=0 where CallNo=" & CallNo
End If
de.con.Execute (sqlText)
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Function getMaxClaimId() As Double
On Error GoTo ErrorHandler
Dim rsMaxClaimId As New ADODB.Recordset
sqlText = "Select Max(CallNo)MaxClaimId From MaintCall"
Set rsMaxClaimId = de.con.Execute(sqlText)
getMaxClaimId = rsMaxClaimId!MaxClaimId
Exit Function
ErrorHandler:
getMaxClaimId = -1
MsgBox Err.Description
End Function
Function GenerateBarcodeNo(reparationProductViewModel As AdhamProductViewModel) As String
On Error GoTo ErrorHandler
    With reparationProductViewModel
        If .GasNo = "" Or .ItemNo = "" Or .BarcodeSerialNo = "" Or Not IsDate(.ProdDate) Or IsNull(.ProdDate) Then
            GenerateBarcodeNo = ""
        Else
            GenerateBarcodeNo = Trim(.GasNo) & Trim(.ItemNo) & Right(Trim(.ProdDate), 2) & Mid(Trim(.ProdDate), 4, 2) & left(Trim(.ProdDate), 2) & Trim(.BarcodeSerialNo)
        End If
    End With
Exit Function
ErrorHandler:
GenerateBarcodeNo = ""
End Function

Public Function GetAllReparations() As ADODB.Recordset
On Error GoTo ErrorHandler
Dim Vindex As Double
Dim rs As New ADODB.Recordset
Dim sqlText As String

sqlText = "select CallNo , RepNo From Reparation order by RepNo"
Set rs = de.con.Execute(sqlText)
Set GetAllReparations = rs

Exit Function
ErrorHandler:
MsgBox Err.Description
Set GetAllReparations = rs
End Function



Public Function GetAllMaintCall() As ADODB.Recordset
On Error GoTo ErrorHandler
Dim Vindex As Double
Dim rsMaintCall As New ADODB.Recordset
Dim sqlText As String

sqlText = "select  CallNo   From MaintCall Order By CallNo "
Set rsMaintCall = de.con.Execute(sqlText)
Set GetAllMaintCall = rsMaintCall

Exit Function
ErrorHandler:
MsgBox Err.Description
Set GetAllMaintCall = rsMaintCall
End Function

Public Function GetTotalFeesReparation(reparationInfo As ReparationViewModel) As Double
On Error GoTo ErrorHandler
Dim vSum As Double
vSum = 0
If reparationInfo.ReparationPiecesViewModel.Count <> 0 Then
    For i = 1 To reparationInfo.ReparationPiecesViewModel.Count
     vSum = vSum + (reparationInfo.ReparationPiecesViewModel.Item(i).Price * reparationInfo.ReparationPiecesViewModel.Item(i).Qty)
    Next
Else
End If
GetTotalFeesReparation = vSum
Exit Function
ErrorHandler:
MsgBox Err.Description
GetTotalFeesReparation = vSum
End Function

Function GetDefaultGasNo(ProdFamNo As Integer) As Integer

Dim famillyListWithGas() As Integer
 famillyListWithGas = GetFamillyListWithGas()
For i = LBound(famillyListWithGas) To UBound(famillyListWithGas)
    If ProdFamNo = famillyListWithGas(i) Then
        GetDefaultGasNo = GasNoForSpecialFamilies
        Exit Function
    End If
Next i
GetDefaultGasNo = GasNoForAllFamilies

End Function

Function GetMaintCallState(vState As EnumMaintCallState) As String

Select Case vState
    Case UnderwayAndPrinted:
        GetMaintCallState = "ﬁÌœ «· ‰›Ì– Ê «·ÿ»«⁄Â"
    Case Repared:
        GetMaintCallState = "«·‘ﬂÊÏ „‰›–Â"
    
    Case Else
        GetMaintCallState = "«·‘ﬂÊÏ „ƒﬁ Â"
End Select

End Function

Function GetReparationState(vState As EnumState) As String

Select Case vState
    Case DefaultRecord:
        GetReparationState = "⁄‹—÷ «·«’·«Õ"
    Case ReadOnly:
        GetReparationState = "«·«’·«Õ „—Õ·"
    Case NewRecord:
        GetReparationState = "≈’·«Õ ÃœÌœ"
    Case UpdateRecord:
        GetReparationState = "Ì „  ⁄œÌ· «·«’·«Õ"
    Case Else
        GetReparationState = "€Ì— „⁄—Ê›"
End Select

End Function

'UpdateREparationState (repNo)
Public Function GetReparationByCallNo(CallNo As Double) As ReparationViewModel
On Error GoTo ErrorHandler
Dim rsReparationInfo As New ADODB.Recordset
Dim reperationInfo As New ReparationViewModel
Dim sqlText As String

sqlText = "select RepNo, m1.CallNo, m1.ModNo , m3.TeamNo, m4.teamname, n1.id StrId , n1.strno ,  r1.ProdNo,   RepDate , CallDateTime, RepTimeBegin, RepTimeEnd, RepPrice, Cash, r1.regestdate, r1.Notes, Carried, CliRecever, Description, VoltBefor, VoltAfter , m1.CallStatus , m1.CliNo , adhamName as ClientName , isnull(r1.EmpNo,0)EmpNo , isnull(e1.FullName,'')FullName   from"
'sqlText = sqlText & "     maintcallOrder m1  inner join"
sqlText = sqlText & "     maintcall m1  left outer join"
sqlText = sqlText & "     MntMoveSummary m3 on m1.callno = m3.CallNo left outer join"
sqlText = sqlText & "     MaintTeam m4 on m3.TeamNo = m4.TeamNo left outer join "
sqlText = sqlText & "     namestr n1 on m4.empno = n1.strno left outer join "
sqlText = sqlText & "     reparation r1 on m1.callno = r1.callno left outer join"
sqlText = sqlText & "     _source.dbo.EmpFullName e1 on r1.EmpNo = e1.EmpNo left outer join"
'sqlText = sqlText & "     adhammodels a2 on a1.modno = a2.modno left outer join"
sqlText = sqlText & "     AdhamView7 a3 on m1.clino = a3.adhamno where  m1.callno= " & CallNo
    
Set rsReparationInfo = de.con.Execute(sqlText)

If Not rsReparationInfo.EOF Then
    With reperationInfo
       ' .ReparationState = vState
        If IsNull(rsReparationInfo!repNo) Then
            .ReparationState = NewRecord
        End If
        .empNo = rsReparationInfo!empNo
        .FullName = rsReparationInfo!FullName
        .CallNo = rsReparationInfo!CallNo
        .ClientName = rsReparationInfo!ClientName
        .TeamNo = IIf(IsNull(rsReparationInfo!TeamNo), 0, rsReparationInfo!TeamNo)
        .TeamName = rsReparationInfo!TeamName & ""
        .StrNo = rsReparationInfo!StrNo
        .Strid = rsReparationInfo!Strid
        .CallDateTime = Format(rsReparationInfo!CallDateTime, "dd/mm/yyyy")
'        .EntryDate = Format(rsReparationInfo!regestdate, "dd/mm/yyyy")
        If IsDate(rsReparationInfo!RepDate) Then
            .RepDate = Format(rsReparationInfo!RepDate, "dd/mm/yyyy")
        Else
            .RepDate = Format(rsReparationInfo!CallDateTime, "dd/mm/yyyy")
        End If
        If Not IsNull(rsReparationInfo!ModNo) And .ReparationState = NewRecord Then
            Set .AdhamProductViewModel = GetProductViewModel(, rsReparationInfo!cliNo, rsReparationInfo!ModNo)
        End If
        If .ReparationState <> NewRecord Then
            .repNo = rsReparationInfo!repNo
            .ProdNo = IIf(IsNull(rsReparationInfo!ProdNo), Null, rsReparationInfo!ProdNo)
            .RepPrice = IIf(IsNull(rsReparationInfo!RepPrice), 0, rsReparationInfo!RepPrice)
            .RepTimeBegin = Format(rsReparationInfo!RepTimeBegin, "HH:mm")
            .RepTimeEnd = Format(rsReparationInfo!RepTimeEnd, "HH:mm")
            .Cash = IIf(IsNull(rsReparationInfo!Cash), -1, rsReparationInfo!Cash)
            .regestdate = Format(rsReparationInfo!regestdate, "dd/mm/yyyy")
            .Notes = Replace(rsReparationInfo!Notes & "", vbNewLine, "")
             If IsNull(rsReparationInfo!Carried) Or rsReparationInfo!Carried = 0 Then
                .Carried = False
'                .ReparationState = DefaultRecord
             Else
                .Carried = True
                .ReparationState = ReadOnly
             End If
            
            .CliRecever = IIf(IsNull(rsReparationInfo!CliRecever), -1, rsReparationInfo!CliRecever)
            .CallStatus = IIf(IsNull(rsReparationInfo!CallStatus), -1, rsReparationInfo!CallStatus)
            .Description = Replace(rsReparationInfo!Description & "", vbNewLine, "")
            .VoltBefor = IIf(IsNull(rsReparationInfo!VoltBefor), 0, rsReparationInfo!VoltBefor)
            .VoltAfter = IIf(IsNull(rsReparationInfo!VoltAfter), 0, rsReparationInfo!VoltAfter)
            Set .AdhamProductViewModel = GetProductViewModel(rsReparationInfo!repNo)
            Set .ReparationWorksViewModel = GetReparationWorksByRepNo(rsReparationInfo!repNo)
            Set .ReparationPiecesViewModel = GetReparationPiecesByRepNo(rsReparationInfo!repNo)
        End If
    End With
    Set GetReparationByCallNo = reperationInfo
End If
Exit Function
ErrorHandler:
MsgBox Err.Description
Set GetReparationByCallNo = reperationInfo
End Function

Public Function GetClientInfoById(clientId As Long) As AdhamViewModel
On Error GoTo ErrorHandler

Dim clientInfo As New AdhamViewModel
Dim rsClientInfo As New ADODB.Recordset
sqlText = "Select AdNo , Adhamno , adhamname ,"
sqlText = sqlText & " AdhamPhon , MobilePhone, AdhamAdress , Zone , ZoneName , Notes , Defindname  from AdhamView7 c1 left outer join CoZone c2 on c1.zone = c2.zoneNo  where Adhamno=" & clientId
Set rsClientInfo = de.con.Execute(sqlText)
If rsClientInfo.RecordCount > 0 Then
    With clientInfo
        .AdNo = rsClientInfo!AdNo
        .adhamNo = rsClientInfo!adhamNo
        .AdhamName = rsClientInfo!AdhamName & ""
        .AdhamAdress = rsClientInfo!AdhamAdress & ""
        .AdhamPhon = rsClientInfo!AdhamPhon & ""
        .MobilePhone = rsClientInfo!MobilePhone & ""
        .Zone = rsClientInfo!Zone
        .ZoneName = rsClientInfo!ZoneName
        .Notes = rsClientInfo!Notes & ""
        .Defindname = rsClientInfo!Defindname & ""
        .ClientState = Default
    End With
    Set GetClientInfoById = clientInfo
End If
Exit Function
ErrorHandler:
Set GetClientInfoById = clientInfo
MsgBox Err.Description
End Function
Function NumberingTheSelectedOrders(selectedClaimsWithoutNumber As String) As Boolean
On Error GoTo ErrorHandler
Dim sqlText As String
sqlText = "Insert into MaintCall(Id_ComNo , CallOrderId , CallState)"
sqlText = sqlText & " Select 1, m1.Id , 0 From MaintCallOrder m1  left outer join maintcall m2 on m1.Id = m2.CallOrderId Where m2.CallOrderId is null and m1.Id in(" & selectedClaimsWithoutNumber & ")"
de.con.Execute (sqlText)
NumberingTheSelectedOrders = True
Exit Function
ErrorHandler:
NumberingTheSelectedOrders = False
MsgBox Err.Description
End Function

Public Function GetTeamNameForRepeationClaim(cliNo As Double, callDate As Date) As String
On Error GoTo ErrorHandler
Dim rsTeamName As New ADODB.Recordset
Dim sqlText As String
sqlText = ""

sqlText = sqlText & " select top 1 TeamName from "
sqlText = sqlText & " maintcall m1 inner join"
sqlText = sqlText & " MntMoveSummary m2 on m1.callno = m2.callno inner join"
sqlText = sqlText & " maintteam m3 on m2.TeamNo = m3.TeamNo"
sqlText = sqlText & " where clino = " & cliNo & " and convert(varchar(10),m1.CallDateTime,101) ='" & ConvertControlDate(callDate) & "'"

Set rsTeamName = de.con.Execute(sqlText)
If rsTeamName.RecordCount > 0 Then
    GetTeamNameForRepeationClaim = rsTeamName!TeamName
Else
    GetTeamNameForRepeationClaim = ""
End If
Exit Function
ErrorHandler:
GetTeamNameForRepeationClaim = ""

End Function



Public Function claimsCountForTheClient(cliNo As Double, callDate As Date, maintCallState As EnumState) As Boolean
On Error GoTo ErrorHandler
Dim rsRepeatClaims As New ADODB.Recordset
Dim sqlText As String
sqlText = ""

sqlText = sqlText & "    select count(*) as countMaintCall from maintcall m1 where clino = " & cliNo & " and convert(varchar(10),CallDateTime,101) = '" & ConvertControlDate(callDate) & "'"

Set rsRepeatClaims = de.con.Execute(sqlText)
 If rsRepeatClaims!countMaintCall = 1 And maintCallState = DefaultRecord Then
    claimsCountForTheClient = False
 
ElseIf rsRepeatClaims!countMaintCall >= 1 Then  'And (maintCallState = NewRecord Or maintCallState = UpdateRecord) Then
  claimsCountForTheClient = True
Else
    claimsCountForTheClient = False
End If

Exit Function
ErrorHandler:
claimsCountForTheClient = False
End Function


Public Function GetMaintCallInfo(CallNo As Double) As MaintCallViewModel
On Error GoTo ErrorHandler
Dim rsMaintCallInfo As New ADODB.Recordset
Dim maintCallInfo As New MaintCallViewModel
Dim sqlText As String

sqlText = ""
sqlText = sqlText & "     Select "
sqlText = sqlText & "            m1.CallNo ,"
sqlText = sqlText & "           CallDateTime ,"
sqlText = sqlText & "           Case When r1.callNo is not null then 2 else m1.CallState end CallState ,"
sqlText = sqlText & "           m1.ModNo ,"
sqlText = sqlText & "           m1.CliNo  ,"
sqlText = sqlText & "           CallDescription ,"
'sqlText = sqlText & "           m1.Notes ,"
'sqlText = sqlText & "           m1.DefindName  ,"
sqlText = sqlText & "           m1.CallReceiverEmpNo ,"
sqlText = sqlText & "           e2.FullName  as CallReceiverName,"
sqlText = sqlText & "           a1.AdNo ,"
sqlText = sqlText & "           a1.Adhamno ,"
sqlText = sqlText & "           a1.adhamname ,"
sqlText = sqlText & "           a1.AdhamPhon ,"
sqlText = sqlText & "           MobilePhone,"
sqlText = sqlText & "           AdhamAdress ,"
sqlText = sqlText & "           a1.Zone ,"
sqlText = sqlText & "           c1.ZoneName,"
sqlText = sqlText & "           a1.Notes as ClientNote,"
sqlText = sqlText & "           a1.Defindname as ClientDefindname,"
'sqlText = sqlText & "           m1.Notes as Note,"
'sqlText = sqlText & "           m1.Defindname Defindname   ,"
sqlText = sqlText & "           LastVisitDate = (select Max(rep.repdate) from reparation rep  inner join maintcall mc on rep.CallNo = mc.CallNo   where m1.clino = mc.clino    and m1.callNo >isnull(r1.CallNo,0) ),"
sqlText = sqlText & "           LastCallNo = (select Max(rep.CallNo) from reparation rep   inner join maintcall mc on rep.CallNo = mc.CallNo   where m1.clino = mc.clino and   m1.callNo >isnull(r1.CallNo,0) )"
sqlText = sqlText & "          From"
sqlText = sqlText & "               maintcall m1   inner join"
sqlText = sqlText & "               empfullname e2 on m1.CallReceiverEmpNo = e2.Empno inner join"
sqlText = sqlText & "               AdhamView7 a1 on m1.clino = a1.AdhamNo inner join"
sqlText = sqlText & "               CoZone c1 on a1.Zone = c1.ZoneNo left outer join"
sqlText = sqlText & "               Reparation r1 on m1.callno = r1.callno"
sqlText = sqlText & "           Where m1.callNo = " & CallNo


    
Set rsMaintCallInfo = de.con.Execute(sqlText)

If Not rsMaintCallInfo.EOF Then
    With maintCallInfo
 '       .Id = rsMaintCallInfo!Id
        .CallNo = rsMaintCallInfo!CallNo
        .callState = IIf(IsNull(rsMaintCallInfo!callState), 0, rsMaintCallInfo!callState)
        .CallDateTime = rsMaintCallInfo!CallDateTime
        .ModNo = rsMaintCallInfo!ModNo
        .cliNo = rsMaintCallInfo!cliNo
        .ClientVisitorCount = GetClientVisitorCount(.cliNo, .CallDateTime)
        .ClaimIsRepeat = claimsCountForTheClient(.cliNo, .CallDateTime, EnumState.DefaultRecord) ' IIf(claimsCountForTheClient(.cliNo, .CallDateTime) >= 1, True, False)
        .ClaimTeamName = GetTeamNameForRepeationClaim(.cliNo, .CallDateTime)
        .LastVisitDate = rsMaintCallInfo!LastVisitDate
        .LastCallNo = rsMaintCallInfo!LastCallNo
        .clientInfo.adhamNo = rsMaintCallInfo!adhamNo
        .clientInfo.AdhamName = rsMaintCallInfo!AdhamName & ""
        .clientInfo.AdhamAdress = rsMaintCallInfo!AdhamAdress & ""
        .clientInfo.AdhamPhon = rsMaintCallInfo!AdhamPhon & ""
        .clientInfo.MobilePhone = rsMaintCallInfo!MobilePhone & ""
        .clientInfo.Zone = rsMaintCallInfo!Zone
        .clientInfo.ZoneName = rsMaintCallInfo!ZoneName & ""
        .clientInfo.Defindname = rsMaintCallInfo!ClientDefindname & ""
        .clientInfo.Notes = rsMaintCallInfo!ClientNote & ""
        .clientInfo.ClientState = Default
        .CallDEscription = rsMaintCallInfo!CallDEscription & ""
'        .Notes = rsMaintCallInfo!Notes & ""
'        .Defindname = rsMaintCallInfo!Defindname & ""
        .CallReceiverEmpNo = rsMaintCallInfo!CallReceiverEmpNo
        .CallReceiverName = rsMaintCallInfo!CallReceiverName
    End With
    Set GetMaintCallInfo = maintCallInfo
End If
Exit Function
ErrorHandler:
MsgBox Err.Description
Set GetMaintCallInfo = maintCallInfo
End Function


Function GetProductViewModel(Optional repNo, Optional cliNo, Optional prodFamily) As AdhamProductViewModel

Dim sqlText As String
    Dim rsAdhamProductViewModelInfo As New ADODB.Recordset
    
    
'    Dim Vindex As Integer
    Dim adhamProductViewModelInfo As New AdhamProductViewModel
    sqlText = "select top 1 m1.ModNo ProdFamNo,  a2.ModNo  , a2.Symbol , a2.ItemNo,  a1.ProdNo , a3.ProdFamNameA , ProdBaCodeNo, ProdDate , ProdPurchaseDate , a1.GasNo, a1.BarcodeSerNo   from"
    sqlText = sqlText & " MaintCall m1 inner join"
'    sqlText = sqlText & " maintcallOrder m1 on m1.Id = m2.CallOrderId left outer join"
    sqlText = sqlText & " AdhamProductFamily a3 on m1.ModNo = a3.ProdFamNo left outer join"
    sqlText = sqlText & " reparation r1 on m1.callno = r1.callno left outer join"
    sqlText = sqlText & " AdhamProducts a1 on a1.ProdNo = r1.ProdNo left outer join"
    sqlText = sqlText & " adhammodels a2 on a1.modno = a2.modno"
    If Not IsMissing(cliNo) And Not IsMissing(prodFamily) Then
        sqlText = sqlText & " Where m1.CliNo=" & cliNo & " and m1.ModNo=" & prodFamily
        sqlText = sqlText & " Order by a1.prodno desc "
    ElseIf Not IsMissing(repNo) Then
        sqlText = sqlText & " Where r1.RepNo=" & repNo
    End If
    
    Set rsAdhamProductViewModelInfo = de.con.Execute(sqlText)
    If Not rsAdhamProductViewModelInfo.EOF Then
        With adhamProductViewModelInfo
            If Not IsMissing(cliNo) And Not IsMissing(prodFamily) Then
                .ProdNo = Null
            Else
                .ProdNo = rsAdhamProductViewModelInfo!ProdNo
            End If
            .ProdFamName = rsAdhamProductViewModelInfo!ProdFamNameA & ""
            .ProdFamNo = IIf(IsNull(rsAdhamProductViewModelInfo!ProdFamNo), 0, rsAdhamProductViewModelInfo!ProdFamNo)
            .ModNo = IIf(IsNull(rsAdhamProductViewModelInfo!ModNo), 0, rsAdhamProductViewModelInfo!ModNo)
            .Symbol = rsAdhamProductViewModelInfo!Symbol & ""
            .ItemNo = rsAdhamProductViewModelInfo!ItemNo & ""
            .ProdBaCodeNo = rsAdhamProductViewModelInfo!ProdBaCodeNo & ""
            .GasNo = rsAdhamProductViewModelInfo!GasNo & ""
            .BarcodeSerialNo = rsAdhamProductViewModelInfo!BarcodeSerNo & ""
'            If Not IsNull(rsAdhamProductViewModelInfo!ProdBaCodeNo) And rsAdhamProductViewModelInfo!ProdBaCodeNo <> "" Then
'                    .GasNo = left(Trim(rsAdhamProductViewModelInfo!ProdBaCodeNo), 1)
'                ElseIf Not IsNull(rsAdhamProductViewModelInfo!ProdFamNo) Then
'                    .GasNo = GetDefaultGasNo(rsAdhamProductViewModelInfo!ProdFamNo)
'                Else
'                    .GasNo = ""
'            End If
'            If rsAdhamProductViewModelInfo!ProdBaCodeNo <> "" And (Len(Trim(rsAdhamProductViewModelInfo!ProdBaCodeNo)) = 16 Or Len(Trim(rsAdhamProductViewModelInfo!ProdBaCodeNo)) = 18) Then
'                .BarcodeSerialNo = Right(Trim(rsAdhamProductViewModelInfo!ProdBaCodeNo), 3)
'            Else
'                .BarcodeSerialNo = ""
'            End If
            If IsDate(rsAdhamProductViewModelInfo!ProdDate) Then
                .ProdDate = Format(rsAdhamProductViewModelInfo!ProdDate, "dd/mm/yyyy")
            Else
                .ProdDate = Null
            End If
            If IsDate(rsAdhamProductViewModelInfo!ProdPurchaseDate) Then
                .ProdPurchaseDate = Format(rsAdhamProductViewModelInfo!ProdPurchaseDate, "dd/mm/yyyy")
            Else
                .ProdPurchaseDate = Null
            End If
        End With
    End If
    Set GetProductViewModel = adhamProductViewModelInfo
End Function

Private Function GetReparationWorksByRepNo(repNo As Double) As ReparationWorksViewModel
Dim sqlText As String
    Dim rsReparationWorkInfo As New ADODB.Recordset
'    Dim Vindex As Integer
    Dim reparationWorksInfo As New ReparationWorksViewModel
    sqlText = "select  R1.RepTypeNo , C1.reptypedescription From ReparationWorks R1 INNER JOIN CoReparationType C1 ON R1.RepTypeNo = C1.reptypeno Where R1.RepNo=" & repNo
    Set rsReparationWorkInfo = de.con.Execute(sqlText)
    If Not rsReparationWorkInfo.EOF Then
        rsReparationWorkInfo.MoveFirst
'            Vindex = 1
            Do While Not rsReparationWorkInfo.EOF
                With reparationWorksInfo
                    .Add rsReparationWorkInfo!RepTypeNo, rsReparationWorkInfo!RepTypeDescription, DefaultRecord, rsReparationWorkInfo!RepTypeNo
'                    Vindex = Vindex + 1
                End With
                rsReparationWorkInfo.MoveNext
            Loop
    End If
    Set GetReparationWorksByRepNo = reparationWorksInfo
End Function

Private Function GetReparationPiecesByRepNo(repNo As Double) As ReparationPiecesViewModel
Dim sqlText As String
'    Dim Vindex As Integer
    Dim rsReparationPircesInfo As New ADODB.Recordset
    Dim reparationPircesInfo As New ReparationPiecesViewModel
    sqlText = "select SerId , RepNo , PieceNo , id as stkid ,  STKNAME ,  Price , Qty  From reparationPieces R1 INNER JOIN CoStock C1 ON R1.PieceNo = C1.STKNO COLLATE ARABIC_CI_AS Where RepNo=" & repNo
    Set rsReparationPircesInfo = de.con.Execute(sqlText)
    If Not rsReparationPircesInfo.EOF Then
        rsReparationPircesInfo.MoveFirst
'            Vindex = 1
            Do While Not rsReparationPircesInfo.EOF
                With reparationPircesInfo
                    .Add rsReparationPircesInfo!stkId, rsReparationPircesInfo!PieceNo, rsReparationPircesInfo!Qty, rsReparationPircesInfo!Price, rsReparationPircesInfo!StkName, rsReparationPircesInfo!SerId, DefaultRecord, rsReparationPircesInfo!PieceNo
'                    Vindex = Vindex + 1
                End With
                rsReparationPircesInfo.MoveNext
            Loop
    End If
    Set GetReparationPiecesByRepNo = reparationPircesInfo
End Function

Public Function Exists(ByVal oCol As Collection, ByVal vKey As Variant) As Boolean

    On Error Resume Next
    oCol.Item vKey
    Exists = (Err.Number = 0)
    Err.Clear
    
End Function

Public Function ReparationWorkTypeNoExists(ByVal oCol As ReparationWorksViewModel, ByVal vKey As Variant) As Boolean
    Dim o As Object
    Set o = oCol.Item(vKey)
    If o Is Nothing Then
        ReparationWorkTypeNoExists = False
    Else
        ReparationWorkTypeNoExists = True
    End If
End Function

Public Function ReparationPieceExists(ByVal oCol As ReparationPiecesViewModel, ByVal vKey As Variant) As Boolean
    Dim o As Object
    Set o = oCol.Item(vKey)
    If o Is Nothing Then
        ReparationPieceExists = False
    Else
        ReparationPieceExists = True
    End If
End Function
