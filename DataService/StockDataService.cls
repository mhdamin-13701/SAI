VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StockDataService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Function GetBalanceItemsByStkNoAndTeamNo(stkno As String, TeamNo As Integer) As Double
On Error GoTo ErrorHandler
    Dim RsBalance As New ADODB.Recordset
    Dim sqlText As String
    
    sqlText = "select s1.FnlQnt     from"
    sqlText = sqlText & " Stkinf s1 inner join"
    sqlText = sqlText & " namestr n1 on s1.StrNo = n1.strno inner join"
    sqlText = sqlText & " maintTeam m1 on n1.strno = m1.Empno inner join"
    sqlText = sqlText & " CoStock c1 on s1.Stkid = c1.Id"
    sqlText = sqlText & " Where c1.StkNo='" & stkno & "' and teamno=" & TeamNo
    Set RsBalance = de.con.Execute(sqlText)
    
    If Not RsBalance.EOF Then
        GetBalanceItemsByStkNoAndTeamNo = RsBalance!fnlqnt
    Else
        GetBalanceItemsByStkNoAndTeamNo = 0
    End If
Exit Function
ErrorHandler:
GetBalanceItemsByStkNoAndTeamNo = 0
MsgBox Err.Description

End Function

Function GetItemNumberDealerPriceWithDiscount(stkno As String) As Double
On Error GoTo ErrorHandler
    Dim RsPrice As New ADODB.Recordset
    Dim sqlText As String
    sqlText = "select  dealpriceafterdiscount from pieces where piecestockno='" & stkno & "'"
   
    Set RsPrice = de.con.Execute(sqlText)
    
    If Not RsPrice.EOF Then
        GetItemNumberDealerPriceWithDiscount = RsPrice!dealpriceafterdiscount
    Else
        GetItemNumberDealerPriceWithDiscount = 0
    End If
Exit Function
ErrorHandler:
GetItemNumberDealerPriceWithDiscount = 0
MsgBox Err.Description
End Function

Function GetMaxByanId() As Double
Dim rsMaxByanId As New ADODB.Recordset
Dim sqlText As String
sqlText = "Select Isnull(Max(ByanId),0) As MaxByanId From Stmov"
Set rsMaxByanId = de.con.Execute(sqlText)

If rsMaxByanId!maxByanId = 0 Then
    GetMaxByanId = 1
Else
    GetMaxByanId = rsMaxByanId!maxByanId + 1
End If
End Function

Function GetStkId(stkno) As Double
Dim slqtext As String
Dim rsStkId As New ADODB.Recordset

sqlText = "Select id from costock where stkno='" & stkno & "'"
Set rsStkId = de.con.Execute(sqlText)

If Not rsStkId.EOF Then
    GetStkId = rsStkId!Id
Else
    GetStkId = -1
End If
End Function
