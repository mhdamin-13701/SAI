VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MaintCallValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Dim MaintCallProdFamillyIsMissing  As String
Dim MaintCallDescriptionIsMissing As String
Dim ClientNameIsMissing As String
Dim ClientPhoneNumberIsMissing As String
Dim ClientAddressIsMissing As String
Dim ClientZoneNameIsMissing As String
Dim MaintCallReceiverName As String

Private Sub Class_Initialize()
    MaintCallProdFamillyIsMissing = "·„ Ì „ ≈œŒ«· «·⁄«∆·Â"
    MaintCallDescriptionIsMissing = "·„ » „ ≈œŒ«· «·⁄ÿ·"
    ClientNameIsMissing = "·«ÌÊÃœ ≈”„ ··“»Ê‰"
    ClientPhoneNumberIsMissing = "·„ Ì „ ≈œŒ«· —ﬁ„ Â« › «·“»Ê‰ «·À«» "
    ClientAddressIsMissing = "·„ Ì „ ≈œŒ«· ⁄‰Ê«‰ «·“»Ê‰"
    ClientZoneNameIsMissing = "·„ Ì „ ≈œŒ«· „‰ÿﬁÂ «·“»Ê‰"
    MaintCallReceiverName = "·„ Ì „ ≈œŒ«· „” ·„ «·‘ﬂÊÏ"
    
End Sub

Function GetMaintCallValidResult(maintCallInfo As MaintCallViewModel) As Collection
Dim validateResult As New Collection
With maintCallInfo
    If IsNull(.ModNo) Then
     validateResult.Add MaintCallProdFamillyIsMissing
    End If
    
    If LTrim(RTrim(.CallDEscription)) = "" Then
        validateResult.Add MaintCallDescriptionIsMissing
    End If
    
    If IsNull(.CallReceiverEmpNo) Then
        validateResult.Add MaintCallReceiverName
    End If

    If LTrim(RTrim(.clientInfo.AdhamName)) = "" Then
        validateResult.Add ClientNameIsMissing
    End If
    
    If IsNull(.clientInfo.AdhamAdress) Or LTrim(RTrim(.clientInfo.AdhamAdress)) = "" Then
        validateResult.Add ClientAddressIsMissing
    End If
    
    If IsNull(.clientInfo.AdhamPhon) Or LTrim(RTrim(.clientInfo.AdhamPhon)) = "" Then
        validateResult.Add ClientPhoneNumberIsMissing
    End If
    
    If IsNull(.clientInfo.Zone) Then
        validateResult.Add ClientZoneNameIsMissing
    End If
    
  
End With

Set GetMaintCallValidResult = validateResult
End Function

