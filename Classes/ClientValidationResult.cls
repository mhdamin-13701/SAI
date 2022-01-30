VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Dim ClientNameIsMissing As String
Dim ClientPhoneNumberIsMissing As String
Dim ClientAddressIsMissing As String
Dim ClientZoneNameIsMissing As String

Private Sub Class_Initialize()
    ClientNameIsMissing = "·«ÌÊÃœ ≈”„ ··“»Ê‰"
    ClientPhoneNumberIsMissing = "·„ Ì „ ≈œŒ«· —ﬁ„ Â« › «·“»Ê‰ «·À«» "
    ClientAddressIsMissing = "·„ Ì „ ≈œŒ«· ⁄‰Ê«‰ «·“»Ê‰"
    ClientZoneNameIsMissing = "·„ Ì „ ≈œŒ«· „‰ÿﬁÂ «·“»Ê‰"
    
End Sub

Function GetClinetValidationResult(clientInfo As AdhamViewModel) As Collection
Dim validationClinetResult As New Collection
With clientInfo

    If LTrim(RTrim(.AdhamName)) = "" Then
        validationClinetResult.Add ClientNameIsMissing
    End If
    
    If IsNull(.AdhamAdress) Or LTrim(RTrim(.AdhamAdress)) = "" Then
        validationClinetResult.Add ClientAddressIsMissing
    End If
    
    If IsNull(.AdhamPhon) Or LTrim(RTrim(.AdhamPhon)) = "" Then
        validationClinetResult.Add ClientPhoneNumberIsMissing
    End If
    
    If IsNull(.Zone) Then
        validationClinetResult.Add ClientZoneNameIsMissing
    End If
    
  
End With

Set GetClinetValidationResult = validationClinetResult
End Function


