VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaiSharedDataService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Function characterCodeIsFund(characterCode As String) As Boolean
On Error GoTo ErrorHandler
Dim sqltext As String
Dim rs As New ADODB.Recordset
    sqltext = "Select count(*)countRec from Character where code='" & LTrim(RTrim(characterCode)) & "'"
    Set rs = de.con.Execute(sqltext)
    If rs!countRec > 0 Then
        characterCodeIsFund = True
    Else
        characterCodeIsFund = False
    End If
Exit Function
ErrorHandler:
characterCodeIsFund = False
MsgBox Err.Description
End Function

Function SaveOrUpdateCharacter(character_ As Character) As SaiResult
On Error GoTo ErrorHandler
Dim result As New SaiResult
Dim characterValidationResult_ As New CharacterValidationResult
Dim validationResult As Collection
Dim item As Variant
Dim sqltext As String

Set validationResult = characterValidationResult_.GetCharacterValidationResult(character_)
If validationResult.Count > 0 Then
    result.ResultState = False
    For Each item In validationResult
        result.ResultDescription = result.ResultDescription & item & Chr(13)
    Next
    Set SaveOrUpdateCharacter = result
    Exit Function
End If

Dim characterCodeFundValue As Boolean
characterCodeFundValue = characterCodeIsFund(character_.Code)
If characterCodeFundValue Then
    sqltext = "Update Character set CommodityTypeId=" & character_.CommodityTypeId & ",Character='" & character_.Character & "' Where Code='" & character_.Code & "'"
Else
    sqltext = "insert into Character( Code , CommodityTypeId , Character )Values('" & character_.Code & "'," & character_.CommodityTypeId & ",'" & character_.Character & "')"
End If
de.con.Execute (sqltext)
Set SaveOrUpdateCharacter = result

Exit Function
ErrorHandler:
result.ResultDescription = Err.Description
result.ResultState = False
Set SaveOrUpdateCharacter = result
End Function
