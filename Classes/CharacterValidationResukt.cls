VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharacterValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim CharacterCodeIsMissing As String
Dim CharacterCommodityTypeIdIsMissing As String
Dim CharacterDescriptionIsMissing As String


Sub Class_initialize()
CharacterCodeIsMissing = "Character code is missing"
CharacterCommodityTypeIdIsMissing = "Commodity type is missing"
CharacterDescriptionIsMissing = "Description is missing"
End Sub


Function GetCharacterValidationResult(character_ As Character) As Collection
Dim validationResult As New Collection

With character_
    If LTrim(RTrim(.Code)) = "" Then
        validationResult.Add (CharacterCodeIsMissing)
    End If
    If LTrim(RTrim(.Character)) = "" Then
        validationResult.Add (CharacterDescriptionIsMissing)
    End If
    If .CommidityTypeId Is Null Then
        validationResult.Add (CharacterCommodityTypeIdIsMissing)
    End If
End With
Set GetCharacterValidationResult = validationResult
End Function
