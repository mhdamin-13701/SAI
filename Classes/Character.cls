VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Character"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarId As Variant 'local copy
Private mvarCode As Variant 'local copy
Private mvarCommodityTypeId As Variant 'local copy
Private mvarCharacter As Variant 'local copy
Public Property Let Character(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Character = 5
    mvarCharacter = vData
End Property


Public Property Set Character(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Character = Form1
    Set mvarCharacter = vData
End Property


Public Property Get Character() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Character
    If IsObject(mvarCharacter) Then
        Set Character = mvarCharacter
    Else
        Character = mvarCharacter
    End If
End Property



Public Property Let CommodityTypeId(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommodityTypeId = 5
    mvarCommodityTypeId = vData
End Property


Public Property Set CommodityTypeId(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CommodityTypeId = Form1
    Set mvarCommodityTypeId = vData
End Property


Public Property Get CommodityTypeId() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommodityTypeId
    If IsObject(mvarCommodityTypeId) Then
        Set CommodityTypeId = mvarCommodityTypeId
    Else
        CommodityTypeId = mvarCommodityTypeId
    End If
End Property



Public Property Let Code(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Code = 5
    mvarCode = vData
End Property


Public Property Set Code(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Code = Form1
    Set mvarCode = vData
End Property


Public Property Get Code() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Code
    If IsObject(mvarCode) Then
        Set Code = mvarCode
    Else
        Code = mvarCode
    End If
End Property



Public Property Let Id(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Id = 5
    mvarId = vData
End Property


Public Property Set Id(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Id = Form1
    Set mvarId = vData
End Property


Public Property Get Id() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Id
    If IsObject(mvarId) Then
        Set Id = mvarId
    Else
        Id = mvarId
    End If
End Property



