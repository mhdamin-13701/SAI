VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFiledId As String 'local copy
Private mvarColHeaderName As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarisChecked As Boolean 'local copy
Public Property Let isChecked(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.isChecked = 5
    mvarisChecked = vData
End Property


Public Property Get isChecked() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.isChecked
    isChecked = mvarisChecked
End Property




Public Property Let colHeaderName(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarColHeaderName = vData
End Property


Public Property Get colHeaderName() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    colHeaderName = mvarColHeaderName
End Property



Public Property Let filedId(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FiledId = 5
    mvarFiledId = vData
End Property


Public Property Get filedId() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FiledId
    filedId = mvarFiledId
End Property





