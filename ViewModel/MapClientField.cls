VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapClientField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFiledId As String 'local copy
Private mvarFiledName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarIsVisible As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarOrder As Integer 'local copy
Private mvarIsListField As Boolean 'local copy
Private mvarIsListColumn As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarIsSearchable As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarIsNumeric As Boolean 'local copy
Public Property Let IsNumeric(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsNumeric = 5
    mvarIsNumeric = vData
End Property


Public Property Get IsNumeric() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsNumeric
    IsNumeric = mvarIsNumeric
End Property



Public Property Let IsSearchable(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsSearchable = 5
    mvarIsSearchable = vData
End Property


Public Property Get IsSearchable() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsSearchable
    IsSearchable = mvarIsSearchable
End Property



Public Property Let IsListColumn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsListColumn = 5
    mvarIsListColumn = vData
End Property


Public Property Get IsListColumn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsListColumn
    IsListColumn = mvarIsListColumn
End Property



Public Property Let IsListField(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsListField = 5
    mvarIsListField = vData
End Property


Public Property Get IsListField() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsListField
    IsListField = mvarIsListField
End Property



Public Property Let Order(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Order = 5
    mvarOrder = vData
End Property


Public Property Get Order() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Order
    Order = mvarOrder
End Property



Public Property Let IsVisible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsVisible = 5
    mvarIsVisible = vData
End Property


Public Property Get IsVisible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsVisible
    IsVisible = mvarIsVisible
End Property



Public Property Let FiledName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FiledName = 5
    mvarFiledName = vData
End Property


Public Property Get FiledName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FiledName
    FiledName = mvarFiledName
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

 Private Sub Class_Initialize()
        Me.filedId = ""
        Me.FiledName = ""
        Me.IsListColumn = False
        Me.IsListField = False
        Me.IsVisible = False
        Me.Order = 0
        Me.IsSearchable = True
 End Sub



