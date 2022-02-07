VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaiResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarResultDescription As String 'local copy
Private mvarResultState As Boolean 'local copy
Public Property Let ResultState(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResultState = 5
    mvarResultState = vData
End Property


Public Property Get ResultState() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResultState
    ResultState = mvarResultState
End Property

Public Property Let ResultDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResultDescription = 5
    mvarResultDescription = vData
End Property


Public Property Get ResultDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResultDescription
    ResultDescription = mvarResultDescription
End Property

Private Sub Class_initialize()
    ResultDescription = ""
    ResultState = True
End Sub
