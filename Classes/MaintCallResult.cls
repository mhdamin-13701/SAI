VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MaintCallResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMaintCallResultDescription As String 'local copy
Private mvarMaintCallResultStatus As Boolean 'local copy
Public Property Let MaintCallResultStatus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReparationResultStatus = 5
    mvarMaintCallResultStatus = vData
End Property


Public Property Get MaintCallResultStatus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationResultStatus
    MaintCallResultStatus = mvarMaintCallResultStatus
End Property



Public Property Let MaintCallResultDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReparationResultDescription = 5
    mvarMaintCallResultDescription = vData
End Property


Public Property Get MaintCallResultDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationResultDescription
    MaintCallResultDescription = mvarMaintCallResultDescription
End Property

Private Sub Class_Initialize()
    MaintCallResultDescription = ""
    MaintCallResultStatus = True
End Sub
