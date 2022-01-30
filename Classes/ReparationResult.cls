VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReparationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarReparationResultDescription As String 'local copy
Private mvarReparationResultStatus As Boolean 'local copy
Public Property Let ReparationResultStatus(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReparationResultStatus = 5
    mvarReparationResultStatus = vData
End Property


Public Property Get ReparationResultStatus() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationResultStatus
    ReparationResultStatus = mvarReparationResultStatus
End Property



Public Property Let ReparationResultDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReparationResultDescription = 5
    mvarReparationResultDescription = vData
End Property


Public Property Get ReparationResultDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationResultDescription
    ReparationResultDescription = mvarReparationResultDescription
End Property

Private Sub Class_Initialize()
    ReparationResultDescription = ""
    ReparationResultStatus = True
End Sub


