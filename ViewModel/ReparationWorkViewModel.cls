VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReparationWorkViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarRepNo As Double 'local copy
Private mvarRepTypeNo As String 'local copy
'local variable(s) to hold property value(s)
Private mvarRepTypeDescription As String 'local copy
Private mvarState As EnumState 'local copy
Public Property Let ReparationWorkState(ByVal vData As EnumState)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property


Public Property Get ReparationWorkState() As EnumState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
     ReparationWorkState = mvarState
End Property



Public Property Let RepTypeDescription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepTypeDescription = 5
    mvarRepTypeDescription = vData
End Property


Public Property Get RepTypeDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepTypeDescription
    RepTypeDescription = mvarRepTypeDescription
End Property



Public Property Let RepTypeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepTypeNo = 5
    mvarRepTypeNo = vData
End Property


Public Property Get RepTypeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepTypeNo
    RepTypeNo = mvarRepTypeNo
End Property



'Public Property Let repNo(ByVal vData As Double)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.RepNo = 5
'    mvarRepNo = vData
'End Property
'
'
'Public Property Get repNo() As Double
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.RepNo
'    repNo = mvarRepNo
'End Property



