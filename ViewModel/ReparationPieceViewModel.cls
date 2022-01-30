VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReparationPieceViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarStkNo As String 'local copy
Private mvarQty As Integer 'local copy
Private mvarPrice As Double 'local copy
Private mvarRepNo As Double 'local copy
Private mvarStkName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSerId As Double 'local copy
Private mvarState As EnumState 'local copy
'local variable(s) to hold property value(s)
Private mvarStkid As Double 'local copy
Public Property Let stkId(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Stkid = 5
    mvarStkid = vData
End Property


Public Property Get stkId() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Stkid
    stkId = mvarStkid
End Property




Public Property Let ReparationPieceState(ByVal vData As EnumState)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property


Public Property Get ReparationPieceState() As EnumState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
     ReparationPieceState = mvarState
End Property



Public Property Let SerId(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SerId = 5
    mvarSerId = vData
End Property


Public Property Get SerId() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SerId
    SerId = mvarSerId
End Property



Public Property Let StkName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StkName = 5
    mvarStkName = vData
End Property


Public Property Get StkName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StkName
    StkName = mvarStkName
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



Public Property Let Price(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Price = 5
    mvarPrice = vData
End Property


Public Property Get Price() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Price
    Price = mvarPrice
End Property



Public Property Let Qty(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Qty = 5
    mvarQty = vData
End Property


Public Property Get Qty() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Qty
    Qty = mvarQty
End Property



Public Property Let stkno(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StkNo = 5
    mvarStkNo = vData
End Property


Public Property Get stkno() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StkNo
    stkno = mvarStkNo
End Property



