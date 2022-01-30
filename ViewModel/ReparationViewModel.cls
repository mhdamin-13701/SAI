VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReparationViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"ReparationPiecesViewModel"
Attribute VB_Ext_KEY = "Member1" ,"ReparationWorksViewModel"
Attribute VB_Ext_KEY = "Member2" ,"ReparationState"
Attribute VB_Ext_KEY = "Member3" ,"AdhamProductViewModel"
'local variable(s) to hold property value(s)
Private mvarCallNo As Double   'local copy
Private mvarRepNo As Double 'local copy
'Private mvarEntryDate As Variant 'local copy
Private mvarReperationDate As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarReparatonPiecesViewModel As ReparationPiecesViewModel 'local copy
Private mvarReparationWorksViewModel As ReparationWorksViewModel 'local copy
'local variable(s) to hold property value(s)
Private mvarRepPrice As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarTeamNo As Integer 'local copy
Private mvarProdNo As Variant 'local copy
Private mvarRepTimeBegin As Variant 'local copy
Private mvarRepTimeEnd As Variant 'local copy
Private mvarCash As Variant 'local copy
Private mvarregestdate As Variant 'local copy
Private mvarNotes As String 'local copy
Private mvarCarried As Boolean 'local copy
Private mvarCliRecever As Variant 'local copy
Private mvarDescription As String 'local copy
Private mvarVoltBefor As Integer 'local copy
Private mvarVoltAfter As Integer 'local copy
'local variable(s) to hold property value(s)
'Private mvarSymbol As String 'local copy
'Private mvarProdBaCodeNo As String 'local copy
'Private mvarProdDate As Variant 'local copy
'Private mvarProdPurchaseDate As Variant 'local copy
'Private mvarModNo As Integer 'local copy
'local variable(s) to hold property value(s)
'Private mvarGasNo As String 'local copy
'local variable(s) to hold property value(s)
'Private mvarBarcodeSerialNo As String 'local copy
'local variable(s) to hold property value(s)
'Private mvarItemNo As String 'local copy
'local variable(s) to hold property value(s)
Private mvarClientName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarCallStatus As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarState As EnumState 'local copy
'local variable(s) to hold property value(s)
'Private mvarProdFamName As String 'local copy
Private mvarTeamName As String 'local copy

Const DefaultVoltBefore = 190
Const DefaultVoltAfter = 210
Const DefaultRepPrice = 0
Const DefaultCliReceiver = 1
Const DefaultPaymentMethod = 0
Const DefaultCallStatus = 1
'local variable(s) to hold property value(s)
Private mvarCallDateTime As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarAdhamProductViewModel As AdhamProductViewModel 'local copy
Private mvarStrId As Variant 'local copy
Private mvarStrNo As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarEmpNo As Integer 'local copy
Private mvarFullName As String 'local copy
Public Property Let FullName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FullName = 5
    mvarFullName = vData
End Property


Public Property Get FullName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FullName
    FullName = mvarFullName
End Property



Public Property Let empNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EmpNo = 5
    mvarEmpNo = vData
End Property


Public Property Get empNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EmpNo
    empNo = mvarEmpNo
End Property



Public Property Let StrNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StrNo = 5
    mvarStrNo = vData
End Property


Public Property Get StrNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StrNo
    StrNo = mvarStrNo
End Property



Public Property Let Strid(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StrId = 5
    mvarStrId = vData
End Property


Public Property Get Strid() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StrId
    Strid = mvarStrId
End Property

Public Property Set AdhamProductViewModel(ByVal vData As AdhamProductViewModel)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AdhamProductViewModel = Form1
    Set mvarAdhamProductViewModel = vData
End Property


Public Property Get AdhamProductViewModel() As AdhamProductViewModel
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdhamProductViewModel
    Set AdhamProductViewModel = mvarAdhamProductViewModel
End Property



Public Property Let CallDateTime(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallDateTime = 5
    mvarCallDateTime = vData
End Property


Public Property Set CallDateTime(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CallDateTime = Form1
    Set mvarCallDateTime = vData
End Property


Public Property Get CallDateTime() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallDateTime
    If IsObject(mvarCallDateTime) Then
        Set CallDateTime = mvarCallDateTime
    Else
        CallDateTime = mvarCallDateTime
    End If
End Property

Public Property Let TeamName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TeamName = 5
    mvarTeamName = vData
End Property


Public Property Get TeamName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TeamName
    TeamName = mvarTeamName
End Property

Public Property Let ReparationState(ByVal vData As EnumState)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property


Public Property Get ReparationState() As EnumState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
     ReparationState = mvarState
End Property



Public Property Let CallStatus(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallStatus = 5
    mvarCallStatus = vData
End Property


Public Property Get CallStatus() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallStatus
    CallStatus = mvarCallStatus
End Property

Public Property Let ClientName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClientName = 5
    mvarClientName = vData
End Property


Public Property Get ClientName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClientName
    ClientName = mvarClientName
End Property


Public Property Let VoltAfter(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VoltAfter = 5
    mvarVoltAfter = vData
End Property


Public Property Get VoltAfter() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VoltAfter
    VoltAfter = mvarVoltAfter
End Property



Public Property Let VoltBefor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VoltBefor = 5
    mvarVoltBefor = vData
End Property


Public Property Get VoltBefor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VoltBefor
    VoltBefor = mvarVoltBefor
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let CliRecever(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CliRecever = 5
    mvarCliRecever = vData
End Property


Public Property Get CliRecever() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CliRecever
    CliRecever = mvarCliRecever
End Property



Public Property Let Carried(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Carried = 5
    mvarCarried = vData
End Property


Public Property Get Carried() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Carried
    Carried = mvarCarried
End Property

Public Property Let Notes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Notes = 5
    mvarNotes = vData
End Property


Public Property Get Notes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Notes
    Notes = mvarNotes
End Property



Public Property Let regestdate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RegestDate = 5
    mvarregestdate = vData
End Property


Public Property Get regestdate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RegestDate
    regestdate = mvarregestdate
End Property



Public Property Let Cash(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cash = 5
    mvarCash = vData
End Property


Public Property Get Cash() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cash
    Cash = mvarCash
End Property



Public Property Let RepTimeEnd(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepTimeEnd = 5
    mvarRepTimeEnd = vData
End Property


Public Property Get RepTimeEnd() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepTimeEnd
    RepTimeEnd = mvarRepTimeEnd
End Property

Public Property Let RepTimeBegin(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepTimeBegin = 5
    mvarRepTimeBegin = vData
End Property

Public Property Get RepTimeBegin() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepTimeBegin
    RepTimeBegin = mvarRepTimeBegin
End Property

Public Property Let ProdNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdNo = 5
    mvarProdNo = vData
End Property

Public Property Get ProdNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdNo
    ProdNo = mvarProdNo
End Property

Public Property Let TeamNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TeamNo = 5
    mvarTeamNo = vData
End Property

Public Property Get TeamNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TeamNo
    TeamNo = mvarTeamNo
End Property

Public Property Let RepPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepPrice = 5
    mvarRepPrice = vData
End Property


Public Property Get RepPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepPrice
    RepPrice = mvarRepPrice
End Property



Public Property Set ReparationWorksViewModel(ByVal vData As ReparationWorksViewModel)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ReparationWorks = Form1
    Set mvarReparationWorksViewModel = vData
End Property


Public Property Get ReparationWorksViewModel() As ReparationWorksViewModel
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationWorks
    Set ReparationWorksViewModel = mvarReparationWorksViewModel
End Property



Public Property Set ReparationPiecesViewModel(ByVal vData As ReparationPiecesViewModel)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ReparatonPieces = Form1
    Set mvarReparatonPiecesViewModel = vData
End Property


Public Property Get ReparationPiecesViewModel() As ReparationPiecesViewModel
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparatonPieces
    Set ReparationPiecesViewModel = mvarReparatonPiecesViewModel
End Property

Public Property Let RepDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReperationDate = 5
    mvarReperationDate = vData
End Property


Public Property Get RepDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReperationDate
    RepDate = mvarReperationDate
End Property



'Public Property Let EntryDate(ByVal vData As Variant)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.EntryDate = 5
'    mvarEntryDate = vData
'End Property
'
'
'Public Property Get EntryDate() As Variant
''used when retrieving value of a property, on the right side of an assignment.
''Syntax: Debug.Print X.EntryDate
'    EntryDate = mvarEntryDate
'End Property

Public Property Let repNo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepNo = 5
    mvarRepNo = vData
End Property

Public Property Get repNo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepNo
    repNo = mvarRepNo
End Property

Public Property Let CallNo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallNo = 5
    mvarCallNo = vData
End Property


Public Property Get CallNo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallNo
    CallNo = mvarCallNo
End Property

Private Sub Class_Initialize()
    CallNo = 0
    repNo = 0
    TeamName = ""
    TeamNo = 0
    RepTimeBegin = Null
    RepTimeEnd = Null
    RepPrice = DefaulrRepprice
    CallStatus = DefaultCallStatus
    Cash = DefaultPaymentMethod
    regestdate = Format(Date, "dd/mm/yyyy")
    RepDate = Format(Date, "dd/mm/yyyy")
    Notes = ""
    Carried = False
    CliRecever = DefaultCliReceiver
    Description = ""
    VoltBefor = DefaultVoltBefore
    VoltAfter = DefaultVoltAfter
    Me.ReparationState = DefaultRecord

    Set AdhamProductViewModel = New AdhamProductViewModel
    Set ReparationPiecesViewModel = New ReparationPiecesViewModel
    Set ReparationWorksViewModel = New ReparationWorksViewModel
End Sub

