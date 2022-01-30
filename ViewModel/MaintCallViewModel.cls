VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MaintCallViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarId_ComNo As Integer 'local copy
Private mvarCallNo As Variant 'local copy
Private mvarCallDateTime As Date 'local copy
Private mvarModNo As Variant 'local copy
Private mvarCliNo As Variant 'local copy
Private mvarCallDescription As String 'local copy
Private mvarZoneNo As String 'local copy
Private mvarCallStatus As Integer 'local copy
Private mvarNotes As String 'local copy
Private mvarDefindName As String 'local copy
Private mvarCallReceiverEmpNo As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarMaintStatus As EnumState 'local copy
'local variable(s) to hold property value(s)
Private mvarClientInfo As AdhamViewModel 'local copy
'local variable(s) to hold property value(s)
Private mvarCallReceiverName As Variant 'local copy
Private mvarCallEntryEmpNo As Variant 'local copy
Private mvarCallEntryName As Variant 'local copy
Private mvarZoneName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarClientVisitorCount As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarCallState As EnumMaintCallState 'local copy
'local variable(s) to hold property value(s)
Private mvarClaimIsRepeat As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarId As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarLastVisitDate As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarLastCallNo As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarClaimTeamName As String 'local copy
Public Property Let ClaimTeamName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClaimTeamName = 5
    mvarClaimTeamName = vData
End Property


Public Property Get ClaimTeamName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClaimTeamName
    ClaimTeamName = mvarClaimTeamName
End Property



Public Property Let LastCallNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastCallNo = 5
    mvarLastCallNo = vData
End Property

Public Property Get LastCallNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastCallNo
    If IsObject(mvarLastCallNo) Then
        Set LastCallNo = mvarLastCallNo
    Else
        LastCallNo = mvarLastCallNo
    End If
End Property



Public Property Let LastVisitDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastVisitDate = 5
    mvarLastVisitDate = vData
End Property


Public Property Get LastVisitDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastVisitDate
    LastVisitDate = mvarLastVisitDate
End Property



Public Property Let Id(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Id = 5
    mvarId = vData
End Property


Public Property Get Id() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Id
    Id = mvarId
End Property



Public Property Let ClaimIsRepeat(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClaimIsRepeat = 5
    mvarClaimIsRepeat = vData
End Property


Public Property Get ClaimIsRepeat() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClaimIsRepeat
    ClaimIsRepeat = mvarClaimIsRepeat
End Property



Public Property Let callState(ByVal vData As EnumMaintCallState)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallState = 5
    mvarCallState = vData
End Property


Public Property Get callState() As EnumMaintCallState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallState
     callState = mvarCallState
End Property



Public Property Let ClientVisitorCount(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClientVisitorCount = 5
    mvarClientVisitorCount = vData
End Property


Public Property Set ClientVisitorCount(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ClientVisitorCount = Form1
    Set mvarClientVisitorCount = vData
End Property


Public Property Get ClientVisitorCount() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClientVisitorCount
    If IsObject(mvarClientVisitorCount) Then
        Set ClientVisitorCount = mvarClientVisitorCount
    Else
        ClientVisitorCount = mvarClientVisitorCount
    End If
End Property



Public Property Let ZoneName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ZoneName = 5
    mvarZoneName = vData
End Property


Public Property Get ZoneName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ZoneName
    ZoneName = mvarZoneName
End Property



Public Property Let CallEntryName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallEntryName = 5
    mvarCallEntryName = vData
End Property


Public Property Get CallEntryName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallEntryName
    CallEntryName = mvarCallEntryName
End Property



Public Property Let CallEntryEmpNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallEntryEmpNo = 5
    mvarCallEntryEmpNo = vData
End Property


Public Property Get CallEntryEmpNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallEntryEmpNo
    CallEntryEmpNo = mvarCallEntryEmpNo
End Property



Public Property Let CallReceiverName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallReceiverName = 5
    mvarCallReceiverName = vData
End Property


Public Property Get CallReceiverName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallReceiverName
    CallReceiverName = mvarCallReceiverName
End Property

Public Property Set clientInfo(ByVal vData As AdhamViewModel)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ClientInfo = Form1
    Set mvarClientInfo = vData
End Property


Public Property Get clientInfo() As AdhamViewModel
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClientInfo
    Set clientInfo = mvarClientInfo
End Property



Public Property Let MaintStatus(ByVal vData As EnumState)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MaintStatus = Form1
     mvarMaintStatus = vData
End Property


Public Property Get MaintStatus() As EnumState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaintStatus
     MaintStatus = mvarMaintStatus
End Property



Public Property Let CallReceiverEmpNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallReceiverEmpNo = 5
    mvarCallReceiverEmpNo = vData
End Property


Public Property Get CallReceiverEmpNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallReceiverEmpNo
    CallReceiverEmpNo = mvarCallReceiverEmpNo
End Property



Public Property Let Defindname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefindName = 5
    mvarDefindName = vData
End Property


Public Property Get Defindname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefindName
    Defindname = mvarDefindName
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



Public Property Let CallStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallStatus = 5
    mvarCallStatus = vData
End Property


Public Property Get CallStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallStatus
    CallStatus = mvarCallStatus
End Property



Public Property Let ZoneNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ZoneNo = 5
    mvarZoneNo = vData
End Property


Public Property Get ZoneNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ZoneNo
    ZoneNo = mvarZoneNo
End Property



Public Property Let CallDEscription(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallDescription = 5
    mvarCallDescription = vData
End Property


Public Property Get CallDEscription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallDescription
    CallDEscription = mvarCallDescription
End Property



Public Property Let cliNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CliNo = 5
    mvarCliNo = vData
End Property


Public Property Get cliNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CliNo
    cliNo = mvarCliNo
End Property



Public Property Let ModNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ModNo = 5
    mvarModNo = vData
End Property


Public Property Get ModNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ModNo
    ModNo = mvarModNo
End Property



Public Property Let CallDateTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallDateTime = 5
    mvarCallDateTime = vData
End Property


Public Property Get CallDateTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallDateTime
    CallDateTime = mvarCallDateTime
End Property



Public Property Let CallNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallNo = 5
    mvarCallNo = vData
End Property


Public Property Get CallNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallNo
    CallNo = mvarCallNo
End Property



Public Property Let Id_ComNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Id_ComNo = 5
    mvarId_ComNo = vData
End Property


Public Property Get Id_ComNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Id_ComNo
    Id_ComNo = mvarId_ComNo
End Property

Private Sub Class_Initialize()
Dim FullName As String
FullName = GetFullName(empNo)
Set clientInfo = New AdhamViewModel
Me.CallNo = Null
Me.clientInfo.adhamNo = Null
Me.CallDEscription = ""
Me.callState = 0
Me.ClientVisitorCount = 0
Me.Defindname = ""
Me.Id_ComNo = 1
Me.ModNo = Null
Me.Notes = ""
Me.clientInfo.Zone = Null
Me.CallEntryEmpNo = empNo
Me.CallEntryName = FullName
Me.CallReceiverEmpNo = empNo
Me.CallReceiverName = FullName


Me.CallDateTime = Now
End Sub



