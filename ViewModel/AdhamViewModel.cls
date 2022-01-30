VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdhamViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClientState"
'local variable(s) to hold property value(s)
Private mvarAdhamNo As Variant 'local copy
Private mvarAdNo As Variant 'local copy
Private mvarAdhamphon As String 'local copy
Private mvarAdhamname As String 'local copy
Private mvarMobilePhone As String 'local copy
Private mvarAdhamAdress As String 'local copy
Private mvarNotes As String 'local copy
Private mvarZone As Variant 'local copy
Private mvarKind As Integer 'local copy
Private mvarDefindName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarZoneName As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarClientState As EnumClientState 'local copy
Public Property Let ClientState(ByVal vData As EnumClientState)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ClientState = 5
    mvarClientState = vData
End Property


Public Property Get ClientState() As EnumClientState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ClientState
     ClientState = mvarClientState
End Property



Public Property Let ZoneName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ZoneName = 5
    mvarZoneName = vData
End Property


Public Property Get ZoneName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ZoneName
    ZoneName = mvarZoneName
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



Public Property Let Kind(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Kind = 5
    mvarKind = vData
End Property


Public Property Get Kind() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Kind
    Kind = mvarKind
End Property



Public Property Let Zone(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Zone = 5
    mvarZone = vData
End Property


Public Property Get Zone() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Zone
    Zone = mvarZone
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



Public Property Let AdhamAdress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdhamAdress = 5
    mvarAdhamAdress = vData
End Property


Public Property Get AdhamAdress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdhamAdress
    AdhamAdress = mvarAdhamAdress
End Property



Public Property Let MobilePhone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MobilePhone = 5
    mvarMobilePhone = vData
End Property


Public Property Get MobilePhone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MobilePhone
    MobilePhone = mvarMobilePhone
End Property



Public Property Let AdhamName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Adhamname = 5
    mvarAdhamname = vData
End Property


Public Property Get AdhamName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Adhamname
    AdhamName = mvarAdhamname
End Property



Public Property Let AdhamPhon(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Adhamphon = 5
    mvarAdhamphon = vData
End Property


Public Property Get AdhamPhon() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Adhamphon
    AdhamPhon = mvarAdhamphon
End Property



Public Property Let AdNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdNo = 5
    mvarAdNo = vData
End Property


Public Property Get AdNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdNo
    AdNo = mvarAdNo
End Property

Public Property Let adhamNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AdhamNo = 5
    mvarAdhamNo = vData
End Property


Public Property Get adhamNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdhamNo
    adhamNo = mvarAdhamNo
End Property



