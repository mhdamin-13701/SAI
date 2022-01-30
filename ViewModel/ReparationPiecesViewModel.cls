VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReparationPiecesViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ReparationPiece"
Attribute VB_Ext_KEY = "Member0" ,"AdhamProduct"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"ReparationPiece"
'local variable to hold collection
Private mCol As Collection
Private mvarAdhamProduct As AdhamProduct


Public Property Get AdhamProduct() As AdhamProduct
    If mvarAdhamProduct Is Nothing Then
        Set mvarAdhamProduct = New AdhamProduct
    End If


    Set AdhamProduct = mvarAdhamProduct
End Property


Public Property Set AdhamProduct(vData As AdhamProduct)
    Set mvarAdhamProduct = vData
End Property


Public Function Add(stkId As Double, stkno As String, Qty As Integer, Price As Double, StkName As String, SerId As Double, ReparationPieceState As EnumState, Optional sKey As String) As ReparationPieceViewModel
    'create a new object
    Dim objNewMember As ReparationPieceViewModel
    Set objNewMember = New ReparationPieceViewModel


    'set the properties passed into the method
'    objNewMember.repNo = repNo
    objNewMember.stkId = stkId
    objNewMember.stkno = stkno
    objNewMember.Qty = Qty
    objNewMember.Price = Price
'    objNewMember.repNo = repNo
    objNewMember.StkName = StkName
    objNewMember.SerId = SerId
    objNewMember.ReparationPieceState = ReparationPieceState
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As ReparationPieceViewModel
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrorHandler
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol.Item(vntIndexKey)
  Exit Property
ErrorHandler:
  Set Item = Nothing
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
  Set mvarAdhamProduct = Nothing
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

