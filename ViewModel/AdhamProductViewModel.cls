VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdhamProductViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarProdNo As Variant 'local copy
Private mvarSymbol As String 'local copy
Private mvarProdBaCodeNo As String 'local copy
Private mvarProdDate As Variant 'local copy
Private mvarProdPurchaseDate As Variant 'local copy
Private mvarModNo As Integer 'local copy
Private mvarItemNo As String 'local copy
Private mvarGasNo As String 'local copy
Private mvarBarcodeSerialNo As String 'local copy
Private mvarProdFamName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarProductState As EnumState 'local copy
'local variable(s) to hold property value(s)
Private mvarProdFamNo As Integer 'local copy
Public Property Let ProdFamNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdFamNo = 5
    mvarProdFamNo = vData
End Property


Public Property Get ProdFamNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdFamNo
    ProdFamNo = mvarProdFamNo
End Property



Public Property Let ProductState(ByVal vData As EnumState)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProductState = Form1
      mvarProductState = vData
End Property


Public Property Get ProductState() As EnumState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductState
      ProductState = mvarProductState
End Property




Public Property Let ProdFamName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdFamName = 5
    mvarProdFamName = vData
End Property


Public Property Get ProdFamName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdFamName
    ProdFamName = mvarProdFamName
End Property



Public Property Let BarcodeSerialNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BarcodeSerialNo = 5
    mvarBarcodeSerialNo = vData
End Property


Public Property Get BarcodeSerialNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BarcodeSerialNo
    BarcodeSerialNo = mvarBarcodeSerialNo
End Property



Public Property Let GasNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GasNo = 5
    mvarGasNo = vData
End Property


Public Property Get GasNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GasNo
    GasNo = mvarGasNo
End Property



Public Property Let ItemNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemNo = 5
    mvarItemNo = vData
End Property


Public Property Get ItemNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemNo
    ItemNo = mvarItemNo
End Property



Public Property Let ModNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ModNo = 5
    mvarModNo = vData
End Property


Public Property Get ModNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ModNo
    ModNo = mvarModNo
End Property



Public Property Let ProdPurchaseDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdPurchaseDate = 5
    mvarProdPurchaseDate = vData
End Property


Public Property Set ProdPurchaseDate(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProdPurchaseDate = Form1
    Set mvarProdPurchaseDate = vData
End Property


Public Property Get ProdPurchaseDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdPurchaseDate
    If IsObject(mvarProdPurchaseDate) Then
        Set ProdPurchaseDate = mvarProdPurchaseDate
    Else
        ProdPurchaseDate = mvarProdPurchaseDate
    End If
End Property



Public Property Let ProdDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdDate = 5
    mvarProdDate = vData
End Property


Public Property Set ProdDate(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProdDate = Form1
    Set mvarProdDate = vData
End Property


Public Property Get ProdDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdDate
    If IsObject(mvarProdDate) Then
        Set ProdDate = mvarProdDate
    Else
        ProdDate = mvarProdDate
    End If
End Property



Public Property Let ProdBaCodeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdBaCodeNo = 5
    mvarProdBaCodeNo = vData
End Property


Public Property Get ProdBaCodeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdBaCodeNo
    ProdBaCodeNo = mvarProdBaCodeNo
End Property



Public Property Let Symbol(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Symbol = 5
    mvarSymbol = vData
End Property


Public Property Get Symbol() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Symbol
    Symbol = mvarSymbol
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



