VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarDatabaseName As Variant 'local copy
Private mvarUserId As Variant 'local copy
Private mvarPassword As Variant 'local copy
Private mvarDSN As Variant 'local copy
Private mvarServerName As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarSystemName As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarYear As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarVersion As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarDatabaseDestination As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarPasswordMethod As Variant 'local copy
Public Property Let PasswordMethod(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PasswordMethod = 5
    mvarPasswordMethod = vData
End Property

Public Property Set PasswordMethod(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PasswordMethod = Form1
    Set mvarPasswordMethod = vData
End Property


Public Property Get PasswordMethod() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PasswordMethod
    If IsObject(mvarPasswordMethod) Then
        Set PasswordMethod = mvarPasswordMethod
    Else
        PasswordMethod = mvarPasswordMethod
    End If
End Property



Public Property Let DatabaseDestination(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MoveStmovToDatabase = 5
    mvarDatabaseDestination = vData
End Property


Public Property Set DatabaseDestination(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MoveStmovToDatabase = Form1
    Set mvarDatabaseDestination = vData
End Property


Public Property Get DatabaseDestination() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MoveStmovToDatabase
    If IsObject(mvarDatabaseDestination) Then
        Set DatabaseDestination = mvarDatabaseDestination
    Else
        DatabaseDestination = mvarDatabaseDestination
    End If
End Property



Public Property Let MainStoreNo(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MainStoreNo = 5
    mvarMainStoreNo = vData
End Property


Public Property Set MainStoreNo(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MainStoreNo = Form1
    Set mvarMainStoreNo = vData
End Property


Public Property Get MainStoreNo() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MainStoreNo
    If IsObject(mvarMainStoreNo) Then
        Set MainStoreNo = mvarMainStoreNo
    Else
        MainStoreNo = mvarMainStoreNo
    End If
End Property



Public Property Let Version(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Version = 5
    mvarVersion = vData
End Property


Public Property Set Version(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Version = Form1
    Set mvarVersion = vData
End Property


Public Property Get Version() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Version
    If IsObject(mvarVersion) Then
        Set Version = mvarVersion
    Else
        Version = mvarVersion
    End If
End Property



Public Property Let Year(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Year = 5
    mvarYear = vData
End Property


Public Property Set Year(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Year = Form1
    Set mvarYear = vData
End Property


Public Property Get Year() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Year
    If IsObject(mvarYear) Then
        Set Year = mvarYear
    Else
        Year = mvarYear
    End If
End Property



Public Property Let SystemName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SystemName = 5
    mvarSystemName = vData
End Property


Public Property Set SystemName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SystemName = Form1
    Set mvarSystemName = vData
End Property


Public Property Get SystemName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SystemName
    If IsObject(mvarSystemName) Then
        Set SystemName = mvarSystemName
    Else
        SystemName = mvarSystemName
    End If
End Property



Public Property Let ServerName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ServerName = 5
    mvarServerName = vData
End Property


Public Property Set ServerName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ServerName = Form1
    Set mvarServerName = vData
End Property


Public Property Get ServerName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ServerName
    If IsObject(mvarServerName) Then
        Set ServerName = mvarServerName
    Else
        ServerName = mvarServerName
    End If
End Property



Public Property Let DSN(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Odbc = 5
    mvarDSN = vData
End Property


Public Property Set DSN(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Odbc = Form1
    Set mvarDSN = vData
End Property


Public Property Get DSN() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Odbc
    If IsObject(mvarOdbc) Then
        Set DSN = mvarDSN
    Else
        DSN = mvarDSN
    End If
End Property



Public Property Let Password(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Public Property Set Password(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Password = Form1
    Set mvarPassword = vData
End Property


Public Property Get Password() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    If IsObject(mvarPassword) Then
        Set Password = mvarPassword
    Else
        Password = mvarPassword
    End If
End Property



Public Property Let UserId(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserId = 5
    mvarUserId = vData
End Property


Public Property Set UserId(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.UserId = Form1
    Set mvarUserId = vData
End Property


Public Property Get UserId() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserId
    If IsObject(mvarUserId) Then
        Set UserId = mvarUserId
    Else
        UserId = mvarUserId
    End If
End Property



Public Property Let DatabaseName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DatabaseName = 5
    mvarDatabaseName = vData
End Property


Public Property Set DatabaseName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DatabaseName = Form1
    Set mvarDatabaseName = vData
End Property


Public Property Get DatabaseName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatabaseName
    If IsObject(mvarDatabaseName) Then
        Set DatabaseName = mvarDatabaseName
    Else
        DatabaseName = mvarDatabaseName
    End If
End Property



