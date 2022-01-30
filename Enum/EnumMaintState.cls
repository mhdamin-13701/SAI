VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumMaintState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EnumMaintState"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum EnumMaintCallState
' Temporary = 1
'    Underway = 2
'    Printed = 4
    UnderwayAndPrinted = 1
    Repared = 2
'    Repared = 8
End Enum

