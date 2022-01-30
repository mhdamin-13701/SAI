VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "Threed20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8C458270-26A2-11D2-9327-00A0C91AD7BF}#5.0#0"; "AnimFX.ocx"
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3870
      Top             =   3390
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   6240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   926
      _Version        =   131074
      Begin VB.TextBox TxtPasswprd 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3060
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   90
         Width           =   2760
      End
      Begin VB.TextBox TxtUserId 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   1
         Top             =   90
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Password"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2190
         TabIndex        =   4
         Top             =   150
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User Id."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   150
         Width           =   555
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   5
      Top             =   690
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   9763
      _Version        =   131074
      Begin AnimationFX.AnimationFX AnimationFX 
         Height          =   5475
         Left            =   0
         OleObjectBlob   =   "FrmAbout.frx":0000
         TabIndex        =   6
         Top             =   0
         Width           =   10815
      End
      Begin VB.Image Image1 
         Height          =   5475
         Left            =   30
         Top             =   30
         Width           =   10815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   6795
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   714
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   780
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbout.frx":0071
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbout.frx":2ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAbout.frx":57FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CheckPassword() As Boolean
On Error GoTo ErrorHandler
    CallByName Me, systemConfigration.PasswordMethod, VbMethod
    ConnectString = "Provider=SQLNCLI11.1 " & ";Initial Catalog=" & systemConfigration.DatabaseName & ";Data Source=" & systemConfigration.ServerName
    If de.con.State <> adStateOpen Then de.con.Open ConnectString, systemConfigration.UserId, systemConfigration.Password
    If de.con.State <> adStateOpen Then de.con.Open ConnectString, systemConfigration.UserId, systemConfigration.Password
    de.con.CommandTimeout = 0
    sqlText = "Select * From users Where UserId=" & IIf(LTrim(RTrim(TxtUserId.text)) = "", 0, TxtUserId.text) & " And Password='" & TxtPasswprd.text & "'"
    Set rs = de.con.Execute(sqlText)
    CheckPassword = Not (rs.EOF)
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function

Sub GetPassword()
systemConfigration.Password = GetPass()
End Sub
Function LoadProgram() As Boolean
If CheckPassword() Then
    Set rs = de.con.Execute("Select * From Users Where UserId=" & TxtUserId.text)
    UserId = TxtUserId.text
    
    LoadProgram = True
Else
        MsgBox "Wrong password", vbExclamation + vbMsgBoxRight, "Attention"
        TxtPasswprd.SetFocus
        Sendkeys "{Home}+{End}"
        LoadProgram = False
End If
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Sendkeys "{Tab}"
    Sendkeys "{Home}+{end}"
End If
End Sub

Private Sub Form_Load()
    ParseXml
    ProgramTitle = systemConfigration.SystemName & " " & systemConfigration.Year & " -version " & systemConfigration.Version
    Me.Caption = ProgramTitle
    Effect = 1
    With AnimationFX
         Image1.Picture = LoadPicture(App.Path + "\photo\graphics_152.jpg")
        .Bitmap = Image1.Picture.Handle
        .Operation = 9
        .Modifier = 0
        .Interval = 1
        .Steps = 10
        .Delta = 1
        .UseThread = False
        AnimationFX.Play
    End With
End Sub


Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
    StatusBar1.SimpleText = Now
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If LoadProgram Then
            Unload FrmAbout
            FrmMain.Show
        End If
    Case 3
        If CheckPassword Then
            UserId = TxtUserId.text
            FrmUpdatePassword.Show
        End If
    Case 5
       Unload Me
End Select
End Sub

Private Sub TxtPasswprd_GotFocus()
ChangeToEnglish
End Sub

Private Sub TxtPasswprd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LoadProgram Then
            Unload FrmAbout
            FrmMain.Show
        End If
    End If
End Sub

