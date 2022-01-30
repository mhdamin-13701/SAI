VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "Threed20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSAIManagment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Managment"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame2 
      Height          =   6045
      Left            =   30
      TabIndex        =   18
      Top             =   750
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10663
      _Version        =   131074
      Begin VB.TextBox TxtUserId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1260
         TabIndex        =   40
         Top             =   30
         Width           =   2445
      End
      Begin VB.ListBox EmployeeList 
         ForeColor       =   &H00000080&
         Height          =   3765
         Left            =   60
         TabIndex        =   19
         Top             =   2220
         Width           =   3615
      End
      Begin VSFlex8Ctl.VSFlexGrid GridUsers 
         Height          =   1875
         Left            =   60
         TabIndex        =   20
         Top             =   330
         Width           =   3645
         _cx             =   6429
         _cy             =   3307
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "User Id"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   60
         Width           =   615
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   780
      Left            =   7890
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   1376
      _Version        =   131074
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   27
         Left            =   2250
         TabIndex        =   37
         Tag             =   "0"
         Top             =   7890
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ßÔÝ ÍÓÇÈ ÇáãæÇÏ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   26
         Left            =   2430
         TabIndex        =   36
         Tag             =   "0"
         Top             =   7650
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÑÕíÏ ÇáãæÇÏ ÇáãÎÒäíå"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   25
         Left            =   2340
         TabIndex        =   35
         Tag             =   "0"
         Top             =   7380
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÍÑßå ãÇÏå"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   24
         Left            =   2280
         TabIndex        =   34
         Tag             =   "0"
         Top             =   7080
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "äÞá ãä ãÓÊæÏÚ Åáì ãÓÊæÏÚ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   23
         Left            =   2280
         TabIndex        =   33
         Tag             =   "0"
         Top             =   6750
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÇáÅÏÎÇá æÇáÅÎÑÇÌ"
         Alignment       =   1
      End
      Begin Threed.SSCheck ChkPrintBills 
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   32
         Tag             =   "6"
         Top             =   5010
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ØÈÇÚÉ ÝÇÊæÑ ÇÕáíÉ"
      End
      Begin Threed.SSCheck ChkExportItemsBills 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   31
         Tag             =   "6"
         Top             =   2280
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÊÕÏíÑ ÅÔÚÇÑÇÊ ÇáãæÇÏ ÇáÃæáíÉ"
      End
      Begin Threed.SSCheck ChkTransferToMvStock 
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   30
         Tag             =   "6"
         Top             =   1410
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÊÑÍíá Åá ÇáãÓÊæÏÚÇÊ"
      End
      Begin Threed.SSCheck ChkCancelTRansfer 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   29
         Tag             =   "6"
         Top             =   1980
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÅáÛÇÁ ÇáÊËÈíÊ"
      End
      Begin Threed.SSCheck ChkAllowTransfer 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Tag             =   "6"
         Top             =   1710
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÊÑÍíá Åáì ÇáãÍÇÓÈÉ"
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   22
         Left            =   2250
         TabIndex        =   27
         Tag             =   "0"
         Top             =   6105
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÊÕÏíÑ ÇáãÚáæãÇÊ ÇáÇÓÇÓíå"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   21
         Left            =   2250
         TabIndex        =   26
         Tag             =   "0"
         Top             =   5745
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÑÝÚ ÇÓÚÇÑ ÇáãæÇÏ ÇáÇæáíå"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   20
         Left            =   2250
         TabIndex        =   25
         Tag             =   "8"
         Top             =   2010
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÊÑãíÒ ÞÇÆãÉ  ÊÇãæÏíáÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   19
         Left            =   2190
         TabIndex        =   24
         Tag             =   "29"
         Top             =   5445
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÇáÈÇÑßæÏ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   18
         Left            =   2220
         TabIndex        =   23
         Tag             =   "28"
         Top             =   5160
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÑÈØ ÇáãæÏíáÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   17
         Left            =   2220
         TabIndex        =   22
         Tag             =   "27"
         Top             =   4890
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ßæßÇ ßæáÇ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   16
         Left            =   2250
         TabIndex        =   7
         Tag             =   "9"
         Top             =   2310
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÊÑãíÒ ÊÇÈÚíÉ ÇáãæÏíáÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   15
         Left            =   2250
         TabIndex        =   9
         Tag             =   "11"
         Top             =   2910
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   14
         Left            =   2250
         TabIndex        =   14
         Tag             =   "26"
         Top             =   4620
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÃÑÔÝÉ ÇáÈíÇäÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   13
         Left            =   2250
         TabIndex        =   13
         Tag             =   "25"
         Top             =   4305
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÅÍÕÇÆíÉ ÚäÇæíä ÇáÕíÇäÉ ÇáÎÇÑÌíÉ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   10
         Left            =   2250
         TabIndex        =   12
         Tag             =   "24"
         Top             =   3945
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ØÈÇÚå ÍÑßÇÊ ÇáãÓÊæÏÚÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   9
         Left            =   2250
         TabIndex        =   8
         Tag             =   "10"
         Top             =   2580
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   12
         Left            =   2250
         TabIndex        =   16
         Tag             =   "23"
         Top             =   8190
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÅÏÇÑÉ ÇáãÓÊæíÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   2
         Left            =   2250
         TabIndex        =   2
         Tag             =   "3"
         Top             =   540
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   6
         Left            =   2250
         TabIndex        =   10
         Tag             =   "21"
         Top             =   3285
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÌÑÏ ÇáãæÇÏ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   3
         Left            =   2250
         TabIndex        =   3
         Tag             =   "4"
         Top             =   795
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÝæÇÊíÑ ÎÏãÉ ÇáãÓÊåáß"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   1
         Left            =   3840
         TabIndex        =   1
         Tag             =   "2"
         Top             =   270
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   5
         Left            =   2250
         TabIndex        =   5
         Tag             =   "6"
         Top             =   1395
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÇáÈÍË Úä ÝÇÊæÑÉ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   7
         Left            =   2250
         TabIndex        =   11
         Tag             =   "22"
         Top             =   3615
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÇáÚäÇæíä ÇáãßÑÑÉ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   0
         Left            =   2250
         TabIndex        =   0
         Tag             =   "1"
         Top             =   30
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   8
         Left            =   2250
         TabIndex        =   6
         Tag             =   "7"
         Top             =   1695
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   11
         Left            =   2250
         TabIndex        =   15
         Tag             =   "0"
         Top             =   6435
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         CaptionStyle    =   1
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÇáãÓÊæÏÚÇÊ"
         Alignment       =   1
      End
      Begin Threed.SSCheck Check 
         Height          =   270
         Index           =   4
         Left            =   2250
         TabIndex        =   4
         Tag             =   "5"
         Top             =   1095
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   476
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2130
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":4E7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":7777
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":A126
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":C64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":EE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":11817
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":14169
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":16EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":1972C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":1C5D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":1F32D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":21CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":24C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":27A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":2A4B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":2CE6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":2F79F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":323DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":34CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":3774B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":3A6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":3D025
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":3F95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":42509
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":44C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":475B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":49E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":4C04F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":4E9AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":511D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":53C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":566C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":591DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":5C176
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":5EF42
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSAIManagment.frx":61BBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6075
      Left            =   3840
      TabIndex        =   38
      Top             =   720
      Width           =   7455
      _cx             =   13150
      _cy             =   10716
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmSAIManagment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColTagId = 1
Const ColTagName = 2
Const ColChk = 3




Const ColUserId = 1
Const ColUserName = 2

Dim rs As New ADODB.Recordset
Dim Flag As Boolean


Sub init()
Dim rs As New ADODB.Recordset
    Me.Top = 0
    Me.Left = 0
    
    sqlText = "Select TagId , TagName , 0 chk From coSaiTag"
    Set rs = de.con.Execute(sqlText)
    Set FlexGrid.DataSource = rs
    FillFormating FlexGrid
    FlexGrid.Editable = flexEDKbdMouse
    
    sqlText = "Select Top 5 userId , UserName From users where UserId=0"
    Set rs = de.con.Execute(sqlText)
    Set GridUsers.DataSource = rs
    GridUsers.FormatString = FillFs
    SetColWidths ColUserId, GridUsers
    SetColWidths ColUserName, GridUsers
    FillList
    msgFlag = True
End Sub
Sub FillFormating(FlexGrid As VSFlexGrid)
    fs = "|>" + "TagId"
    fs = fs + "|>" + "Tag Name"
    fs = fs + "|>" + "Chk"
   With FlexGrid
        .FormatString = fs
        .Cols = 4
        .ColWidth(ColChk) = 400
        .ColDataType(ColChk) = flexDTBoolean
        SetColWidths ColTagId, FlexGrid
        SetColWidths ColTagName, FlexGrid
   End With
End Sub

Sub SaveRec()

End Sub
Function FoundUserId(TUserId As Integer) As Boolean
sqlText = "Select * From MaintUsers where UserId= " & TUserId
Set rs = de.con.Execute(sqlText)
If rs.EOF And rs.BOF Then
    FoundUserId = False
Else
    FoundUserId = True
End If
End Function

Sub ClearChk()
For i = 0 To Check.UBound
    Check(i).Value = ssCBUnchecked
Next
End Sub

Sub FillList()
    sqlText = "Select UserId , UserName  From Users  "
    Set rs = de.con.Execute(sqlText)
    If rs.EOF And rs.BOF Then Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
        With EmployeeList
           .AddItem rs!UserName
           .ItemData(.NewIndex) = rs!UserId
           rs.MoveNext
        End With
    Loop
    EmployeeList.ListIndex = 0
End Sub

Function Found(Vrow As Integer) As Boolean
With EmployeeList
    For i = 0 To .ListCount - 1
        If .ItemData(i) = GridUsers.TextMatrix(Vrow, ColUserId) Then
            Found = True
            Exit Function
        Else
            Found = False
        End If
    Next i
End With
End Function

Function FillFs() As String
    fs = "|<" + "TagId"
    fs = fs + "|<" + "Tag Name"
    
    FillFs = fs
End Function


Sub SetColWidths(ColNo As Integer, MSHFlexGrid1 As VSFlexGrid)
    With MSHFlexGrid1
        .AutoSize (ColNo)
    End With
End Sub

Private Sub CmdDetailsReport_Click()
                msgFlag = False
                CmdSave_Click
                msgFlag = True
                FrmSubDetailsReport.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
SaveRec
End Sub
Private Sub Command1_Click()
MsgBox SubReport
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With FlexGrid
    sqlText = "Delete From SAIPermissions Where UserId =" & EmployeeList.ItemData(EmployeeList.ListIndex) & " and TagId=" & .TextMatrix(Row, ColTagId)
    de.con.Execute (sqlText)
    If .TextMatrix(Row, ColChk) Then
        sqlText = "Insert Into SAIPermissions(TagId , UserId) Values(" & .TextMatrix(Row, ColTagId) & "," & EmployeeList.ItemData(EmployeeList.ListIndex) & ")"
        de.con.Execute (sqlText)
    End If
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If Col <> ColChk Then cancel = True
End Sub

Private Sub Form_Load()
    init
End Sub

Private Sub EmployeeList_Click()
Dim sqlText As String
Dim selectedUserId As Integer

Flag = False

With EmployeeList
selectedUserId = .ItemData(.ListIndex)
End With

    
    sqlText = ""
    sqlText = sqlText & "select   isnull(t1.tagid , t2.tagid) tagid ,  isnull(t1.tagname , t2.tagname) tagname , isnull(t1.chk,t2.chk)chk   "
    sqlText = sqlText & " From"
    sqlText = sqlText & " ("
    sqlText = sqlText & " select 1 chk , s1.tagid , tagname from"
    sqlText = sqlText & " saipermissions s1 inner join"
    sqlText = sqlText & " users u1 on s1.userid = u1.userid inner join"
    sqlText = sqlText & " CoSAITag c1 on s1.tagid = c1.tagid"
    sqlText = sqlText & " Where s1.UserId = " & selectedUserId
    sqlText = sqlText & " )t1"
    sqlText = sqlText & " full outer join"
    sqlText = sqlText & " ("
    sqlText = sqlText & " select 0 chk , tagid , tagname from CoSAITag"
    sqlText = sqlText & " )t2 on t1.tagid = t2.tagid"
    sqlText = sqlText & " Order by TagId"
    Set rs = de.con.Execute(sqlText)
    Set FlexGrid.DataSource = rs
    FillFormating FlexGrid

Flag = True
End Sub

Private Sub EmployeeList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    With EmployeeList
        sqlText = "Delete From saipermissions where UserId=" & .ItemData(.ListIndex)
        de.con.Execute (sqlText)
        .RemoveItem .ListIndex
    End With
End If
End Sub

Private Sub GridUsers_DblClick()
With GridUsers
    If Not Found(.Row) Then
        EmployeeList.AddItem .TextMatrix(.Row, ColUserName)
        EmployeeList.ItemData(EmployeeList.NewIndex) = .TextMatrix(.Row, ColUserId)
    End If
End With
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Unload Me
    Case 3
        SaveRec
End Select
End Sub

Private Sub TxtUserId_Change()
'On Error Resume Next
    If IsNumeric(TxtUserId.text) Then
        sqlText = "Select  UserId , UserName From Dbo.Users Where UserId =" & TxtUserId.text
        Set rs = de.con.Execute(sqlText)
    Else
        sqlText = "Select  UserId , UserName From Dbo.Users Where UserName like '%" & IIf(LTrim(RTrim(TxtUserId.text)) = "", 0, LTrim(RTrim(TxtUserId.text))) & "%' Order By userName"
        Set rs = de.con.Execute(sqlText)
    End If
    Set GridUsers.DataSource = rs
    GridUsers.FormatString = FillFs
    SetColWidths ColUserId, GridUsers
    SetColWidths ColUserName, GridUsers
End Sub

Private Sub TxtUserId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With GridUsers
            If .Rows = 1 Then Exit Sub
                If Not Found(1) Then
                    EmployeeList.AddItem .TextMatrix(1, ColUserName)
                    EmployeeList.ItemData(EmployeeList.NewIndex) = .TextMatrix(1, ColUserId)
                End If
        End With
    End If
End Sub
