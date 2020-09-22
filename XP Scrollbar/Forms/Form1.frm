VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Exactly XP Scrollbar :: Osen Kusnadi © 2005"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar9 
      Height          =   5835
      Left            =   4620
      TabIndex        =   17
      Top             =   180
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   10292
      LargeChange     =   10
      Max             =   1000
      Colorscheme     =   2
   End
   Begin XPScrollbar.OsenXPHScrollBar OsenXPHScrollBar2 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6210
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      Enabled         =   0   'False
      LargeChange     =   10000
      Max             =   1000000
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar8 
      Height          =   6195
      Left            =   5280
      TabIndex        =   15
      Top             =   180
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   10927
      Enabled         =   0   'False
      LargeChange     =   100
      Max             =   1000
      Colorscheme     =   1
   End
   Begin XPScrollbar.OsenXPHScrollBar OsenXPHScrollBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6570
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   450
      LargeChange     =   10
      Max             =   255
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar7 
      Height          =   6855
      Left            =   5730
      TabIndex        =   13
      Top             =   150
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   12091
      LargeChange     =   10
      Max             =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More ..."
      Height          =   435
      Left            =   2070
      TabIndex        =   12
      Top             =   4560
      Width           =   1635
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar3 
      Height          =   2505
      Left            =   600
      TabIndex        =   5
      Top             =   1710
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4419
      LargeChange     =   20
      Max             =   100
      Value           =   100
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar2 
      Height          =   2505
      Left            =   1140
      TabIndex        =   4
      Top             =   1710
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4419
      LargeChange     =   20
      Max             =   50
      Value           =   40
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar1 
      Height          =   2505
      Left            =   1680
      TabIndex        =   3
      Top             =   1710
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4419
      LargeChange     =   20
      Max             =   20
      Value           =   15
      Colorscheme     =   1
   End
   Begin XPScrollbar.OsenXPHScrollBar BBAr 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1260
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   450
      LargeChange     =   50
      Max             =   255
      Colorscheme     =   2
   End
   Begin XPScrollbar.OsenXPHScrollBar GBar 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   750
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   450
      LargeChange     =   50
      Max             =   255
      Colorscheme     =   1
   End
   Begin XPScrollbar.OsenXPHScrollBar Rbar 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   210
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   450
      LargeChange     =   50
      Max             =   255
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar4 
      Height          =   2505
      Left            =   2280
      TabIndex        =   9
      Top             =   1710
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4419
      LargeChange     =   20
      Max             =   100
      Value           =   100
      Colorscheme     =   1
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar5 
      Height          =   2505
      Left            =   2820
      TabIndex        =   10
      Top             =   1710
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4419
      LargeChange     =   20
      Max             =   50
      Value           =   25
      Colorscheme     =   2
   End
   Begin XPScrollbar.OsenXPVScrollBar OsenXPVScrollBar6 
      Height          =   2505
      Left            =   3360
      TabIndex        =   11
      Top             =   1710
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4419
      LargeChange     =   20
      Max             =   20
      Value           =   20
      Colorscheme     =   2
   End
   Begin VB.Label B 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3900
      TabIndex        =   8
      Top             =   1290
      Width           =   75
   End
   Begin VB.Label G 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3900
      TabIndex        =   7
      Top             =   780
      Width           =   75
   End
   Begin VB.Label R 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3900
      TabIndex        =   6
      Top             =   240
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub BBAr_Change()
    ChangeBackGround
End Sub

Private Sub Command1_Click()
    MyOpenBrowser
End Sub

Private Sub GBar_Change()
    ChangeBackGround
End Sub

Private Sub Rbar_Change()
    ChangeBackGround
End Sub

Private Sub ChangeBackGround()

    BackColor = RGB(Rbar.Value, GBar.Value, BBAr.Value)
    R = Rbar.Value
    G = GBar.Value
    B = BBAr.Value
   
End Sub

Private Sub MyOpenBrowser(Optional IpHomePage As String = "http://osenxpsuite.net/download")

    ShellExecute hwnd, "open", IpHomePage, vbNullString, vbNullString, 1

End Sub



