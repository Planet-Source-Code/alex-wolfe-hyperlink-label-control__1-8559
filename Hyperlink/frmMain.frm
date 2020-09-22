VERSION 5.00
Object = "*\AprjHyperlink.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hyperlink Test"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin HyperlinkAX.Hyperlink Hyperlink2 
      Height          =   420
      Left            =   1140
      ToolTipText     =   "Email us today!"
      Top             =   765
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Email Us"
      HoverColor      =   65535
      HyperLink       =   "mailto:someone@yourdomain.com"
      BorderStyle     =   1
      ForeColor       =   65280
      BackColor       =   0
      AutoToolTip     =   0   'False
      SoundOn         =   0   'False
   End
   Begin HyperlinkAX.Hyperlink Hyperlink1 
      Height          =   435
      Left            =   705
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Microsoft"
      HyperLink       =   "http://www.microsoft.com"
      ForeColor       =   16711680
      MOUnderline     =   0   'False
      MOPointer       =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Hyperlink1.Caption = "Microsoft" & Chr(174)
End Sub
