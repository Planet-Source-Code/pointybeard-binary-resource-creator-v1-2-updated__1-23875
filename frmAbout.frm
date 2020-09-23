VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About SRI"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2610
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3075
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Special thanks to Lucky's VB gaming site for the initial code idea."
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Website: http://srinteractive.cjb.net"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact: al_kearney@mailcity.com"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright SRInteractive 2001"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Programed by Alistair Kearney"
      Height          =   255
      Left            =   4095
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Binary Resouce Creator 1.2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   2880
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
CenterForm Me
End Sub

