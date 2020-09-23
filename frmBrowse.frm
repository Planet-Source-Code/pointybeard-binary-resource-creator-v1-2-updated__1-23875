VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Browse for folder"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Choose a save path."
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
OKCancel = False
Unload Me

End Sub

Private Sub cmdOK_Click()
OKCancel = True
ExtractPath = Dir1.Path
MsgBox ExtractPath
Unload Me
End Sub

Private Sub Dir1_Change()
cmdOK.Enabled = True
End Sub

Private Sub Form_Load()
Dir1.Path = "c:\"
End Sub
