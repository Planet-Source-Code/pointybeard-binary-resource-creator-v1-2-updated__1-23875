VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Binary Resouce Creator Help"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbHelpText 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   12726
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmHelp.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
rtbHelpText.filename = App.Path & "\readme.txt"
CenterForm Me
End Sub

