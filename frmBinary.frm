VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBinary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary Resource Creator 1.2 - www.srinteractive.cjb.net"
   ClientHeight    =   3465
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7425
   Icon            =   "frmBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddAll 
      Caption         =   "Add All >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3735
      TabIndex        =   9
      Top             =   600
      Width           =   1385
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "< Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3730
      TabIndex        =   8
      Top             =   960
      Width           =   1385
   End
   Begin MSComDlg.CommonDialog comdiag1 
      Left            =   4200
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "<< Remove All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3730
      TabIndex        =   7
      Top             =   1320
      Width           =   1385
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3730
      TabIndex        =   6
      Top             =   240
      Width           =   1385
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   5160
      TabIndex        =   5
      Top             =   0
      Width           =   2250
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   3045
      Width           =   1920
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   2070
      TabIndex        =   3
      Top             =   -30
      Width           =   1620
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   105
      TabIndex        =   2
      Top             =   -15
      Width           =   1905
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3730
      TabIndex        =   1
      Top             =   1680
      Width           =   1385
   End
   Begin VB.CommandButton cmdCombine 
      Caption         =   "Combine Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3730
      TabIndex        =   0
      Top             =   2040
      Width           =   1385
   End
   Begin VB.Menu mnuopts 
      Caption         =   "Options"
      Begin VB.Menu mnuDelFiles 
         Caption         =   "Delete Files after combining"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FILEINFO
    fileNamearray() As String 'Holds file names (no paths)
    filePatharray() As String 'holds File paths (No names)
End Type

Dim Files As FILEINFO
Dim TempFiles As FILEINFO


Private Sub cmdAdd_Click()
    Call AddItem
End Sub

Private Sub cmdAddAll_Click()
    Call AddAllItems
End Sub

Private Sub cmdClear_Click()

'Clear the list of files
List1.Clear

'Clear the arrays
ReDim Files.filePatharray(0)
ReDim Files.fileNamearray(0)
End Sub

Private Sub cmdCombine_Click()
'If no files were selected then say so
If List1.ListCount <= 0 Then MsgBox "Please choose files to compile", vbCritical, "Warning": Exit Sub

'Give the user the options of selecting the destination directory
comdiag1.ShowSave

If mnuDelFiles.Checked = True Then
    'if the user wants the input files deleted then let the function know
    BinaryCombine Files.filePatharray, comdiag1.filename, Files.fileNamearray, True
Else
    'opposite of above
    BinaryCombine Files.filePatharray, comdiag1.filename, Files.fileNamearray, False
End If

End Sub

Private Sub cmdExtract_Click()

'If not binary file was selected then let the user know
If File1.filename = "" Or File1.filename = " " Then MsgBox "Please choose a binary file to extract files from", vbCritical, "Error": Exit Sub

frmBrowse.Show 1, Me
'Extract all the files to the directory that the binary file is in.
If OKCancel = True Then BinaryExtract File1.Path & "\" & File1.filename, ExtractPath, True

End Sub

Private Sub cmdRemove_Click()
Dim i As Long
recheck:
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        List1.RemoveItem (i)
        RemoveFile i
        GoTo recheck
    End If
Next i
List1.Refresh
End Sub

Private Sub RemoveFile(lFileNum As Long)
Dim i

ReDim TempFiles.fileNamearray(UBound(Files.fileNamearray))
ReDim TempFiles.filePatharray(UBound(Files.filePatharray))

For i = LBound(Files.fileNamearray) To UBound(Files.fileNamearray)
    TempFiles.fileNamearray(i) = Files.fileNamearray(i)
    TempFiles.filePatharray(i) = Files.filePatharray(i)
Next

Dim CurrNum
Dim CurrNum2

CurrNum = 0
CurrNum2 = 0

For i = LBound(Files.fileNamearray) To UBound(Files.fileNamearray)
     If i <> lFileNum Then
        Files.fileNamearray(CurrNum2) = TempFiles.fileNamearray(CurrNum)
        Files.filePatharray(CurrNum2) = TempFiles.filePatharray(CurrNum)
        CurrNum2 = CurrNum2 + 1
     End If
     
     CurrNum = CurrNum + 1
Next

ReDim Preserve Files.filePatharray(UBound(Files.fileNamearray) - 1)
ReDim Preserve Files.fileNamearray(UBound(Files.fileNamearray) - 1)
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
    Call AddItem
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuDelFiles_Click()
    If Not mnuDelFiles.Checked = True Then mnuDelFiles.Checked = True Else mnuDelFiles.Checked = False
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelp_Click()
frmHelp.Show
End Sub

Private Sub AddItem()
'Add the filename to the list of files to be combined
List1.AddItem File1.filename

'Resize the arrays that hold the file information
ReDim Preserve Files.filePatharray(List1.ListCount - 1)
ReDim Preserve Files.fileNamearray(List1.ListCount - 1)

'Populate the arrays with the latest info
Files.fileNamearray(List1.ListCount - 1) = "\" & File1.filename
Files.filePatharray(List1.ListCount - 1) = File1.Path
End Sub

Private Sub AddAllItems()

For i = 0 To File1.ListCount - 1
    File1.Selected(i) = True

    'Add the filename to the list of files to be combined
    List1.AddItem File1.filename
    
    'Resize the arrays that hold the file information
    ReDim Preserve Files.filePatharray(List1.ListCount - 1)
    ReDim Preserve Files.fileNamearray(List1.ListCount - 1)
    
    'Populate the arrays with the latest info
    Files.fileNamearray(List1.ListCount - 1) = "\" & File1.filename
    Files.filePatharray(List1.ListCount - 1) = File1.Path

Next
End Sub
