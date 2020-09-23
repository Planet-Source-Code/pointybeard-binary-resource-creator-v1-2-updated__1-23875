Attribute VB_Name = "modBinary"
'This structure will describe our binary file's
'size and number of contained files
Private Type FILEHEADER
    intNumFiles As Integer      'How many files are inside?
    lngFileSize As Long         'How big is this file? (Used to check integrity)
End Type

'This structure will describe each file contained
'in our binary file
Private Type INFOHEADER
    lngFileSize As Long         'How big is this chunk of stored data?
    lngFileStart As Long        'Where does the chunk start?
    strFileName As String * 16  'What's the name of the file this data came from?
End Type

Private Type BYTEDATA
    bytData() As Byte 'Holds the data of each file.
End Type

Public OKCancel As Boolean
Public ExtractPath As String

'-------------------------------
'Extract files from a binary file. Must have
'Been created with this program or the same code.
'INPUT: sfile - Location of binary file
'       sdest - Destination diectory for extracted files
'       ExtractFiles - If true then files are extracted. If False then a list of files
'                       in the binary file is put in the list box. (Not properly implemented)
'--------------------------------

Public Sub BinaryExtract(sFile As String, sDest As String, ExtractFiles As Boolean)

On Local Error GoTo errout

Dim i As Integer
Dim intSampleFile As Integer
Dim intBinaryFile As Integer
Dim bytSampleData() As Byte
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
    

    'Open the binary file
    intBinaryFile = FreeFile
    Open sFile For Binary Access Read Lock Write As intBinaryFile
    
    'Extract the FILEHEADER
    Get intBinaryFile, 1, FileHead
    
    'Check the file for validity
    If LOF(intBinaryFile) <> FileHead.lngFileSize Then
        MsgBox "This is not a valid file format.", vbOKOnly, "Invalid File"
        Exit Sub
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.intNumFiles - 1)
  
    'Extract the INFOHEADER
    Get intBinaryFile, , InfoHead

If ExtractFiles = True Then
    'Extract all of the files from the binary file
    For i = 0 To UBound(InfoHead)
        'Resize the byte data array
        ReDim bytSampleData(InfoHead(i).lngFileSize - 1)
        'Get the data
        Get intBinaryFile, InfoHead(i).lngFileStart, bytSampleData
        'Open a new file and store the data
        intSampleFile = FreeFile
        Open sDest & InfoHead(i).strFileName For Binary Access Write Lock Write As intSampleFile
        Put intSampleFile, 1, bytSampleData
        Close intSampleFile
    Next
Else

    frmBinary.List1.Clear
    For i = LBound(InfoHead) To UBound(InfoHead)
        frmBinary.List1.AddItem InfoHead(i).strFileName
    Next i
End If

'Close the binary file
    Close intBinaryFile
    Exit Sub
errout:

MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Warning"
End Sub

'-------------------------------
'Combine files into a binary file.
'INPUT: PathToFiles() - an array to hold all the file PATHS (not names) for the files to be combines
'       sDestintion - Destination diectory for the binary file
'       sSource() - an array to hold all the file NAMES (not path) for the files to be combines
'       DelFiles - Flag to tell the fuction to delete the source files once combined.
'--------------------------------

Public Sub BinaryCombine(PathToFile() As String, ByVal sDestination As String, sSource() As String, DelFiles As Boolean)

On Local Error GoTo errout

Dim intFile() As Integer
Dim intBinaryFile As Integer
Dim bytFileData() As BYTEDATA
Dim FileHead As FILEHEADER
Dim InfoHead() As INFOHEADER
Dim lngFileStart As Long
Dim i

'resize the arrays
ReDim intFile(UBound(sSource))
ReDim bytFileData(UBound(sSource))

    For i = LBound(sSource) To UBound(sSource)
        'Find some free file numbers to use and open the files
        intFile(i) = FreeFile
        Open PathToFile(i) & sSource(i) For Binary Access Read Lock Write As intFile(i)
    Next
    
    'Find out how large the files are and
    'resize the data arrays appropriately
    For i = LBound(bytFileData) To UBound(bytFileData)
        With bytFileData(i)
            'MsgBox LOF(intFile(i)) - 1
            ReDim .bytData(LOF(intFile(i)))
        End With
        Get intFile(i), 1, bytFileData(i).bytData
    Next
   
    If DelFiles = True Then
        For i = LBound(sSource) To UBound(sSource)
            'Close and delete the files
            Close intFile(i)
            Kill PathToFile(i) & sSource(i)
        Next
    End If
    
    'Set up the file header
    FileHead.intNumFiles = UBound(sSource) + 1
    
    For i = LBound(bytFileData) To UBound(bytFileData)
        FileHead.lngFileSize = FileHead.lngFileSize + (UBound(bytFileData(i).bytData) + 1)
    Next
    
    FileHead.lngFileSize = FileHead.lngFileSize + (6) + (FileHead.intNumFiles * 24)
   
    'Set up the info headers
    ReDim InfoHead(FileHead.intNumFiles - 1)
    lngFileStart = (6) + (FileHead.intNumFiles * 24) + 1

    For i = LBound(bytFileData) To UBound(bytFileData)
        InfoHead(i).lngFileSize = UBound(bytFileData(i).bytData) + 1
        InfoHead(i).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(i).lngFileSize
    Next
    
    For i = LBound(sSource) To UBound(sSource)
        InfoHead(i).strFileName = sSource(i)
    Next
    
    'Open a new file
    intBinaryFile = FreeFile
    Open sDestination For Binary Access Write Lock Write As intBinaryFile
    
    'Store the data in the file
    Put intBinaryFile, 1, FileHead
    Put intBinaryFile, , InfoHead
    
    For i = LBound(bytFileData) To UBound(bytFileData)
        Put intBinaryFile, , bytFileData(i).bytData
    Next
    
    'Close the file
    Close intBinaryFile
    Exit Sub
errout:

MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Warning"
End Sub

Sub CenterForm(Frm As Form)

'Input:  CenterForm Me

Frm.Move (Screen.Width - Frm.Width) \ 2, (Screen.Height - Frm.Height) \ 2

End Sub
