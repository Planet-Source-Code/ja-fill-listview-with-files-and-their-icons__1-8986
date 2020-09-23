Attribute VB_Name = "Module1"
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
         
Public Enum IconSource
    internal = 0
    associated = 1
    None = 2
End Enum
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Function GetFilename(mFile As String) As String
Dim fs As New FileSystemObject
GetFilename = fs.GetFilename(mFile)
End Function

   Public Function GetShortName(ByVal sLongFileName As String) As String
       'Not used i the sample
       'I used it for a while and I think it is usefull just to be
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
   End Function

Sub AddFileTolView(hwnd As Long, lView As ListView, mFile As String, Icons As ImageList, SmallIcons As ImageList, picLarge As PictureBox, picSmall As PictureBox)
'Adds the mFile to a listview using the index of the imagelists returned by the following function

imgIndex = GetImgListIndex(hwnd, mFile, Icons, SmallIcons, picLarge, picSmall)
Dim it As ListItem
x$ = GetFilename(mFile)
Set it = lView.ListItems.Add(, , x$, imgIndex, imgIndex)

End Sub
Function GetFileExtension(mFile As String) As String
Dim fs As New FileSystemObject
GetFileExtension = fs.GetExtensionName(mFile)
Exit Function
End Function


Function GetIconHandleEx(hwnd As Long, fName As String, extension As String, exeName As String, icSource As IconSource) As Long
'This function will return an Icon handle no matter what
'Also it will return the file from which the icon handle
'was retreived and if it was retreived from the given file
'or another one
'These two last are usefull to avoid duplicate icons in the imagelist

Dim hIcon As Long, nIcon As Long
Dim AssoFile As String * 250
Dim execName As String
Dim ret As Long
Dim IconFileName As String
AssoFile = fName
hIcon = ExtractIcon(hwnd, AssoFile, nIcon)
IconFileName = Trim(AssoFile)
Select Case hIcon
    Case 0
    hIcon = ExtractAssociatedIcon(hwnd, AssoFile, nIcon)
    exeName = Trim(AssoFile)
    icSource = associated
    Case Is > 0
    If UCase(IconFileName) = UCase(fName) Then
    exeName = fName
    icSource = internal
    Else
    exeName = IconFileName
    icSource = associated
    End If
    
End Select
  
GetIconHandleEx = hIcon
   
End Function


Function GetImgListIndex(hwnd As Long, mFile As String, Icons As ImageList, SmallIcons As ImageList, picLarge As PictureBox, picSmall As PictureBox) As Integer
'this function will add an icon to the image lists only if they do not exist
'allready, and will return the index of this icon

Dim hIcon As Long
Dim icSource As IconSource
Dim exeName As String
Dim extension As String
Dim limg1 As ListImage
Dim limg2 As ListImage
Dim tmpImgKey As String

extension = UCase(GetFileExtension(mFile))
hIcon = GetIconHandleEx(hwnd, mFile, extension, exeName, icSource)
'If we have a file that the icon is stored in it then do not check for duplicates
'Just add the image to the imagelist
If icSource = internal Then
    picLarge = LoadPicture
    picSmall = LoadPicture
    ret1& = DrawIconEx(Form1.Picture1.hdc, 0, 0, hIcon, 16, 16, 0, 0, DI_NORMAL)
    ret2& = DrawIconEx(Form1.Picture2.hdc, 0, 0, hIcon, 32, 32, 0, 0, DI_NORMAL)
    DestroyIcon hIcon
    tmpImgKey = exeName
    For i% = 1 To Icons.ListImages.Count
    If Icons.ListImages(i%).Key = exeName Then
    GetImgListIndex = i%
    Exit Function
    End If
    Next i%
    Set limg1 = Icons.ListImages.Add(, tmpImgKey, picLarge.Image)
    Set limg2 = SmallIcons.ListImages.Add(, tmpImgKey, picSmall.Image)
    GetImgListIndex = limg1.Index
    Form1.ListView1.Icons = Icons
    Form1.ListView1.SmallIcons = SmallIcons
    Exit Function
ElseIf icSource = associated Then
'We have a file that the icon is outside so use as key the extension and check for dublicates
    tmpImgKey = "ID" & extension
    For i% = 1 To Icons.ListImages.Count
        If Icons.ListImages(i%).Key = "ID" & extension Then
        DestroyIcon hIcon
        GetImgListIndex = i%
        Exit Function
        End If
    Next i%
    'The item did not found so add it
    picLarge = LoadPicture
    picSmall = LoadPicture
    ret1& = DrawIconEx(picSmall.hdc, 0, 0, hIcon, 16, 16, 0, 0, DI_NORMAL)
    ret2& = DrawIconEx(picLarge.hdc, 0, 0, hIcon, 32, 32, 0, 0, DI_NORMAL)
    DestroyIcon hIcon
    Set limg1 = Icons.ListImages.Add(, tmpImgKey, picLarge.Image)
    Set limg2 = SmallIcons.ListImages.Add(, tmpImgKey, picSmall.Image)
    Form1.ListView1.Icons = Icons
    Form1.ListView1.SmallIcons = SmallIcons
    GetImgListIndex = limg1.Index
    Exit Function
Else
tmpImgKey = "IDNOICON1100"
    For i% = 1 To Icons.ListImages.Count
    If Icons.ListImages(i%).Key = tmpImgKey Then
    GetImgListIndex = i%
    DestroyIcon hIcon
    Exit Function
    End If
    Next i%
    picLarge = LoadPicture
    picSmall = LoadPicture
    ret1& = DrawIconEx(picSmall.hdc, 0, 0, hIcon, 16, 16, 0, 0, DI_NORMAL)
    ret2& = DrawIconEx(picLarge.hdc, 0, 0, hIcon, 32, 32, 0, 0, DI_NORMAL)
    DestroyIcon hIcon
    Set limg1 = Icons.ListImages.Add(, tmpImgKey, picLarge.Image)
    Set limg2 = SmallIcons.ListImages.Add(, tmpImgKey, picSmall.Image)
    Form1.ListView1.Icons = Icons
    Form1.ListView1.SmallIcons = SmallIcons
    GetImgListIndex = limg1.Index
    Exit Function
End If

        
    
End Function


