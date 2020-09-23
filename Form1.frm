VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   556
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   135
      TabIndex        =   15
      Top             =   405
      Width           =   2970
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   480
      Left            =   5790
      TabIndex        =   13
      Top             =   2130
      Width           =   1260
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7680
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7635
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.OptionButton Option1 
      Height          =   450
      Index           =   0
      Left            =   810
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      Height          =   450
      Index           =   1
      Left            =   1290
      Picture         =   "Form1.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      Height          =   450
      Index           =   2
      Left            =   1770
      Picture         =   "Form1.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      Height          =   450
      Index           =   3
      Left            =   2250
      Picture         =   "Form1.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   480
      Left            =   5775
      TabIndex        =   7
      Top             =   1590
      Width           =   1260
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add all Files"
      Height          =   480
      Left            =   5775
      TabIndex        =   5
      Top             =   510
      Width           =   1260
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7530
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   1785
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   135
      TabIndex        =   3
      Top             =   750
      Width           =   2985
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7710
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   2475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   3135
      TabIndex        =   1
      Top             =   405
      Width           =   2340
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add File"
      Height          =   480
      Left            =   5775
      TabIndex        =   0
      Top             =   1050
      Width           =   1260
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2880
      Left            =   30
      TabIndex        =   14
      Top             =   2895
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5080
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   158750
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "View"
      Height          =   210
      Left            =   255
      TabIndex        =   12
      Top             =   2565
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   45
      TabIndex        =   6
      Top             =   150
      Width           =   4440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Right$(File1.Path, 1) = "\" Then
p$ = File1.Path
Else
p$ = File1.Path & "\"
End If

AddFileTolView hwnd, ListView1, p$ & File1.List(File1.ListIndex), ImageList2, ImageList1, Picture2, Picture1








Label1.Caption = "Files Added: " & ListView1.ListItems.Count & " images in ImageList: " & ImageList1.ListImages.Count

End Sub


Private Sub Command2_Click()
Dim StartTime As Date
StartTime = Time
MousePointer = 11
retVal& = LockWindowUpdate(hwnd)
If Right$(File1.Path, 1) = "\" Then
p$ = File1.Path
Else
p$ = File1.Path & "\"
End If

For i% = 0 To File1.ListCount - 1
AddFileTolView hwnd, ListView1, p$ & File1.List(i%), ImageList2, ImageList1, Picture2, Picture1
Label1.Caption = "Files Added: " & ListView1.ListItems.Count & ", images in ImageList: " & ImageList1.ListImages.Count
Label1.Refresh
Next i%
retVal& = LockWindowUpdate(&O0)
ListView1.Arrange = lvwAutoTop
MousePointer = 0
Label1.Caption = Label1.Caption & " in " & DateDiff("s", StartTime, Time) & "seconds"






End Sub


Private Sub Command3_Click()
Unload Me

End Sub


Private Sub Command4_Click()
ListView1.Icons = Nothing
ListView1.SmallIcons = Nothing
ImageList1.ListImages.Clear
ImageList2.ListImages.Clear
ListView1.ListItems.Clear


End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub


Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()
Option1(0).Value = True
ImageList1.MaskColor = Picture1.BackColor
ImageList2.MaskColor = Picture2.BackColor

End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Move ListView1.Left, ListView1.Top, ScaleWidth - 5, ScaleHeight - 5 - ListView1.Top


End Sub


Private Sub Option1_Click(Index As Integer)
ListView1.View = Index
End Sub


