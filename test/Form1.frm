VERSION 5.00
Object = "*\A..\prjBrowseFile.vbp"
Begin VB.Form Form1 
   Caption         =   "Test Form"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEdit 
      BackColor       =   &H8000000A&
      Height          =   3600
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   4095
   End
   Begin prjBrowseFile.uclBrowseFile uclBrowseFile1 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6482
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3600
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   30
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub open_file(fname As String, tbox As TextBox)
'//load file to textbox
On Error GoTo LoadFileError

Dim fnum As Integer
Dim flen As Long
Dim txt As String

DoEvents
fnum = FreeFile
Open fname For Input As fnum
    txt = Input$(LOF(fnum), #fnum)
Close #fnum
tbox.Text = txt
Exit Sub
LoadFileError:
    Beep
    MsgBox "Error " & Err.Number & " : " & Err.Description
    fname = ""
    tbox.Text = ""
    Exit Sub
    
End Sub
Private Sub Form_Resize()
'//resize usercontrol
uclBrowseFile1.Height = ScaleHeight

End Sub

Private Sub uclBrowseFile1_FileClick()
'//example how to load a file
On Error GoTo dont_blame_me
'//load picture to image1
Image1.Visible = True
txtEdit.Visible = False
Image1.Picture = LoadPicture(uclBrowseFile1.TheFile)
Exit Sub
'//load text to textbox
dont_blame_me:
Image1.Visible = False
txtEdit.Visible = True
Call open_file(uclBrowseFile1.TheFile, txtEdit)

End Sub
