VERSION 5.00
Begin VB.UserControl uclBrowseFile 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   Begin VB.CheckBox chkOptionClick 
      Caption         =   "single click folder"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.FileListBox filFiles 
      Height          =   1455
      Hidden          =   -1  'True
      Left            =   0
      System          =   -1  'True
      TabIndex        =   2
      Top             =   2610
      Width           =   2415
   End
   Begin VB.DirListBox dirFolder 
      Height          =   1665
      Left            =   30
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox cboPattern 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2130
      Width           =   2400
   End
End
Attribute VB_Name = "uclBrowseFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//uclBrowseFile v0.01
'//quick/simple file browser
'//by Mil-X Pro
'//07 april 2003
'//i give this usercontrol f.o.c, take it and make
'//it yours. i don't mind if you claim it's yours,
'//but i hope it might help beginner that interested
'//in creating active-x. other than that, maybe
'//this usercontrol can simplified your project
'//that need drivelistbox/dirlistbox/filelistbox control,
'//who knows...
'//but don't put a blame on me if this usercontrol
'//spoil your machine...

'//event declarations
Event FileClick()

Public Function exact_file_path() As String
'//check if file in directory
Dim File As String

If filFiles.FileName = "" Then Exit Function

If Len(filFiles.Path) > 3 Then
    '// file in directory
    File = filFiles.Path & "\" & filFiles.FileName
Else
    '//file not in directory
    File = filFiles.Path & filFiles.FileName
End If

exact_file_path = File

End Function

Private Sub cboPattern_Click()
'//change filelistbox pattern(filetype)
Dim ptrn As String
Dim pat1 As Integer
Dim pat2 As Integer

ptrn = cboPattern.List(cboPattern.ListIndex)
pat1 = InStr(ptrn, "(")
pat2 = InStr(ptrn, ")")

filFiles.Pattern = Mid$(ptrn, pat1 + 1, pat2 - pat1 - 1)

End Sub

Private Sub dirFolder_Change()
'//list selected folder's contents in filelistbox
filFiles.Path = dirFolder.Path

End Sub

Private Sub dirFolder_Click()
'//list selected folder's contents in filelistbox,
'//use single click, except if the checkbox is
'//checked than this command will not execute
If chkOptionClick.Value = 1 Then Exit Sub
dirFolder.Path = dirFolder.List(dirFolder.ListIndex)

End Sub

Private Sub drvDrive_Change()
'//list selected drive's folders in dirlistbox,
'//if error occur such as "floppy a" is empty
'//then prompt user about the error
On Error GoTo no_crash

dirFolder.Path = drvDrive.Drive

Exit Sub
'//this command will trap the error
'//so your app will not crash
no_crash:
MsgBox "Drive reading error." & vbCr & _
    "Error " & Err.Number & " : " & Err.Description, _
    vbCritical, App.Title

End Sub

Private Sub filFiles_Click()
'//activate FileClick event this event will be
'//listed in the usercontrol dropdown list
RaiseEvent FileClick

End Sub

Private Sub UserControl_Initialize()
'//add combobox with these items,
'//change these items to suit your project
cboPattern.AddItem "All Graphics Files (*.bmp;*.jpg;*.gif)"
cboPattern.AddItem "All Text Files (*.txt;*.log;*.ini;*.diz)"
cboPattern.ListIndex = 0

End Sub

Private Sub UserControl_Resize()
'//resize all controls
'//except all error, silly code
On Error Resume Next
'//prevent user from resize the usercontrol too small
If UserControl.Width < 2000 Then UserControl.Width = 2000
If UserControl.Height < 2000 Then UserControl.Height = 2000
'//now place all controls nicely if the usercontrol
'//is resizing, using simple/logic maths
drvDrive.Move 2, 2, ScaleWidth - 4
dirFolder.Move 2, 4 + drvDrive.Height, ScaleWidth - 4, _
    (ScaleHeight / 2) - cboPattern.Height
cboPattern.Move 2, 2 + dirFolder.Top + dirFolder.Height, _
    ScaleWidth - 4
filFiles.Move 2, 2 + cboPattern.Top + cboPattern.Height, _
    ScaleWidth - 4, ScaleHeight - cboPattern.Top - _
    cboPattern.Height - 2

End Sub

Public Property Get TheFile() As String
'//define "property get" for TheFile
TheFile = exact_file_path
    
End Property

Public Property Let TheFile(ByVal New_TheFile As String)
'//define "property let" for TheFile
exact_file_path = New_TheFile
PropertyChanged "TheFile"

End Property
