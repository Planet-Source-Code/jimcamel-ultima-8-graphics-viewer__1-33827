VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ultima 8 Graphics Viewer"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit Program"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   285
      Left            =   3120
      Picture         =   "frmMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   285
   End
   Begin VB.ComboBox cmbFrame 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cmbType 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":1047
      Left            =   720
      List            =   "frmMain.frx":1049
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton cmdShowPal 
      Caption         =   "Show Palette"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load File"
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "JimCamel 2002"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Frame:"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lbl1 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbFrame_Click()
'When someone selects the frame, draw it
DrawImage cmbType.ListIndex, cmbFrame.ListIndex
End Sub

Private Sub cmbType_Click()
Dim i As Long
'clear the listbox
cmbFrame.Clear
'Read the frames with the arguement of the current value of cmbtype
'and with the value returned...
For i = 0 To ReadFrames(cmbType.ListIndex)
    '...add all the numbers to the fram listbox
    cmbFrame.AddItem (i)
Next i
'and set the index to 0
cmbFrame.ListIndex = 0
End Sub

Private Sub cmdExit_Click()
'End the program.... right
End
End Sub

Private Sub cmdLoad_Click()
Dim i As Long
'empty the type list
cmbType.Clear
'enable the list boxes
cmbType.Enabled = True: cmbFrame.Enabled = True
'Load the palette file (which SHOULD be in the same directory, called u8pal.pal
LoadPal (Left(txtPath.Text, InStrRev(txtPath.Text, "\")) & "u8pal.pal")
'Call openfile with the arguement of the filepath, and with the returned value
For i = 0 To OpenFile(txtPath.Text)
    'add the numbers to the type listbox
    cmbType.AddItem i
Next i
'and set it's value to 0
cmbType.ListIndex = 0
End Sub

Private Sub cmdOpen_Click()
'Declare a new FileDialog Class
Dim FD As New clsFileDialog
'Set the init directory to that in txtPath
FD.InitDir = txtPath.Text
'Set the filter to the Graphics file format
FD.Filter = "Ultima 8 Graphics file (FLX)|*.flx|All files (*.*)|*.*"
'Show the FileDialog
FD.ShowOpen
'Set the value to txtPath
txtPath.Text = FD.Filename
'Clear the memory out
Set FD = Nothing
End Sub

Private Sub cmdShowPal_Click()
'Shows the Palette form
frmPalette.Show
End Sub

Private Sub Form_Load()
'Clear everything out
cmbFrame.Clear: cmbType.Clear
txtPath.Text = App.Path
End Sub
