VERSION 5.00
Begin VB.Form frmPalette 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Palette"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "Palette.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3800
      Left            =   0
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   0
      Top             =   0
      Width           =   3800
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long, j As Long, val As Long
    'Use 2 loops to make a lovely grid
    For i = 0 To 15
        For j = 0 To 15
            'Get the fill value
            picPal.FillColor = LongPal((j + 1) + (i * 15))
            'Draw the box
            picPal.Line (j * 15, i * 15)-((j + 1) * 15, (i + 1) * 15), vbBlack, B
        Next j
    Next i
End Sub
