VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPreview 
      AutoSize        =   -1  'True
      Height          =   2100
      Left            =   315
      ScaleHeight     =   136
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   300
      Width           =   3225
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    picPreview.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
