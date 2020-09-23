VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3120
      Left            =   4440
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3150
      Left            =   0
      ScaleHeight     =   206
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   294
      TabIndex        =   0
      Top             =   30
      Width           =   4470
      Begin VB.Image imgBrowse 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   810
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   1110
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
    
Private Filenames() As String
Private Path As String
Private Suffixes As String

Private Sub Form_Load()
    ReDim Preserve Filenames(0 To 0)
End Sub

Private Sub Form_Resize()
    VScroll1.Move ScaleWidth - VScroll1.Width, 0, VScroll1.Width, ScaleHeight
    picBackground.Move 0, 0, ScaleWidth - VScroll1.Width, ScaleHeight
    DisplayPreviews Filenames
End Sub

Private Sub imgBrowse_Click(Index As Integer)
    frmPreview.picPreview.Picture = LoadPicture(imgBrowse(Index).Tag)
    frmPreview.Show vbModal
End Sub

Private Sub mnuFileOpen_Click()
    'Build a list of files to process.
    Path = GetFolderName
    If Len(Path) = 0 Then
        Exit Sub
    End If
    Suffixes = ".bmp,.gif,.jpg"
    Filenames = BuildFileList(Path, Suffixes)
    If UBound(Filenames) = 0 Then
        MsgBox "No picture files found in that directory!"
        Exit Sub
    End If
    DisplayPreviews Filenames
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Function GetFolderName() As String
    'Opens a Treeview control that displays the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "This is the title"

    With tBrowseInfo
        .hWndOwner = 0 'Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    
    GetFolderName = sBuffer
End Function

Private Function BuildFileList(ByVal Pathname As String, ByVal Suffixes As String) As String()
    Dim s() As String, FileName As String, Suffix As String, Max As Long
    
    If Right$(Pathname, 1) <> "\" Then
        Pathname = Pathname & "\"
    End If
    FileName = Dir(Pathname & "*.*", vbNormal Or vbArchive)
    
    ReDim s(0 To 0)
        
    Do While FileName > ""
        If Len(FileName) > 3 Then
            Suffix = LCase(Right$(FileName, 4))
            If InStr(1, Suffixes, Suffix) > 0 Then
                'it has a correct suffix for loading in vb
                Max = UBound(s, 1) + 1
                ReDim Preserve s(0 To Max)
                s(Max) = Pathname & FileName
            End If
        End If
        FileName = Dir()
    Loop
    BuildFileList = s
End Function

Private Sub DisplayPreviews(Filenames() As String)
    Dim i As Long, FormWidth As Long, FormHeight As Long
    Dim PicturesAcross As Long, PicturesDown As Long
    Dim TheCol As Long, TheRow As Long
    Dim BrowseWidth As Long, BrowseHeight As Long
    On Error GoTo Err_Init
    
    'Remove old pictures (form may have resized)
    For i = 1 To imgBrowse.Count - 1
        imgBrowse(i).Picture = LoadPicture()
        imgBrowse(i).Visible = False
    Next i
    
    'Make sure we have some files to process
    If UBound(Filenames, 1) = 0 Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Determine how many picture previews will fit across and down.
    FormWidth = picBackground.Width
    FormHeight = picBackground.Height
    BrowseWidth = imgBrowse(0).Width
    BrowseHeight = imgBrowse(0).Height
    PicturesAcross = FormWidth \ BrowseWidth
    PicturesDown = FormHeight \ BrowseHeight
    
    'Set the scrollbar settings
    With VScroll1
        .Max = PicturesDown * PicturesAcross
        If .Max > UBound(Filenames, 1) Then
            .Max = UBound(Filenames, 1)
        End If
        .Min = 1
        .LargeChange = PicturesDown
        .SmallChange = 1
    End With
    
    'Load the picture previews into the image controls.
    For i = ((VScroll1.Value - 1) * PicturesAcross) + 1 To ((VScroll1.Value - 1) * PicturesAcross) + (PicturesAcross * PicturesDown)
        If i > UBound(Filenames, 1) Then
            'exit out early if we don't have enough pictures to fill form
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Load imgBrowse(i)
        imgBrowse(i).Picture = LoadPicture(Filenames(i))
        imgBrowse(i).Tag = Filenames(i)
        imgBrowse(i).ToolTipText = Filenames(i)
        imgBrowse(i).Move (TheCol * BrowseWidth), (TheRow * BrowseHeight)
        imgBrowse(i).Visible = True
        TheCol = TheCol + 1
        If TheCol > (PicturesAcross - 1) Then
            TheCol = 0
            TheRow = TheRow + 1
        End If
    Next i
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Err_Init:
    If Err.Number = 360 Then
        'control already loaded - no biggie.
        Resume Next
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub VScroll1_Change()
    DisplayPreviews Filenames()
End Sub
