VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   4995
   ClientTop       =   4395
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   4140
   Begin VB.CommandButton Command2 
      Caption         =   "2. create WRI file"
      Height          =   435
      Left            =   2250
      TabIndex        =   2
      Top             =   90
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1. load picture"
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   75
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   90
      ScaleHeight     =   4110
      ScaleWidth      =   2790
      TabIndex        =   0
      Top             =   630
      Width           =   2790
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String
Dim PicColors(15) As Long

Private Sub Command1_Click()
    'load 16 shades of gray picture to picturebox1
    Dim PictureFile As String
    PictureFile = Path + "Seba.bmp"
    Picture1.Picture = LoadPicture(PictureFile)
End Sub

Private Sub Command2_Click()
    Dim GLine As String
    Dim PicX As Single
    Dim PicY As Single
    Dim PicColor As Double
    Dim OutFile As String
    
    OutFile = Path + "Sebastjan.wri"
    Open OutFile For Output As #1
       
    
    'counts for colors in picture.
    'it will generate an error if there is more then 16 colors
    For PicY = 0 To Picture1.Height - (Screen.TwipsPerPixelY) Step Screen.TwipsPerPixelY * 2
        For PicX = 0 To Picture1.Width - (Screen.TwipsPerPixelX) Step Screen.TwipsPerPixelX
        
            PicColor = Picture1.Point(PicX, PicY) 'get pixle color
            found = False
            For ColCheck = 0 To ColorUsed
                
                If PicColors(ColCheck) = PicColor Then
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                PicColors(ColorUsed) = PicColor
                ColorUsed = ColorUsed + 1
            End If
        Next
    Next
    
    'sorts the PicColors  (palete)
    For sort1 = 0 To 15
        For sort2 = sort1 To 15
            If PicColors(sort1) < PicColors(sort2) Then
                TempBuffer = PicColors(sort1)
                PicColors(sort1) = PicColors(sort2)
                PicColors(sort2) = TempBuffer
            End If
        Next
    Next
    
    
    For PicY = 0 To Picture1.Height - (Screen.TwipsPerPixelY) Step Screen.TwipsPerPixelY * 2
    '<-- every second line is used because... _
    pixle.width     / pixle.height     = 1 / 1   => TRUE _
    character.width / character.height = 1 / 2   => TRUE
        GLine = ""
        For PicX = 0 To Picture1.Width - (Screen.TwipsPerPixelX) Step Screen.TwipsPerPixelX
            PicColor = Picture1.Point(PicX, PicY) 'get pixle color
            
            Select Case PicColor
            Case PicColors(0)          'white
                GLine = GLine + " "
            Case PicColors(1)
                GLine = GLine + "."
            Case PicColors(2)
                GLine = GLine + ","
            Case PicColors(3)
                GLine = GLine + ":"
            Case PicColors(4)
                GLine = GLine + "÷"
            Case PicColors(5)
                GLine = GLine + "¤"
            Case PicColors(6)
                GLine = GLine + "l"
            Case PicColors(7)
                GLine = GLine + "J"
            Case PicColors(8)
                GLine = GLine + "F"
            Case PicColors(9)
                GLine = GLine + "E"
            Case PicColors(10)
                GLine = GLine + "9"
            Case PicColors(11)
                GLine = GLine + "$"
            Case PicColors(12)
                GLine = GLine + "€"
            Case PicColors(13)
                GLine = GLine + "8"
            Case PicColors(14)
                GLine = GLine + "@"
            Case PicColors(15)                'Black
                GLine = GLine + "#"
            End Select
        Next
        Print #1, GLine
    Next
    Close
    x = MsgBox("Done. Now open file " + OutFile + ". To see entire picture...: Select all and reduce font size (to about '4')")
    End
End Sub

Private Sub Form_Load()
    Path = App.Path
    If Right$(Path, 1) <> "\" Then Path = Path + "\"
End Sub
