VERSION 5.00
Begin VB.Form MainWindow 
   Caption         =   "The Window"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Expand_Over 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   6330
      Picture         =   "Window.frx":0000
      ScaleHeight     =   765
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Reverse_Over 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   6510
      Picture         =   "Window.frx":00A0
      ScaleHeight     =   765
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Reverse 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   6150
      Picture         =   "Window.frx":0138
      ScaleHeight     =   765
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Expand 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   5970
      Picture         =   "Window.frx":01D2
      ScaleHeight     =   765
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Expander 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5
      Picture         =   "Window.frx":026A
      ScaleHeight     =   705
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   1770
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   "Put your program here, or whatever."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   1320
      TabIndex        =   5
      Top             =   1260
      Width           =   4605
   End
   Begin VB.Shape SideBar 
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4725
      Left            =   0
      Top             =   -60
      Width           =   105
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Expander_Click()

Toggle

If ShowPic = "Expand" Then
Expander.Picture = Expand.Picture
End If
If ShowPic = "Reverse" Then
Expander.Picture = Reverse.Picture
End If

End Sub

Private Sub Expander_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseOver

End Sub

Sub MouseAway()

If ShowPic = "Expand" Then
Expander.Picture = Expand.Picture
End If
If ShowPic = "Reverse" Then
Expander.Picture = Reverse.Picture
End If

End Sub


Sub MouseOver()

If ShowPic = "Expand" Then
Expander.Picture = Expand_Over.Picture
End If
If ShowPic = "Reverse" Then
Expander.Picture = Reverse_Over.Picture
End If

End Sub

Private Sub Form_Load()

' init with Reverse picture
INorOUT = "IN"
ShowPic = "Reverse"
Expander.Picture = Reverse.Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseAway
End Sub

Private Sub Form_Resize()
On Error Resume Next
SideBar.Height = Me.Height + 100
Expander.Top = (Me.Height / 2) - 560
End Sub
