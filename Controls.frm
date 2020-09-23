VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "You can remove this form from the project"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   855
      Left            =   2790
      TabIndex        =   1
      Top             =   2220
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Toggle"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   1110
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "You can also toggle the form to Expand and Reverse, by clicking the little arrow on the left side of the Main Window form."
      Height          =   1035
      Left            =   270
      TabIndex        =   4
      Top             =   2010
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "You can delete this form. You probably won't need it. It's just here to test the thing."
      Height          =   585
      Left            =   1830
      TabIndex        =   3
      Top             =   1110
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Use this to test the window. In case it hides from you, just hit the End button, and edit your dimensions."
      Height          =   555
      Left            =   420
      TabIndex        =   2
      Top             =   390
      Width           =   2865
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Toggle

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
MainWindow.Show
With MainWindow
    .Top = 0
    .Left = Screen.Width - .Width
End With
INorOUT = "IN" ' Init at IN
End Sub
