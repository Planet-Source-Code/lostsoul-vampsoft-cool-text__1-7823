VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Vampire Software: Cool Form!"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Change Colors"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Blue:"
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   7
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Green:"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Red:"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   5
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   4
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "72"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   2
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1, y1 As Integer
Private Sub Command1_Click()
    End
End Sub
Private Sub Command2_Click()
    Dim Red1 As Integer
    Dim Green1 As Integer
    Dim Blue1 As Integer
    
    Randomize
    Red1 = Rnd * 255
    Green1 = Rnd * 255
    Blue1 = Rnd * 255
    Call FormEffect(Form1, (Red1), (Green1), (Blue1), True, True)
    Label1(0) = Red1
    Label1(1) = Green1
    Label1(2) = Blue1
End Sub
Private Sub Form_Click()
    'Determines if the Minimize, and Exit Button have been Pushed
    If x1 >= Me.Width - 485 And x1 <= Me.Width - 300 And y1 >= 75 And y1 <= 255 Then
        Me.WindowState = 1
    End If
    If x1 >= Me.Width - 275 And x1 <= Me.Width - 90 And y1 >= 75 And y1 <= 255 Then
        End
    End If
End Sub
Private Sub Form_Load()
    Randomize
    'Form Caption; Red Color; Green Color; Blue Color; X Button; Min Button
    Dim Red1 As Integer
    Dim Green1 As Integer
    Dim Blue1 As Integer
    
    Randomize
    Red1 = Rnd * 255
    Green1 = Rnd * 255
    Blue1 = Rnd * 255
    Call FormEffect(Form1, (Red1), (Green1), (Blue1), True, True)
    Label1(0) = Red1
    Label1(1) = Green1
    Label1(2) = Blue1
   
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag the Form
    If x1 >= 60 And x1 <= Me.Width - 485 And y1 >= 75 And y1 <= 255 Then
        FormDrag Me
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    x1 = X
    y1 = Y
End Sub

Private Sub Label3_Click()
End Sub

