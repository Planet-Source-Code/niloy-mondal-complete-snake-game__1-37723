VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Snake"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdelay 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Text            =   "1"
      Top             =   6360
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10800
      Top             =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Increase delay to slow down game speed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Delay  :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.Shape shpfood 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   9600
      Top             =   2880
      Width           =   255
   End
   Begin VB.Shape shpsnake 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2640
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer :- Niloy Mondal. Email - niloygk@yahoo.com"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   5760
      Width           =   5295
   End
   Begin VB.Label lblfoodcount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label lblfood 
      Caption         =   "Food Eaten:-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Shape shprect 
      Height          =   6000
      Left            =   240
      Top             =   240
      Width           =   10800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Snake"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1395
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   5235
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim button, i, randomtop, randomleft, newgame As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight And button <> vbKeyLeft Then
    button = vbKeyRight
End If
If KeyCode = vbKeyLeft And button <> vbKeyRight Then
    button = vbKeyLeft
End If
If KeyCode = vbKeyUp And button <> vbKeyDown Then
    button = vbKeyUp
End If
If KeyCode = vbKeyDown And button <> vbKeyUp Then
    button = vbKeyDown
End If
End Sub

Private Function set_new_game()
lblfoodcount.Caption = "0"
shpsnake(0).Left = 2640
shpsnake(0).Top = 3600
shpfood.Left = 5280
shpfood.Top = 3600
button = 0
Do                  'Remove every dynamically created snake body piece
    If shpsnake.UBound = 0 Then
        Exit Function
    Else
        Unload shpsnake(shpsnake.UBound)
    End If
Loop
End Function

Private Function detect_collision()
If shpsnake(0).Left < shprect.Left Or shpsnake(0).Left >= shprect.Width + 240 Then
    newgame = MsgBox("You have collided with the wall, Start a new game?", vbYesNo, "Game Over")
    If newgame = vbYes Then
        set_new_game
    Else
        End
    End If
End If
If shpsnake(0).Top < shprect.Top Or shpsnake(0).Top >= shprect.Height + 240 Then
    newgame = MsgBox("You have collided with the wall,Start a new game?", vbYesNo, "Game Over")
    If newgame = vbYes Then
        set_new_game
    Else
        End
    End If
End If
End Function

Private Function detect_self_collision()
For i = 1 To shpsnake.UBound
    If shpsnake(0).Top = shpsnake(i).Top And shpsnake(0).Left = shpsnake(i).Left Then
        newgame = MsgBox("You have collided yourself, Start a new Game?", vbYesNo, "Game Over")
        If newgame = vbYes Then
            set_new_game
            Exit Function
        Else
            End
        End If
    End If
Next i
End Function

Private Sub Timer1_Timer()
If txtdelay.Text = "" Then
    Timer1.Interval = 1
Else
    If Int(txtdelay.Text) > 0 Then
        Timer1.Interval = Int(txtdelay.Text)
    Else
        Timer1.Interval = 1
    End If
End If
'Moves the snake horizontally
For i = shpsnake().UBound To 1 Step -1
    shpsnake(i).Left = shpsnake(i - 1).Left
Next i
If button = vbKeyRight Then
    shpsnake(0).Left = shpsnake(0).Left + 240
End If
If button = vbKeyLeft Then
    shpsnake(0).Left = shpsnake(0).Left - 240
End If
'Moves the snake vertically
For i = shpsnake().UBound To 1 Step -1
    shpsnake(i).Top = shpsnake(i - 1).Top
Next i
If button = vbKeyDown Then
    shpsnake(0).Top = shpsnake(0).Top + 240
End If
If button = vbKeyUp Then
    shpsnake(0).Top = shpsnake(0).Top - 240
End If
detect_collision        'detects collision with wall
detect_self_collision   'detects self collision
'Checks if snake has eaten the food and increase the snake's size
If shpfood.Top = shpsnake(0).Top And shpfood.Left = shpsnake(0).Left Then
    Load shpsnake(shpsnake.UBound + 1)
    shpsnake(shpsnake.UBound).Visible = True
    Randomize
    randomleft = Int((Rnd * 10800) + 240)
    randomleft = Int(randomleft / 240)
    randomleft = randomleft * 240
    randomtop = Int((Rnd * 6000) + 240)
    randomtop = Int(randomtop / 240)
    randomtop = randomtop * 240
    shpfood.Top = randomtop
    shpfood.Left = randomleft
    lblfoodcount = lblfoodcount + 1
End If
End Sub
