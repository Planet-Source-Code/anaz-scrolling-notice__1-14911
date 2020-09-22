VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   16545
   ClientTop       =   15360
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Tag             =   "Not required"
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3240
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move mouse over these click ; check the code and use it whereever you need"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please give your comment and vote for me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   248
      TabIndex        =   1
      ToolTipText     =   "Dont forget me"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You have no new messages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   630
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   2010
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000018&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   8
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'For helping to declare all variables
'To open the site.This is not needed for the animation
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
Form1.Width = 3270
Form1.Height = 2700 'So that only the required part of the form is shown even if
                    'we change size accidently during design

Label1.MouseIcon = LoadResPicture(101, 1) 'Custom Mouse icon from resource

Form1.ScaleMode = 1 'Scale 1-Twip
Form1.WindowState = 0 '0-Normal window

Form1.Left = Screen.Width - Form1.Width
Form1.Top = Screen.Height 'we have to position the form above the tray

Form1.Height = 0 'Reset height to 0. We then have to increase it from 0 to max ht.
Form1.Visible = True 'Or else animation may not work

C_Ofrm Me, 1, False 'False to open and true to close form

'A simple FOR loop may too will do the trick
Timer1.Interval = 4000 'The time after which to unload the form
End Sub


Public Function C_Ofrm(frm As Form, Speed As Integer, tag As Boolean)
If Speed = 0 Then
    Exit Function 'The form will not be Closed/Opened
End If

If tag Then
    Do Until frm.Height <= 5
        DoEvents
        frm.Height = frm.Height - Speed * 5
        frm.Top = frm.Top + Speed * 5
    Loop
    Unload frm
Else
    Do Until frm.Height >= 2700
        DoEvents
        frm.Height = frm.Height + Speed * 5
        frm.Top = frm.Top - Speed * 5
    Loop
End If

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(0, 0, 255)
'Might have finished reading ?
Timer1.Enabled = True
End Sub

Private Sub Label1_Click()
Dim conSwNormal As Long
ShellExecute hWnd, "open", "http://the hyperlink here", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(255, 0, 0)
'Stop the timer so that the form doesnt unload while reading
Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()
    C_Ofrm Me, 1, True 'Close or unload form
End Sub

Private Sub Timer2_Timer()
'Not require by the project;but you can find any use,probabily like this :-)
DoEvents
Label2.ForeColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd) 'Blink the text
End Sub
