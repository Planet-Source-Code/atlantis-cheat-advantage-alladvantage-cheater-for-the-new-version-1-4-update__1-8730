VERSION 5.00
Begin VB.Form cheatadvantage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   720
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   315
      ScaleWidth      =   2985
      TabIndex        =   8
      Top             =   0
      Width           =   2985
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   2880
         TabIndex        =   9
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      Height          =   495
      Left            =   1680
      Picture         =   "Form1.frx":3A44
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      Height          =   495
      Left            =   360
      Picture         =   "Form1.frx":3DFA
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   0
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2150
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   760
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   640
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   0
      Picture         =   "Form1.frx":41B0
      ScaleHeight     =   1500
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   3030
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "status: disabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   1245
         Width           =   3015
      End
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "cheatadvantage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
End Sub

Private Sub Form_Load()
Call Win_StayOnTop(cheatadvantage)
Call Win_CenterForm(cheatadvantage)
Timer1.Enabled = False

End Sub

Private Sub Label1_Click()
End                'click on X then exit program
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Picture2.Picture <> 0 Then Picture2.Picture = Picture6.Picture    'if the button pictures are visible
If Picture3.Picture <> 0 Then Picture3.Picture = Picture6.Picture    'then turn them off because mouse isnt over them

End Sub

Private Sub Picture2_Click()
Timer1.Enabled = True                        'start button
Timer2.Enabled = True                        'turn on random mouse and user mouse movement timers
Label4.Caption = "status: enabled"           'change status
ie& = FindWindow("ieframe", vbNullString)    'get explorer window
If ie& = 0 Then                              'if it is closed, then shell it to www.altavista.com as a maximized window
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE www.altavista.com", vbMaximizedFocus
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Picture2.Picture = 0 Then Picture2.Picture = Picture4.Picture  'the mouse is over the start button, so change the picture for the user
End Sub

Private Sub Picture3_Click()
Timer1.Enabled = False                 'turn off the random web page timer (mouse movement will have already been turned off)
Label4.Caption = "status: disabled"    'change status
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Picture3.Picture = 0 Then Picture3.Picture = Picture5.Picture  'mouse is over the stop button, so change the picture for the user

End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Call FormDrag(cheatadvantage)          'if the user wants to move the window, then let them
End Sub

Private Sub Timer1_Timer()
Randomize                                                                   'used to make rnd not the same
ie& = FindWindow("ieframe", vbNullString)                                   'find internet explorer
If ie& = 0 Then                                                             'if explorer is not loaded
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE", vbMaximizedFocus   'then shell it as a maximized
End If                                                                      'window
Do
DoEvents
ie& = FindWindow("ieframe", vbNullString)                                   'wait untill it is visible
Loop Until ie& <> 0
ie& = FindWindow("ieframe", vbNullString)                                   'find explorer window agin
worker& = FindChildByClass(ie&, "workera")                                  'find the sub class for the bar across the top
rebar& = FindChildByClass(worker&, "rebarwindow32")                         'find the bar across the top
combo32& = FindChildByClass(rebar&, "comboboxex32")                         'find the sub class for the combo box
combo& = FindChildByClass(combo32&, "combobox")                             'find the combo box
edit& = FindChildByClass(combo&, "edit")                                    'find where to put the adress in
X = Int(Rnd * 16) + 1
If X = 16 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.altavista.com")      'this just randomly picks a web
If X = 15 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.hotbot.com")         'page and sets it in the adress box
If X = 14 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.askjeeves.com")
If X = 13 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.winfiles.com")
If X = 12 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.skateboarding.com")
If X = 11 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.yahoo.com")
If X = 10 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.espn.com")
If X = 9 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.planetsourcecode.com")
If X = 8 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.windows.com")
If X = 7 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.sports.com")
If X = 6 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.asdf.com")
If X = 5 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.microsoft.com")
If X = 4 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.wwwheels.com")
If X = 3 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.basketball.com")
If X = 2 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.football.com")
If X = 1 Then giveweb& = SendMessageByString(edit&, WM_SETTEXT, 0, "www.autobytel.com")
gobutton& = FindChildByClass(combo32, "toolbarwindow32")                    'find the go button
Click& = SendMessage(gobutton&, WM_LBUTTONDOWN, 0, 0&)                      'click the go button
Click& = SendMessage(gobutton&, WM_LBUTTONUP, 0, 0&)
End Sub

Private Sub Timer2_Timer()
Randomize                                                             'used to make rnd different
wdth = Screen.Width / Screen.TwipsPerPixelX                           'find the users resolution
hight = Screen.Height / Screen.TwipsPerPixelY
Label2.Caption = Int(Rnd * wdth) + 1                                  'randomly get a point on the screen
Label3.Caption = Int(Rnd * hight) + 1                                 'to put the mouse (label2,label3)
Call SetMousePos(Label2.Caption, Label3.Caption)                      'set the pointer to randomb point
If Timer3.Enabled = False Then Timer3.Enabled = True                  'turn on timer3 wich checks for mouse movement
ie& = FindWindow("ieframe", vbNullString)                             'find explorer window
statusbar = FindChildByClass(ie&, "msctls_statusbar32")               'find the stats bar at the bottom of the explorer window
essages = Left(Get_Text(statusbar), 5)                                'get the first 5 letters of the status bar
If essages = "http:" Then                                             'if the the stats is a web page
Call LeftClick                                                        'then click on it
End If
End Sub

Private Sub Timer3_Timer()
xxx = GetX()                                                          'get the current x and y coords
yyy = GetY()                                                          'of the pointer
If Label2.Caption <> xxx Or Label3.Caption <> yyy Then                'if they are not the same as what is in (label2,label3) then
Timer2.Enabled = False                                                'the user moved the mouse
Timer3.Enabled = False                                                'so turn off the random mouse pointer timer
Label4.Caption = "status: enabled, but not moving mouse"              'change stats label
End If
End Sub
