VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eyecare 20-20-20"
   ClientHeight    =   3405
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   240
   End
   Begin VB.Label lblLookAwayTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblLookAway 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Don't touch the computer and look away for 20 seconds."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblCounter 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Shape shpFlash 
      FillColor       =   &H000000FF&
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI: X As Long: Y As Long: End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type MousePosition: Current As POINTAPI: Last As POINTAPI: End Type
Private Enum WindowState: Normal = 0: Minimized = 1: Maximized = 2: End Enum
Private Enum Fillstyle: Solid = 0: Transparent = 1: End Enum

Private Const IDLE_TIME As Integer = 30
Private Const LEAVE_TIME As Integer = 300
Private Const FULL_TIME As Integer = 1200




'1 Second timer.
Private Sub Timer1_Timer()

    Static numIdleSeconds As Long
    Static CursorPos_Last As POINTAPI
    Static numActivityTimer
    Static numTick
    Static numLookAwayTime
    Static MousePos As MousePosition

    Dim i As Integer
    

    numIdleSeconds = numIdleSeconds + 1
    If numIdleSeconds < IDLE_TIME Then
        numActivityTimer = numActivityTimer + 1
        If numActivityTimer > FULL_TIME Then numActivityTimer = FULL_TIME     'Cap at 20 minutes.
    End If
    
    
    'Update idle timer.
    '--------------------------------------------------------------------------------------
    Call GetCursorPos(MousePos.Current)
    If (MousePos.Current.X <> MousePos.Last.X) Or (MousePos.Current.Y <> MousePos.Last.Y) Then
        MousePos.Last = MousePos.Current
        numIdleSeconds = 0
    Else
        For i = 0 To 255
            If GetAsyncKeyState(i) <> 0 Then
                numIdleSeconds = 0
            End If
        Next
    End If
    
    
    '20 Minutes is up.
    'Flash Background, set app to front.
    '-----------------------------------------------------------------
    If numActivityTimer >= FULL_TIME Then
    
        Call WindownOnTop(Me.hWnd)
        Me.WindowState = WindowState.Normal
        lblLookAway.Visible = True: lblLookAwayTime.Visible = True
        lblLookAwayTime.Caption = SecondsToTime(numLookAwayTime)
        If numTick Mod 2 = 0 Then
            shpFlash.Fillstyle = Fillstyle.Transparent
        Else
            shpFlash.Fillstyle = Fillstyle.Solid
        End If
        
        If numIdleSeconds > 0 Then
            numLookAwayTime = numLookAwayTime + 1
        Else
            numLookAwayTime = 0
        End If
        
        'Looked away long enough, reset timer.
        If numLookAwayTime >= 22 Then
            numLookAwayTime = 0
            lblLookAway.Visible = False: lblLookAwayTime.Visible = False
            lblLookAwayTime.Caption = "00:00"
            numActivityTimer = Fillstyle.Solid
            shpFlash.Fillstyle = Fillstyle.Transparent
            Me.WindowState = WindowState.Minimized
            Call WindownOnBottom(Me.hWnd)
        End If
        
    End If
    '-----------------------------------------------------------------
    
    
    'Idle Time Exceeded - assuming user has left the office, reset the activity Timer.
    If numIdleSeconds > LEAVE_TIME Then
        numActivityTimer = 0
    End If

    
    'Update tick timer, and update screen.
    numTick = numTick + 1:  If numTick > 100 Then numTick = 0
    Label1.Caption = numIdleSeconds
    lblCounter.Caption = SecondsToTime(numActivityTimer)
    
End Sub


'Converts seconds to a string.
Private Function SecondsToTime(inputSeconds)
    
    Dim strMinutes As String
    Dim strSeconds As String
    
    strMinutes = Int(inputSeconds / 60)
    If Len(strMinutes) = 1 Then strMinutes = "0" & strMinutes
    strSeconds = Round(((inputSeconds / 60) - Int(inputSeconds / 60)) * 60, 0)
    If Len(strSeconds) = 1 Then strSeconds = "0" & strSeconds
    SecondsToTime = strMinutes & ":" & strSeconds

End Function


Private Sub About_Click()
    frmAbout.Left = Me.Left: frmAbout.Top = Me.Top
    frmAbout.Show vbModal
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure that you want to close Eyecare 20-20-20?", vbYesNo, "Close Application?") = vbYes Then
        Exit Sub
    Else
        Cancel = True
    End If
End Sub
