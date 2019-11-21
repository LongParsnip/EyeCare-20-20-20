VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Label Label4 
         Caption         =   "<- Click to copy"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblBTC 
         Caption         =   "1MNCT8yBn8uH9z8BeWHRsXQhwAQREGDt4y"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "If you like this app and want to support me, please consider donating to my bitcoin address.:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3480
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   $"frmAbout.frx":000C
         Height          =   855
         Left            =   2280
         TabIndex        =   4
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   $"frmAbout.frx":00B5
         Height          =   1095
         Left            =   2280
         TabIndex        =   3
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmAbout.frx":01BC
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCompany 
         Caption         =   "By: Paul Atkins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Eyecare 20-20-20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   2
         Top             =   225
         Width           =   2865
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblBTC_Click()
    Clipboard.Clear
    Clipboard.SetText lblBTC.Caption
    Unload Me
End Sub
