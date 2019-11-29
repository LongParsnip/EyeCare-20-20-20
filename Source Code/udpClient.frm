VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form udpClient 
   Caption         =   "udpClient"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton udpClientSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin MSWinsockLib.Winsock udpClient 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "udpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Static strMsg As String
    strMsg = "Enter the IP address of the server"
    udpClient.RemoteHost = frmMain.udpIP.Text
    If udpClient.RemoteHost = "" Then
         udpClient.RemoteHost = "localhost"
    End If
    udpClient.RemotePort = frmMain.udpPort.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
     udpClient.Close
End Sub

Private Sub udpClient_DataArrival(ByVal bytesTotal As Long)
Static strMsg As String
Static arrMsg As Variant
    
    udpClient.GetData strMsg
    arrMsg = Split(strMsg, ",")
    
    numIdleSeconds = Int(arrMsg(0))
    numActivityTimer = Int(arrMsg(1))
    numLookAwayTime = Int(arrMsg(2))
    
End Sub

Public Sub udpClientSend_Click()
On Error Resume Next
    udpClient.SendData Environ("COMPUTERNAME") & "," & Format(Now(), "HH:MM:SS") & "," & bMovement
End Sub
