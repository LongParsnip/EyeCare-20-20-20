VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form udpServer 
   Caption         =   "udpServer"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton udpServerSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin MSWinsockLib.Winsock udpServer 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "udpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Note the server works by replying to clients after they send data, it does not routinely send updates.

Private Sub Form_Load()
     udpServer.Bind frmMain.udpPort.Text
     frmMain.lblLocalIP.Text = findLocalIP
     frmMain.lblIPAddress.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
     udpServer.Close
End Sub

Private Sub udpServer_DataArrival(ByVal bytesTotal As Long)

Static strMsg As String
Static arrMsg As Variant
Static arrPos As Integer
Static clientLength As Integer

    
    clientLength = UBound(arrClients)
    
    udpServer.GetData strMsg
    arrMsg = Split(strMsg, ",")
    arrPos = findArrPos(arrMsg(0))
    
    'No array entry found.
    If arrPos = -1 Then
        MsgBox "Too many clients connected!", "vbOkOnly", "Network Error"
        Unload Me
        Unload frmMain
    End If
    
    'Write data to array.
    With arrClients(arrPos)
        .ComputerName = arrMsg(0)
        .UpdateTime = Format(Now(), "HH:MM:SS")     'arrMsg(1).... holy fkn herp derp.
        If arrMsg(2) = "True" Then .Movement = True
        If arrMsg(2) = "False" Then .Movement = False
    End With
    
    'Reply To Client
    Call udpServerSend_Click

End Sub

Public Sub udpServerSend_Click()
On Error Resume Next
    udpServer.SendData numIdleSeconds & "," & numActivityTimer & "," & numLookAwayTime
End Sub

Private Function findArrPos(strComputerName)

Dim i
    
    For i = 0 To UBound(arrClients)
        If arrClients(i).ComputerName = strComputerName Then Exit For
    Next i
    
    'Nothing found, find next empty slot.
    If i = UBound(arrClients) + 1 Then
        For i = 0 To UBound(arrClients)
            If arrClients(i).ComputerName = "" Then Exit For
        Next i
    End If
    
    If i = UBound(arrClients) + 1 Then
        findArrPos = -1
    Else
        findArrPos = i
    End If

End Function
