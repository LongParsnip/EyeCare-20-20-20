Attribute VB_Name = "A_Common"
Public numIdleSeconds As Long       'Global variable for tracking user idle time.
Public numActivityTimer As Long     'Global variable for tracking user activity time.
Public numLookAwayTime As Long      'Global variable for tracking user look away time, used when 20 minutes is up and screen is flashing.
Public bMovement As Boolean         'Any movement sets this flag to true, used for network mode.
Public bNetworkMode As Boolean      'True when network mode enabled.
Public numNetworkMode As Integer    '0 = Client, 1 = Server
Public arrClients(9) As clientData  'An array of clients and their data, max 10 connections, added by computer names.... disconnects are not removed.

Public Type clientData
    ComputerName As String  'Environ("COMPUTERNAME")
    UpdateTime As String    'Format(Now(),"HH:MM:SS")
    Movement As Boolean     'bMovement
End Type


'Evaluates whether or not clients are reporting movement.
Public Function evalClientData() As Boolean

On Error Resume Next

Dim i
Dim bMovement
    
    For i = 0 To UBound(arrClients)
        'Check update was within the last 10 seconds.
        If DateDiff("s", Format(Now(), "HH:MM:SS"), arrClients(i).UpdateTime) < 10 Then
            If arrClients(i).Movement = "True" Then
                bMovement = True
            End If
        End If
    Next i
    
    evalClientData = bMovement

End Function


'Gets a list of all the potential Local IP addresses.
'Winsock.LocalIP only returns the first one, which is no good as there may be more than one available.
Public Function findLocalIP()

Dim objWMI     As Object: Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Dim objWMIqry  As Object: Set objWMIqry = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
Dim Item    As Variant
Dim strTemp As String

    For Each Item In objWMIqry
        strTemp = strTemp & Item.IPAddress(0) & vbCrLf
    Next
    
    findLocalIP = strTemp

End Function


