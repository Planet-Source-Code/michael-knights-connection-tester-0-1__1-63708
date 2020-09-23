Attribute VB_Name = "Settings"
'Copyright Michael Knights - 2005

Public Sub SaveSettings()
'save the settings to the registry
On Error GoTo ERR_L
SaveSetting App.EXEName, "Settings", "ConnectionType", ConnectionType
SaveSetting App.EXEName, "Settings", "ConnectionHost", ConnectionHost
SaveSetting App.EXEName, "Settings", "ConnectionPort", ConnectionPort
SaveSetting App.EXEName, "Settings", "SendData", frmmain.txtSend
Exit Sub
ERR_L:
MsgBox "Unable To Save Settings.", vbInformation, App.EXEName & " Error"
End Sub

Public Sub LoadSettings()
'load the settings from the registry
On Error GoTo ERR_S
Dim Temp1 As String
ConnectionType = GetSetting(App.EXEName, "Settings", "ConnectionType", "TCP")
ConnectionHost = GetSetting(App.EXEName, "Settings", "ConnectionHost", "0.0.0.0")
ConnectionPort = GetSetting(App.EXEName, "Settings", "ConnectionPort", "2000")
Temp1 = GetSetting(App.EXEName, "Settings", "SendData", "")
frmmain.txtPort = ConnectionPort
frmmain.txtHost = ConnectionHost
frmmain.txtSend = Temp1
If ConnectionType = "TCP" Then
frmmain.cmdTCP.BackColor = &H8000000A
frmmain.cmdUPD.BackColor = &H8000000F
ElseIf ConnectionType = "UPD" Then
frmmain.cmdUPD.BackColor = &H8000000A
frmmain.cmdTCP.BackColor = &H8000000F
End If
Temp1 = "" ' Clear Variable To Save Memory
Exit Sub
ERR_S:
MsgBox "Unable To Load Settings.", vbInformation, App.EXEName & " Error"
End Sub



' ::: www.mksoftware.co.uk :::
