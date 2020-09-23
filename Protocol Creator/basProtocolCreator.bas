Attribute VB_Name = "basProtocolCreator"
'  Title: ProtocolCreator
'Version: 1.0
' Author: Josh McCullough
'Contact: FistfulOfSteel on AIM
'   Info: Put this module in any of your programs that require a protocol to be assigned
'to them.  Simply pass the following to the 'CreateProtocol' function:
'   Protocol - what you want the protocol to be
'   Description - a description of your protocol
'   Path - the path that the protocol should execute
'   Run - whether the REG file should be run after creation
'   Delete - whether the REG file should be deleted after execution
'   AllowAll - whether the most common protocols should be protected
'
'Example: You want the protocol 'myApp', description 'Launches myApp', you want to run
'it after it is created, you do not want to delete it after it is created, and you want
'to disallow any protected names - you would type this:
'   intResult=CreateProtocol("myApp","Launches myApp",True,False,False)
'
'You can catch errors (as the line above does) and get that error by typing:
'   strError=GetProtocolCreationError(intResult)
'
'Put it all together now:
'   strError=GetProtocolCreationError(CreateProtocol("myApp","Launches myApp",True,False,False))
'
'Get it? If not message me on AIM at: FistfulOfSteel82
'
'You can use this as much as you want, and distribute it where-ever, when-ever you want,
'all I ask is that you keep this comment intact.
'
'Please vote for me if you like this code!
'Please go to my website: http://jsment.com
'(c) 2002 Josh McCullough.

Option Explicit
Dim mintError As Integer
Dim strProtocols(23) As String
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub BuildArray()
    strProtocols(0) = ""
    strProtocols(1) = "cifs"
    strProtocols(2) = "dhcp"
    strProtocols(3) = "dns"
    strProtocols(4) = "ftp"
    strProtocols(5) = "http"
    strProtocols(6) = "icmp"
    strProtocols(7) = "ip"
    strProtocols(8) = "irc"
    strProtocols(9) = "kerberos"
    strProtocols(10) = "mail"
    strProtocols(11) = "nntp"
    strProtocols(12) = "ntp"
    strProtocols(13) = "pkix"
    strProtocols(14) = "ppp"
    strProtocols(15) = "radius"
    strProtocols(16) = "rtsp"
    strProtocols(17) = "rwhios"
    strProtocols(18) = "samba"
    strProtocols(19) = "snmp"
    strProtocols(20) = "ssl-tls"
    strProtocols(21) = "tcp"
    strProtocols(22) = "udp"
    strProtocols(23) = "webdav"
End Sub

Public Function CreateProtocol(strProtocol As String, strDescription As String, strPath As String, blnRun As Boolean, blnAllowAll As Boolean, frmSource As Form) As Integer
    Dim intCheck, intShell As Integer
    Dim oFile As Integer
    
    If blnAllowAll = False Then
        intCheck = CheckProtocol(LCase(strProtocol))
        If intCheck <> 0 Then
            CreateProtocol = intCheck
            Exit Function
        End If
    End If
    
    oFile = FreeFile
    If Dir(App.Path & "\REG Files", vbDirectory) = "" Then MkDir (App.Path & "\REG Files")
    Open App.Path & "\REG Files\" & strProtocol & "_protocol.reg" For Output As #oFile
    Print #oFile, "Windows Registry Editor Version 5.00" & vbCrLf
    Print #oFile, "[HKEY_CLASSES_ROOT\" & strProtocol & "]"
    Print #oFile, "@=""" & strDescription & """"
    Print #oFile, """URL Protocol""=""""" & vbCrLf
    Print #oFile, "[HKEY_CLASSES_ROOT\" & strProtocol & "\shell]" & vbCrLf
    Print #oFile, "[HKEY_CLASSES_ROOT\" & strProtocol & "\shell\open]" & vbCrLf
    Print #oFile, "[HKEY_CLASSES_ROOT\" & strProtocol & "\shell\open\command]"
    Print #oFile, "@=""\""" & Replace(strPath, "\", "\\") & """"
    Close #oFile
    If blnRun Then ShellExecute frmSource.hwnd, "Open", App.Path & "\REG Files\" & strProtocol & "_protocol.reg", vbNullString, vbNullString, 3
    CreateProtocol = 0
End Function

Private Function CheckProtocol(strProtocol) As Integer
    Dim i As Integer
    
    BuildArray
    For i = 0 To UBound(strProtocols)
        If strProtocol = strProtocols(i) Then
            CheckProtocol = i
            Exit Function
        End If
    Next i
End Function

Public Function GetProtocolCreationError(intError As Integer) As String
    If intError = 0 Then
        GetProtocolCreationError = "Protocol was accepted.  REG file was successfully created."
    Else
        GetProtocolCreationError = "Can not replace " & strProtocols(intError) & " protocol, it is protected."
    End If
End Function
