Imports System.Security.Principal
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography.X509Certificates
Imports NetFwTypeLib
Imports System.IO


<ComVisible(True)>
Public Class Admin
    Public Shared FirewallSet As Boolean = False
    'Declare API  
    Private Declare Ansi Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Const BCM_FIRST As Int32 = &H1600
    Private Const BCM_SETSHIELD As Int32 = (BCM_FIRST + &HC)

    '<ComVisible(False)> _
    'Public Shared Function IsWin7orHigher() As Boolean
    '    If Environment.OSVersion.Version.Build < 6000 Then
    '        Return False
    '    Else
    '        Return True
    '    End If

    'End Function

    Public Shared Function IsWinVista() As Boolean
        If Environment.OSVersion.Version.Build >= 6000 AndAlso Environment.OSVersion.Version.Build < 7000 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function IsWin7() As Boolean
        If Environment.OSVersion.Version.Build >= 7000 AndAlso Environment.OSVersion.Version.Build < 8000 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function IsWin8orHigher() As Boolean
        If Environment.OSVersion.Version.Build >= 8000 Then
            Return True
        Else
            Return False
        End If

    End Function

    <ComVisible(True)>
    Public Shared Function IsVistaOrHigher() As Boolean
        If Environment.OSVersion.Version.Major < 6 Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Shared Function InstallADECert()
        Try
            Dim store As New X509Store(StoreName.Root, StoreLocation.LocalMachine)
            store.Open(OpenFlags.ReadWrite)
            store.Add(New X509Certificate2(X509Certificate2.CreateFromCertFile("ADE_RNDIS.cer")))
            store.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Shared Sub AddFirewallException(ByVal AppName As String, ByVal AppLoc As String)
        Try
            If IsVistaOrHigher() = False Then
                Dim appType As Type = Type.GetTypeFromProgID("HnetCfg.FwAuthorizedApplication")
                Dim app As INetFwAuthorizedApplication
                app = DirectCast(Activator.CreateInstance(appType), INetFwAuthorizedApplication)

                ' Set the application properties
                app.Name = AppName

                app.ProcessImageFileName = AppLoc
                app.Enabled = True

                ' Get the firewall manager, so we can get the list of authorized apps
                Dim fwMgrType As Type = Type.GetTypeFromProgID("HnetCfg.FwMgr")
                Dim fwMgr As INetFwMgr
                fwMgr = DirectCast(Activator.CreateInstance(fwMgrType), INetFwMgr)

                ' Get the list of authorized applications from the Firewall Manager, so we can add our app to that list
                Dim apps As INetFwAuthorizedApplications
                apps = fwMgr.LocalPolicy.CurrentProfile.AuthorizedApplications
                apps.Add(app)
                apps.Item(AppName).Scope = NET_FW_SCOPE_.NET_FW_SCOPE_ALL
                apps.Item(AppName).Enabled = True

                'Dim OpenPort As INetFwOpenPort = DirectCast(Activator.CreateInstance(appType), INetFwOpenPort)
                'OpenPort.Name = AppName
                'OpenPort.Enabled = True
                'OpenPort.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_ANY
                'OpenPort.Port = 25213
                'OpenPort.Scope = NET_FW_SCOPE_.NET_FW_SCOPE_ALL
                'Dim ports As INetFwOpenPorts = fwMgr.LocalPolicy.CurrentProfile.GloballyOpenPorts
                'ports.Add(OpenPort)
            Else
                Dim FWPolicy As INetFwPolicy2 = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWPolicy2")), INetFwPolicy2)

                Dim FWRule As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                FWRule.Name = AppName & " IN (DOMAIN)"
                FWRule.ApplicationName = AppLoc
                FWRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN
                FWRule.Description = "FIREWALL RULE 1 FOR " & AppName
                'FWRule.InterfaceTypes = "All"
                FWRule.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP
                'FWRule.LocalPorts = 25213
                FWRule.LocalAddresses = "*"
                FWRule.Enabled = True
                FWRule.Grouping = "@firewallapi.dll,-23255"
                FWRule.Profiles = 1
                FWRule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                Try
                    Dim temp As INetFwRule

                    For Each i As INetFwRule In FWPolicy.Rules
                        If i.Name = AppName & " IN (DOMAIN)" Then
                            temp = i
                            Exit For
                        End If
                    Next
                    If temp IsNot Nothing Then
                        FWPolicy.Rules.Remove(temp.Name)
                    End If

                    FWPolicy.Rules.Add(FWRule)
                Catch
                    FWPolicy.Rules.Add(FWRule)
                End Try

                Dim FWRule2 As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                FWRule2.Name = AppName & " IN (PRIVATE)"
                FWRule2.ApplicationName = AppLoc
                FWRule2.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN
                FWRule2.Description = "FIREWALL RULE 2 FOR " & AppName
                'FWRule.InterfaceTypes = "All"
                FWRule2.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP
                'FWRule2.LocalPorts = 25213
                FWRule2.LocalAddresses = "*"
                FWRule2.Enabled = True
                FWRule2.Grouping = "@firewallapi.dll,-23255"
                FWRule2.Profiles = 2
                FWRule2.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                Try
                    Dim temp As INetFwRule

                    For Each i As INetFwRule In FWPolicy.Rules
                        If i.Name = AppName & " IN (PRIVATE)" Then
                            temp = i
                            Exit For
                        End If
                    Next
                    If temp IsNot Nothing Then
                        FWPolicy.Rules.Remove(temp.Name)
                    End If

                    FWPolicy.Rules.Add(FWRule2)
                Catch
                    FWPolicy.Rules.Add(FWRule2)
                End Try

                Dim FWRule3 As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                FWRule3.Name = AppName & " IN (PUBLIC)"
                FWRule3.ApplicationName = AppLoc
                FWRule3.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN
                FWRule3.Description = "FIREWALL RULE 3 FOR " & AppName
                'FWRule.InterfaceTypes = "All"
                FWRule3.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP
                'FWRule3.LocalPorts = 25213
                FWRule3.LocalAddresses = "*"
                FWRule3.Enabled = True
                FWRule3.Grouping = "@firewallapi.dll,-23255"
                FWRule3.Profiles = 4
                FWRule3.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                Try
                    Dim temp As INetFwRule

                    For Each i As INetFwRule In FWPolicy.Rules
                        If i.Name = AppName & " IN (PUBLIC)" Then
                            temp = i
                            Exit For
                        End If
                    Next
                    If temp IsNot Nothing Then
                        FWPolicy.Rules.Remove(temp.Name)
                    End If

                    FWPolicy.Rules.Add(FWRule3)
                Catch
                    FWPolicy.Rules.Add(FWRule3)
                End Try

                Dim FWRule4 As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                FWRule4.Name = AppName & " OUT (DOMAIN)"
                FWRule4.ApplicationName = AppLoc
                FWRule4.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT
                FWRule4.Description = "FIREWALL RULE 4 FOR " & AppName
                'FWRule2.InterfaceTypes = "All"
                FWRule4.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP
                'FWRule4.LocalPorts = 25213
                FWRule4.LocalAddresses = "*"
                FWRule4.Enabled = True
                FWRule4.Grouping = "@firewallapi.dll,-23255"
                FWRule4.Profiles = 1
                FWRule4.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                Try
                    Dim temp As INetFwRule

                    For Each i As INetFwRule In FWPolicy.Rules
                        If i.Name = AppName & " OUT (DOMAIN)" Then
                            temp = i
                            Exit For
                        End If
                    Next
                    If temp IsNot Nothing Then
                        FWPolicy.Rules.Remove(temp.Name)
                    End If

                    FWPolicy.Rules.Add(FWRule4)
                Catch
                    FWPolicy.Rules.Add(FWRule4)
                End Try

                Dim FWRule5 As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                FWRule5.Name = AppName & " OUT (PRIVATE)"
                FWRule5.ApplicationName = AppLoc
                FWRule5.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT
                FWRule5.Description = "FIREWALL RULE 5 FOR " & AppName
                'FWRule2.InterfaceTypes = "All"
                FWRule5.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP
                'FWRule5.LocalPorts = 25213
                FWRule5.LocalAddresses = "*"
                FWRule5.Enabled = True
                FWRule5.Grouping = "@firewallapi.dll,-23255"
                FWRule5.Profiles = 2
                FWRule5.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                Try
                    Dim temp As INetFwRule

                    For Each i As INetFwRule In FWPolicy.Rules
                        If i.Name = AppName & " OUT (PRIVATE)" Then
                            temp = i
                            Exit For
                        End If
                    Next
                    If temp IsNot Nothing Then
                        FWPolicy.Rules.Remove(temp.Name)
                    End If

                    FWPolicy.Rules.Add(FWRule5)
                Catch
                    FWPolicy.Rules.Add(FWRule5)
                End Try

                Dim FWRule6 As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                FWRule6.Name = AppName & " OUT (PUBLIC)"
                FWRule6.ApplicationName = AppLoc
                FWRule6.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT
                FWRule6.Description = "FIREWALL RULE 6 FOR " & AppName
                'FWRule2.InterfaceTypes = "All"
                FWRule6.Protocol = NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP
                'FWRule6.LocalPorts = 25213
                FWRule6.LocalAddresses = "*"
                FWRule6.Enabled = True
                FWRule6.Grouping = "@firewallapi.dll,-23255"
                FWRule6.Profiles = 4
                FWRule6.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                Try
                    Dim temp As INetFwRule

                    For Each i As INetFwRule In FWPolicy.Rules
                        If i.Name = AppName & " OUT (PUBLIC)" Then
                            temp = i
                            Exit For
                        End If
                    Next
                    If temp IsNot Nothing Then
                        FWPolicy.Rules.Remove(temp.Name)
                    End If

                    FWPolicy.Rules.Add(FWRule6)
                Catch
                    FWPolicy.Rules.Add(FWRule6)
                End Try
                'Dim FWRule3 As INetFwRule = DirectCast(Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule")), INetFwRule)
                'FWRule3.Name = "ICMP RULE"
                'FWRule3.Description = "Allow ICMP network traffic"
                ''FWRule2.InterfaceTypes = "All"
                'FWRule3.Protocol = 1
                'FWRule3.IcmpTypesAndCodes = "1:1"
                'FWRule3.Enabled = True
                'FWRule3.Grouping = "@firewallapi.dll,-23255"
                'FWRule3.Profiles = FWPolicy.CurrentProfileTypes
                'FWRule3.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW
                'FWPolicy.Rules.Add(FWRule3)
                For Each r As INetFwRule In FWPolicy.Rules
                    If UCase(r.Name).Contains("ICMP") Then
                        r.Enabled = True
                    End If
                Next
            End If




        Catch ex As Exception
            EventLog.WriteEntry("RTIS Vulcan SVC", ex.ToString)
        End Try

    End Sub

    ' Checks if the process is elevated  

    <ComVisible(True)>
    Public Shared Function IsAdmin() As Boolean
        Dim id As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim p As WindowsPrincipal = New WindowsPrincipal(id)
        Return p.IsInRole(WindowsBuiltInRole.Administrator)
    End Function

    ' Add a shield icon to a button  

    'Public Sub AddShieldToButton(ByRef b As Button)
    '    b.FlatStyle = FlatStyle.System
    '    SendMessage(b.Handle, BCM_SETSHIELD, 0, &HFFFFFFFF)
    'End Sub


    '' Restart the current process with administrator credentials  
    <ComVisible(True)>
    Public Shared Function RestartElevated(ByVal ApplicationPath As String) As Integer
        Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        startInfo.UseShellExecute = True
        startInfo.WorkingDirectory = Environment.CurrentDirectory
        startInfo.FileName = ApplicationPath
        startInfo.Verb = "runas"
        Try
            Dim p As Process = Process.Start(startInfo)
        Catch ex As Exception
            Return 0 'If cancelled, do nothing  
        End Try
        Return 1
    End Function



End Class

