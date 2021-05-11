Imports System.Timers
Public Class Service
    Dim tmrTransfer As New System.Timers.Timer()
    Dim triggerd As Boolean = False
    Protected Overrides Sub OnStart(ByVal args() As String)
        tmrTransfer.Interval = 10000
        AddHandler tmrTransfer.Elapsed, AddressOf tmrTransfer_Tick
        tmrTransfer.Start()
    End Sub
    Private Sub tmrTransfer_Tick(sender As Object, e As ElapsedEventArgs)
        tmrTransfer.Stop()
        Dim rand As New Random()
        Dim lower As Integer = 180000
        Dim upper As Integer = 360000
        Dim intRandomNumber = rand.Next(lower, upper)
        tmrTransfer.Interval = intRandomNumber
        proccessWhseTransfers()
    End Sub

    Public Sub proccessWhseTransfers()
        Try
            tmrTransfer.Stop()
            WarehouseTransfers.proccessWhseTransfers()
            tmrTransfer.Start()
        Catch ex As Exception
            EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + ex.ToString())
            triggerd = False
            tmrTransfer.Start()
        End Try
    End Sub

    'Public Sub proccessWhseTransfers()
    '    Try
    '        triggerd = True
    '        Dim timerTriggered = WarehouseTransfers.GetWHTriggered()
    '        Select Case timerTriggered.Split("*")(0)
    '            Case "1"
    '                timerTriggered = timerTriggered.Remove(0, 2)
    '                If Convert.ToBoolean(timerTriggered) = False Then
    '                    Dim setToTrue As String = WarehouseTransfers.SetWHTriggeredTrue()
    '                    Select Case setToTrue.Split("*")(0)
    '                        Case "1"
    '                            WarehouseTransfers.proccessWhseTransfers()
    '                            ' WarehouseTransfers.SetWHTriggeredFalse()
    '                            triggerd = False
    '                            tmrTransfer.Start()
    '                        Case "-1"
    '                            EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + setToTrue.Remove(0, 2))
    '                            triggerd = False
    '                            tmrTransfer.Start()
    '                        Case Else
    '                            EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + setToTrue)
    '                            triggerd = False
    '                            tmrTransfer.Start()
    '                    End Select
    '                    WarehouseTransfers.SetWHTriggeredFalse()
    '                Else
    '                    triggerd = False
    '                    tmrTransfer.Start()
    '                End If
    '            Case "-1"
    '                EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + timerTriggered.Remove(0, 2))
    '                triggerd = False
    '                tmrTransfer.Start()
    '            Case Else
    '                EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + timerTriggered)
    '                triggerd = False
    '                tmrTransfer.Start()
    '        End Select
    '    Catch ex As Exception
    '        EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + ex.ToString())
    '        triggerd = False
    '        tmrTransfer.Start()
    '    End Try
    'End Sub
    Protected Overrides Sub OnStop()

    End Sub
End Class
