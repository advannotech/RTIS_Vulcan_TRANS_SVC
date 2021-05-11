Imports System.Data.SqlClient

Public Class WarehouseTransfers

    Public Shared RTString As String = "Data Source=" + My.Settings.RTServer + "; Initial Catalog=" + My.Settings.RTDB +
    "; user ID=" + My.Settings.RTUser + "; password=" + My.Settings.RTPassword + ";Max Pool Size=99999;"

    Public Shared Sub proccessWhseTransfers()
        Try
            Dim WhseTransferLines As String = WhseTransfers.RTSQL.Retreive.SVC_GetPendingWhseTransfers()
            Select Case WhseTransferLines.Split("*")(0)
                Case "1"
                    WhseTransferLines = WhseTransferLines.Remove(0, 2)
                    Dim allWhseTransfers As String() = WhseTransferLines.Split("~")
                    For Each whseTransfer As String In allWhseTransfers
                        If whseTransfer <> String.Empty Then

                            Dim lineID As String = whseTransfer.Split("|")(0)
                            Dim code As String = whseTransfer.Split("|")(1)
                            Dim lotNumber As String = whseTransfer.Split("|")(2)
                            Dim whseFrom As String = whseTransfer.Split("|")(3)
                            Dim whseTo As String = whseTransfer.Split("|")(4)
                            Dim qty As String = whseTransfer.Split("|")(5)
                            Dim username As String = whseTransfer.Split("|")(6)
                            Dim proccess As String = whseTransfer.Split("|")(7)
                            Dim transDesc As String = whseTransfer.Split("|")(8)
                            Dim transDate As String = whseTransfer.Split("|")(9)

                            Dim transferred As String = WhseTransfers.EvolutionSDK.CTransferItem("RTIS Auto transfer", whseFrom, whseTo, code, lotNumber, qty)
                            Select Case transferred.Split("*")(0)
                                Case "1"
                                    Dim completedInserted As String = WhseTransfers.RTSQL.Insert.SVC_InsertWHTLineCompleted(lineID, code, lotNumber, whseFrom, whseTo, qty, transDate, username, proccess, transDesc)
                                    Select Case completedInserted.Split("*")(0)
                                        Case "1"
                                            Dim lineRemoved As String = WhseTransfers.RTSQL.Delete.SVC_DeleteWhseTransLineComplete(lineID)
                                            Select Case lineRemoved.Split("*")(0)
                                                Case "1"

                                                Case "-1"
                                                    lineRemoved = lineRemoved.Remove(0, 3)
                                                    EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + lineRemoved)
                                            End Select
                                        Case "-1"
                                            completedInserted = completedInserted.Remove(0, 3)
                                            EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + completedInserted)
                                    End Select
                                Case "-1"
                                    Dim failureReason As String = transferred.Split("*")(1)
                                    Dim failureUpdated As String = WhseTransfers.RTSQL.Update.SVC_UpdateWhseTransferFailed(lineID, failureReason)
                                    Select Case failureUpdated.Split("*")(0)
                                        Case "1"
                                            'Line set as failed
                                        Case "-1"
                                            failureUpdated = failureUpdated.Remove(0, 3)
                                            EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + failureUpdated)
                                    End Select
                            End Select
                        End If
                    Next
                Case "0"
                    'No Lines Found
                Case "-1"
                    WhseTransferLines = WhseTransferLines.Remove(0, 3)
                    EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + WhseTransferLines)
                Case Else

            End Select
        Catch ex As Exception
            EventLog.WriteEntry("RTIS Vulcan SVC", "proccessWhseTransfers: " + Environment.NewLine + ex.ToString())
        End Try
    End Sub

    Public Shared Function GetWHTriggered() As String
        Try
            Dim ReturnData As String = ""
            Dim sqlConn As New SqlConnection(RTString)
            Dim sqlComm As New SqlCommand("  SELECT [SettingValue] FROM [tbl_RTSettings] WHERE [Setting_Name] = 'WHTimer'", sqlConn)
            sqlConn.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()
            While sqlReader.Read()
                ReturnData = Convert.ToString(sqlReader.Item(0))
            End While
            sqlReader.Close()
            sqlComm.Dispose()
            sqlConn.Close()

            If ReturnData <> "" Then
                Return "1*" + ReturnData
            Else
                Return "-1*Missing setting: WHTimer"
            End If
        Catch ex As Exception
            If ex.Message = "Invalid attempt to read when no data is present." Then
                Return "-1*Missing setting: WHTimer"
            Else
                EventLog.WriteEntry("RTIS Vulcan SVC", "GetWHTriggered: " + ex.ToString())
                Return ExHandler.returnErrorEx(ex)
            End If
        End Try
    End Function

    Public Shared Function SetWHTriggeredTrue() As String
        Try
            Dim ReturnData As String = ""
            Dim sqlConn As New SqlConnection(RTString)
            Dim sqlComm As New SqlCommand("UPDATE [tbl_RTSettings] SET [SettingValue] = 'True' WHERE [Setting_Name] = 'WHTimer'", sqlConn)
            sqlConn.Open()
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()
            sqlConn.Close()
            Return "1*"
        Catch ex As Exception
            EventLog.WriteEntry("RTIS Vulcan SVC", "SetWHTriggered: " + ex.ToString())
            Return ExHandler.returnErrorEx(ex)
        End Try
    End Function

    Public Shared Function SetWHTriggeredFalse() As String
        Try
            Dim ReturnData As String = ""
            Dim sqlConn As New SqlConnection(RTString)
            Dim sqlComm As New SqlCommand("UPDATE [tbl_RTSettings] SET [SettingValue] = 'False' WHERE [Setting_Name] = 'WHTimer'", sqlConn)
            sqlConn.Open()
            sqlComm.ExecuteNonQuery()
            sqlComm.Dispose()
            sqlConn.Close()
            Return "1*"
        Catch ex As Exception
            EventLog.WriteEntry("RTIS Vulcan SVC", "SetWHTriggered: " + ex.ToString())
            Return ExHandler.returnErrorEx(ex)
        End Try
    End Function
End Class
