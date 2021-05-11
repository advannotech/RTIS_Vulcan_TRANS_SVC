Imports System.Data.SqlClient
Imports System.Threading
Imports Pastel.Evolution

Public Class WhseTransfers

    Public Shared RTString As String = "Data Source=" + My.Settings.RTServer + "; Initial Catalog=" + My.Settings.RTDB +
    "; user ID=" + My.Settings.RTUser + "; password=" + My.Settings.RTPassword + ";Max Pool Size=99999;"
    Public Shared EvoString As String = "Data Source=" + My.Settings.EvoServer + "; Initial Catalog=" + My.Settings.EvoDB +
     "; user ID=" + My.Settings.EvoUser + "; password=" + My.Settings.EvoPassword + ";Max Pool Size=99999;"
    Public Shared sep As String = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

    Public Class RTSQL
        Partial Public Class Retreive
            Public Shared Function SVC_GetPendingWhseTransfers() As String
                Try
                    Dim ReturnData As String = ""
                    Dim sqlConn As New SqlConnection(RTString)
                    Dim sqlComm As New SqlCommand("  SELECT TOP 1000 [iLineID], [vItemCode], [vLotNumber], [vWarehouse_From], [vWarehouse_To], [dQtyTransfered], [vUsername] ,[vProcess] ,[vTransDesc], [dtDateTransfered] 
                                                     FROM [tbl_WHTPending] WHERE [vStatus] = 'Pending'", sqlConn)
                    sqlConn.Open()
                    Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()
                    sqlReader.Read()
                    ReturnData = Convert.ToString(sqlReader.Item(0)) + "|" + Convert.ToString(sqlReader.Item(1)) + "|" + Convert.ToString(sqlReader.Item(2)) + "|" + Convert.ToString(sqlReader.Item(3)) + "|" + Convert.ToString(sqlReader.Item(4)) + "|" + Convert.ToString(sqlReader.Item(5)) + "|" + Convert.ToString(sqlReader.Item(6)) + "|" + Convert.ToString(sqlReader.Item(7)) + "|" + Convert.ToString(sqlReader.Item(8)) + "|" + Convert.ToString(sqlReader.Item(9)) + "~"
                    While sqlReader.Read()
                        ReturnData &= Convert.ToString(sqlReader.Item(0)) + "|" + Convert.ToString(sqlReader.Item(1)) + "|" + Convert.ToString(sqlReader.Item(2)) + "|" + Convert.ToString(sqlReader.Item(3)) + "|" + Convert.ToString(sqlReader.Item(4)) + "|" + Convert.ToString(sqlReader.Item(5)) + "|" + Convert.ToString(sqlReader.Item(6)) + "|" + Convert.ToString(sqlReader.Item(7)) + "|" + Convert.ToString(sqlReader.Item(8)) + "|" + Convert.ToString(sqlReader.Item(9)) + "~"
                    End While
                    sqlReader.Close()
                    sqlComm.Dispose()
                    sqlConn.Close()

                    If ReturnData <> "" Then
                        Return "1*" + ReturnData
                    Else
                        Return "0*Pending Lines!"
                    End If
                Catch ex As Exception
                    If ex.Message = "Invalid attempt to read when no data is present." Then
                        Return "0*Pending Lines!"
                    Else
                        EventLog.WriteEntry("RTIS Vulcan SVC", "SVC_GetPendingWhseTransfers: " + ex.ToString())
                        Return ExHandler.returnErrorEx(ex)
                    End If
                End Try
            End Function
        End Class

        Partial Public Class Update
            Public Shared Function SVC_UpdateWhseTransferFailed(ByVal id As String, ByVal reason As String) As String
                Try
                    Dim ReturnData As String = ""
                    Dim sqlConn As New SqlConnection(RTString)
                    Dim sqlComm As New SqlCommand("  UPDATE [tbl_WHTPending] SET [vStatus] = 'Failed', [vFailureReason] = @2, [dtDateFailed] = GETDATE()
                                                     WHERE [iLineID] = @1", sqlConn)
                    sqlComm.Parameters.Add(New SqlParameter("@1", id))
                    sqlComm.Parameters.Add(New SqlParameter("@2", reason))
                    sqlConn.Open()
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()
                    sqlConn.Close()
                    Return "1*Success"
                Catch ex As Exception
                    EventLog.WriteEntry("RTIS Vulcan SVC", "SVC_UpdateWhseTransferFailed: " + ex.ToString())
                    Return ExHandler.returnErrorEx(ex)
                End Try
            End Function
        End Class

        Partial Public Class Insert
            Public Shared Function SVC_InsertWHTLineCompleted(ByVal id As String, ByVal code As String, ByVal lotnumber As String, ByVal whseFrom As String, ByVal whseTo As String, ByVal qty As String,
                                                              ByVal dtEntered As String, ByVal username As String, ByVal process As String, ByVal transDesc As String) As String
                Try
                    Dim ReturnData As String = ""
                    Dim sqlConn As New SqlConnection(RTString)
                    Dim sqlComm As New SqlCommand(" INSERT INTO [tbl_WHTCompleted] 
                                                    ([iLineID], [vItemCode], [vLotNumber], [vWarehouse_From], [vWarehouse_To], [dQtyTransfered], [dtDateEntered], [dtDateTransfered], [vUsername], [vProcess], [vTransDesc]) 
                                                    VALUES 
                                                    (@1, @2, @3, @4, @5, @6, @7, GETDATE(), @8, @9, @10)", sqlConn)
                    sqlComm.Parameters.Add(New SqlParameter("@1", id))
                    sqlComm.Parameters.Add(New SqlParameter("@2", code))
                    sqlComm.Parameters.Add(New SqlParameter("@3", lotnumber))
                    sqlComm.Parameters.Add(New SqlParameter("@4", whseFrom))
                    sqlComm.Parameters.Add(New SqlParameter("@5", whseTo))
                    sqlComm.Parameters.Add(New SqlParameter("@6", qty))
                    sqlComm.Parameters.Add(New SqlParameter("@7", dtEntered))
                    sqlComm.Parameters.Add(New SqlParameter("@8", username))
                    sqlComm.Parameters.Add(New SqlParameter("@9", process))
                    sqlComm.Parameters.Add(New SqlParameter("@10", transDesc))
                    sqlConn.Open()
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()
                    sqlConn.Close()
                    Return "1*Success"
                Catch ex As Exception
                    EventLog.WriteEntry("RTIS Vulcan SVC", "SVC_InsertWHTLineCompleted: " + ex.ToString())
                    Return ExHandler.returnErrorEx(ex)
                End Try
            End Function
        End Class

        Partial Public Class Delete
            Public Shared Function SVC_DeleteWhseTransLineComplete(ByVal id As String) As String
                Try
                    Dim ReturnData As String = ""
                    Dim sqlConn As New SqlConnection(RTString)
                    Dim sqlComm As New SqlCommand(" DELETE FROM [tbl_WHTPending]  WHERE [iLineID] = @1", sqlConn)
                    sqlComm.Parameters.Add(New SqlParameter("@1", id))
                    sqlConn.Open()
                    sqlComm.ExecuteNonQuery()
                    sqlComm.Dispose()
                    sqlConn.Close()
                    Return "1*Success"
                Catch ex As Exception
                    EventLog.WriteEntry("RTIS Vulcan SVC", "SVC_DeleteWhseTransLineComplete: " + ex.ToString())
                    Return ExHandler.returnErrorEx(ex)
                End Try
            End Function
        End Class
    End Class
    Public Class EvolutionSDK
        Public Shared Function Initialize() As String
            Try
                DatabaseContext.CreateCommonDBConnection(My.Settings.EvoComServer, My.Settings.EvoComDB, My.Settings.EvoComUser, My.Settings.EvoComPassword, False)
                DatabaseContext.SetLicense("DE10110181", "7725371")
                DatabaseContext.CreateConnection(My.Settings.EvoServer, My.Settings.EvoDB, My.Settings.EvoUser, My.Settings.EvoPassword, False)
                Return "1*Success"
            Catch ex As Exception
                EventLog.WriteEntry("RTIS Vulcan SVC", "Server Response Evolution SDK Error: " + ex.Message)
                Return ExHandler.returnErrorEx(ex)
            End Try
        End Function

        Public Shared Function CGetItem(ByVal ItemCode As String) As InventoryItem
            Initialize()
            Dim errorItem As InventoryItem = New InventoryItem()
            errorItem.Code = "ErrorItem"
            Try
                Dim thisitem As InventoryItem = New InventoryItem(InventoryItem.Find("Code='" & ItemCode & "'"))
                Return thisitem
            Catch ex As Exception
                EventLog.WriteEntry("RTIS Vulcan SVC", "CGetItem: " + ex.ToString)
                errorItem.Description = ex.Message
                Return errorItem
            End Try
        End Function

        Public Shared Function CGetLot(ByVal lotNumber As String, ByVal stockId As String) As Lot
            Initialize()
            Dim errorLot As Lot = New Lot()
            errorLot.Code = lotNumber
            Try
                Dim thisitem As Lot = New Lot(Lot.Find("cLotDescription='" & lotNumber & "' AND iStockID = '" & stockId & "'"))
                Return thisitem
            Catch ex As Exception
                EventLog.WriteEntry("RTIS Vulcan SVC", "CGetItem: " + ex.ToString)
                Return errorLot
            End Try
        End Function

        Public Shared Function CGetWhse(ByVal whseCode As String) As Warehouse
            Try
                Dim thisitem As Warehouse = New Warehouse(Warehouse.Find("Code='" & whseCode & "'"))
                Return thisitem
            Catch ex As Exception
                EventLog.WriteEntry("RTIS Vulcan SVC", "CGetWhse: " + ex.ToString)
                Return Nothing
            End Try
        End Function

        Public Shared Function CTransferItem(ByVal OrderNo As String, ByVal FromWarehouseCode As String, ByVal ToWarehouseCode As String, ByVal ItemCode As String, ByVal Lot As String, ByVal Qty As String) As String
            Try
                Initialize()
                Dim thistransfer As New WarehouseTransfer()
                Dim invItem As New InventoryItem()
                invItem = CGetItem(ItemCode)
                If invItem.Code <> "ErrorItem" Then
                    thistransfer.Date = Now

                    thistransfer.Description = "RTIS Warehouse Transfer from: " + FromWarehouseCode + " to: " + ToWarehouseCode
                    thistransfer.ExtOrderNo = OrderNo
                    thistransfer.FromWarehouse = CGetWhse(FromWarehouseCode)
                    thistransfer.ToWarehouse = CGetWhse(ToWarehouseCode)
                    thistransfer.InventoryItem = invItem
                    thistransfer.Lot = CGetLot(Lot, invItem.ID)
                    thistransfer.Quantity = Convert.ToDouble(Qty.Replace(".", sep).Replace(",", sep))
                    thistransfer.Reference = OrderNo
                    If thistransfer.Validate() Then
                        thistransfer.Post()
                        Return "1*Success"
                    Else
                        Return "-1*The warehouse transfer could not be completed"
                    End If
                Else
                    Return invItem.Description
                End If
            Catch ex As Exception
                If ex.Message.Contains("Negative lot tracking quantity not allowed!") Then
                    EventLog.WriteEntry("RTIS SVC", "CTransferItem: " + ex.ToString)
                    Return "-1*Insufficient qty for lot"
                Else
                    EventLog.WriteEntry("RTIS SVC", "CTransferItem: " + ex.ToString)
                    Return "-1*Reason: " + ex.Message
                End If

            End Try
        End Function
    End Class
End Class
