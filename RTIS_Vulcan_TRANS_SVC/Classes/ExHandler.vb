Public Class ExHandler

    Public Shared Function returnErrorEx(ByVal exc As Exception) As String
        Dim msg As String = exc.Message
        Dim info As String = String.Empty
        Dim st As StackTrace = New StackTrace(exc, True)
        Dim line As String = String.Empty
        Dim name As String = String.Empty
        Dim meth As String = String.Empty
        For Each frame As StackFrame In st.GetFrames()
            Try
                line = frame.GetFileLineNumber().ToString()
                name = frame.GetFileName().ToString()
                meth = frame.GetMethod().ToString()
            Catch ex As Exception

            End Try
        Next
        info = exc.ToString() + Environment.NewLine + Environment.NewLine + "Method:" + Environment.NewLine + meth + Environment.NewLine + Environment.NewLine + "File: " + Environment.NewLine + name + Environment.NewLine + Environment.NewLine + "Line Number: " + Environment.NewLine + line
        Return "-1*" + msg + "|" + info
    End Function

End Class
