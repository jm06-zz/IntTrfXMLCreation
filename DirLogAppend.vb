Imports System.IO

Module DirLogAppend

    Dim logFilePath As String = "C:/Users/jmadmin/Documents/TransferDesk/log.txt"

    Sub Log(logMessage As String)
        Dim log As String = DateString + " " + TimeString + ": " + logMessage
        Dim objWriter As New System.IO.StreamWriter(logFilePath, True)
        objWriter.WriteLine(log)
        objWriter.Close()
    End Sub

End Module
