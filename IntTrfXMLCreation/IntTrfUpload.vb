Module IntTrfUpload
    Public sState As String
    Sub Main()

        Dim sBusinessUnit As String
        Dim lMaxJrnl As Long
        Dim sPeriod As String

        ' connection to the database
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand

        Dim sdrRow As SqlClient.SqlDataReader
        Dim ErrorMail As New EMail

        ' find the current state of the application (test or live)

        sqCmd.Connection = sqCon
        sqCon.Open()
        sqCmd.CommandText = "SELECT CurrentState FROM TFR_CONFIG"
        sdrRow = sqCmd.ExecuteReader()
        While sdrRow.Read()
            sState = sdrRow.GetValue(sdrRow.GetOrdinal("CurrentState"))
        End While
        sqCon.Close()

        ' we will check for error first and send out notification if any files remains un-imported
        If TfrFunctions.sCompileEMail <> "No Mail" Then
            ' need to compile an error e-mail and send to recipients
            ErrorMail.SetConnectionSettings()
            ErrorMail.sBody = TfrFunctions.sCompileEMail()
            ErrorMail.sSubject = "Transfer Error - Auto Generated E-Mail"
            ErrorMail.SendErrorMail()
            Console.WriteLine("Error Found in previous uploads. E-Mail Sent")

        Else
            Console.WriteLine("All imports up to date - no error and no e-mail sent")
        End If

        ' truncate the current_upload table that will be filled with transactions to export
        Call TfrFunctions.TruncateUpload()

        ' need to loop through business units and add transactions to the upload current table

        ' read business units from database which are live
        sqCmd.Connection = sqCon
        sqCmd.CommandText = "SELECT BusinessUnit FROM dbo.TFR_BU_DEFINITION where Status = 'live'"

        sqCon.Open()
        sdrRow = sqCmd.ExecuteReader()

        ' each business unit is handled individually
        While sdrRow.Read()
            sBusinessUnit = sdrRow.GetValue(sdrRow.GetOrdinal("BusinessUnit"))

            ' determine the last journal number that has been created for the specific business unit
            lMaxJrnl = TfrFunctions.GetMaxJournal(sBusinessUnit)
            ' determine period from the business unit definition table
            sPeriod = TfrFunctions.GetFromPeriod(sBusinessUnit)

            ' now add any new transactions within the local office bu view to the upload current table
            Call AddTrxData.AddLocalData(sBusinessUnit, lMaxJrnl, sPeriod)

        End While
        sdrRow.Close()

        ' add the additional lines that relates to the transactions that have been added from the local offices to the current upload table
        ' based on the T8 codes
        Call AddTrxData.AddAdditionalLines()
        ' create the output files
        Call CreateXML.CreateFiles()
        ' add the data that has been output to files to the upload history table
        Call AddHistory()

        ' close database connection
        sqCon.Close()

    End Sub

End Module
