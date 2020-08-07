Option Explicit On

Module TfrFunctions

    Function SQL_Connectionstring() As String
        ' set the connection string for the relevant database
        Dim sConnection As String
        '        MsgBox(sState)
        '       If sState = "Live" Then
        sConnection = "Server=HQDB1; Database=ExclBudget; Trusted_Connection=True;"
        'Else
        ' connection string for local testing
        ' sConnection = "Server=JM-EXCLUSERV\SQLEXPRESS; Database=Fin_Budget; Trusted_Connection=True;"
        'End If
        Return sConnection
    End Function

    ' this function returns the correct Reserve account which should be paired to the relevant T8 analysis code when creating additional journal lines
    Function sGetReserveAccount(sT8 As String) As String

        Dim sReserveAccountTemp As String

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader


        ' see if T8 code mapping is present in the transfer mapping table
        sqCmd.Connection = sqCon
        sqCon.Open()

        sqCmd.CommandText = "SELECT ReserveAccount FROM TFR_MAPPING WHERE T8 = '" & sT8 & "'"

        sdrRow = sqCmd.ExecuteReader

        ' sdrRow.Read()
        sReserveAccountTemp = "X"
        ' If sdrRow.IsDBNull(sdrRow.GetOrdinal("ReserveAccount")) Then
        If sdrRow.Read() Then
            sGetReserveAccount = sdrRow.GetValue(sdrRow.GetOrdinal("ReserveAccount"))

        Else
            ' if the name cannot be found in the mapping table then need to determine the name through logic and then match to the accounts table
            If Left(sT8, 3) = "ICD" Then
                sReserveAccountTemp = "Q" & sT8 ' the legth of this should be 7 and correspond to one of the account codes in the central business unit
            ElseIf Left(sT8, 2) = "SD" Then
                sReserveAccountTemp = "R" & sT8 ' length should be 7
            ElseIf Left(sT8, 2) = "DF" Then
                sReserveAccountTemp = "S" & sT8 ' length should be 7
            End If
            ' retrieve the account from the accounts table by looking for the first 7 characters
            sGetReserveAccount = sGetAccount(sReserveAccountTemp)
        End If
        sqCon.Close()

    End Function

    ' function to truncate the current upload table
    Sub TruncateUpload()
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim lNumRows As Long

        Try
            sqCmd.Connection = sqCon
            sqCon.Open()

            sqCmd.CommandText = "DELETE FROM TFR_UPLOAD_CURRENT"
            lNumRows = sqCmd.ExecuteNonQuery()
            sqCon.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)

        End Try

        Exit Sub
ErrorHandler:
        MsgBox("There was an error truncating the current upload table.")
    End Sub

    ' select the maximum journal number from the upload history table based on the business unit
    Function GetMaxJournal(sBusinessUnit As String) As Long
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand

        Dim sdrMaxJrnl As SqlClient.SqlDataReader
        Dim lMaxJrnl As Long

        sqCmd.Connection = sqCon
        sqCon.Open()

        sqCmd.CommandText = "SELECT MAX(JRNAL_NO) as MAX_JRNL_NO  from TFR_UPLOAD_HISTORY where BU = '" & sBusinessUnit & "'"
        sdrMaxJrnl = sqCmd.ExecuteReader()
        sdrMaxJrnl.Read()

        ' if no records ie then the max journal number is given the value 0
        If sdrMaxJrnl.IsDBNull(sdrMaxJrnl.GetOrdinal("MAX_JRNL_NO")) Then
            lMaxJrnl = 0
        Else
            lMaxJrnl = sdrMaxJrnl.GetValue(sdrMaxJrnl.GetOrdinal("MAX_JRNL_NO"))
        End If
        GetMaxJournal = lMaxJrnl

        sqCon.Close()

    End Function

    ' select period that the application should import data from
    Function GetFromPeriod(sBusinessUnit As String) As String
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand

        Dim sdrMaxJrnl As SqlClient.SqlDataReader
        Dim sPeriod As String

        sqCmd.Connection = sqCon
        sqCon.Open()

        sqCmd.CommandText = "SELECT PeriodFrom  from TFR_BU_DEFINITION where BusinessUnit = '" & sBusinessUnit & "'"
        sdrMaxJrnl = sqCmd.ExecuteReader()
        sdrMaxJrnl.Read()

        ' if no period is selected then the period is set from 2013001 - this should never be the case if the busines unit has been set up in the TFR_BU_DEFINITION table
        If sdrMaxJrnl.IsDBNull(sdrMaxJrnl.GetOrdinal("PeriodFrom")) Then
            sPeriod = "2013001"
        Else
            sPeriod = sdrMaxJrnl.GetValue(sdrMaxJrnl.GetOrdinal("PeriodFrom"))
        End If
        GetFromPeriod = sPeriod

        sqCon.Close()

    End Function

    ' get the account from the live CAF accounts table (TFR_CENTRAL_ACCOUNTS) view
    Function sGetAccount(sAccountCode As String) As String
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand

        Dim sdrRow As SqlClient.SqlDataReader

        sqCmd.Connection = sqCon
        sqCon.Open()

        ' check to match first 7 characters
        sqCmd.CommandText = "SELECT ACNT_CODE FROM TFR_CENTRAL_ACCOUNTS where (LEFT(ACNT_CODE,7) = '" & sAccountCode & "' OR LEFT(ACNT_CODE,8) = '" & sAccountCode & "')"
        ' changed to allow for 
        sdrRow = sqCmd.ExecuteReader()

        If sdrRow.Read() Then
            sGetAccount = sdrRow.GetValue(sdrRow.GetOrdinal("ACNT_CODE"))
        Else
            ' if no account can be found in either the mapping table or through the logic applied then the account is given the code UNMATCHED
            sGetAccount = "UNMATCHED"
        End If

        sqCon.Close()

    End Function

    Function sGetAllocationCode(sAccountCode As String) As String

        ' the allocation code is determined by looking at whether the account code is an income or expenditure code

        If Left(CStr(sAccountCode), 1) = "1" Then
            ' income code requires some logic to be applied

            sGetAllocationCode = "52" & Right(Trim(sAccountCode), 2)
        Else
            sGetAllocationCode = "5215"
        End If

    End Function

    Function sXMLFormat(sSource As String) As String
        Dim sTemp As String

        sTemp = Replace(Trim(sSource), "&", "&amp;")
        sTemp = Replace(sTemp, """", "&quot;")
        sTemp = Replace(sTemp, "'", "&apos;")
        sTemp = Replace(sTemp, "<", "&lt;")
        sTemp = Replace(sTemp, ">", "&gt;")

        sXMLFormat = sTemp

    End Function

    
    Function sGetFolder(sType As String) As String
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sQuery As String
        Dim sdrRow As SqlClient.SqlDataReader
        Dim sFolder As String

        sFolder = "No Folder"

        ' sTest = TfrFunctions.sCurrentState()
        sqCmd.Connection = sqCon
        sqCon.Open()
        sQuery = "SELECT Folder FROM TFR_FOLDERS WHERE Test = '" & sState & "' AND Type = '" & sType & "'"
        sqCmd.CommandText = sQuery
        sdrRow = sqCmd.ExecuteReader()

        While sdrRow.Read()
            sFolder = sdrRow.GetValue(sdrRow.GetOrdinal("Folder"))
        End While


        sGetFolder = sFolder
        sqCon.Close()

    End Function

    Function sCompileEMail() As String
        ' this function exist to compile the error message to be sent to the recipient email addresses

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader

        Try
            sqCmd.Connection = sqCon
            sqCon.Open()
            sqCmd.CommandText = "SELECT DISTINCT 'Date Produced: ' + CAST(UPLOAD_DATE as nvarchar(50)) + ' - Journal Type: ' + JRNAL_TYPE FROM TFR_ERROR_LINES"
            sdrRow = sqCmd.ExecuteReader

            If sdrRow.Read() Then
                ' a reminder needs to be sent
                sCompileEMail = "Errors have occured in files created on the following dates for the specified journal types. Please review: " + vbCrLf + vbCrLf
                sCompileEMail = sCompileEMail + sdrRow.GetValue(0) + vbCrLf

                While sdrRow.Read()
                    sCompileEMail = sCompileEMail + vbCrLf + sdrRow.GetValue(0) + vbCrLf
                End While
                sCompileEMail = sCompileEMail + vbCrLf + vbCrLf + "Please find and import outstanding files to eliminate errors."

#If DEBUG Then
                DirLogAppend.Log("reminder sCompileEMail() " & sCompileEMail.ToString)
#End If

            Else
                ' no error occurred
                sCompileEMail = "No Mail"

#If DEBUG Then
                DirLogAppend.Log("no error sCompileEMail() " & sCompileEMail.ToString)
#End If
            End If
        Catch ex As Exception
            sCompileEMail = "Error in reading the error view! Please contact developer - " & ex.Message
        End Try


    End Function



End Module
