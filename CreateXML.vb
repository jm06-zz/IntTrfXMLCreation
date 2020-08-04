Module CreateXML

    ' all the CSV and XML files are created within this module.
    ' It also adds all records that has been extracted to the upload history table
    ' there are 3 seperate instances where the folder configuration need to change if that is needed
    Sub CreateFiles()

        ' this module creates xml files for each of the distinct journal types

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader

        Dim sQuery As String
        sqCmd.Connection = sqCon
        sqCon.Open()

        sQuery = "SELECT DISTINCT JRNAL_TYPE FROM TFR_UPLOAD_CURRENT"
        sqCmd.CommandText = sQuery

        sdrRow = sqCmd.ExecuteReader

        ' loop through all the different journal types and create files for each
        While sdrRow.Read()

            CreateJrnlFile(sdrRow.GetValue(0))
            CreateJrnlFileCSV(sdrRow.GetValue(0))
        End While
        CreateFileCSV()
        sqCon.Close()

    End Sub

    ' procedure to create XML files
    Sub CreateJrnlFile(sTemplateType As String)

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader

        Dim sQuery As String
        Dim MyFileName As String
        Dim MyNewLine As String


        Dim sAnalysis0 As String
        Dim sAnalysis1 As String
        Dim sAnalysis2 As String
        Dim sAnalysis3 As String
        Dim sAnalysis4 As String
        Dim sAnalysis5 As String
        Dim sAnalysis6 As String
        Dim sAnalysis7 As String

        Dim sMonth As String
        Dim sDay As String

        sqCmd.Connection = sqCon
        sqCon.Open()

        sQuery = "SELECT ACCNT_CODE" & _
                          ",PERIOD" & _
                          ",TRANS_DATETIME" & _
                          ",AMOUNT" & _
                          ",D_C" & _
                          ",TREFERENCE" & _
                          ",DESCRIPTN" & _
                          ",CONV_CODE" & _
                          ",CONV_RATE" & _
                          ",OTHER_AMT" & _
                          ",ANAL_T0" & _
                          ",ANAL_T1" & _
                          ",ANAL_T2" & _
                          ",ANAL_T3" & _
                          ",ANAL_T4" & _
                          ",ANAL_T5" & _
                          ",ANAL_T6" & _
                          ",ANAL_T7 " & _
                          ",LINK_REF_1 " & _
                          "FROM TFR_UPLOAD_CURRENT WHERE JRNAL_TYPE = '" & sTemplateType & "'"
        sqCmd.CommandText = sQuery

        sdrRow = sqCmd.ExecuteReader

        ' local filename for testing
        MyFileName = TfrFunctions.sGetFolder("XML") & "Upload_" & Year(Now) & Month(Now) & Day(Now) & Second(Now) & "_" & Trim(sTemplateType) & ".xml"

        ' write header of xml file. NB the header needs to be in this format for transfer desk to recognise file
        Dim myWriter As New System.IO.StreamWriter(MyFileName)
        MyNewLine = "<?xml version=""1.0"" encoding=""utf-8""?>"
        myWriter.WriteLine(MyNewLine)
        MyNewLine = "<SSC>"
        myWriter.WriteLine(MyNewLine)
        MyNewLine = "<Payload>"
        myWriter.WriteLine(MyNewLine)
        MyNewLine = "<Ledger>"
        myWriter.WriteLine(MyNewLine)

        While sdrRow.Read()

            Try
                ' ensure that each analysis code has a value
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T0"))) = "" Then
                    sAnalysis0 = "X"
                Else
                    sAnalysis0 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T0")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T1"))) = "" Then
                    sAnalysis1 = "X"
                Else
                    sAnalysis1 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T1")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T2"))) = "" Then
                    sAnalysis2 = "X"
                Else
                    sAnalysis2 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T2")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T3"))) = "" Then
                    sAnalysis3 = "X"
                Else
                    sAnalysis3 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T3")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T4"))) = "" Then
                    sAnalysis4 = "X"
                Else
                    sAnalysis4 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T4")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T5"))) = "" Then
                    sAnalysis5 = "X"
                Else
                    sAnalysis5 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T5")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T6"))) = "" Then
                    sAnalysis6 = "X"
                Else
                    sAnalysis6 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T6")))
                End If
                If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7"))) = "" Then
                    sAnalysis7 = "X"
                Else
                    sAnalysis7 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7")))
                End If

                ' set day and month string values
                ' this is in order to create the date in the correct format for transfer desk to be picked up
                sMonth = Month(Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TRANS_DATETIME"))))
                sDay = Day(Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TRANS_DATETIME"))))

                If Len(sMonth) = 1 Then
                    sMonth = "0" & sMonth
                End If
                If Len(sDay) = 1 Then
                    sDay = "0" & sDay
                End If



                ' write lines to the xml file
                MyNewLine = "<Line>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AccountCode>" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ACCNT_CODE"))) & "</AccountCode>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AccountingPeriod>" & Right(Trim(sdrRow.GetValue(sdrRow.GetOrdinal("PERIOD"))), 3) & Left(Trim(sdrRow.GetValue(sdrRow.GetOrdinal("PERIOD"))), 4) & "</AccountingPeriod>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <TransactionDate>" & sDay & sMonth & Year(Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TRANS_DATETIME")))) & "</TransactionDate>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <BaseAmount>" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("AMOUNT"))) & "</BaseAmount>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <DebitCredit>" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("D_C"))) & "</DebitCredit>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <TransactionReference>" & TfrFunctions.sXMLFormat(sdrRow.GetValue(sdrRow.GetOrdinal("TREFERENCE"))) & "</TransactionReference>"
                myWriter.WriteLine(MyNewLine)
                If IsDBNull(sdrRow.GetValue(sdrRow.GetOrdinal("DESCRIPTN"))) Then
                    MyNewLine = "   <Description>No Description</Description>"
                Else
                    MyNewLine = "   <Description>" & TfrFunctions.sXMLFormat(sdrRow.GetValue(sdrRow.GetOrdinal("DESCRIPTN"))) & "</Description>"
                End If
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <CurrencyCode>" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("CONV_CODE"))) & "</CurrencyCode>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <TransactionAmount>" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("OTHER_AMT"))) & "</TransactionAmount>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode1>" & sAnalysis0 & "</AnalysisCode1>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode2>" & sAnalysis1 & "</AnalysisCode2>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode3>" & sAnalysis2 & "</AnalysisCode3>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode4>" & sAnalysis3 & "</AnalysisCode4>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode5>" & sAnalysis4 & "</AnalysisCode5>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode6>" & sAnalysis5 & "</AnalysisCode6>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode7>" & sAnalysis6 & "</AnalysisCode7>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <AnalysisCode8>" & sAnalysis7 & "</AnalysisCode8>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "   <Line_Ref_1>" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("LINK_REF_1"))) & "</Line_Ref_1>"
                myWriter.WriteLine(MyNewLine)
                MyNewLine = "</Line>"
                myWriter.WriteLine(MyNewLine)

            Catch e As Exception

                MsgBox(e.InnerException)
                MsgBox(e.Message & " " & sdrRow.GetValue(sdrRow.GetOrdinal("TREFERENCE")))
                MsgBox(MyNewLine)
            End Try
        End While
        ' closing header tags
        MyNewLine = "</Ledger>"
        myWriter.WriteLine(MyNewLine)
        MyNewLine = "</Payload>"
        myWriter.WriteLine(MyNewLine)
        MyNewLine = "</SSC>"
        myWriter.WriteLine(MyNewLine)
        myWriter.Close()

        sqCon.Close()

    End Sub
    ' similar to the XML file procedure but outputs the CSV equivalents
    Sub CreateJrnlFileCSV(sTemplateType As String)

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader
        Dim sDescription As String

        Dim sQuery As String
        Dim MyFileName As String
        Dim MyNewLine As String

        Dim sAnalysis0 As String
        Dim sAnalysis1 As String
        Dim sAnalysis2 As String
        Dim sAnalysis3 As String
        Dim sAnalysis4 As String
        Dim sAnalysis5 As String
        Dim sAnalysis6 As String
        Dim sAnalysis7 As String

        sqCmd.Connection = sqCon
        sqCon.Open()

        sQuery = "SELECT ACCNT_CODE" & _
                          ",PERIOD" & _
                          ",TRANS_DATETIME" & _
                          ",AMOUNT" & _
                          ",D_C" & _
                          ",TREFERENCE" & _
                          ",DESCRIPTN" & _
                          ",CONV_CODE" & _
                          ",CONV_RATE" & _
                          ",OTHER_AMT" & _
                          ",ANAL_T0" & _
                          ",ANAL_T1" & _
                          ",ANAL_T2" & _
                          ",ANAL_T3" & _
                          ",ANAL_T4" & _
                          ",ANAL_T5" & _
                          ",ANAL_T6" & _
                          ",ANAL_T7 " & _
                          ",LINK_REF_1 " & _
                          "FROM TFR_UPLOAD_CURRENT WHERE JRNAL_TYPE = '" & sTemplateType & "'"
        sqCmd.CommandText = sQuery

        sdrRow = sqCmd.ExecuteReader

        ' local filename for testing
        MyFileName = TfrFunctions.sGetFolder("CSV") & "Upload_" & Year(Now) & Month(Now) & Day(Now) & Second(Now) & "_" & Trim(sTemplateType) & ".csv"

        Console.WriteLine("Upload_" & Year(Now) & Month(Now) & Day(Now) & Second(Now) & "_" & Trim(sTemplateType) & ".csv")

        Dim myWriter As New System.IO.StreamWriter(MyFileName)

        While sdrRow.Read()
            ' ensure that analysis codes have a value
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T0"))) = "" Then
                sAnalysis0 = "X"
            Else
                sAnalysis0 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T0")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T1"))) = "" Then
                sAnalysis1 = "X"
            Else
                sAnalysis1 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T1")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T2"))) = "" Then
                sAnalysis2 = "X"
            Else
                sAnalysis2 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T2")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T3"))) = "" Then
                sAnalysis3 = "X"
            Else
                sAnalysis3 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T3")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T4"))) = "" Then
                sAnalysis4 = "X"
            Else
                sAnalysis4 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T4")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T5"))) = "" Then
                sAnalysis5 = "X"
            Else
                sAnalysis5 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T5")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T6"))) = "" Then
                sAnalysis6 = "X"
            Else
                sAnalysis6 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T6")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7"))) = "" Then
                sAnalysis7 = "X"
            Else
                sAnalysis7 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7")))
            End If

            If IsDBNull(sdrRow.GetValue(sdrRow.GetOrdinal("DESCRIPTN"))) Then
                sDescription = "No Description - CSV"
            Else
                sDescription = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("DESCRIPTN")))
            End If



            MyNewLine = """" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ACCNT_CODE"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("PERIOD"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TRANS_DATETIME"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("AMOUNT"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("D_C"))) & """,=""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TREFERENCE"))) & """,""" & _
                                sDescription & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("CONV_CODE"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("OTHER_AMT"))) & """,""" & _
                                sAnalysis0 & """,""" & _
                                sAnalysis1 & """,""" & _
                                sAnalysis2 & """,""" & _
                                sAnalysis3 & """,""" & _
                                sAnalysis4 & """,""" & _
                                sAnalysis5 & """,""" & _
                                sAnalysis6 & """,""" & _
                                sAnalysis7 & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("LINK_REF_1"))) & """"
            myWriter.WriteLine(MyNewLine)

        End While

        myWriter.Close()

        sqCon.Close()

    End Sub

    ' this procedure updates the upload history table with all the extracted transactions
    Sub AddHistory()

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sQuery As String
        sqCmd.Connection = sqCon
        sqCon.Open()

        sQuery = "INSERT INTO TFR_UPLOAD_HISTORY " & _
                                    "(UPLOAD_DATE,BU,BU_JRNL_JRNLINE,ACCNT_CODE, PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, AMOUNT,D_C,ALLOCATION,JRNAL_TYPE,JRNAL_SRCE,TREFERENCE," & _
                                    "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," & _
                                    "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG,ANAL_T0,ANAL_T1,ANAL_T2,ANAL_T3," & _
                                    "ANAL_T4,ANAL_T5,ANAL_T6" & _
                                      ",ANAL_T7" & _
                                      ",ANAL_T8" & _
                                      ",ANAL_T9" & _
                                      ",POSTING_DATETIME" & _
                                      ",ALLOC_IN_PROGRESS" & _
                                      ",HOLD_REF" & _
                                      ",HOLD_OP_ID" & _
                                      ",BASE_RATE" & _
                                      ",BASE_OPERATOR" & _
                                      ",CONV_OPERATOR" & _
                                      ",REPORT_RATE" & _
                                      ",REPORT_OPERATOR" & _
                                      ",REPORT_AMT" & _
                                      ",MEMO_AMT" & _
                                      ",EXCLUDE_BAL" & _
                                      ",LE_DETAILS_IND" & _
                                      ",CONSUMED_BDGT_ID" & _
                                      ",CV4_CONV_CODE" & _
                                      ",CV4_AMT" & _
                                      ",CV4_CONV_RATE" & _
                                      ",CV4_OPERATOR" & _
                                      ",CV4_DP" & _
                                      ",CV5_CONV_CODE" & _
                                      ",CV5_AMT" & _
                                      ",CV5_CONV_RATE" & _
                                      ",CV5_OPERATOR" & _
                                      ",CV5_DP" & _
                                      ",LINK_REF_1" & _
                                      ",LINK_REF_2" & _
                                      ",LINK_REF_3" & _
                                      ",ALLOCN_CODE" & _
                                      ",ALLOCN_STMNTS" & _
                                      ",OPR_CODE" & _
                                      ",SPLIT_ORIG_LINE" & _
                                      ",VAL_DATETIME" & _
                                      ",SIGNING_DETAILS" & _
                                      ",INSTLMT_DATETIME" & _
                                      ",PRINCIPAL_REQD" & _
                                      ",BINDER_STATUS" & _
                                      ",AGREED_STATUS" & _
                                      ",SPLIT_LINK_REF" & _
                                      ",PSTG_REF" & _
                                      ",TRUE_RATED" & _
                                      ",HOLD_DATETIME" & _
                                      ",HOLD_TEXT" & _
                                      ",INSTLMT_NUM" & _
                                      ",SUPPLMNTRY_EXTSN" & _
                                      ",APRVLS_EXTSN" & _
                                      ",REVAL_LINK_REF" & _
                                      ",SAVED_SET_NUM" & _
                                      ",AUTHORISTN_SET_REF" & _
                                      ",PYMT_AUTHORISTN_SET_REF" & _
                                      ",MAN_PAY_OVER" & _
                                      ",PYMT_STAMP" & _
                                      ",AUTHORISTN_IN_PROGRESS" & _
                                      ",SPLIT_IN_PROGRESS" & _
                                      ",VCHR_NUM" & _
                                      ",JNL_CLASS_CODE" & _
                                      ",ORIGINATOR_ID" & _
                                      ",ORIGINATED_DATETIME" & _
                                      ",LAST_CHANGE_USER_ID" & _
                                      ",LAST_CHANGE_DATETIME" & _
                                      ",AFTER_PSTG_ID" & _
                                      ",AFTER_PSTG_DATETIME" & _
                                      ",POSTER_ID" & _
                                      ",ALLOC_ID" & _
                                      ",JNL_REVERSAL_TYPE) SELECT CURRENT_TIMESTAMP," & _
                                      "BU" & _
                                      ",(BU + '_' + LTRIM(STR(JRNAL_NO)) + '_' + LTRIM(STR(JRNAL_LINE))) AS BU_JRNAL_NO_JRNAL_LINE," & _
                                      "ACCNT_CODE, " & _
                                      "PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, ABS(AMOUNT), D_C, ALLOCATION, JRNAL_TYPE, JRNAL_SRCE, TREFERENCE, " & _
                                      "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," & _
                                      "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG,ANAL_T0,ANAL_T1,ANAL_T2,ANAL_T3," & _
                                      "ANAL_T4,ANAL_T5,ANAL_T6" & _
                                      ",ANAL_T7" & _
                                      ",ANAL_T8" & _
                                      ",ANAL_T9" & _
                                      ",POSTING_DATETIME" & _
                                      ",ALLOC_IN_PROGRESS" & _
                                      ",HOLD_REF" & _
                                      ",HOLD_OP_ID" & _
                                      ",BASE_RATE" & _
                                      ",BASE_OPERATOR" & _
                                      ",CONV_OPERATOR" & _
                                      ",REPORT_RATE" & _
                                      ",REPORT_OPERATOR" & _
                                      ",REPORT_AMT" & _
                                      ",MEMO_AMT" & _
                                      ",EXCLUDE_BAL" & _
                                      ",LE_DETAILS_IND" & _
                                      ",CONSUMED_BDGT_ID" & _
                                      ",CV4_CONV_CODE" & _
                                      ",CV4_AMT" & _
                                      ",CV4_CONV_RATE" & _
                                      ",CV4_OPERATOR" & _
                                      ",CV4_DP" & _
                                      ",CV5_CONV_CODE" & _
                                      ",CV5_AMT" & _
                                      ",CV5_CONV_RATE" & _
                                      ",CV5_OPERATOR" & _
                                      ",CV5_DP" & _
                                      ",LINK_REF_1" & _
                                      ",LINK_REF_2" & _
                                      ",LINK_REF_3" & _
                                      ",ALLOCN_CODE" & _
                                      ",ALLOCN_STMNTS" & _
                                      ",OPR_CODE" & _
                                      ",SPLIT_ORIG_LINE" & _
                                      ",VAL_DATETIME" & _
                                      ",SIGNING_DETAILS" & _
                                      ",INSTLMT_DATETIME" & _
                                      ",PRINCIPAL_REQD" & _
                                      ",BINDER_STATUS" & _
                                      ",AGREED_STATUS" & _
                                      ",SPLIT_LINK_REF" & _
                                      ",PSTG_REF" & _
                                      ",TRUE_RATED" & _
                                      ",HOLD_DATETIME" & _
                                      ",HOLD_TEXT" & _
                                      ",INSTLMT_NUM" & _
                                      ",SUPPLMNTRY_EXTSN" & _
                                      ",APRVLS_EXTSN" & _
                                      ",REVAL_LINK_REF" & _
                                      ",SAVED_SET_NUM" & _
                                      ",AUTHORISTN_SET_REF" & _
                                      ",PYMT_AUTHORISTN_SET_REF" & _
                                      ",MAN_PAY_OVER" & _
                                      ",PYMT_STAMP" & _
                                      ",AUTHORISTN_IN_PROGRESS" & _
                                      ",SPLIT_IN_PROGRESS" & _
                                      ",VCHR_NUM" & _
                                      ",JNL_CLASS_CODE" & _
                                      ",ORIGINATOR_ID" & _
                                      ",ORIGINATED_DATETIME" & _
                                      ",LAST_CHANGE_USER_ID" & _
                                      ",LAST_CHANGE_DATETIME" & _
                                      ",AFTER_PSTG_ID" & _
                                      ",AFTER_PSTG_DATETIME" & _
                                      ",POSTER_ID" & _
                                      ",ALLOC_ID" & _
                                      ",JNL_REVERSAL_TYPE FROM TFR_UPLOAD_CURRENT"

        sqCmd.CommandText = sQuery
        sqCmd.ExecuteNonQuery()
        sqCon.Close()

    End Sub


    ' this procedure creates a file combining all of the different journal types
    Sub CreateFileCSV()

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader

        Dim sQuery As String
        Dim MyFileName As String
        Dim MyNewLine As String

        Dim sAnalysis0 As String
        Dim sAnalysis1 As String
        Dim sAnalysis2 As String
        Dim sAnalysis3 As String
        Dim sAnalysis4 As String
        Dim sAnalysis5 As String
        Dim sAnalysis6 As String
        Dim sAnalysis7 As String

        Dim sDescription As String


        sqCmd.Connection = sqCon
        sqCon.Open()

        sQuery = "SELECT ACCNT_CODE" & _
                          ",PERIOD" & _
                          ",TRANS_DATETIME" & _
                          ",AMOUNT" & _
                          ",D_C" & _
                          ",TREFERENCE" & _
                          ",DESCRIPTN" & _
                          ",CONV_CODE" & _
                          ",CONV_RATE" & _
                          ",OTHER_AMT" & _
                          ",ANAL_T0" & _
                          ",ANAL_T1" & _
                          ",ANAL_T2" & _
                          ",ANAL_T3" & _
                          ",ANAL_T4" & _
                          ",ANAL_T5" & _
                          ",ANAL_T6" & _
                          ",ANAL_T7 " & _
                          ",LINK_REF_1 " & _
                          ",JRNAL_TYPE " & _
                          "FROM TFR_UPLOAD_CURRENT"
        sqCmd.CommandText = sQuery

        sdrRow = sqCmd.ExecuteReader

        ' local filename for testing
        MyFileName = TfrFunctions.sGetFolder("CSV") & "Upload_" & Year(Now) & Month(Now) & Day(Now) & Second(Now) & "_ALL_JOURNALS.csv"

        Console.WriteLine("Upload_" & Year(Now) & Month(Now) & Day(Now) & Second(Now) & "_ALL_JOURNALS.csv")
        Dim myWriter As New System.IO.StreamWriter(MyFileName)


        MyNewLine = "ACCOUNT_CODE," & _
                                    "PERIOD," & _
                                    "TRANSACTION_DATE," & _
                                    "AMOUNT," & _
                                    "DEBIT_CREDIT," & _
                                    "TRX_REFERENCE," & _
                                    "DESCRIPTION," & _
                                    "CONV_CODE," & _
                                    "OTHER_AMOUNT," & _
                                    "T1," & _
                                    "T2," & _
                                    "T3," & _
                                    "T4," & _
                                    "T5," & _
                                    "T6," & _
                                    "T7," & _
                                    "T8," & _
                                    "LINK_REF_1," & _
                                    "JRNAL_TYPE"

        myWriter.WriteLine(MyNewLine)


        While sdrRow.Read()

            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T0"))) = "" Then
                sAnalysis0 = "X"
            Else
                sAnalysis0 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T0")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T1"))) = "" Then
                sAnalysis1 = "X"
            Else
                sAnalysis1 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T1")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T2"))) = "" Then
                sAnalysis2 = "X"
            Else
                sAnalysis2 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T2")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T3"))) = "" Then
                sAnalysis3 = "X"
            Else
                sAnalysis3 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T3")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T4"))) = "" Then
                sAnalysis4 = "X"
            Else
                sAnalysis4 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T4")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T5"))) = "" Then
                sAnalysis5 = "X"
            Else
                sAnalysis5 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T5")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T6"))) = "" Then
                sAnalysis6 = "X"
            Else
                sAnalysis6 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T6")))
            End If
            If Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7"))) = "" Then
                sAnalysis7 = "X"
            Else
                sAnalysis7 = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7")))
            End If

            If IsDBNull(sdrRow.GetValue(sdrRow.GetOrdinal("DESCRIPTN"))) Then
                sDescription = "No Description - CSV"
            Else
                sDescription = Trim(sdrRow.GetValue(sdrRow.GetOrdinal("DESCRIPTN")))
            End If


            MyNewLine = """" & Trim(sdrRow.GetValue(sdrRow.GetOrdinal("ACCNT_CODE"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("PERIOD"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TRANS_DATETIME"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("AMOUNT"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("D_C"))) & """,=""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("TREFERENCE"))) & """,""" & _
                                sDescription & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("CONV_CODE"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("OTHER_AMT"))) & """,""" & _
                                sAnalysis0 & """,""" & _
                                sAnalysis1 & """,""" & _
                                sAnalysis2 & """,""" & _
                                sAnalysis3 & """,""" & _
                                sAnalysis4 & """,""" & _
                                sAnalysis5 & """,""" & _
                                sAnalysis6 & """,""" & _
                                sAnalysis7 & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("LINK_REF_1"))) & """,""" & _
                                Trim(sdrRow.GetValue(sdrRow.GetOrdinal("JRNAL_TYPE"))) & """"

            myWriter.WriteLine(MyNewLine)

        End While

        myWriter.Close()

        sqCon.Close()

    End Sub

End Module
