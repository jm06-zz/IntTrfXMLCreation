Module AddTrxData

    ' this procedure adds all the relevant transaction data from the views into the UPLOAD_CURRENT table
    Sub AddLocalData(sBusinessUnit As String, lMaxJrnlNo As Long, sPeriod As String)
        ' looks at the view on SUN and select new transactions to be added to migration process
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand

        sqCmd.Connection = sqCon
        sqCon.Open()

        ' debtor and creditor account transformations occur within this query
        sqCmd.CommandText = "INSERT INTO TFR_UPLOAD_CURRENT " &
                                    "(BU,ACCNT_CODE, PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, AMOUNT,D_C,ALLOCATION,JRNAL_TYPE,JRNAL_SRCE,TREFERENCE," &
                                    "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," &
                                    "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG,ANAL_T0,ANAL_T1,ANAL_T2,ANAL_T3," &
                                    "ANAL_T4,ANAL_T5,ANAL_T6" &
                                      ",ANAL_T7" &
                                      ",ANAL_T8" &
                                      ",ANAL_T9" &
                                      ",POSTING_DATETIME" &
                                      ",ALLOC_IN_PROGRESS" &
                                      ",HOLD_REF" &
                                      ",HOLD_OP_ID" &
                                      ",BASE_RATE" &
                                      ",BASE_OPERATOR" &
                                      ",CONV_OPERATOR" &
                                      ",REPORT_RATE" &
                                      ",REPORT_OPERATOR" &
                                      ",REPORT_AMT" &
                                      ",MEMO_AMT" &
                                      ",EXCLUDE_BAL" &
                                      ",LE_DETAILS_IND" &
                                      ",CONSUMED_BDGT_ID" &
                                      ",CV4_CONV_CODE" &
                                      ",CV4_AMT" &
                                      ",CV4_CONV_RATE" &
                                      ",CV4_OPERATOR" &
                                      ",CV4_DP" &
                                      ",CV5_CONV_CODE" &
                                      ",CV5_AMT" &
                                      ",CV5_CONV_RATE" &
                                      ",CV5_OPERATOR" &
                                      ",CV5_DP" &
                                      ",LINK_REF_1" &
                                      ",LINK_REF_2" &
                                      ",LINK_REF_3" &
                                      ",ALLOCN_CODE" &
                                      ",ALLOCN_STMNTS" &
                                      ",OPR_CODE" &
                                      ",SPLIT_ORIG_LINE" &
                                      ",VAL_DATETIME" &
                                      ",SIGNING_DETAILS" &
                                      ",INSTLMT_DATETIME" &
                                      ",PRINCIPAL_REQD" &
                                      ",BINDER_STATUS" &
                                      ",AGREED_STATUS" &
                                      ",SPLIT_LINK_REF" &
                                      ",PSTG_REF" &
                                      ",TRUE_RATED" &
                                      ",HOLD_DATETIME" &
                                      ",HOLD_TEXT" &
                                      ",INSTLMT_NUM" &
                                      ",SUPPLMNTRY_EXTSN" &
                                      ",APRVLS_EXTSN" &
                                      ",REVAL_LINK_REF" &
                                      ",SAVED_SET_NUM" &
                                      ",AUTHORISTN_SET_REF" &
                                      ",PYMT_AUTHORISTN_SET_REF" &
                                      ",MAN_PAY_OVER" &
                                      ",PYMT_STAMP" &
                                      ",AUTHORISTN_IN_PROGRESS" &
                                      ",SPLIT_IN_PROGRESS" &
                                      ",VCHR_NUM" &
                                      ",JNL_CLASS_CODE" &
                                      ",ORIGINATOR_ID" &
                                      ",ORIGINATED_DATETIME" &
                                      ",LAST_CHANGE_USER_ID" &
                                      ",LAST_CHANGE_DATETIME" &
                                      ",AFTER_PSTG_ID" &
                                      ",AFTER_PSTG_DATETIME" &
                                      ",POSTER_ID" &
                                      ",ALLOC_ID" &
                                      ",JNL_REVERSAL_TYPE) SELECT " &
                                      "'" & sBusinessUnit & "'," &
                                      "CASE " &
                                        "WHEN LEFT(ACCNT_CODE,1)='D' AND ACCNT_CODE <> 'DMABAN' AND LEFT(ACCNT_CODE, 4) <> 'DSDN' " &
                                        "THEN 'DZ' + '" & sBusinessUnit & "' + 'GEN' " &
                                        "WHEN LEFT(ACCNT_CODE,1)='C' THEN '8101' + '" & sBusinessUnit & "' " &
                                        "ELSE ACCNT_CODE " &
                                      "END, " &
                                      "PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, ABS(AMOUNT), D_C, ALLOCATION, JRNAL_TYPE, JRNAL_SRCE, TREFERENCE, " &
                                      "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," &
                                      "CONV_CODE,CONV_RATE,ABS(OTHER_AMT),OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG,ANAL_T0, " &
                                      "ANAL_T1," &
                                      "ANAL_T2," &
                                      "ANAL_T3," &
                                      "ANAL_T4," &
                                      "CASE " &
                                        "WHEN LEN(ANAL_T0) > 0 " &
                                        "THEN 'X' ELSE ANAL_T5 " &
                                      "END, " &
                                      "CASE " &
                                        "WHEN LEN(ANAL_T0) > 0 " &
                                        "THEN 'X' ELSE ANAL_T6 " &
                                      "END " &
                                      ",ANAL_T7" &
                                      ",ANAL_T8" &
                                      ",ANAL_T9" &
                                      ",POSTING_DATETIME" &
                                      ",ALLOC_IN_PROGRESS" &
                                      ",HOLD_REF" &
                                      ",HOLD_OP_ID" &
                                      ",BASE_RATE" &
                                      ",BASE_OPERATOR" &
                                      ",CONV_OPERATOR" &
                                      ",REPORT_RATE" &
                                      ",REPORT_OPERATOR" &
                                      ",REPORT_AMT" &
                                      ",MEMO_AMT" &
                                      ",EXCLUDE_BAL" &
                                      ",LE_DETAILS_IND" &
                                      ",CONSUMED_BDGT_ID" &
                                      ",CV4_CONV_CODE" &
                                      ",CV4_AMT" &
                                      ",CV4_CONV_RATE" &
                                      ",CV4_OPERATOR" &
                                      ",CV4_DP" &
                                      ",CV5_CONV_CODE" &
                                      ",CV5_AMT" &
                                      ",CV5_CONV_RATE" &
                                      ",CV5_OPERATOR" &
                                      ",CV5_DP" &
                                      ",('" & sBusinessUnit & "_' + LTRIM(STR(JRNAL_NO)) + '_' + LTRIM(STR(JRNAL_LINE))) AS BU_JRNAL_NO_JRNAL_LINE" &
                                      ",LINK_REF_2" &
                                      ",LINK_REF_3" &
                                      ",ALLOCN_CODE" &
                                      ",ALLOCN_STMNTS" &
                                      ",OPR_CODE" &
                                      ",SPLIT_ORIG_LINE" &
                                      ",VAL_DATETIME" &
                                      ",SIGNING_DETAILS" &
                                      ",INSTLMT_DATETIME" &
                                      ",PRINCIPAL_REQD" &
                                      ",BINDER_STATUS" &
                                      ",AGREED_STATUS" &
                                      ",SPLIT_LINK_REF" &
                                      ",PSTG_REF" &
                                      ",TRUE_RATED" &
                                      ",HOLD_DATETIME" &
                                      ",HOLD_TEXT" &
                                      ",INSTLMT_NUM" &
                                      ",SUPPLMNTRY_EXTSN" &
                                      ",APRVLS_EXTSN" &
                                      ",REVAL_LINK_REF" &
                                      ",SAVED_SET_NUM" &
                                      ",AUTHORISTN_SET_REF" &
                                      ",PYMT_AUTHORISTN_SET_REF" &
                                      ",MAN_PAY_OVER" &
                                      ",PYMT_STAMP" &
                                      ",AUTHORISTN_IN_PROGRESS" &
                                      ",SPLIT_IN_PROGRESS" &
                                      ",VCHR_NUM" &
                                      ",JNL_CLASS_CODE" &
                                      ",ORIGINATOR_ID" &
                                      ",ORIGINATED_DATETIME" &
                                      ",LAST_CHANGE_USER_ID" &
                                      ",LAST_CHANGE_DATETIME" &
                                      ",AFTER_PSTG_ID" &
                                      ",AFTER_PSTG_DATETIME" &
                                      ",POSTER_ID" &
                                      ",ALLOC_ID" &
                                      ",JNL_REVERSAL_TYPE FROM TFR_LEDGERS_VIEW_ALL WHERE BU = '" & sBusinessUnit & "' AND JRNAL_NO > '" & lMaxJrnlNo & "' AND PERIOD >= '" & sPeriod & "' AND ABS(AMOUNT) > 0"

        sqCmd.ExecuteNonQuery()
        sqCon.Close()

    End Sub

    ' this procedure adds all the additional journal lines based on the T8 codes that require us to do so
    Sub AddAdditionalLines()
        ' this is the procedure to add in the 4 legged journals to the current upload table. 
        ' this is irrespective of the business unit so any transaction line that qualifies are analysed and 2 additional lines added
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader
        Dim sAccount As String
        Dim sT8 As String
        Dim Id As Long
        Dim sReserveAccount As String
        Dim sAllocationCode As String
        Dim sBusinessUnit As String
        Dim sDrCr As String
        Dim lMaxId As Long
        sqCmd.Connection = sqCon
        sqCon.Open()

        sqCmd.CommandText = "SELECT MAX(ID) FROM TFR_UPLOAD_CURRENT"
        sdrRow = sqCmd.ExecuteReader()

        sdrRow.Read()

        If sdrRow.IsDBNull(0) Then
            Exit Sub
        Else
            lMaxId = sdrRow.GetValue(0)

            sdrRow.Close()
            ' if there are records in the upload current table and any of the t8 codes and account codes meet the criteria then we need to add additional rows
            sqCmd.CommandText = "SELECT ID, BU, ACCNT_CODE, D_C, ANAL_T7 FROM TFR_UPLOAD_CURRENT WHERE " & _
                                    "(LEFT(ANAL_T7,3) = 'ICD' OR LEFT(ANAL_T7,2) = 'SD' OR LEFT(ANAL_T7,2) = 'DF') AND " & _
                                    "(LEFT(ACCNT_CODE,1) >= '1' AND LEFT(ACCNT_CODE,1) <= '4') AND ID <= " & lMaxId



            sdrRow = sqCmd.ExecuteReader()

            While sdrRow.Read()

                ' determine the various details to obtain the correct details for the allocation and reserve account rows
                sAccount = sdrRow.GetValue(sdrRow.GetOrdinal("ACCNT_CODE"))
                sT8 = sdrRow.GetValue(sdrRow.GetOrdinal("ANAL_T7"))
                Id = sdrRow.GetValue(sdrRow.GetOrdinal("ID"))
                sBusinessUnit = sdrRow.GetValue(sdrRow.GetOrdinal("BU"))
                sDrCr = sdrRow.GetValue(sdrRow.GetOrdinal("D_C"))

                ' determine account codes for reserve and allocation rows
                sReserveAccount = sGetReserveAccount(sT8)
                sAllocationCode = sGetAllocationCode(sAccount)

                ' add allocation row only if the account code does not start with 11
                If Left(sAccount, 2) <> "11" And sT8 <> "SD0001" Then ' update for Salesforce - added SD0001 filter
                    Call AddAllocationRow(Id, sAllocationCode, sBusinessUnit, sDrCr)
                    Call AddReserveAccountRow(Id, sReserveAccount, sBusinessUnit)
                End If

            End While
        End If

        sqCon.Close()

    End Sub


    ' execute the SQL query to add the allocatoion row with correct allocation code
    Sub AddAllocationRow(Id As Long, sAllocationCode As String, sBusinessUnit As String, sDrCr As String)
        ' create a new row with the allocation code
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sQuery As String
        sqCmd.Connection = sqCon
        sqCon.Open()

        If sDrCr = "D" Then
            sDrCr = "C"
        Else
            sDrCr = "D"
        End If
        sQuery = "INSERT INTO TFR_UPLOAD_CURRENT " &
                                    "(BU,ACCNT_CODE, PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, AMOUNT,D_C,ALLOCATION,JRNAL_TYPE,JRNAL_SRCE,TREFERENCE," &
                                    "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," &
                                    "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG,ANAL_T0,ANAL_T1,ANAL_T2,ANAL_T3," &
                                    "ANAL_T4,ANAL_T5,ANAL_T6" &
                                      ",ANAL_T7" &
                                      ",ANAL_T8" &
                                      ",ANAL_T9" &
                                      ",POSTING_DATETIME" &
                                      ",ALLOC_IN_PROGRESS" &
                                      ",HOLD_REF" &
                                      ",HOLD_OP_ID" &
                                      ",BASE_RATE" &
                                      ",BASE_OPERATOR" &
                                      ",CONV_OPERATOR" &
                                      ",REPORT_RATE" &
                                      ",REPORT_OPERATOR" &
                                      ",REPORT_AMT" &
                                      ",MEMO_AMT" &
                                      ",EXCLUDE_BAL" &
                                      ",LE_DETAILS_IND" &
                                      ",CONSUMED_BDGT_ID" &
                                      ",CV4_CONV_CODE" &
                                      ",CV4_AMT" &
                                      ",CV4_CONV_RATE" &
                                      ",CV4_OPERATOR" &
                                      ",CV4_DP" &
                                      ",CV5_CONV_CODE" &
                                      ",CV5_AMT" &
                                      ",CV5_CONV_RATE" &
                                      ",CV5_OPERATOR" &
                                      ",CV5_DP" &
                                      ",LINK_REF_1" &
                                      ",LINK_REF_2" &
                                      ",LINK_REF_3" &
                                      ",ALLOCN_CODE" &
                                      ",ALLOCN_STMNTS" &
                                      ",OPR_CODE" &
                                      ",SPLIT_ORIG_LINE" &
                                      ",VAL_DATETIME" &
                                      ",SIGNING_DETAILS" &
                                      ",INSTLMT_DATETIME" &
                                      ",PRINCIPAL_REQD" &
                                      ",BINDER_STATUS" &
                                      ",AGREED_STATUS" &
                                      ",SPLIT_LINK_REF" &
                                      ",PSTG_REF" &
                                      ",TRUE_RATED" &
                                      ",HOLD_DATETIME" &
                                      ",HOLD_TEXT" &
                                      ",INSTLMT_NUM" &
                                      ",SUPPLMNTRY_EXTSN" &
                                      ",APRVLS_EXTSN" &
                                      ",REVAL_LINK_REF" &
                                      ",SAVED_SET_NUM" &
                                      ",AUTHORISTN_SET_REF" &
                                      ",PYMT_AUTHORISTN_SET_REF" &
                                      ",MAN_PAY_OVER" &
                                      ",PYMT_STAMP" &
                                      ",AUTHORISTN_IN_PROGRESS" &
                                      ",SPLIT_IN_PROGRESS" &
                                      ",VCHR_NUM" &
                                      ",JNL_CLASS_CODE" &
                                      ",ORIGINATOR_ID" &
                                      ",ORIGINATED_DATETIME" &
                                      ",LAST_CHANGE_USER_ID" &
                                      ",LAST_CHANGE_DATETIME" &
                                      ",AFTER_PSTG_ID" &
                                      ",AFTER_PSTG_DATETIME" &
                                      ",POSTER_ID" &
                                      ",ALLOC_ID" &
                                      ",JNL_REVERSAL_TYPE) SELECT " &
                                      "'" & sBusinessUnit & "'," &
                                      "'" & sAllocationCode & "', " &
                                      "PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, ABS(AMOUNT), '" & sDrCr & "', ALLOCATION, JRNAL_TYPE, JRNAL_SRCE, TREFERENCE, " &
                                      "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," &
                                      "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG,ANAL_T0, ANAL_T1,ANAL_T2,ANAL_T3," &
                                      "ANAL_T4,'X','X'" &
                                      ",ANAL_T7" &
                                      ",ANAL_T8" &
                                      ",ANAL_T9" &
                                      ",POSTING_DATETIME" &
                                      ",ALLOC_IN_PROGRESS" &
                                      ",HOLD_REF" &
                                      ",HOLD_OP_ID" &
                                      ",BASE_RATE" &
                                      ",BASE_OPERATOR" &
                                      ",CONV_OPERATOR" &
                                      ",REPORT_RATE" &
                                      ",REPORT_OPERATOR" &
                                      ",REPORT_AMT" &
                                      ",MEMO_AMT" &
                                      ",EXCLUDE_BAL" &
                                      ",LE_DETAILS_IND" &
                                      ",CONSUMED_BDGT_ID" &
                                      ",CV4_CONV_CODE" &
                                      ",CV4_AMT" &
                                      ",CV4_CONV_RATE" &
                                      ",CV4_OPERATOR" &
                                      ",CV4_DP" &
                                      ",CV5_CONV_CODE" &
                                      ",CV5_AMT" &
                                      ",CV5_CONV_RATE" &
                                      ",CV5_OPERATOR" &
                                      ",CV5_DP" &
                                      ",('" & sBusinessUnit & "_' + LTRIM(STR(JRNAL_NO)) + '_' + LTRIM(STR(JRNAL_LINE))) AS BU_JRNAL_NO_JRNAL_LINE" &
                                      ",LINK_REF_2" &
                                      ",LINK_REF_3" &
                                      ",ALLOCN_CODE" &
                                      ",ALLOCN_STMNTS" &
                                      ",OPR_CODE" &
                                      ",SPLIT_ORIG_LINE" &
                                      ",VAL_DATETIME" &
                                      ",SIGNING_DETAILS" &
                                      ",INSTLMT_DATETIME" &
                                      ",PRINCIPAL_REQD" &
                                      ",BINDER_STATUS" &
                                      ",AGREED_STATUS" &
                                      ",SPLIT_LINK_REF" &
                                      ",PSTG_REF" &
                                      ",TRUE_RATED" &
                                      ",HOLD_DATETIME" &
                                      ",HOLD_TEXT" &
                                      ",INSTLMT_NUM" &
                                      ",SUPPLMNTRY_EXTSN" &
                                      ",APRVLS_EXTSN" &
                                      ",REVAL_LINK_REF" &
                                      ",SAVED_SET_NUM" &
                                      ",AUTHORISTN_SET_REF" &
                                      ",PYMT_AUTHORISTN_SET_REF" &
                                      ",MAN_PAY_OVER" &
                                      ",PYMT_STAMP" &
                                      ",AUTHORISTN_IN_PROGRESS" &
                                      ",SPLIT_IN_PROGRESS" &
                                      ",VCHR_NUM" &
                                      ",JNL_CLASS_CODE" &
                                      ",ORIGINATOR_ID" &
                                      ",ORIGINATED_DATETIME" &
                                      ",LAST_CHANGE_USER_ID" &
                                      ",LAST_CHANGE_DATETIME" &
                                      ",AFTER_PSTG_ID" &
                                      ",AFTER_PSTG_DATETIME" &
                                      ",POSTER_ID" &
                                      ",ALLOC_ID" &
                                      ",JNL_REVERSAL_TYPE FROM TFR_UPLOAD_CURRENT WHERE ID = '" & Id & "'"

        sqCmd.CommandText = sQuery
        sqCmd.ExecuteNonQuery()
        sqCon.Close()

    End Sub

    ' execute the SQL query to add the reserve row with correct allocation code
    Sub AddReserveAccountRow(Id As Long, sReserveAccount As String, sBusinessUnit As String)
        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sQuery As String
        sqCmd.Connection = sqCon
        sqCon.Open()

        sQuery = "INSERT INTO TFR_UPLOAD_CURRENT " &
                                    "(BU,ACCNT_CODE, PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, AMOUNT,D_C,ALLOCATION,JRNAL_TYPE,JRNAL_SRCE,TREFERENCE," &
                                    "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," &
                                    "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG" &
                                    ",ANAL_T0" &
                                      ",ANAL_T1" &
                                      ",ANAL_T2" &
                                      ",ANAL_T3" &
                                      ",ANAL_T4" &
                                      ",ANAL_T5" &
                                      ",ANAL_T6" &
                                      ",ANAL_T7" &
                                      ",ANAL_T8" &
                                      ",ANAL_T9" &
                                      ",POSTING_DATETIME" &
                                      ",ALLOC_IN_PROGRESS" &
                                      ",HOLD_REF" &
                                      ",HOLD_OP_ID" &
                                      ",BASE_RATE" &
                                      ",BASE_OPERATOR" &
                                      ",CONV_OPERATOR" &
                                      ",REPORT_RATE" &
                                      ",REPORT_OPERATOR" &
                                      ",REPORT_AMT" &
                                      ",MEMO_AMT" &
                                      ",EXCLUDE_BAL" &
                                      ",LE_DETAILS_IND" &
                                      ",CONSUMED_BDGT_ID" &
                                      ",CV4_CONV_CODE" &
                                      ",CV4_AMT" &
                                      ",CV4_CONV_RATE" &
                                      ",CV4_OPERATOR" &
                                      ",CV4_DP" &
                                      ",CV5_CONV_CODE" &
                                      ",CV5_AMT" &
                                      ",CV5_CONV_RATE" &
                                      ",CV5_OPERATOR" &
                                      ",CV5_DP" &
                                      ",LINK_REF_1" &
                                      ",LINK_REF_2" &
                                      ",LINK_REF_3" &
                                      ",ALLOCN_CODE" &
                                      ",ALLOCN_STMNTS" &
                                      ",OPR_CODE" &
                                      ",SPLIT_ORIG_LINE" &
                                      ",VAL_DATETIME" &
                                      ",SIGNING_DETAILS" &
                                      ",INSTLMT_DATETIME" &
                                      ",PRINCIPAL_REQD" &
                                      ",BINDER_STATUS" &
                                      ",AGREED_STATUS" &
                                      ",SPLIT_LINK_REF" &
                                      ",PSTG_REF" &
                                      ",TRUE_RATED" &
                                      ",HOLD_DATETIME" &
                                      ",HOLD_TEXT" &
                                      ",INSTLMT_NUM" &
                                      ",SUPPLMNTRY_EXTSN" &
                                      ",APRVLS_EXTSN" &
                                      ",REVAL_LINK_REF" &
                                      ",SAVED_SET_NUM" &
                                      ",AUTHORISTN_SET_REF" &
                                      ",PYMT_AUTHORISTN_SET_REF" &
                                      ",MAN_PAY_OVER" &
                                      ",PYMT_STAMP" &
                                      ",AUTHORISTN_IN_PROGRESS" &
                                      ",SPLIT_IN_PROGRESS" &
                                      ",VCHR_NUM" &
                                      ",JNL_CLASS_CODE" &
                                      ",ORIGINATOR_ID" &
                                      ",ORIGINATED_DATETIME" &
                                      ",LAST_CHANGE_USER_ID" &
                                      ",LAST_CHANGE_DATETIME" &
                                      ",AFTER_PSTG_ID" &
                                      ",AFTER_PSTG_DATETIME" &
                                      ",POSTER_ID" &
                                      ",ALLOC_ID" &
                                      ",JNL_REVERSAL_TYPE) SELECT " &
                                      "'" & sBusinessUnit & "'," &
                                      "'" & sReserveAccount & "', " &
                                      "PERIOD, TRANS_DATETIME, JRNAL_NO, JRNAL_LINE, ABS(AMOUNT), D_C, ALLOCATION, JRNAL_TYPE, JRNAL_SRCE, TREFERENCE, " &
                                      "DESCRIPTN,ENTRY_DATETIME,ENTRY_PRD,DUE_DATETIME,ALLOC_REF,ALLOC_DATETIME,ALLOC_PERIOD,ASSET_IND,ASSET_CODE,ASSET_SUB," &
                                      "CONV_CODE,CONV_RATE,OTHER_AMT,OTHER_DP,CLEARDOWN,REVERSAL,LOSS_GAIN,ROUGH_FLAG,IN_USE_FLAG" &
                                      ",ANAL_T0" &
                                      ",ANAL_T1" &
                                      ",ANAL_T2" &
                                      ",ANAL_T3" &
                                      ",ANAL_T4" &
                                      ",'X'" &
                                      ",'X'" &
                                      ",ANAL_T7" &
                                      ",ANAL_T8" &
                                      ",ANAL_T9" &
                                      ",POSTING_DATETIME" &
                                      ",ALLOC_IN_PROGRESS" &
                                      ",HOLD_REF" &
                                      ",HOLD_OP_ID" &
                                      ",BASE_RATE" &
                                      ",BASE_OPERATOR" &
                                      ",CONV_OPERATOR" &
                                      ",REPORT_RATE" &
                                      ",REPORT_OPERATOR" &
                                      ",REPORT_AMT" &
                                      ",MEMO_AMT" &
                                      ",EXCLUDE_BAL" &
                                      ",LE_DETAILS_IND" &
                                      ",CONSUMED_BDGT_ID" &
                                      ",CV4_CONV_CODE" &
                                      ",CV4_AMT" &
                                      ",CV4_CONV_RATE" &
                                      ",CV4_OPERATOR" &
                                      ",CV4_DP" &
                                      ",CV5_CONV_CODE" &
                                      ",CV5_AMT" &
                                      ",CV5_CONV_RATE" &
                                      ",CV5_OPERATOR" &
                                      ",CV5_DP" &
                                      ",('" & sBusinessUnit & "_' + LTRIM(STR(JRNAL_NO)) + '_' + LTRIM(STR(JRNAL_LINE))) AS BU_JRNAL_NO_JRNAL_LINE" &
                                      ",LINK_REF_2" &
                                      ",LINK_REF_3" &
                                      ",ALLOCN_CODE" &
                                      ",ALLOCN_STMNTS" &
                                      ",OPR_CODE" &
                                      ",SPLIT_ORIG_LINE" &
                                      ",VAL_DATETIME" &
                                      ",SIGNING_DETAILS" &
                                      ",INSTLMT_DATETIME" &
                                      ",PRINCIPAL_REQD" &
                                      ",BINDER_STATUS" &
                                      ",AGREED_STATUS" &
                                      ",SPLIT_LINK_REF" &
                                      ",PSTG_REF" &
                                      ",TRUE_RATED" &
                                      ",HOLD_DATETIME" &
                                      ",HOLD_TEXT" &
                                      ",INSTLMT_NUM" &
                                      ",SUPPLMNTRY_EXTSN" &
                                      ",APRVLS_EXTSN" &
                                      ",REVAL_LINK_REF" &
                                      ",SAVED_SET_NUM" &
                                      ",AUTHORISTN_SET_REF" &
                                      ",PYMT_AUTHORISTN_SET_REF" &
                                      ",MAN_PAY_OVER" &
                                      ",PYMT_STAMP" &
                                      ",AUTHORISTN_IN_PROGRESS" &
                                      ",SPLIT_IN_PROGRESS" &
                                      ",VCHR_NUM" &
                                      ",JNL_CLASS_CODE" &
                                      ",ORIGINATOR_ID" &
                                      ",ORIGINATED_DATETIME" &
                                      ",LAST_CHANGE_USER_ID" &
                                      ",LAST_CHANGE_DATETIME" &
                                      ",AFTER_PSTG_ID" &
                                      ",AFTER_PSTG_DATETIME" &
                                      ",POSTER_ID" &
                                      ",ALLOC_ID" &
                                      ",JNL_REVERSAL_TYPE FROM TFR_UPLOAD_CURRENT WHERE ID = '" & Id & "'"

        sqCmd.CommandText = sQuery
        sqCmd.ExecuteNonQuery()
        sqCon.Close()


    End Sub

End Module
