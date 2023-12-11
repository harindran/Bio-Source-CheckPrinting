Imports SAPbouiCOM.Framework
Namespace CheckPrinting_ProdOrdRoundOff
    Public Class ClsBatchSetup
        Public Const Formtype = "41"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset


        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("3").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            Try

                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Try
                                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then Exit Sub
                                Assign_Batch(FormUID, ProdOrderEntry)
                            Catch ex As Exception

                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_CLICK


                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Assign_Batch(ByVal FormUID As String, ByVal ProductionOrder_Entry As String)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("3").Specific
                'If objaddon.HANA Then
                '    strSQL = "select ""U_MND_BTNO"" ""Batch_No"",To_Varchar(""U_MND_EXPD"",'yyyyMMdd') ""Exp_Date"",To_Varchar(""U_MND_MFGD"",'yyyyMMdd') ""Mfg_Date"" from OWOR where ""DocEntry""=" & ProductionOrder_Entry & ""
                'Else
                '    strSQL = "select U_MND_BTNO Batch_No,U_MND_EXPD Exp_Date,U_MND_MFGD Mfg_Date from OWOR where DocEntry=" & ProductionOrder_Entry & ""
                'End If
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRs.DoQuery(strSQL)
                Dim Batch As String = objaddon.objglobalmethods.getSingleValue("select ""U_MND_BTNO"" ""Batch_No"" from OWOR where ""DocEntry""=" & ProductionOrder_Entry & "")
                Dim ExpDate As String = objaddon.objglobalmethods.getSingleValue("select To_Varchar(""U_MND_EXPD"",'yyyyMMdd') ""Exp_Date"" from OWOR where ""DocEntry""=" & ProductionOrder_Entry & "")
                Dim MfgDate As String = objaddon.objglobalmethods.getSingleValue("select To_Varchar(""U_MND_MFGD"",'yyyyMMdd') ""Mfg_Date""from OWOR where ""DocEntry""=" & ProductionOrder_Entry & "")
                objmatrix.Columns.Item("2").Cells.Item(1).Specific.String = Batch 'objRs.Fields.Item(0).Value.ToString 
                objmatrix.Columns.Item("10").Cells.Item(1).Specific.String = ExpDate
                objmatrix.Columns.Item("11").Cells.Item(1).Specific.String = MfgDate
                objmatrix.Columns.Item("2").Cells.Item(1).Click()
                objaddon.objapplication.StatusBar.SetText("Batches Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Assign_Batch :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

    End Class
End Namespace

