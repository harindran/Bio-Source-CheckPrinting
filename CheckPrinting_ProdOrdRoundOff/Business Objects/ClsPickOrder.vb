Imports System.Drawing
Imports System.Windows.Forms

Namespace CheckPrinting_ProdOrdRoundOff
    Public Class ClsPickOrder
        Public Const Formtype = "CT_PF_PickOrder"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Dim odbdsHeader, odbdsDetails As SAPbouiCOM.DBDataSource

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnqty" Then
                                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then BubbleEvent = False : Exit Sub
                                'If objaddon.objapplication.MessageBox("Do you want to load the Picked Quantity?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                            End If


                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            MOCount = False
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            MOCount = True
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnqty" Then
                                Get_PickedQty(FormUID)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                            If pVal.ActionSuccess And MOCount = True Then
                                CreateButton(FormUID)
                            End If

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub


        Private Sub CreateButton(ByVal FormUID As String)
            Try
                Dim objItem As SAPbouiCOM.Item
                Dim objButton As SAPbouiCOM.Button
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Try
                    If objform.Items.Item("btnqty").UniqueID = "btnqty" Then Exit Sub
                Catch ex As Exception
                End Try

                objItem = objform.Items.Add("btnqty", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                objItem.Left = objform.Items.Item("porStCb").Left + objform.Items.Item("porStCb").Width + 10
                'objItem.Left = objForm.Items.Item("10002056").Left + objForm.Items.Item("10002056").Width + 60
                Dim Fieldsize As Size = TextRenderer.MeasureText("Pick Qty", New Font("Arial", 12.0F))
                objItem.Width = Fieldsize.Width '120
                objItem.Top = objform.Items.Item("porStCb").Top
                objItem.Height = 19 ' objform.Items.Item("6").Height
                objItem.LinkTo = "porStCb"
                objButton = objItem.Specific
                objButton.Caption = "Pick Qty"
                objform.Items.Item("btnqty").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoFormMode.fm_FIND_MODE, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Get_PickedQty(ByVal FormUID As String)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OPOR")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_POR1")
                If odbdsHeader.GetValue("U_Ref2", 0) <> "" Then Exit Sub
                Dim objmatrix As SAPbouiCOM.Matrix
                objmatrix = objform.Items.Item("porRqIMtx").Specific
                'strsql = objmatrix.Columns.Item("rqIBDECol").Cells.Item(1).Specific.string
                'strsql = odbdsDetails.GetValue("U_BaseEntry", 0)
                strSQL = "Select T1.""U_Quantity"" ""Header Qty"" ,T0.""U_Result"" ""Line Planned Qty"" from ""@CT_PF_MOR3"" T0 join ""@CT_PF_OMOR"" T1 On T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""=" & odbdsDetails.GetValue("U_BaseEntry", 0) & ""
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(strSQL)
                If objRs.RecordCount = 0 Then Exit Sub
                objaddon.objapplication.StatusBar.SetText("Picked Qty loading. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                odbdsHeader.SetValue("U_Ref2", 0, objRs.Fields.Item("Header Qty").Value)
                For i As Integer = 0 To objRs.RecordCount - 1
                    odbdsDetails.SetValue("U_PickedQty", i, objRs.Fields.Item("Line Planned Qty").Value)
                    'If i = 0 Then MOCount = False
                    objRs.MoveNext()
                Next
                objmatrix.LoadFromDataSource()
                objaddon.objapplication.StatusBar.SetText("Picked Qty loaded successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception

            End Try
        End Sub

    End Class
End Namespace

