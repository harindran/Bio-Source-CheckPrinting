
Imports System.Drawing
Imports System.Windows.Forms

Namespace CheckPrinting_ProdOrdRoundOff
    Public Class ClsManufacturingOrder
        Public Const Formtype = "CT_PF_ManufacOrd"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Dim odbdsHeader, odbdsDetails As SAPbouiCOM.DBDataSource
        Dim CalcFlag As Boolean = False

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "Items" And pVal.ColUID = "U_OQty" Then
                                BubbleEvent = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnDCP" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If objform.Items.Item("9").Specific.Selected.Value = "RL" Then BubbleEvent = False : Exit Sub
                                'If objform.Items.Item("etDCP").Specific.String <> "" Then
                                '    If objaddon.objapplication.MessageBox("Do you want to adjust the Planned Quantity?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                                'End If
                            End If


                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            Try

                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            If pVal.ActionSuccess Then
                                CreateButton(FormUID)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            CalcFlag = False
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If pVal.ItemUID = "11" And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ItemChanged = True Then
                                If CalcFlag = True Then CalcFlag = False
                            ElseIf pVal.ItemUID = "Quantity" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) And pVal.ItemChanged = True Then
                                If CalcFlag = True Then CalcFlag = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            If (pVal.ItemUID = "11" Or pVal.ItemUID = "Quantity") And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If CalcFlag = True Then Exit Sub
                                'If Val(objform.Items.Item("etDCP").Specific.String) <> 0 Then Exit Sub
                                objmatrix = objform.Items.Item("Items").Specific
                                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_MOR3")
                                odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OMOR")
                                Dim Qty As Double
                                For i As Integer = 0 To odbdsDetails.Size - 1
                                    If odbdsDetails.GetValue("U_ItemCode", i) = "" Then Continue For
                                    odbdsDetails.SetValue("U_OQty", i, odbdsDetails.GetValue("U_Quantity", i) * odbdsHeader.GetValue("U_Quantity", 0))
                                    Qty = CDbl(odbdsDetails.GetValue("U_Result", i))
                                    CalcFlag = True
                                    If Math.Round(Qty, 1) = 0 Then
                                        odbdsDetails.SetValue("U_Result", i, 0.1)
                                    Else
                                        odbdsDetails.SetValue("U_Result", i, Math.Round(odbdsDetails.GetValue("U_Quantity", i) * odbdsHeader.GetValue("U_Quantity", 0), 1))
                                    End If
                                Next
                                objmatrix.LoadFromDataSource()
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "txtTANK" And pVal.CharPressed = 9 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnDCP" And (objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                If objform.Items.Item("9").Specific.Selected.Value = "RL" Then Exit Sub
                                If objform.Items.Item("etDCP").Specific.String <> "" Then
                                    Rounding_PlannedQty(FormUID, CInt(objform.Items.Item("etDCP").Specific.String))
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            CreateButton(BusinessObjectInfo.FormUID)
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Sub CreateButton(ByVal FormUID As String)
            Try
                Dim objItem As SAPbouiCOM.Item
                'Dim objLabel As SAPbouiCOM.StaticText
                Dim objButton As SAPbouiCOM.Button
                Dim objedit As SAPbouiCOM.EditText
                'Dim objlink As SAPbouiCOM.LinkedButton
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Try
                    If objform.Items.Item("btnDCP").UniqueID = "btnDCP" Then Exit Sub
                Catch ex As Exception
                End Try

                objItem = objform.Items.Add("btnDCP", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                objItem.Left = objform.Items.Item("17").Left + objform.Items.Item("17").Width + 10
                'objItem.Left = objForm.Items.Item("10002056").Left + objForm.Items.Item("10002056").Width + 60
                Dim Fieldsize As Size = TextRenderer.MeasureText("Adjust Decimal", New Font("Arial", 12.0F))
                objItem.Width = Fieldsize.Width '120
                objItem.Top = objform.Items.Item("17").Top
                objItem.Height = 19 ' objform.Items.Item("6").Height
                objItem.LinkTo = "17"
                objButton = objItem.Specific
                objButton.Caption = "Adjust Decimal"

                'Dim objedit As SAPbouiCOM.EditText
                objItem = objform.Items.Add("etDCP", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                objItem.Left = objform.Items.Item("btnDCP").Left + objform.Items.Item("btnDCP").Width + 5
                objItem.Width = 50
                objItem.Top = objform.Items.Item("17").Top
                objItem.Height = objform.Items.Item("17").Height ' 14 'objform.Items.Item("btnDCP").Height
                objItem.LinkTo = "btnDCP"
                objedit = objItem.Specific
                'objedit.Item.Enabled = False
                objedit.DataBind.SetBound(True, "@CT_PF_OMOR", "U_AdjDec")


            Catch ex As Exception
            End Try

        End Sub

        Private Function Rounding_PlannedQty(ByVal FormUID As String, ByVal Rounddec As Integer)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("Items").Specific
                odbdsDetails = objform.DataSources.DBDataSources.Item("@CT_PF_MOR3")
                odbdsHeader = objform.DataSources.DBDataSources.Item("@CT_PF_OMOR")
                Dim PlanQty As Double
                objform.Freeze(True)
                For i As Integer = 0 To odbdsDetails.Size - 1
                    If odbdsDetails.GetValue("U_ItemCode", i) = "" Then Continue For
                    PlanQty = CDbl(odbdsDetails.GetValue("U_OQty", i))
                    If Math.Round(PlanQty, Rounddec) = 0 Then
                        odbdsDetails.SetValue("U_Result", i, 0.1)
                    Else
                        odbdsDetails.SetValue("U_Result", i, Math.Round(CDbl(odbdsDetails.GetValue("U_OQty", i)), Rounddec)) 'If Math.Round(PlanQty, Rounddec) <> 0.1 Then 
                    End If
                Next
                objmatrix.LoadFromDataSource()
                objaddon.objapplication.StatusBar.SetText("Planned Qty adjusted successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objmatrix.Columns.Item("col_1").Cells.Item(1).Click()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Function

    End Class

End Namespace
