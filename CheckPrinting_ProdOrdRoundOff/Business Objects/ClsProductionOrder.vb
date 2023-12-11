Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace CheckPrinting_ProdOrdRoundOff
    Public Class ClsProductionOrder
        Public Const Formtype = "65211"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("37").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnDCP" And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If objform.Items.Item("etDCP").Specific.String <> "" Then
                                    objform.Freeze(True)
                                    For i As Integer = 1 To objmatrix.VisualRowCount
                                        If objmatrix.Columns.Item("4").Cells.Item(i).Specific.String = "" Then Continue For
                                        If CDbl(objmatrix.Columns.Item("U_MND_PLNQTY").Cells.Item(i).Specific.String) <> 0 Then Exit For
                                        objmatrix.Columns.Item("U_MND_PLNQTY").Cells.Item(i).Specific.String = objmatrix.Columns.Item("14").Cells.Item(i).Specific.String
                                    Next
                                    objform.Freeze(False)
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            Try

                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.ActionSuccess Then
                                CreateButton(FormUID)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "txtTANK" And pVal.CharPressed = 9 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If objform.Items.Item("6").Specific.String = "" Then Exit Sub
                                If Not objaddon.FormExist("SELFRM") Then
                                    If objaddon.HANA Then
                                        TankQuery = "Select T1.""U_MND_TANK"" as ""Tank"",T1.""U_MND_SIZE"" ""Tank Size"" from ""@MND_OTNKM"" T0 join ""@MND_TNKM1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_MND_TANK""<>'' and T1.""U_MND_ACT""='Y'"
                                    Else
                                        TankQuery = "Select T1.U_MND_TANK as Tank,T1.U_MND_SIZE Tank Size from [@MND_OTNKM] T0 join [@MND_TNKM1] T1 on T0.Code=T1.Code where T1.U_MND_TANK<>'' and T1.U_MND_ACT='Y'"
                                    End If
                                    objform.Freeze(True)
                                    FrmSelForm = objaddon.objapplication.Forms.ActiveForm
                                    Dim activeform As New FrmSelectionForm
                                    activeform.Show()
                                    bModal = True
                                    activeform.UIAPIRawForm.Left = objform.Left + 450
                                    activeform.UIAPIRawForm.Top = objform.Top + 200
                                    objform.Freeze(False)
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnDCP" And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
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

                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Sub CreateButton(ByVal FormUID As String)
            Try
                Dim objItem As SAPbouiCOM.Item
                Dim objLabel As SAPbouiCOM.StaticText
                Dim objButton As SAPbouiCOM.Button
                Dim objedit As SAPbouiCOM.EditText
                'Dim objlink As SAPbouiCOM.LinkedButton
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub

                objItem = objform.Items.Add("etTANK", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                objItem.Left = objform.Items.Item("540000152").Left
                objItem.Width = objform.Items.Item("540000152").Width '80
                objItem.Top = objform.Items.Item("540000152").Top + objform.Items.Item("540000152").Height + 2
                objItem.Height = 14
                objLabel = objItem.Specific
                objLabel.Caption = "TANK"


                objItem = objform.Items.Add("txtTANK", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                objItem.Left = objform.Items.Item("540000153").Left '+ objform.Items.Item("52").Width + 10
                objItem.Width = objform.Items.Item("540000153").Width
                objItem.Top = objform.Items.Item("etTANK").Top
                objItem.Height = objform.Items.Item("6").Height '14' objform.Items.Item("etTANK").Height
                objItem.LinkTo = "etTANK"
                objedit = objItem.Specific
                'objedit.Item.Enabled = False
                objedit.DataBind.SetBound(True, "OWOR", "U_MND_TANK")
                'objItem.Enabled = False

                objItem = objform.Items.Add("btnDCP", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                objItem.Left = objform.Items.Item("76").Left + objform.Items.Item("76").Width + 10
                'objItem.Left = objForm.Items.Item("10002056").Left + objForm.Items.Item("10002056").Width + 60
                objItem.Width = 120
                objItem.Top = objform.Items.Item("76").Top
                objItem.Height = 19 ' objform.Items.Item("6").Height
                objItem.LinkTo = "76"
                objButton = objItem.Specific
                objButton.Caption = "Adjust Decimal Point"


                'Dim objedit As SAPbouiCOM.EditText
                objItem = objform.Items.Add("etDCP", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                objItem.Left = objform.Items.Item("btnDCP").Left + objform.Items.Item("btnDCP").Width + 5
                objItem.Width = 50
                objItem.Top = objform.Items.Item("76").Top
                objItem.Height = objform.Items.Item("76").Height ' 14 'objform.Items.Item("btnDCP").Height
                objItem.LinkTo = "btnDCP"
                objedit = objItem.Specific
                'objedit.Item.Enabled = False
                objedit.DataBind.SetBound(True, "OWOR", "U_MND_DCMP")


            Catch ex As Exception
            End Try

        End Sub

        Private Function Rounding_PlannedQty(ByVal FormUID As String, ByVal Rounddec As Integer)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("37").Specific
                Dim PlanQty As Double

                objform.Freeze(True)
                For i As Integer = 1 To objmatrix.VisualRowCount
                    If objmatrix.Columns.Item("4").Cells.Item(i).Specific.String = "" Then Continue For
                    PlanQty = CDbl(objmatrix.Columns.Item("14").Cells.Item(i).Specific.String)
                    'Dim ConQty As Double = Math.Round(PlanQty, Rounddec)
                    If Math.Round(PlanQty, Rounddec) = 0 Then 'Math.Floor(PlanQty) 
                        Dim DecVal As Integer = PlanQty.ToString().Substring(PlanQty.ToString().IndexOf(".")).Length - 1
                        'Dim ss As String = CStr(ConQty) + "." + CStr(ConQty).PadRight(Rounddec - 1, "0"c) + "1"
                        If DecVal <= Rounddec Then objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = PlanQty : Continue For
                        If Rounddec - 1 = 0 Then
                            objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = 0.1
                        Else
                            objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = "0." + CStr(Math.Floor(0.0)).PadRight(Rounddec - 1, "0"c) + "1" 'CStr(Math.Floor(PlanQty)) + "." + CStr(Math.Floor(PlanQty)).PadRight(Rounddec - 1, "0"c) + "1"
                        End If
                    Else
                        objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = Math.Round(PlanQty, Rounddec)
                    End If

                Next
                objmatrix.Columns.Item("4").Cells.Item(1).Click()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Function

    End Class
End Namespace
