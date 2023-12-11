Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace CheckPrinting_ProdOrdRoundOff

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx

                    Case "141", "170", "-170", "426", "-426", "392", "-392"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)

                    Case "65211"
                        Production_Order_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim oUDFForm As SAPbouiCOM.Form
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "6005"

                        Case "6913"

                        Case "1284" 'Cancel

                    End Select
                Else
                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Select Case pval.MenuUID
                        Case "1284" 'Cancel

                        Case "1281" 'Find

                        Case "1287" 'Duplicate



                        Case "1282"


                        Case Else

                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Production Order"

        Private Sub Production_Order_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                'Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    'Dim DBSource As SAPbouiCOM.DBDataSource
                    'DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OAPI")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("btnDCP").Enabled = False

                        Case "1282" ' Add Mode
                            objform.Items.Item("btnDCP").Enabled = True


                        Case "1288", "1289", "1290", "1291"

                        Case "1293"
                            'DeleteRow(Matrix0, "@MIPL_API1")
                        Case "1292"
                            'objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region



        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub

    End Class
End Namespace