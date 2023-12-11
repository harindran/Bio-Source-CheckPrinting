Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace CheckPrinting_ProdOrdRoundOff
    <FormAttribute("SELFRM", "Business Objects/FrmSelectionForm.b1f")>
    Friend Class FrmSelectionForm
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim DecRound As Integer
        Dim PlanQty As Double
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("griddata").Specific, SAPbouiCOM.Grid)
            Me.StaticText0 = CType(Me.GetItem("lblfind").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtfind").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
#End Region

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("SELFRM", 0)
                LoadGrid(TankQuery)
                objform.Update()
                objform.Refresh()
            Catch ex As Exception

            End Try
        End Sub



        Private Sub LoadGrid(ByVal query As String)
            Try
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(query)
                Grid0.DataTable.ExecuteQuery(query)
                If objRs.RecordCount = 0 Then objaddon.objapplication.StatusBar.SetText("No data found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objform.Close() : Exit Sub
                Grid0.Columns.Item(0).Editable = False
                Grid0.Columns.Item(1).Editable = False
                Grid0.RowHeaders.TitleObject.Caption = "#"
                For i As Integer = 0 To Grid0.Rows.Count - 1
                    Grid0.RowHeaders.SetText(i, i + 1)
                Next
                TankQuery = ""
            Catch ex As Exception

            End Try
        End Sub


        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim Code As String = "", Size As String = ""

                For i As Integer = 0 To Grid0.Rows.Count - 1
                    If Grid0.Rows.IsSelected(i) = True Then
                        Code = Grid0.DataTable.GetValue(0, i).ToString
                        Size = Grid0.DataTable.GetValue(1, i).ToString
                    End If
                Next
                If FrmSelForm.Items.Item("6").Specific.string = "" Then Exit Sub
                FrmSelForm.Items.Item("txtTANK").Specific.string = Code
                FrmSelForm.Items.Item("12").Specific.string = Size

                If FrmSelForm.Items.Item("etDCP").Specific.string = "" Then
                    DecRound = 0
                Else
                    DecRound = FrmSelForm.Items.Item("etDCP").Specific.string
                End If
                If DecRound <> 0 Then
                    objmatrix = FrmSelForm.Items.Item("37").Specific
                    For i As Integer = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("4").Cells.Item(i).Specific.String = "" Then Continue For
                        PlanQty = CDbl(objmatrix.Columns.Item("14").Cells.Item(i).Specific.String)
                        'objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = Math.Round(PlanQty, DecRound)
                        If Math.Round(PlanQty, DecRound) = 0 Then
                            Dim DecVal As Integer = PlanQty.ToString().Substring(PlanQty.ToString().IndexOf(".")).Length - 1
                            If DecVal <= DecRound Then objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = PlanQty : Continue For
                            If DecRound - 1 = 0 Then
                                objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = 0.1
                            Else
                                objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = "0." + CStr(Math.Floor(0.0)).PadRight(DecRound - 1, "0"c) + "1"
                            End If
                        Else
                            objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = Math.Round(PlanQty, DecRound)
                        End If
                    Next
                    objmatrix.Columns.Item("4").Cells.Item(1).Click()
                End If

                FrmSelForm = Nothing
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.KeyDownAfter
            Try
                Dim FindString As String
                FindString = EditText0.Value
                For i As Integer = 0 To Grid0.Rows.Count - 1
                    If Grid0.DataTable.GetValue(0, i).ToString.ToUpper Like FindString.ToUpper Then
                        Grid0.Rows.SelectedRows.Add(i)
                        Exit For
                    End If
                Next
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.ClickAfter
            Try
                Grid0.Rows.SelectedRows.Add(pVal.Row)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Grid0_DoubleClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.DoubleClickAfter
            Try
                If pVal.Row = -1 Then Exit Sub
                Dim Code As String = "", Size As String = ""
                For i As Integer = 0 To Grid0.Rows.Count - 1
                    If Grid0.Rows.IsSelected(i) = True Then
                        Code = Grid0.DataTable.GetValue(0, i).ToString
                        Size = Grid0.DataTable.GetValue(1, i).ToString
                    End If
                Next

                If FrmSelForm.Items.Item("6").Specific.string = "" Then Exit Sub
                FrmSelForm.Items.Item("txtTANK").Specific.string = Code
                FrmSelForm.Items.Item("12").Specific.string = Size
                If FrmSelForm.Items.Item("etDCP").Specific.string = "" Then
                    DecRound = 0
                Else
                    DecRound = FrmSelForm.Items.Item("etDCP").Specific.string
                End If
                If DecRound <> 0 Then
                    objmatrix = FrmSelForm.Items.Item("37").Specific
                    For i As Integer = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("4").Cells.Item(i).Specific.String = "" Then Continue For
                        PlanQty = CDbl(objmatrix.Columns.Item("14").Cells.Item(i).Specific.String)
                        'objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = Math.Round(PlanQty, DecRound)
                        If Math.Round(PlanQty, DecRound) = 0 Then
                            Dim DecVal As Integer = PlanQty.ToString().Substring(PlanQty.ToString().IndexOf(".")).Length - 1
                            If DecVal <= DecRound Then objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = PlanQty : Continue For
                            If DecRound - 1 = 0 Then
                                objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = 0.1
                            Else
                                objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = "0." + CStr(Math.Floor(0.0)).PadRight(DecRound - 1, "0"c) + "1"
                            End If
                        Else
                            objmatrix.Columns.Item("14").Cells.Item(i).Specific.String = Math.Round(PlanQty, DecRound)
                        End If
                    Next
                    objmatrix.Columns.Item("4").Cells.Item(1).Click()
                End If
                FrmSelForm = Nothing
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
