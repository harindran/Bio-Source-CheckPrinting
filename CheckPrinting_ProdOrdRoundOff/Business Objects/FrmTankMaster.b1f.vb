Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace CheckPrinting_ProdOrdRoundOff
    <FormAttribute("TNKMSTR", "Business Objects/FrmTankMaster.b1f")>
    Friend Class FrmTankMaster
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Dim RecCount As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("lblcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtcode").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
#End Region

        Private Sub OnCustomInitialize()
            Try

                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "tank", "#")
                If objaddon.HANA Then
                    objform.Items.Item("txtcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("@MND_OTNKM")
                Else
                    objform.Items.Item("txtcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("[@MND_OTNKM]")
                End If
                If objaddon.HANA Then
                    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) from ""@MND_OTNKM"";")
                Else
                    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) from [@MND_OTNKM]")
                End If
                If RecCount = "1" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText0.Item.Enabled = True
                    EditText0.Value = "1"
                    'objform.ActiveItem = "txtdoc"
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    EditText0.Item.Enabled = False
                    'Exit Sub
                End If
                Matrix0.AutoResizeColumns()
            Catch ex As Exception

            End Try
        End Sub


        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("TNKMSTR", pVal.FormTypeCount)
            Catch ex As Exception

            End Try

        End Sub


        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    If pVal.ActionSuccess Then
                        If objaddon.HANA Then
                            RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) +1 from ""@MND_OTNKM"";")
                        Else
                            RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) +1 from [@MND_OTNKM]")
                        End If
                        If RecCount <> "2" Then
                            objform.Close()
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try


        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "tank", "#")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                RemoveLastrow(Matrix0, "tank")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

    End Class
End Namespace
