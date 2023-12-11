Option Strict Off
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Printing
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports SAPbouiCOM.Framework

Namespace CheckPrinting_ProdOrdRoundOff
    <FormAttribute("MULCHK", "Business Objects/FrmCheckPrinting.b1f")>
    Friend Class FrmCheckPrinting
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objDTable As SAPbouiCOM.DataTable
        Dim StrQuery As String
        Private WithEvents objCheck As SAPbouiCOM.CheckBox
        Dim objRs As SAPbobsCOM.Recordset
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("Mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("lfdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tfdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("ltdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("ttdate").Specific, SAPbouiCOM.EditText)
            Me.Button2 = CType(Me.GetItem("btnload").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub


        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MULCHK", 0)
                Matrix0.AutoResizeColumns()
                objform.EnableMenu("1281", False)
                objform.EnableMenu("1282", False)
            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Button2 As SAPbouiCOM.Button

#End Region

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                If objform.DataSources.DataTables.Count.Equals(0) Then
                    objform.DataSources.DataTables.Add("DT_List")
                Else
                    objform.DataSources.DataTables.Item("DT_List").Clear()
                End If
                objDTable = objform.DataSources.DataTables.Item("DT_List")
                objDTable.Clear()
                If objaddon.HANA Then
                    StrQuery = "Select ROW_NUMBER() OVER (order by T2.""DocEntry"" desc) as ""LineId"",'N' as ""Select"",T2.""DocEntry"",T1.""CheckKey"",T1.""PmntDate"",T1.""VendorName"",T1.""CheckNum"",T1.""CheckSum"",T1.""TotalWords"",T4.""Street"",T4.""City"",T4.""ZipCode"","
                    StrQuery += vbCrLf + "(Select ""Name"" from OCST Where ""Code""=T4.""State"" and ""Country""=T4.""Country"") as ""State"",(Select ""Name"" from OCRY Where ""Code""=T4.""Country"") as ""Country"""
                    StrQuery += vbCrLf + "from OCHO T1 INNER JOIN OVPM T2 ON T1.""TransRef""=T2.""DocNum"" and T1.""PmntDate""=T2.""DocDate"" "
                    StrQuery += vbCrLf + "LEFT OUTER JOIN CRD1 T4 ON T4.""CardCode""=T1.""VendorCode"" AND T4.""Address""=T1.""AddrName"" AND T2.""PayToCode""=T4.""Address"" AND T4.""AdresType""='B'"
                    StrQuery += vbCrLf + "where T1.""PmntDate"" between '" & EditText0.Value & "' and '" & EditText1.Value & "' and T2.""U_MND_CHKP""='N' "
                Else
                End If
                objform.Freeze(True)
                objDTable.ExecuteQuery(StrQuery)
                Matrix0.Clear()
                Matrix0.LoadFromDataSourceEx()
                Matrix0.AutoResizeColumns()
                If Matrix0.VisualRowCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("No records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                For i As Integer = 0 To Matrix0.Columns.Count - 1
                    Matrix0.Columns.Item(i).BackColor = Matrix0.Item.BackColor
                Next

                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Button2_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button2.ClickBefore
            Try
                If EditText0.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("From Date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText1.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("To Date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix0.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.PressedAfter
            Try
                objCheck = Matrix0.Columns.Item("sel").Cells.Item(pVal.Row).Specific
                If pVal.ColUID = "sel" Then
                    If objCheck.Checked = True Then
                        'Matrix0.SelectRow(pVal.Row, True, True)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                    Else
                        'Matrix0.SelectRow(pVal.Row, True, False)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Matrix0.Item.BackColor)
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If Matrix0.VisualRowCount = 0 Then objaddon.objapplication.StatusBar.SetText("Line Data is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                Dim Flag As Boolean = False
                For i As Integer = 1 To Matrix0.VisualRowCount
                    objCheck = Matrix0.Columns.Item("sel").Cells.Item(i).Specific
                    If objCheck.Checked = True Then
                        Flag = True
                        Exit For
                    End If
                Next
                If Flag = False Then objaddon.objapplication.StatusBar.SetText("Please select a row...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False

            Catch ex As Exception
            End Try
        End Sub

        Private Function Print_Multiple_Checks() As Boolean
            Dim cryRpt As New ReportDocument
            Try
                Dim Filename, Foldername, DBUserName, DBPassword As String, OutPayEntry As String = ""
                'Filename = "E:\Chitra\BioSource US Customer\CheckPrinting.rpt"
                Dim initialpath As String = objaddon.objglobalmethods.getSingleValue("select ""AttachPath"" from OADP")
                Foldername = initialpath + "CheckPrint\RptFile"
                If Directory.Exists(Foldername) Then
                Else
                    Directory.CreateDirectory(Foldername)
                End If
                Filename = Foldername & "\CheckPrinting.rpt"
                DBUserName = objaddon.objglobalmethods.getSingleValue("select ""U_MND_USER"" from OADM")
                DBPassword = objaddon.objglobalmethods.getSingleValue("select ""U_MND_PASS"" from OADM")
                cryRpt.Load(Filename)
                cryRpt.DataSourceConnections(0).SetConnection(objaddon.objcompany.Server, objaddon.objcompany.CompanyDB, False)
                cryRpt.DataSourceConnections(0).SetLogon(Trim(DBUserName), Trim(DBPassword))
                Try
                    cryRpt.Refresh()
                    cryRpt.VerifyDatabase()
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("VerifyDatabase: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End Try
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                For i As Integer = 1 To Matrix0.VisualRowCount
                    objCheck = Matrix0.Columns.Item("sel").Cells.Item(i).Specific
                    If objCheck.Checked = True Then
                        OutPayEntry = Matrix0.Columns.Item("outpay").Cells.Item(i).Specific.string
                        cryRpt.SetParameterValue("DocKey@", OutPayEntry)
                    Else
                        Continue For
                    End If
                    Dim defaultprinter As String = Get_DefaultPrinter()

                    cryRpt.PrintOptions.PrinterName = defaultprinter ' "EPSON4A7C9D (L6170 Series) (redirected 25)" ' strDefaultPrinter
                    cryRpt.PrintToPrinter(1, False, 0, 0)

                    If objaddon.HANA Then
                        StrQuery = "Update OVPM Set ""U_MND_CHKP""='Y' where ""DocEntry""='" & Matrix0.Columns.Item("outpay").Cells.Item(i).Specific.string & "' "
                    Else
                        StrQuery = "Update OVPM Set U_MND_CHKP='Y' where DocEntry='" & Matrix0.Columns.Item("outpay").Cells.Item(i).Specific.string & "' "
                    End If
                    objRs.DoQuery(StrQuery)
                Next
                Matrix0.Clear()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Print_Multiple_Checks: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Finally
                cryRpt.Close()
                cryRpt.Dispose()
            End Try
        End Function

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objaddon.objapplication.StatusBar.SetText("Checks Printing. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If Print_Multiple_Checks() = True Then
                    objaddon.objapplication.StatusBar.SetText("Checks Printed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Print_ClickAfter: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Function Get_DefaultPrinter() As String
            Try
                Dim defprinter As String = ""
                Dim settings As PrinterSettings
                settings = New PrinterSettings

                For Each printer In PrinterSettings.InstalledPrinters
                    settings.PrinterName = printer
                    If settings.IsDefaultPrinter Then
                        defprinter = printer
                        Exit For
                    End If
                Next
                objaddon.objapplication.StatusBar.SetText("Get_DefaultPrinter: " & defprinter, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return defprinter
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Get_DefaultPrinter: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Function
    End Class
End Namespace
