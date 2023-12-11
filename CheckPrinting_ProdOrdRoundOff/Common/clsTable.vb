Namespace CheckPrinting_ProdOrdRoundOff

    Public Class clsTable

        Public Sub FieldCreation()
            AddFields("OWOR", "MND_TANK", "Tank", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("OWOR", "MND_DCMP", "Decimal Point", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("WOR1", "MND_PLNQTY", "Original Planned Qty.", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity,, "0")
            AddFields("OVPM", "MND_CHKP", "Is Check Printed", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", , {"Y,Yes", "N,No"})
            AddFields("OADM", "MND_USER", "UserId", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("OADM", "MND_PASS", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("OWOR", "MND_BTNO", "Batch", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("OWOR", "MND_EXPD", "Expiration Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("OWOR", "MND_MFGD", "Mfg Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("@CT_PF_OMOR", "AdjDec", "Adjust Decimal", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@CT_PF_MOR3", "OQty", "Original Qty", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            AddFields("@CT_PF_OMOR", "BatchNo", "Batch No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@CT_PF_OMOR", "MfgDate", "Mfg Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@CT_PF_OMOR", "ExpDate", "Expiration Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("OITB", "ASFGMOC", "Auto SFG MO Creation", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "Y", , {"Y,Yes", "N,No"})

            Tank_Master()
        End Sub

#Region "Document Data Creation"

        Private Sub MultiBranchAPInvoice()
            AddTables("MIPL_OAPI", "AP Service Inv Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MIPL_API1", "AP Service Inv Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@MIPL_OAPI", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OAPI", "DueDate", "DocDue Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OAPI", "PosDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("@MIPL_API1", "SACCode", "SAC Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "VendorCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "VendorName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "Desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "GLCode", "GL Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_API1", "GLName", "GL Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "TaxCode", "Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddFields("@MIPL_API1", "OTaxCode", "O Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddFields("@MIPL_API1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MIPL_API1", "GTotal", "Gross Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MIPL_API1", "Branch", "Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "TranEntry", "Transaction Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "RefNo", "Header Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "LRefNo", "Line Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("OPCH", "ymhbpref", "YMH Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("PCH1", "ymhbpref", "YMH Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            AddUDO("MIAPSI", "AP Service Invoice", SAPbobsCOM.BoUDOObjType.boud_Document, "MIPL_OAPI", {"MIPL_API1"}, {"DocEntry", "DocNum", "U_PosDate", "U_DocDate", "U_DueDate"}, True, True)
        End Sub

#End Region

#Region "Master Data Creation"
        Private Sub Tank_Master()

            AddTables("MND_OTNKM", "Tank Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("MND_TNKM1", "Tank Master Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            AddFields("@MND_TNKM1", "MND_TANK", "Tank", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            AddFields("@MND_TNKM1", "MND_SIZE", "Tank Size", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MND_TNKM1", "MND_TNKSG", "Tank Size(Gallon)", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MND_TNKM1", "MND_ACT", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1)
            AddFields("@MND_TNKM1", "MND_HTEMP", "Heat Temperature", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            AddFields("@MND_TNKM1", "MND_HTIME", "Heat Time", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            AddFields("@MND_TNKM1", "MND_SSMPL", "Sterile Sample", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)

            AddUDO("MND_UDO_TNKM", "Tank Master UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MND_OTNKM", {"MND_TNKM1"}, {"Code", "Name"}, True, False)
        End Sub

#End Region

#Region "Table Creation Common Functions"

        Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            Try
                oUserTablesMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                'Adding Table
                If Not oUserTablesMD.GetByKey(strTab) Then
                    oUserTablesMD.TableName = strTab
                    oUserTablesMD.TableDescription = strDesc
                    oUserTablesMD.TableType = nType

                    If oUserTablesMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription & strTab)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Sub AddFields(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoFieldTypes,
                             Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO,
                              Optional ByVal defaultvalue As String = "", Optional ByVal Yesno As Boolean = False, Optional ByVal Validvalues() As String = Nothing, Optional ByVal LinkedSystemObject As SAPbobsCOM.UDFLinkedSystemObjectTypesEnum = Nothing)
            Dim oUserFieldMD1 As SAPbobsCOM.UserFieldsMD
            oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                'If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                '    strTab = "@" + strTab
                'End If
                If Not IsColumnExists(strTab, strCol) Then
                    'If Not oUserFieldMD1 Is Nothing Then
                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    'End If
                    'oUserFieldMD1 = Nothing
                    'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc
                    oUserFieldMD1.Name = strCol
                    oUserFieldMD1.Type = nType
                    oUserFieldMD1.SubType = nSubType
                    oUserFieldMD1.TableName = strTab
                    oUserFieldMD1.EditSize = nEditSize
                    oUserFieldMD1.Mandatory = Mandatory
                    oUserFieldMD1.DefaultValue = defaultvalue

                    If Yesno = True Then
                        oUserFieldMD1.ValidValues.Value = "Y"
                        oUserFieldMD1.ValidValues.Description = "Yes"
                        oUserFieldMD1.ValidValues.Add()
                        oUserFieldMD1.ValidValues.Value = "N"
                        oUserFieldMD1.ValidValues.Description = "No"
                        oUserFieldMD1.ValidValues.Add()
                    End If
                    If LinkedSystemObject <> 0 Then oUserFieldMD1.LinkedSystemObject = LinkedSystemObject

                    Dim split_char() As String
                    If Not Validvalues Is Nothing Then
                        If Validvalues.Length > 0 Then
                            For i = 0 To Validvalues.Length - 1
                                If Trim(Validvalues(i)) = "" Then Continue For
                                split_char = Validvalues(i).Split(",")
                                If split_char.Length <> 2 Then Continue For
                                oUserFieldMD1.ValidValues.Value = split_char(0)
                                oUserFieldMD1.ValidValues.Description = split_char(1)
                                oUserFieldMD1.ValidValues.Add()
                            Next
                        End If
                    End If
                    Dim val As Integer
                    val = oUserFieldMD1.Add
                    If val <> 0 Then
                        objaddon.objapplication.SetStatusBarMessage(objaddon.objcompany.GetLastErrorDescription & " " & strTab & " " & strCol, True)
                    End If
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                End If
            Catch ex As Exception
                Throw ex
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                oUserFieldMD1 = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strSQL As String
            Try
                If objaddon.HANA Then
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
                Else
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
                End If

                oRecordSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strSQL)

                If oRecordSet.Fields.Item(0).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Function

        Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
            Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

            Try
                '// The meta-data object must be initialized with a
                '// regular UserKeys object
                oUserKeysMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

                If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                    '// Set the table name and the key name
                    oUserKeysMD.TableName = strTab
                    oUserKeysMD.KeyName = strKey

                    '// Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn
                    oUserKeysMD.Elements.Add()
                    oUserKeysMD.Elements.ColumnAlias = "RentFac"

                    '// Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                    '// Add the key
                    If oUserKeysMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
                oUserKeysMD = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub AddUDO(ByVal strUDO As String, ByVal strUDODesc As String, ByVal nObjectType As SAPbobsCOM.BoUDOObjType, ByVal strTable As String, ByVal childTable() As String, ByVal sFind() As String,
                           Optional ByVal canlog As Boolean = False, Optional ByVal Manageseries As Boolean = False)

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim tablecount As Integer = 0
            Try
                oUserObjectMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                If oUserObjectMD.GetByKey(strUDO) = 0 Then

                    oUserObjectMD.Code = strUDO
                    oUserObjectMD.Name = strUDODesc
                    oUserObjectMD.ObjectType = nObjectType
                    oUserObjectMD.TableName = strTable

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES

                    If Manageseries Then oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES Else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    If canlog Then
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.LogTableName = "A" + strTable.ToString
                    Else
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                        oUserObjectMD.LogTableName = ""
                    End If

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.ExtensionName = ""

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    tablecount = 1
                    If sFind.Length > 0 Then
                        For i = 0 To sFind.Length - 1
                            If Trim(sFind(i)) = "" Then Continue For
                            oUserObjectMD.FindColumns.ColumnAlias = sFind(i)
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount)
                            tablecount = tablecount + 1
                        Next
                    End If

                    tablecount = 0
                    If Not childTable Is Nothing Then
                        If childTable.Length > 0 Then
                            For i = 0 To childTable.Length - 1
                                If Trim(childTable(i)) = "" Then Continue For
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount)
                                oUserObjectMD.ChildTables.TableName = childTable(i)
                                oUserObjectMD.ChildTables.Add()
                                tablecount = tablecount + 1
                            Next
                        End If
                    End If

                    If oUserObjectMD.Add() <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If

            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub

#End Region

    End Class
End Namespace

