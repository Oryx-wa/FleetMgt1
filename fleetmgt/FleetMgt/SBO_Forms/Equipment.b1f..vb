Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace FleetMgt
    <FormAttribute("FleetMgt.Equipment", "SBO_Forms/Equipment.b1f")>
    Friend Class Equipment
        Inherits UserFormBaseClassold

        Dim oNavInicio As Boolean, oFleetRec As SAPbobsCOM.Recordset

        Public Sub New()

            LoadFilters()
            GetData()

            oFleetRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'FilterRecordSet()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.txtModel = CType(Me.GetItem("txtModel").Specific, SAPbouiCOM.EditText)
            Me.txtCode = CType(Me.GetItem("txtCode").Specific, SAPbouiCOM.EditText)
            Me.cboCostCentre = CType(Me.GetItem("cboCC").Specific, SAPbouiCOM.ComboBox)
            Me.oGrid = CType(Me.GetItem("gdJornal").Specific, SAPbouiCOM.Grid)
            Me.txtMtcSch = CType(Me.GetItem("txtMtcSch").Specific, SAPbouiCOM.EditText)
            Me.cboLoca = CType(Me.GetItem("cboLoca").Specific, SAPbouiCOM.ComboBox)
            Me.cboFType = CType(Me.GetItem("cboFType").Specific, SAPbouiCOM.ComboBox)
            Me.matMtcLine = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix)
            Me.txtSvrCode = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colSvrCode")
            Me.colLastMtceDate = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colLMDate")
            Me.colNextMtceDate = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colNMDate")
            Me.colMtceInterval = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colIntDays")
            Me.txtPrj = CType(Me.GetItem("txtPrj").Specific, SAPbouiCOM.EditText)
            Me.txtAsset = CType(Me.GetItem("txtAsset").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("Item_26").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("Item_27").Specific, SAPbouiCOM.EditText)
            Me.EditText0 = CType(Me.GetItem("Item_90").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("Item_89").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("Item_88").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("Item_87").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("Item_86").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("Item_85").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("Item_84").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("Item_83").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("Item_82").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("Item_81").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("Item_80").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText6 = CType(Me.GetItem("Item_67").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("Item_102").Specific, SAPbouiCOM.EditText)
            Me.EditText7 = CType(Me.GetItem("Item_101").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("Item_100").Specific, SAPbouiCOM.StaticText)
            Me.EditText8 = CType(Me.GetItem("Item_99").Specific, SAPbouiCOM.EditText)
            Me.StaticText8 = CType(Me.GetItem("Item_98").Specific, SAPbouiCOM.StaticText)
            Me.EditText9 = CType(Me.GetItem("Item_97").Specific, SAPbouiCOM.EditText)
            Me.StaticText9 = CType(Me.GetItem("Item_96").Specific, SAPbouiCOM.StaticText)
            Me.StaticText10 = CType(Me.GetItem("Item_95").Specific, SAPbouiCOM.StaticText)
            Me.EditText10 = CType(Me.GetItem("Item_94").Specific, SAPbouiCOM.EditText)
            Me.StaticText11 = CType(Me.GetItem("Item_93").Specific, SAPbouiCOM.StaticText)
            Me.EditText11 = CType(Me.GetItem("Item_92").Specific, SAPbouiCOM.EditText)
            Me.StaticText12 = CType(Me.GetItem("Item_91").Specific, SAPbouiCOM.StaticText)
            Me.Grid0 = CType(Me.GetItem("grdWOrder").Specific, SAPbouiCOM.Grid)
            Me.Button1 = CType(Me.GetItem("Item_14").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub


        'Private Sub FilterRecordSet()
        '    Try
        '        If Not IsNothing(oFleetRec) Then
        '            oNavInicio = False
        '            oFleetRec.DoQuery(String.Format("select * from [@OWA_FMFLEET] where U_Category = 'E' order by docentry"))
        '            'LoadByDocEntry(oFleetRec.Fields.Item("Code").Value)
        '        End If

        '    Catch ex As Exception
        '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    End Try
        'End Sub

        'Private Sub LoadByDocEntry(ByVal code As String)
        '    '
        '    If oForm.TypeEx = "FleetMgt.Equipment" Then
        '        oConditions = New SAPbouiCOM.Conditions
        '        oCondition = Nothing
        '        oCondition = oConditions.Add
        '        With oCondition
        '            .Alias = "Code"
        '            .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        '            .CondVal = Convert.ToString(code).Trim
        '        End With
        '         oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET").Query(oConditions)
        '        oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET").Offset = 0

        '    End If

        'End Sub

        Private Sub LoadFilters()
            Try
                'filter for Vendor
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add
                With oCondition
                    .Alias = "CardType"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "S"
                End With
                cflVendor = oForm.ChooseFromLists.Item("CFLVendor")
                cflVendor.SetConditions(oConditions)

                'filter for Asset
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add

                With oCondition
                    .Alias = "AssetItem"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "Y"
                End With

                cflAsset = oForm.ChooseFromLists.Item("CFLAsset")
                cflAsset.SetConditions(oConditions)

                'filter for services
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add

                With oCondition
                    .Alias = "U_PM"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "Y"
                End With

                cflService = oForm.ChooseFromLists.Item("CFLService")
                cflService.SetConditions(oConditions)

                oForm.Select()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Finally
            End Try
        End Sub

        Private Sub GetData()
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMSETUP")
                oConditions = New SAPbouiCOM.Conditions
                oCondition = oConditions.Add
                oCondition.Alias = "Code"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = "FLEET"
                oDataSource.Query(oConditions)

                If oDataSource.Size > 0 Then
                    oDataSource.Offset = 0
                End If

                fillCombo("PrcCode", "PrcName", "OPRC", cboCostCentre, "DimCode='" & oDataSource.GetValue("U_DimCode", oDataSource.Offset).ToString & "'", True)
                fillCombo("Code", "Name", "@OWA_FMLOCATION", cboLoca, , True)
                fillCombo("Code", "Name", "@OWA_FMFLEETTYPE", cboFType, "U_Category='E'", True)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub HandleNavigation()
            Try

                txtCode = oForm.Items.Item("txtCode").Specific()

                getPostedJournals()
                DisplayWorkOrderList(txtCode.Value)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub DisplayWorkOrderList(ByVal EquipCode As String)
            Try
                Dim sQuery As String

                sQuery = "select docentry [S/N],U_FleetID [Equip COde],U_DocumentDate [Date Issued],U_PostingDate [Completion Date], "
                sQuery = sQuery + "U_TotalWOAmt [Total Cost] ,U_Status Status from [@owa_fmworkorder] where U_FleetID='" & EquipCode & "' and U_Status='Closed'"

                oGrid = oForm.Items.Item("grdWOrder").Specific
                oDataTable = oForm.DataSources.DataTables.Item("DTWOrder")
                oDataTable.ExecuteQuery(sQuery)
                oGrid.DataTable = oDataTable

                oGrid.Columns.Item(0).TitleObject.Caption = "S/N"
                oGrid.Columns.Item(1).TitleObject.Caption = "Equip Code"
                oGrid.Columns.Item(2).TitleObject.Caption = "Date Issued"
                oGrid.Columns.Item(3).TitleObject.Caption = "Completion Date"
                oGrid.Columns.Item(4).TitleObject.Caption = "Total Cost"
                oGrid.Columns.Item(5).TitleObject.Caption = "Status"


                'oGrid.AutoResizeColumns()


            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub
        Private Sub getPostedJournals()
            Try
                Dim sQuery As String
                cboCostCentre = oForm.Items.Item("cboCC").Specific

                sQuery = "Select IsNull(T0.TransType ,'') TransType , T0.RefDate TransDate, T0.LineMemo Descr,T0.Debit DebitAmt,T0.Credit CreditAmt From "
                sQuery += "[JDT1] T0 Join [@OWA_FMFLEET] T1 On T0.ProfitCode=T1.U_CostCentre " + _
                     " WHERE T0.ProfitCode = '" + cboCostCentre.Value.Trim & "' and T1.U_Category='E'"

                oGrid = oForm.Items.Item("gdJornal").Specific
                oDataTable = oForm.DataSources.DataTables.Item("DTJournal")
                oDataTable.ExecuteQuery(sQuery)
                oGrid.DataTable = oDataTable

                oGrid.Columns.Item(0).TitleObject.Caption = "Trans Type"
                oGrid.Columns.Item(1).TitleObject.Caption = "Trans Date"
                oGrid.Columns.Item(2).TitleObject.Caption = "Description"
                oGrid.Columns.Item(3).TitleObject.Caption = "Debit Amount"
                oGrid.Columns.Item(4).TitleObject.Caption = "Credit Amount"

                oGrid.AutoResizeColumns()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private WithEvents oEditText As SAPbouiCOM.EditText
        Private WithEvents txtCode As SAPbouiCOM.EditText
        Private WithEvents txtModel As SAPbouiCOM.EditText
        Private WithEvents cflDriver As SAPbouiCOM.ChooseFromList
        Private WithEvents cflVendor As SAPbouiCOM.ChooseFromList
        Private WithEvents cboCostCentre As SAPbouiCOM.ComboBox
        Private WithEvents cflAsset As SAPbouiCOM.ChooseFromList
        Private WithEvents cflService As SAPbouiCOM.ChooseFromList
        Private WithEvents cflPrj As SAPbouiCOM.ChooseFromList
        Private WithEvents oGrid As SAPbouiCOM.Grid
        Private WithEvents txtMtcSch As SAPbouiCOM.EditText
        Private WithEvents btnMtcSch As SAPbouiCOM.Button
        Private WithEvents cboLoca As SAPbouiCOM.ComboBox
        Private WithEvents matMtcLine As SAPbouiCOM.Matrix
        Private WithEvents txtSvrCode As SAPbouiCOM.Column
        Private WithEvents colLastMtceDate As SAPbouiCOM.Column
        Private WithEvents colNextMtceDate As SAPbouiCOM.Column
        Private WithEvents colMtceInterval As SAPbouiCOM.Column
        Private WithEvents cboFType As SAPbouiCOM.ComboBox

        Private Sub Equipment_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataAddAfter
            'FilterRecordSet()
        End Sub


        Private Sub Equipment_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Me.DataAddBefore
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
                'oDataSource.SetValue("Name", oDataSource.Offset, txtModel.Value)
                oDataSource.SetValue("U_Category", oDataSource.Offset, "E")

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub Equipment_DataLoadBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Me.DataLoadBefore

        End Sub


        Private Sub Equipment_DataUpdateAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataUpdateAfter
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
                'oDataSource.SetValue("Name", oDataSource.Offset, txtModel.Value)
                oDataSource.SetValue("U_Category", oDataSource.Offset, "E")

                'oForm.Update()

                'FilterRecordSet()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub


        Private Sub txtPrj_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtPrj.ChooseFromListAfter

            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            oDataSource.SetValue("U_Project", 0, val)

            oDataSource = oForm.DataSources.DBDataSources.Item("OPRJ")
            If getOffset(val, "PrjCode", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub txtAsset_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtAsset.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            oDataSource.SetValue("u_Asset", 0, val)

            oDataSource = oForm.DataSources.DBDataSources.Item("OITM")
            If getOffset(val, "ItemCode", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub txtMtcSch_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtMtcSch.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            oDataSource.SetValue("U_PrevMtceSch", 0, val)

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMMTCESHEDHDR")
            If getOffset(val, "Code", oDataSource) Then
                oDataSource.Offset = 0
            End If

            'update the mtce line
            'get the Vehicl ID
            oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            If Not String.IsNullOrEmpty(oDataSource2.GetValue("Code", 0)) Then
                oForm.Freeze(True)
                InsertSchedMtceLines(Trim(oDataSource2.GetValue("Code", 0)), val)
                oForm.Freeze(False)

            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub InsertSchedMtceLines(ByVal FleetID As String, ByVal ShedMtceCode As String)
            Dim strSQL As String
            oDataTable = Nothing
            strSQL = Me.GetSQLString("owa_insertschdmtcelines", FleetID, ShedMtceCode)
            oDataTable = oForm.DataSources.DataTables.Item("DT_Base")
            oDataTable.ExecuteQuery(strSQL)

            'strSQL = " EXEC owa_fminsertschdmtcelines '" + FleetID + "','" & ShedMtceCode & "'"
            'oRecordSet.DoQuery(strSQL)

        End Sub


        Private Sub OnCustomInitialize()

        End Sub


        Private Sub txtSvrCode_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtSvrCode.ChooseFromListAfter
            Try
                Dim sQuery As String = "", val As String

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEETMTCLINES")

                matMtcLine = oForm.Items.Item("matMtcLine").Specific
                txtSvrCode = matMtcLine.Columns.Item("colSvrCode")
                oEditText = txtSvrCode.Cells.Item(pVal.Row).Specific()


                If oEditText.Value = String.Empty Then
                    oCFLEventArg = pVal

                    oDataTable = oCFLEventArg.SelectedObjects
                    If oDataTable Is Nothing Then
                        Exit Sub
                    End If

                    val = oDataTable.GetValue(0, 0)

                    sQuery = "Select * From [@OWA_FMFLEETMTCLINES] "
                    sQuery &= "Where U_ServiceCode='" & val.Trim & "' "
                    sQuery &= "And Code='" & DBDS("@OWA_FMFLEET").GetValue("Code", 0).Trim & "'"

                    oFleetRec.DoQuery(sQuery)
                    If oFleetRec.RecordCount > 0 Then
                        SBO_Application.StatusBar.SetText("Service Code already exists for this Maintenance Line", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If


                    'oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEETMTCLINES")
                    matMtcLine.FlushToDataSource()
                    oDataSource.SetValue("U_ServiceCode", pVal.Row - 1, val)

                    getOffset(val, "Code", DBDS("@OWA_FMSERVICES"))
                    oDataSource.SetValue("U_SvrDesc", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("Name", 0).Trim)
                    oDataSource.SetValue("U_LastMtceDate", pVal.Row - 1, sboDate(Today.Date))
                    oDataSource.SetValue("U_MtceFreq", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("U_MtceIntervalDays", 0).Trim)
                    oDataSource.SetValue("U_NextMtceDate", pVal.Row - 1, sboDate(DateAdd("d", CInt(DBDS("@OWA_FMSERVICES").GetValue("U_MtceIntervalDays", 0)), Today.Date)))

                    matMtcLine.LoadFromDataSource()
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                Else

                    SBO_Application.SetStatusBarMessage("You cannot change the Maintenance Service Code again! Delete this line and add another.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End If

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub



        Private Sub colLastMtceDate_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles colLastMtceDate.ValidateBefore

            Dim a, b, c As SAPbouiCOM.EditText

            Try
                matMtcLine = oForm.Items.Item("matMtcLine").Specific
                colLastMtceDate = matMtcLine.Columns.Item("colLMDate")
                colNextMtceDate = matMtcLine.Columns.Item("colNMDate")
                colMtceInterval = matMtcLine.Columns.Item("colIntDays")

                a = colLastMtceDate.Cells.Item(pVal.Row).Specific()
                b = colNextMtceDate.Cells.Item(pVal.Row).Specific()
                c = colMtceInterval.Cells.Item(pVal.Row).Specific()

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEETMTCLINES")
                matMtcLine.FlushToDataSource()
                oDataSource.SetValue("U_NextMtceDate", pVal.Row - 1, sboDate(DateAdd("d", CInt(c.Value), sboDate(a.Value.ToString))))
                matMtcLine.LoadFromDataSource()
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub


        Private Sub colMtceInterval_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles colMtceInterval.ValidateBefore
            Dim a, b, c As SAPbouiCOM.EditText

            Try
                matMtcLine = oForm.Items.Item("matMtcLine").Specific
                colLastMtceDate = matMtcLine.Columns.Item("colLMDate")
                colNextMtceDate = matMtcLine.Columns.Item("colNMDate")
                colMtceInterval = matMtcLine.Columns.Item("colIntDays")

                a = colLastMtceDate.Cells.Item(pVal.Row).Specific()
                b = colNextMtceDate.Cells.Item(pVal.Row).Specific()
                c = colMtceInterval.Cells.Item(pVal.Row).Specific()

                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEETMTCLINES")
                matMtcLine.FlushToDataSource()
                oDataSource.SetValue("U_NextMtceDate", pVal.Row - 1, sboDate(DateAdd("d", CInt(c.Value), sboDate(a.Value.ToString))))
                matMtcLine.LoadFromDataSource()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub matMtcLine_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles matMtcLine.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_FMFLEETMTCLINES").Clear()
            matMtcLine = oForm.Items.Item("matMtcLine").Specific

            If pVal.Row = matMtcLine.RowCount + 1 Then
                If pVal.Row = 1 Then
                    matMtcLine.AddRow(1)
                Else
                    matMtcLine.AddRow(1, matMtcLine.RowCount)
                End If
                matMtcLine.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub
      

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            oForm = SBO_Application.Forms.ActiveForm

            If oForm.TypeEx = "FleetMgt.Equipment" Then
                If (pVal.BeforeAction) Then
                    HandleNavigation()
                End If
            End If


        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter
            AddHandler DataAddAfter, AddressOf Me.Equipment_DataAddAfter
            AddHandler DataAddBefore, AddressOf Me.Equipment_DataAddBefore
            AddHandler DataLoadBefore, AddressOf Me.Equipment_DataLoadBefore
            AddHandler DataUpdateAfter, AddressOf Me.Equipment_DataUpdateAfter

        End Sub
        Private WithEvents txtPrj As SAPbouiCOM.EditText
        Private WithEvents txtAsset As SAPbouiCOM.EditText
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText8 As SAPbouiCOM.EditText
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText
        Private WithEvents EditText9 As SAPbouiCOM.EditText
        Private WithEvents StaticText9 As SAPbouiCOM.StaticText
        Private WithEvents StaticText10 As SAPbouiCOM.StaticText
        Private WithEvents EditText10 As SAPbouiCOM.EditText
        Private WithEvents StaticText11 As SAPbouiCOM.StaticText
        Private WithEvents EditText11 As SAPbouiCOM.EditText
        Private WithEvents StaticText12 As SAPbouiCOM.StaticText
        Private WithEvents Grid0 As SAPbouiCOM.Grid
        Private WithEvents Button1 As SAPbouiCOM.Button



        Private Sub Button1_ClickAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            'Throw New System.NotImplementedException()
            txtMtcSch = oForm.Items.Item("txtMtcSch").Specific
            txtCode = oForm.Items.Item("txtCode").Specific
            matMtcLine = oForm.Items.Item("matMtcLine").Specific

            If Not String.IsNullOrEmpty(txtCode.Value) And Not String.IsNullOrEmpty(txtMtcSch.Value) Then
                InsertSchedMtceLines(txtCode.Value, txtMtcSch.Value)
                'FilterRecordSet()
            End If

        End Sub
    End Class
End Namespace
