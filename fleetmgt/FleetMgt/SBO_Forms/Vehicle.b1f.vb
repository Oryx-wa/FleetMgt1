Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace FleetMgt
    <FormAttribute("FleetMgt.Vehicle", "SBO_Forms/Vehicle.b1f")>
    Friend Class Vehicle
        Inherits UserFormBaseClassOld

        Dim oFleetRec As SAPbobsCOM.Recordset

        Public Sub New()

            LoadFilters()
            GetData()

            oFleetRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        End Sub

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

                'filter for Driver
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add

                With oCondition
                    .Alias = "u_fltEmpType"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "Driver"
                End With

                cflDriver = oForm.ChooseFromLists.Item("CFLDriver")
                cflDriver.SetConditions(oConditions)

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


        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            oForm = SBO_Application.Forms.ActiveForm

            If oForm.TypeEx = "FleetMgt.Vehicle" Then
                If (pVal.BeforeAction) Then
                    HandleNavigation()
                End If
            End If

     
        End Sub

        Private Sub btnAddRow_ClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles btnAddNew.ClickAfter
            Try
                oForm.DataSources.DBDataSources.Item(1).Clear()
                MatTire.AddRow(1)
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

                fillCombo("Name", "Remarks", "OUDP", cboDept, , True)
                fillCombo("PrcCode", "PrcName", "OPRC", cboCostCentre, "DimCode='" & oDataSource.GetValue("U_DimCode", oDataSource.Offset).ToString & "'", True)
                fillCombo("Code", "Name", "@OWA_FMLOCATION", cboLoca, , True)
                'fillCombo("Code", "Name", "@OWA_FMFLEETTYPE", cboFType, "U_Category='V'", True)

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
                     " WHERE T0.ProfitCode = '" + cboCostCentre.Value.Trim & "' and T1.U_Category='V'"

                'oForm.DataSources.DataTables.Item("DTJournal").ExecuteQuery(sQuery)

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

        Private WithEvents btnDrvLkp As SAPbouiCOM.Button
        Private WithEvents txtVIN As SAPbouiCOM.EditText
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents txtCode As SAPbouiCOM.EditText
        Private WithEvents txtModel As SAPbouiCOM.EditText
        Private WithEvents cboDept As SAPbouiCOM.ComboBox
        Private WithEvents cflDriver As SAPbouiCOM.ChooseFromList
        Private WithEvents cflAsset As SAPbouiCOM.ChooseFromList
        Private WithEvents cflService As SAPbouiCOM.ChooseFromList
        Private WithEvents cflPrj As SAPbouiCOM.ChooseFromList
        Private WithEvents MatTire As SAPbouiCOM.Matrix
        Private WithEvents matMtcLine As SAPbouiCOM.Matrix
        Private WithEvents btnAddNew As SAPbouiCOM.Button
        Private WithEvents cflVendor As SAPbouiCOM.ChooseFromList
        Private WithEvents cboVendor As SAPbouiCOM.Column
        Private WithEvents oGrid As SAPbouiCOM.Grid
        Private WithEvents cboLoca As SAPbouiCOM.ComboBox
        Private WithEvents cboFType As SAPbouiCOM.ComboBox
        Private WithEvents txtSvrCode As SAPbouiCOM.Column
        Private WithEvents colLastMtceDate As SAPbouiCOM.Column
        Private WithEvents colNextMtceDate As SAPbouiCOM.Column
        Private WithEvents colMtceInterval As SAPbouiCOM.Column
        Private WithEvents txtMtcSch As SAPbouiCOM.EditText
        Private WithEvents btnMtcSch As SAPbouiCOM.Button
        Private WithEvents oEditText As SAPbouiCOM.EditText
        Private WithEvents btnPrj As SAPbouiCOM.Button
        Private WithEvents txtAsset As SAPbouiCOM.EditText
        Private WithEvents btnAsset As SAPbouiCOM.Button
        Private WithEvents cboCostCentre As SAPbouiCOM.ComboBox

        Private Sub Vehicle_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            'FilterRecordSet()
        End Sub

        Private Sub Vehicle_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
                'oDataSource.SetValue("Name", oDataSource.Offset, txtModel.Value)
                oDataSource.SetValue("U_Category", oDataSource.Offset, "V")
                'oForm.Update()
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub


        Private Sub Vehicle_DataUpdateAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
                ' oDataSource.SetValue("Name", oDataSource.Offset, txtModel.Value)
                oDataSource.SetValue("U_Category", oDataSource.Offset, "V")

                oForm.Update()

                'FilterRecordSet()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub




        Private Sub Driver_ChooseFromListAfter(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtDrv.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            oDataSource.SetValue("U_DriverID", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OHEM")
            If getOffset(val, "empID", oDataSource) Then
                oDataSource.Offset = 0
                getOffset(val, "empID", oDataSource)
                oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
                oDataSource2.SetValue("U_DriverName", 0, Trim(oDataSource.GetValue("lastName", 0)) & ", " & Trim(oDataSource.GetValue("firstName", 0)))
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
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


        Private Sub cboVendor_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles cboVendor.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMTIRETRACK")
            MatTire.FlushToDataSource()
            oDataSource.SetValue("U_VendorCode", oDataSource.Offset, val)
            MatTire.LoadFromDataSource()

        End Sub

        Private Sub MtcSch_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtMtcSch.ChooseFromListAfter
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

                'get the Vehicl ID
                oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_FMMTCESHEDHDR")
                If Not String.IsNullOrEmpty(oDataSource2.GetValue("Code", 0)) Then
                    oForm.Freeze(True)
                    InsertSchedMtceLines(Trim(oDataSource2.GetValue("Code", 0)), val)
                    oForm.Freeze(False)
                    oForm.Refresh()
                End If
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
        End Sub


        Private Sub txtSvrCode_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtSvrCode.ChooseFromListAfter
            Try
                Dim sQuery As String = "", val As String,
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



        Private Sub OnCustomInitialize()
            AddHandler DataAddBefore, AddressOf Me.Vehicle_DataAddBefore
            AddHandler DataAddAfter, AddressOf Me.Vehicle_DataAddAfter

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private WithEvents txtPrj As SAPbouiCOM.EditText


        Public Overrides Sub OnInitializeComponent()
            Me.txtPrj = CType(Me.GetItem("txtPrj").Specific, SAPbouiCOM.EditText)
            Me.txtAsset = CType(Me.GetItem("txtAsset").Specific, SAPbouiCOM.EditText)
            Me.txtMtcSch = CType(Me.GetItem("txtMtcSch").Specific, SAPbouiCOM.EditText)
            Me.cboCostCentre = CType(Me.GetItem("cboCC").Specific, SAPbouiCOM.ComboBox)
            Me.cboLoca = CType(Me.GetItem("cboLoca").Specific, SAPbouiCOM.ComboBox)
            Me.cboDept = CType(Me.GetItem("cboDept").Specific, SAPbouiCOM.ComboBox)
            Me.txtCode = CType(Me.GetItem("txtCode").Specific, SAPbouiCOM.EditText)
            Me.MatTire = CType(Me.GetItem("matTire").Specific, SAPbouiCOM.Matrix)
            Me.matMtcLine = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix)
            Me.txtSvrCode = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colSvrCode")
            Me.colLastMtceDate = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colLMDate")
            Me.colNextMtceDate = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colNMDate")
            Me.colMtceInterval = CType(Me.GetItem("matMtcLine").Specific, SAPbouiCOM.Matrix).Columns.Item("colIntDays")
            Me.cboVendor = CType(Me.GetItem("matTire").Specific, SAPbouiCOM.Matrix).Columns.Item("Vendor")
            Me.btnAddNew = CType(Me.GetItem("btnAdd").Specific, SAPbouiCOM.Button)
            Me.oGrid = CType(Me.GetItem("gdJornal").Specific, SAPbouiCOM.Grid)
            Me.txtDrv = CType(Me.GetItem("txtDrv").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("Item_62").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("Item_63").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("Item_64").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("Item_65").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("Item_66").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("Item_90").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub
        Private WithEvents EditText20 As SAPbouiCOM.EditText
        Private WithEvents Grid2 As SAPbouiCOM.Grid
        Private WithEvents txtDrv As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
    End Class
End Namespace
