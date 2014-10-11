Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace FleetMgt
    <FormAttribute("FleetMgt.WorkOrder", "SBO_Forms/WorkOrder.b1f")>
    Friend Class WorkOrder
        Inherits UserFormBaseClassOld

        Public Sub New()
            LoadFilters()
            optVehi = oForm.Items.Item("optVehi").Specific
            optVehi.GroupWith("optEquip")

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.txtVehicle = CType(Me.GetItem("txtVehi").Specific, SAPbouiCOM.EditText)
            Me.lkVehi = CType(Me.GetItem("lkVehi").Specific, SAPbouiCOM.LinkedButton)
            Me.btnAdd = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.btnCancel = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.cboWOType = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("cboWOType")
            Me.txtWOCode = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtWOCode")
            Me.txtDesc = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtDesc")
            Me.txtVendor = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtVendor")
            Me.oLnkWOCode = CType(CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtWOCode").ExtendedObject, SAPbouiCOM.LinkedButton)
            Me.txtQty = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtQty")
            Me.txtUPrice = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtUPrice")
            Me.txtAmount = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix).Columns.Item("txtAmount")
            Me.txtToWOAmt = CType(Me.GetItem("txtToWOAmt").Specific, SAPbouiCOM.EditText)
            Me.oMatrix = CType(Me.GetItem("WODetail").Specific, SAPbouiCOM.Matrix)
            Me.btnPost = CType(Me.GetItem("btnPost").Specific, SAPbouiCOM.Button)
            Me.DocEntry = CType(Me.GetItem("txtDocNo").Specific, SAPbouiCOM.EditText)
            Me.optEquip = CType(Me.GetItem("optEquip").Specific, SAPbouiCOM.OptionBtn)
            Me.optVehi = CType(Me.GetItem("optVehi").Specific, SAPbouiCOM.OptionBtn)
            Me.EditText0 = CType(Me.GetItem("Item_2").Specific, SAPbouiCOM.EditText)
            Me.txtrptdby = CType(Me.GetItem("txtrptdby").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Private WithEvents oEditText As SAPbouiCOM.EditText
        Private WithEvents oLnkWOCode As SAPbouiCOM.LinkedButton
        Private WithEvents txtVehicle As SAPbouiCOM.EditText
        Private WithEvents btnAdd As SAPbouiCOM.Button
        Private WithEvents btnCancel As SAPbouiCOM.Button
        Private WithEvents lkVehi As SAPbouiCOM.LinkedButton
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents cboWOType As SAPbouiCOM.Column
        Private WithEvents txtWOCode As SAPbouiCOM.Column
        Private WithEvents txtDesc As SAPbouiCOM.Column
        Private WithEvents txtVendor As SAPbouiCOM.Column
        Private WithEvents cflVendor As SAPbouiCOM.ChooseFromList
        Private WithEvents cflService As SAPbouiCOM.ChooseFromList
        Private WithEvents cflFleet As SAPbouiCOM.ChooseFromList
        Private WithEvents txtQty As SAPbouiCOM.Column
        Private WithEvents txtUPrice As SAPbouiCOM.Column
        Private WithEvents txtAmount As SAPbouiCOM.Column
        Private WithEvents txtToWOAmt As SAPbouiCOM.EditText
        Private WithEvents oMatrix As SAPbouiCOM.Matrix
        Private WithEvents btnPost As SAPbouiCOM.Button
        Private WithEvents DocEntry As SAPbouiCOM.EditText


        Private Sub LoadFilters()
            Try
                oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER").Offset = 0
                oForm.Select()

                oConditions = New SAPbouiCOM.Conditions
                oCondition = oConditions.Add

                With oCondition
                    .Alias = "CardType"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "S"
                End With

                cflVendor = oForm.ChooseFromLists.Item("CFLVendor")
                cflVendor.SetConditions(oConditions)

                'filter for services
                oConditions = New SAPbouiCOM.Conditions
                oCondition = Nothing
                oCondition = oConditions.Add

                With oCondition
                    .Alias = "U_PM"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = "Y"
                End With

                cflService = oForm.ChooseFromLists.Item("CFLServices")
                cflService.SetConditions(oConditions)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Finally
                'oForm.Freeze(False)
            End Try

        End Sub

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "1288", "1289", "1290", "1291"
                        oForm = SBO_Application.Forms.ActiveForm
                        If oForm.TypeEx = "FleetMgt.WorkOrder" Then
                            HandleNavigation()
                        End If

                    Case "1282", "1283"
                        Reset()
                    Case "519"

                        PrintRequest()
                        BubbleEvent = False
                End Select
            Else

            End If

        End Sub

        Private Sub CalcSubTotal()
            Dim oEditTotal As SAPbouiCOM.EditText   ' Total = Item Price * Item Quantity
            Dim CalcTotal As Double
            Dim i As Integer

            CalcTotal = 0
            ' Iterate all the matrix rows
            oMatrix = oForm.Items.Item("WODetail").Specific
            txtAmount = oMatrix.Columns.Item("txtAmount")
            txtToWOAmt = oForm.Items.Item("txtToWOAmt").Specific

            For i = 1 To oMatrix.RowCount
                oEditTotal = txtAmount.Cells.Item(i).Specific
                CalcTotal += oEditTotal.Value
            Next

            txtToWOAmt.Item.Enabled = False
            txtToWOAmt.Value = CalcTotal
        End Sub
        Private Sub cboWOType_ComboSelectAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles cboWOType.ComboSelectAfter
            'checks the index value
            If pVal.ItemChanged Then
                Select Case pVal.PopUpIndicator
                    Case 1
                        txtWOCode.ChooseFromListUID = "CFLServices"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSERVICES"
                        FilterServiceByType("N")
                    Case 2
                        txtWOCode.ChooseFromListUID = "CFLSparePart"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSPAREPART"
                    Case 3
                        txtWOCode.ChooseFromListUID = "CFLPM"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSERVICES"
                        FilterServiceByType("Y")
                End Select
            End If

        End Sub

        Private Sub FilterServiceByType(ByVal type As String)

            'filter for services
            oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER").Offset = 0
            oForm.Select()

            oConditions = New SAPbouiCOM.Conditions
            oCondition = oConditions.Add

            With oCondition
                .Alias = "U_PM"
                .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                .CondVal = type
            End With

            cflService = oForm.ChooseFromLists.Item("CFLServices")
            cflService.SetConditions(oConditions)


        End Sub

        Private Sub txtWOCode_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtWOCode.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal

            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDERDET")
            oMatrix.FlushToDataSource()
            oDataSource.SetValue("U_WOCode", pVal.Row - 1, val)
            If cboWOType.Cells.Item(pVal.Row).Specific.Value = "SV" Then
                getOffset(val, "Code", DBDS("@OWA_FMSERVICES"))
                oDataSource.SetValue("U_Description", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("Name", 0).Trim)
            End If

            If cboWOType.Cells.Item(pVal.Row).Specific.Value = "SP" Then
                getOffset(val, "Code", DBDS("@OWA_FMSPAREPARTS"))
                oDataSource.SetValue("U_Description", pVal.Row - 1, DBDS("@OWA_FMSPAREPARTS").GetValue("Name", 0).Trim)
            End If

            If cboWOType.Cells.Item(pVal.Row).Specific.Value = "PM" Then
                getOffset(val, "Code", DBDS("@OWA_FMSERVICES"))
                oDataSource.SetValue("U_Description", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("Name", 0).Trim)
            End If

            oMatrix.LoadFromDataSource()

            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Private Sub txtWOCode_GotFocusAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtWOCode.GotFocusAfter
            Try

                Select Case cboWOType.Cells.Item(pVal.Row).Specific.Value
                    Case "SV"
                        txtWOCode.ChooseFromListUID = "CFLServices"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSERVICES"
                    Case "SP"
                        txtWOCode.ChooseFromListUID = "CFLSparePart"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSPAREPART"
                    Case "PM"
                        txtWOCode.ChooseFromListUID = "CFLPM"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSERVICES"
                End Select
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub txtWOCode_LinkPressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles txtWOCode.LinkPressedBefore
            Try

                Select Case cboWOType.Cells.Item(pVal.Row).Specific.Value
                    Case "SV"
                        txtWOCode.ChooseFromListUID = "CFLServices"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSERVICES"
                    Case "SP"
                        txtWOCode.ChooseFromListUID = "CFLSparePart"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSPAREPART"
                    Case "PM"
                        txtWOCode.ChooseFromListUID = "CFLPM"
                        txtWOCode.ChooseFromListAlias = "Code"
                        oLnkWOCode.LinkedObjectType = "OWA_FMSERVICES"
                End Select
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Finally
                'oForm.Freeze(False)
            End Try
        End Sub

        Private Sub txtVehi_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtVehicle.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")
            oDataSource.SetValue("U_FleetID", 0, val)

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            If getOffset(val, "Code", oDataSource) Then
                oDataSource.Offset = 0
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End Sub


        Private Sub HandleNavigation()
            Try

                txtVehicle = oForm.Items.Item("txtVehi").Specific
                getOffset(txtVehicle.Value, "Code", oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET"))
                CalcSubTotal()


            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub


        Private Sub txtVendor_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtVendor.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDERDET")
            oMatrix.FlushToDataSource()
            oDataSource.SetValue("U_CardCode", pVal.Row - 1, val)
            oMatrix.LoadFromDataSource()
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter
            AddHandler DataAddAfter, AddressOf Me.WorkOrder_DataAddAfter
            AddHandler DataLoadAfter, AddressOf Me.WorkOrder_DataLoadAfter

        End Sub

        Private Sub Reset()
            Try
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
                oDataSource.SetValue("Name", 0, "")

                oDataSource = oForm.DataSources.DBDataSources.Item("OHEM")
                oDataSource.SetValue("FirstName", 0, "")

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub txtQty_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles txtQty.ValidateBefore
            Try
                Dim oEditPrice As SAPbouiCOM.EditText   ' Item Price
                Dim oEditQuan As SAPbouiCOM.EditText    ' Item Quantity
                Dim oEditTotal As SAPbouiCOM.EditText   ' Total = Item Price * Item Quantity

                ' Get the items from the matrix
                oEditPrice = txtUPrice.Cells.Item(pVal.Row).Specific
                oEditQuan = txtQty.Cells.Item(pVal.Row).Specific
                oEditTotal = txtAmount.Cells.Item(pVal.Row).Specific

                txtAmount.Editable = False
                oEditTotal.Value = If(oEditQuan.Value.Length = 0, 0, oEditQuan.Value) * If(oEditPrice.Value.Length = 0, 0, oEditPrice.Value)

                ' Calc the document total

                Dim CalcTotal As Double
                Dim i As Integer

                CalcTotal = 0
                ' Iterate all the matrix rows
                For i = 1 To oMatrix.RowCount
                    oEditTotal = txtAmount.Cells.Item(i).Specific
                    CalcTotal += oEditTotal.Value
                Next
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")
                oDataSource.SetValue("U_TotalWOAmt", oDataSource.Offset, CalcTotal)
                oForm.Update()
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub txtUPrice_ValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles txtUPrice.ValidateBefore
            Try
                Dim oEditPrice As SAPbouiCOM.EditText   ' Item Price
                Dim oEditQuan As SAPbouiCOM.EditText    ' Item Quantity
                Dim oEditTotal As SAPbouiCOM.EditText   ' Total = Item Price * Item Quantity

                ' Get the items from the matrix
                oEditPrice = txtUPrice.Cells.Item(pVal.Row).Specific
                oEditQuan = txtQty.Cells.Item(pVal.Row).Specific
                oEditTotal = txtAmount.Cells.Item(pVal.Row).Specific

                txtAmount.Editable = False
                oEditTotal.Value = If(oEditQuan.Value.Length = 0, 0, oEditQuan.Value) * If(oEditPrice.Value.Length = 0, 0, oEditPrice.Value)

                ' Calc the document total

                Dim CalcTotal As Double
                Dim i As Integer

                CalcTotal = 0
                ' Iterate all the matrix rows
                For i = 1 To oMatrix.RowCount
                    oEditTotal = txtAmount.Cells.Item(i).Specific
                    CalcTotal += oEditTotal.Value
                Next
                oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")
                oDataSource.SetValue("U_TotalWOAmt", oDataSource.Offset, CalcTotal)
                oForm.Update()
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub oMatrix_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles oMatrix.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDERDET").Clear()
            oMatrix = oForm.Items.Item("WODetail").Specific

            If pVal.Row = oMatrix.RowCount + 1 Then
                If pVal.Row = 1 Then
                    oMatrix.AddRow(1)
                Else
                    oMatrix.AddRow(1, oMatrix.RowCount)
                End If
                oMatrix.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub

        Private Sub OnCustomInitialize()

        End Sub

        Private Sub WorkOrder_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataAddAfter
            Reset()
        End Sub


        Private Sub btnPost_ClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles btnPost.ClickAfter
            Dim EquipID As String, DateCompld As String

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")

            If SBO_Application.MessageBox("Do you want to post Work Order?", 2, "Yes", "No", "") = 2 Then
                Exit Sub
            Else
                'tests if the work order is opened
                If oDataSource.GetValue("U_Status", 0) = "Closed" Then
                    SBO_Application.MessageBox("Work Order already posted")
                    Exit Sub
                End If

                'tests if the rquired field are supplied

                EquipID = oDataSource.GetValue("U_FleetID", 0)
                DateCompld = oDataSource.GetValue("U_PostingDate", 0)
                DocEntry = oForm.Items.Item("txtDocNo").Specific

                If EquipID <> "" And DateCompld <> "" Then
                    PostWorkOrder(DocEntry.Value)
                    oForm.Update()
                    SBO_Application.MessageBox("Work Order successfully posted")
                Else
                    SBO_Application.MessageBox("Check if Date of completion is supplied")
                End If

            End If

        End Sub


        Private Sub PostWorkOrder(ByVal docno As Integer)
            Dim strSQL As String
            strSQL = Me.GetSQLString("owa_postworkorder", docno)
            oDataTable = oForm.DataSources.DataTables.Item("DT_Seq")
            oDataTable.ExecuteQuery(strSQL)
        End Sub

        Private Sub PrintRequest()

            Try
                Dim ReportFile As String = System.Windows.Forms.Application.StartupPath + "\Work Order Document.rpt"
                Dim orep As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                orep.Load(ReportFile)
                DocEntry = oForm.Items.Item("txtDocNo").Specific
                orep.SetParameterValue("DocEntry@", CInt(DocEntry.Value.Trim()))
                Dim frm As frmReports = New frmReports(orep)
                frm.ShowDialog()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub


        Private Sub WorkOrder_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataLoadAfter
            'checks if the work order's status is closed. If yes, disable d grid
            oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")
            oMatrix = oForm.Items.Item("WODetail").Specific
            If oDataSource2.GetValue("U_Status", 0).Trim = "Closed" Then
                oMatrix.Columns.Item(0).Editable = False
                oMatrix.Columns.Item(1).Editable = False
                oMatrix.Columns.Item(2).Editable = False
                oMatrix.Columns.Item(3).Editable = False
                oMatrix.Columns.Item(4).Editable = False
                oMatrix.Columns.Item(5).Editable = False
                oMatrix.Columns.Item(6).Editable = False
                oMatrix.Columns.Item(7).Editable = False

            Else
                oMatrix.Columns.Item(0).Editable = True
                oMatrix.Columns.Item(1).Editable = True
                oMatrix.Columns.Item(2).Editable = True
                oMatrix.Columns.Item(3).Editable = True
                oMatrix.Columns.Item(4).Editable = True
                oMatrix.Columns.Item(5).Editable = True
                oMatrix.Columns.Item(6).Editable = True
                oMatrix.Columns.Item(7).Editable = True
            End If

        End Sub
        Private WithEvents optEquip As SAPbouiCOM.OptionBtn
        Private WithEvents optVehi As SAPbouiCOM.OptionBtn

        Private Sub optEquip_ClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles optEquip.ClickAfter
            optEquip = oForm.Items.Item("optEquip").Specific
            If optEquip.Selected = False Then
                FilterFleet("E")
            End If
        End Sub

        Private Sub FilterFleet(ByVal type As String)
            Try
                oConditions = New SAPbouiCOM.Conditions
                oCondition = oConditions.Add

                With oCondition
                    .Alias = "U_Category"
                    .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    .CondVal = type
                End With

                cflFleet = oForm.ChooseFromLists.Item("CFLVehi")
                cflFleet.SetConditions(oConditions)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Finally
                'oForm.Freeze(False)
            End Try

        End Sub

        Private Sub optVehi_ClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles optVehi.ClickAfter
            optVehi = oForm.Items.Item("optVehi").Specific
            If optVehi.Selected = False Then
                FilterFleet("V")
            End If
        End Sub

        Private Sub txtrptdby_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles txtrptdby.ChooseFromListAfter

            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")
            oDataSource.SetValue("U_ReportedBy", 0, val)
            oDataSource = oForm.DataSources.DBDataSources.Item("OHEM")
            If getOffset(val, "empID", oDataSource) Then
                oDataSource.Offset = 0
                getOffset(val, "empID", oDataSource)
                oDataSource2 = oForm.DataSources.DBDataSources.Item("@OWA_FMWORKORDER")
                oDataSource2.SetValue("U_ReportedByName", 0, Trim(oDataSource.GetValue("lastName", 0)) & ", " & Trim(oDataSource.GetValue("firstName", 0)))
            End If

            ' Put form in UPDATE Mode when in OK Mode
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents txtrptdby As SAPbouiCOM.EditText
        Private WithEvents EditText1 As SAPbouiCOM.EditText

    End Class



End Namespace
