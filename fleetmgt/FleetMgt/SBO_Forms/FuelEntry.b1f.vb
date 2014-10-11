Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase


Namespace FleetMgt
    <FormAttribute("FleetMgt.FuelEntry", "SBO_Forms/FuelEntry.b1f")>
    Friend Class FuelEntry
        Inherits UserFormBaseClassOld

        Dim oNavInicio As Boolean

        Public Sub New()

            LoadByPostingStatus()

        End Sub

        Private Sub LoadByPostingStatus()

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
            cflDriver = Nothing
            With oCondition
                .Alias = "u_fltEmpType"
                .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                .CondVal = "Driver"
            End With
            cflDriver = oForm.ChooseFromLists.Item("CFLDriver")
            cflDriver.SetConditions(oConditions)

        End Sub


        Private Sub FilterRecordSet()
            Try
                If Not IsNothing(oRecordSet) Then
                    oRecordSet.DoQuery(String.Format("select * from [@OWA_FMFUELFILL] where U_Post = 'N' order by docentry"))

                    oNavInicio = True
                End If

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub

        Private Sub LoadByDocEntry(ByVal aDocEntry As Integer)
            '
            'If oForm.UniqueID = "FuelEntry" Then
            oConditions = New SAPbouiCOM.Conditions
            oCondition = Nothing
            oCondition = oConditions.Add
            With oCondition
                .Alias = "DocEntry"
                .Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                .CondVal = Convert.ToString(aDocEntry)
            End With
            oForm.DataSources.DBDataSources.Item("@OWA_FMFUELFILL").Query(oConditions)
            oForm.DataSources.DBDataSources.Item("@OWA_FMFUELFILL").Offset = 0
            'End If

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.txtDriver = CType(Me.GetItem("txtDriver").Specific, SAPbouiCOM.EditText)
            Me.txtVehi = CType(Me.GetItem("txtVehi").Specific, SAPbouiCOM.EditText)
            Me.txtVendor = CType(Me.GetItem("txtVendor").Specific, SAPbouiCOM.EditText)
            Me.btnDriver = CType(Me.GetItem("btnDriver").Specific, SAPbouiCOM.Button)
            Me.btnVehi = CType(Me.GetItem("btnVehi").Specific, SAPbouiCOM.Button)
            Me.btnVendor = CType(Me.GetItem("btnVendor").Specific, SAPbouiCOM.Button)
            Me.lblDrvName = CType(Me.GetItem("lblDrvName").Specific, SAPbouiCOM.StaticText)
            Me.lblVenName = CType(Me.GetItem("lblVenName").Specific, SAPbouiCOM.StaticText)
            Me.btnPost = CType(Me.GetItem("btnPost").Specific, SAPbouiCOM.Button)

        End Sub

        Private WithEvents txtDriver As SAPbouiCOM.EditText
        Private WithEvents txtVehi As SAPbouiCOM.EditText
        Private WithEvents txtVendor As SAPbouiCOM.EditText
        Private WithEvents btnDriver As SAPbouiCOM.Button

        Private WithEvents btnVehi As SAPbouiCOM.Button
        Private WithEvents btnVendor As SAPbouiCOM.Button

        Private WithEvents lblDrvName As SAPbouiCOM.StaticText
        Private WithEvents lblVenName As SAPbouiCOM.StaticText

        Private WithEvents cflVendor As SAPbouiCOM.ChooseFromList
        Private WithEvents cflDriver As SAPbouiCOM.ChooseFromList


        'Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        '    Try
        '        If pVal.FormUID = FormUID Then

        '            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

        '                oCFLEvento = pVal

        '                Dim sCFL_ID As String
        '                sCFL_ID = oCFLEvento.ChooseFromListUID

        '                Dim oCFL As SAPbouiCOM.ChooseFromList
        '                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
        '                If oCFLEvento.BeforeAction = False Then
        '                    Dim oDataTable As SAPbouiCOM.DataTable
        '                    oDataTable = oCFLEvento.SelectedObjects
        '                    If oDataTable Is Nothing Then
        '                        Exit Sub
        '                    End If

        '                    Dim val As String

        '                    val = oDataTable.GetValue(0, 0)
        '                    Select Case pVal.ItemUID

        '                        Case "btnDriver"
        '                            txtDriver.Value = val
        '                            oDataSource = oForm.DataSources.DBDataSources.Item("OHEM")
        '                            If getOffset(val, "empid", oDataSource) Then
        '                                oDataSource.Offset = 0
        '                                lblDrvName.Caption = Trim(DBDS("OHEM").GetValue("firstName", 0)) & ", " & Trim(DBDS("OHEM").GetValue("lastName", 0))
        '                                oForm.Update()
        '                            End If

        '                        Case "btnVendor"
        '                            txtVendor.Value = val
        '                            oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
        '                            If getOffset(val, "CardCode", oDataSource) Then
        '                                oDataSource.Offset = 0
        '                                lblVenName.Caption = Trim(DBDS("OCRD").GetValue("CardName", 0))
        '                                oForm.Update()
        '                            End If

        '                        Case "btnVehi"
        '                            txtVehi.Value = val
        '                            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
        '                            If getOffset(val, "Code", oDataSource) Then
        '                                oDataSource.Offset = 0
        '                                oForm.Update()
        '                            End If
        '                    End Select
        '                End If
        '            End If
        '        End If

        '    Catch ex As Exception
        '        SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        '    Finally
        '        'If Not IsNothing(oForm) Then
        '        '    oForm.Freeze(False)
        '        'End If
        '    End Try
        'End Sub

        Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent

            Dim lBubbleEvent As Boolean

            lBubbleEvent = True

            oForm = SBO_Application.Forms.ActiveForm

            If oForm.TypeEx = "FleetMgt.FuelEntry" Then

                If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") Then

                    If (pVal.BeforeAction) Then
                        Select Case pVal.MenuUID

                            Case "1288"
                                If (oNavInicio) Then

                                    oRecordSet.MoveFirst()
                                    oNavInicio = False
                                Else
                                    If (Not oRecordSet.EoF) Then

                                        oRecordSet.MoveNext()
                                        If (oRecordSet.EoF) Then
                                            oRecordSet.MoveFirst()

                                        End If
                                    End If
                                End If

                            Case "1289"
                                If (oNavInicio) Then
                                    oRecordSet.MoveLast()
                                    oNavInicio = False

                                Else
                                    If (Not oRecordSet.BoF) Then
                                        oRecordSet.MovePrevious()
                                    Else
                                        oRecordSet.MoveLast()
                                    End If
                                End If

                            Case "1290"
                                oRecordSet.MoveFirst()
                                'Exit Select

                            Case "1291"
                                oRecordSet.MoveLast()
                                'Exit Select
                        End Select

                        oForm.Freeze(True)

                        ' Put form in OK Mode  
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                        ' Filter the DBDataSource with RecordSet's current record.
                        If oRecordSet.RecordCount > 0 Then
                            LoadByDocEntry(oRecordSet.Fields.Item("DocEntry").Value)
                        End If

                        oForm.Freeze(False)

                        ' Don't let the following the normal flow of SBO event  
                        lBubbleEvent = False

                    End If
                    HandleNavigation()
                    BubbleEvent = lBubbleEvent
                End If
            End If


        End Sub

        Private Sub HandleNavigation()
            Try
                If oForm.TypeEx = "FleetMgt.FuelEntry" Then
                    lblVenName = oForm.Items.Item("lblVenName").Specific
                    lblDrvName = oForm.Items.Item("lblDrvName").Specific

                    lblVenName.Caption = String.Empty
                    lblDrvName.Caption = String.Empty

                    oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFUELFILL")
                    oDataSource2 = oForm.DataSources.DBDataSources.Item("OCRD")

                    'Vendor Caption
                    If getOffset(oDataSource.GetValue("U_VendorID", 0), "CardCode", oDataSource2) Then
                        lblVenName.Caption = Trim(oDataSource2.GetValue("CardName", 0))
                    End If

                    oDataSource2 = oForm.DataSources.DBDataSources.Item("OHEM")
                    'Driver Name
                    If getOffset(oDataSource.GetValue("U_EmplyeeID", 0), "empID", oDataSource2) Then
                        lblDrvName.Caption = Trim(oDataSource2.GetValue("firstName", 0)) & ", " & Trim(oDataSource2.GetValue("lastName", 0))
                    End If
                    oForm.Update()
                End If
            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub
        Private WithEvents btnPost As SAPbouiCOM.Button


        Private Sub Post() Handles btnPost.ClickAfter

            'Try

            '    Dim yesno = MsgBox("Do you want to post fuel?", vbYesNo, "Fleet Manager")

            '    If yesno = vbYes Then
            '        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '        oRecordSet.DoQuery("Update [@OWA_FMFUELFILL] set U_Post='Y' where docentry=" & txtDocEnt.Value)

            '        FilterRecordSet()

            '        Sync()

            '        SBO_Application.StatusBar.SetText("Fuel Entry Successfully Posted!")

            '        oForm.Freeze(False)
            '    End If


            'Catch ex As Exception
            '    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            'End Try


        End Sub


        Private Sub Sync()

            If oRecordSet.RecordCount > 0 Then
                If oRecordSet.EoF Then
                    oRecordSet.MovePrevious()
                End If

                If oRecordSet.BoF Then
                    oRecordSet.MoveNext()
                End If

                oRecordSet.MoveNext()
                oRecordSet.MovePrevious()

                LoadByDocEntry(oRecordSet.Fields.Item("DocEntry").Value)
            End If
        End Sub

        Private Sub FuelEntry_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo) Handles Me.DataAddAfter
            FilterRecordSet()
        End Sub

        Private Sub FuelEntry_VisibleAfter(ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles Me.VisibleAfter
            FilterRecordSet()
        End Sub


        Public Overrides Sub OnInitializeFormEvents()
            AddHandler VisibleAfter, AddressOf Me.FuelEntry_VisibleAfter
            AddHandler VisibleAfter, AddressOf Me.FuelEntry_VisibleAfter

        End Sub

        Private Sub btnDriver_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles btnDriver.ChooseFromListAfter
            oCFLEventArg = pVal

            Dim oCFL As SAPbouiCOM.ChooseFromList, val As String, oDataTable As SAPbouiCOM.DataTable
            oCFL = oForm.ChooseFromLists.Item(btnDriver.ChooseFromListUID)

            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            txtDriver.Value = val
            oDataSource = oForm.DataSources.DBDataSources.Item("OHEM")
            If getOffset(val, "empid", oDataSource) Then
                oDataSource.Offset = 0
                lblDrvName.Caption = Trim(DBDS("OHEM").GetValue("firstName", 0)) & ", " & Trim(DBDS("OHEM").GetValue("lastName", 0))
                oForm.Update()
            End If
        End Sub


        Private Sub btnVendor_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles btnVendor.ChooseFromListAfter
            oCFLEventArg = pVal

            Dim oCFL As SAPbouiCOM.ChooseFromList, val As String, oDataTable As SAPbouiCOM.DataTable
            oCFL = oForm.ChooseFromLists.Item(btnVendor.ChooseFromListUID)

            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            txtVendor.Value = val
            oDataSource = oForm.DataSources.DBDataSources.Item("OCRD")
            If getOffset(val, "CardCode", oDataSource) Then
                oDataSource.Offset = 0
                lblVenName.Caption = Trim(DBDS("OCRD").GetValue("CardName", 0))
                oForm.Update()
            End If
        End Sub


        Private Sub btnVehi_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles btnVehi.ChooseFromListAfter
            oCFLEventArg = pVal

            Dim oCFL As SAPbouiCOM.ChooseFromList, val As String, oDataTable As SAPbouiCOM.DataTable
            oCFL = oForm.ChooseFromLists.Item(btnVehi.ChooseFromListUID)

            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            txtVehi.Value = val
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMFLEET")
            If getOffset(val, "Code", oDataSource) Then
                oDataSource.Offset = 0
                oForm.Update()
            End If
        End Sub

    End Class
End Namespace
