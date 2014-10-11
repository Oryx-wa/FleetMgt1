Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace FleetMgt
    <FormAttribute("FleetMgt.FleetSetupScreen", "SBO_Forms/FleetSetupScreen.b1f")>
    Friend Class FleetSetupScreen
        Inherits UserFormBaseClassOld

        Public Sub New()

            Load()

            GetData()

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.btnAdd = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.btnCancel = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.cboDim = CType(Me.GetItem("cboDim").Specific, SAPbouiCOM.ComboBox)
            Me.lblDimName = CType(Me.GetItem("lblDimName").Specific, SAPbouiCOM.StaticText)
            Me.StaticText0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.StaticText)
            Me.StaticText1 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("Item_2").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("Item_3").Specific, SAPbouiCOM.StaticText)
            Me.Button0 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.Button)
            Me.StaticText4 = CType(Me.GetItem("Item_7").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("Item_9").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("Item_10").Specific, SAPbouiCOM.StaticText)
            Me.StaticText6 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.StaticText)
            Me.Button1 = CType(Me.GetItem("Item_12").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub


        Private WithEvents btnAdd As SAPbouiCOM.Button
        Private WithEvents btnCancel As SAPbouiCOM.Button
        Private WithEvents cboDim As SAPbouiCOM.ComboBox
        Private WithEvents lblDimName As SAPbouiCOM.StaticText


        Private Function Save() As Boolean

            Dim oSetupTable As SAPbobsCOM.UserTable, lAdd As Boolean
            oSetupTable = getUserTables("@OWA_FMSETUP")

            If Not oSetupTable.GetByKey("FLEET") Then lAdd = True

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMSETUP")

            oSetupTable.UserFields.Fields.Item("U_DimCode").Value = oDataSource.GetValue("U_DimCode", 0)

            If lAdd Then
                oSetupTable.Code = "FLEET"
                oSetupTable.Name = "FLEET Setup"
                If oSetupTable.Add() <> 0 Then
                    oCompany.GetLastError(errCode, errMsg)
                    SBO_Application.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Else
                If oSetupTable.Update() <> 0 Then
                    oCompany.GetLastError(errCode, errMsg)
                    SBO_Application.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True

        End Function

        Private Sub GetData()
            Try
                fillCombo("DimCode", "DimDesc", "ODIM", cboDim, "DimActive='Y'", True)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Load()
            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMSETUP")
            oConditions = New SAPbouiCOM.Conditions
            oCondition = oConditions.Add
            oCondition.Alias = "Code"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "FLEET"
            oDataSource.Query(oConditions)

            If oDataSource.Size > 0 Then
                oDataSource.Offset = 0
            Else
                oDataSource.InsertRecord(0)
                oDataSource.Offset = 0
                Save()
            End If

            Refresh()
        End Sub

        Private Sub btnAdd_ClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles btnAdd.ClickAfter
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                Save()
            End If
        End Sub

        Private Sub Refresh()
            If validatealias(DBDS("@OWA_FMSETUP").GetValue("U_DimCode", 0), "DimCode", DBDS("ODIM")) Then
                lblDimName.Caption = DBDS("ODIM").GetValue("DimName", 0)
            End If

  
            'GetMtceOverDueCount
            'oRecordSet.DoQuery("Update [@OWA_FMSETUP] set U_NoOfMtceDue=" & GetMtceOverDueCount() & " where Code='FLEET'")
           
            'GetMtceDueTodayCount
            oRecordSet.DoQuery("Update [@OWA_FMSETUP] set U_NoOfMtceDue2Day=" & GetMtceDueTodayCount() & " where Code='FLEET'")

            'GetWorkOrderCount
            oRecordSet.DoQuery("Update [@OWA_FMSETUP] set U_NoOfWO=" & GetWorkOrderCount() & " where Code='FLEET'")

            oForm.Update()
            oForm.Refresh()

        End Sub

        Private Function GetMtceOverDueCount() As Integer
            Dim strSQL As String
            strSQL = Me.GetSQLString("owa_getmtceoverdue", "", "")
            oDataTable = oForm.DataSources.DataTables.Item("DT_Seq")
            oDataTable.ExecuteQuery(strSQL)
            If oDataTable.Rows.Count = 1 Then
                Return oDataTable.GetValue("kount", 0)
            Else
                Return 0
            End If
        End Function

        Private Function GetMtceDueTodayCount() As Integer
            Dim strSQL As String
            strSQL = Me.GetSQLString("owa_getmtcedue2day", "", "")
            oDataTable = oForm.DataSources.DataTables.Item("DT_Seq")
            oDataTable.ExecuteQuery(strSQL)
            If oDataTable.Rows.Count = 1 Then
                Return oDataTable.GetValue("kount", 0)
            Else
                Return 0
            End If
        End Function


        Private Function GetWorkOrderCount() As Integer
            Dim strSQL As String
            strSQL = Me.GetSQLString("owa_getnoofworkorder", "", "")
            oDataTable = oForm.DataSources.DataTables.Item("DT_Seq")
            oDataTable.ExecuteQuery(strSQL)
            If oDataTable.Rows.Count = 1 Then
                Return oDataTable.GetValue("kount", 0)
            Else
                Return 0
            End If
        End Function

        Private Function GetEquipDueForMtce() As SAPbouiCOM.DataTable
            Dim strSQL As String
            oDataTable = Nothing
            strSQL = Me.GetSQLString("owa_getequipdueformtce", "", "")
            oDataTable = oForm.DataSources.DataTables.Item("DT_Seq")
            oDataTable.ExecuteQuery(strSQL)
            Return oDataTable
        End Function

        Private Function GetMtceLinesByEquipment(ByVal code As String) As SAPbouiCOM.DataTable
            Dim strSQL As String
            oDataTable = Nothing
            strSQL = Me.GetSQLString("owa_getmtcelinesbycode", code)
            oDataTable = oForm.DataSources.DataTables.Item("DT_Base")
            oDataTable.ExecuteQuery(strSQL)
            Return oDataTable
        End Function

        Private Function GetMtceLines() As SAPbouiCOM.DataTable
            Dim strSQL As String
            oDataTable = Nothing
            strSQL = Me.GetSQLString("owa_getmtcelines")
            oDataTable = oForm.DataSources.DataTables.Item("DT_Base")
            oDataTable.ExecuteQuery(strSQL)
            Return oDataTable
        End Function

        Private Sub GenerateWorkOrders()

            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oMaster As SAPbobsCOM.GeneralData
            Dim oChild As SAPbobsCOM.GeneralData
            Dim oChildren As SAPbobsCOM.GeneralDataCollection
            Dim sCmp As SAPbobsCOM.CompanyService = oCompany.GetCompanyService

            oGeneralService = sCmp.GetGeneralService("FMWKORDER")

            'get the number of work order to create
            Dim i As Integer
            Dim dtMtceLines As SAPbouiCOM.DataTable

            'Get all the Equipment due for Mtce
            dtMtceLines = GetMtceLines()

            For i = 0 To dtMtceLines.Rows.Count - 1
                'Add master here
                'Create data for new row in main UDO
                oMaster = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                With oMaster
                    .SetProperty("U_FleetID", dtMtceLines.GetValue("Equipment Code", i))
                    .SetProperty("U_DateStarted", Today.Date)
                    '.SetProperty("U_PostingDate", Today.Date)
                    '.SetProperty("U_DocumentDate", Today.Date)
                    .SetProperty("U_JobDesc", dtMtceLines.GetValue("Job Description", i))
                End With

                'Create data for a row in the child table
                oChildren = oMaster.Child("OWA_FMWORKORDERDET")

                'Add details here
                oChild = oChildren.Add
                oChild.SetProperty("U_WOType", "PM")
                oChild.SetProperty("U_WOCode", dtMtceLines.GetValue("Mtce Service Code", i))
                oChild.SetProperty("U_Description", dtMtceLines.GetValue("Mtce Service Name", i))

                oGeneralService.Add(oMaster)
                oChildren = Nothing
                oChild = Nothing
                oMaster = Nothing
            Next i


        End Sub

        Private Sub PrintRequest()
            Try
                Dim ReportFile As String = System.Windows.Forms.Application.StartupPath + "\Work Order List.rpt"

                Dim orep As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                orep.Load(ReportFile)
                orep.SetParameterValue("pComplDate", Today.Date.ToShortDateString)
                orep.SetParameterValue("pStatus", "Closed")
                Dim frm As frmReports = New frmReports(orep)
                frm.ShowDialog()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Button1_ClickBefore(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button1.ClickBefore
            'Throw New System.NotImplementedException()
            PrintRequest()
        End Sub

        Private Sub Button0_ClickBefore(ByVal sboObject As System.Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            'Throw New System.NotImplementedException()

            If SBO_Application.MessageBox("Do you want to Generate Work Orders?", 2, "Yes", "No", "") = 2 Then
                Exit Sub
            Else
                GenerateWorkOrders()
                SBO_Application.MessageBox("Work Orders successfully generated")
            End If

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents Button1 As SAPbouiCOM.Button



        
        Public Overrides Sub OnInitializeFormEvents()

        End Sub
       
    End Class
End Namespace
