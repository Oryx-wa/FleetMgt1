Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace FleetMgt
    <FormAttribute("FleetMgt.MtceShedule", "SBO_Forms/MtceShedule.b1f")>
    Friend Class MtceShedule
        Inherits UserFormBaseClass

        Public Sub New()
            GetData()
        End Sub

        Private Sub GetData()
            Try

                fillCombo("Code", "Name", "@OWA_FMUOM", cboOdoType, "U_Type='Mileage'", True)

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.matMtcSchd = CType(Me.GetItem("matMtcSchd").Specific, SAPbouiCOM.Matrix)
            Me.colSvrCode = CType(Me.GetItem("matMtcSchd").Specific, SAPbouiCOM.Matrix).Columns.Item("colSvrCode")
            Me.cboOdoType = CType(Me.GetItem("cboOdoType").Specific, SAPbouiCOM.ComboBox)
            Me.OnCustomInitialize()

        End Sub

        Private WithEvents matMtcSchd As SAPbouiCOM.Matrix
        Private WithEvents cboOdoType As SAPbouiCOM.ComboBox
        Private WithEvents colSvrCode As SAPbouiCOM.Column

        Private Sub OnCustomInitialize()

        End Sub

        Private Sub colSvrCode_ChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles colSvrCode.ChooseFromListAfter
            Dim val As String
            oCFLEventArg = pVal
            oDataTable = oCFLEventArg.SelectedObjects
            If oDataTable Is Nothing Then
                Exit Sub
            End If

            val = oDataTable.GetValue(0, 0)

            oDataSource = oForm.DataSources.DBDataSources.Item("@OWA_FMMTCESHEDDET")
            matMtcSchd.FlushToDataSource()
            oDataSource.SetValue("U_SvrCode", pVal.Row - 1, val)

            getOffset(val, "Code", DBDS("@OWA_FMSERVICES"))
            oDataSource.SetValue("U_SvrDesc", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("Name", 0).Trim)
            oDataSource.SetValue("U_MtceIntervalDays", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("U_MtceIntervalDays", 0).Trim)
            oDataSource.SetValue("U_MtceIntervalMeter", pVal.Row - 1, DBDS("@OWA_FMSERVICES").GetValue("U_MtceIntervalMeter", 0).Trim)

            matMtcSchd.LoadFromDataSource()

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub


        Private Sub oMatrix_PressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles matMtcSchd.PressedBefore
            oForm.DataSources.DBDataSources.Item("@OWA_FMMTCESHEDDET").Clear()
            matMtcSchd = oForm.Items.Item("matMtcSchd").Specific

            If pVal.Row = matMtcSchd.RowCount + 1 Then
                If pVal.Row = 1 Then
                    matMtcSchd.AddRow(1)
                Else
                    matMtcSchd.AddRow(1, matMtcSchd.RowCount)
                End If
                matMtcSchd.Columns.Item(1).Cells.Item(pVal.Row).Click()
            End If
        End Sub

    End Class
End Namespace
