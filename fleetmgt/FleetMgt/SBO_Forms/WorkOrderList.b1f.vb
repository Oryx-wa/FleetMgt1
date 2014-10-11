Option Strict Off
Option Explicit On
Imports SBO.SboAddOnBase

Imports SAPbouiCOM.Framework

Namespace FleetMgt
    <FormAttribute("FleetMgt.WorkOrderList", "SBO_Forms/WorkOrderList.b1f")>
    Friend Class WorkOrderList
        Inherits UserFormBaseclass

        Private WithEvents oGrid As SAPbouiCOM.Grid

        Public Sub New()
            DisplayWorkOrderList("")
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.cboStatus = CType(Me.GetItem("cboStatus").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText0 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub



        Private Sub DisplayWorkOrderList(ByVal status As String)
            Try
                Dim sQuery As String

                sQuery = "select docentry [S/N],U_FleetID [Equip COde],U_DocumentDate [Date Issued],U_PostingDate [Completion Date], "
                sQuery = sQuery + "U_TotalWOAmt [Total Cost] , U_Status Status from [@owa_fmworkorder] where U_Status='" & status & "'"

                oGrid = oForm.Items.Item("grdWOrder").Specific
                oDataTable = oForm.DataSources.DataTables.Item("DTWOrder")
                oDataTable.ExecuteQuery(sQuery)
                oGrid.DataTable = oDataTable

                'oGrid.Columns.Item(0).

                oGrid.Columns.Item(0).TitleObject.Caption = "S/N"
                oGrid.Columns.Item(1).TitleObject.Caption = "Equip Code"
                oGrid.Columns.Item(2).TitleObject.Caption = "Date Issued"
                oGrid.Columns.Item(3).TitleObject.Caption = "Completion Date"
                oGrid.Columns.Item(4).TitleObject.Caption = "Total Cost"
                oGrid.Columns.Item(5).TitleObject.Caption = "Status"

                oGrid.AutoResizeColumns()

            Catch ex As Exception
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try


        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()

        End Sub
        Private WithEvents cboStatus As SAPbouiCOM.ComboBox
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText

   
        Private Sub cboStatus_ComboSelectAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg) Handles cboStatus.ComboSelectAfter
            cboStatus = oForm.Items.Item("cboStatus").Specific
            DisplayWorkOrderList(cboStatus.Value)
        End Sub
    End Class
End Namespace
