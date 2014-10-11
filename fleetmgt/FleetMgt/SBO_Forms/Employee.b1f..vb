
Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase

Namespace FleetMgt
    <FormAttribute("60100", "SBO_Forms/Employee.b1f")>
    Friend Class Employee
        Inherits SystemFormBase

        Public Sub New()

        End Sub

        Public Overrides Sub OnInitializeComponent()
            'Me.cboEmpType = CType(Me.GetItem("cboEmpType").Specific, SAPbouiCOM.ComboBox)
        End Sub

        Private WithEvents cboEmpType As SAPbouiCOM.ComboBox

    End Class
End Namespace
