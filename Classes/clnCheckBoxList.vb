Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Web.UI.WebControls

Namespace CheckBoxListWin
    Public Class clnCheckBoxList

        Public Shared Sub pfCarregaCheckList(ByVal lslList As Windows.Forms.ListBox, ByVal dsCombo As DataSet, Optional ByVal blnOpcao As Boolean = False)
            With lslList
                .DataSource = dsCombo.Tables(0)
                .DisplayMember = dsCombo.Tables(0).Columns(1).ColumnName
                .ValueMember = dsCombo.Tables(0).Columns(0).ColumnName
            End With
        End Sub

    End Class
End Namespace

Namespace CheckBoxListWeb
    Public Class clnCheckBoxList

    End Class
End Namespace
