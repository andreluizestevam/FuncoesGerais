Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Web.UI.WebControls

Namespace ListBoxWin
    Public Class clnListBox

        Public Overloads Shared Sub pfCarregaList(ByVal lslList As Windows.Forms.ListBox, ByVal strProcedure As String, Optional ByVal blnOpcao As Boolean = False)
            Dim dsList As DataSet
            dsList = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dsList.Tables(0).NewRow
                dr.Item(0) = 0
                dr.Item(1) = "Selecione"
                dsList.Tables(0).Rows.InsertAt(dr, 0)
            End If
            With lslList
                .DataSource = dsList.Tables(0)
                .DisplayMember = dsList.Tables(0).Columns(1).ColumnName
                .ValueMember = dsList.Tables(0).Columns(0).ColumnName
            End With
        End Sub
        Public Overloads Shared Sub pfCarregaList(ByVal dsList As DataSet, ByVal lslList As Windows.Forms.ListBox, ByVal strProcedure As String, Optional ByVal blnOpcao As Boolean = False)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dsList.Tables(0).NewRow
                dr.Item(0) = 0
                dr.Item(1) = "Selecione"
                dsList.Tables(0).Rows.InsertAt(dr, 0)
            End If
            With lslList
                .DataSource = dsList.Tables(0)
                .DisplayMember = dsList.Tables(0).Columns(1).ColumnName
                .ValueMember = dsList.Tables(0).Columns(0).ColumnName
            End With
        End Sub

    End Class
End Namespace

Namespace ListBoxWeb
    Public Class clnListBox

    End Class
End Namespace
