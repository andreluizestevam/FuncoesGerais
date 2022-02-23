Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Web.UI.WebControls

Namespace DropDownListWin
    Public Class clnDropDownList
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Windows.Forms.ComboBox, ByVal dsCombo As DataSet, Optional ByVal blnOpcao As Boolean = False)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dsCombo.Tables(0).NewRow
                dr.Item(0) = 0
                dr.Item(1) = "Selecione / Select"
                dsCombo.Tables(0).Rows.InsertAt(dr, 0)
            End If
            With cmbCombo
                .DisplayMember = dsCombo.Tables(0).Columns(1).ColumnName
                .ValueMember = dsCombo.Tables(0).Columns(0).ColumnName
                .DataSource = dsCombo.Tables(0)
            End With
        End Sub

        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Windows.Forms.ComboBox, ByVal strProcedure As String, ByVal blnOpcao As Boolean, ByVal ParamArray commandParameters() As SqlParameter)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure, commandParameters)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dsCombo.Tables(0).NewRow
                dr.Item(0) = 0
                dr.Item(1) = "Selecione / Select"
                dsCombo.Tables(0).Rows.InsertAt(dr, 0)
            End If
            With cmbCombo
                .DisplayMember = dsCombo.Tables(0).Columns(1).ColumnName
                .ValueMember = dsCombo.Tables(0).Columns(0).ColumnName
                .DataSource = dsCombo.Tables(0)
            End With
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Windows.Forms.ComboBox, ByVal strProcedure As String, ByVal strColunaValor As String, ByVal strColunaDescricao As String, ByVal blnOpcao As Boolean, ByVal ParamArray commandParameters() As SqlParameter)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure, commandParameters)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dsCombo.Tables(0).NewRow
                dr.Item(strColunaValor) = 0
                dr.Item(strColunaDescricao) = "Selecione / Select"
                dsCombo.Tables(0).Rows.InsertAt(dr, 0)
            End If
            With cmbCombo
                .DisplayMember = strColunaDescricao
                .ValueMember = strColunaValor
                .DataSource = dsCombo.Tables(0)
            End With
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Windows.Forms.ComboBox, ByVal dsCombo As DataSet, ByVal strColunaValor As String, ByVal strColunaDescricao As String, ByVal blnOpcao As Boolean)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dsCombo.Tables(0).NewRow
                dr.Item(strColunaValor) = 0
                dr.Item(strColunaDescricao) = "Selecione / Select"
                dsCombo.Tables(0).Rows.InsertAt(dr, 0)
            End If
            With cmbCombo
                .DisplayMember = strColunaDescricao
                .ValueMember = strColunaValor
                .DataSource = dsCombo.Tables(0)
            End With
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Windows.Forms.ComboBox, ByVal dtCombo As DataTable, ByVal strColunaValor As String, ByVal strColunaDescricao As String, ByVal blnOpcao As Boolean)
            If blnOpcao = True Then
                Dim dr As DataRow
                dr = dtCombo.NewRow
                dr.Item(strColunaValor) = 0
                dr.Item(strColunaDescricao) = "Selecione / Select"
                dtCombo.Rows.InsertAt(dr, 0)
            End If
            With cmbCombo
                .DisplayMember = strColunaDescricao
                .ValueMember = strColunaValor
                .DataSource = dtCombo
            End With
        End Sub
    End Class
End Namespace

Namespace DropDownListWeb
    Public Class clnDropDownList

        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal dsCombo As DataSet, Optional ByVal blnOpcao As Boolean = False)
            With cmbCombo
                .DataSource = dsCombo.Tables(0)
                .DataTextField = dsCombo.Tables(0).Columns(1).ColumnName
                .DataValueField = dsCombo.Tables(0).Columns(0).ColumnName
                .DataBind()
            End With

           If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal dsCombo As DataSet, ByVal strValueMember As String, ByVal strTextMember As String, Optional ByVal blnOpcao As Boolean = False)
            With cmbCombo
                .DataSource = dsCombo.Tables(0)
                .DataTextField = strTextMember
                .DataValueField = strValueMember
                .DataBind()
            End With

           If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal dtCombo As DataTable, ByVal strValueMember As String, ByVal strTextMember As String, Optional ByVal blnOpcao As Boolean = False)
            With cmbCombo
                .DataSource = dtCombo
                .DataTextField = strTextMember
                .DataValueField = strValueMember
                .DataBind()
            End With

           If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal dvCombo As DataView, ByVal strValueMember As String, ByVal strTextMember As String, Optional ByVal blnOpcao As Boolean = False)
            With cmbCombo
                .DataSource = dvCombo.Table
                .DataTextField = strTextMember
                .DataValueField = strValueMember
                .DataBind()
            End With

           If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal strProcedure As String, ByVal strValueMember As String, ByVal strTextMember As String, ByVal blnOpcao As Boolean, ByVal ParamArray commandParameters() As SqlParameter)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure, commandParameters)
            With cmbCombo
                .DataSource = dsCombo.Tables(0)
                .DataTextField = dsCombo.Tables(0).Columns(1).ColumnName
                .DataValueField = dsCombo.Tables(0).Columns(0).ColumnName
                .DataBind()
            End With

           If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal strProcedure As String, ByVal blnOpcao As Boolean, ByVal ParamArray commandParameters() As SqlParameter)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure, commandParameters)
            With cmbCombo
                .DataSource = dsCombo.Tables(0)
                .DataTextField = dsCombo.Tables(0).Columns(1).ColumnName
                .DataValueField = dsCombo.Tables(0).Columns(0).ColumnName
                .DataBind()
            End With

            If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal strProcedure As String, ByVal blnOpcao As Boolean)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure)
            With cmbCombo
                .DataSource = dsCombo.Tables(0)
                .DataTextField = dsCombo.Tables(0).Columns(1).ColumnName
                .DataValueField = dsCombo.Tables(0).Columns(0).ColumnName
                .DataBind()
            End With

            If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
        Public Overloads Shared Sub pfCarregaCombo(ByVal cmbCombo As Web.UI.WebControls.DropDownList, ByVal strProcedure As String, ByVal strValueMember As String, ByVal strTextMember As String, ByVal blnOpcao As Boolean)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure)
            With cmbCombo
                .DataSource = dsCombo.Tables(0)
                .DataTextField = dsCombo.Tables(0).Columns(1).ColumnName
                .DataValueField = dsCombo.Tables(0).Columns(0).ColumnName
                .DataBind()
            End With

            If blnOpcao = True Then
                cmbCombo.Items.Insert(0, New ListItem("Selecione / Select", "0"))
            End If
        End Sub
    End Class
End Namespace
