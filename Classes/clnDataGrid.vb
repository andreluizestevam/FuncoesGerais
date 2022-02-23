Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Web.UI.WebControls

Namespace DataGridWin
    Public Class clnDataGrid

        Public Overloads Shared Sub pfCarrega(ByVal grdGrid As Windows.Forms.DataGrid, ByVal strProcedure As String)
            Dim dsGrade As DataSet
            dsGrade = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure)
            With grdGrid
                .DataSource = dsGrade.Tables(0)
            End With
        End Sub
        Public Overloads Shared Sub pfCarrega(ByVal grdGrid As Windows.Forms.DataGrid, ByVal strProcedure As String, ByVal ParamArray commandParameters() As SqlParameter)
            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure _
                                                , strProcedure, _
                                                commandParameters)
            With grdGrid
                .DataSource = dsCombo.Tables(0)
            End With
        End Sub

    End Class
End Namespace

Namespace DataGridWeb
    Public Class clnDataGrid

        Public Overloads Shared Sub pfCarregaDataGrid(ByVal grdGrid As Web.UI.WebControls.DataGrid, ByVal strProcedure As String)

            Dim dsGrade As DataSet
            dsGrade = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure, strProcedure)

            With grdGrid
                .DataSource = dsGrade.Tables(0)
                .DataBind()
            End With

        End Sub
        Public Overloads Shared Sub pfCarregaDataGrid(ByVal grdGrid As Web.UI.WebControls.DataGrid, ByVal strProcedure As String, ByVal ParamArray commandParameters() As SqlParameter)

            Dim dsCombo As DataSet
            dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure _
                                                , strProcedure, _
                                                commandParameters)

            With grdGrid
                .DataSource = dsCombo.Tables(0)
                .DataBind()
            End With

        End Sub
        Public Overloads Shared Sub pfCarregaDataGrid(ByVal grdGrid As Web.UI.WebControls.DataGrid, _
                        ByVal strProcedure As String, _
                        ByVal pagPagina As System.Web.UI.Page, _
                        ByVal strNomeCache As String, _
                        ByVal douTempo As Double, _
                        ByVal blnRemoveCache As Boolean, _
                        ByVal ParamArray commandParameters() As SqlParameter)

            Dim dsCombo As DataSet
            dsCombo = DirectCast(pagPagina.Cache(strNomeCache), DataSet)

            If dsCombo Is Nothing Then
                dsCombo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure _
                                                    , strProcedure, _
                                                    commandParameters)

                pagPagina.Cache.Insert(strNomeCache, dsCombo, Nothing, System.DateTime.Now.AddMinutes(douTempo), System.Web.Caching.Cache.NoSlidingExpiration)

            End If

            With grdGrid
                .DataSource = dsCombo.Tables(0)
                .DataBind()
            End With

        End Sub
    End Class
End Namespace
