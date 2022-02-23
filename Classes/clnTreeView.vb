Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Web.UI.WebControls

Namespace TreeViewWin
    Public Class clnTreeView
        Public Overloads Shared Sub pfCarrega(ByVal trvGrupo As Windows.Forms.TreeView, ByVal strProcedure As String, ByVal ParamArray commandParameters() As SqlParameter)

            Dim contGrupo As Integer
            Dim imgTreeView As System.Windows.Forms.ImageList
            Dim dsGrupo As New DataSet

            dsGrupo = SqlHelper.ExecuteDataset(clnServidor.fpuObterStrConn(), CommandType.StoredProcedure _
                                                , strProcedure, _
                                                commandParameters)

            trvGrupo.Nodes.Clear()
            trvGrupo.ImageList = imgTreeView

            Dim rootNode As New Windows.Forms.TreeNode("GCM OnLine")
            trvGrupo.Nodes.Add(rootNode)
            rootNode.ImageIndex = 1
            rootNode.SelectedImageIndex = 1

            Dim strDescricao As String
            Dim contUsuario As Integer = 0

            For contGrupo = 0 To dsGrupo.Tables(0).Rows.Count - 1
                If strDescricao <> dsGrupo.Tables(0).Rows(contGrupo).Item("gcmonline_gru_descricao").ToString Then
                    strDescricao = dsGrupo.Tables(0).Rows(contGrupo).Item("gcmonline_gru_descricao").ToString
                    Dim Filho As New Windows.Forms.TreeNode(strDescricao)
                    trvGrupo.Nodes(0).Nodes.Add(Filho)
                    Filho.ImageIndex = 2
                    Filho.SelectedImageIndex = 2
                End If
            Next

            trvGrupo.ExpandAll()



        End Sub

        Public Overloads Shared Sub pfCarrega(ByVal trvGrupo As Windows.Forms.TreeView, ByVal dsGrupo As DataSet, ByVal imgTreeView As System.Windows.Forms.ImageList, ByVal strColuna As String, ByVal ImageIndexPai As Integer, ByVal ImageIndexFilho As Integer)

            Dim contGrupo As Integer

            trvGrupo.Nodes.Clear()
            trvGrupo.ImageList = imgTreeView

            Dim rootNode As New Windows.Forms.TreeNode("GCM OnLine")
            trvGrupo.Nodes.Add(rootNode)
            rootNode.ImageIndex = ImageIndexPai
            rootNode.SelectedImageIndex = ImageIndexPai

            Dim strDescricao As String
            Dim contUsuario As Integer = 0

            For contGrupo = 0 To dsGrupo.Tables(0).Rows.Count - 1
                If strDescricao <> dsGrupo.Tables(0).Rows(contGrupo).Item(strColuna).ToString Then
                    strDescricao = dsGrupo.Tables(0).Rows(contGrupo).Item(strColuna).ToString
                    Dim Filho As New Windows.Forms.TreeNode(strDescricao)
                    trvGrupo.Nodes(0).Nodes.Add(Filho)
                    Filho.ImageIndex = ImageIndexFilho
                    Filho.SelectedImageIndex = ImageIndexFilho
                End If
            Next

            'trvGrupo.ExpandAll()

        End Sub

    End Class
End Namespace
