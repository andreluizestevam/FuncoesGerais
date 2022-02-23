Public Class frmRelatorioErro

    Private _Categoria As Decimal
    Private _SubCategoria As Decimal

    Public Sub New(ByVal Categoria As Decimal, ByVal SubCategoria As Decimal)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        _Categoria = Categoria
        _SubCategoria = SubCategoria
    End Sub

    Private Sub frmRelatorioErro_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            'Dim sAppSubCategoria As New Configuration.AppSettingsReader

            'Dim clsSISCOP As New SISCOP.clnSISCOP
            'clsSISCOP.Categoria = _Categoria
            'clsSISCOP.Descricao = txtDescricao.Text.Trim
            'clsSISCOP.SubCategoria = _SubCategoria
            'clsSISCOP.Gravar()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
End Class