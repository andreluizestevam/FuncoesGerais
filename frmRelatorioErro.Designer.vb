<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRelatorioErro
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRelatorioErro))
        Me.picErro = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDescricao = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        CType(Me.picErro, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picErro
        '
        Me.picErro.Image = CType(resources.GetObject("picErro.Image"), System.Drawing.Image)
        Me.picErro.Location = New System.Drawing.Point(12, 12)
        Me.picErro.Name = "picErro"
        Me.picErro.Size = New System.Drawing.Size(39, 37)
        Me.picErro.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picErro.TabIndex = 0
        Me.picErro.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(57, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(292, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "SISCOP - Sistema de Controle de Operações"
        '
        'txtDescricao
        '
        Me.txtDescricao.BackColor = System.Drawing.SystemColors.Control
        Me.txtDescricao.Location = New System.Drawing.Point(15, 97)
        Me.txtDescricao.Multiline = True
        Me.txtDescricao.Name = "txtDescricao"
        Me.txtDescricao.ReadOnly = True
        Me.txtDescricao.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDescricao.Size = New System.Drawing.Size(439, 184)
        Me.txtDescricao.TabIndex = 3
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(362, 289)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(92, 23)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "Fechar"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(442, 42)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "O Sistema executou uma operação ilegal. Foi aberto um chamado automaticamente par" & _
            "a a resolução deste problema. Para acompanhar o chamado, acesse o Portal GCMOnli" & _
            "ne."
        '
        'frmRelatorioErro
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(469, 317)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtDescricao)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.picErro)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRelatorioErro"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Relatório de Erro"
        CType(Me.picErro, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents picErro As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDescricao As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
