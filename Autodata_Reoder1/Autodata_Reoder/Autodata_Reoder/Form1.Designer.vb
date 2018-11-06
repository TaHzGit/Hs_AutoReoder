<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Auto_Reoder
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.lbxx = New System.Windows.Forms.Label()
        Me.btnRe = New System.Windows.Forms.Button()
        Me.btnRt = New System.Windows.Forms.Button()
        Me.btnPy = New System.Windows.Forms.Button()
        Me.btnPK = New System.Windows.Forms.Button()
        Me.btnSw = New System.Windows.Forms.Button()
        Me.lbSrv = New System.Windows.Forms.Label()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.btnNr = New System.Windows.Forms.Button()
        Me.lbcc = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'lbxx
        '
        Me.lbxx.AutoSize = True
        Me.lbxx.Location = New System.Drawing.Point(134, 22)
        Me.lbxx.Name = "lbxx"
        Me.lbxx.Size = New System.Drawing.Size(0, 13)
        Me.lbxx.TabIndex = 0
        '
        'btnRe
        '
        Me.btnRe.Location = New System.Drawing.Point(142, 38)
        Me.btnRe.Name = "btnRe"
        Me.btnRe.Size = New System.Drawing.Size(75, 23)
        Me.btnRe.TabIndex = 1
        Me.btnRe.Text = "HSRE"
        Me.btnRe.UseVisualStyleBackColor = True
        '
        'btnRt
        '
        Me.btnRt.Location = New System.Drawing.Point(226, 38)
        Me.btnRt.Name = "btnRt"
        Me.btnRt.Size = New System.Drawing.Size(75, 23)
        Me.btnRt.TabIndex = 2
        Me.btnRt.Text = "HSRT"
        Me.btnRt.UseVisualStyleBackColor = True
        '
        'btnPy
        '
        Me.btnPy.Location = New System.Drawing.Point(404, 38)
        Me.btnPy.Name = "btnPy"
        Me.btnPy.Size = New System.Drawing.Size(75, 23)
        Me.btnPy.TabIndex = 3
        Me.btnPy.Text = "HSPY"
        Me.btnPy.UseVisualStyleBackColor = True
        '
        'btnPK
        '
        Me.btnPK.Location = New System.Drawing.Point(312, 38)
        Me.btnPK.Name = "btnPK"
        Me.btnPK.Size = New System.Drawing.Size(75, 23)
        Me.btnPK.TabIndex = 4
        Me.btnPK.Text = "HSPK"
        Me.btnPK.UseVisualStyleBackColor = True
        '
        'btnSw
        '
        Me.btnSw.Location = New System.Drawing.Point(492, 38)
        Me.btnSw.Name = "btnSw"
        Me.btnSw.Size = New System.Drawing.Size(75, 23)
        Me.btnSw.TabIndex = 5
        Me.btnSw.Text = "HSSW"
        Me.btnSw.UseVisualStyleBackColor = True
        '
        'lbSrv
        '
        Me.lbSrv.AllowDrop = True
        Me.lbSrv.AutoSize = True
        Me.lbSrv.Location = New System.Drawing.Point(363, 22)
        Me.lbSrv.Name = "lbSrv"
        Me.lbSrv.Size = New System.Drawing.Size(0, 13)
        Me.lbSrv.TabIndex = 6
        '
        'txtMsg
        '
        Me.txtMsg.Location = New System.Drawing.Point(9, 99)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.Size = New System.Drawing.Size(872, 438)
        Me.txtMsg.TabIndex = 7
        '
        'btnNr
        '
        Me.btnNr.Location = New System.Drawing.Point(581, 38)
        Me.btnNr.Name = "btnNr"
        Me.btnNr.Size = New System.Drawing.Size(75, 23)
        Me.btnNr.TabIndex = 8
        Me.btnNr.Text = "HSNR"
        Me.btnNr.UseVisualStyleBackColor = True
        '
        'lbcc
        '
        Me.lbcc.AllowDrop = True
        Me.lbcc.AutoSize = True
        Me.lbcc.Location = New System.Drawing.Point(62, 73)
        Me.lbcc.Name = "lbcc"
        Me.lbcc.Size = New System.Drawing.Size(0, 13)
        Me.lbcc.TabIndex = 9
        '
        'Auto_Reoder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(887, 549)
        Me.Controls.Add(Me.lbcc)
        Me.Controls.Add(Me.btnNr)
        Me.Controls.Add(Me.txtMsg)
        Me.Controls.Add(Me.lbSrv)
        Me.Controls.Add(Me.btnSw)
        Me.Controls.Add(Me.btnPK)
        Me.Controls.Add(Me.btnPy)
        Me.Controls.Add(Me.btnRt)
        Me.Controls.Add(Me.btnRe)
        Me.Controls.Add(Me.lbxx)
        Me.Name = "Auto_Reoder"
        Me.Text = "Auto_Reoder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Timer1 As Timer
    Friend WithEvents lbxx As Label
    Friend WithEvents btnRe As Button
    Friend WithEvents btnRt As Button
    Friend WithEvents btnPy As Button
    Friend WithEvents btnPK As Button
    Friend WithEvents btnSw As Button
    Friend WithEvents lbSrv As Label
    Friend WithEvents txtMsg As TextBox
    Friend WithEvents btnNr As Button
    Friend WithEvents lbcc As Label
End Class
