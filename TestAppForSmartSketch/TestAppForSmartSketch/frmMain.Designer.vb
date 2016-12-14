<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.btnHello = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnHello
        '
        Me.btnHello.Font = New System.Drawing.Font("Arial Rounded MT Bold", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHello.Location = New System.Drawing.Point(12, 12)
        Me.btnHello.Name = "btnHello"
        Me.btnHello.Size = New System.Drawing.Size(320, 78)
        Me.btnHello.TabIndex = 0
        Me.btnHello.Text = "Hello!"
        Me.btnHello.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 102)
        Me.Controls.Add(Me.btnHello)
        Me.Name = "frmMain"
        Me.Text = "Hello SmartSketch"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnHello As System.Windows.Forms.Button

End Class
