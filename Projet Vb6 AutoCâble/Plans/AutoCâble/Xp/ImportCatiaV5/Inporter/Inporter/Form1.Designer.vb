<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button
        Me.AxSpreadsheet1 = New AxOWC10.AxSpreadsheet
        CType(Me.AxSpreadsheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(398, 564)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(113, 48)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'AxSpreadsheet1
        '
        Me.AxSpreadsheet1.DataSource = Nothing
        Me.AxSpreadsheet1.Enabled = True
        Me.AxSpreadsheet1.Location = New System.Drawing.Point(20, 20)
        Me.AxSpreadsheet1.Name = "AxSpreadsheet1"
        Me.AxSpreadsheet1.OcxState = CType(resources.GetObject("AxSpreadsheet1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxSpreadsheet1.Size = New System.Drawing.Size(961, 515)
        Me.AxSpreadsheet1.TabIndex = 2
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1022, 630)
        Me.Controls.Add(Me.AxSpreadsheet1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.AxSpreadsheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents AxSpreadsheet1 As AxOWC10.AxSpreadsheet

End Class
