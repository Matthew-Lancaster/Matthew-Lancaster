<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtFile = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.chkConfirmDelete = New System.Windows.Forms.CheckBox
        Me.btnEmpty = New System.Windows.Forms.Button
        Me.chkProgress = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.radDeletePermanently = New System.Windows.Forms.RadioButton
        Me.radSendToWastebasket = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkConfirmEmpty = New System.Windows.Forms.CheckBox
        Me.chkPlaySound = New System.Windows.Forms.CheckBox
        Me.tmrCheckBin = New System.Windows.Forms.Timer(Me.components)
        Me.lblNumItems = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(26, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "File:"
        '
        'txtFile
        '
        Me.txtFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFile.Location = New System.Drawing.Point(40, 8)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(438, 20)
        Me.txtFile.TabIndex = 0
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnDelete.Location = New System.Drawing.Point(112, 72)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(112, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "Delete File"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'chkConfirmDelete
        '
        Me.chkConfirmDelete.AutoSize = True
        Me.chkConfirmDelete.Checked = True
        Me.chkConfirmDelete.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkConfirmDelete.Location = New System.Drawing.Point(24, 56)
        Me.chkConfirmDelete.Name = "chkConfirmDelete"
        Me.chkConfirmDelete.Size = New System.Drawing.Size(61, 17)
        Me.chkConfirmDelete.TabIndex = 2
        Me.chkConfirmDelete.Text = "Confirm"
        Me.chkConfirmDelete.UseVisualStyleBackColor = True
        '
        'btnEmpty
        '
        Me.btnEmpty.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btnEmpty.Location = New System.Drawing.Point(112, 112)
        Me.btnEmpty.Name = "btnEmpty"
        Me.btnEmpty.Size = New System.Drawing.Size(112, 32)
        Me.btnEmpty.TabIndex = 4
        Me.btnEmpty.Text = "Empty Wastebasket"
        Me.btnEmpty.UseVisualStyleBackColor = True
        '
        'chkProgress
        '
        Me.chkProgress.AutoSize = True
        Me.chkProgress.Checked = True
        Me.chkProgress.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkProgress.Location = New System.Drawing.Point(24, 24)
        Me.chkProgress.Name = "chkProgress"
        Me.chkProgress.Size = New System.Drawing.Size(97, 17)
        Me.chkProgress.TabIndex = 5
        Me.chkProgress.Text = "Show Progress"
        Me.chkProgress.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.radDeletePermanently)
        Me.GroupBox1.Controls.Add(Me.radSendToWastebasket)
        Me.GroupBox1.Controls.Add(Me.chkConfirmDelete)
        Me.GroupBox1.Controls.Add(Me.btnDelete)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 120)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Delete File"
        '
        'radDeletePermanently
        '
        Me.radDeletePermanently.AutoSize = True
        Me.radDeletePermanently.Location = New System.Drawing.Point(176, 24)
        Me.radDeletePermanently.Name = "radDeletePermanently"
        Me.radDeletePermanently.Size = New System.Drawing.Size(117, 17)
        Me.radDeletePermanently.TabIndex = 4
        Me.radDeletePermanently.Text = "Delete Permanently"
        Me.radDeletePermanently.UseVisualStyleBackColor = True
        '
        'radSendToWastebasket
        '
        Me.radSendToWastebasket.AutoSize = True
        Me.radSendToWastebasket.Checked = True
        Me.radSendToWastebasket.Location = New System.Drawing.Point(24, 24)
        Me.radSendToWastebasket.Name = "radSendToWastebasket"
        Me.radSendToWastebasket.Size = New System.Drawing.Size(128, 17)
        Me.radSendToWastebasket.TabIndex = 3
        Me.radSendToWastebasket.TabStop = True
        Me.radSendToWastebasket.Text = "Send to Wastebasket"
        Me.radSendToWastebasket.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkConfirmEmpty)
        Me.GroupBox2.Controls.Add(Me.chkPlaySound)
        Me.GroupBox2.Controls.Add(Me.chkProgress)
        Me.GroupBox2.Controls.Add(Me.btnEmpty)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 168)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(336, 160)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Empty Wastebasket"
        '
        'chkConfirmEmpty
        '
        Me.chkConfirmEmpty.AutoSize = True
        Me.chkConfirmEmpty.Checked = True
        Me.chkConfirmEmpty.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkConfirmEmpty.Location = New System.Drawing.Point(24, 88)
        Me.chkConfirmEmpty.Name = "chkConfirmEmpty"
        Me.chkConfirmEmpty.Size = New System.Drawing.Size(61, 17)
        Me.chkConfirmEmpty.TabIndex = 8
        Me.chkConfirmEmpty.Text = "Confirm"
        Me.chkConfirmEmpty.UseVisualStyleBackColor = True
        '
        'chkPlaySound
        '
        Me.chkPlaySound.AutoSize = True
        Me.chkPlaySound.Checked = True
        Me.chkPlaySound.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPlaySound.Location = New System.Drawing.Point(24, 56)
        Me.chkPlaySound.Name = "chkPlaySound"
        Me.chkPlaySound.Size = New System.Drawing.Size(80, 17)
        Me.chkPlaySound.TabIndex = 6
        Me.chkPlaySound.Text = "Play Sound"
        Me.chkPlaySound.UseVisualStyleBackColor = True
        '
        'tmrCheckBin
        '
        Me.tmrCheckBin.Enabled = True
        Me.tmrCheckBin.Interval = 1000
        '
        'lblNumItems
        '
        Me.lblNumItems.AutoSize = True
        Me.lblNumItems.Location = New System.Drawing.Point(8, 336)
        Me.lblNumItems.Name = "lblNumItems"
        Me.lblNumItems.Size = New System.Drawing.Size(0, 13)
        Me.lblNumItems.TabIndex = 8
        '
        'Form1
        '
        Me.AcceptButton = Me.btnDelete
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(490, 364)
        Me.Controls.Add(Me.lblNumItems)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtFile)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents chkConfirmDelete As System.Windows.Forms.CheckBox
    Friend WithEvents btnEmpty As System.Windows.Forms.Button
    Friend WithEvents chkProgress As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents radDeletePermanently As System.Windows.Forms.RadioButton
    Friend WithEvents radSendToWastebasket As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkPlaySound As System.Windows.Forms.CheckBox
    Friend WithEvents chkConfirmEmpty As System.Windows.Forms.CheckBox
    Friend WithEvents tmrCheckBin As System.Windows.Forms.Timer
    Friend WithEvents lblNumItems As System.Windows.Forms.Label

End Class
