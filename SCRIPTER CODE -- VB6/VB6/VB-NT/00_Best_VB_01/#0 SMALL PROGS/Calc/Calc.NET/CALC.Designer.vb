<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Calculator
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _Number_7 As System.Windows.Forms.Button
	Public WithEvents _Number_8 As System.Windows.Forms.Button
	Public WithEvents _Number_9 As System.Windows.Forms.Button
	Public WithEvents Cancel As System.Windows.Forms.Button
	Public WithEvents CancelEntry As System.Windows.Forms.Button
	Public WithEvents _Number_4 As System.Windows.Forms.Button
	Public WithEvents _Number_5 As System.Windows.Forms.Button
	Public WithEvents _Number_6 As System.Windows.Forms.Button
	Public WithEvents _Operator_1 As System.Windows.Forms.Button
	Public WithEvents _Operator_3 As System.Windows.Forms.Button
	Public WithEvents _Number_1 As System.Windows.Forms.Button
	Public WithEvents _Number_2 As System.Windows.Forms.Button
	Public WithEvents _Number_3 As System.Windows.Forms.Button
	Public WithEvents _Operator_2 As System.Windows.Forms.Button
	Public WithEvents _Operator_0 As System.Windows.Forms.Button
	Public WithEvents _Number_0 As System.Windows.Forms.Button
	Public WithEvents Decimal_Renamed As System.Windows.Forms.Button
	Public WithEvents _Operator_4 As System.Windows.Forms.Button
	Public WithEvents Percent As System.Windows.Forms.Button
	Public WithEvents Readout As System.Windows.Forms.Label
	Public WithEvents Number As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents Operator_Renamed As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Calculator))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._Number_7 = New System.Windows.Forms.Button
		Me._Number_8 = New System.Windows.Forms.Button
		Me._Number_9 = New System.Windows.Forms.Button
		Me.Cancel = New System.Windows.Forms.Button
		Me.CancelEntry = New System.Windows.Forms.Button
		Me._Number_4 = New System.Windows.Forms.Button
		Me._Number_5 = New System.Windows.Forms.Button
		Me._Number_6 = New System.Windows.Forms.Button
		Me._Operator_1 = New System.Windows.Forms.Button
		Me._Operator_3 = New System.Windows.Forms.Button
		Me._Number_1 = New System.Windows.Forms.Button
		Me._Number_2 = New System.Windows.Forms.Button
		Me._Number_3 = New System.Windows.Forms.Button
		Me._Operator_2 = New System.Windows.Forms.Button
		Me._Operator_0 = New System.Windows.Forms.Button
		Me._Number_0 = New System.Windows.Forms.Button
		Me.Decimal_Renamed = New System.Windows.Forms.Button
		Me._Operator_4 = New System.Windows.Forms.Button
		Me.Percent = New System.Windows.Forms.Button
		Me.Readout = New System.Windows.Forms.Label
		Me.Number = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.Operator_Renamed = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Number, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Operator_Renamed, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Calculator"
		Me.ClientSize = New System.Drawing.Size(218, 198)
		Me.Location = New System.Drawing.Point(172, 99)
		Me.Icon = CType(resources.GetObject("Calculator.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Calculator"
		Me._Number_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_7.Text = "7"
		Me._Number_7.Size = New System.Drawing.Size(32, 32)
		Me._Number_7.Location = New System.Drawing.Point(8, 40)
		Me._Number_7.TabIndex = 7
		Me._Number_7.BackColor = System.Drawing.SystemColors.Control
		Me._Number_7.CausesValidation = True
		Me._Number_7.Enabled = True
		Me._Number_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_7.TabStop = True
		Me._Number_7.Name = "_Number_7"
		Me._Number_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_8.Text = "8"
		Me._Number_8.Size = New System.Drawing.Size(32, 32)
		Me._Number_8.Location = New System.Drawing.Point(48, 40)
		Me._Number_8.TabIndex = 8
		Me._Number_8.BackColor = System.Drawing.SystemColors.Control
		Me._Number_8.CausesValidation = True
		Me._Number_8.Enabled = True
		Me._Number_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_8.TabStop = True
		Me._Number_8.Name = "_Number_8"
		Me._Number_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_9.Text = "9"
		Me._Number_9.Size = New System.Drawing.Size(32, 32)
		Me._Number_9.Location = New System.Drawing.Point(88, 40)
		Me._Number_9.TabIndex = 9
		Me._Number_9.BackColor = System.Drawing.SystemColors.Control
		Me._Number_9.CausesValidation = True
		Me._Number_9.Enabled = True
		Me._Number_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_9.TabStop = True
		Me._Number_9.Name = "_Number_9"
		Me.Cancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Cancel.Text = "C"
		Me.Cancel.Size = New System.Drawing.Size(32, 32)
		Me.Cancel.Location = New System.Drawing.Point(136, 40)
		Me.Cancel.TabIndex = 10
		Me.Cancel.BackColor = System.Drawing.SystemColors.Control
		Me.Cancel.CausesValidation = True
		Me.Cancel.Enabled = True
		Me.Cancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Cancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Cancel.TabStop = True
		Me.Cancel.Name = "Cancel"
		Me.CancelEntry.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelEntry.Text = "CE"
		Me.CancelEntry.Size = New System.Drawing.Size(32, 32)
		Me.CancelEntry.Location = New System.Drawing.Point(176, 40)
		Me.CancelEntry.TabIndex = 11
		Me.CancelEntry.BackColor = System.Drawing.SystemColors.Control
		Me.CancelEntry.CausesValidation = True
		Me.CancelEntry.Enabled = True
		Me.CancelEntry.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CancelEntry.Cursor = System.Windows.Forms.Cursors.Default
		Me.CancelEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CancelEntry.TabStop = True
		Me.CancelEntry.Name = "CancelEntry"
		Me._Number_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_4.Text = "4"
		Me._Number_4.Size = New System.Drawing.Size(32, 32)
		Me._Number_4.Location = New System.Drawing.Point(8, 80)
		Me._Number_4.TabIndex = 4
		Me._Number_4.BackColor = System.Drawing.SystemColors.Control
		Me._Number_4.CausesValidation = True
		Me._Number_4.Enabled = True
		Me._Number_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_4.TabStop = True
		Me._Number_4.Name = "_Number_4"
		Me._Number_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_5.Text = "5"
		Me._Number_5.Size = New System.Drawing.Size(32, 32)
		Me._Number_5.Location = New System.Drawing.Point(48, 80)
		Me._Number_5.TabIndex = 5
		Me._Number_5.BackColor = System.Drawing.SystemColors.Control
		Me._Number_5.CausesValidation = True
		Me._Number_5.Enabled = True
		Me._Number_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_5.TabStop = True
		Me._Number_5.Name = "_Number_5"
		Me._Number_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_6.Text = "6"
		Me._Number_6.Size = New System.Drawing.Size(32, 32)
		Me._Number_6.Location = New System.Drawing.Point(88, 80)
		Me._Number_6.TabIndex = 6
		Me._Number_6.BackColor = System.Drawing.SystemColors.Control
		Me._Number_6.CausesValidation = True
		Me._Number_6.Enabled = True
		Me._Number_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_6.TabStop = True
		Me._Number_6.Name = "_Number_6"
		Me._Operator_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Operator_1.Text = "+"
		Me._Operator_1.Size = New System.Drawing.Size(32, 32)
		Me._Operator_1.Location = New System.Drawing.Point(136, 80)
		Me._Operator_1.TabIndex = 12
		Me._Operator_1.BackColor = System.Drawing.SystemColors.Control
		Me._Operator_1.CausesValidation = True
		Me._Operator_1.Enabled = True
		Me._Operator_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Operator_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Operator_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Operator_1.TabStop = True
		Me._Operator_1.Name = "_Operator_1"
		Me._Operator_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Operator_3.Text = "-"
		Me._Operator_3.Size = New System.Drawing.Size(32, 32)
		Me._Operator_3.Location = New System.Drawing.Point(176, 80)
		Me._Operator_3.TabIndex = 13
		Me._Operator_3.BackColor = System.Drawing.SystemColors.Control
		Me._Operator_3.CausesValidation = True
		Me._Operator_3.Enabled = True
		Me._Operator_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Operator_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Operator_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Operator_3.TabStop = True
		Me._Operator_3.Name = "_Operator_3"
		Me._Number_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_1.Text = "1"
		Me._Number_1.Size = New System.Drawing.Size(32, 32)
		Me._Number_1.Location = New System.Drawing.Point(8, 120)
		Me._Number_1.TabIndex = 1
		Me._Number_1.BackColor = System.Drawing.SystemColors.Control
		Me._Number_1.CausesValidation = True
		Me._Number_1.Enabled = True
		Me._Number_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_1.TabStop = True
		Me._Number_1.Name = "_Number_1"
		Me._Number_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_2.Text = "2"
		Me._Number_2.Size = New System.Drawing.Size(32, 32)
		Me._Number_2.Location = New System.Drawing.Point(48, 120)
		Me._Number_2.TabIndex = 2
		Me._Number_2.BackColor = System.Drawing.SystemColors.Control
		Me._Number_2.CausesValidation = True
		Me._Number_2.Enabled = True
		Me._Number_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_2.TabStop = True
		Me._Number_2.Name = "_Number_2"
		Me._Number_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_3.Text = "3"
		Me._Number_3.Size = New System.Drawing.Size(32, 32)
		Me._Number_3.Location = New System.Drawing.Point(88, 120)
		Me._Number_3.TabIndex = 3
		Me._Number_3.BackColor = System.Drawing.SystemColors.Control
		Me._Number_3.CausesValidation = True
		Me._Number_3.Enabled = True
		Me._Number_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_3.TabStop = True
		Me._Number_3.Name = "_Number_3"
		Me._Operator_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Operator_2.Text = "X"
		Me._Operator_2.Size = New System.Drawing.Size(32, 32)
		Me._Operator_2.Location = New System.Drawing.Point(136, 120)
		Me._Operator_2.TabIndex = 14
		Me._Operator_2.BackColor = System.Drawing.SystemColors.Control
		Me._Operator_2.CausesValidation = True
		Me._Operator_2.Enabled = True
		Me._Operator_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Operator_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Operator_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Operator_2.TabStop = True
		Me._Operator_2.Name = "_Operator_2"
		Me._Operator_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Operator_0.Text = "/"
		Me._Operator_0.Size = New System.Drawing.Size(32, 32)
		Me._Operator_0.Location = New System.Drawing.Point(176, 120)
		Me._Operator_0.TabIndex = 15
		Me._Operator_0.BackColor = System.Drawing.SystemColors.Control
		Me._Operator_0.CausesValidation = True
		Me._Operator_0.Enabled = True
		Me._Operator_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Operator_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Operator_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Operator_0.TabStop = True
		Me._Operator_0.Name = "_Operator_0"
		Me._Number_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Number_0.Text = "0"
		Me._Number_0.Size = New System.Drawing.Size(72, 32)
		Me._Number_0.Location = New System.Drawing.Point(8, 160)
		Me._Number_0.TabIndex = 0
		Me._Number_0.BackColor = System.Drawing.SystemColors.Control
		Me._Number_0.CausesValidation = True
		Me._Number_0.Enabled = True
		Me._Number_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Number_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Number_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Number_0.TabStop = True
		Me._Number_0.Name = "_Number_0"
		Me.Decimal_Renamed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Decimal_Renamed.Text = "."
		Me.Decimal_Renamed.Size = New System.Drawing.Size(32, 32)
		Me.Decimal_Renamed.Location = New System.Drawing.Point(88, 160)
		Me.Decimal_Renamed.TabIndex = 18
		Me.Decimal_Renamed.BackColor = System.Drawing.SystemColors.Control
		Me.Decimal_Renamed.CausesValidation = True
		Me.Decimal_Renamed.Enabled = True
		Me.Decimal_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Decimal_Renamed.Cursor = System.Windows.Forms.Cursors.Default
		Me.Decimal_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Decimal_Renamed.TabStop = True
		Me.Decimal_Renamed.Name = "Decimal_Renamed"
		Me._Operator_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Operator_4.Text = "="
		Me._Operator_4.Size = New System.Drawing.Size(32, 32)
		Me._Operator_4.Location = New System.Drawing.Point(136, 160)
		Me._Operator_4.TabIndex = 16
		Me._Operator_4.BackColor = System.Drawing.SystemColors.Control
		Me._Operator_4.CausesValidation = True
		Me._Operator_4.Enabled = True
		Me._Operator_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Operator_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Operator_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Operator_4.TabStop = True
		Me._Operator_4.Name = "_Operator_4"
		Me.Percent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Percent.Text = "%"
		Me.Percent.Size = New System.Drawing.Size(32, 32)
		Me.Percent.Location = New System.Drawing.Point(176, 160)
		Me.Percent.TabIndex = 17
		Me.Percent.BackColor = System.Drawing.SystemColors.Control
		Me.Percent.CausesValidation = True
		Me.Percent.Enabled = True
		Me.Percent.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Percent.Cursor = System.Windows.Forms.Cursors.Default
		Me.Percent.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Percent.TabStop = True
		Me.Percent.Name = "Percent"
		Me.Readout.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Readout.BackColor = System.Drawing.Color.Yellow
		Me.Readout.Text = "0."
		Me.Readout.ForeColor = System.Drawing.Color.Black
		Me.Readout.Size = New System.Drawing.Size(200, 25)
		Me.Readout.Location = New System.Drawing.Point(8, 7)
		Me.Readout.TabIndex = 19
		Me.Readout.Enabled = True
		Me.Readout.Cursor = System.Windows.Forms.Cursors.Default
		Me.Readout.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Readout.UseMnemonic = True
		Me.Readout.Visible = True
		Me.Readout.AutoSize = False
		Me.Readout.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Readout.Name = "Readout"
		Me.Controls.Add(_Number_7)
		Me.Controls.Add(_Number_8)
		Me.Controls.Add(_Number_9)
		Me.Controls.Add(Cancel)
		Me.Controls.Add(CancelEntry)
		Me.Controls.Add(_Number_4)
		Me.Controls.Add(_Number_5)
		Me.Controls.Add(_Number_6)
		Me.Controls.Add(_Operator_1)
		Me.Controls.Add(_Operator_3)
		Me.Controls.Add(_Number_1)
		Me.Controls.Add(_Number_2)
		Me.Controls.Add(_Number_3)
		Me.Controls.Add(_Operator_2)
		Me.Controls.Add(_Operator_0)
		Me.Controls.Add(_Number_0)
		Me.Controls.Add(Decimal_Renamed)
		Me.Controls.Add(_Operator_4)
		Me.Controls.Add(Percent)
		Me.Controls.Add(Readout)
		Me.Number.SetIndex(_Number_7, CType(7, Short))
		Me.Number.SetIndex(_Number_8, CType(8, Short))
		Me.Number.SetIndex(_Number_9, CType(9, Short))
		Me.Number.SetIndex(_Number_4, CType(4, Short))
		Me.Number.SetIndex(_Number_5, CType(5, Short))
		Me.Number.SetIndex(_Number_6, CType(6, Short))
		Me.Number.SetIndex(_Number_1, CType(1, Short))
		Me.Number.SetIndex(_Number_2, CType(2, Short))
		Me.Number.SetIndex(_Number_3, CType(3, Short))
		Me.Number.SetIndex(_Number_0, CType(0, Short))
		Me.Operator_Renamed.SetIndex(_Operator_1, CType(1, Short))
		Me.Operator_Renamed.SetIndex(_Operator_3, CType(3, Short))
		Me.Operator_Renamed.SetIndex(_Operator_2, CType(2, Short))
		Me.Operator_Renamed.SetIndex(_Operator_0, CType(0, Short))
		Me.Operator_Renamed.SetIndex(_Operator_4, CType(4, Short))
		CType(Me.Operator_Renamed, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Number, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class