<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class DIALER
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
	Public WithEvents Timer_TICKER_CHAIN_DOG_SHOW_RUN As System.Windows.Forms.Timer
	Public WithEvents Timer20 As System.Windows.Forms.Timer
	Public WithEvents MMControl9 As AxMCI.AxMMControl
	Public WithEvents Timer_ERROR As System.Windows.Forms.Timer
	Public WithEvents TIMER_FRONT_DOOR As System.Windows.Forms.Timer
	Public WithEvents MSComm_DOOR As AxMSCommLib.AxMSComm
	Public WithEvents TIMER_PIR As System.Windows.Forms.Timer
	Public WithEvents RichTextBox1 As System.Windows.Forms.PictureBox
	Public WithEvents List1 As System.Windows.Forms.ListBox
	Public WithEvents TIMER_1 As System.Windows.Forms.Timer
	Public WithEvents TimerComm4 As System.Windows.Forms.Timer
	Public WithEvents MMControl1 As System.Windows.Forms.PictureBox
	Public WithEvents Dir1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
	Public WithEvents List2 As System.Windows.Forms.ListBox
	Public WithEvents MSComm_PIR As AxMSCommLib.AxMSComm
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DIALER))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Timer_TICKER_CHAIN_DOG_SHOW_RUN = New System.Windows.Forms.Timer(components)
		Me.Timer20 = New System.Windows.Forms.Timer(components)
		Me.MMControl9 = New AxMCI.AxMMControl
		Me.Timer_ERROR = New System.Windows.Forms.Timer(components)
		Me.TIMER_FRONT_DOOR = New System.Windows.Forms.Timer(components)
		Me.MSComm_DOOR = New AxMSCommLib.AxMSComm
		Me.TIMER_PIR = New System.Windows.Forms.Timer(components)
		Me.RichTextBox1 = New System.Windows.Forms.PictureBox
		Me.List1 = New System.Windows.Forms.ListBox
		Me.TIMER_1 = New System.Windows.Forms.Timer(components)
		Me.TimerComm4 = New System.Windows.Forms.Timer(components)
		Me.MMControl1 = New System.Windows.Forms.PictureBox
		Me.Dir1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
		Me.List2 = New System.Windows.Forms.ListBox
		Me.MSComm_PIR = New AxMSCommLib.AxMSComm
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.MMControl9, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MSComm_DOOR, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.MSComm_PIR, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "RS232_LOGGER"
		Me.ClientSize = New System.Drawing.Size(1029, 216)
		Me.Location = New System.Drawing.Point(672, 21)
		Me.Visible = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.Name = "DIALER"
		Me.Timer_TICKER_CHAIN_DOG_SHOW_RUN.Interval = 1
		Me.Timer_TICKER_CHAIN_DOG_SHOW_RUN.Enabled = True
		Me.Timer20.Enabled = False
		Me.Timer20.Interval = 20000
		MMControl9.OcxState = CType(resources.GetObject("MMControl9.OcxState"), System.Windows.Forms.AxHost.State)
		Me.MMControl9.Size = New System.Drawing.Size(189, 44)
		Me.MMControl9.Location = New System.Drawing.Point(772, 129)
		Me.MMControl9.TabIndex = 5
		Me.MMControl9.Visible = False
		Me.MMControl9.Name = "MMControl9"
		Me.Timer_ERROR.Interval = 1000
		Me.Timer_ERROR.Enabled = True
		Me.TIMER_FRONT_DOOR.Interval = 1000
		Me.TIMER_FRONT_DOOR.Enabled = True
		MSComm_DOOR.OcxState = CType(resources.GetObject("MSComm_DOOR.OcxState"), System.Windows.Forms.AxHost.State)
		Me.MSComm_DOOR.Location = New System.Drawing.Point(824, 9)
		Me.MSComm_DOOR.Name = "MSComm_DOOR"
		Me.TIMER_PIR.Interval = 1000
		Me.TIMER_PIR.Enabled = True
		Me.RichTextBox1.Size = New System.Drawing.Size(97, 44)
		Me.RichTextBox1.Location = New System.Drawing.Point(969, 12)
		Me.RichTextBox1.TabIndex = 4
		Me.RichTextBox1.Visible = False
		Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.None
		Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Control
		Me.RichTextBox1.CausesValidation = True
		Me.RichTextBox1.Enabled = True
		Me.RichTextBox1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.RichTextBox1.Cursor = System.Windows.Forms.Cursors.Default
		Me.RichTextBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.RichTextBox1.TabStop = True
		Me.RichTextBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.RichTextBox1.Name = "RichTextBox1"
		Me.List1.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me.List1.Font = New System.Drawing.Font("Lucida Console", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.List1.ForeColor = System.Drawing.Color.White
		Me.List1.Size = New System.Drawing.Size(96, 115)
		Me.List1.Location = New System.Drawing.Point(867, 11)
		Me.List1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
		Me.List1.TabIndex = 3
		Me.List1.Visible = False
		Me.List1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.List1.CausesValidation = True
		Me.List1.Enabled = True
		Me.List1.IntegralHeight = True
		Me.List1.Cursor = System.Windows.Forms.Cursors.Default
		Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List1.Sorted = False
		Me.List1.TabStop = True
		Me.List1.MultiColumn = False
		Me.List1.Name = "List1"
		Me.TIMER_1.Interval = 1000
		Me.TIMER_1.Enabled = True
		Me.TimerComm4.Enabled = False
		Me.TimerComm4.Interval = 2000
		Me.MMControl1.Size = New System.Drawing.Size(148, 46)
		Me.MMControl1.Location = New System.Drawing.Point(968, 72)
		Me.MMControl1.TabIndex = 2
		Me.MMControl1.Visible = False
		Me.MMControl1.Dock = System.Windows.Forms.DockStyle.None
		Me.MMControl1.BackColor = System.Drawing.SystemColors.Control
		Me.MMControl1.CausesValidation = True
		Me.MMControl1.Enabled = True
		Me.MMControl1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.MMControl1.Cursor = System.Windows.Forms.Cursors.Default
		Me.MMControl1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.MMControl1.TabStop = True
		Me.MMControl1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.MMControl1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.MMControl1.Name = "MMControl1"
		Me.Dir1.Size = New System.Drawing.Size(121, 51)
		Me.Dir1.Location = New System.Drawing.Point(392, 360)
		Me.Dir1.TabIndex = 1
		Me.Dir1.Visible = False
		Me.Dir1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Dir1.BackColor = System.Drawing.SystemColors.Window
		Me.Dir1.CausesValidation = True
		Me.Dir1.Enabled = True
		Me.Dir1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Dir1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Dir1.TabStop = True
		Me.Dir1.Name = "Dir1"
		Me.List2.Size = New System.Drawing.Size(313, 71)
		Me.List2.Location = New System.Drawing.Point(72, 360)
		Me.List2.TabIndex = 0
		Me.List2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.List2.BackColor = System.Drawing.SystemColors.Window
		Me.List2.CausesValidation = True
		Me.List2.Enabled = True
		Me.List2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.List2.IntegralHeight = True
		Me.List2.Cursor = System.Windows.Forms.Cursors.Default
		Me.List2.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.List2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List2.Sorted = False
		Me.List2.TabStop = True
		Me.List2.Visible = True
		Me.List2.MultiColumn = False
		Me.List2.Name = "List2"
		MSComm_PIR.OcxState = CType(resources.GetObject("MSComm_PIR.OcxState"), System.Windows.Forms.AxHost.State)
		Me.MSComm_PIR.Location = New System.Drawing.Point(784, 8)
		Me.MSComm_PIR.Name = "MSComm_PIR"
		Me.Label6.BackColor = System.Drawing.Color.White
		Me.Label6.Text = "DOOR"
		Me.Label6.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Size = New System.Drawing.Size(495, 28)
		Me.Label6.Location = New System.Drawing.Point(7, 175)
		Me.Label6.TabIndex = 11
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label5.BackColor = System.Drawing.Color.White
		Me.Label5.Text = "DOOR"
		Me.Label5.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Size = New System.Drawing.Size(495, 28)
		Me.Label5.Location = New System.Drawing.Point(7, 141)
		Me.Label5.TabIndex = 10
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.BackColor = System.Drawing.Color.White
		Me.Label4.Text = "DOOR"
		Me.Label4.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(495, 28)
		Me.Label4.Location = New System.Drawing.Point(7, 108)
		Me.Label4.TabIndex = 9
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.BackColor = System.Drawing.Color.White
		Me.Label3.Text = "PIR"
		Me.Label3.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(495, 28)
		Me.Label3.Location = New System.Drawing.Point(7, 75)
		Me.Label3.TabIndex = 8
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.BackColor = System.Drawing.Color.White
		Me.Label2.Text = "PIR"
		Me.Label2.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(495, 28)
		Me.Label2.Location = New System.Drawing.Point(7, 42)
		Me.Label2.TabIndex = 7
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.White
		Me.Label1.Text = "PIR"
		Me.Label1.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(495, 28)
		Me.Label1.Location = New System.Drawing.Point(7, 9)
		Me.Label1.TabIndex = 6
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(MMControl9)
		Me.Controls.Add(MSComm_DOOR)
		Me.Controls.Add(RichTextBox1)
		Me.Controls.Add(List1)
		Me.Controls.Add(MMControl1)
		Me.Controls.Add(Dir1)
		Me.Controls.Add(List2)
		Me.Controls.Add(MSComm_PIR)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		CType(Me.MSComm_PIR, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MSComm_DOOR, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.MMControl9, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class