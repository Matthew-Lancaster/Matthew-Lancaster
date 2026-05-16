Option Strict Off
Option Explicit On
Friend Class Form2
	Inherits System.Windows.Forms.Form
	Private Sub Form2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
        'Load(DIALER)
        'Call DIALER.DIALER_Load(Nothing, New System.EventArgs())
		' DIALER.Show
        'Timer1.Enabled = True
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'Me.Close()
	End Sub
End Class