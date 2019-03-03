Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class Calculator
	Inherits System.Windows.Forms.Form
	' ------------------------------------------------------------------------
	'               Copyright (C) 1994 Microsoft Corporation
	'
	' You have a royalty-free right to use, modify, reproduce and distribute
	' the Sample Application Files (and/or any modified version) in any way
	' you find useful, provided that you agree that Microsoft has no warranty,
	' obligations or liability for any Sample Application Files.
	' ------------------------------------------------------------------------
	Dim Op1, Op2 As Object ' Previously input operand.
	Dim DecimalFlag As Short ' Decimal point present yet?
	Dim NumOps As Short ' Number of operands.
	Dim LastInput As Object ' Indicate type of last keypress event.
	Dim OpFlag As Object ' Indicate pending operation.
	Dim TempReadout As Object
	
	' Click event procedure for C (cancel) key.
	' Reset the display and initializes variables.
	Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
		Readout.Text = VB6.Format(0, "0.")
		'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Op1 = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Op2 = 0
		Calculator_Load(Me, New System.EventArgs())
	End Sub
	
	' Click event procedure for CE (cancel entry) key.
	Private Sub CancelEntry_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CancelEntry.Click
		Readout.Text = VB6.Format(0, "0.")
		DecimalFlag = False
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastInput = "CE"
	End Sub
	
	' Click event procedure for decimal point (.) key.
	' If last keypress was an operator, initialize
	' readout to "0." Otherwise, append a decimal
	' point to the display.
	Private Sub Decimal_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Decimal_Renamed.Click
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LastInput = "NEG" Then
			Readout.Text = VB6.Format(0, "-0.")
			'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf LastInput <> "NUMS" Then 
			Readout.Text = VB6.Format(0, "0.")
		End If
		DecimalFlag = True
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastInput = "NUMS"
	End Sub
	
	' Initialization routine for the form.
	' Set all variables to initial values.
	Private Sub Calculator_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		DecimalFlag = False
		NumOps = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastInput = "NONE"
		'UPGRADE_WARNING: Couldn't resolve default property of object OpFlag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		OpFlag = " "
		Readout.Text = VB6.Format(0, "0.")
		'Decimal.Caption = Format(0, ".")
		
		'End
		
	End Sub
	
	' Click event procedure for number keys (0-9).
	' Append new number to the number in the display.
	Private Sub Number_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Number.Click
		Dim Index As Short = Number.GetIndex(eventSender)
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LastInput <> "NUMS" Then
			Readout.Text = VB6.Format(0, ".")
			DecimalFlag = False
		End If
		If DecimalFlag Then
			Readout.Text = Readout.Text & Number(Index).Text
		Else
			Readout.Text = VB.Left(Readout.Text, InStr(Readout.Text, VB6.Format(0, ".")) - 1) & Number(Index).Text & VB6.Format(0, ".")
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LastInput = "NEG" Then Readout.Text = "-" & Readout.Text
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastInput = "NUMS"
	End Sub
	
	' Click event procedure for operator keys (+, -, x, /, =).
	' If the immediately preceeding keypress was part of a
	' number, increments NumOps. If one operand is present,
	' set Op1. If two are present, set Op1 equal to the
	' result of the operation on Op1 and the current
	' input string, and display the result.
	Private Sub Operator_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Operator_Renamed.Click
		Dim Index As Short = Operator_Renamed.GetIndex(eventSender)
		'UPGRADE_WARNING: Couldn't resolve default property of object TempReadout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempReadout = Readout.Text
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LastInput = "NUMS" Then
			NumOps = NumOps + 1
		End If
		Select Case NumOps
			Case 0
				'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Operator_Renamed(Index).Text = "-" And LastInput <> "NEG" Then
					Readout.Text = "-" & Readout.Text
					'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LastInput = "NEG"
				End If
			Case 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Op1 = Readout.Text
				'UPGRADE_WARNING: Couldn't resolve default property of object OpFlag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Operator_Renamed(Index).Text = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
					Readout.Text = "-"
					'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					LastInput = "NEG"
				End If
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object TempReadout. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Op2 = TempReadout
				Select Case OpFlag
					Case "+"
						'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Op1 = CDbl(Op1) + CDbl(Op2)
					Case "-"
						'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Op1 = CDbl(Op1) - CDbl(Op2)
					Case "X"
						'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Op1 = CDbl(Op1) * CDbl(Op2)
					Case "/"
						'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Op2 = 0 Then
							MsgBox("Can't divide by zero", 48, "Calculator")
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Op1 = CDbl(Op1) / CDbl(Op2)
						End If
					Case "="
						'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Op1 = CDbl(Op2)
					Case "%"
						'UPGRADE_WARNING: Couldn't resolve default property of object Op2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Op1 = CDbl(Op1) * CDbl(Op2)
				End Select
				'UPGRADE_WARNING: Couldn't resolve default property of object Op1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Readout.Text = Op1
				NumOps = 1
		End Select
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LastInput <> "NEG" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LastInput = "OPS"
			'UPGRADE_WARNING: Couldn't resolve default property of object OpFlag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OpFlag = Operator_Renamed(Index).Text
		End If
	End Sub
	
	' Click event procedure for percent key (%).
	' Compute and display a percentage of the first operand.
	Private Sub Percent_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Percent.Click
		Readout.Text = CStr(CDbl(Readout.Text) / 100)
		'UPGRADE_WARNING: Couldn't resolve default property of object LastInput. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LastInput = "Ops"
		'UPGRADE_WARNING: Couldn't resolve default property of object OpFlag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		OpFlag = "%"
		NumOps = NumOps + 1
		DecimalFlag = True
	End Sub
End Class