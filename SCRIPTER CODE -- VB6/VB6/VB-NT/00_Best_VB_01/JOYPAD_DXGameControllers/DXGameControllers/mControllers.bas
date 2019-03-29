Attribute VB_Name = "mControllers"
' Name:    mControllers
' Author:  Michael A Seel
' Date:    Feb 2004
'
Option Explicit
'---------------------------------------------
'FRMDXAMEJOY
'---------------------------------------------
Public JOY_CONTROLLER_F710_Y_4_BUTTON_PRESSER


Public Enum ControllerElement
   ceUnsupported
   ceBtn01
   ceBtn02
   ceBtn03
   ceBtn04
   ceBtn05
   ceBtn06
   ceBtn07
   ceBtn08
   ceBtn09
   ceBtn10
   ceBtn11
   ceBtn12
   ceBtn13
   ceBtn14
   ceBtn15
   ceBtn16
   ceBtn17
   ceBtn18
   ceBtn19
   ceBtn20
   ceBtn21
   ceBtn22
   ceBtn23
   ceBtn24
   ceBtn25
   ceBtn26
   ceBtn27
   ceBtn28
   ceBtn29
   ceBtn30
   ceBtn31
   ceBtn32
   ceXNeg
   ceXPos
   ceYNeg
   ceYPos
   ceZNeg
   ceZPos
   ceRXNeg
   ceRXPos
   ceRYNeg
   ceRYPos
   ceRZNeg
   ceRZPos
   cePOV0Neg
   cePOV0Pos
   cePOV1Neg
   cePOV1Pos
   cePOV2Neg
   cePOV2Pos
   cePOV3Neg
   cePOV3Pos
   ceSlider0Neg
   ceSlider0Pos
   ceSlider1Neg
   ceSlider1Pos
End Enum

Public Enum ControllerElementStyles
   cesUnsupported
   cesButton
   cesDirLeft
   cesDirRight
   cesDirUp
   cesDirDown
   cesDirOut
   cesDirIn
   cesSliderLower
   cesSliderRaise
   cesPOVDecrement
   cesPOVIncrement
End Enum


Public Function ElementName(ByVal Element As ControllerElement) As String
' Return the display name of the specified Element
   Select Case Element
      Case ceBtn01 To ceBtn32
         ElementName = "Button " & CStr(Element)
      Case ceXNeg To ceRZPos
         ElementName = Split("Left,Right,Up,Down,In,Out,L-Left,L-Right,L-Up,L-Down,L-Out,L-In", ",")(Element - ceXNeg)
      Case cePOV0Neg To cePOV3Pos
         ElementName = Split("POVa Neg,POVa Pos,POVb Neg,POVb Pos,POVc Neg,POVc Pos,POVd Neg,POVd Pos", ",")(Element - cePOV0Neg)
      Case ceSlider0Neg To ceSlider1Pos
         ElementName = Split("Increase1,Decrease1,Increase2,Decrease2", ",")(Element - ceSlider0Neg)
      Case Else
         ElementName = "Unsupported"
   End Select
End Function

Public Function ElementStyle(ByVal Element As ControllerElement) As ControllerElementStyles
' Return the style of the specified Element
   Select Case Element
      Case ceBtn01 To ceBtn32
         ElementStyle = cesButton
      Case ceXNeg, ceRXNeg
         ElementStyle = cesDirLeft
      Case ceXPos, ceRXPos
         ElementStyle = cesDirRight
      Case ceYNeg, ceRYNeg
         ElementStyle = cesDirUp
      Case ceYPos, ceRYPos
         ElementStyle = cesDirDown
      Case ceZNeg, ceRZNeg
         ElementStyle = cesDirOut
      Case ceZPos, ceRZPos
         ElementStyle = cesDirIn
      Case ceSlider0Neg, ceSlider1Neg
         ElementStyle = cesSliderLower
      Case ceSlider0Pos, ceSlider1Pos
         ElementStyle = cesSliderRaise
      Case cePOV0Neg, cePOV1Neg, cePOV2Neg, cePOV3Neg
         ElementStyle = cesPOVDecrement
      Case cePOV0Pos, cePOV1Pos, cePOV2Pos, cePOV3Pos
         ElementStyle = cesPOVIncrement
   End Select
End Function

Public Function ElementStyleName(ByVal Style As ControllerElementStyles) As String
' Return the display name of the specified Style
   If (Style > 0) And (Style < 12) Then
      ElementStyleName = Split("Button,Left,Right,Up,Down,Out,In,Lower,Raise,Decrement,Increment", ",")(Style - 1)
   Else
      ElementStyleName = "Unsupported"
   End If
End Function
