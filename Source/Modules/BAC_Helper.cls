VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Friend Function getPositionString(ByVal ThisPosition As chPosition) As String

    Select Case ThisPosition
        Case chPosition.posTop
            getPositionString = "'top'"
        Case chPosition.posLeft
            getPositionString = "'left'"
        Case chPosition.posBottom
            getPositionString = "'bottom'"
        Case chPosition.posRight
            getPositionString = "'right'"
        Case Else
            getPositionString = "'top'"
    End Select

End Function

