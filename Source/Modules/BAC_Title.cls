VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Show As Boolean
Private m_Position As chPosition
Private m_Text As String

' ### Properties

Public Property Let Show(ByVal This_Show As Boolean)
    m_Show = This_Show
End Property
Public Property Get Show() As Boolean
    Show = m_Show
End Property

Public Property Let Position(ByVal This_Position As chPosition)
    m_Position = This_Position
End Property
Public Property Get Position() As chPosition
    Position = m_Position
End Property

Public Property Let Text(ByVal This_Text As String)
    m_Text = This_Text
End Property
Public Property Get Text() As String
    Text = m_Text
End Property

' ### Generator Function

Friend Function getTitle() As String

    'Declarations
    Dim strText As String

    If Me.Show = False Then
        strText = strText & "            title: {" & vbCrLf
        strText = strText & "                display: false," & vbCrLf
        strText = strText & "            }," & vbCrLf
    Else
        strText = strText & "            title: {" & vbCrLf
        strText = strText & "                display: true," & vbCrLf
        strText = strText & "                position: " & BAC.Helper.getPositionString(Me.Position) & "," & vbCrLf
        strText = strText & "                text: '" & Me.Text & "'" & vbCrLf
        strText = strText & "            }," & vbCrLf
    End If
    
    getTitle = strText

End Function

