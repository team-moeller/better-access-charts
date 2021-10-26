VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAC_Defaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_FontColor As String
Private m_FontFamily As String
Private m_FontSize As Integer


Public Property Let FontColor(ByVal This_FontColor As String)
    m_FontColor = This_FontColor
End Property
Public Property Get FontColor() As String
    FontColor = m_FontColor
End Property

Public Property Let FontFamily(ByVal This_FontFamily As String)
    m_FontFamily = This_FontFamily
End Property
Public Property Get FontFamily() As String
    FontFamily = m_FontFamily
End Property

Public Property Let FontSize(ByVal This_FontSize As Integer)
    m_FontSize = This_FontSize
End Property
Public Property Get FontSize() As Integer
    FontSize = m_FontSize
End Property


