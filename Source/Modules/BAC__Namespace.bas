Attribute VB_Name = "BAC__Namespace"
'###########################################################################################
'# Copyright (c) 2020, 2021 Thomas Möller                                                  #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 1.31.12  published: 29.10.2021                                                  #
'###########################################################################################

Option Compare Database
Option Explicit

Private m_BetterAccessCharts As BAC__Factory
'# Version 1.31.12  published: 29.10.2021                                                  #
Public Property Get BetterAccessCharts() As BAC__Factory
  If m_BetterAccessCharts Is Nothing Then Set m_BetterAccessCharts = New BAC__Factory
  Set BetterAccessCharts = m_BetterAccessCharts
End Property

Public Property Get BAC() As BAC__Factory
  Set BAC = BetterAccessCharts
End Property


