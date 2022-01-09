Attribute VB_Name = "BAC__Namespace"
'###########################################################################################
'# Copyright (c) 2020 - 2022 Thomas Möller, supported by K.D.Gundermann                    #
'# MIT License  => https://github.com/team-moeller/better-access-charts/blob/main/LICENSE  #
'# Version 2.01.09  published: 09.01.2022                                                  #
'###########################################################################################

Option Compare Database
Option Explicit

Private m_BetterAccessCharts As BAC__Factory

Public Property Get BetterAccessCharts() As BAC__Factory
  If m_BetterAccessCharts Is Nothing Then Set m_BetterAccessCharts = New BAC__Factory
  Set BetterAccessCharts = m_BetterAccessCharts
End Property

Public Property Get BAC() As BAC__Factory
  Set BAC = BetterAccessCharts
End Property


