VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLanczosFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ClsLanczosFilter
Implements IGenericFilter

Private m_dWidth As Double

Public Property Let IGenericFilter_Width(ByVal dWidth As Double)
    m_dWidth = dWidth
End Property

Public Property Get IGenericFilter_Width() As Double
    IGenericFilter_Width = m_dWidth
End Property

Public Function IGenericFilter_Filter(ByVal x As Double) As Double
    If (x < 0) Then x = (-x)
    If (x < 3#) Then
        IGenericFilter_Filter = (Sinc(x) * Sinc(x / 3#))
    Else
        IGenericFilter_Filter = 0#
    End If
End Function

Private Function Sinc(x As Double) As Double
    x = x + (x * FILTER_PI)
    If (x <> 0#) Then
        Sinc = (Sin(x) / x)
    Else
        Sinc = (1#)
    End If
End Function

