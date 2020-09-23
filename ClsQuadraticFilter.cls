VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsQuadraticFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ClsQuadraticFilter
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
    If (x < 0.5) Then IGenericFilter_Filter = (0.75 - x * x): Exit Function
    If (x < 1.5) Then
        x = x - 1.5
        IGenericFilter_Filter = (0.5 * x * x)
        Exit Function
    End If
    IGenericFilter_Filter = (0#)
End Function


