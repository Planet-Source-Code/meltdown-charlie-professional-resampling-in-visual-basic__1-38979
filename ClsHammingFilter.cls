VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsHammingFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ClsHammingFilter
Implements IGenericFilter

Private m_dWidth As Double

Public Property Let IGenericFilter_Width(ByVal dWidth As Double)
    m_dWidth = dWidth
End Property

Public Property Get IGenericFilter_Width() As Double
    IGenericFilter_Width = m_dWidth
End Property

Public Function IGenericFilter_Filter(ByVal dVal As Double) As Double
    Dim dWindow As Double
    Dim dSinc As Double
    
    If (Abs(dVal) > m_dWidth) Then
        IGenericFilter_Filter = 0#
        Exit Function
    End If
    dWindow = 0.54 + 0.46 * Cos(FILTER_2PI * dVal)
    If dVal = 0 Then
        dSinc = 1#
    Else
        dSinc = Sin(FILTER_PI * dVal) / (FILTER_PI * dVal)
    End If
    IGenericFilter_Filter = dWindow * dSinc
End Function

Private Sub Class_Initialize()
    m_dWidth = 0.5
End Sub

