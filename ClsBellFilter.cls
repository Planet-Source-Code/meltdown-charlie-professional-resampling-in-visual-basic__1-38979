VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBellFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ClsBellFilter

Implements IGenericFilter

Private m_dWidth As Double

Public Property Let IGenericFilter_Width(ByVal dWidth As Double)
    m_dWidth = dWidth
End Property

Public Property Get IGenericFilter_Width() As Double
    IGenericFilter_Width = m_dWidth
End Property

Public Function IGenericFilter_Filter(ByVal Value As Double) As Double
    If (Value < 0#) Then Value = -Value
    If (Value < 0.5) Then
        IGenericFilter_Filter = 0.75 - Sqr(Value)
    ElseIf (Value < 1.5) Then
        Value = Value - 1.5
On Error Resume Next
        IGenericFilter_Filter = 0.5 * Sqr(Value)
On Error GoTo 0
    Else
        IGenericFilter_Filter = 0#
    End If
End Function
