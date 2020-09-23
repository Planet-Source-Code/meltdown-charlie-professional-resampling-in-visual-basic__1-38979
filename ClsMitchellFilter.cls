VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMitchellFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ClsMitchellFilter
Implements IGenericFilter

Private m_dWidth As Double

Public Property Let IGenericFilter_Width(ByVal dWidth As Double)
    m_dWidth = dWidth
End Property

Public Property Get IGenericFilter_Width() As Double
    IGenericFilter_Width = m_dWidth
End Property

Public Function IGenericFilter_Filter(ByVal x As Double) As Double
    Dim b As Double, c As Double
    
    b = 1# / 3#
    c = 1# / 3#
    If (x < 0) Then x = (-x)
    If (x < 1#) Then
        x = ((12# - 9# * b - 6# * c) * (x * x * x)) + ((-18# + 12# * b + 6# * c) * x * x) + (6# - 2# * b)
        IGenericFilter_Filter = x / 6#
        Exit Function
    End If
    If (x < 2#) Then
        x = ((-1# * b - 6# * c) * (x * x * x)) + ((6# * b + 30# * c) * x * x) + ((-12# * b - 48# * c) * x) + (8# * b + 24# * c)
        IGenericFilter_Filter = x / 6#
        Exit Function
    End If
    IGenericFilter_Filter = 0#
End Function
