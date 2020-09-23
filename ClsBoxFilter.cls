VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBoxFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ClsBoxFilter
Implements IGenericFilter

Private m_dWidth As Double

Public Property Let IGenericFilter_Width(ByVal dWidth As Double)
    m_dWidth = dWidth
End Property

Public Property Get IGenericFilter_Width() As Double
    IGenericFilter_Width = m_dWidth
End Property

Public Function IGenericFilter_Filter(ByVal dVal As Double) As Double
    Dim n As Double
    
    If m_dWidth <> 0 Then n = 1# Else n = 0#
    If Abs(dVal) <= n Then
        IGenericFilter_Filter = 1#
    Else
        IGenericFilter_Filter = 0#
    End If
End Function

Private Sub Class_Initialize()
    m_dWidth = 0.5
End Sub

