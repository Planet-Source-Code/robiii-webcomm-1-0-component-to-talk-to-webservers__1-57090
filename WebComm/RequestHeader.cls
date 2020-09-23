VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Class to hold response/request headers
Private m_strHeaderName As String
Private m_strValue As String

Private Sub Class_Initialize()
    m_strHeaderName = ""
    m_strValue = ""
End Sub

Public Property Get HeaderName() As String
    HeaderName = m_strHeaderName
End Property

Public Property Let HeaderName(ByVal strName As String)
    m_strHeaderName = strName
End Property

Public Property Get Value() As String
    Value = m_strValue
End Property

Public Property Let Value(ByVal strValue As String)
    m_strValue = strValue
End Property
