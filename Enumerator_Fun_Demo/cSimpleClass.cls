VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSimpleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit
  ' demo project showing how to use a typelibrary to access the
  ' Skip, Reset and Clone methods of a NewEnum enumerator object
  ' by Bryan Stafford of New Vision Software® - newvision@mvps.org
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.

  Private m_sName As String
  Private m_nNumber As Long

Public Property Get Name() As String
  Name = m_sName
End Property
Public Property Let Name(ByVal sNewVal$)
  m_sName = sNewVal
End Property

Public Property Get Number() As Long
  Number = m_nNumber
End Property
Public Property Let Number(ByVal nNewVal&)
  m_nNumber = nNewVal
End Property

