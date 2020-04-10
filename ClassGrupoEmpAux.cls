VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGrupoEmpAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iGrupoEmp As Integer 'Grupo Empresarial
Public lEmpresa As Long
Public iFilialEmpresa As Long
Public iNumCmd As Integer
Public lConexao As Long
Public sStringConexao As String

Private mvaralComando(1 To 50) As Long

Public Property Get alComando(ByVal iIndice As Integer) As Long
    alComando = mvaralComando(iIndice)
End Property

Public Property Let alComando(ByVal iIndice As Integer, ByVal lComando As Long)
    mvaralComando(iIndice) = lComando
End Property

