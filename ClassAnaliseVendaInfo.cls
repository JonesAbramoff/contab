VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAnaliseVendaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public iFilialEmpresa As Integer
Public lCliente As Long
Public iFilial As Integer
Public dtDataEmissao As Date
Public iTabelaPreco As Integer
Public iVendedor As Integer
Public iMoeda As Integer

Public colItens As New Collection

