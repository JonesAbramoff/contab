VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaRecDig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolBaixaParcRec As colBaixaParcRec
Private mvarobjMovCCI As ClassMovContaCorrente
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function

Public Property Let colBaixaParcRec(ByVal vData As colBaixaParcRec)
    Set mvarcolBaixaParcRec = vData
End Property

Public Property Get colBaixaParcRec() As colBaixaParcRec
    Set colBaixaParcRec = mvarcolBaixaParcRec
End Property

Public Property Set colBaixaParcRec(ByVal vData As colBaixaParcRec)
    Set mvarcolBaixaParcRec = vData
End Property

Public Property Let objMovCCI(ByVal vData As ClassMovContaCorrente)
    Set mvarobjMovCCI = vData
End Property

Public Property Get objMovCCI() As ClassMovContaCorrente
    Set objMovCCI = mvarobjMovCCI
End Property

Public Property Set objMovCCI(ByVal vData As ClassMovContaCorrente)
    Set mvarobjMovCCI = vData
End Property



