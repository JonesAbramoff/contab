VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDocContabilIntf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlDoc As Long 'guarda o numero do voucher
Private mvardtData As Date 'guarda a data do voucher
Private mvarobjLote As Object 'guarda o lote associado ao documento
Private mvariUltSeq As Integer 'guarda o ultimo sequencial utilizado do documento

Public Property Set objLote(ByVal vData As Object)
    Set mvarobjLote = vData
End Property

Public Property Get objLote() As Object
    Set objLote = mvarobjLote
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lDoc(ByVal vData As Long)
    mvarlDoc = vData
End Property

Public Property Get lDoc() As Long
    lDoc = mvarlDoc
End Property

Public Function IncluirLcto(ByVal sContaCredito As String, ByVal sContaDebito As String, ByVal sCcustoCredito As String, ByVal sCcustoDebito As String, ByVal dValor As Double, ByVal sHistoricoCredito As String, ByVal sHistoricoDebito As String) As Long
'inclui um par de lançamentos pendentes com os dados passados como parametro

Dim lErro As Long
Dim lDoc As Long
Dim objPeriodo As New ClassPeriodo
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection

On Error GoTo Erro_IncluirLcto

    'Preenche Objeto Lançamento_Cabeçalho
    objLancamento_Cabecalho.iFilialEmpresa = mvarobjLote.iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = mvarobjLote.sOrigem
    objLancamento_Cabecalho.iLote = mvarobjLote.iLote
    objLancamento_Cabecalho.lDoc = mvarlDoc
    objLancamento_Cabecalho.dtData = mvardtData
    
    Set objLancamento_Detalhe = New ClassLancamento_Detalhe
    
    mvariUltSeq = mvariUltSeq + 1
    
    objLancamento_Detalhe.iSeq = mvariUltSeq
        
    objLancamento_Detalhe.sConta = sContaCredito
    
    objLancamento_Detalhe.sCcl = sCcustoCredito
    
    objLancamento_Detalhe.dValor = dValor
    
    objLancamento_Detalhe.sProduto = ""
    
    If dValor = 0 Then gError 92934
    
    objLancamento_Detalhe.iSeqContraPartida = mvariUltSeq + 1

    'Armazena Histórico e Ccl
    objLancamento_Detalhe.sHistorico = sHistoricoCredito
        
    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
    colLancamento_Detalhe.Add objLancamento_Detalhe
    
    Set objLancamento_Detalhe = New ClassLancamento_Detalhe
    
    mvariUltSeq = mvariUltSeq + 1
    
    objLancamento_Detalhe.iSeq = mvariUltSeq
        
    objLancamento_Detalhe.sConta = sContaDebito
    
    objLancamento_Detalhe.sCcl = sCcustoDebito
    
    objLancamento_Detalhe.dValor = -dValor
    
    objLancamento_Detalhe.sProduto = ""
    
    objLancamento_Detalhe.iSeqContraPartida = mvariUltSeq - 1

    'Armazena Histórico
    objLancamento_Detalhe.sHistorico = sHistoricoDebito
        
    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
    colLancamento_Detalhe.Add objLancamento_Detalhe
    
    lErro = CF("Lancamento_Grava", objLancamento_Cabecalho, colLancamento_Detalhe, NAO_EXCLUI_LANCAMENTOS_CONTABEIS)
    If lErro <> SUCESSO Then gError 92940

    IncluirLcto = SUCESSO
    
    Exit Function

Erro_IncluirLcto:

    IncluirLcto = gErr

    Select Case gErr

        Case 92934
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_LANCAMENTO_NAO_PREENCHIDO1", gErr, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.lDoc, objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl, objLancamento_Cabecalho.iLote, objLancamento_Detalhe.iSeq)
    
        Case 92940

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149268)

    End Select
    
    Exit Function

End Function


