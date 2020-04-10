VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpCoinfoNFCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para contabilizacao automatica da importacao de notas fiscais do sistema coinfo

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjNFiscal As ClassNFiscal
Private mvarobjFormCTBAux As Object

Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const CLIENTE1 As String = "Cliente"
Private Const CLIENTE_CODIGO As String = "Cliente_Codigo"
Private Const FILIAL1 As String = "Filial_Cliente"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const SERIE_NOTA As String = "Serie"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_SAIDA As String = "Data_Saida"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Preco_Unitario"
Private Const PRECO_TOTAL As String = "Preco_Total"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"

'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"
Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const VALOR_TOTAL As String = "Valor_Total"
'###########Almoxarifado############
Private Const PRODUTO_ALMOX As String = "Produto_Almox"
Private Const ALMOX1 As String = "Almoxarifado"
Private Const QUANT_ALOCADA As String = "Quant_Alocada"
Private Const UNIDADE_MED_EST As String = "Unidade_Med_Est"
                            '###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const PIS_DEBITO_TOTAL As String = "PIS_Debito_Total"
Private Const COFINS_DEBITO_TOTAL As String = "COFINS_Debito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
'fim da contabilidade

Public Property Get Controls() As Object
    
Dim objControls As New Collection

'    If mvarobjFormCTBAux Is Nothing Then
'
'        Set mvarobjFormCTBAux = New FormAuxCtbCoinfo
'
'    End If
'
'    Set Controls = mvarobjFormCTBAux.Controls
    Set Controls = objControls
    
End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

Dim lErro As Long, lDoc As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjNFiscal = vParams(0)
    
    lErro = CF("Periodo_Le", gobjNFiscal.dtDataEmissao, objPeriodo)
    If lErro <> SUCESSO Then gError 130631
    
    lErro = CF("Voucher_Automatico_Trans", gobjNFiscal.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "FAT", lDoc)
    If lErro <> SUCESSO Then gError 130630
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro2(Me, "Nfiscal", gobjNFiscal.lNumIntDoc, gobjNFiscal.lCliente, gobjNFiscal.iFilialCli, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjNFiscal.iFilialEmpresa, 0, gobjNFiscal.lNumNotaFiscal, 0, "", 70, "NF Coinfo")
    If lErro <> SUCESSO Then gError 32161
        
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 32161, 130630, 130631
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184002)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
            
        Case DATA_EMISSAO
            objMnemonicoValor.colValor.Add gobjNFiscal.dtDataEmissao
            
        Case VALOR_TOTAL 'Valor líquido das notas
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorTotal
            
        Case CSLL_RETIDO 'Comissão retida nas agências das faturas que estão sendo pagas
            objMnemonicoValor.colValor.Add gobjNFiscal.dTaxaMoeda
            
        Case SERIE_NOTA 'UF do Cliente
        
'            objCliente.lCodigo = gobjNFiscal.lCliente
'
'            'faz a leitura do cliente
'            lErro = CF("Cliente_Le", objCliente)
'            If lErro <> SUCESSO And lErro <> 12293 Then gError 189430
'
'            objEndereco.lCodigo = objCliente.lEndereco
'
'            'Lê o endereço à partir do Código
'            lErro = CF("Endereco_Le", objEndereco)
'            If lErro <> SUCESSO And lErro <> 12309 Then gError 189431
'
'            objMnemonicoValor.colValor.Add objEndereco.sSiglaEstado

            objMnemonicoValor.colValor.Add CStr(gobjNFiscal.iFilialEmpresa)
            
        Case NATUREZA_OP
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacao.sNaturezaOpInterna
             
        Case VALOR_PRODUTOS
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorProdutos
            
        Case ISS_VALOR
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacao.dISSValor

        Case PIS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacao.dPISCredito
        
        Case COFINS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacao.dCOFINSCredito
            
        Case Else
            gError 39695
    
    End Select
    
    Calcula_Mnemonico = SUCESSO
     
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
     
    Select Case gErr
          
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 189430, 189431
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184003)
     
    End Select
     
    Exit Function

End Function

Public Sub Encerrando()
    Set gobjContabAutomatica = Nothing
    Set gobjNFiscal = Nothing
    Set mvarobjFormCTBAux = Nothing
End Sub

Private Sub Class_Terminate()
    Call Encerrando
End Sub

