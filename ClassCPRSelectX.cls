VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRSelectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function NFPag_Testa_Repeticao(ByVal objNFsPag As ClassNFsPag) As Long
'Teste se há NotaFiscalPagar (lançada ou baixada)
'com os mesmos Fornecedor, Filial, Número e DataEmissao dentro da faixa PERIODO_EMISSAO
'Função retorna SUCESSO se NÃO há ou se há e deseja prosseguir na INSERÇÃO de nova NotaFiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataEmissao As Date
Dim sSQLAux As String

On Error GoTo Erro_NFPag_Testa_Repeticao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18531

    sSQLAux = "SELECT NumIntDoc, DataEmissao FROM NFsPag WHERE "
        
    If objNFsPag.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
        
    If objNFsPag.dtDataEmissao <> DATA_NULA Then
    
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao > ?  AND DataEmissao < ? AND Status <> ?"
        
        'Seleciona NotaFiscalPagar com os mesmos Fornecedor, Filial, Número e DataEmissao dentro da faixa PERIODO_EMISSAO
        lErro = Comando_Executar(lComando, sSQLAux, lNumIntDoc, dtDataEmissao, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao - PERIODO_EMISSAO, objNFsPag.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO)
        
    Else
    
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND Status <> ?"
        
        'Seleciona NotaFiscalPagar com os mesmos Fornecedor, Filial, Número independente de DataEmissao
        lErro = Comando_Executar(lComando, sSQLAux, lNumIntDoc, dtDataEmissao, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, STATUS_EXCLUIDO)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 18499

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18500
    
    If lErro = AD_SQL_SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NFPAG_MESMO_NUMERO", objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, dtDataEmissao)
        
        If vbMsgRes = vbNo Then Error 18501
    
    End If
           
    sSQLAux = "SELECT NumIntDoc FROM NFsPagBaixadas WHERE "
        
    If objNFsPag.iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        sSQLAux = sSQLAux & " FilialEmpresa <= ? AND "
    Else
        sSQLAux = sSQLAux & " FilialEmpresa > ? AND "
    End If
        
    If objNFsPag.dtDataEmissao <> DATA_NULA Then
    
        sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND DataEmissao > ?  AND DataEmissao < ? AND Status <> ?"
        
        'Seleciona NotaFiscalPagarBaixada com os mesmos Fornecedor, Filial, Número e DataEmissao dentro da faixa PERIODO_EMISSAO
        lErro = Comando_Executar(lComando, sSQLAux, lNumIntDoc, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao - PERIODO_EMISSAO, objNFsPag.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO)
        
    Else
    
         sSQLAux = sSQLAux & " Fornecedor = ? AND Filial = ? AND NumNotaFiscal = ? AND Status <> ?"
         
        'Seleciona NotaFiscalPagarBaixada com os mesmos Fornecedor, Filial, Número independente de DataEmissao
        lErro = Comando_Executar(lComando, sSQLAux, lNumIntDoc, DELTA_FILIALREAL_OFICIAL, objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, STATUS_EXCLUIDO)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 18535

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18536
    
    If lErro = AD_SQL_SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NFPAG_BAIXADA_MESMO_NUMERO", objNFsPag.lFornecedor, objNFsPag.iFilial, objNFsPag.lNumNotaFiscal, objNFsPag.dtDataEmissao)
        
        If vbMsgRes = vbNo Then Error 18537
    
    End If
    
    Call Comando_Fechar(lComando)
    
    NFPag_Testa_Repeticao = SUCESSO
    
    Exit Function
    
Erro_NFPag_Testa_Repeticao:

    NFPag_Testa_Repeticao = Err

    Select Case Err
    
        Case 18499, 18500
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG1", Err, objNFsPag.lNumNotaFiscal)
        
        Case 18531
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 18501, 18537
        
        Case 18535, 18536
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS", Err, objNFsPag.lNumNotaFiscal)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179003)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


