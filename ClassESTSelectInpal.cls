VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ALTERAÇÕES INPAL
'1 - 08/2000 - Marcio - Incluido o Preço Base na leitura dos ItensPV

Option Explicit

Const STRING_UM_PV_SIGLA = 5

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function Produto_ConvPrecoUMAnalise(ByVal sProduto As String, ByVal sUMOrig As String, ByVal dPrecoUnitarioOrig As Double, dPrecoUnitarioConv As Double) As Long
'converte um preco na unidade sUM no preco na unidade de analise de margem de conribuicao

Dim lErro As Long, sUMConv As String, objProduto As New ClassProduto, dFator As Double

On Error GoTo Erro_Produto_ConvPrecoUMAnalise

    objProduto.sCodigo = sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 124304
    
    sUMConv = "KG" '??? INPAL (o padrao usa a UM de Venda)
    
    'converter o custo p/a unidade necessaria
    If sUMConv <> sUMOrig Then
    
        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUMConv, sUMOrig, dFator)
        If lErro <> SUCESSO Then gError 124305
        
        dPrecoUnitarioConv = dPrecoUnitarioOrig * dFator
        
    Else
        
        dPrecoUnitarioConv = dPrecoUnitarioOrig
    
    End If
    
    Produto_ConvPrecoUMAnalise = SUCESSO
     
    Exit Function
    
Erro_Produto_ConvPrecoUMAnalise:

    Produto_ConvPrecoUMAnalise = gErr
     
    Select Case gErr
          
        Case 124304, 124305
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Public Function NFiscalEntrada_Verifica_Existencia2_Cust(ByVal lErro As Long, ByVal objNFiscal As ClassNFiscal, vbMsgRes As VbMsgBoxResult, Optional ByVal bSemAviso As Boolean = False) As Long

Dim lErro1 As Long
Dim objUsuario As New ClassUsuarios

On Error GoTo Erro_NFiscalEntrada_Verifica_Existencia2_Cust

    'Nota Fiscal Externa Igual
    If lErro = 35384 Then
        If Not bSemAviso Then
            
            objUsuario.sCodUsuario = gsUsuario
            lErro1 = CF("Usuarios_Le", objUsuario)
            If lErro1 <> SUCESSO Then gError 184411
            
            If UCase(gsUsuario) = UCase("supervisor") Or UCase(objUsuario.sCodGrupo) = UCase("decon") Or UCase(objUsuario.sCodGrupo) = UCase("Decon-Fis") Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_EXTERNA_CONTAB", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            Else
                gError 184411
            End If
        End If
    'Nota Fiscal Interna Igual
    ElseIf lErro = 42083 Then
        If Not bSemAviso Then
            
            objUsuario.sCodUsuario = gsUsuario
            lErro1 = CF("Usuarios_Le", objUsuario)
            If lErro1 <> SUCESSO Then gError 184412
            
            If UCase(gsUsuario) = UCase("supervisor") Or UCase(objUsuario.sCodGrupo) = UCase("decon") Or UCase(objUsuario.sCodGrupo) = UCase("Decon-Fis") Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            Else
                gError 184412
            End If
        End If
    End If
    
    NFiscalEntrada_Verifica_Existencia2_Cust = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Verifica_Existencia2_Cust:

    NFiscalEntrada_Verifica_Existencia2_Cust = gErr

    Select Case gErr
    
        Case 184411, 184412, 184413, 184414
            Call Rotina_Erro(vbOKOnly, "ERRO_NFENT_ALT_SUPERVISOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184415)
    
    End Select
    
    Exit Function
    
End Function
