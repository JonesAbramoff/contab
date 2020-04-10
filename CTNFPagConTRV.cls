VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFPagConTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BotaoExcluir_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag
Dim vbMsgRes As VbMsgBoxResult
Dim iFilialAux As Integer

On Error GoTo Erro_BotaoExcluir_Click

    iFilialAux = giFilialEmpresa
    giFilialEmpresa = objCT.ifilialempresatela

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o fornecedor foi preenchido
    If Len(Trim(objCT.Fornecedor.ClipText)) = 0 Then gError 196421
    
    'Verifica se a filial foi digitada
    If Len(Trim(objCT.Filial.Text)) = 0 Then gError 196422
    
    'Verifica se o numero da nota fiscal foi informado
    If Len(Trim(objCT.NumeroNota.ClipText)) = 0 Then gError 196423
    
    'Move os dados da tela para a memória
    lErro = objCT.Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then gError 196424
    
    objNFsPag.iFilialEmpresa = objCT.ifilialempresatela
    
    'Busca a nota na tabela de notas fiscais à pagar
    lErro = CF("NFPag_Le_Numero", objNFsPag)
    If lErro <> SUCESSO And lErro <> 18338 Then gError 196425
    
    'Se não encontra
    If lErro = 18338 Then
    
        'Procura na tabela de notas baixadas
        lErro = CF("NFPagBaixada_Le_Numero", objNFsPag)
        If lErro <> SUCESSO And lErro <> 18344 Then gError 196426
    
        'A nota está baixada
        If lErro = AD_SQL_SUCESSO Then gError 196427
        
        'A nota não está cadastrada
        If lErro = 18344 Then gError 196428

    End If
    
    'Pede a confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NFPAG", objNFsPag.lNumNotaFiscal)
    
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Faz a exclusão da Nota Fiscal (inclusive dados contábeis)
    lErro = CF("NFPag_Exclui", objNFsPag, objCT.objContabil)
    If lErro <> SUCESSO Then gError 196429
    
    'Limpa a tela de Notas Fiscais
    Call objCT.Limpa_Tela_NotaPagar
    
    giFilialEmpresa = iFilialAux
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    giFilialEmpresa = iFilialAux

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 196421
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 196422
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 196423
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 196424, 196425, 196426, 196429
                    
        Case 196427
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_BAIXADA_EXCLUSAO", gErr, objNFsPag.lNumNotaFiscal)
                    
        Case 196428
            Call Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA1", objNFsPag.lNumNotaFiscal)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196430)
            
    End Select
    
    Exit Sub

End Sub
