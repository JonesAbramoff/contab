VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatchFIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Revisada e alterada por Luiz Nogueira em 05/02/04
Public Function Rotina_Inicia_Batch_IN86(ByVal sNomeArqParam As String, ByVal objIN86ModeloOrig As Object, ByVal colIN86TiposArquivo As Collection) As Long
'Inicia o batch responsável pela geraçãos dos arquivos IN86

Dim lErro As Long, objIN86Modelo As New ClassIN86Modelos

On Error GoTo Erro_Rotina_Inicia_Batch_IN86

    'Faz uma cópia dos dados de objIN86Modelo
    Call objIN86Modelo.Copiar(objIN86ModeloOrig)
    
    'Guarda uma referência global de objIN86Modelo na tela de acompanhamento do batch
    Set TelaAcompanhaBatchFIS.objIN86Modelo = objIN86Modelo
    
    'Guarda uma referência global de colIN86TiposArquivo na tela de acompanhamento do batch
    Set TelaAcompanhaBatchFIS.colIN86TiposArquivo = colIN86TiposArquivo
    
    'Indica a caption da tela de acompanhamento do batch
    TelaAcompanhaBatchFIS.Caption = TITULO_TELABATCH_GERACAO_ARQ_IN86
    
    'Guarda em uma variável global o nome do arquivo de parâmetros do batch
    TelaAcompanhaBatchFIS.sNomeArqParam = sNomeArqParam
    
    'Indica qual a rotina o batch deve disparar
    TelaAcompanhaBatchFIS.iRotinaBatch = ROTINA_ARQIN86_BATCH
    
    'Inicializa a variável que controla o cancelamento do batch
    TelaAcompanhaBatchFIS.iCancelaBatch = 0
    
    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchFIS.Show

    Rotina_Inicia_Batch_IN86 = SUCESSO

    Exit Function

Erro_Rotina_Inicia_Batch_IN86:

    Rotina_Inicia_Batch_IN86 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144655)

    End Select
    
    Exit Function
    
End Function

Function Rotina_SpedFiscal(ByVal sNomeArqParam As String, ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal iIncluiRegInv As Integer, ByVal iMotivoRegInv As Integer, ByVal dtDataInv As Date, ByVal iFiltroNatureza As Integer, ByVal iIncluiRCPE As Integer) As Long
'executa o diario sped

Dim lErro As Long

On Error GoTo Erro_Rotina_SpedFiscal

    TelaAcompanhaBatchFIS.Caption = TITULO_SPED_FISCAL
    TelaAcompanhaBatchFIS.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchFIS.iFilialEmpresa = iFilialEmpresa
    TelaAcompanhaBatchFIS.sDiretorio = sDiretorio
    TelaAcompanhaBatchFIS.dtDataIni = dtDataIni
    TelaAcompanhaBatchFIS.dtDataFim = dtDataFim
    TelaAcompanhaBatchFIS.iRotinaBatch = ROTINA_SPED_FISCAL
    
    TelaAcompanhaBatchFIS.iIncluiRegInv = iIncluiRegInv
    TelaAcompanhaBatchFIS.iMotivoRegInv = iMotivoRegInv
    TelaAcompanhaBatchFIS.dtDataInv = dtDataInv
    TelaAcompanhaBatchFIS.iFiltroNatureza = iFiltroNatureza
    TelaAcompanhaBatchFIS.iIncluiRCPE = iIncluiRCPE
    
    TelaAcompanhaBatchFIS.Show

    Rotina_SpedFiscal = SUCESSO
    
    Exit Function

Erro_Rotina_SpedFiscal:

    Rotina_SpedFiscal = gErr

    Select Case gErr

        Case 204107
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204108)

    End Select
    
    Exit Function

End Function

Function Rotina_SpedFiscalPis(ByVal sNomeArqParam As String, ByVal objEFD As ClassEFDPisCofinsSel) As Long
'executa o diario sped

Dim lErro As Long

On Error GoTo Erro_Rotina_SpedFiscalPis

    TelaAcompanhaBatchFIS.Caption = TITULO_SPED_FISCAL
    TelaAcompanhaBatchFIS.sNomeArqParam = sNomeArqParam
    Set TelaAcompanhaBatchFIS.objEFD = objEFD
    TelaAcompanhaBatchFIS.iRotinaBatch = ROTINA_SPED_FISCAL_PIS
    
    TelaAcompanhaBatchFIS.Show

    Rotina_SpedFiscalPis = SUCESSO
    
    Exit Function

Erro_Rotina_SpedFiscalPis:

    Rotina_SpedFiscalPis = gErr

    Select Case gErr

        Case 204107
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204108)

    End Select
    
    Exit Function

End Function


Function Rotina_Sped_ECF(ByVal sNomeArqParam As String, ByVal objSped As ClassSpedECFSel) As Long
'executa o diario sped

Dim lErro As Long

On Error GoTo Erro_Rotina_Sped_ECF

    TelaAcompanhaBatchFIS.Caption = TITULO_SPED_FISCAL
    TelaAcompanhaBatchFIS.sNomeArqParam = sNomeArqParam
    Set TelaAcompanhaBatchFIS.objSpedECF = objSped
    TelaAcompanhaBatchFIS.iRotinaBatch = ROTINA_ECF
    
    TelaAcompanhaBatchFIS.Show

    Rotina_Sped_ECF = SUCESSO
    
    Exit Function

Erro_Rotina_Sped_ECF:

    Rotina_Sped_ECF = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213948)

    End Select
    
    Exit Function

End Function

