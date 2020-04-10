VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpRaizCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para contabilizacao automatica da importacao de nfs de raiz de mandioca

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjNFiscal As ClassNFiscal
Private mvarobjFormCTBAux As Object

Public Property Get Controls() As Object
    
    If mvarobjFormCTBAux Is Nothing Then
    
        Set mvarobjFormCTBAux = New FormAuxCtbRaiz
        
    End If
    
    Set Controls = mvarobjFormCTBAux.Controls
    
End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

Dim lErro As Long, lDoc As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjNFiscal = vParams(0)
    
    lErro = CF("Periodo_Le", gobjNFiscal.dtDataEntrada, objPeriodo)
    If lErro <> SUCESSO Then gError 130631
    
    lErro = CF("Voucher_Automatico_Trans", gobjNFiscal.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "EST", lDoc)
    If lErro <> SUCESSO Then gError 130630
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro2(Me, "NFiscalFatEntrada", gobjNFiscal.lNumIntDoc, gobjNFiscal.lFornecedor, gobjNFiscal.iFilialForn, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjNFiscal.iFilialEmpresa, 0, gobjNFiscal.lNumNotaFiscal, 0, "", 56, "Raiz mandioca")
    If lErro <> SUCESSO Then Error 32161
        
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 32161, 130630, 130631
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, objFornecedor As New ClassFornecedor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case "Valor_Total_Produto"
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorTotal
        
        Case "ICMS_Credito_Valor"
            objMnemonicoValor.colValor.Add 0
        
        Case "Valor_Frete"
            objMnemonicoValor.colValor.Add 0
        
        Case "ICMS_Credito_Frete"
            objMnemonicoValor.colValor.Add 0
        
        Case "Valor_Seguro"
            objMnemonicoValor.colValor.Add 0
        
        Case "ICMS_Credito_Seguro"
            objMnemonicoValor.colValor.Add 0
        
        Case "Valor_Despesas"
            objMnemonicoValor.colValor.Add 0
        
        Case "ICMS_Cred_OutrasDesp"
            objMnemonicoValor.colValor.Add 0
        
        Case "Valor_INSS"
            objMnemonicoValor.colValor.Add IIf(gobjNFiscal.objTributacaoNF.iINSSRetido <> 0, gobjNFiscal.objTributacaoNF.dValorINSS, 0)
        
        Case "Total_Nota_Fiscal"
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorTotal
        
        Case "Fornecedor"
            objFornecedor.lCodigo = gobjNFiscal.lFornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 130625
            If lErro <> SUCESSO Then gError 130626
            
            objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
            
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
            
        Case 130625
        
        Case 130626
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Private Sub Class_Terminate()
    Set mvarobjFormCTBAux = Nothing
End Sub

