VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjUltMovCaixa As New ClassMovimentoCaixa
Dim gobjUltMovCaixa1 As New ClassMovimentoCaixa

Public Function Obtem_Operador_Caixa(objMovCaixa As ClassMovimentoCaixa, sOperador As String) As Long
'guarda o operador do caixa

Dim lErro As Long
Dim objOperador As New ClassOperador
Dim sUltOperador As String

On Error GoTo Erro_Obtem_Operador_Caixa

    ''Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If objMovCaixa.dtDataMovimento <> gobjUltMovCaixa1.dtDataMovimento Or objMovCaixa.iCaixa <> gobjUltMovCaixa1.iCaixa Or objMovCaixa.iFilialEmpresa <> gobjUltMovCaixa1.iFilialEmpresa Then
    
        'Alterado por Luiz Nogueira em 22/04/04
        'Lê as informações necessárias para o relatório RELCXS
        lErro = CF("MovimentosCaixa_Le_MovAbertSessao", objMovCaixa)
        If lErro <> SUCESSO And lErro <> 127104 Then gError 116228
        
        'Se não encontrou => erro
        If lErro = 127104 Then gError 116236
        
        'Guarda o código do operador a ser lido
        objOperador.iCodigo = objMovCaixa.iCodOperador
        objOperador.iFilialEmpresa = objMovCaixa.iFilialEmpresa
        
        'Lê os dados do operador encontrado
        lErro = CF("Operador_Le", objOperador)
        If lErro <> SUCESSO And lErro <> 81026 Then gError 102386
        
        'Se não encontrou o operador => erro
        If lErro = 81026 Then gError 102387
        
        'Guarda os dados lidos como dados da última leitura executada
        Set gobjUltMovCaixa = objMovCaixa
        
        'Atualiza a string que guarda o operador
        sUltOperador = objOperador.iCodigo & SEPARADOR & objOperador.sNome
        
    End If
 
    'Devolve o código do operador, encontrado na última leitura feita
    sOperador = sUltOperador
    
    Obtem_Operador_Caixa = SUCESSO
    
    Exit Function

Erro_Obtem_Operador_Caixa:

    Obtem_Operador_Caixa = gErr
    
    Select Case gErr
    
        Case 116228, 102386
        
        Case 116236
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_NAO_ENCONTRADA", gErr)
        
        Case 102387
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_CADASTRADO", gErr, objMovCaixa.iCodOperador)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153051)
        
    End Select

    Exit Function
    
End Function

Public Function Obtem_HoraAbertura_Caixa(objMovCaixa As ClassMovimentoCaixa) As Long
'Guarda a hora de abertura do caixa

Dim lErro As Long

On Error GoTo Erro_Obtem_HoraAbertura_Caixa

    'Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If objMovCaixa.dtDataMovimento <> gobjUltMovCaixa.dtDataMovimento Or objMovCaixa.iCaixa <> gobjUltMovCaixa.iCaixa Or objMovCaixa.iFilialEmpresa <> gobjUltMovCaixa.iFilialEmpresa Then
            
        'Alterado por Luiz Nogueira em 22/04/04
        'Lê as informações necessárias para o relatório RELCXS.
        lErro = CF("MovimentosCaixa_Le_MovAbertSessao", objMovCaixa)
        If lErro <> SUCESSO And lErro <> 127104 Then gError 116229
        
        'Se não encontrou => erro
        If lErro = 127104 Then
            
            objMovCaixa.dHora = -1
            
            gError 116237
        
        End If
        
        Set gobjUltMovCaixa = objMovCaixa
        
    End If
    
    'Devolve a hora de abertura do caixa, encontrado na última leitura feita
    objMovCaixa.dHora = gobjUltMovCaixa.dHora
    
    Obtem_HoraAbertura_Caixa = SUCESSO
    
    Exit Function

Erro_Obtem_HoraAbertura_Caixa:

    Obtem_HoraAbertura_Caixa = gErr
    
    Select Case gErr
    
        Case 116229
        
        Case 116237
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153052)
        
    End Select

    Exit Function
    
End Function

