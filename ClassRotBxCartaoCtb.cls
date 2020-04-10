VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotBxCartaoCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjFormCTBAux As Object

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloRec As ClassTituloReceber
Private gobjParcelaRec As ClassParcelaReceber
Private gobjBaixaParcRec As ClassBaixaParcRec
Private gobjBaixaReceber As ClassBaixaReceber
Private gsContaCtaCorrente As String 'conta contabil da conta corrente dos depositos
Private giFilialEmpresaConta As Integer 'filial empresa possuidora da conta corrente utilizada p/o deposito

Private mvarobjMov As ClassAdmExtFinMov
Private mvarobjDebitoRecCli As ClassDebitoRecCli

Private Const DATA_CREDITO As String = "Data_Credito"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"

Public Property Get Controls() As Object
    
    If mvarobjFormCTBAux Is Nothing Then
    
        Set mvarobjFormCTBAux = New FormAuxCTB
        
    End If
    
    Set Controls = mvarobjFormCTBAux.Controls
    
End Property

Public Property Set objMov(ByVal vData As ClassAdmExtFinMov)
    Set mvarobjMov = vData
End Property

Public Property Get objMov() As ClassAdmExtFinMov
    Set objMov = mvarobjMov
End Property

Public Property Set objDebitoRecCli(ByVal vData As ClassDebitoRecCli)
    Set mvarobjDebitoRecCli = vData
End Property

Public Property Get objDebitoRecCli() As ClassDebitoRecCli
    Set objDebitoRecCli = mvarobjDebitoRecCli
End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcrec e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iConta As Integer, dValorLivroAux As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcRec = vParams(0)
    Set gobjParcelaRec = vParams(1)
    Set gobjTituloRec = vParams(2)
    Set gobjBaixaReceber = vParams(3)

    'se ainda nao obtive a filial empresa onde vai ser feito o deposito
    If giFilialEmpresaConta = 0 Then
    
        iConta = objMov.iCodConta
    
        lErro = CF("ContaCorrenteInt_Le", iConta, objContasCorrentesInternas)
        If lErro <> SUCESSO Then Error 32243
    
        giFilialEmpresaConta = objContasCorrentesInternas.iFilialEmpresa
        
    End If
    
    'obtem numero de doc para a filial onde vai ser feito o deposito
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32244
    
    Controls("GridParcelas").Tag = gobjBaixaReceber.colBaixaParcRec.Count
        
    'grava a contabilizacao na filial da cta
    lErro = objContabAutomatica.Gravar_Registro(Me, "ExtratoCartaoCredito", objMov.lNumMovto, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32248
            
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32243 To 32248
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144657)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor
Dim iLinha As Integer, iConta As Integer, sContaContabil As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, sContaTela As String
Dim dValor As Double
Dim objBaixaParcRec As ClassBaixaParcRec

On Error GoTo Erro_Calcula_Mnemonico

        Select Case objMnemonicoValor.sMnemonico
                        
            Case "Conta_Contabil_Conta", DATA_CREDITO
                lErro = Calcula_Mnemonico_Comum(objMnemonicoValor)
                If lErro <> SUCESSO Then Error 32262
                
            Case "Valor_Recebido"
            
                dValor = Arredonda_Moeda(objMov.dValorLiq)
                objMnemonicoValor.colValor.Add dValor
            
            Case "Valor_Baixado"
                For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                    dValor = Arredonda_Moeda(dValor + objBaixaParcRec.dValorBaixado)
                Next
                objMnemonicoValor.colValor.Add dValor
            
            Case "Valor_Nao_Localizado"
                dValor = Arredonda_Moeda(objDebitoRecCli.dValorTotal)
                objMnemonicoValor.colValor.Add dValor
            
            Case Else
                Error 39695
    
        End Select

   
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 32253, 56527, 56528, 56529, 56530, 56803, 32262, 32263, 32274, 39586, 56500
        
        Case 56804
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144658)

    End Select

    Exit Function

End Function

Function Calcula_Mnemonico_Comum(objMnemonicoValor As ClassMnemonicoValor) As Long
'codigo em comum do Calcula_Mnemonico entre o processamento de baixas de titulos e o processamento de tarifas e custas

Dim lErro As Long, objCliente As New ClassCliente, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor
Dim iLinha As Integer, iConta As Integer, sContaContabil As String, sBancoNomeRed As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, sContaTela As String
Dim dValor As Double, objBanco As New ClassBanco

On Error GoTo Erro_Calcula_Mnemonico_Comum

    Select Case objMnemonicoValor.sMnemonico

        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
            'calcula-la apenas uma vez e deixa-la guardada
                
            If gsContaCtaCorrente = "" Then
                
                iConta = objMov.iCodConta
                lErro = CF("ContaCorrenteInt_Le", iConta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 56546
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56547
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                
        Case DATA_CREDITO
            objMnemonicoValor.colValor.Add objMov.dtData
        
        Case Else
            Error 39695

    End Select
    
    Calcula_Mnemonico_Comum = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico_Comum:

    Calcula_Mnemonico_Comum = Err

    Select Case Err

        Case 32280, 32253, 56527, 56528, 56529, 56530, 56803
        
        Case 56804
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico_Comum = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144660)

    End Select

    Exit Function

End Function

Private Sub Class_Terminate()
    Set mvarobjFormCTBAux = Nothing
End Sub

