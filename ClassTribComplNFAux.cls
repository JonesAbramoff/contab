VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribComplNFAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ClassTributoDocCompl

Private mvarobjTribNFAux As ClassTribNFAux 'local copy
Private mvarobjTribItemNFAux As ClassTribItemNFAux 'local copy

Public Sub Iniciar(objTribNFAux As ClassTribNFAux, objTribItemNFAux As ClassTribItemNFAux)

    Set mvarobjTribNFAux = objTribNFAux
    Set mvarobjTribItemNFAux = objTribItemNFAux

End Sub

Private Function ClassTributoDocCompl_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long
Dim lErro As Long, objVariavelCalculo As New ClassVariavelCalculo
On Error GoTo Erro_ClassTributoDocCompl_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDocCompl_GuardarVarTela = ClassTributoDocCompl_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDocCompl_GuardarVarTela:

    ClassTributoDocCompl_GuardarVarTela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153700)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocCompl_objDoc() As GlobaisTRB.ClassTributoDoc
    Set ClassTributoDocCompl_objDoc = mvarobjTribNFAux
End Property

Private Function ClassTributoDocCompl_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocCompl_ObterVar

    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador

        Case "ITEM_TIPO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.iItem
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "NAT_OPERACAO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.sNaturezaOp
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iNaturezaOpManual
                    
        Case "TIPO_TRIB"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.iTipoTributacao
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iTipoTributacaoManual

        Case "PRODUTO_VALOR"
            Select Case mvarobjTribItemNFAux.iItem
                
                Case ITEM_TIPO_FRETE
                    objVariavelCalculo.vValor = mvarobjTribNFAux.dValorFrete
                
                Case ITEM_TIPO_SEGURO
                    objVariavelCalculo.vValor = mvarobjTribNFAux.dValorSeguro
                
                Case ITEM_TIPO_DESCONTO
                    objVariavelCalculo.vValor = mvarobjTribNFAux.dValorDesconto
                
                Case ITEM_TIPO_OUTRAS_DESP
                    If mvarobjTribNFAux.iTipoNFiscal <> DOCINFO_NFIEIMP Then
                        objVariavelCalculo.vValor = mvarobjTribNFAux.dValorOutrasDespesas
                    Else
                        objVariavelCalculo.vValor = 0
                    End If
                
                Case Else
                    'tipo invalido
                    Error 27827
                    
            End Select
            
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        'para ICMS
        Case "ICMS_TIPO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.iICMSTipo
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSTipoManual
        
        Case "ICMS_BASE"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSBase
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSBaseManual
        
        Case "ICMS_RED_BASE"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSPercRedBase
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSPercRedBaseManual
        
        Case "ICMS_ALIQUOTA"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSAliquota
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSAliquotaManual
        
        Case "ICMS_VALOR"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSValor
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSValorManual
        
        Case "ICMS_BASE_SUBST"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSSubstBase
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSSubstBaseManual
            
        Case "ICMS_ALIQUOTA_SUBST"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSSubstAliquota
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSSubstAliquotaManual
        
        Case "ICMS_VALOR_SUBST"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSSubstValor
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iICMSSubstValorManual
        
        Case "ICMS_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dICMSCredito
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        'para IPI
        Case "IPI_TIPO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.iIPITipo
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iIPITipoManual
        
        Case "IPI_BASE"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dIPIBaseCalculo
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iIPIBaseManual
        
        Case "IPI_RED_BASE"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dIPIPercRedBase
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iIPIPercRedBaseManual
        
        Case "IPI_ALIQUOTA"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dIPIAliquota
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iIPIAliquotaManual
        
        Case "IPI_VALOR"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dIPIValor
            objVariavelCalculo.iStatus = mvarobjTribItemNFAux.iIPIValorManual

        Case "IPI_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dIPICredito
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "PIS_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dPISCredito
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "COFINS_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTribItemNFAux.dCOFINSCredito
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case Else
            Error 27828

    End Select

    ClassTributoDocCompl_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDocCompl_ObterVar:

    ClassTributoDocCompl_ObterVar = Err

    Select Case Err

        Case 27827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_TIPO_INVALIDO", Err, mvarobjTribItemNFAux.iItem)
        
        Case 27828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_MNEMONICO_INVALIDO", Err, sIdentificador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153701)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocCompl_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocCompl_GuardarVar

    Select Case objVariavelCalculo.sIdentificador

        Case "NAT_OPERACAO"
            mvarobjTribItemNFAux.sNaturezaOp = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iNaturezaOpManual = objVariavelCalculo.iStatus
        
        Case "TIPO_TRIB"
            mvarobjTribItemNFAux.iTipoTributacao = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iTipoTributacaoManual = objVariavelCalculo.iStatus

        'para ICMS
        Case "ICMS_TIPO"
            mvarobjTribItemNFAux.iICMSTipo = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSTipoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE"
            mvarobjTribItemNFAux.dICMSBase = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_RED_BASE"
            mvarobjTribItemNFAux.dICMSPercRedBase = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSPercRedBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA"
            mvarobjTribItemNFAux.dICMSAliquota = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR"
            mvarobjTribItemNFAux.dICMSValor = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSValorManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE_SUBST"
            mvarobjTribItemNFAux.dICMSSubstBase = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSSubstBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA_SUBST"
            mvarobjTribItemNFAux.dICMSSubstAliquota = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSSubstAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_SUBST"
            mvarobjTribItemNFAux.dICMSSubstValor = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iICMSSubstValorManual = objVariavelCalculo.iStatus

        Case "ICMS_VALOR_CREDITO"
            mvarobjTribItemNFAux.dICMSCredito = objVariavelCalculo.vValor
        
        'para IPI
        Case "IPI_TIPO"
            mvarobjTribItemNFAux.iIPITipo = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iIPITipoManual = objVariavelCalculo.iStatus
        
        Case "IPI_BASE"
            mvarobjTribItemNFAux.dIPIBaseCalculo = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iIPIValorManual = objVariavelCalculo.iStatus
            
        Case "IPI_RED_BASE"
            mvarobjTribItemNFAux.dIPIPercRedBase = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iIPIPercRedBaseManual = objVariavelCalculo.iStatus
            
        Case "IPI_ALIQUOTA"
            mvarobjTribItemNFAux.dIPIAliquota = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iIPIAliquotaManual = objVariavelCalculo.iStatus
        
        Case "IPI_VALOR"
            mvarobjTribItemNFAux.dIPIValor = objVariavelCalculo.vValor
            mvarobjTribItemNFAux.iIPIValorManual = objVariavelCalculo.iStatus

        Case "IPI_VALOR_CREDITO"
            mvarobjTribItemNFAux.dIPICredito = objVariavelCalculo.vValor
        
        Case "PIS_VALOR_CREDITO"
            mvarobjTribItemNFAux.dPISCredito = objVariavelCalculo.vValor
        
        Case "COFINS_VALOR_CREDITO"
            mvarobjTribItemNFAux.dCOFINSCredito = objVariavelCalculo.vValor
        
        Case Else
            Error 27829

    End Select

    ClassTributoDocCompl_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDocCompl_GuardarVar:

    ClassTributoDocCompl_GuardarVar = Err

    Select Case Err

        Case 27829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_MNEMONICO_INVALIDO", Err, objVariavelCalculo.sIdentificador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153702)

    End Select

    Exit Function

End Function



