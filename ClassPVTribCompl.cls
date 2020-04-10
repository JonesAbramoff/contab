VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVTribCompl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ClassTributoDocCompl

Private mvarobjPedidoDeVenda As Object 'local copy
Private mvarobjTributacaoItemPV As ClassTributacaoItemPV 'local copy

Public Sub Iniciar(objPedidoDeVenda As Object, objTributacaoItemPV As ClassTributacaoItemPV)

    Set mvarobjPedidoDeVenda = objPedidoDeVenda
    Set mvarobjTributacaoItemPV = objTributacaoItemPV

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152754)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocCompl_objDoc() As GlobaisTRB.ClassTributoDoc
    Set ClassTributoDocCompl_objDoc = mvarobjPedidoDeVenda
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
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.lNumIntDoc
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "NAT_OPERACAO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.sNaturezaOp
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iNaturezaOpManual
                    
        Case "TIPO_TRIB"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.iTipoTributacao
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iTipoTributacaoManual

        Case "PRODUTO_VALOR"
            Select Case mvarobjTributacaoItemPV.lNumIntDoc
                
                Case ITEM_TIPO_FRETE
                    objVariavelCalculo.vValor = mvarobjPedidoDeVenda.dValorFrete
                
                Case ITEM_TIPO_SEGURO
                    objVariavelCalculo.vValor = mvarobjPedidoDeVenda.dValorSeguro
                
                Case ITEM_TIPO_DESCONTO
                    objVariavelCalculo.vValor = mvarobjPedidoDeVenda.dValorDesconto
                
                Case ITEM_TIPO_OUTRAS_DESP
                    objVariavelCalculo.vValor = mvarobjPedidoDeVenda.dValorOutrasDespesas
                
                Case Else
                    'tipo invalido
                    Error 27443
                    
            End Select
            
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        'para ICMS
        Case "ICMS_TIPO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.iICMSTipo
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSTipoManual
        
        Case "ICMS_BASE"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSBase
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSBaseManual
        
        Case "ICMS_RED_BASE"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSPercRedBase
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSPercRedBaseManual
        
        Case "ICMS_ALIQUOTA"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSAliquota
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSAliquotaManual
        
        Case "ICMS_VALOR"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSValor
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSValorManual
        
        Case "ICMS_BASE_SUBST"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstBase
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSSubstBaseManual
            
        Case "ICMS_ALIQUOTA_SUBST"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstAliquota
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSSubstAliquotaManual
        
        Case "ICMS_VALOR_SUBST"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSSubstValor
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iICMSSubstValorManual
        
        Case "ICMS_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dICMSCredito
        
        'para IPI
        Case "IPI_TIPO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.iIPITipo
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPITipoManual
        
        Case "IPI_BASE"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIBaseCalculo
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIBaseManual
        
        Case "IPI_RED_BASE"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIPercRedBase
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIPercRedBaseManual
        
        Case "IPI_ALIQUOTA"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIAliquota
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIAliquotaManual
        
        Case "IPI_VALOR"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPIValor
            objVariavelCalculo.iStatus = mvarobjTributacaoItemPV.iIPIValorManual

        Case "IPI_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dIPICredito
        
        Case "PIS_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dPISCredito
        
        Case "COFINS_VALOR_CREDITO"
            objVariavelCalculo.vValor = mvarobjTributacaoItemPV.dCOFINSCredito
        
        Case Else
            Error 27444

    End Select

    ClassTributoDocCompl_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDocCompl_ObterVar:

    ClassTributoDocCompl_ObterVar = Err

    Select Case Err

        Case 27443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_TIPO_INVALIDO", Err, mvarobjTributacaoItemPV.lNumIntDoc)
        
        Case 27444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_MNEMONICO_INVALIDO", Err, sIdentificador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152755)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocCompl_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocCompl_GuardarVar

    Select Case objVariavelCalculo.sIdentificador

        Case "NAT_OPERACAO"
            mvarobjTributacaoItemPV.sNaturezaOp = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iNaturezaOpManual = objVariavelCalculo.iStatus
        
        Case "TIPO_TRIB"
            mvarobjTributacaoItemPV.iTipoTributacao = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iTipoTributacaoManual = objVariavelCalculo.iStatus

        'para ICMS
        Case "ICMS_TIPO"
            mvarobjTributacaoItemPV.iICMSTipo = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSTipoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE"
            mvarobjTributacaoItemPV.dICMSBase = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_RED_BASE"
            mvarobjTributacaoItemPV.dICMSPercRedBase = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSPercRedBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA"
            mvarobjTributacaoItemPV.dICMSAliquota = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR"
            mvarobjTributacaoItemPV.dICMSValor = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSValorManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE_SUBST"
            mvarobjTributacaoItemPV.dICMSSubstBase = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSSubstBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA_SUBST"
            mvarobjTributacaoItemPV.dICMSSubstAliquota = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSSubstAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_SUBST"
            mvarobjTributacaoItemPV.dICMSSubstValor = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iICMSSubstValorManual = objVariavelCalculo.iStatus

        Case "ICMS_VALOR_CREDITO"
            mvarobjTributacaoItemPV.dICMSCredito = objVariavelCalculo.vValor
        
        'para IPI
        Case "IPI_TIPO"
            mvarobjTributacaoItemPV.iIPITipo = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iIPITipoManual = objVariavelCalculo.iStatus
        
        Case "IPI_BASE"
            mvarobjTributacaoItemPV.dIPIBaseCalculo = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iIPIValorManual = objVariavelCalculo.iStatus
            
        Case "IPI_RED_BASE"
            mvarobjTributacaoItemPV.dIPIPercRedBase = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iIPIPercRedBaseManual = objVariavelCalculo.iStatus
            
        Case "IPI_ALIQUOTA"
            mvarobjTributacaoItemPV.dIPIAliquota = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iIPIAliquotaManual = objVariavelCalculo.iStatus
        
        Case "IPI_VALOR"
            mvarobjTributacaoItemPV.dIPIValor = objVariavelCalculo.vValor
            mvarobjTributacaoItemPV.iIPIValorManual = objVariavelCalculo.iStatus

        Case "IPI_VALOR_CREDITO"
            mvarobjTributacaoItemPV.dIPICredito = objVariavelCalculo.vValor
        
        Case "PIS_VALOR_CREDITO"
            mvarobjTributacaoItemPV.dPISCredito = objVariavelCalculo.vValor

        Case "COFINS_VALOR_CREDITO"
            mvarobjTributacaoItemPV.dCOFINSCredito = objVariavelCalculo.vValor
        
        Case Else
            Error 27445

    End Select

    ClassTributoDocCompl_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDocCompl_GuardarVar:

    ClassTributoDocCompl_GuardarVar = Err

    Select Case Err

        Case 27445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_MNEMONICO_INVALIDO", Err, objVariavelCalculo.sIdentificador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152756)

    End Select

    Exit Function

End Function


