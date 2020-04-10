VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoCOFINSEver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ClassTributoDriver

Private mvardFatorValor As Double 'para multiplicar o valor real obtendo o oficial
Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo

Dim lErro As Long

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then gError 32026
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc, iRecalculaAutomatica)
    If lErro <> SUCESSO Then gError 106538
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = gErr

    Select Case gErr

        Case 106538, 32026
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179028)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Ativar() As Long
    'permite a inicializacao do driver
    'pode ser utilizado p/criar cache etc

    ClassTributoDriver_Ativar = SUCESSO
    
End Function

Private Function ClassTributoDriver_Desativar() As Long
    'informa que o driver nao vai mais ser necessario
    'pode ser utilizado p/liberar cache etc

    ClassTributoDriver_Desativar = SUCESSO
    
End Function

Private Function AtualizarDoc(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long

Dim lErro As Long, colItens As Collection, dValorCredito As Double, dValorCreditoDoc As Double, objVarValorCredito As New ClassVariavelCalculo
Dim objDocItem As ClassTributoDocItem, dtDataEmissao As Date, objVarValorCompl As New ClassVariavelCalculo
Dim objVarDataEmissao As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTrib As New ClassTipoDeTributacaoMovto
Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection, objVarItemTipo As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, dtDataRef As Date
Dim objVarDataEntrada As New ClassVariavelCalculo, dtDataEntrada As Date
Dim objVarSerie As New ClassVariavelCalculo, dBaseCalculo As Double

On Error GoTo Erro_AtualizarDoc

    objVarValorCredito.sIdentificador = "COFINS_VALOR_CREDITO"
    
    lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
    If lErro <> SUCESSO Then gError 106539
    
    dtDataEmissao = objVarDataEmissao.vValor
    
    lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataEntrada)
    If lErro <> SUCESSO Then gError 106552
    
    dtDataEntrada = objVarDataEntrada.vValor
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then gError 106550
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 106551
    
    lErro = objDoc.ObterVar("SERIE", objVarSerie)
    If lErro <> SUCESSO Then gError 106551
            
    lErro = CF("NFiscal_ObtemFatorValor", mvariFilialEmpresa, objVarTipoDocInfo.vValor, objVarSerie.vValor, mvardFatorValor)
    If lErro <> SUCESSO Then gError 106551
    
    Select Case objTipoDocInfo.iTipoTrib
    
        Case TRIB_ENTRADA_CLI, TRIB_ENTRADA_FORN
            dtDataRef = dtDataEntrada
            
        Case TRIB_SAIDA_CLI, TRIB_SAIDA_FORN
            dtDataRef = dtDataEmissao
            
    End Select
    
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
    
        dValorCredito = 0
        
        If dtDataRef >= #2/1/2004# Then

            'obter do item "TIPO_TRIB" e "PRODUTO_VALOR"
            
            lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then gError 106540
            
            'obter mais info sobre o tipo trib
            objTipoTrib.iTipo = objVarTipoTrib.vValor
            lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
            If lErro <> SUCESSO Then gError 106541
    
            If objTipoTrib.iCOFINSCredita <> TIPOTRIB_SEMCREDDEB Then
            
                lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, False, True, dBaseCalculo)
                If lErro <> SUCESSO Then gError 106542
                
                'usa aliquota de 7,6%
                dValorCredito = CDbl(Format(dBaseCalculo * 0.076, "0.00"))
                
            End If
    
        End If
            
        'guarda o valor do credito do COFINS do item
        objVarValorCredito.vValor = dValorCredito
        lErro = objDocItem.GuardarVar(objVarValorCredito)
        If lErro <> SUCESSO Then gError 106543
        
        dValorCreditoDoc = dValorCreditoDoc + dValorCredito
            
    Next
    
    Call objDoc.ObterColDocCompl(colDocCompl)
    
    'percorrer os complementos (frete, seguro,...) atualizando-os
    For Each objDocCompl In colDocCompl
    
        dValorCredito = 0
        
        If dtDataRef >= #2/1/2004# Then

            lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
            If lErro <> SUCESSO Then gError 106544
        
            'obter do complemento "TIPO_TRIB" e "PRODUTO_VALOR"
            
            lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then gError 106545
            
            'obter mais info sobre o tipo trib
            objTipoTrib.iTipo = objVarTipoTrib.vValor
            lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
            If lErro <> SUCESSO Then gError 106546
    
            If objTipoTrib.iCOFINSCredita <> TIPOTRIB_SEMCREDDEB Then
            
                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVarValorCompl)
                If lErro <> SUCESSO Then gError 106547
                
                objVarValorCompl.vValor = Round(objVarValorCompl.vValor * mvardFatorValor, 2)
                
                'usa aliquota de 7,6%
                dValorCredito = CDbl(Format(objVarValorCompl.vValor * 0.076, "0.00"))
                
            End If
    
        End If
            
        'guarda o valor do credito do COFINS do complemento
        objVarValorCredito.vValor = dValorCredito
        lErro = objDocCompl.GuardarVar(objVarValorCredito)
        If lErro <> SUCESSO Then gError 106548
        
        If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
            dValorCreditoDoc = dValorCreditoDoc + dValorCredito
            
        Else
        
            dValorCreditoDoc = dValorCreditoDoc - dValorCredito
        
        End If
    
    Next
    
    'guarda o valor do credito do COFINS do doc
    objVarValorCredito.vValor = dValorCreditoDoc
    lErro = objDoc.GuardarVar(objVarValorCredito)
    If lErro <> SUCESSO Then gError 106549
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = gErr

    Select Case gErr

        Case 106539 To 106552
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179029)

    End Select

    Exit Function

End Function

Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objVarFilialFat As New ClassVariavelCalculo
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32025
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32025
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179030)

    End Select

    Exit Function

End Function


