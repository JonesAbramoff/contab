VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemonicoComissCalcTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjTela As Object

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Function Calcula_Mnemonico_Comissoes(objMnemonicoValor As ClassMnemonicoValor) As Long
'Essa função é utilizada para calcular mnemônicos customizados para cada cliente
'Na versão 2 ela existe apenas para "constar" e sempre retornará o valor NAO_ENCONTRADO,
'pois não existem mnemônicos customizados na versão 2

Dim lErro As Long

On Error GoTo Erro_Calcula_Mnemonico_Comissoes

    'Faz uma seleção pelo mnemônico
    Select Case objMnemonicoValor.sMnemonico
    
        Case MNEMONICO_COMISSOES_VENDEDOR
            'Calcula o valor do mnemonico Vendedor
            lErro = Calcula_MnemonicoComissoes_Vendedor(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197326
    
        Case MNEMONICO_COMISSOES_TRVTIPOCLIENTE
        
            'Calcula o valor do mnemonico TRVTipoCliente
            lErro = Calcula_MnemonicoComissoes_TRVTipoCliente(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197322
        
        Case MNEMONICO_COMISSOES_TRVPRODUTO
            lErro = Calcula_MnemonicoComissoes_TRVProduto(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197403
        
        Case MNEMONICO_COMISSOES_TRVPRODUTONOMERED
            lErro = Calcula_MnemonicoComissoes_TRVProdutoNomeRed(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197404
        
        
        Case MNEMONICO_COMISSOES_TRVVENDAANT
            If objTela.gobjTRVVoucher.idiasantc > 0 Then
                objMnemonicoValor.colValor.Add 1
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case MNEMONICO_COMISSOES_TRVCMA
                objMnemonicoValor.colValor.Add objTela.giTemCMA
        
        Case MNEMONICO_COMISSOES_TRVCMCC
                objMnemonicoValor.colValor.Add objTela.giTemCMCC
        
        Case MNEMONICO_COMISSOES_TRVVENDAANUAL
            lErro = Calcula_MnemonicoComissoes_TRVVendaAnual(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 197409
        
        Case Else
        
            'Retorna uma indicação de que o mnemônico não foi tratado
            Calcula_Mnemonico_Comissoes = MNEMONICOCOMISSOES_NAO_ENCONTRADO
            
            Exit Function
            
    End Select
    
    Calcula_Mnemonico_Comissoes = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico_Comissoes:
    
    Calcula_Mnemonico_Comissoes = gErr
    
    Select Case gErr
    
        Case 197322, 197326, 197403, 197404, 197409
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197323)
    
    End Select
    
End Function
    
Private Function Calcula_MnemonicoComissoes_TRVTipoCliente(ByVal objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente


On Error GoTo Erro_Calcula_MnemonicoComissoes_TRVTipoCliente
    
    objCliente.lCodigo = objTela.gobjTRVVoucher.lCliente
    
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO Then gError 197324
    
    'Adiciona à coleção o tipo do cliente
    objMnemonicoValor.colValor.Add objCliente.iTipo
    
    Calcula_MnemonicoComissoes_TRVTipoCliente = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_TRVTipoCliente:

    Calcula_MnemonicoComissoes_TRVTipoCliente = gErr
    
    Select Case gErr
    
        Case 197324
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197325)
    
    End Select
    
    Exit Function

End Function
    
Private Function Calcula_MnemonicoComissoes_Vendedor(objMnemonicoValor As ClassMnemonicoValor) As Long
'Lê os dados da filial do cliente e devolve o código do vendedor
'objMnemonicoValor RETORNA(Output) o código do vendedor

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProduto As New ClassProduto
Dim objProdutoTRV As ClassProdutoTRV

On Error GoTo Erro_Calcula_MnemonicoComissoes_Vendedor

    objProduto.sCodigo = objTela.gobjTRVVoucher.sProduto

    lErro = CF("Produto_Le_Customizado", objProduto)
    If lErro <> SUCESSO Then gError 197327

    Set objProdutoTRV = objProduto.objInfoUsu

    If objProdutoTRV.iCodVendedor > 0 Then
        
        objMnemonicoValor.colValor.Add objProdutoTRV.iCodVendedor
    
    Else
    
        objCliente.lCodigo = objTela.gobjTRVVoucher.icliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO Then gError 197328
        
        'Guarda na coleção o código do vendedor
        objMnemonicoValor.colValor.Add objCliente.iVendedor
    
    End If
    
    Calcula_MnemonicoComissoes_Vendedor = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_Vendedor:

    Calcula_MnemonicoComissoes_Vendedor = gErr
    
    Select Case gErr
    
        Case 197327, 197328
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197329)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TRVProduto(objMnemonicoValor As ClassMnemonicoValor) As Long
'objMnemonicoValor RETORNA(Output) o código do produto

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProduto As New ClassProduto
Dim objProdutoTRV As ClassProdutoTRV

On Error GoTo Erro_Calcula_MnemonicoComissoes_TRVProduto

    objMnemonicoValor.colValor.Add objTela.gobjTRVVoucher.sProduto
    
    Calcula_MnemonicoComissoes_TRVProduto = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TRVProduto:

    Calcula_MnemonicoComissoes_TRVProduto = gErr
    
    Select Case gErr
    
        Case 197327, 197328
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197329)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TRVProdutoNomeRed(objMnemonicoValor As ClassMnemonicoValor) As Long
'objMnemonicoValor RETORNA(Output) o código do produto

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Calcula_MnemonicoComissoes_TRVProdutoNomeRed

    objProduto.sCodigo = objTela.gobjTRVVoucher.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 197406

    objMnemonicoValor.colValor.Add objProduto.sNomeReduzido
    
    Calcula_MnemonicoComissoes_TRVProdutoNomeRed = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TRVProdutoNomeRed:

    Calcula_MnemonicoComissoes_TRVProdutoNomeRed = gErr
    
    Select Case gErr
    
        Case 197406
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197407)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_TRVVendaAnual(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objTRVVoucherInfo As New ClassTRVVoucherInfo
Dim lPos As Long
Dim lVigencia As Long

On Error GoTo Erro_Calcula_MnemonicoComissoes_TRVVendaAnual

    objTRVVoucherInfo.sTipo = objTela.gobjTRVVoucher.sTipVou
    objTRVVoucherInfo.sSerie = objTela.gobjTRVVoucher.sSerie
    objTRVVoucherInfo.lNumVou = objTela.gobjTRVVoucher.lNumVou

    lErro = CF("TRVVoucherInfoSigav_Le", objTRVVoucherInfo)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 197410

    lPos = InStr(1, objTRVVoucherInfo.sVigencia, " ")

    If lPos > 0 Then
        lVigencia = StrParaLong(Mid(1, objTRVVoucherInfo.sVigencia, lPos - 1))
        If lVigencia > 90 Then
            objMnemonicoValor.colValor.Add 1
        Else
            objMnemonicoValor.colValor.Add 0
        End If
    Else
        objMnemonicoValor.colValor.Add 0
    End If

    Calcula_MnemonicoComissoes_TRVVendaAnual = SUCESSO
    
    Exit Function

Erro_Calcula_MnemonicoComissoes_TRVVendaAnual:

    Calcula_MnemonicoComissoes_TRVVendaAnual = gErr
    
    Select Case gErr
    
        Case 197410
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197411)
    
    End Select
    
    Exit Function


End Function


