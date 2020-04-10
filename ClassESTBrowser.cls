VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function RecebMaterialLista_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_RecebMaterialLista_BotaoEdita

    If Not objNFiscal Is Nothing Then
        
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 95191
            
        If objTipoDocInfo.sNomeTelaNFiscal = "RecebMaterialF" Then
        
            Call Chama_Tela("RecebMaterialF", objNFiscal)
        
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "RecebMaterialC" Then
        
            Call Chama_Tela("RecebMaterialC", objNFiscal)
        
        ElseIf objTipoDocInfo.sNomeTelaNFiscal = "RecebMaterialFCom" Then
        
            Call Chama_Tela("RecebMaterialFCom", objNFiscal)
            
        End If
    
    End If
        
    RecebMaterialLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_RecebMaterialLista_BotaoEdita:

    RecebMaterialLista_BotaoEdita = gErr
    
    Select Case gErr

        Case 95191
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149457)

    End Select

    Exit Function

End Function

Function RastroMovEstoqueLista_BotaoEdita(objRastroMovEstoque As ClassRastroMovEstoque, lErro As Long) As Long

Dim objMovEstoque As New ClassMovEstoque
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim objInventario As New ClassInventario
Dim sCodInventario As String

On Error GoTo Erro_RastroMovEstoqueLista_BotaoEdita

    If lErro = SUCESSO Then
        
        objMovEstoque.lCodigo = objRastroMovEstoque.lCodigo
        objMovEstoque.iFilialEmpresa = objRastroMovEstoque.iFilialEmpresa

        'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
        Select Case objRastroMovEstoque.iTipoNumIntDocOrigem

            Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL

                'Lê o Nome da tela e os dados da Nota Fiscal
                lErro = CF("MovEstoque_Le_NomeTela_NFiscal", objMovEstoque, sTela, objNFiscal)
                If lErro <> SUCESSO Then gError 80327 'WW

'                If giTipoVersao = VERSAO_FULL Then
                    Call Chama_Tela(sTela, objNFiscal)
'                ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'                    If objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_SAIDA Then
'                        Call Chama_Tela("NFiscalFatura", objNFiscal)
'                    ElseIf objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_ENTRADA Or objNFiscal.iTipoDocInfo = DOCINFO_NF_EXTERNA Then
'                        Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
'                    End If
'
'                End If

            Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL

                'Lê o Nome da tela e os dados da Nota Fiscal
                lErro = CF("MovEstoque_Le_NomeTela_ItemNFiscal", objMovEstoque, sTela, objNFiscal)
                If lErro <> SUCESSO Then gError 80328 'WW

'                If giTipoVersao = VERSAO_FULL Then
                    Call Chama_Tela(sTela, objNFiscal)
'                ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'                    If objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_SAIDA Then
'                        Call Chama_Tela("NFiscalFatura", objNFiscal)
'                    ElseIf objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_ENTRADA Or objNFiscal.iTipoDocInfo = DOCINFO_NF_EXTERNA Then
'                        Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
'                    End If
'
'                End If

            Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO

                'Le o nome da tela e o inventario de Origem
                lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
                If lErro <> SUCESSO Then gError 80329

                objInventario.sCodigo = sCodInventario

                Call Chama_Tela(sTela, objInventario)

            Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE

                Call Chama_Tela("MovEstoque", objMovEstoque)

            Case Else

                'Le o Nome da Tela
                lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
                If lErro <> SUCESSO Then gError 80330

                If Trim(sTela) = "" Then sTela = "MovEstoque"

                Call Chama_Tela(sTela, objMovEstoque)

        End Select
        
    End If
    
    RastroMovEstoqueLista_BotaoEdita = SUCESSO

    Exit Function

Erro_RastroMovEstoqueLista_BotaoEdita:

    RastroMovEstoqueLista_BotaoEdita = gErr

    Select Case gErr

        Case 80326, 80327, 80328, 80329, 80330 'Tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149458)

    End Select

    Exit Function

End Function

Function RastroLoteSaldoLista_BotaoEdita(objRastroLoteSaldo As ClassRastroLoteSaldo, lErro As Long) As Long

'Atenção!!! Variavel especifica para esta tela (RastroLoteSaldoLista)
Dim objRastroLote As New ClassRastreamentoLote

On Error GoTo Erro_RastroLoteSaldoLista_BotaoEdita

        'Verifica se o obj é Nothing
        If objRastroLoteSaldo Is Nothing Then
            
            Set objRastroLoteSaldo = New ClassRastroLoteSaldo
        
        End If
        
        'Atenção!!! Este trecho de código é especifico desta tela (RastroLoteSaldoLista)
        objRastroLote.sProduto = objRastroLoteSaldo.sProduto
        objRastroLote.sCodigo = objRastroLoteSaldo.sLote
        objRastroLote.iFilialOP = objRastroLoteSaldo.iFilialOP
    
    Call Chama_Tela("RastreamentoLote", objRastroLote)
    
    RastroLoteSaldoLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_RastroLoteSaldoLista_BotaoEdita:

    RastroLoteSaldoLista_BotaoEdita = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149459)

    End Select

    Exit Function

End Function

Function NFiscalInternaEntradaLista_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscalInternaEntradaLista_BotaoEdita
    
    If Not objNFiscal Is Nothing Then
        
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 95192
    
'        If giTipoVersao = VERSAO_FULL Then
            
            If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalEntrada" Then
            
                Call Chama_Tela("NFiscalEntrada", objNFiscal)
                        
            ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalEntRem" Then
            
                Call Chama_Tela("NFiscalEntRem", objNFiscal)
                        
            ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalEntDev" Then
            
                Call Chama_Tela("NFiscalEntDev", objNFiscal)
                        
            ElseIf objTipoDocInfo.sNomeTelaNFiscal = "NFiscalFatEntrada" Then
            
                Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
            
            End If
        
'        End If
    
    End If
    
    NFiscalInternaEntradaLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_NFiscalInternaEntradaLista_BotaoEdita:

    NFiscalInternaEntradaLista_BotaoEdita = gErr
      
    Select Case gErr

        Case 95192
        
        Case Else
        
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149460)

    End Select

    Exit Function

End Function


Function NFiscalFatEntradaLista_BotaoEdita(objNFiscalFatEntrada As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFiscalFatEntradaLista_BotaoEdita

    If lErro = SUCESSO Then
    
        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscalFatEntrada.iTipoNFiscal
    
        'Chamada de funcao incluida por Leo em 12/10/01 ******
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 94280
        If lErro = 31415 Then gError 94281 'Não encontrou
    
        objNFiscalFatEntrada.sNomeTelaNFiscal = objTipoDocInfo.sNomeTelaNFiscal
        
        'Leo em 12/10/01 ate aqui *****
        
        If objNFiscalFatEntrada.sNomeTelaNFiscal = "NFiscalFatEntrada" Then
        
            Call Chama_Tela("NFiscalFatEntrada", objNFiscalFatEntrada)
        
        ElseIf objNFiscalFatEntrada.sNomeTelaNFiscal = "NFiscalFatEntradaCom" Then
        
            Call Chama_Tela("NFiscalFatEntradaCom", objNFiscalFatEntrada)
        
        End If
        
    End If
    
    NFiscalFatEntradaLista_BotaoEdita = SUCESSO
        
    Exit Function
    
Erro_NFiscalFatEntradaLista_BotaoEdita:

    NFiscalFatEntradaLista_BotaoEdita = gErr
    
    Select Case gErr

       Case 94280, 94281
        
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149461)

    End Select

    Exit Function

End Function

Function NFiscalEntradaTodasLista_BotaoEdita(objNFiscalEntrada As ClassNFiscal, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFiscalEntradaTodasLista_BotaoEdita

    If lErro = SUCESSO Then
    
        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscalEntrada.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then Error 63051
        If lErro = 31415 Then Error 63052 'Não encontrou
    
        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscalEntrada)
    
    End If
    
    NFiscalEntradaTodasLista_BotaoEdita = SUCESSO

    Exit Function

Erro_NFiscalEntradaTodasLista_BotaoEdita:

    NFiscalEntradaTodasLista_BotaoEdita = Err

    Select Case Err

        Case 63045, 63051

        Case 63052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149462)

    End Select

    Exit Function

End Function


Function NFiscalEntradaLista_BotaoEdita(objNFiscalEntrada As ClassNFiscal, lErro As Long) As Long

On Error GoTo Erro_NFiscalEntradaLista_BotaoEdita

    If lErro = SUCESSO Then
    
        If lErro <> SUCESSO And lErro <> 55693 Then Error 29515
        
        If objNFiscalEntrada.sNomeTelaNFiscal = "NFiscalEntradaCom" Then
            Call Chama_Tela("NFiscalEntradaCom", objNFiscalEntrada)
        Else
            Call Chama_Tela("NFiscalEntrada", objNFiscalEntrada)
        End If
        
    End If
    
    NFiscalEntradaLista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_NFiscalEntradaLista_BotaoEdita:

    NFiscalEntradaLista_BotaoEdita = Err

    Select Case Err

        Case 29515
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149463)

    End Select

    Exit Function

End Function

Function NFInternaDestinatarioLista_BotaoEdita(objNFiscal As ClassNFiscal, lErro As Long)
'Função alterado por Leo em 12/12/01

Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFInternaDestinatarioLista_BotaoEdita

    If lErro = SUCESSO Then
                      
        Set objTipoDocInfo = New ClassTipoDocInfo

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 94282
        If lErro = 31415 Then gError 94283 'Não encontrou
           
        objNFiscal.sNomeTelaNFiscal = objTipoDocInfo.sNomeTelaNFiscal
        
        If objNFiscal.sNomeTelaNFiscal = "NFiscal" Then
         
            Call Chama_Tela("NFiscal", objNFiscal)
                    
        ElseIf objNFiscal.sNomeTelaNFiscal = "NFiscalRem" Then
        
            Call Chama_Tela("NFiscalRem", objNFiscal)
                    
        ElseIf objNFiscal.sNomeTelaNFiscal = "NFiscalDev" Then
        
            Call Chama_Tela("NFiscalDev", objNFiscal)
                    
        ElseIf objNFiscal.sNomeTelaNFiscal = "NFiscalFatura" Then
        
            Call Chama_Tela("NFiscalFatura", objNFiscal)
                    
        End If
        
    Else
        
        Call Chama_Tela("NFiscal")
    
    End If
    
    NFInternaDestinatarioLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_NFInternaDestinatarioLista_BotaoEdita:

    NFInternaDestinatarioLista_BotaoEdita = gErr

    Select Case gErr

        Case 94282, 94283
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149464)

    End Select

    Exit Function

End Function

Function NFEntNaoDevLista_BotaoEdita(objNFEntNaoDev As ClassNFiscal, lErro As Long)

Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFEntNaoDevLista_BotaoEdita

     If lErro = SUCESSO Then

        Set objTipoDocInfo = New ClassTipoDocInfo
        objTipoDocInfo.iCodigo = objNFEntNaoDev.iTipoNFiscal

        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then Error 19518
        If lErro = 31415 Then Error 19519 'Não encontrou

        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFEntNaoDev)

    Else

        Call Chama_Tela("NFiscalEntrada")

    End If
    
    NFEntNaoDevLista_BotaoEdita = SUCESSO

    Exit Function

Erro_NFEntNaoDevLista_BotaoEdita:

    NFEntNaoDevLista_BotaoEdita = Err

    Select Case Err

        Case 16991, 19518

        Case 19519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149465)

    End Select

    Exit Function

End Function

Function MovEstoqueOPLista_BotaoEdita(objMovEstoque As ClassMovEstoque, lErro As Long) As Long

Dim objItemMovEstoque As New ClassItemMovEstoque
Dim objInventario As New ClassInventario
Dim sTela As String
Dim sCodInventario As String
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_MovEstoqueOPLista_BotaoEdita

    If lErro = SUCESSO Then
        
        objItemMovEstoque.lCodigo = objMovEstoque.lCodigo
        objItemMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa
    
        'Lê o tipo Numero interno de Origem do Primeiro Movimento no BD
        lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstoque)
        If lErro <> SUCESSO And lErro <> 52891 Then Error 54786
    
        'if não encontrou  - - -  > ERRO
        If lErro = 52891 Then Error 54785
    
        'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
        Select Case objItemMovEstoque.iTipoNumIntDocOrigem
    
            Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
    
                'Lê o Nome da tela e os dados da Nota Fiscal
                lErro = CF("MovEstoque_Le_NomeTela_NFiscal", objMovEstoque, sTela, objNFiscal)
                If lErro <> SUCESSO Then Error 54787
    
                Call Chama_Tela(sTela, objNFiscal)
    
            Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
    
                'Lê o Nome da tela e os dados da Nota Fiscal
                lErro = CF("MovEstoque_Le_NomeTela_ItemNFiscal", objMovEstoque, sTela, objNFiscal)
                If lErro <> SUCESSO Then Error 54788
    
                Call Chama_Tela(sTela, objNFiscal)
    
            Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
    
                'Le o nome da tela e o inventario de Origem
                lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
                If lErro <> SUCESSO Then Error 54789
    
                objInventario.sCodigo = sCodInventario
    
                Call Chama_Tela(sTela, objInventario)
    
            Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
    
                Call Chama_Tela("MovEstoque", objMovEstoque)
    
            Case Else
    
                'Le o Nome da Tela
                lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
                If lErro <> SUCESSO Then Error 54790
    
                If Trim(sTela) = "" Then sTela = "MovEstoque"
    
                Call Chama_Tela(sTela, objMovEstoque)
    
            End Select
            
    End If
    
    MovEstoqueOPLista_BotaoEdita = SUCESSO

    Exit Function

Erro_MovEstoqueOPLista_BotaoEdita:

    MovEstoqueOPLista_BotaoEdita = Err

    Select Case Err

        Case 54784, 54786, 54787, 54788, 54789, 54790 'Tratados na rotinas chamadas

        Case 54785
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149466)

    End Select

    Exit Function

End Function

Function MovEstoqueOPLista_Trata_Parametros(objTela As Object, Optional colSelecao As Collection, Optional objMovEstoque As ClassMovEstoque, Optional objEvento1 As AdmEvento) As Long

Dim lErro As Long

On Error GoTo Erro_MovEstoqueOPLista_Trata_Parametros

    If colSelecao.Item(1) = MOV_EST_REQ_PRODUCAO Then
        objTela.Caption = "Lista de Itens Requisitados para Produção"
    End If

    If colSelecao.Item(1) = MOV_EST_PRODUCAO Then
        objTela.Caption = "Lista de Itens Produzidos"
    End If

''    If colSelecao.Item(1) = MOV_EST_CONSUMO Then
''        Me.Caption = "Lista de Requisições para Consumo"
''    End If

    MovEstoqueOPLista_Trata_Parametros = SUCESSO

    Exit Function

Erro_MovEstoqueOPLista_Trata_Parametros:

    MovEstoqueOPLista_Trata_Parametros = Err

    Select Case Err

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149467)

    End Select

    Exit Function

End Function



Function MovEstoqueLista_Consulta_BotaoEdita(objMovEstoque As ClassMovEstoque, lErro As Long) As Long

Dim objItemMovEstoque As ClassItemMovEstoque
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim objInventario As New ClassInventario
Dim sCodInventario As String
Dim objTipoMovEst As New ClassTipoMovEst

On Error GoTo Erro_MovEstoqueLista_Consulta_BotaoEdita
    
    If lErro = SUCESSO Then
    
        Set objItemMovEstoque = New ClassItemMovEstoque
        objItemMovEstoque.lCodigo = objMovEstoque.lCodigo
        objItemMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa
            
        'Lê o tipo Numero interno de Origem do Primeiro Movimento no BD
        lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstoque)
        If lErro <> SUCESSO And lErro <> 52891 Then Error 52895
        
        'if não encontrou  - - -  > ERRO
        If lErro = 52891 Then Error 52896
        
        'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
        Select Case objItemMovEstoque.iTipoNumIntDocOrigem
        
            Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
            
                'Lê o Nome da tela e os dados da Nota Fiscal
                lErro = CF("MovEstoque_Le_NomeTela_NFiscal", objMovEstoque, sTela, objNFiscal)
                If lErro <> SUCESSO Then Error 52897
                
'                If giTipoVersao = VERSAO_FULL Then
                    Call Chama_Tela(sTela, objNFiscal)
'                ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'                    If objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_SAIDA Then
'                        Call Chama_Tela("NFiscalFatura", objNFiscal)
'                    ElseIf objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_ENTRADA Or objNFiscal.iTipoDocInfo = DOCINFO_NF_EXTERNA Then
'                        Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
'                    End If
'
'                End If
                
            Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
            
                'Lê o Nome da tela e os dados da Nota Fiscal
                lErro = CF("MovEstoque_Le_NomeTela_ItemNFiscal", objMovEstoque, sTela, objNFiscal)
                If lErro <> SUCESSO Then Error 52898
            
'                If giTipoVersao = VERSAO_FULL Then
                    Call Chama_Tela(sTela, objNFiscal)
'                ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'                    If objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_SAIDA Then
'                        Call Chama_Tela("NFiscalFatura", objNFiscal)
'                    ElseIf objNFiscal.iTipoDocInfo = DOCINFO_NF_INT_ENTRADA Or objNFiscal.iTipoDocInfo = DOCINFO_NF_EXTERNA Then
'                        Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
'                    End If
'
'                End If
            
            Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
                
                'Le o nome da tela e o inventario de Origem
                lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
                If lErro <> SUCESSO Then Error 52903
                
                objInventario.sCodigo = sCodInventario
                
                Call Chama_Tela(sTela, objInventario)
                    
            Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
            
                Call Chama_Tela("MovEstoque", objMovEstoque)
                    
            Case MOVEST_TIPONUMINTDOCORIGEM_ITEMOS
            
                objTipoMovEst.iCodigo = objMovEstoque.iTipoMov
    
                lErro = CF("TipoMovEstoque_Le", objTipoMovEst)
                If lErro <> SUCESSO And lErro <> 30372 Then Error 52899
            
                If objTipoMovEst.iValidoMovInt = MOVESTOQUEINTSRV Then
                    Call Chama_Tela("MovEstoqueSRV", objMovEstoque)
                Else
                    Call Chama_Tela("OSApontamento", objMovEstoque)
                End If
            Case Else
            
                'Le o Nome da Tela
                lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
                If lErro <> SUCESSO Then Error 52899
                
                If Trim(sTela) = "" Then sTela = "MovEstoque"
                        
                Call Chama_Tela(sTela, objMovEstoque)
                    
        End Select
        
    End If
    
    MovEstoqueLista_Consulta_BotaoEdita = SUCESSO
    
    Exit Function

Erro_MovEstoqueLista_Consulta_BotaoEdita:

    MovEstoqueLista_Consulta_BotaoEdita = Err

    Select Case Err

        Case 29830, 52895, 52897, 52898, 52899, 52903 'Tratados nas rotinas chamadas
            
        Case 52896
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149468)

    End Select

    Exit Function

End Function

Function MovEstoqueListaEST_BotaoEdita(objMovEstoque As ClassMovEstoque, lErro As Long) As Long

Dim objItemMovEstoque As ClassItemMovEstoque
Dim objInventario As New ClassInventario
Dim sTela As String
Dim sCodInventario As String
Dim objNFiscal As New ClassNFiscal
Dim objTipoMovEst As New ClassTipoMovEst

On Error GoTo Erro_MovEstoqueListaEST_BotaoEdita
         
'        If giTipoVersao = VERSAO_FULL Then
        
            Set objItemMovEstoque = New ClassItemMovEstoque
            objItemMovEstoque.lCodigo = objMovEstoque.lCodigo
            objItemMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa
                
            'Lê o tipo Numero interno de Origem do Primeiro Movimento no BD
            lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstoque)
            If lErro <> SUCESSO And lErro <> 0 Then Error 52904
            
            'if não encontrou  - - -  > ERRO
            If lErro = 52891 Then Error 52905
                
            'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
            Select Case objItemMovEstoque.iTipoNumIntDocOrigem
            
                Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
                
                    'Lê o Nome da tela e os dados da Nota Fiscal
                    lErro = CF("MovEstoque_Le_NomeTela_NFiscal", objMovEstoque, sTela, objNFiscal)
                    If lErro <> SUCESSO Then Error 52906
                   
                    Call Chama_Tela(sTela, objNFiscal)
                
                Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                
                    'Lê o Nome da tela e os dados da Nota Fiscal
                    lErro = CF("MovEstoque_Le_NomeTela_ItemNFiscal", objMovEstoque, sTela, objNFiscal)
                    If lErro <> SUCESSO Then Error 52907
                
                    Call Chama_Tela(sTela, objNFiscal)
                
                Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
                    
                    'Le o nome da tela e o inventario de Origem
                    lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
                    If lErro <> SUCESSO Then Error 52908
                    
                    objInventario.sCodigo = sCodInventario
                    
                    Call Chama_Tela(sTela, objInventario)
                        
                Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
                
                    Call Chama_Tela("MovEstoque", objMovEstoque)
                    
                 Case MOVEST_TIPONUMINTDOCORIGEM_ITEMOS
                    
                        objTipoMovEst.iCodigo = objMovEstoque.iTipoMov
            
                        lErro = CF("TipoMovEstoque_Le", objTipoMovEst)
                        If lErro <> SUCESSO And lErro <> 30372 Then Error 52909
                    
                        If objTipoMovEst.iValidoMovInt = MOVESTOQUEINTSRV Then
                            Call Chama_Tela("MovEstoqueSRV", objMovEstoque)
                        Else
                            Call Chama_Tela("OSApontamento", objMovEstoque)
                        End If
                        
                Case Else
                
                    'Le o Nome da Tela
                    lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
                    If lErro <> SUCESSO Then Error 52909
                            
                    If Trim(sTela) = "" Then sTela = "MovEstoque"
                            
                    Call Chama_Tela(sTela, objMovEstoque)
                        
                End Select
'        Else
'
'            Call Chama_Tela("MovEstoque")
'
'        End If
        
    MovEstoqueListaEST_BotaoEdita = SUCESSO
        
    Exit Function

Erro_MovEstoqueListaEST_BotaoEdita:

    MovEstoqueListaEST_BotaoEdita = Err

    Select Case Err

        Case 24789, 52904, 52906, 52907, 52908, 52909 'Tratados na rotinas chamadas
            
        Case 52905
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149469)

    End Select

    Exit Function

End Function

Function MovEstoqueInternoLista_BotaoEdita(objMovEstoque As ClassMovEstoque, lErro As Long) As Long

Dim objItemMovEstoque As ClassItemMovEstoque
Dim sTela As String
Dim sCodInventario As String
Dim objInventario As New ClassInventario

On Error GoTo Erro_MovEstoqueInternoLista_BotaoEdita
       
'   If giTipoVersao = VERSAO_FULL Then
   
       Set objItemMovEstoque = New ClassItemMovEstoque
       objItemMovEstoque.lCodigo = objMovEstoque.lCodigo
       objItemMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa
           
       'Lê o tipo Numero interno de Origem do Primeiro Movimento no BD
       lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstoque)
       If lErro <> SUCESSO And lErro <> 52891 Then Error 55409
       
       'if não encontrou  - - -  > ERRO
       If lErro = 52891 Then Error 55410
   
       'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
       Select Case objItemMovEstoque.iTipoNumIntDocOrigem
       
           Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
               
               'Le o nome da tela e o inventario de Origem
               lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
               If lErro <> SUCESSO Then Error 55411
               
               objInventario.sCodigo = sCodInventario
               
               Call Chama_Tela(sTela, objInventario)
                   
           Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
           
               Call Chama_Tela("MovEstoque", objMovEstoque)
                   
           Case Else
           
               'Le o Nome da Tela
               lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
               If lErro <> SUCESSO Then Error 55412
               
               If Trim(sTela) = "" Then sTela = "MovEstoque"
                       
               Call Chama_Tela(sTela, objMovEstoque)
                   
       End Select
   
'   Else
'
'       Call Chama_Tela("MovEstoque")
'
'   End If
   
    MovEstoqueInternoLista_BotaoEdita = SUCESSO
            
    Exit Function

Erro_MovEstoqueInternoLista_BotaoEdita:

    MovEstoqueInternoLista_BotaoEdita = Err

    Select Case Err

        Case 24819, 55409, 55411, 55412
            
        Case 55410
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149470)

    End Select

    Exit Function

End Function

Function ItensNFiscalTodosSaida_Lista_BotaoEdita(objItemNFiscal As ClassItemNF, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_ItensNFiscalTodosSaida_Lista_BotaoEdita

    If lErro = SUCESSO Then
                   
        objNFiscal.lNumIntDoc = objItemNFiscal.lNumIntNF
                
        'lê a Nota Fiscal
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 80491
        If lErro = 31442 Then gError 80494 'Não encontrou
        
        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 80492
        If lErro = 31415 Then gError 80493 'Não encontrou
    
        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
            
    End If
    
    ItensNFiscalTodosSaida_Lista_BotaoEdita = SUCESSO
    
    Exit Function
    
Erro_ItensNFiscalTodosSaida_Lista_BotaoEdita:

    ItensNFiscalTodosSaida_Lista_BotaoEdita = gErr

    Select Case gErr

        Case 80491, 80492
        
        Case 80493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            
        Case 80494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149471)

    End Select

    Exit Function

End Function


Function MovEstoqueLista_TrataParametros(objTela As Object, colSelecao As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_MovEstoqueLista_TrataParametros

    If colSelecao.Item(1) = MOV_EST_REQ_PRODUCAO Then
        objTela.Caption = "Lista de Requisições para Produção"
    End If

    If colSelecao.Item(1) = MOV_EST_PRODUCAO Then
        objTela.Caption = "Lista de Entradas de Material Produzido no Estoque"
    End If

    If colSelecao.Item(1) = MOV_EST_CONSUMO Then
        objTela.Caption = "Lista de Requisições para Consumo"
    End If


    MovEstoqueLista_TrataParametros = SUCESSO

    Exit Function

Erro_MovEstoqueLista_TrataParametros:

    MovEstoqueLista_TrataParametros = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149472)

    End Select
    
    Exit Function
    
End Function

Function MovEstoqueLista_BotaoEdita(objMovEstoque As ClassMovEstoque, lErro As Long) As Long

Dim objItemMovEstoque As ClassItemMovEstoque
Dim objInventario As New ClassInventario
Dim sTela As String
Dim sCodInventario As String
Dim objNFiscal As New ClassNFiscal
Dim objTipoMovEst As New ClassTipoMovEst

On Error GoTo Erro_MovEstoqueLista_BotaoEdita

    If lErro = SUCESSO Then
    
'        If giTipoVersao = VERSAO_FULL Then
        
            Set objItemMovEstoque = New ClassItemMovEstoque
            objItemMovEstoque.lCodigo = objMovEstoque.lCodigo
            objItemMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa
                
            'Lê o tipo Numero interno de Origem do Primeiro Movimento no BD
            lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstoque)
            If lErro <> SUCESSO And lErro <> 0 Then gError 52904
            
            'if não encontrou  - - -  > ERRO
            If lErro = 52891 Then gError 52905
                
            'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
            Select Case objItemMovEstoque.iTipoNumIntDocOrigem
            
                Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
                
                    'Lê o Nome da tela e os dados da Nota Fiscal
                    lErro = CF("MovEstoque_Le_NomeTela_NFiscal", objMovEstoque, sTela, objNFiscal)
                    If lErro <> SUCESSO Then gError 52906
                   
                    Call Chama_Tela(sTela, objNFiscal)
                
                Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                
                    'Lê o Nome da tela e os dados da Nota Fiscal
                    lErro = CF("MovEstoque_Le_NomeTela_ItemNFiscal", objMovEstoque, sTela, objNFiscal)
                    If lErro <> SUCESSO Then gError 52907
                
                    Call Chama_Tela(sTela, objNFiscal)
                
                Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
                    
                    'Le o nome da tela e o inventario de Origem
                    lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
                    If lErro <> SUCESSO Then gError 52908
                    
                    objInventario.sCodigo = sCodInventario
                    
                    Call Chama_Tela(sTela, objInventario)
                        
                Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
                
                    Call Chama_Tela("MovEstoque", objMovEstoque)
                        
                Case MOVEST_TIPONUMINTDOCORIGEM_ITEMOS
                
                    objTipoMovEst.iCodigo = objMovEstoque.iTipoMov
        
                    lErro = CF("TipoMovEstoque_Le", objTipoMovEst)
                    If lErro <> SUCESSO And lErro <> 30372 Then gError 52909
                
                    If objTipoMovEst.iValidoMovInt = MOVESTOQUEINTSRV Then
                        Call Chama_Tela("MovEstoqueSRV", objMovEstoque)
                    Else
                        Call Chama_Tela("OSApontamento", objMovEstoque)
                    End If
                        
                Case Else
                
                    'Le o Nome da Tela
                    lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
                    If lErro <> SUCESSO Then gError 52909
                            
                    If Trim(sTela) = "" Then sTela = "MovEstoque"
                            
                    Call Chama_Tela(sTela, objMovEstoque)
                        
                End Select
'        Else
'
'            Call Chama_Tela("MovEstoque")
'
'        End If
        
    Else
                
        Call Chama_Tela("MovEstoque")
        
    End If

    MovEstoqueLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_MovEstoqueLista_BotaoEdita:

    MovEstoqueLista_BotaoEdita = gErr

    Select Case gErr

        Case 52904, 52906, 52907, 52908, 52909
            
        Case 52905
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149473)

    End Select
    
    Exit Function

End Function

Function MovEstoqueLista1_TrataParametros(objTela As Object, colSelecao As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_MovEstoqueLista1_TrataParametros

    If colSelecao.Item(1) = MOV_EST_REQ_PRODUCAO Then
        objTela.Caption = "Lista de Requisições para Produção"
    End If

    If colSelecao.Item(1) = MOV_EST_PRODUCAO Then
        objTela.Caption = "Lista de Entradas de Material Produzido no Estoque"
    End If

    MovEstoqueLista1_TrataParametros = SUCESSO

    Exit Function

Erro_MovEstoqueLista1_TrataParametros:

    MovEstoqueLista1_TrataParametros = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149474)

    End Select
    
    Exit Function
    
End Function

Function MovEstoqueLista1_BotaoEdita(objMovEstoque As ClassMovEstoque, lErro As Long) As Long

Dim objItemMovEstoque As ClassItemMovEstoque
Dim objInventario As New ClassInventario
Dim sTela As String
Dim sCodInventario As String
Dim objNFiscal As New ClassNFiscal
Dim objTipoMovEst As New ClassTipoMovEst

On Error GoTo Erro_MovEstoqueLista1_BotaoEdita

    If lErro = SUCESSO Then
    
'        If giTipoVersao = VERSAO_FULL Then
        
            Set objItemMovEstoque = New ClassItemMovEstoque
            objItemMovEstoque.lCodigo = objMovEstoque.lCodigo
            objItemMovEstoque.iFilialEmpresa = objMovEstoque.iFilialEmpresa
                
            'Lê o tipo Numero interno de Origem do Primeiro Movimento no BD
            lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstoque)
            If lErro <> SUCESSO And lErro <> 0 Then gError 91283
            
            'if não encontrou  - - -  > ERRO
            If lErro = 52891 Then gError 91284
                
            'Verifica se o Doc. de origem é uma NFiscal ou um item de uma NFiscal
            Select Case objItemMovEstoque.iTipoNumIntDocOrigem
            
                Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
                
                    'Lê o Nome da tela e os dados da Nota Fiscal
                    lErro = CF("MovEstoque_Le_NomeTela_NFiscal", objMovEstoque, sTela, objNFiscal)
                    If lErro <> SUCESSO Then gError 91285
                   
                    Call Chama_Tela(sTela, objNFiscal)
                
                Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                
                    'Lê o Nome da tela e os dados da Nota Fiscal
                    lErro = CF("MovEstoque_Le_NomeTela_ItemNFiscal", objMovEstoque, sTela, objNFiscal)
                    If lErro <> SUCESSO Then gError 91286
                
                    Call Chama_Tela(sTela, objNFiscal)
                
                Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
                    
                    'Le o nome da tela e o inventario de Origem
                    lErro = CF("MovEstoque_Le_NomeTela_Inventario", objMovEstoque, sTela, sCodInventario)
                    If lErro <> SUCESSO Then gError 91287
                    
                    objInventario.sCodigo = sCodInventario
                    
                    Call Chama_Tela(sTela, objInventario)
                        
                Case MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
                
                    Call Chama_Tela("MovEstoque", objMovEstoque)
                    
                Case MOVEST_TIPONUMINTDOCORIGEM_ITEMOS
                
                    objTipoMovEst.iCodigo = objMovEstoque.iTipoMov
        
                    lErro = CF("TipoMovEstoque_Le", objTipoMovEst)
                    If lErro <> SUCESSO And lErro <> 30372 Then gError 91288
                
                    If objTipoMovEst.iValidoMovInt = MOVESTOQUEINTSRV Then
                        Call Chama_Tela("MovEstoqueSRV", objMovEstoque)
                    Else
                        Call Chama_Tela("OSApontamento", objMovEstoque)
                    End If
                        
                Case Else
                
                    'Le o Nome da Tela
                    lErro = CF("MovEstoque_Le_NomeTela", objMovEstoque, sTela)
                    If lErro <> SUCESSO Then gError 91288
                            
                    If Trim(sTela) = "" Then sTela = "MovEstoque"
                            
                    Call Chama_Tela(sTela, objMovEstoque)
                        
                End Select
'        Else
'
'            Call Chama_Tela("MovEstoque")
'
'        End If
        
    Else
                
        Call Chama_Tela("MovEstoque")
        
    End If

    MovEstoqueLista1_BotaoEdita = SUCESSO
    
    Exit Function

Erro_MovEstoqueLista1_BotaoEdita:

    MovEstoqueLista1_BotaoEdita = gErr

    Select Case gErr

        Case 91283, 91285, 91286, 91287, 91288
            
        Case 91284
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149475)

    End Select
    
    Exit Function

End Function

Function ItensNFiscalTodasEnt_lista_BotaoEdita(objItemNFiscal As ClassItemNF, lErro As Long) As Long

Dim objTipoDocInfo As ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_ItensNFiscalTodasEnt_lista_BotaoEdita

    If lErro = SUCESSO Then
    
        'lê a Nota Fiscal
        objNFiscal.lNumIntDoc = objItemNFiscal.lNumIntNF
        
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 80484
        If lErro = 31442 Then gError 80485 'Não encontrou
        
        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 80478
        If lErro = 31415 Then gError 80479 'Não encontrou
    
        Call Chama_Tela(objTipoDocInfo.sNomeTelaNFiscal, objNFiscal)
        
    End If
    
    ItensNFiscalTodasEnt_lista_BotaoEdita = SUCESSO
        
    Exit Function

Erro_ItensNFiscalTodasEnt_lista_BotaoEdita:

    ItensNFiscalTodasEnt_lista_BotaoEdita = gErr
    
    Select Case gErr

        Case 80477, 80478
        
        Case 80479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            
        Case 80485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149476)

    End Select

    Exit Function

End Function

Function ProdutoEmbalagemLista_BotaoEdita(objProdutoEmbalagem As ClassProdutoEmbalagem, lErro As Long) As Long

Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoEmbalagemLista_BotaoEdita
        
    objProduto.sCodigo = objProdutoEmbalagem.sProduto
   
    Call Chama_Tela("ProdutoEmbalagem", objProduto)
    
    ProdutoEmbalagemLista_BotaoEdita = SUCESSO
    
    Exit Function

Erro_ProdutoEmbalagemLista_BotaoEdita:
    
    ProdutoEmbalagemLista_BotaoEdita = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149477)
    
    End Select
    
    Exit Function
    
End Function

Function NFiscalTipoDocInfoLista_Trata_Parametros(objTela As Object, colSelecao As Collection) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscalTipoDocInfoLista_Trata_Parametros
    
    objTipoDocInfo.iCodigo = colSelecao.Item(1)
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 94879
    
    objTela.Caption = "Lista de notas fiscais do tipo:" & objTipoDocInfo.sNomeReduzido
    
    NFiscalTipoDocInfoLista_Trata_Parametros = SUCESSO

    Exit Function

Erro_NFiscalTipoDocInfoLista_Trata_Parametros:

    NFiscalTipoDocInfoLista_Trata_Parametros = gErr

    Select Case gErr

        Case 94879
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149478)

    End Select

    Exit Function

End Function

