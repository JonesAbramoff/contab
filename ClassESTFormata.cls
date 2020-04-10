VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
         
Public Sub Converte_Situacao_Kit(vCampo As Variant, sSituacao As String)
'Devolve em sSituacao a Descricao da Situacao passada em vCampo

    Select Case vCampo
    
        Case KIT_SITUACAO_PADRAO
            sSituacao = KIT_SITUACAO_STRING_PADRAO
        
        Case KIT_SITUACAO_ATIVO
            sSituacao = KIT_SITUACAO_STRING_ATIVO
        
        Case KIT_SITUACAO_INATIVO
            sSituacao = KIT_SITUACAO_STRING_INATIVO
            
        Case Else
            sSituacao = ""
        
    End Select

End Sub

Function Browse_Formata_Campo(objBrowse As AdmBrowse, colRegCampo As Collection, objValorCampo As AdmValorCampo, vCampo As Variant, iProximoCampo As Integer, iLinha As Integer, iIndice As Integer, iExercicio As Integer, iPeriodo As Integer, sNomeExterno As String) As Long

Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sConta As String
Dim sCcl As String
Dim lErro As Long
Dim sData As String
Dim sProdutoEnxuto As String
Dim sProduto As String
Dim sProdutoMascarado As String
Dim sTipoContaCcl As String
Dim sNaturezaConta As String
Dim sNaturezaProduto As String
Dim sSituacao As String
Dim sDestinacao As String
Dim sApropriacao As String
Dim sControleEstoque As String
Dim sGerencial As String
Dim sTipoDocReserva As String
Dim sStatus As String
Dim lNumIntNF As Long
Dim dPrazoMedio As Double
Dim dQuantOP As Double
Dim lNumIntDocLote As Long
'**********************************
'Alteracao Daniel em 11/07/2002
Dim sSituacaoKit As String
'**********************************
Dim sCgc As String

On Error GoTo Erro_Browse_Formata_Campo

    Select Case objValorCampo.iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Percent")
        
                Case ADM_SUBTIPO_HORA
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CDate(vCampo)
        
                Case ADM_SUBTIPO_PRECOUNITARIO
                    If Len(Trim(gobjCOM.sFormatoPrecoUnitario)) > 0 Then
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, gobjCOM.sFormatoPrecoUnitario)
                    Else
                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "Standard")
                    End If
                    
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Format(vCampo, "#,##0.00##")
                    
            End Select
        
        
        Case ADM_TIPO_DATE
            sData = Format(vCampo, "dd/mm/yyyy")
        
            If sData = "07/09/1822" Then sData = ""
            
            objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sData
    
        Case ADM_TIPO_SMALLINT
    
            Select Case objValorCampo.iSubTipo
        
                '***************************************************
                'Alteracao Daniel em 11/07/2002
                Case ADM_SUBTIPO_KIT
                    Call CF("Converte_Situacao_Kit", vCampo, sSituacaoKit)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sSituacaoKit
                '***************************************************
                
                Case ADM_SUBTIPO_STATUS
                    Call CF("Converte_Status", vCampo, sStatus)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sStatus

                Case ADM_SUBTIPO_TIPOCONTACCL
                    Call CF("Converte_TipoContaCcl", vCampo, sTipoContaCcl)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoContaCcl
        
                Case ADM_SUBTIPO_NATUREZACONTA
                    Call CF("Converte_NaturezaConta", vCampo, sNaturezaConta)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaConta
        
                Case ADM_SUBTIPO_NATUREZAPRODUTO
                    Call CF("Converte_NaturezaProduto", vCampo, sNaturezaProduto)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sNaturezaProduto
                        
                Case ADM_SUBTIPO_SITUACAO
                    Call Converte_Situacao(vCampo, sSituacao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sSituacao
                     
                Case ADM_SUBTIPO_DESTINACAO
                    Call Converte_Destinacao(vCampo, sDestinacao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sDestinacao
                                
                Case ADM_SUBTIPO_APROPRIACAO
                    Call Converte_Apropriacao(vCampo, sApropriacao)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sApropriacao
                
                Case ADM_SUBTIPO_CONTROLE_ESTOQUE
                    Call Converte_ControleEstoque(vCampo, sControleEstoque)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sControleEstoque
                
                Case ADM_SUBTIPO_GERENCIAL
                    Call Converte_Gerencial(vCampo, sGerencial)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sGerencial
                
                Case ADM_SUBTIPO_TIPODOC_RESERVA
                    Call Converte_TipoDoc_Reserva(vCampo, sTipoDocReserva)
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sTipoDocReserva
                
                Case ADM_SUBTIPO_PRAZOMEDIO
                
                        lNumIntNF = vCampo

                        'le o periodo do BD e exibe o nome externo
                        lErro = CF("ParcelasRecNF_Le_PrazoMedio", lNumIntNF, dPrazoMedio)
                        If lErro <> SUCESSO Then gError 126145

                        objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Round(dPrazoMedio, 0)
                
                Case Else
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
    
        Case Else
            
            Select Case objValorCampo.iSubTipo
        
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for uma conta, mascara a conta antes de exibi-la
                        sContaMascarada = String(STRING_CONTA, 0)
                
                        sConta = vCampo
                    
                        lErro = Mascara_MascararConta(sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 24291
                        
                    Else
                
                        sContaMascarada = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sContaMascarada
        
                Case ADM_SUBTIPO_CCL
            
                    If Len(vCampo) > 0 Then
            
                        'mascara o centro de custo
                        sCclMascarado = String(STRING_CCL, 0)
                    
                        sCcl = vCampo
        
                        lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
                        If lErro <> SUCESSO Then Error 24292
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCclMascarado
                    
                Case ADM_SUBTIPO_PRODUTO
                
                    If Len(vCampo) > 0 Then
                    
                        'se o campo for um produto, mascara o produto antes de exibi-lo
                
                        sProduto = vCampo
                    
                    
                        sProdutoMascarado = String(STRING_PRODUTO, 0)

                        'coloca o Produto no formato que é exibido na tela
                        lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
                        If lErro <> SUCESSO Then gError 20772
                    
                    Else
                
                        sProdutoMascarado = ""
                        
                    End If
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sProdutoMascarado

                Case ADM_SUBTIPO_QUANTOP_LOTE
                
                    lNumIntDocLote = vCampo
                
                    lErro = CF("RastroLoteSaldo_Obter_QuantOP_Customizado", lNumIntDocLote, dQuantOP)
                    If lErro <> SUCESSO Then Error 20772
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = Formata_Estoque(dQuantOP)

                Case ADM_SUBTIPO_CGC
                
                    Select Case Len(Trim(vCampo))
                    
                        Case STRING_CPF 'CPF
                                        
                            sCgc = Format(vCampo, "000\.000\.000-00; ; ; ")
                    
                        Case STRING_CGC 'CGC
                    
                            sCgc = Format(vCampo, "00\.000\.000\/0000-00; ; ; ")
                            
                    End Select
                    
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = sCgc

                Case Else
                
                    objBrowse.objGrid.TextMatrix(iLinha, objValorCampo.iPosicaoTela - 1) = CStr(vCampo)
                    
            End Select
            
    End Select
    
    Browse_Formata_Campo = SUCESSO
    
    Exit Function
    
Erro_Browse_Formata_Campo:

    Browse_Formata_Campo = gErr

    Select Case gErr
        
        Case 20772
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, sProduto)
        
        Case 24291
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vCampo))

        Case 24292
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vCampo))
            
        Case 126145
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149479)

    End Select

    Exit Function

End Function

Private Sub Converte_Situacao(ByVal iTipo As Integer, sTipo As String)
'converte o tipo Situação inteiro para o seu correspondente textual

    Select Case iTipo
    
        Case ITEMOP_SITUACAO_NORMAL
            sTipo = STRING_NORMAL
            
        Case ITEMOP_SITUACAO_DESAB
            sTipo = STRING_DESABILITADA
            
        Case ITEMOP_SITUACAO_SACR
            sTipo = STRING_SACRAMENTADA
        
        Case ITEMOP_SITUACAO_BAIXADA
            sTipo = STRING_BAIXADA
    
    End Select
    
End Sub

Private Sub Converte_Destinacao(ByVal iTipo As Integer, sTipo As String)
'converte o tipo Destinação inteiro para o seu correspondente textual

    Select Case iTipo
    
        Case ITEMOP_DESTINACAO_PV
            sTipo = STRING_PEDIDOVENDA
            
        Case ITEMOP_DESTINACAO_ESTOQUE
            sTipo = STRING_ESTOQUE
            
        Case ITEMOP_DESTINACAO_CONSUMO
            sTipo = STRING_CONSUMO
            
    End Select

End Sub

Private Sub Converte_Apropriacao(ByVal iTipo As Integer, sTipo As String)
'converte o tipo Apropriacao inteiro para o seu correspondente textual

    Select Case iTipo

        Case PRODUTO_CUSTO_MEDIO
            sTipo = STRING_PRODUTO_CUSTO_MEDIO

        Case PRODUTO_CUSTO_STANDARD
            sTipo = STRING_PRODUTO_CUSTO_STANDARD

        Case PRODUTO_CUSTO_INFORMADO
            sTipo = STRING_PRODUTO_CUSTO_INFORMADO

        Case PRODUTO_CUSTO_PRODUCAO
            sTipo = STRING_PRODUTO_CUSTO_PRODUCAO


    End Select

End Sub

Private Sub Converte_ControleEstoque(ByVal iTipo As Integer, sTipo As String)
'converte o tipo Controle Estoque inteiro para o seu correspondente textual

    Select Case iTipo

        Case PRODUTO_RESERVA_ESTOQUE
            sTipo = STRING_PRODUTO_RESERVA_ESTOQUE

        Case PRODUTO_ESTOQUE
            sTipo = STRING_PRODUTO_ESTOQUE

        Case PRODUTO_SEM_ESTOQUE
            sTipo = STRING_PRODUTO_SEM_ESTOQUE


    End Select

End Sub

Private Sub Converte_Gerencial(ByVal iTipo As Integer, sTipo As String)
'converte o tipo Gerencial inteiro para o seu correspondente textual

    Select Case iTipo

        Case PRODUTO_GERENCIAL
            sTipo = STRING_PRODUTO_GERENCIAL

        Case PRODUTO_FINAL
            sTipo = STRING_PRODUTO_FINAL

    End Select

End Sub

Private Sub Converte_TipoDoc_Reserva(ByVal iTipo As Integer, sTipo As String)
'converte o tipo de documento de reserva inteiro para o seu correspondente textual

    Select Case iTipo
    
        Case RESERVA_MANUT_RESERVA
            sTipo = STRING_MANUT_RESERVA
            
        Case RESERVA_PEDIDO
            sTipo = STRING_RESERVA_PEDIDO
                
    End Select

End Sub

Public Function Aglutina_Pedidos(ByVal colPedidoCompra As Collection, ByVal colPedidos As Collection, colPedidosRec As Collection) As Long

Dim objPedidoCompra As ClassPedidoCompras
Dim objPedidoCompraRec As ClassPedidoCompras
Dim iIndice As Integer
Dim iIndiceRec As Integer
Dim lErro As Long

On Error GoTo Erro_Aglutina_Pedidos

    iIndice = 1
    iIndiceRec = 1
   
    If colPedidos.Count > 0 Then Set objPedidoCompra = colPedidos(1)
    If colPedidosRec.Count > 0 Then Set objPedidoCompraRec = colPedidosRec(1)
        
    Do While iIndice <= colPedidos.Count And iIndiceRec <= colPedidosRec.Count

        If objPedidoCompra.lCodigo < objPedidoCompraRec.lCodigo Then
            colPedidoCompra.Add objPedidoCompra
            iIndice = iIndice + 1
            If iIndice <= colPedidos.Count Then Set objPedidoCompra = colPedidos(iIndice)
        End If

        If objPedidoCompraRec.lCodigo < objPedidoCompra.lCodigo Then
            colPedidoCompra.Add objPedidoCompraRec
            iIndiceRec = iIndiceRec + 1
            If iIndiceRec <= colPedidosRec.Count Then Set objPedidoCompraRec = colPedidosRec(iIndiceRec)
        End If

        If objPedidoCompra.lCodigo = objPedidoCompraRec.lCodigo Then
            colPedidoCompra.Add objPedidoCompra
            iIndice = iIndice + 1
            iIndiceRec = iIndiceRec + 1
            If iIndice <= colPedidos.Count Then Set objPedidoCompra = colPedidos(iIndice)
            If iIndiceRec <= colPedidosRec.Count Then Set objPedidoCompraRec = colPedidosRec(iIndiceRec)
        End If

    Loop

    Do While iIndice <= colPedidos.Count
        Set objPedidoCompra = colPedidos(iIndice)
        colPedidoCompra.Add objPedidoCompra
        iIndice = iIndice + 1
    Loop
    
    Do While iIndiceRec <= colPedidosRec.Count
        Set objPedidoCompraRec = colPedidosRec(iIndiceRec)
        colPedidoCompra.Add objPedidoCompraRec
        iIndiceRec = iIndiceRec + 1
    Loop
        
    Aglutina_Pedidos = SUCESSO

    Exit Function

Erro_Aglutina_Pedidos:

    Aglutina_Pedidos = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149480)

    End Select

    Exit Function
    
End Function

