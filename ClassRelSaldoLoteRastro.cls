VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelSaldoLoteRastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private alComando(0 To 3) As Long
Private alComando2(0 To 3) As Long

'Identificação. Obs.: só preencher se conseguir obter os dados solicitados, ou seja, se os "valores obtidos" forem válidos.
Private sUltProduto As String
Private dtUltDataInicial As Date
Private sUltLote As String
Private iUltFilialOP As Integer
Private iUltAlmoxarifado As Integer

'valores obtidos
Private dUltQuantInicial As Double

Private sUltProduto2 As String
Private iUltFilialEmpresa As Integer
Private dUltQuantDispFilial As Double

Private sUltProduto3 As String
Private iUltFilialEmpresa3 As Integer
Private dUltQuantDispFilial3 As Double
Private dUltQuantBenef3Filial3 As Double

Function ObterQuantDispLoteFilial(dQuantDisp As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, dQuant As Double, iIndice As Integer

On Error GoTo Erro_ObterQuantDispLoteFilial

    If sProduto <> sUltProduto2 Or iFilialEmpresa <> iUltFilialEmpresa Then
    
        If alComando(0) = 0 Then
        
            For iIndice = LBound(alComando) To UBound(alComando)
            
                alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
                If alComando(0) = 0 Then gError 184100
                
            Next
            
        End If
        
        lErro = Comando_Executar(alComando(3), "SELECT SUM(QuantDispNossa) FROM RastreamentoLoteSaldo, Almoxarifado WHERE RastreamentoLoteSaldo.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ? AND Produto = ?", dQuant, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 184101
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184102
    
        sUltProduto2 = sProduto
        iUltFilialEmpresa = iFilialEmpresa
        dUltQuantDispFilial = dQuant
        
    End If
    
    dQuantDisp = dUltQuantDispFilial
    
    ObterQuantDispLoteFilial = SUCESSO
     
    Exit Function
    
Erro_ObterQuantDispLoteFilial:

    ObterQuantDispLoteFilial = gErr
     
    Select Case gErr
          
        Case 184100
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 184101, 184102
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROLOTESALDO", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184103)
     
    End Select
     
    Exit Function

End Function

Function ObterQuantDispLoteFilialComLoc(dQuantDisp As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ObterQuantDispLoteFilialComLoc

    lErro = ObterQuantLoteFilialComLoc(sProduto, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 190916
    
    dQuantDisp = dUltQuantDispFilial3
    
    ObterQuantDispLoteFilialComLoc = SUCESSO
     
    Exit Function
    
Erro_ObterQuantDispLoteFilialComLoc:

    ObterQuantDispLoteFilialComLoc = gErr
     
    Select Case gErr
          
        Case 190916
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190919)
     
    End Select
     
    Exit Function

End Function

Function ObterQuantBenef3LoteFilialComLoc(dQuantBenef3 As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ObterQuantBenef3LoteFilialComLoc

    lErro = ObterQuantLoteFilialComLoc(sProduto, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 190916
    
    dQuantBenef3 = dUltQuantBenef3Filial3
    
    ObterQuantBenef3LoteFilialComLoc = SUCESSO
     
    Exit Function
    
Erro_ObterQuantBenef3LoteFilialComLoc:

    ObterQuantBenef3LoteFilialComLoc = gErr
     
    Select Case gErr
          
        Case 190916
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190919)
     
    End Select
     
    Exit Function

End Function

Function ObterQuantLoteFilialComLoc(ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long

Dim lErro As Long, dQuant As Double, iIndice As Integer, dQuantBenef3 As Double

On Error GoTo Erro_ObterQuantLoteFilialComLoc

    If sProduto <> sUltProduto3 Or iFilialEmpresa <> iUltFilialEmpresa3 Then
    
        If alComando2(0) = 0 Then
        
            For iIndice = LBound(alComando2) To UBound(alComando2)
            
                alComando2(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
                If alComando2(0) = 0 Then gError 190916
                
            Next
            
        End If
        
        lErro = Comando_Executar(alComando2(3), "SELECT SUM(S.QuantDispNossa), SUM(S.QuantBenef3) FROM RastreamentoLoteSaldo AS S, Almoxarifado AS A,RastreamentoLote AS L  WHERE S.Almoxarifado = A.Codigo AND L.NumIntDoc = S.NumIntDocLote AND A.FilialEmpresa = ? AND S.Produto = ? AND L.Localizacao <> ?", dQuant, dQuantBenef3, iFilialEmpresa, sProduto, "")
        If lErro <> AD_SQL_SUCESSO Then gError 190917
        
        lErro = Comando_BuscarProximo(alComando2(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190918
    
        sUltProduto3 = sProduto
        iUltFilialEmpresa3 = iFilialEmpresa
        dUltQuantDispFilial3 = dQuant
        dUltQuantBenef3Filial3 = dQuantBenef3
        
    End If
    
    ObterQuantLoteFilialComLoc = SUCESSO
     
    Exit Function
    
Erro_ObterQuantLoteFilialComLoc:

    ObterQuantLoteFilialComLoc = gErr
     
    Select Case gErr
          
        Case 190916
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 190917, 190918
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROLOTESALDO", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190919)
     
    End Select
     
    Exit Function

End Function

Function ObterQuantInicialLote(dQuantInicial As Double, ByVal sProduto As String, ByVal sLote As String, ByVal iFilialOP As Integer, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date) As Long

Dim lErro As Long
Dim objRastroLoteSaldo As New ClassRastreamentoLoteSaldo

On Error GoTo Erro_ObterQuantInicialLote

    'se mudou a identificacao
    If sProduto <> sUltProduto Or dtDataInicial <> dtUltDataInicial Or sLote <> sUltLote Or iFilialOP <> iUltFilialOP Or iAlmoxarifado <> iUltAlmoxarifado Then
    
        With objRastroLoteSaldo
            .sProduto = sProduto
            .sLote = sLote
            .iFilialOP = iFilialOP
            .iAlmoxarifado = iAlmoxarifado
        End With
        
        'vou precisar buscar no bd
        lErro = CF("RastroLoteSaldo_ObterNaData", objRastroLoteSaldo, dtDataInicial, alComando)
        If lErro <> SUCESSO Then gError 130541
        
        sUltProduto = sProduto
        dtUltDataInicial = dtDataInicial
        sUltLote = sLote
        iUltFilialOP = iFilialOP
        iUltAlmoxarifado = iAlmoxarifado
        dUltQuantInicial = objRastroLoteSaldo.dQuantDispNossa
        
    End If
    
    dQuantInicial = dUltQuantInicial
        
    ObterQuantInicialLote = SUCESSO
     
    Exit Function
    
Erro_ObterQuantInicialLote:

    ObterQuantInicialLote = gErr
     
    Select Case gErr
          
        Case 130541
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130540)
     
    End Select
     
    Exit Function

End Function

'Function RastroLoteSaldo_ObterNaData(ByVal objRastroLoteSaldo As ClassRastreamentoLoteSaldo, ByVal dtDataInicial As Date, alComando() As Long) As Long
''obtem o saldo inicial do lote-almoxarifado na data informada
'
'Dim lErro As Long, iIndice As Integer, lNumIntDocLote As Long, dQuantidadeEst As Double, dFator As Double
'Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo, objTipoMovEstoque As New ClassTipoMovEst
'Dim iTipoMov As Integer, dQuantidade As Double, sSiglaUM As String, sSiglaUMEstoque As String, iClasseUM As Integer
'
'On Error GoTo Erro_RastroLoteSaldo_ObterNaData
'
'    If alComando(0) = 0 Then
'
'        For iIndice = LBound(alComando) To UBound(alComando)
'
'            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
'            If alComando(0) = 0 Then gError 130543
'
'        Next
'
'    End If
'
'    lNumIntDocLote = objRastroLoteSaldo.lNumIntDocLote
'    If lNumIntDocLote = 0 Then
'
'        With objRastroLoteSaldo
'            lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", lNumIntDocLote, .sProduto, .sLote, .iFilialOP)
'        End With
'        If lErro <> AD_SQL_SUCESSO Then gError 130544
'
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130545
'
'        'Não encontrou o lote ==> Erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 130546
'
'    End If
'
'    'obter o saldo atual
'    With tRastroLoteSaldo
'        lErro = Comando_Executar(alComando(1), "SELECT QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, " & _
'            "QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, " & _
'            "QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3 FROM RastreamentoLoteSaldo WHERE NumIntDocLote = ? AND Almoxarifado = ?", _
'            .dQuantDispNossa, .dQuantReservada, .dQuantReservadaConsig, .dQuantEmpenhada, .dQuantPedida, _
'            .dQuantRecIndl, .dQuantIndOutras, .dQuantDefeituosa, .dQuantConsig3, .dQuantConsig, .dQuantDemo3, .dQuantDemo, .dQuantConserto3, .dQuantConserto, _
'            .dQuantOutras3, .dQuantOutras, .dQuantOP, .dQuantBenef, .dQuantBenef3, lNumIntDocLote, objRastroLoteSaldo.iAlmoxarifado)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then gError 130547
'
'    lErro = Comando_BuscarProximo(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130548
'
'    'Não encontrou o lote ==> Erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 130549
'
'    'copiar dados
'    With tRastroLoteSaldo
'        objRastroLoteSaldo.dQuantDispNossa = .dQuantDispNossa
'        objRastroLoteSaldo.dQuantReservada = .dQuantReservada
'        objRastroLoteSaldo.dQuantReservadaConsig = .dQuantReservadaConsig
'        objRastroLoteSaldo.dQuantEmpenhada = .dQuantEmpenhada
'        objRastroLoteSaldo.dQuantPedida = .dQuantPedida
'        objRastroLoteSaldo.dQuantRecIndl = .dQuantRecIndl
'        objRastroLoteSaldo.dQuantIndOutras = .dQuantIndOutras
'        objRastroLoteSaldo.dQuantDefeituosa = .dQuantDefeituosa
'        objRastroLoteSaldo.dQuantConsig3 = .dQuantConsig3
'        objRastroLoteSaldo.dQuantConsig = .dQuantConsig
'        objRastroLoteSaldo.dQuantDemo3 = .dQuantDemo3
'        objRastroLoteSaldo.dQuantDemo = .dQuantDemo
'        objRastroLoteSaldo.dQuantConserto3 = .dQuantConserto3
'        objRastroLoteSaldo.dQuantConserto = .dQuantConserto
'        objRastroLoteSaldo.dQuantOutras3 = .dQuantOutras3
'        objRastroLoteSaldo.dQuantOutras = .dQuantOutras
'        objRastroLoteSaldo.dQuantOP = .dQuantOP
'        objRastroLoteSaldo.dQuantBenef = .dQuantBenef
'        objRastroLoteSaldo.dQuantBenef3 = .dQuantBenef3
'    End With
'
'    'descontar os movimentos até a data informada
'    sSiglaUM = String(STRING_UM_SIGLA, 0)
'    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
'    lErro = Comando_Executar(alComando(2), "SELECT MovimentoEstoque.TipoMov, RastreamentoMovto.Quantidade, RastreamentoMovto.SiglaUM, Produtos.SiglaUMEstoque, Produtos.ClasseUM FROM Produtos, RastreamentoMovto, MovimentoEstoque WHERE Produtos.Codigo = RastreamentoMovto.Produto AND RastreamentoMovto.NumIntDocOrigem = MovimentoEstoque.NumIntDoc AND RastreamentoMovto.TipoDocOrigem = ? AND RastreamentoMovto.NumIntDocLote = ? AND MovimentoEstoque.Data >= ? AND MovimentoEstoque.NumIntDocEst = 0 ORDER BY MovimentoEstoque.Data DESC", _
'        iTipoMov, dQuantidade, sSiglaUM, sSiglaUMEstoque, iClasseUM, 0, lNumIntDocLote, dtDataInicial)
'    If lErro <> AD_SQL_SUCESSO Then gError 130550
'
'    lErro = Comando_BuscarProximo(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130551
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'obter qtde do movimento na unidade de estoque do produto
'        'Converte a unidade de medida da quantidade do Movimento de Estoque para a UM de Estoque
'        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 130552
'
'        dQuantidadeEst = dQuantidade * dFator
'
'        'para desfazer os movimentos
'        dQuantidadeEst = -dQuantidadeEst
'
'        objTipoMovEstoque.iCodigo = iTipoMov
'        lErro = CF("TipoMovEstoque_Le", objTipoMovEstoque)
'        If lErro <> SUCESSO Then gError 130553
'
'        'se a movimentação é referente a um conserto
'        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantConserto = objRastroLoteSaldo.dQuantConserto + dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantConserto3 = objRastroLoteSaldo.dQuantConserto3 + dQuantidadeEst
'            End If
'        End If
'
'        'se é uma movimentação referente a um conserto
'        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantConserto = objRastroLoteSaldo.dQuantConserto - dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantConserto3 = objRastroLoteSaldo.dQuantConserto3 - dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a uma demonstração
'        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantDemo = objRastroLoteSaldo.dQuantDemo + dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantDemo3 = objRastroLoteSaldo.dQuantDemo3 + dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a uma demonstração
'        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantDemo = objRastroLoteSaldo.dQuantDemo - dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantDemo3 = objRastroLoteSaldo.dQuantDemo3 - dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a material em consignação
'        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantConsig = objRastroLoteSaldo.dQuantConsig + dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantConsig3 = objRastroLoteSaldo.dQuantConsig3 + dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a material em consignação
'        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantConsig = objRastroLoteSaldo.dQuantConsig - dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantConsig3 = objRastroLoteSaldo.dQuantConsig3 - dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a outras movimentações de material
'        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantOutras = objRastroLoteSaldo.dQuantOutras + dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantOutras3 = objRastroLoteSaldo.dQuantOutras3 + dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a outras movimentações de material
'        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantOutras = objRastroLoteSaldo.dQuantOutras - dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantOutras3 = objRastroLoteSaldo.dQuantOutras3 - dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a material em beneficiamento
'        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantBenef = objRastroLoteSaldo.dQuantBenef + dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantBenef3 = objRastroLoteSaldo.dQuantBenef3 + dQuantidadeEst
'            End If
'        End If
'
'        'se é uma movimentação referente a um conserto
'        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'                objRastroLoteSaldo.dQuantBenef = objRastroLoteSaldo.dQuantBenef - dQuantidadeEst
'            Else
'                objRastroLoteSaldo.dQuantBenef3 = objRastroLoteSaldo.dQuantBenef3 - dQuantidadeEst
'            End If
'        End If
'
'        'se a movimentação é referente a outras movimentações de material indisponivel
'        If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
'            objRastroLoteSaldo.dQuantIndOutras = objRastroLoteSaldo.dQuantIndOutras + dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a outras movimentações de material indisponivel
'        If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
'            objRastroLoteSaldo.dQuantIndOutras = objRastroLoteSaldo.dQuantIndOutras - dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a nosso material disponivel
'        If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
'            objRastroLoteSaldo.dQuantDispNossa = objRastroLoteSaldo.dQuantDispNossa + dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a nosso material disponivel
'        If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
'            objRastroLoteSaldo.dQuantDispNossa = objRastroLoteSaldo.dQuantDispNossa - dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a material defeituoso
'        If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
'            objRastroLoteSaldo.dQuantDefeituosa = objRastroLoteSaldo.dQuantDefeituosa + dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a material defeituoso
'        If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
'            objRastroLoteSaldo.dQuantDefeituosa = objRastroLoteSaldo.dQuantDefeituosa - dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a material recebido e indisponível
'        If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
'            objRastroLoteSaldo.dQuantRecIndl = objRastroLoteSaldo.dQuantRecIndl + dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a material recebido e indisponível
'        If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
'            objRastroLoteSaldo.dQuantRecIndl = objRastroLoteSaldo.dQuantRecIndl - dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a material em ordem de producao
'        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
'            objRastroLoteSaldo.dQuantOP = objRastroLoteSaldo.dQuantOP + dQuantidadeEst
'        End If
'
'        'se a movimentação é referente a material em ordem de producao
'        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
'            objRastroLoteSaldo.dQuantOP = objRastroLoteSaldo.dQuantOP - dQuantidadeEst
'        End If
'
'        lErro = Comando_BuscarProximo(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130554
'
'    Loop
'
'    RastroLoteSaldo_ObterNaData = SUCESSO
'
'    Exit Function
'
'Erro_RastroLoteSaldo_ObterNaData:
'
'    RastroLoteSaldo_ObterNaData = gErr
'
'    Select Case gErr
'
'        Case 130543
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 130544, 130545, 130546
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROLOTE", gErr)
'
'        Case 130547, 130548, 130549
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROLOTESALDO", gErr)
'
'        Case 130550, 130551, 130554
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROMOVTO", gErr)
'
'        Case 130552, 130553
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130542)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub Class_Terminate()

Dim iIndice As Integer

    For iIndice = LBound(alComando) To UBound(alComando)
    
        If alComando(iIndice) <> 0 Then Call Comando_Fechar(alComando(iIndice))
    
    Next
    
    For iIndice = LBound(alComando2) To UBound(alComando2)
    
        If alComando2(iIndice) <> 0 Then Call Comando_Fechar(alComando2(iIndice))
    
    Next

End Sub
