VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function UsuProdArtlux_Le(ByVal objUsuProdArtlux As ClassUsuProdArtlux) As Long

Dim lErro As Long
Dim lComando As Long
Dim tUsuProdArtlux As typeUsuProdArtlux

On Error GoTo Erro_UsuProdArtlux_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 206621

    'Le a tabelaUsuProdArtlux
    lErro = Comando_Executar(lComando, "SELECT AcessoCorte, AcessoForro, AcessoMontagem FROM UsuProdArtlux WHERE FilialEmpresa= ?  AND CodUsuario= ? ", _
                tUsuProdArtlux.iAcessoCorte, tUsuProdArtlux.iAcessoForro, tUsuProdArtlux.iAcessoMontagem, _
                objUsuProdArtlux.iFilialEmpresa, objUsuProdArtlux.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 206622

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206623

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objUsuProdArtlux.iAcessoCorte = tUsuProdArtlux.iAcessoCorte
    objUsuProdArtlux.iAcessoForro = tUsuProdArtlux.iAcessoForro
    objUsuProdArtlux.iAcessoMontagem = tUsuProdArtlux.iAcessoMontagem

    'Fecha Comando
    Call Comando_Fechar(lComando)

    UsuProdArtlux_Le = SUCESSO

    Exit Function

Erro_UsuProdArtlux_Le:

    UsuProdArtlux_Le = gErr

    Select Case gErr

        Case 206621
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206622, 206623
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUPRODARTLUX", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206624)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function UsuProdArtlux_Le_Todos(ByVal colUsuProdArtlux As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tUsuProdArtlux As typeUsuProdArtlux
Dim objUsuProdArtlux As ClassUsuProdArtlux

On Error GoTo Erro_UsuProdArtlux_Le_Todos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 206625

    'Alocação de espaço no buffer
    tUsuProdArtlux.sCodUsuario = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaUsuProdArtlux
    lErro = Comando_Executar(lComando, "SELECT CodUsuario, AcessoCorte, AcessoForro, AcessoMontagem FROM UsuProdArtlux WHERE FilialEmpresa = ? ORDER BY CodUsuario ", _
                tUsuProdArtlux.sCodUsuario, tUsuProdArtlux.iAcessoCorte, tUsuProdArtlux.iAcessoForro, tUsuProdArtlux.iAcessoMontagem, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 206626

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206627

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objUsuProdArtlux = New ClassUsuProdArtlux

        objUsuProdArtlux.sCodUsuario = tUsuProdArtlux.sCodUsuario
        objUsuProdArtlux.iFilialEmpresa = giFilialEmpresa
        objUsuProdArtlux.iAcessoCorte = tUsuProdArtlux.iAcessoCorte
        objUsuProdArtlux.iAcessoForro = tUsuProdArtlux.iAcessoForro
        objUsuProdArtlux.iAcessoMontagem = tUsuProdArtlux.iAcessoMontagem
        
        colUsuProdArtlux.Add objUsuProdArtlux
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206628
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    UsuProdArtlux_Le_Todos = SUCESSO

    Exit Function

Erro_UsuProdArtlux_Le_Todos:

    UsuProdArtlux_Le_Todos = gErr

    Select Case gErr

        Case 206625
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206626 To 206628
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUPRODARTLUX", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206629)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function OrdensDeCorteArtlux_Le(ByVal colOCArtluxRet As Collection, ByVal iEtapa As Integer) As Long

Dim lErro As Long
Dim tOCArtlux As typeOCArtlux
Dim objOCArtlux As New ClassOCArtlux
Dim sFiltro As String
Dim dLote As Double, iNumLotes As Integer, iPrioridade As Integer
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim tOCProdArtlux As typeOCProdArtlux
Dim objOCProdArtlux As ClassOCProdArtlux
Dim dQtdEst As Double, dQtdPV As Double, dEstSeg As Double
Dim sDescProd As String, sTipoCouro As String
Dim colOCArtlux As New Collection, colCampos As New Collection
Dim sGrupo As String

On Error GoTo Erro_OrdensDeCorteArtlux_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206630
    Next

    'Alocação de espaço no buffer
    tOCArtlux.sProduto = String(STRING_PRODUTO, 0)
    tOCArtlux.sUsuCorte = String(STRING_USUARIO_CODIGO, 0)
    tOCArtlux.sUsuForro = String(STRING_USUARIO_CODIGO, 0)
    sDescProd = String(STRING_PRODUTO_DESCRICAO, 0)
    sGrupo = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

    'Le a tabela OrdensDeCorteArtlux
    Select Case iEtapa
        Case ETAPA_CORTE
            sFiltro = "O.DataFimCorte = ? AND O.DataFimCorte = ?"
        Case ETAPA_FORRO
            sFiltro = "DataFimForro = ? AND DataFimForro = ?"
        Case ETAPA_MONTAGEM
            sFiltro = "O.DataFimCorte <> ? AND O.DataFimForro <> ? AND O.Finalizado = 0"
    End Select
    
    lErro = Comando_Executar(alComando(1), "SELECT O.NumIntDoc, O.FilialEmpresa, O.Produto, O.Quantidade, O.UsuCorte, O.DataIniCorte, " & _
                "O.HoraIniCorte, O.DataFimCorte, O.HoraFimCorte, O.UsuForro, O.DataIniForro, O.HoraIniForro, O.DataFimForro, " & _
                "O.HoraFimForro, O.QuantidadeProd, P.QuantEst, P.QuantPV, P.EstoqueSeguranca, P.Descricao, P.Grupo FROM OrdensDeCorteArtlux AS O, ProdutoOrdemARTX AS P WHERE O.Produto = P.Produto AND O.FilialEmpresa = P.FilialEmpresa AND  O.FilialEmpresa= ? AND " & sFiltro, _
                tOCArtlux.lNumIntDoc, tOCArtlux.iFilialEmpresa, tOCArtlux.sProduto, tOCArtlux.dQuantidade, _
                tOCArtlux.sUsuCorte, tOCArtlux.dtDataIniCorte, tOCArtlux.dHoraIniCorte, tOCArtlux.dtDataFimCorte, tOCArtlux.dHoraFimCorte, _
                tOCArtlux.sUsuForro, tOCArtlux.dtDataIniForro, tOCArtlux.dHoraIniForro, tOCArtlux.dtDataFimForro, tOCArtlux.dHoraFimForro, _
                tOCArtlux.dQuantidadeProd, dQtdEst, dQtdPV, dEstSeg, sDescProd, sGrupo, giFilialEmpresa, DATA_NULA, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 206631

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206632

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
       
        Set objOCArtlux = New ClassOCArtlux
    
        objOCArtlux.lNumIntDoc = tOCArtlux.lNumIntDoc
        objOCArtlux.iFilialEmpresa = tOCArtlux.iFilialEmpresa
        objOCArtlux.sProduto = tOCArtlux.sProduto
        objOCArtlux.dQuantidade = tOCArtlux.dQuantidade
        objOCArtlux.sUsuCorte = tOCArtlux.sUsuCorte
        objOCArtlux.dtDataIniCorte = tOCArtlux.dtDataIniCorte
        objOCArtlux.dHoraIniCorte = tOCArtlux.dHoraIniCorte
        objOCArtlux.dtDataFimCorte = tOCArtlux.dtDataFimCorte
        objOCArtlux.dHoraFimCorte = tOCArtlux.dHoraFimCorte
        objOCArtlux.sUsuForro = tOCArtlux.sUsuForro
        objOCArtlux.dtDataIniForro = tOCArtlux.dtDataIniForro
        objOCArtlux.dHoraIniForro = tOCArtlux.dHoraIniForro
        objOCArtlux.dtDataFimForro = tOCArtlux.dtDataFimForro
        objOCArtlux.dHoraFimForro = tOCArtlux.dHoraFimForro
        objOCArtlux.dQuantidadeProd = tOCArtlux.dQuantidadeProd
        objOCArtlux.dQuantidadeEst = dQtdEst
        objOCArtlux.dQuantidadePV = dQtdPV
        objOCArtlux.dEstoqueSeguranca = dEstSeg
        objOCArtlux.sProdutoDesc = sDescProd
        objOCArtlux.sGrupo = sGrupo
        objOCArtlux.dQuantidadeNecess1 = objOCArtlux.dQuantidadeEst - objOCArtlux.dQuantidadePV
        objOCArtlux.dQuantidadeNecess2 = objOCArtlux.dQuantidadeEst - objOCArtlux.dQuantidadePV - objOCArtlux.dEstoqueSeguranca
        
        colOCArtlux.Add objOCArtlux
        
        If iEtapa = ETAPA_MONTAGEM Then
        
            lErro = CF("OrdensDeCorteProdArtlux_Le", alComando(3), objOCArtlux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206636
        
    Loop
    
    If iEtapa <> ETAPA_MONTAGEM Then
    
        tOCArtlux.sProduto = String(STRING_PRODUTO, 0)
        sDescProd = String(STRING_PRODUTO_DESCRICAO, 0)
        sGrupo = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
        lErro = Comando_Executar(alComando(2), "SELECT Produto, Descricao, LoteEconomico, QuantEst, QuantPedida, EstoqueSeguranca, QuantidadeNecess, Grupo FROM OCArtlux WHERE FilialEmpresa = ? ", _
                    tOCArtlux.sProduto, sDescProd, dLote, dQtdEst, dQtdPV, dEstSeg, tOCArtlux.dQuantidade, sGrupo, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 206637
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206638
    
        'Sem Dados
        Do While lErro <> AD_SQL_SEM_DADOS
                   
            If tOCArtlux.dQuantidade > QTDE_ESTOQUE_DELTA Then
            
                iNumLotes = Round(tOCArtlux.dQuantidade / dLote, 0)
                If (iNumLotes * dLote) + QTDE_ESTOQUE_DELTA < tOCArtlux.dQuantidade Then iNumLotes = iNumLotes + 1
            
                For iIndice = 1 To iNumLotes
            
                    Set objOCArtlux = New ClassOCArtlux
                    colOCArtlux.Add objOCArtlux
                
                    objOCArtlux.iFilialEmpresa = giFilialEmpresa
                    objOCArtlux.sProduto = tOCArtlux.sProduto
                    objOCArtlux.dQuantidade = dLote
                    'objOCArtlux.dtDataIniCorte = DATA_NULA
                    'objOCArtlux.dtDataFimCorte = DATA_NULA
                    'objOCArtlux.dtDataIniForro = DATA_NULA
                    'objOCArtlux.dtDataFimForro = DATA_NULA
                    objOCArtlux.dQuantidadeEst = dQtdEst
                    objOCArtlux.dQuantidadePV = dQtdPV
                    objOCArtlux.dEstoqueSeguranca = dEstSeg
                    objOCArtlux.sProdutoDesc = sDescProd
                    objOCArtlux.sGrupo = sGrupo
                    objOCArtlux.dQuantidadeNecess1 = objOCArtlux.dQuantidadeEst - objOCArtlux.dQuantidadePV
                    objOCArtlux.dQuantidadeNecess2 = objOCArtlux.dQuantidadeEst - objOCArtlux.dQuantidadePV - objOCArtlux.dEstoqueSeguranca
                    
                Next
                
            End If
            
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206639
            
        Loop
    
    End If
    
    colCampos.Add "dQuantidadeNecess1"
    colCampos.Add "dQuantidadeNecess2"
    colCampos.Add "sProduto"
    
    Call Ordena_Colecao(colOCArtlux, colOCArtluxRet, colCampos)
    
    iPrioridade = 0
    For Each objOCArtlux In colOCArtluxRet
    
        iPrioridade = iPrioridade + 1
        
        lErro = CF("Produto_TipoCouro_Le", objOCArtlux.sProduto, sTipoCouro)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOCArtlux.sTipoCouro = sTipoCouro
        objOCArtlux.iPrioridade = iPrioridade
        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdensDeCorteArtlux_Le = SUCESSO

    Exit Function

Erro_OrdensDeCorteArtlux_Le:

    OrdensDeCorteArtlux_Le = gErr

    Select Case gErr

        Case 206630
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206631, 206632, 206636
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDECORTEARTLUX", gErr)

        Case 206637 To 206639
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206640)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrdensDeCorteProdArtlux_Le(ByVal lComando As Long, ByVal objOCArtlux As ClassOCArtlux) As Long

Dim lErro As Long
Dim tOCProdArtlux As typeOCProdArtlux
Dim objOCProdArtlux As ClassOCProdArtlux
Dim dProd As Double, dPreProd As Double
Dim sUsuMontagem As String

On Error GoTo Erro_OrdensDeCorteProdArtlux_Le

    'Alocação de espaço no buffer
    tOCProdArtlux.sUsuMontagem = String(STRING_USUARIO_CODIGO, 0)

    'Le a tabelaOrdensDeCorteProdArtlux
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Seq, NumIntDocMovEst, UsuMontagem, DataIniMontagem, " & _
                "DataFimMontagem, QuantidadePreProd, QuantidadeProd FROM OrdensDeCorteProdArtlux WHERE NumIntDocOC= ? ORDER BY Seq", _
                tOCProdArtlux.lNumIntDoc, tOCProdArtlux.iSeq, tOCProdArtlux.lNumIntDocMovEst, _
                tOCProdArtlux.sUsuMontagem, tOCProdArtlux.dtDataIniMontagem, tOCProdArtlux.dtDataFimMontagem, tOCProdArtlux.dQuantidadePreProd, tOCProdArtlux.dQuantidadeProd, _
                objOCArtlux.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 206633

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206634

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOCProdArtlux = New ClassOCProdArtlux
    
        objOCProdArtlux.lNumIntDoc = tOCProdArtlux.lNumIntDoc
        objOCProdArtlux.lNumIntDocOC = objOCArtlux.lNumIntDoc
        objOCProdArtlux.iSeq = tOCProdArtlux.iSeq
        objOCProdArtlux.lNumIntDocMovEst = tOCProdArtlux.lNumIntDocMovEst
        objOCProdArtlux.sUsuMontagem = tOCProdArtlux.sUsuMontagem
        objOCProdArtlux.dtDataIniMontagem = tOCProdArtlux.dtDataIniMontagem
        objOCProdArtlux.dtDataFimMontagem = tOCProdArtlux.dtDataFimMontagem
        objOCProdArtlux.dQuantidadePreProd = tOCProdArtlux.dQuantidadePreProd
        objOCProdArtlux.dQuantidadeProd = tOCProdArtlux.dQuantidadeProd
        
        dProd = dProd + objOCProdArtlux.dQuantidadeProd
        dPreProd = dPreProd + objOCProdArtlux.dQuantidadePreProd
        
        If Len(Trim(sUsuMontagem)) = 0 Then
            sUsuMontagem = tOCProdArtlux.sUsuMontagem
        Else
            If sUsuMontagem <> tOCProdArtlux.sUsuMontagem Then sUsuMontagem = "Vários"
        End If
        
        objOCArtlux.colItens.Add objOCProdArtlux
        
        'Busca Próximo
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206635

    Loop
    
    objOCArtlux.dQuantidadeProd = dProd
    objOCArtlux.dQuantidadePreProd = dPreProd
    objOCArtlux.sUsuMontagem = sUsuMontagem

    OrdensDeCorteProdArtlux_Le = SUCESSO

    Exit Function

Erro_OrdensDeCorteProdArtlux_Le:

    OrdensDeCorteProdArtlux_Le = gErr

    Select Case gErr

        Case 206633 To 206635
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDECORTEPRODARTLUX", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206661)

    End Select

    Exit Function

End Function

Public Function Produto_TipoCouro_Le(ByVal sProduto As String, sTipoCouro As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer, sItem As String, sProdPai As String

On Error GoTo Erro_Produto_TipoCouro_Le

    sTipoCouro = ""

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206818
    Next
    
    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

    lErro = Comando_Executar(alComando(0), "SELECT Item FROM ProdutoCategoria WHERE Produto = ? AND Categoria = ? ", _
                sItem, sProduto, ARTX_CATEGORIA_TIPOCOURO)
    If lErro <> AD_SQL_SUCESSO Then gError 206819

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206820
    
    If lErro <> AD_SQL_SUCESSO Then
        
        sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        sProdPai = String(STRING_PRODUTO, 0)
        
        Call Mascara_RetornaProdutoPai(sProduto, sProdPai)
                       
        lErro = Comando_Executar(alComando(1), "SELECT Item FROM ProdutoCategoria WHERE Produto = ? AND Categoria = ? ", _
                    sItem, sProdPai, ARTX_CATEGORIA_TIPOCOURO)
        If lErro <> AD_SQL_SUCESSO Then gError 206821
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206822
        
    End If
    
    If Len(Trim(sItem)) > 0 Then sTipoCouro = sItem
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Produto_TipoCouro_Le = SUCESSO

    Exit Function

Erro_Produto_TipoCouro_Le:

    Produto_TipoCouro_Le = gErr

    Select Case gErr

        Case 206818
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206819 To 206822
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCATEGORIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206823)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function FilialEmpresaAlmox_Customiza(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FilialEmpresaAlmox_Customiza

    iFilialEmpresa = 1

    FilialEmpresaAlmox_Customiza = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresaAlmox_Customiza:

    FilialEmpresaAlmox_Customiza = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142613)
        
    End Select
    
    Exit Function

End Function

Function Customizacoes_MovEstoque(ByVal objMovEst As ClassMovEstoque) As Long
'Lê Almoxarifado Padrão do Produto da iFilialEmpresa

Dim lErro As Long

On Error GoTo Erro_Customizacoes_MovEstoque

    objMovEst.iFilialEmpresa = 1

    Customizacoes_MovEstoque = SUCESSO

    Exit Function

Erro_Customizacoes_MovEstoque:

    Customizacoes_MovEstoque = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175935)

    End Select

    Exit Function

End Function

Function Customizacao_FilialEmpresa(iFilialEmpresa As Integer) As Long
'Altera a filialempresa passada como parametro

Dim lErro As Long

On Error GoTo Erro_Customizacao_FilialEmpresa

    iFilialEmpresa = 1

    Customizacao_FilialEmpresa = SUCESSO

    Exit Function

Erro_Customizacao_FilialEmpresa:

    Customizacao_FilialEmpresa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175941)

    End Select

    Exit Function

End Function

