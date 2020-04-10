VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function UsuProdArtlux_Grava(ByVal objUsuProdArtlux As ClassUsuProdArtlux) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_UsuProdArtlux_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206641

    lErro = UsuProdArtlux_Grava_EmTrans(objUsuProdArtlux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 206642

    UsuProdArtlux_Grava = SUCESSO

    Exit Function

Erro_UsuProdArtlux_Grava:

    UsuProdArtlux_Grava = gErr

    Select Case gErr

        Case 206641
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 206642
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206643)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function UsuProdArtlux_Grava_EmTrans(ByVal objUsuProdArtlux As ClassUsuProdArtlux) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_UsuProdArtlux_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206644
    Next

    'Le a tabelaUsuProdArtlux
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM UsuProdArtlux WHERE FilialEmpresa= ?  AND CodUsuario= ? ", _
                0, iAux, objUsuProdArtlux.iFilialEmpresa, objUsuProdArtlux.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 206645

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206646

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 206647

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE UsuProdArtlux SET AcessoCorte= ? , AcessoForro= ? , AcessoMontagem= ? ", alComando(0), _
                objUsuProdArtlux.iAcessoCorte, objUsuProdArtlux.iAcessoForro, objUsuProdArtlux.iAcessoMontagem)
        If lErro <> AD_SQL_SUCESSO Then gError 206648

    Else

        lErro = Comando_Executar(alComando(1), "INSERT INTO UsuProdArtlux( FilialEmpresa, CodUsuario, AcessoCorte, AcessoForro, " & _
                "AcessoMontagem) VALUES (?,?,?,?,?)", _
                objUsuProdArtlux.iFilialEmpresa, objUsuProdArtlux.sCodUsuario, objUsuProdArtlux.iAcessoCorte, objUsuProdArtlux.iAcessoForro, _
                objUsuProdArtlux.iAcessoMontagem)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206649

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    UsuProdArtlux_Grava_EmTrans = SUCESSO

    Exit Function

Erro_UsuProdArtlux_Grava_EmTrans:

    UsuProdArtlux_Grava_EmTrans = gErr

    Select Case gErr

        Case 206644
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206645, 206646
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUPRODARTLUX", gErr)

        Case 206647
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_USUPRODARTLUX", gErr)

        Case 206648
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_USUPRODARTLUX", gErr)

        Case 206649
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUPRODARTLUX", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206650)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function OrdensDeCorteArtlux_Grava(ByVal objOCArtlux As ClassOCArtlux) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OrdensDeCorteArtlux_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206651

    lErro = OrdensDeCorteArtlux_Grava_EmTrans(objOCArtlux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 206652

    OrdensDeCorteArtlux_Grava = SUCESSO

    Exit Function

Erro_OrdensDeCorteArtlux_Grava:

    OrdensDeCorteArtlux_Grava = gErr

    Select Case gErr

        Case 206651
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 206652
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206653)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function OrdensDeCorteArtlux_Grava_EmTrans(ByVal objOCArtlux As ClassOCArtlux) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long
Dim objOCArtluxBD As New ClassOCArtlux
Dim objOCProd As ClassOCProdArtlux
Dim iSeq As Integer
Dim dQtdPreProd As Double
Dim dQtdProd As Double
Dim sUsuMontagem As String

On Error GoTo Erro_OrdensDeCorteArtlux_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206654
    Next

    'Le a tabelaOrdensDeCorteArtlux
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM OrdensDeCorteArtlux WHERE NumIntDoc = ? ", _
                0, iAux, objOCArtlux.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 206655

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206656
    
    'Se finalizou o Corte considera montado
    If objOCArtlux.dtDataFimCorte <> DATA_NULA And objOCArtlux.dtDataFimForro <> DATA_NULA Then
        
        If objOCArtlux.dQuantidadeProd <> objOCArtlux.dQuantidade Then
        
            Set objOCProd = New ClassOCProdArtlux
        
            objOCProd.dQuantidadePreProd = objOCArtlux.dQuantidade - objOCArtlux.dQuantidadeProd
            objOCProd.dQuantidadeProd = objOCProd.dQuantidadePreProd
            objOCProd.dtDataFimMontagem = Date
            objOCProd.dtDataIniMontagem = Date
            objOCProd.sUsuMontagem = gsUsuario
        
            objOCArtlux.colItens.Add objOCProd
        
        End If
        
        iSeq = 0
        For Each objOCProd In objOCArtlux.colItens
            iSeq = iSeq + 1
            objOCProd.iSeq = iSeq
            dQtdPreProd = dQtdPreProd + objOCProd.dQuantidadePreProd
            dQtdProd = dQtdProd + objOCProd.dQuantidadeProd
            If Len(Trim(sUsuMontagem)) = 0 Then
                sUsuMontagem = objOCProd.sUsuMontagem
            Else
                If sUsuMontagem <> objOCProd.sUsuMontagem Then sUsuMontagem = "Vários"
            End If
        Next
        objOCArtlux.dQuantidadeProd = dQtdProd
        objOCArtlux.dQuantidadePreProd = dQtdPreProd
        objOCArtlux.sUsuMontagem = sUsuMontagem
    
    End If

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 206657

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE OrdensDeCorteArtlux SET Quantidade= ? , UsuCorte= ? , DataIniCorte= ? , HoraIniCorte= ? , " & _
                "DataFimCorte= ? , HoraFimCorte= ? , UsuForro= ? , DataIniForro= ? , HoraIniForro= ? , " & _
                "DataFimForro= ? , HoraFimForro= ? , QuantidadeProd= ?, QuantidadeEmProd = ?, Finalizado = ?, OrdemCorte = ? , OrdemForro = ? ", alComando(0), _
                objOCArtlux.dQuantidade, objOCArtlux.sUsuCorte, objOCArtlux.dtDataIniCorte, objOCArtlux.dHoraIniCorte, _
                objOCArtlux.dtDataFimCorte, objOCArtlux.dHoraFimCorte, objOCArtlux.sUsuForro, objOCArtlux.dtDataIniForro, objOCArtlux.dHoraIniForro, _
                objOCArtlux.dtDataFimForro, objOCArtlux.dHoraFimForro, objOCArtlux.dQuantidadeProd, objOCArtlux.dQuantidade - objOCArtlux.dQuantidadeProd, IIf(objOCArtlux.dQuantidade - objOCArtlux.dQuantidadeProd < QTDE_ESTOQUE_DELTA, 1, 0), IIf(objOCArtlux.sUsuCorte = "", 1, 0), IIf(objOCArtlux.sUsuForro = "", 1, 0))
        If lErro <> AD_SQL_SUCESSO Then gError 206658
        
        objOCArtluxBD.lNumIntDoc = objOCArtlux.lNumIntDoc
        
        lErro = CF("OrdensDeCorteProdArtlux_Le", alComando(2), objOCArtluxBD)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else
    
        If objOCArtlux.lNumIntDoc <> 0 Then gError 206778

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_ORDENSDECORTEARTLUX", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objOCArtlux.lNumIntDoc = lNumIntDoc
        
        If objOCArtlux.iManual = DESMARCADO Then
            objOCArtlux.dtDataManual = Date
            objOCArtlux.sUsuManual = gsUsuario
        End If

        lErro = Comando_Executar(alComando(1), "INSERT INTO OrdensDeCorteArtlux( NumIntDoc, FilialEmpresa, Produto, Quantidade, " & _
                "UsuCorte, DataIniCorte, HoraIniCorte, DataFimCorte, HoraFimCorte, " & _
                "UsuForro, DataIniForro, HoraIniForro, DataFimForro, HoraFimForro, " & _
                "QuantidadeProd, QuantidadeEmProd, Finalizado, OrdemCorte, OrdemForro, Manual, DataManual, UsuManual) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objOCArtlux.lNumIntDoc, objOCArtlux.iFilialEmpresa, objOCArtlux.sProduto, objOCArtlux.dQuantidade, _
                objOCArtlux.sUsuCorte, objOCArtlux.dtDataIniCorte, objOCArtlux.dHoraIniCorte, objOCArtlux.dtDataFimCorte, objOCArtlux.dHoraFimCorte, _
                objOCArtlux.sUsuForro, objOCArtlux.dtDataIniForro, objOCArtlux.dHoraIniForro, objOCArtlux.dtDataFimForro, objOCArtlux.dHoraFimForro, _
                objOCArtlux.dQuantidadeProd, objOCArtlux.dQuantidade - objOCArtlux.dQuantidadeProd, IIf(objOCArtlux.dQuantidade - objOCArtlux.dQuantidadeProd < QTDE_ESTOQUE_DELTA, 1, 0), IIf(objOCArtlux.sUsuCorte = "", 1, 0), IIf(objOCArtlux.sUsuForro = "", 1, 0), objOCArtlux.iManual, objOCArtlux.dtDataManual, objOCArtlux.sUsuManual)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206659

    End If

    lErro = OrdensDeCorteProdArtlux_Grava(objOCArtlux, objOCArtluxBD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdensDeCorteArtlux_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OrdensDeCorteArtlux_Grava_EmTrans:

    OrdensDeCorteArtlux_Grava_EmTrans = gErr

    Select Case gErr

        Case 206654
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206655, 206656
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDECORTEARTLUX", gErr)

        Case 206657
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ORDENSDECORTEARTLUX", gErr)

        Case 206658
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ORDENSDECORTEARTLUX", gErr)

        Case 206659
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDECORTEARTLUX", gErr)
            
        Case 206778
            Call Rotina_Erro(vbOKOnly, "ERRO_OC_INEXISTENTE", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206660)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function OrdensDeCorteProdArtlux_Grava(ByVal objOCArtlux As ClassOCArtlux, ByVal objOCArtluxBD As ClassOCArtlux) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim colNovos As New Collection
Dim colAlterados As New Collection
Dim colExcluidos As New Collection
Dim colMovEst As New Collection
Dim objOCProdArtlux As ClassOCProdArtlux
Dim objOCProdArtluxBD As ClassOCProdArtlux
Dim bAchou As Boolean
Dim objMovEstoque As ClassMovEstoque, lCodigo As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colRatreamentoMovto As New Collection
Dim objProduto As New ClassProduto

On Error GoTo Erro_OrdensDeCorteProdArtlux_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206662
    Next
    
    objProduto.sCodigo = objOCArtlux.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    'Com base na coleção da tela separa o que é novo e o que é alterado e o que vai gerar movimento de estoque
    For Each objOCProdArtlux In objOCArtlux.colItens
        If objOCProdArtlux.lNumIntDoc = 0 Then
            colNovos.Add objOCProdArtlux
            If objOCProdArtlux.dQuantidadeProd > QTDE_ESTOQUE_DELTA Then colMovEst.Add objOCProdArtlux
        Else
            bAchou = False
            For Each objOCProdArtluxBD In objOCArtluxBD.colItens
                If objOCProdArtluxBD.lNumIntDoc = objOCProdArtlux.lNumIntDoc Then
                    If Abs(objOCProdArtluxBD.dQuantidadeProd - objOCProdArtlux.dQuantidadeProd) > QTDE_ESTOQUE_DELTA And objOCProdArtluxBD.dQuantidadeProd > QTDE_ESTOQUE_DELTA Then gError 206663 'Não pode ter alterado a quantidade do Movimento de estoque
                    If Abs(objOCProdArtluxBD.dQuantidadeProd - objOCProdArtlux.dQuantidadeProd) > QTDE_ESTOQUE_DELTA Then colMovEst.Add objOCProdArtlux
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then gError 206664 ' Se tem NumIntDoc mas ele não está no BD -> Erro
            colAlterados.Add objOCProdArtlux
        End If
    Next
    
    'Com base na coleção do BD busca o que foi excluido, se tiver movimento de estoque vai ter que excluir também
    For Each objOCProdArtluxBD In objOCArtluxBD.colItens
        bAchou = False
        For Each objOCProdArtlux In objOCArtlux.colItens
            If objOCProdArtluxBD.lNumIntDoc = objOCProdArtlux.lNumIntDoc Then
                 bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            colExcluidos.Add objOCProdArtluxBD
        End If
    Next
    
    'Grava os movimentos de estoque novos
    For Each objOCProdArtlux In colMovEst
    
        lErro = CF("MovEstoque_Automatico_EmTransacao", giFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Set objMovEstoque = New ClassMovEstoque
        objMovEstoque.iFilialEmpresa = giFilialEmpresa
        objMovEstoque.dtData = objOCProdArtlux.dtDataFimMontagem
        objMovEstoque.lCodigo = lCodigo
    
        Set objItemMovEstoque = objMovEstoque.colItens.Add(0, MOV_EST_ENTRADA_OC, 0, 0, objOCArtlux.sProduto, "", objProduto.sSiglaUMEstoque, objOCProdArtlux.dQuantidadeProd, 1, "", 0, "", 0, "", "", "", "", 0, colRatreamentoMovto, Nothing, DATA_NULA)
        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
                           
        objMovEstoque.sObservacao = "Usuário: " & objOCProdArtlux.sUsuMontagem & " Saída: " & Formata_Estoque(objOCProdArtlux.dQuantidadePreProd) & " Entrada: " & Formata_Estoque(objOCProdArtlux.dQuantidadeProd)

        'Grava no BD (inclusive os dados contabeis) (contabilidade)
        lErro = CF("MovEstoque_Interno_Grava0", objMovEstoque, Nothing)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOCProdArtlux.lNumIntDocMovEst = objItemMovEstoque.lNumIntDoc

    Next
    
    For Each objOCProdArtlux In colNovos
    
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_ORDENSDECORTEPRODARTLUX", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objOCProdArtlux.lNumIntDoc = lNumIntDoc
        objOCProdArtlux.lNumIntDocOC = objOCArtlux.lNumIntDoc

        lErro = Comando_Executar(alComando(0), "INSERT INTO OrdensDeCorteProdArtlux( NumIntDoc, NumIntDocOC, Seq, NumIntDocMovEst, " & _
                "UsuMontagem, DataIniMontagem, DataFimMontagem, QuantidadePreProd, QuantidadeProd) VALUES (?,?,?,?,?,?,?,?,?)", _
                objOCProdArtlux.lNumIntDoc, objOCProdArtlux.lNumIntDocOC, objOCProdArtlux.iSeq, objOCProdArtlux.lNumIntDocMovEst, _
                objOCProdArtlux.sUsuMontagem, objOCProdArtlux.dtDataIniMontagem, objOCProdArtlux.dtDataFimMontagem, objOCProdArtlux.dQuantidadePreProd, objOCProdArtlux.dQuantidadeProd)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206665
    
    Next
    
    For Each objOCProdArtlux In colAlterados

        lErro = Comando_Executar(alComando(1), "UPDATE OrdensDeCorteProdArtlux SET Seq = ?, NumIntDocMovEst= ? , UsuMontagem= ? , DataIniMontagem= ? , DataFimMontagem= ? , " & _
                "QuantidadePreProd= ? , QuantidadeProd= ? WHERE NumIntDoc = ?", objOCProdArtlux.iSeq, objOCProdArtlux.lNumIntDocMovEst, objOCProdArtlux.sUsuMontagem, objOCProdArtlux.dtDataIniMontagem, objOCProdArtlux.dtDataFimMontagem, _
                objOCProdArtlux.dQuantidadePreProd, objOCProdArtlux.dQuantidadeProd, objOCProdArtlux.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 206666

    Next

    For Each objOCProdArtlux In colExcluidos

        lErro = Comando_Executar(alComando(2), "DELETE FROM OrdensDeCorteProdArtlux WHERE NumIntDoc = ?", objOCProdArtlux.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 206667
        
        If objOCProdArtlux.lNumIntDocMovEst <> 0 Then
            
            Set objItemMovEstoque = New ClassItemMovEstoque
            
            objItemMovEstoque.lNumIntDoc = objOCProdArtlux.lNumIntDocMovEst
        
            lErro = CF("MovEstoque_Le_NumIntDoc", objItemMovEstoque)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            Set objMovEstoque = New ClassMovEstoque
        
            objMovEstoque.lCodigo = objItemMovEstoque.lCodigo
            objMovEstoque.iFilialEmpresa = objItemMovEstoque.iFilialEmpresa
             
            'trata a exclusão do movimento de estoque
            lErro = CF("MovEstoque_Exclui_Trans", objMovEstoque, Nothing)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdensDeCorteProdArtlux_Grava = SUCESSO

    Exit Function

Erro_OrdensDeCorteProdArtlux_Grava:

    OrdensDeCorteProdArtlux_Grava = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206662
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 206663
        
        Case 206664

        Case 206665
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ORDENSDECORTEPRODARTLUX", gErr)

        Case 206666
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDENSDECORTEPRODARTLUX", gErr)

        Case 206667
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_ORDENSDECORTEPRODARTLUX", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206668)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function UsuProdArtlux_Exclui(ByVal objUsuProdArtlux As ClassUsuProdArtlux) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_UsuProdArtlux_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206669

    lErro = UsuProdArtlux_Exclui_EmTrans(objUsuProdArtlux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 206670

    UsuProdArtlux_Exclui = SUCESSO

    Exit Function

Erro_UsuProdArtlux_Exclui:

    UsuProdArtlux_Exclui = gErr

    Select Case gErr

        Case 206669
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 206670
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206671)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function UsuProdArtlux_Exclui_EmTrans(ByVal objUsuProdArtlux As ClassUsuProdArtlux) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_UsuProdArtlux_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206672
    Next

    'Le a tabelaUsuProdArtlux
    lErro = Comando_ExecutarPos(alComando(0), "SELECT AcessoCorte FROM UsuProdArtlux WHERE FilialEmpresa= ?  AND CodUsuario= ? ", _
                0, iAux, objUsuProdArtlux.iFilialEmpresa, objUsuProdArtlux.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 206673

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206674

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 206675

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 206676

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM UsuProdArtlux", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 206677

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    UsuProdArtlux_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_UsuProdArtlux_Exclui_EmTrans:

    UsuProdArtlux_Exclui_EmTrans = gErr

    Select Case gErr

        Case 206672
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206673, 206674
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUPRODARTLUX", gErr)

        Case 206675
            Call Rotina_Erro(vbOKOnly, "ERRO_USUPRODARTLUX_NAO_CADASTRADO", gErr, objUsuProdArtlux.iFilialEmpresa, objUsuProdArtlux.sCodUsuario)

        Case 206676
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_USUPRODARTLUX", gErr)

        Case 206677
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUPRODARTLUX", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206678)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'Exclusão
Public Function OrdensDeCorteArtlux_Exclui(ByVal objOCArtlux As ClassOCArtlux) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_OrdensDeCorteArtlux_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206764

    lErro = OrdensDeCorteArtlux_Exclui_EmTrans(objOCArtlux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 206765

    OrdensDeCorteArtlux_Exclui = SUCESSO

    Exit Function

Erro_OrdensDeCorteArtlux_Exclui:

    OrdensDeCorteArtlux_Exclui = gErr

    Select Case gErr

        Case 206764
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 206765
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206766)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Private Function OrdensDeCorteArtlux_Exclui_EmTrans(ByVal objOCArtlux As ClassOCArtlux) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_OrdensDeCorteArtlux_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206767
    Next
    
    'Le a tabelaOrdensDeCorteArtlux
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Seq FROM OrdensDeCorteProdArtlux WHERE NumIntDocOC = ? ", _
                0, iAux, objOCArtlux.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 206775

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206776

    If lErro = AD_SQL_SUCESSO Then gError 206777

    'Le a tabelaOrdensDeCorteArtlux
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM OrdensDeCorteArtlux WHERE NumIntDoc = ? ", _
                0, iAux, objOCArtlux.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 206768

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206769

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 206770

        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM OrdensDeCorteArtlux ", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 206771
        

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdensDeCorteArtlux_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_OrdensDeCorteArtlux_Exclui_EmTrans:

    OrdensDeCorteArtlux_Exclui_EmTrans = gErr

    Select Case gErr

        Case 206767
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206768, 206769
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDECORTEARTLUX", gErr)

        Case 206770
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ORDENSDECORTEARTLUX", gErr)

        Case 206771
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORDENSDECORTEARTLUX", gErr)

        Case 206775, 206776
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDECORTEPRODARTLUX", gErr)

        Case 206777
            Call Rotina_Erro(vbOKOnly, "ERRO_OC_EM_PRODUCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206772)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function GV_Exporta(ByVal sNomeArq As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Gera txt para uso pelo Geerenciador de Vendas comprado pela Artlux
'http://www.gerenciadordevendas.com.br/produtos/mapa/como_implantarMapa.htm

Dim lErro As Long, lComando As Long, sReg As String, sCidade As String, sUF As String, sRegiao As String, sDescProd As String
Dim lNumNotaFiscal As Long, iFilialEmpresa As Integer, dtDataEmissao As Date, dValorItem As Double, dQtde As Double
Dim sGerenteVendas As String, sNomeVendedor As String, sCanalDistrib As String, sEndereco As String, sCEP As String, sCEPAux As String
Dim sTipoDeCouro As String, sCor As String, sTam As String, sGrupo As String, sNomeRedCli As String, sProduto As String, sCgc As String, sCGCAux As String
Dim iFaturamento As Integer, sTelefone1 As String, iTamTel As Integer

On Error GoTo Erro_GV_Exporta

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 201542
    
    'Cria o arquivo de remessa
    Open sNomeArq For Output As #1
    
    sGerenteVendas = "ND"
    sCanalDistrib = "ND"
    
    sCidade = String(STRING_CIDADE, 0)
    sUF = String(STRING_ESTADOS_SIGLA, 0)
    sNomeVendedor = String(STRING_MAXIMO, 0)
    sTipoDeCouro = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sCor = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sTam = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sGrupo = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sDescProd = String(STRING_PRODUTO_DESCRICAO, 0)
    sNomeRedCli = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sCgc = String(STRING_CGC, 0)
    sEndereco = String(STRING_ENDERECO_COMPLEMENTO, 0)
    sRegiao = String(STRING_MAXIMO, 0)
    sCEP = String(STRING_CEP, 0)
    sTelefone1 = String(STRING_TELEFONE, 0)
    
    If dtDataInicial <> DATA_NULA Or dtDataFinal <> DATA_NULA Then
    
        If dtDataFinal = DATA_NULA Then dtDataFinal = gdtDataAtual
        lErro = Comando_Executar(lComando, "SELECT Faturamento, NumNotaFiscal, FilialEmpresa, UF, Cidade, DescRegiao, NomeVendedor, TipoDeCouro, Cor, Tam, Grupo, DescProd, NomeRedCli, ProdutoRef, DataEmissao, (CASE WHEN Qtde = 0 THEN PrecoUnitario ELSE Qtde*PrecoUnitario END)-ValorDesconto AS ValorItem, Qtde, CGC, Endereco, CEP, Telefone1 FROM GV_Vendas WHERE Faturamento IN (1) AND DataEmissao >= ? AND DataEmissao <= ? ORDER BY DataEmissao, FilialEmpresa, NumNotaFiscal", iFaturamento, lNumNotaFiscal, iFilialEmpresa, sUF, sCidade, sRegiao, sNomeVendedor, sTipoDeCouro, sCor, sTam, sGrupo, sDescProd, sNomeRedCli, sProduto, dtDataEmissao, dValorItem, dQtde, sCgc, sEndereco, sCEP, sTelefone1, dtDataInicial, dtDataFinal)
    
    Else
        
        lErro = Comando_Executar(lComando, "SELECT Faturamento, NumNotaFiscal, FilialEmpresa, UF, Cidade, DescRegiao, NomeVendedor, TipoDeCouro, Cor, Tam, Grupo, DescProd, NomeRedCli, ProdutoRef, DataEmissao, (CASE WHEN Qtde = 0 THEN PrecoUnitario ELSE Qtde*PrecoUnitario END)-ValorDesconto AS ValorItem, Qtde, CGC, Endereco, CEP, Telefone1 FROM GV_Vendas WHERE Faturamento IN (1) ORDER BY DataEmissao, FilialEmpresa, NumNotaFiscal", iFaturamento, lNumNotaFiscal, iFilialEmpresa, sUF, sCidade, sRegiao, sNomeVendedor, sTipoDeCouro, sCor, sTam, sGrupo, sDescProd, sNomeRedCli, sProduto, dtDataEmissao, dValorItem, dQtde, sCgc, sEndereco, sCEP, sTelefone1)
        
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 201543
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201544
        
    Do While lErro <> AD_SQL_SEM_DADOS

        sReg = FormataCpoTexto_GV(CStr(lNumNotaFiscal), 6)
        sReg = sReg & FormataCpoTexto_GV(IIf(iFilialEmpresa = 1, "ARTLUX", "RELEVO"), 20)
        sReg = sReg & FormataCpoTexto_GV(sUF, 2)
        sReg = sReg & FormataCpoTexto_GV(sCidade, 30)
        sReg = sReg & FormataCpoTexto_GV(sRegiao, 20)
        sReg = sReg & FormataCpoTexto_GV(sGerenteVendas, 20)
        sReg = sReg & FormataCpoTexto_GV(sNomeVendedor, 30)
        sReg = sReg & FormataCpoTexto_GV(sCanalDistrib, 20)
        sReg = sReg & FormataCpoTexto_GV(sTam, 20) 'segmento de mercado
        sReg = sReg & FormataCpoTexto_GV(sTipoDeCouro, 20) 'marca ou fornecedor
        sReg = sReg & FormataCpoTexto_GV(sCor, 20) 'linha
        sReg = sReg & FormataCpoTexto_GV(sGrupo, 20)
        
'        sReg = sReg & FormataCpoTexto_GV(sDescProd, 50)
        sReg = sReg & FormataCpoTexto_GV(left(sProduto, 5), 5)
        sReg = sReg & " "
        sReg = sReg & FormataCpoTexto_GV(sDescProd, 44)
                
        sReg = sReg & FormataCpoTexto_GV(left(sNomeRedCli, 36), 36)
        sReg = sReg & FormataCpoTexto_GV(Trim(sTelefone1), 14)
        
'        sReg = sReg & FormataCpoTexto_GV(sProduto, 20)
        sReg = sReg & FormataCpoTexto_GV("ND", 20)
        
        sReg = sReg & Format(dtDataEmissao, "ddmmyyyy")
        sReg = sReg & FormataCpoTexto_GV(Format(IIf(iFaturamento = 1, dValorItem, -dValorItem), "0.00"), 11)
        sReg = sReg & FormataCpoTexto_GV(Format(IIf(iFaturamento = 1, dValorItem, -dValorItem), "0.00"), 11)
        sReg = sReg & FormataCpoTexto_GV(Format(IIf(iFaturamento = 1, dQtde, -dQtde), "0"), 11)
        sReg = sReg & FormataCpoTexto_GV(Format(IIf(iFaturamento = 1, dQtde, -dQtde), "0.000"), 11)
        sReg = sReg & FormataCpoTexto_GV(Format(IIf(iFaturamento = 1, dQtde, -dQtde), "0.000"), 11)
        sReg = sReg & FormataCpoTexto_GV(Format(IIf(iFaturamento = 1, dQtde, -dQtde), "0.000"), 11)
        If Len(Trim(sCgc)) > 11 Then
            sCGCAux = right("0000" + sCgc, 14)
            sReg = sReg & "0" & left(sCGCAux, 8)
            sReg = sReg & Mid(sCGCAux, 9, 4)
        Else
            sCGCAux = right("00000000000" + sCgc, 11)
            sReg = sReg & left(sCGCAux, 9)
            sReg = sReg & "0001"
        End If
        sReg = sReg & FormataCpoTexto_GV(sEndereco, 50)
                
        sCEPAux = sCEP
        If Len(Trim(sCEPAux)) < 8 Then sCEPAux = left(sCEPAux + "00000000", 8)
        sCEPAux = right("00000000" & sCEPAux, 8)
        sReg = sReg & left(sCEPAux, 5) & "-" & right(sCEPAux, 3)
        
        Print #1, sReg
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201545
    
    Loop
    
    Close #1
    
    Call Comando_Fechar(lComando)

    GV_Exporta = SUCESSO
    
    Exit Function
    
Erro_GV_Exporta:

    GV_Exporta = gErr

    Select Case gErr
    
        Case 201542
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 201543 To 201545
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201541)

    End Select
    
    Close #1
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function FormataCpoTexto_GV(ByVal sTexto As String, ByVal iTam As Integer) As String
    FormataCpoTexto_GV = FormataCpoTexto(IIf(sTexto = "", "ND", sTexto), iTam)
End Function
