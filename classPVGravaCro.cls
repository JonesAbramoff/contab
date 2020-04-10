VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SEQ_TIPO_CARGA = 1
Const SEQ_TIPO_EMBRANCO = 2
Const SEQ_TIPO_OUTROS = 3
Const SEQ_TIPO_LINHASCOMPLEMENTO = 4
Const SEQ_TIPO_PRODACABADO = 5

Function OrdemDeProducao_Exclui_Usu(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao

Dim lErro As Long

On Error GoTo Erro_OrdemDeProducao_Exclui_Usu

    '#########################################
    'Inserido por Wagner - CROMATON 08/11/04
    lErro = OPFichaProcesso_Exclui(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 131144
    
    lErro = ItemOPCarga_Exclui(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 131179
    '#########################################

    OrdemDeProducao_Exclui_Usu = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui_Usu:

    OrdemDeProducao_Exclui_Usu = Err

    Select Case Err

        Case 131144, 131179

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Exit Function

End Function

Function OrdemDeProducao_Grava_Customizado(ByVal objOP As ClassOrdemDeProducao) As Long

Dim objItemOP As ClassItemOP
Dim lErro As Long

On Error GoTo Erro_OrdemDeProducao_Grava_Customizado

    lErro = OPFichaProcesso_Grava(objOP)
    If lErro <> SUCESSO Then gError 131179
    
    lErro = ItemOPCarga_Exclui(objOP)
    If lErro <> SUCESSO Then gError 131179

    For Each objItemOP In objOP.colItens
    
        lErro = ItemOPCarga_Insere(objItemOP)
        If lErro <> SUCESSO Then gError 131179
    
    Next
    
    OrdemDeProducao_Grava_Customizado = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava_Customizado:

    OrdemDeProducao_Grava_Customizado = gErr

    Select Case gErr

        Case 131179

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

'################################################
'FUNÇÕES NOVAS
Function OPFichaProcesso_Grava(objOP As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_OPFichaProcesso_Grava

    If objOP.objInfoUsu Is Nothing Then Exit Function

    'Abre Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131158
    Next

    'Le Contrato
    lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM OPFichaProcesso WHERE CodigoOP = ? AND FilialEmpresa = ? ", 0, iAux, objOP.sCodigo, objOP.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 131159

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131160

    'Se não encontrou insere, senão altera
    If lErro = AD_SQL_SEM_DADOS Then
        
        With objOP.objInfoUsu
               
            'Insere Ficha de Processo
            lErro = Comando_Executar(alComando(2), "INSERT INTO OPFichaProcesso(CodigoOP, FilialEmpresa, Mistura, TempoMistura, CargaDrays, Embalagem, " & _
                                "ProcessoAgua, ProcessoOleo, Amperagem, VelRotoresDe, VelRotoresAte, AberturaGate, TempCamaraZ1De, TempCamaraZ1Ate, TempCamaraZ2De, TempCamaraZ2Ate, " & _
                                "TempCamaraZ3De, TempCamaraZ3Ate, Tela, TempExtrusoraZ1De, TempExtrusoraZ1Ate, TempExtrusoraZ2De, TempExtrusoraZ2Ate, TempExtrusoraZ3De, TempExtrusoraZ3Ate, " & _
                                "TempExtrusoraZ4De, TempExtrusoraZ4Ate, TempExtrusoraZ5De, TempExtrusoraZ5Ate, TempExtrusoraZ6De, TempExtrusoraZ6Ate, TempExtrusoraZ7De, TempExtrusoraZ7Ate, " & _
                                "VelVariadorDe, VelVariadorAte, Observacao, Amostra " & _
                                ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                .sCodigoOP, .iFilialEmpresa, .iMistura, .iTempoMistura, .iCargaDrays, .iEmbalagem, _
                                .dProcessoAgua, .dProcessoOleo, .iAmperagem, .dVelRotoresDe, .dVelRotoresAte, .dAberturaGate, .dTempCamaraZDe(1), .dTempCamaraZAte(1), .dTempCamaraZDe(2), .dTempCamaraZAte(2), .dTempCamaraZDe(3), _
                                .dTempCamaraZAte(3), .sTela, .dTempExtrusoraZDe(1), .dTempExtrusoraZAte(1), .dTempExtrusoraZDe(2), .dTempExtrusoraZAte(2), .dTempExtrusoraZDe(3), .dTempExtrusoraZAte(3), .dTempExtrusoraZDe(4), .dTempExtrusoraZAte(4), _
                                .dTempExtrusoraZDe(5), .dTempExtrusoraZAte(5), .dTempExtrusoraZDe(6), .dTempExtrusoraZAte(6), .dTempExtrusoraZDe(7), .dTempExtrusoraZAte(7), .dVelVariadorDe, .dVelVariadorAte, .sObservacao, .sAmostra)
            If lErro <> AD_SQL_SUCESSO Then gError 131161
    
        End With
        
    Else

        With objOP.objInfoUsu

            'Altera Ficha de Processo
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE OPFichaProcesso SET Mistura=?, TempoMistura=?, CargaDrays=?, Embalagem=?, " & _
                                "ProcessoAgua=?, ProcessoOleo=?, Amperagem=?, VelRotoresDe=?, VelRotoresAte=?, AberturaGate=?, TempCamaraZ1De=?, TempCamaraZ1Ate=?, TempCamaraZ2De=?, TempCamaraZ2Ate=?, " & _
                                "TempCamaraZ3De=?, TempCamaraZ3Ate=?, Tela=?, TempExtrusoraZ1De=?, TempExtrusoraZ1Ate=?, TempExtrusoraZ2De=?, TempExtrusoraZ2Ate=?, TempExtrusoraZ3De=?, TempExtrusoraZ3Ate=?, " & _
                                "TempExtrusoraZ4De=?, TempExtrusoraZ4Ate=?, TempExtrusoraZ5De=?, TempExtrusoraZ5Ate=?, TempExtrusoraZ6De=?, TempExtrusoraZ6Ate=?, TempExtrusoraZ7De=?, TempExtrusoraZ7Ate=?, " & _
                                "VelVariadorDe=?, VelVariadorAte=?, Observacao=?, Amostra=? ", _
                                alComando(1), .iMistura, .iTempoMistura, .iCargaDrays, .iEmbalagem, _
                                .dProcessoAgua, .dProcessoOleo, .iAmperagem, .dVelRotoresDe, .dVelRotoresAte, .dAberturaGate, .dTempCamaraZDe(1), .dTempCamaraZAte(1), .dTempCamaraZDe(2), .dTempCamaraZAte(2), .dTempCamaraZDe(3), _
                                .dTempCamaraZAte(3), .sTela, .dTempExtrusoraZDe(1), .dTempExtrusoraZAte(1), .dTempExtrusoraZDe(2), .dTempExtrusoraZAte(2), .dTempExtrusoraZDe(3), .dTempExtrusoraZAte(3), .dTempExtrusoraZDe(4), .dTempExtrusoraZAte(4), _
                                .dTempExtrusoraZDe(5), .dTempExtrusoraZAte(5), .dTempExtrusoraZDe(6), .dTempExtrusoraZAte(6), .dTempExtrusoraZDe(7), .dTempExtrusoraZAte(7), .dVelVariadorDe, .dVelVariadorAte, .sObservacao, .sAmostra)
            If lErro <> AD_SQL_SUCESSO Then gError 131162

        End With

    End If
 
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OPFichaProcesso_Grava = SUCESSO

    Exit Function

Erro_OPFichaProcesso_Grava:

    OPFichaProcesso_Grava = gErr

        Select Case gErr

        Case 131158
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131159, 131160
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPFICHAPROCESSO", gErr)

        Case 131161
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OPFICHAPROCESSO", gErr, objOP.sCodigo)

        Case 131162
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_OPFICHAPROCESSO", gErr, objOP.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function OPFichaProcesso_Exclui(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Função que exclui uma ordem de produção a partir do código passado em objOrdemDeProducao

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iFilial As Integer

On Error GoTo Erro_OPFichaProcesso_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comando
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131145
    Next

    'tenta ler se existe a Ficha de Processo
    lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM OPFichaProcesso WHERE FilialEmpresa = ? AND CodigoOP = ? ", 0, iFilial, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 131146

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131147

    If lErro <> AD_SQL_SEM_DADOS Then

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 131149
    
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM OPFichaProcesso", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 131150
        
    End If

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OPFichaProcesso_Exclui = SUCESSO

    Exit Function

Erro_OPFichaProcesso_Exclui:

    OPFichaProcesso_Exclui = gErr

    Select Case gErr

        Case 131145
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131146, 131147
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPFICHAPROCESSO", gErr)

        Case 131148
            Call Rotina_Erro(vbOKOnly, "ERRO_OPFICHAPROCESSO_NAO_CADASTRADO", gErr, objOrdemDeProducao.sCodigo)

        Case 131149
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OPFICHAPROCESSO", gErr)

        Case 131150
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPFICHAPROCESSO", gErr, objOrdemDeProducao.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemOPCarga_Insere(objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lAux As Long

On Error GoTo Erro_ItemOPCarga_Insere

    'Se não existe carga sai da função
    If objItemOP.objInfoUsu Is Nothing Then Exit Function
    
    'Se a carga foi excluída sai da função
    If objItemOP.objInfoUsu.iItem = 0 Then Exit Function

    'Abre Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131152
    Next

    'Le a Carga
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntItemOP FROM ItensOPCarga WHERE NumIntItemOP= ? ", 0, lAux, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 131153

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131154

    'Se não encontrou insere, senão altera
    If lErro = AD_SQL_SEM_DADOS Then
        
        With objItemOP.objInfoUsu
        
            .lNumIntItemOP = objItemOP.lNumIntDoc
               
            'Insere a Carga
            lErro = Comando_Executar(alComando(2), "INSERT INTO ItensOPCarga(NumIntItemOP, ProdutoBase, KgProdBase, KgCarga, QtdCarga) VALUES (?,?,?,?,?)", _
                                                                       .lNumIntItemOP, .sProdutoBase, .dKgProdBase, .dKgCarga, .dQtdCarga)
            If lErro <> AD_SQL_SUCESSO Then gError 131155
    
        End With
        
    Else

        With objItemOP.objInfoUsu

            'Altera Insere a Carga
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensOPCarga SET ProdutoBase=?, KgProdBase=?, KgCarga=?, QtdCarga=? ", alComando(1), _
                                                    .sProdutoBase, .dKgProdBase, .dKgCarga, .dQtdCarga)
            If lErro <> AD_SQL_SUCESSO Then gError 131156

        End With

    End If
    
    lErro = ItemOPCargaInsumo_Exclui(objItemOP.objInfoUsu)
    If lErro <> SUCESSO Then gError 131157
    
    lErro = ItemOPCargaInsumo_Insere(objItemOP.objInfoUsu)
    If lErro <> SUCESSO Then gError 131172
 
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemOPCarga_Insere = SUCESSO

    Exit Function

Erro_ItemOPCarga_Insere:

    ItemOPCarga_Insere = gErr

        Select Case gErr

        Case 131152
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131153, 131154
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)

        Case 131155
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMOPCARGA", gErr, objItemOP.iItem)

        Case 131156
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITEMOPCARGA", gErr, objItemOP.iItem)

        Case 131157, 131172

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ItemOPCargaInsumo_Insere(objItemOPCarga As ClassItemOPCarga) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objItemOPCargaInsumo As ClassItemOPCargaInsumo
Dim lNumProx As Long

On Error GoTo Erro_ItemOPCargaInsumo_Insere

    'Abre Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131163
    Next

    For Each objItemOPCargaInsumo In objItemOPCarga.colInsumos
        
        With objItemOPCargaInsumo
               
            'Obtem identificado (NumIntDoc)
            lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_PROX_ITENSOPCARGAINSUMO", lNumProx)
            If lErro <> SUCESSO Then gError 131164
            
            .lNumIntDoc = lNumProx
            .lNumIntItemOP = objItemOPCarga.lNumIntItemOP
               
            'Insere a Carga
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensOPCargaInsumo(KgCarga,NumIntDoc,NumIntItemOP,Produto, Seq, Fornecedor, Lote) VALUES (?,?,?,?,?,?,?)", _
                                                                     .dKgCarga, .lNumIntDoc, .lNumIntItemOP, .sProduto, .iSeq, .lFornecedor, .sLote)
            If lErro <> AD_SQL_SUCESSO Then gError 131165
    
        End With
        
    Next
 
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemOPCargaInsumo_Insere = SUCESSO

    Exit Function

Erro_ItemOPCargaInsumo_Insere:

    ItemOPCargaInsumo_Insere = gErr

    Select Case gErr

        Case 131163
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131164
        
        Case 131165
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMOPCARGA", gErr, objItemOPCarga.iItem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ItemOPCargaInsumo_Exclui(ByVal objItemOPCarga As ClassItemOPCarga) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lAux As Long

On Error GoTo Erro_ItemOPCargaInsumo_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comando
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131166
    Next

    'tenta ler se existe os Insumos
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensOPCargaInsumo WHERE NumIntItemOP = ? ", 0, lAux, objItemOPCarga.lNumIntItemOP)
    If lErro <> AD_SQL_SUCESSO Then gError 131167

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131168

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 131170
    
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensOPCargaInsumo", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 131171

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131169
    
    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPCargaInsumo_Exclui = SUCESSO

    Exit Function

Erro_ItemOPCargaInsumo_Exclui:

    ItemOPCargaInsumo_Exclui = gErr

    Select Case gErr

        Case 131166
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131167, 131168, 131169
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGAINSUMO", gErr)

        Case 131170
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ITEMOPCARGAINSUMO", gErr)

        Case 131171
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOPCARGAINSUMO", gErr, objItemOPCarga.iItem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemOPCarga_Exclui(ByVal objOP As ClassOrdemDeProducao) As Long
'Exclui todos itens de Carga de uma OP

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lAux As Long
Dim objItemOP As ClassItemOP

On Error GoTo Erro_ItemOPCarga_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comando
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131173
    Next

    For Each objItemOP In objOP.colItens

        'tenta ler se existe as Cargas
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntItemOP FROM ItensOPCarga WHERE NumIntItemOP = ? ", 0, lAux, objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 131174
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131175
        
        If lErro <> AD_SQL_SEM_DADOS Then

            lErro = Comando_LockExclusive(alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 131176
        
            lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ItensOPCarga", alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 131177
            
            If objItemOP.objInfoUsu Is Nothing Then
            
                Set objItemOP.objInfoUsu = New ClassItemOPCarga
                
                objItemOP.objInfoUsu.lNumIntItemOP = objItemOP.lNumIntDoc
                
            End If
            
            lErro = ItemOPCargaInsumo_Exclui(objItemOP.objInfoUsu)
            If lErro <> SUCESSO Then gError 131178
            
        End If
    
    Next

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPCarga_Exclui = SUCESSO

    Exit Function

Erro_ItemOPCarga_Exclui:

    ItemOPCarga_Exclui = gErr

    Select Case gErr

        Case 131173
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131174, 131175
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)

        Case 131176
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ITEMOPCARGA", gErr)

        Case 131177
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOPCARGA", gErr, lAux)

        Case 131178

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelOPCarga_Prepara(ByVal objOP As ClassOrdemDeProducao, lNumIntRel As Long, Optional ByVal bEspelho As Boolean = False) As Long
'Monta a tabela temporária para o relatório de OP

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(1 To 11) As Long
Dim colOPCarga As New Collection

On Error GoTo Erro_RelOPCarga_Prepara

    'Abertura de transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 140350
    
   'Abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140351
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "MATConfig", "NUMINTREL_RELOPCARGA", lNumIntRel)
    If lErro <> SUCESSO Then gError 140352
    
    'Insere os Itens que fazem parte da Carga
    lErro = RelOPCarga_Le_ItensCarga(objOP, colOPCarga, alComando(1))
    If lErro <> SUCESSO Then gError 140353
    
    'Insere os itens que foram empenhados e não fazem parte da carga
    lErro = RelOPCarga_Le_Empenhos_Fora_Carga(objOP, colOPCarga, alComando(2))
    If lErro <> SUCESSO Then gError 140354
       
    'Insere os Itens requisitados que não fizeram parte da carga (Produto/Fornecedor/Lote)
    lErro = RelOPCarga_Le_ReqProd_Fora_Carga_Empenho(objOP, colOPCarga, alComando(3), alComando(4))
    If lErro <> SUCESSO Then gError 140355
    
    If Not bEspelho Then
        'Insere uma linha em branco para cada Item da OP
        lErro = RelOPCarga_Insere_Linha_EmBranco(objOP, colOPCarga)
        If lErro <> SUCESSO Then gError 140356
    End If
    
    'Le as entradas da produção para cada item da OP e busca o responsável
    lErro = RelOPCarga_Le_ProducaoEntrada(objOP, alComando(5))
    If lErro <> SUCESSO Then gError 140357
    
    'Retira da coleção os itens que forem embalagens
    lErro = RelOPCarga_Retira_Embalagens(colOPCarga)
    If lErro <> SUCESSO Then gError 140358
    
    'Le os Lotes dos Produtos/Fornecedor/Lote da Coleção
    lErro = RelOPCarga_Le_Lote(colOPCarga, alComando(6), alComando(7))
    If lErro <> SUCESSO Then gError 140359
    
    'Le a quantidade requisitada de cada item da coleção
    lErro = RelOPCarga_Le_Qtd_Requisitada(colOPCarga, alComando(8), alComando(9))
    If lErro <> SUCESSO Then gError 140360
    
    'Insere uma linha em branco para cada Item da OP
    lErro = RelOPCarga_Insere_Linha_EmBranco2(objOP, colOPCarga, bEspelho)
    If lErro <> SUCESSO Then gError 140356
    
    If bEspelho Then
        lErro = RelOPCarga_Espelho(objOP, colOPCarga)
        If lErro <> SUCESSO Then gError 140361
    End If
    
    'Replica as informações do item da OP para os itens da coleção
    lErro = RelOPCarga_Replica_Dados(objOP, colOPCarga, alComando(11))
    If lErro <> SUCESSO Then gError 140361
      
    'Insere os Itens na tabela temporária
    lErro = RelOPCarga_Insere(lNumIntRel, colOPCarga, alComando(10), bEspelho)
    If lErro <> SUCESSO Then gError 140362
    
    'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 140363
    
    RelOPCarga_Prepara = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Prepara:

    RelOPCarga_Prepara = gErr
    
    Select Case gErr
        
        Case 140350
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 140351
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140352 To 140362

        Case 140363
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
        
End Function

Private Function RelOPCarga_Le_ItensCarga(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection, ByVal lComando As Long) As Long
'Le os itens que fizeram parte da Carga e Insere na Coleção

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objItemOPCarga As ClassItemOPCarga
Dim objItemOPCargaInsumo As ClassItemOPCargaInsumo
Dim objOPCarga As ClassOPCarga
Dim objFornecedor As ClassFornecedor
Dim iFilialForn As Integer
Dim bbeneficiamento As Boolean

On Error GoTo Erro_RelOPCarga_Le_ItensCarga

    'Para cada item da OP
    For Each objItemOP In objOP.colItens

        'Para cada Item da Carga
        Set objItemOPCarga = objItemOP.objInfoUsu
            
        'Para cada Insumo da Carga
        For Each objItemOPCargaInsumo In objItemOPCarga.colInsumos
    
            Set objOPCarga = New ClassOPCarga
            
            objOPCarga.dCarga = objItemOPCargaInsumo.dKgCarga
            objOPCarga.dQuantidade = objItemOPCarga.dQtdCarga * objItemOPCargaInsumo.dKgCarga
            objOPCarga.dQtdCarga = objItemOPCarga.dQtdCarga
            objOPCarga.iSeq = SEQ_TIPO_CARGA
            objOPCarga.iOrd = objItemOPCargaInsumo.iSeq
            objOPCarga.lItemOP = objItemOP.lNumIntDoc
            objOPCarga.sProduto = objItemOPCargaInsumo.sProduto
            objOPCarga.sCodigoOP = objOP.sCodigo
            objOPCarga.iFilialEmpresa = objOP.iFilialEmpresa
            'objOPCarga.lFornecedor = objItemOPCargaInsumo.lFornecedor
            objOPCarga.sLote = objItemOPCargaInsumo.sLote
            
            Set objFornecedor = New ClassFornecedor

            If Len(Trim(objOPCarga.sLote)) > 0 Then
            
                If objItemOP.iBeneficiamento = MARCADO Then
                    bbeneficiamento = True
                Else
                    bbeneficiamento = False
                End If
            
                lErro = RelOPCarga_Le_Fornecedor_Lote(objOPCarga.sProduto, objOPCarga.sLote, objFornecedor, lComando, iFilialForn, bbeneficiamento)
                If lErro <> SUCESSO Then gError 141774
                
                objOPCarga.lFornecedor = objFornecedor.lCodigo
                objOPCarga.sNomeFornecedor = objFornecedor.sNomeReduzido
                
            End If
    
'
'            If objOPCarga.lFornecedor <> 0 Then
'
'                objFornecedor.lCodigo = objOPCarga.lFornecedor
'                lErro = CF("Fornecedor_Le", objFornecedor)
'                If lErro <> SUCESSO And lErro <> 12729 Then gError 141774
'
'                objOPCarga.sNomeFornecedor = objFornecedor.sNomeReduzido
'
'            End If
            
            'Insere na coleção
            colOPCarga.Add objOPCarga
    
        Next
        
    Next
    
    RelOPCarga_Le_ItensCarga = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Le_ItensCarga:

    RelOPCarga_Le_ItensCarga = gErr
    
    Select Case gErr
    
        Case 141774
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
End Function

Private Function RelOPCarga_Le_Empenhos_Fora_Carga(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection, ByVal lComando As Long) As Long
'Le os itens que fizeram parte da Carga e Insere na Coleção

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objOPCarga As ClassOPCarga
Dim sSQL As String
Dim sProduto As String
Dim iOrd  As Integer

On Error GoTo Erro_RelOPCarga_Le_Empenhos_Fora_Carga

    For Each objItemOP In objOP.colItens
    
        sProduto = String(STRING_PRODUTO, 0)

        sSQL = "SELECT E.Produto FROM Empenho AS E WHERE E.NumIntDocItemOP = ? AND E.Produto NOT IN (SELECT I.Produto FROM ItensOPCargaInsumo AS I WHERE I.NumIntItemOP = E.NumIntDocItemOP)"
    
        'Busca os itens empenhados
        lErro = Comando_Executar(lComando, sSQL, sProduto, objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 140337
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140338
    
        iOrd = 0
        
        Do While lErro <> AD_SQL_SEM_DADOS
    
            Set objOPCarga = New ClassOPCarga
    
            iOrd = iOrd + 1
    
            objOPCarga.iSeq = SEQ_TIPO_OUTROS
            objOPCarga.iOrd = iOrd
            objOPCarga.lItemOP = objItemOP.lNumIntDoc
            objOPCarga.sProduto = sProduto
            objOPCarga.sCodigoOP = objOP.sCodigo
            objOPCarga.iFilialEmpresa = objOP.iFilialEmpresa
    
            'Insere na Coleção
            colOPCarga.Add objOPCarga
    
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140339
    
        Loop

    Next
    
    RelOPCarga_Le_Empenhos_Fora_Carga = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Le_Empenhos_Fora_Carga:

    RelOPCarga_Le_Empenhos_Fora_Carga = gErr
    
    Select Case gErr
    
        Case 140337 To 140339
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO2", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
End Function

Private Function RelOPCarga_Le_ReqProd_Fora_Carga_Empenho(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Le os Movimentos de Estoque para obter os itens de Requisição de produção que não estão na colOPCarga

Dim lErro As Long
Dim sProduto As String
Dim objOPCarga As ClassOPCarga
Dim iOrd As Integer
Dim bAchou As Boolean
Dim objItemOP As ClassItemOP
Dim objFornecedor As ClassFornecedor
Dim sLote As String
Dim iFilialForn As Integer
Dim bbeneficiamento As Boolean
Dim iTipo As Integer

On Error GoTo Erro_RelOPCarga_Le_ReqProd_Fora_Carga_Empenho
    
    For Each objItemOP In objOP.colItens
    
        'Aloca espaço no buffer
        sProduto = String(STRING_PRODUTO, 0)
        sLote = String(STRING_LOTE_RASTREAMENTO, 0)
           
        'Le as Requisições para produção
        lErro = Comando_Executar(lComando1, "SELECT ME.Produto, L.Lote,ME.TipoMov FROM RastreamentoLote AS L INNER JOIN RastreamentoMovto AS R ON L.NumIntDoc = R.NumIntDocLote RIGHT OUTER JOIN MovimentoEstoque AS ME ON R.NumIntDocOrigem = ME.NumIntDoc " & _
                                            "WHERE ME.NumIntDocOrigem = ? AND ME.TipoMov IN (?,?) AND ME.NumIntDocEst=0 AND (R.TipoDocOrigem = 0 OR R.TipoDocOrigem IS NULL) AND ME.TipoNumIntDocOrigem=5 ORDER BY ME.Produto ", _
                                            sProduto, sLote, iTipo, objItemOP.lNumIntDoc, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO)
        If lErro <> AD_SQL_SUCESSO Then gError 132123
           
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132124
                        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objFornecedor = New ClassFornecedor
            
            If iTipo = MOV_EST_REQ_PRODUCAO_BENEF3 Then
                bbeneficiamento = True
            Else
                bbeneficiamento = False
            End If
            
            lErro = RelOPCarga_Le_Fornecedor_Lote(sProduto, sLote, objFornecedor, lComando2, iFilialForn, bbeneficiamento)
            If lErro <> SUCESSO Then gError 140364

            bAchou = False
            
            'Verifica se o Item requisitado já está na coleção
            iOrd = 0
            For Each objOPCarga In colOPCarga
                If UCase(sProduto) = UCase(objOPCarga.sProduto) And objOPCarga.lItemOP = objItemOP.lNumIntDoc And UCase(objOPCarga.sLote) = UCase(sLote) Then bAchou = True
                If objOPCarga.iSeq = SEQ_TIPO_OUTROS Then iOrd = iOrd + 1
            Next
            
            'Senão não está na coleção, Insere.
            If bAchou = False Then
            
                Set objOPCarga = New ClassOPCarga
                
                iOrd = iOrd + 1
                
                objOPCarga.lFornecedor = objFornecedor.lCodigo
                objOPCarga.sNomeFornecedor = objFornecedor.sNomeReduzido
                objOPCarga.iFilialEmpresa = objItemOP.iFilialEmpresa
                objOPCarga.sCodigoOP = objItemOP.sCodigo
                objOPCarga.lItemOP = objItemOP.lNumIntDoc
                objOPCarga.sProduto = sProduto
                objOPCarga.iSeq = SEQ_TIPO_OUTROS
                objOPCarga.iOrd = iOrd
                objOPCarga.sLote = sLote
                
                colOPCarga.Add objOPCarga
            
            End If
                     
            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132125
        
        Loop
           
    Next
    
    RelOPCarga_Le_ReqProd_Fora_Carga_Empenho = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Le_ReqProd_Fora_Carga_Empenho:
    
    RelOPCarga_Le_ReqProd_Fora_Carga_Empenho = gErr
        
    Select Case gErr
    
        Case 132123 To 132125
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)
            
        Case 140364
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

Private Function RelOPCarga_Insere_Linha_EmBranco(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection) As Long
'Le os itens que fizeram parte da Carga e Insere na Coleção

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objOPCarga As ClassOPCarga

On Error GoTo Erro_RelOPCarga_Insere_Linha_EmBranco

    'Para cada item da OP
    For Each objItemOP In objOP.colItens

        Set objOPCarga = New ClassOPCarga

        objOPCarga.iSeq = SEQ_TIPO_EMBRANCO
        objOPCarga.iOrd = 1
        objOPCarga.lItemOP = objItemOP.lNumIntDoc
        objOPCarga.sCodigoOP = objOP.sCodigo
        objOPCarga.iFilialEmpresa = objOP.iFilialEmpresa

        'Insere na Coleção
        colOPCarga.Add objOPCarga
        
    Next
    
    RelOPCarga_Insere_Linha_EmBranco = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Insere_Linha_EmBranco:

    RelOPCarga_Insere_Linha_EmBranco = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
End Function

Private Function RelOPCarga_Insere_Linha_EmBranco2(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection, ByVal bEspelho As Boolean) As Long
'Le os itens que fizeram parte da Carga e Insere na Coleção

Dim lErro As Long
Dim iQtdLinhas As Integer
Dim iIndice As Integer
Dim objItemOP As ClassItemOP
Dim objOPCarga As ClassOPCarga

On Error GoTo Erro_RelOPCarga_Insere_Linha_EmBranco2

    'Para cada item da OP
    For Each objItemOP In objOP.colItens
    
        iQtdLinhas = 0
    
        If bEspelho Then

            For iIndice = 1 To 6

                Set objOPCarga = New ClassOPCarga
        
                objOPCarga.iSeq = SEQ_TIPO_LINHASCOMPLEMENTO
                objOPCarga.iOrd = 1
                objOPCarga.lItemOP = objItemOP.lNumIntDoc
                objOPCarga.sCodigoOP = objOP.sCodigo
                objOPCarga.iFilialEmpresa = objOP.iFilialEmpresa
        
                'Insere na Coleção
                colOPCarga.Add objOPCarga
                
            Next
                
        Else
    
            For Each objOPCarga In colOPCarga
            
                If objOPCarga.lItemOP = objItemOP.lNumIntDoc Then
            
                    iQtdLinhas = iQtdLinhas + 1
                    
                    If objOPCarga.lRequisicao(2) <> 0 Then iQtdLinhas = iQtdLinhas + 1
                    If objOPCarga.lRequisicao(3) <> 0 Then iQtdLinhas = iQtdLinhas + 1
                
                End If
            
            Next
    
            If iQtdLinhas < NUM_LINHAS_RELOPCARGA Then
            
                For iIndice = iQtdLinhas + 1 To NUM_LINHAS_RELOPCARGA
    
                    Set objOPCarga = New ClassOPCarga
            
                    objOPCarga.iSeq = SEQ_TIPO_LINHASCOMPLEMENTO
                    objOPCarga.iOrd = 1
                    objOPCarga.lItemOP = objItemOP.lNumIntDoc
                    objOPCarga.sCodigoOP = objOP.sCodigo
                    objOPCarga.iFilialEmpresa = objOP.iFilialEmpresa
            
                    'Insere na Coleção
                    colOPCarga.Add objOPCarga
                    
                Next
                    
            End If
        
        End If
        
    Next
    
    RelOPCarga_Insere_Linha_EmBranco2 = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Insere_Linha_EmBranco2:

    RelOPCarga_Insere_Linha_EmBranco2 = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
End Function

Private Function RelOPCarga_Le_ProducaoEntrada(ByVal objOP As ClassOrdemDeProducao, ByVal lComando As Long) As Long
'Le a Quantidade Produzida e a data de fechamento do Item da OP

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim dtDataFechamento As Date
Dim dQtdProduzido As Double
Dim dQtdTotal As Double
Dim dtDataAux As Date
Dim sUM As String
Dim dFator As Double
Dim objProduto As ClassProduto
Dim objMovEstoqueCro As ClassMovEstoqueCro
Dim sResponsavel As String

On Error GoTo Erro_RelOPCarga_Le_ProducaoEntrada
    
    For Each objItemOP In objOP.colItens
    
        Set objProduto = New ClassProduto
        Set objMovEstoqueCro = New ClassMovEstoqueCro
    
        'Descrição
        objProduto.sCodigo = objItemOP.sProduto
    
        'Le o Produto para obter a descrição
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 132137
        
        'Se o produto não está cadastrado => Erro
        If lErro <> SUCESSO Then gError 132138
        
        sUM = String(STRING_UM_SIGLA, 0)
        sResponsavel = String(STRING_MOVESTOQUE_RESPONSAVEL, 0)
            
        'Le a Quantidade Produzida e a data do movimento de estoque com o NumIntDoc do Item da OP
        lErro = Comando_Executar(lComando, "SELECT ME.Quantidade, ME.Data, ME.SiglaUM, ME.Responsavel  FROM MovimentoEstoque AS ME " & _
                                            "WHERE ME.NumIntDocOrigem = ? AND ME.TipoMov IN (?,?) AND ME.NumIntDocEst=0 AND ME.TipoNumIntDocOrigem=5 ORDER BY ME.Produto ", _
                                            dQtdProduzido, dtDataAux, sUM, sResponsavel, objItemOP.lNumIntDoc, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3)
        If lErro <> AD_SQL_SUCESSO Then gError 132133
            
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132134
           
        If lErro = AD_SQL_SEM_DADOS Then sResponsavel = ""
           
        dQtdTotal = 0
        dtDataFechamento = DATA_NULA
        
        Do While lErro <> AD_SQL_SEM_DADOS
            
             'Lê o fator de conversão para "KG"
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, "KG", dFator)
            If lErro <> SUCESSO Then gError 132139
           
            'Soma as Movimentações
            dQtdTotal = dQtdTotal + (dQtdProduzido * dFator)
            
            'Obtem a última data
            If dtDataAux > dtDataFechamento Then dtDataFechamento = dtDataAux
                     
            'Busca próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132135
        
        Loop
        
        objMovEstoqueCro.sResponsavel = sResponsavel
        
        objItemOP.dQuantidadeProd = Formata_Estoque(dQtdTotal)
        objItemOP.dtDataFimProd = dtDataFechamento
        
        Set objItemOP.objInfoUsu = objMovEstoqueCro
           
    Next
    
    RelOPCarga_Le_ProducaoEntrada = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Le_ProducaoEntrada:
    
    RelOPCarga_Le_ProducaoEntrada = gErr
        
    Select Case gErr
        
        Case 132133 To 132135
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)
            
        Case 132138
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                
        Case 132137, 132139
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

Private Function RelOPCarga_Retira_Embalagens(ByVal colOPCarga As Collection) As Long
'Le os itens que fizeram parte da Carga e Insere na Coleção

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As ClassProduto
Dim objOPCarga As ClassOPCarga

On Error GoTo Erro_RelOPCarga_Retira_Embalagens

    'Retira as embalagens da coleção
    iIndice = 0

    For Each objOPCarga In colOPCarga

        iIndice = iIndice + 1

        'Se não for a linha em branco
        If objOPCarga.iSeq <> SEQ_TIPO_EMBRANCO Then

            Set objProduto = New ClassProduto

            objProduto.sCodigo = objOPCarga.sProduto

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 132199

            If lErro <> SUCESSO Then gError 132200

            If objProduto.iNatureza = NATUREZA_PROD_EMBALAGENS Then

                colOPCarga.Remove (iIndice)
                iIndice = iIndice - 1

            End If

        End If

    Next

    RelOPCarga_Retira_Embalagens = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Retira_Embalagens:

    RelOPCarga_Retira_Embalagens = gErr
    
    Select Case gErr
    
        Case 132199
        
        Case 132200
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
End Function

Private Function RelOPCarga_Le_Lote(ByVal colOPCarga As Collection, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Le os Lotes de Um insumo de um Item de OP

Dim lErro As Long
Dim sLote As String
Dim iIndice As Integer
Dim sRastreamento As String
Dim sLocalizacao As String
Dim objFornecedor As ClassFornecedor
Dim objOPCarga As ClassOPCarga
Dim iFilialForn As Integer
Dim bbeneficiamento As Boolean
Dim iTipo As Integer

On Error GoTo Erro_RelOPCarga_Le_Lote
    
    For Each objOPCarga In colOPCarga
    
        If objOPCarga.sLote = "" Then
    
            'Aloca espaço no buffer
            sLote = String(STRING_LOTE_RASTREAMENTO, 0)
            sLocalizacao = String(STRING_RASTRO_LOCALIZACAO, 0)
            
            lErro = Comando_Executar(lComando1, "SELECT L.Lote, L.Localizacao,ME.TipoMov FROM MovimentoEstoque AS ME, RastreamentoMovto AS R, RastreamentoLote AS L " & _
                                            "WHERE Me.NumIntDocOrigem = ? AND ME.TipoMov IN (?,?) AND R.NumIntDocOrigem = ME.NumIntDoc AND ME.NumIntDocEst=0 AND (R.TipoDocOrigem = 0 OR R.TipoDocOrigem IS NULL) AND ME.TipoNumIntDocOrigem=5 AND R.NumIntDocLote = L.NumIntDoc AND ME.Produto = ? ", _
                                            sLote, sLocalizacao, iTipo, objOPCarga.lItemOP, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO, objOPCarga.sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 131647
                
            'Busca primeiro
            lErro = Comando_BuscarPrimeiro(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131648
            
            sRastreamento = ""
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                Set objFornecedor = New ClassFornecedor
                
                If iTipo = MOV_EST_REQ_PRODUCAO_BENEF3 Then
                    bbeneficiamento = True
                Else
                    bbeneficiamento = False
                End If
            
                'Obtém o fornecedor do Lote
                lErro = RelOPCarga_Le_Fornecedor_Lote(objOPCarga.sProduto, sLote, objFornecedor, lComando2, iFilialForn, bbeneficiamento)
                If lErro <> SUCESSO Then gError 140320
                
                'Se não tiver sido detalhado o fornecedor ou o Lote for do fornecedor especificado
                If objOPCarga.sLote = sLote Then
                                
                    If objOPCarga.lFornecedor = objFornecedor.lCodigo Then objOPCarga.sNomeFornecedor = objFornecedor.sNomeReduzido
                    
                    If Len(Trim(sRastreamento)) > 0 Then sRastreamento = sRastreamento & ", "
                    sRastreamento = sRastreamento & sLote
                
                    objOPCarga.sLocalizacao = sLocalizacao
                
                End If
                
                'Busca próximo
                lErro = Comando_BuscarProximo(lComando1)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131649
            
            Loop
            
            objOPCarga.sLote = sRastreamento
        
        End If
        
    Next
    
    RelOPCarga_Le_Lote = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Le_Lote:
    
    RelOPCarga_Le_Lote = gErr
        
    Select Case gErr
        
        Case 131647 To 131649
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTEMOVTO", gErr)
                
        Case 140320
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

Public Function RelOPCarga_Le_Fornecedor_Lote(ByVal sProduto As String, ByVal sLote As String, ByVal objFornecedor As ClassFornecedor, ByVal lComando As Long, iFilialForn As Integer, Optional ByVal bbeneficiamento As Boolean = False) As Long
'Le o Fornecedor associado a NF de compra

Dim lErro As Long
Dim sSQL As String
Dim lNumIntDocNF As Long
Dim lNumIntDocItemNF As Long
Dim lFornecedor As Long
Dim sFornecedor As String
Dim iFilial As Integer
Dim bAbriuComando As Boolean
Dim objRastroLote As New ClassRastreamentoLote
Dim objRastroLoteCro As New ClassRastreamentoLoteCro

On Error GoTo Erro_RelOPCarga_Le_Fornecedor_Lote

    objRastroLote.sProduto = sProduto
    objRastroLote.sCodigo = sLote
    
    'Verifica se RastreamentoLote existe
    lErro = CF("RastreamentoLote_Le_Customizado", objRastroLote)
    If lErro <> SUCESSO Then gError 185910
    
    Set objRastroLoteCro = objRastroLote.objInfoUsu

    If objRastroLoteCro.lFornecedor <> 0 And Not bbeneficiamento Then
    
        objFornecedor.lCodigo = objRastroLoteCro.lFornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 185931
    
        objFornecedor.lCodigo = objRastroLoteCro.lFornecedor
        objFornecedor.sNomeReduzido = objFornecedor.sNomeReduzido
        iFilialForn = objRastroLoteCro.iFilialForn
        
    Else

        If lComando = 0 Then
            lComando = Comando_Abrir()
            If lComando = 0 Then gError 185911
            bAbriuComando = True
        Else
            bAbriuComando = False
        End If
    
        sFornecedor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
            
        sSQL = ""
        sSQL = sSQL & "SELECT  NF.NumIntDoc, INF.NumIntDoc,  NF.Fornecedor ,NF.FilialForn, F.NomeReduzido "
        sSQL = sSQL & "FROM    MovimentoEstoque AS ME, "
        sSQL = sSQL & "    RastreamentoMovto AS R, "
        sSQL = sSQL & "    RastreamentoLote AS L, "
        sSQL = sSQL & "    TiposMovimentoEstoque AS T, "
        sSQL = sSQL & "    ItensNFiscal AS INF, "
        sSQL = sSQL & "    NFiscal As NF, "
        If Not bbeneficiamento Then
            sSQL = sSQL & "    Fornecedores AS F "
        Else
            sSQL = sSQL & "    Clientes AS F "
        End If
        sSQL = sSQL & "WHERE   ME.TipoMov = T.Codigo AND "
        sSQL = sSQL & "    R.NumIntDocOrigem = ME.NumIntDoc AND "
        sSQL = sSQL & "    R.NumIntDocLote = L.NumIntDoc AND "
        sSQL = sSQL & "    INF.NumIntDoc = ME.NumIntDocOrigem AND "
        sSQL = sSQL & "    INF.NumIntNF = NF.NumIntDoc AND "
        If Not bbeneficiamento Then
            sSQL = sSQL & "    F.Codigo = NF.Fornecedor AND "
        Else
            sSQL = sSQL & "    F.Codigo = NF.Cliente AND "
        End If
        sSQL = sSQL & "    R.TipoDocOrigem = 0  AND "
        sSQL = sSQL & "    ME.NumIntDocEst=0 AND "
        sSQL = sSQL & "    T.EntradaOuSaida = 'E' AND "
        sSQL = sSQL & "    ME.TipoNumIntDocOrigem =1 AND "
        sSQL = sSQL & "    ME.Produto = ? AND "
        sSQL = sSQL & "    L.Lote = ? "
        sSQL = sSQL & "ORDER BY ME.Data "
            
        lErro = Comando_Executar(lComando, sSQL, lNumIntDocNF, lNumIntDocItemNF, lFornecedor, iFilial, sFornecedor, sProduto, sLote)
        If lErro <> AD_SQL_SUCESSO Then gError 140366
            
        'Busca primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140367
        
        If lErro = AD_SQL_SUCESSO Then
            objFornecedor.lCodigo = lFornecedor
            objFornecedor.sNomeReduzido = sFornecedor
            iFilialForn = iFilial
        End If
        
        If bAbriuComando Then
            Call Comando_Fechar(lComando)
        End If
        
    End If
    
    RelOPCarga_Le_Fornecedor_Lote = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Le_Fornecedor_Lote:
    
    RelOPCarga_Le_Fornecedor_Lote = gErr
        
    Select Case gErr
        
        Case 185910
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 140366, 140367
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTEMOVTO", gErr)
            
        Case 185911, 185931
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    If bAbriuComando Then
        Call Comando_Fechar(lComando)
    End If
    
    Exit Function

End Function

Private Function RelOPCarga_Le_Qtd_Requisitada(ByVal colOPCarga As Collection, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Le as Requisições de consumo e calcula a diferença entre o que deveria ter sido
'requisitado e o que realmente foi.

Dim lErro As Long
Dim dQtdAux As Double
Dim dTotal As Double
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sUM As String
Dim iTipo As Integer
Dim objOPCarga As ClassOPCarga
Dim sLote As String
Dim objFornecedor As ClassFornecedor
Dim iContador As Integer
Dim lRequisicao As Long
Dim iAlmoxarifado As Integer
Dim sAlmoxarifado As String
Dim iFilialForn As Integer
Dim bbeneficiamento As Boolean

On Error GoTo Erro_RelOPCarga_Le_Qtd_Requisitada
    
    For Each objOPCarga In colOPCarga
    
        If objOPCarga.iSeq <> SEQ_TIPO_EMBRANCO Then
    
            'Descrição
            objProduto.sCodigo = objOPCarga.sProduto
            
            'Le o Produto para obter a descrição
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 132116
            
            'Se o produto não está cadastrado => Erro
            If lErro <> SUCESSO Then gError 132117
                    
            objOPCarga.sDescricao = objProduto.sDescricao
            
            sUM = String(STRING_UM_SIGLA, 0)
            sLote = String(STRING_LOTE_RASTREAMENTO, 0)
            sAlmoxarifado = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
                
            'Le as requisições de produção
            lErro = Comando_Executar(lComando1, "SELECT A.NomeReduzido, ME.Codigo, ME.Almoxarifado, ME.Quantidade, ME.SiglaUM, ME.TipoMov, L.Lote FROM Almoxarifado AS A, RastreamentoLote AS L INNER JOIN RastreamentoMovto AS R ON L.NumIntDoc = R.NumIntDocLote RIGHT OUTER JOIN MovimentoEstoque AS ME ON R.NumIntDocOrigem = ME.NumIntDoc " & _
                                                "WHERE A.Codigo = ME.Almoxarifado AND ME.NumIntDocOrigem = ? AND ME.TipoMov IN (?,?,?) AND ME.NumIntDocEst=0 AND (R.TipoDocOrigem = 0 OR R.TipoDocOrigem IS NULL) AND ME.TipoNumIntDocOrigem=5 AND ME.Produto = ? ", _
                                                sAlmoxarifado, lRequisicao, iAlmoxarifado, dQtdAux, sUM, iTipo, sLote, objOPCarga.lItemOP, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO, MOV_EST_DEV_MATERIAL_PRODUCAO, objOPCarga.sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 131643
                
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131644
                
            dTotal = 0
            iContador = 0
            Do While lErro <> AD_SQL_SEM_DADOS
                
                Set objFornecedor = New ClassFornecedor
                
                If iTipo = MOV_EST_REQ_PRODUCAO_BENEF3 Then
                    bbeneficiamento = True
                Else
                    bbeneficiamento = False
                End If
                
                lErro = RelOPCarga_Le_Fornecedor_Lote(objOPCarga.sProduto, sLote, objFornecedor, lComando2, iFilialForn, bbeneficiamento)
                If lErro <> SUCESSO Then gError 140364
                    
                If UCase(objOPCarga.sLote) = UCase(sLote) Then
                
                    iContador = iContador + 1
                    
                    'Lê o fator de conversão para "KG"
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, "KG", dFator)
                    If lErro <> SUCESSO Then gError 132125
                    
                    'Abate do que foi requisitado o que foi devolvido
                    If iTipo = MOV_EST_DEV_MATERIAL_PRODUCAO Then dQtdAux = dQtdAux * -1
                    
                    dTotal = dTotal + (dQtdAux * dFator)
                    
                    objOPCarga.sAlmoxarifado(iContador) = sAlmoxarifado
                    objOPCarga.lRequisicao(iContador) = lRequisicao
                    objOPCarga.iAlmoxarifado(iContador) = iAlmoxarifado
                    objOPCarga.dQuantReq(iContador) = dQtdAux * dFator
                    objOPCarga.iNumRequisicoes = iContador
            
                End If
                
                'Busca Próxima
                lErro = Comando_BuscarProximo(lComando1)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131645
                
            Loop
            
            objOPCarga.dAcerto = Formata_Estoque(dTotal - objOPCarga.dQuantidade)
       
        End If
    
    Next
       
    RelOPCarga_Le_Qtd_Requisitada = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Le_Qtd_Requisitada:
    
    RelOPCarga_Le_Qtd_Requisitada = gErr
        
    Select Case gErr
    
        Case 131643 To 131645
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPCARGA", gErr)
                
        Case 132116, 132125
        
        Case 132117
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

Private Function RelOPCarga_Replica_Dados(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection, ByVal lComando As Long) As Long
'Le os itens que fizeram parte da Carga e Insere na Coleção

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objOPCarga As ClassOPCarga
Dim dTotalSemAcerto As Double
Dim dTotalComAcerto As Double
Dim bTeveReq As Double
Dim objMovEstoqueCro As ClassMovEstoqueCro
Dim lNumIntOPAux As Long
Dim objProduto As ClassProduto
Dim sDescricaoAux  As String
Dim sMaquina1Aux As String
Dim sMaquina2Aux As String
Dim objMaquina As ClassMaquinas
Dim sGrupoPesagem As String

On Error GoTo Erro_RelOPCarga_Replica_Dados
        
    For Each objOPCarga In colOPCarga

        'Se não for a linha em branco
        If objOPCarga.iSeq <> 2 Then

            dTotalComAcerto = dTotalComAcerto + Formata_Estoque(objOPCarga.dQuantidade) + objOPCarga.dAcerto
            dTotalSemAcerto = dTotalSemAcerto + Formata_Estoque(objOPCarga.dQuantidade)
     
        End If

        'Obtem os demais dados do Item da OP
        For Each objItemOP In objOP.colItens

            If objItemOP.lNumIntDoc = objOPCarga.lItemOP Then

                'Obtém o Responsável
                If Not (objItemOP.objInfoUsu Is Nothing) Then
                    Set objMovEstoqueCro = objItemOP.objInfoUsu
                    objOPCarga.sResponsavel = objMovEstoqueCro.sResponsavel
                End If
                
                '##################################################
                'Obtém o grupo de pesagem
                
                sGrupoPesagem = String(STRING_GRUPO_PESAGEM, 0)
                
                lErro = Comando_Executar(lComando, "SELECT GrupoPesagem FROM ProdutoKit WHERE ProdutoRaiz = ? AND Versao = ? AND Produto = ? AND Nivel <> 0", sGrupoPesagem, objItemOP.sProduto, objItemOP.sVersao, objOPCarga.sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 190911
            
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190912
                
                objOPCarga.sGrupoPesagem = sGrupoPesagem
                '##################################################
        
        

                'Dados do item da OP
                objOPCarga.dQtdItemOP = objItemOP.dQuantidade
                objOPCarga.dQtdItemOPProduzida = objItemOP.dQuantidadeProd
                objOPCarga.dtDataIni = objItemOP.dtDataInicioProd
                objOPCarga.dtDataFim = objItemOP.dtDataFimProd
                objOPCarga.sItemOPProduto = objItemOP.sProduto

                'Se for diferenta do Item anterior
                If objOPCarga.lItemOP <> lNumIntOPAux Then

                    'Pega a descrição do Produto
                    Set objProduto = New ClassProduto

                    objProduto.sCodigo = objItemOP.sProduto

                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 132129

                    If lErro <> SUCESSO Then gError 132130

                    objOPCarga.sItemOPDescricao = objProduto.sDescricao

                    sDescricaoAux = objProduto.sDescricao
                    lNumIntOPAux = objItemOP.lNumIntDoc

                    'Pega a primeira máquina
                    If objItemOP.lNumIntEquipamento <> 0 Then

                        Set objMaquina = New ClassMaquinas

                        'Le a Maquina a partir do NumIntDoc
                        objMaquina.lNumIntDoc = objItemOP.lNumIntEquipamento

                        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
                        If lErro <> SUCESSO And lErro <> 106353 Then gError 132131

                        'Se nao encontrou => Erro
                        If lErro = 106353 Then gError 132132

                        objOPCarga.sMaquina1 = objMaquina.sNomeReduzido
                        sMaquina1Aux = objMaquina.sNomeReduzido

                    End If

                    'Pega a segunda máquina
                    If objItemOP.lNumIntEquipamento2 <> 0 Then

                        Set objMaquina = New ClassMaquinas

                        'Le a Maquina a partir do numIntDoc
                        objMaquina.lNumIntDoc = objItemOP.lNumIntEquipamento2

                        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
                        If lErro <> SUCESSO And lErro <> 106353 Then gError 132131

                        'Se nao encontrou => Erro
                        If lErro = 106353 Then gError 132132

                        objOPCarga.sMaquina2 = objMaquina.sNomeReduzido
                        sMaquina2Aux = objMaquina.sNomeReduzido

                    End If


                Else
                    'Se for o mesmo item de OP anterior pega as informações que já foram lidas

                    objOPCarga.sItemOPDescricao = sDescricaoAux
                    objOPCarga.sMaquina1 = sMaquina1Aux
                    objOPCarga.sMaquina2 = sMaquina2Aux

                End If

            End If

        Next

    Next
    
    For Each objOPCarga In colOPCarga

        'Se ainda não houve entrada de material desconsidera o acerto TotalCarga - TotalAcerto = 0
        If Abs(dTotalComAcerto) < QTDE_ESTOQUE_DELTA2 Then
            objOPCarga.dAcerto = 0
            If dTotalSemAcerto <> 0 Then
                objOPCarga.dFormFinal = 100 * ((objOPCarga.dQuantidade) / dTotalSemAcerto)
            End If
        Else
            objOPCarga.dFormFinal = 100 * (objOPCarga.dAcerto + objOPCarga.dQuantidade) / dTotalComAcerto
        End If
    
    Next
    
    RelOPCarga_Replica_Dados = SUCESSO
    
    Exit Function
    
Erro_RelOPCarga_Replica_Dados:

    RelOPCarga_Replica_Dados = gErr
    
    Select Case gErr
    
        Case 132129, 132131
        
        Case 132130
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 132132
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.iCodigo)
            
        Case 190911, 190912
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
End Function

Private Function RelOPCarga_Insere(ByVal lNumIntRel As Long, ByVal colOPCarga As Collection, ByVal lComando As Long, ByVal bEspelho As Boolean) As Long

Dim lErro As Long
Dim iSeqRel As Integer
Dim objOPCarga As ClassOPCarga

On Error GoTo Erro_RelOPCarga_Insere

    iSeqRel = 0
    For Each objOPCarga In colOPCarga

        iSeqRel = iSeqRel + 1

        With objOPCarga
            'Insere os dados na tabela que o relatório vai usar
            lErro = Comando_Executar(lComando, "INSERT INTO RelOPCarga (NumIntRel, CodigoOP, FilialEmpresa, NumIntDocItemOP, Produto, Descricao, Lote, Carga, Quantidade,QtdCarga, Acerto, FormulaFinal,DataIni, DataFim, ItemOPProduto, ItemOPDescricao, QtdItemOP, QtdItemOPProduzida, Maquina1, Maquina2, Sequencia1, Sequencia2, Responsavel, Seq, Fornecedor,NomeFornecedor,Requisicao1,Requisicao2,Requisicao3,Almoxarifado1,Almoxarifado2,Almoxarifado3,NumRequisicoes,NomeAlmoxarifado1,NomeAlmoxarifado2,NomeAlmoxarifado3,QuantReq1,QuantReq2,QuantReq3, Localizacao, GrupoPesagem, Via) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                                                        lNumIntRel, .sCodigoOP, .iFilialEmpresa, .lItemOP, .sProduto, .sDescricao, .sLote, .dCarga, .dQuantidade, .dQtdCarga, .dAcerto, .dFormFinal, .dtDataIni, .dtDataFim, .sItemOPProduto, .sItemOPDescricao, .dQtdItemOP, .dQtdItemOPProduzida, .sMaquina1, .sMaquina2, .iSeq, .iOrd, .sResponsavel, iSeqRel, .lFornecedor, .sNomeFornecedor, .lRequisicao(1), .lRequisicao(2), .lRequisicao(3), .iAlmoxarifado(1), .iAlmoxarifado(2), .iAlmoxarifado(3), .iNumRequisicoes, .sAlmoxarifado(1), .sAlmoxarifado(2), .sAlmoxarifado(3), .dQuantReq(1), .dQuantReq(2), .dQuantReq(3), .sLocalizacao, .sGrupoPesagem, 1)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 132120
        
        If bEspelho Then
        
            iSeqRel = iSeqRel + 1
    
            With objOPCarga
                'Insere os dados na tabela que o relatório vai usar
                lErro = Comando_Executar(lComando, "INSERT INTO RelOPCarga (NumIntRel, CodigoOP, FilialEmpresa, NumIntDocItemOP, Produto, Descricao, Lote, Carga, Quantidade,QtdCarga, Acerto, FormulaFinal,DataIni, DataFim, ItemOPProduto, ItemOPDescricao, QtdItemOP, QtdItemOPProduzida, Maquina1, Maquina2, Sequencia1, Sequencia2, Responsavel, Seq, Fornecedor,NomeFornecedor,Requisicao1,Requisicao2,Requisicao3,Almoxarifado1,Almoxarifado2,Almoxarifado3,NumRequisicoes,NomeAlmoxarifado1,NomeAlmoxarifado2,NomeAlmoxarifado3,QuantReq1,QuantReq2,QuantReq3, Localizacao, GrupoPesagem, Via) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                                                            lNumIntRel, .sCodigoOP, .iFilialEmpresa, .lItemOP, .sProduto, .sDescricao, .sLote, .dCarga, .dQuantidade, .dQtdCarga, .dAcerto, .dFormFinal, .dtDataIni, .dtDataFim, .sItemOPProduto, .sItemOPDescricao, .dQtdItemOP, .dQtdItemOPProduzida, .sMaquina1, .sMaquina2, .iSeq, .iOrd, .sResponsavel, iSeqRel, .lFornecedor, .sNomeFornecedor, .lRequisicao(1), .lRequisicao(2), .lRequisicao(3), .iAlmoxarifado(1), .iAlmoxarifado(2), .iAlmoxarifado(3), .iNumRequisicoes, .sAlmoxarifado(1), .sAlmoxarifado(2), .sAlmoxarifado(3), .dQuantReq(1), .dQuantReq(2), .dQuantReq(3), .sLocalizacao, .sGrupoPesagem, 2)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 132120
        
        End If

    Next

    RelOPCarga_Insere = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Insere:
    
    RelOPCarga_Insere = gErr
        
    Select Case gErr
    
        Case 132120
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELOPCARGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

'#####################################################################
'Inserido por Wagner 21/09/05
Function PedidoVenda_Grava_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim sTabela As String
Dim objPVCro As ClassPedidoDeVendaCro

On Error GoTo Erro_PedidoVenda_Grava_Customizado

    If Not (objPV.objInfoUsu Is Nothing) Then
    
        Set objPVCro = objPV.objInfoUsu

        'Abre Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 140300
        Next
       
        sTabela = "PedidosDeVenda"
       
        'Tenta ler Pedido de Venda e Pedido de Venda Baixado
        For iIndice = 1 To 2
           
            'Le o Pedido de Venda
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM " & sTabela & " WHERE NumIntDoc= ? ", 0, lNumIntDoc, objPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 140301
       
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140302
       
            If lErro = AD_SQL_SUCESSO Then Exit For
           
            sTabela = "PedidosDeVendaBaixados"
       
        Next
       
        'Não encontrou
        If lErro <> AD_SQL_SUCESSO Then gError 140303
       
        'Altera a Observação
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE " & sTabela & " SET Observacao=? ", alComando(1), objPVCro.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 140304
    
        'Fecha comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
       
    End If
    
    PedidoVenda_Grava_Customizado = SUCESSO

    Exit Function

Erro_PedidoVenda_Grava_Customizado:

    PedidoVenda_Grava_Customizado = gErr

        Select Case gErr

        Case 140300
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140301, 140302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 140303
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPV.lCodigo)
    
        Case 140304
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'#####################################################################

'###########################################################
'Inserido por Wagner 09/12/05
Public Function ItensOV_Grava_Customizado(ByVal objItensOV As ClassItemOV) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objItensOVCro As ClassItemOVCro
Dim lNumIntDoc As Long

On Error GoTo Erro_ItensOV_Grava_Customizado

    If Not (objItensOV.objInfoUsu Is Nothing) Then
    
        Set objItensOVCro = objItensOV.objInfoUsu

        'Abre Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 141364
        Next
    
        'Lê os itens do orçamento passado
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensOrcamentoVenda WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItensOV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141365
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 141366
    
        If lErro = AD_SQL_SEM_DADOS Then gError 141367
            
        'Altera a Observação
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensOrcamentoVenda SET CondPagto=?, Escolhido=?, Concorrente=? ", alComando(1), objItensOVCro.iCondPagto, objItensOVCro.iEscolhido, objItensOVCro.sConcorrente)
        If lErro <> AD_SQL_SUCESSO Then gError 141368

        'Fecha comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next

    End If

    ItensOV_Grava_Customizado = SUCESSO

    Exit Function

Erro_ItensOV_Grava_Customizado:

    ItensOV_Grava_Customizado = gErr

    Select Case gErr

        Case 141364
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131365, 131366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOV", gErr)

        Case 131367
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_VENDA_SEM_ITENS", gErr)
            
        Case 131368
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSOV", gErr, objItensOV.lCodOrcamento)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
'########################################################################

Private Function RelOPCarga_Espelho(ByVal objOP As ClassOrdemDeProducao, ByVal colOPCarga As Collection) As Long
'Le os Lotes de Um insumo de um Item de OP

Dim lErro As Long
Dim colOPCargaAux As New Collection
Dim colOPCargaBkp As New Collection
Dim objOPCarga As ClassOPCarga
Dim objOPCargaAux As ClassOPCarga
Dim iIndice As Integer
Dim bAchou As Boolean
Dim objItemOP As ClassItemOP

On Error GoTo Erro_RelOPCarga_Espelho

    For Each objOPCarga In colOPCarga
        colOPCargaBkp.Add objOPCarga
    Next
    
    iIndice = 0
    For Each objOPCarga In colOPCarga
        bAchou = False
        If Len(Trim(objOPCarga.sProduto)) > 0 Then
            For Each objOPCargaAux In colOPCargaAux
                If objOPCarga.lItemOP = objOPCargaAux.lItemOP And objOPCarga.sProduto = objOPCargaAux.sProduto Then
                    objOPCargaAux.dCarga = objOPCargaAux.dCarga + objOPCarga.dCarga
                    objOPCargaAux.dFormFinal = objOPCargaAux.dFormFinal + objOPCarga.dFormFinal
                    bAchou = True
                End If
            Next
        End If
        If Not bAchou Then
            colOPCargaAux.Add objOPCarga
        End If
        If Len(Trim(objOPCarga.sProduto)) = 0 Then
            iIndice = iIndice + 1
            If iIndice < 7 Then objOPCarga.sProduto = "ZZZZZ" & CStr(iIndice)
        End If
    Next
    
    'Para cada item da OP
    For Each objItemOP In objOP.colItens

        Set objOPCarga = New ClassOPCarga

        objOPCarga.iSeq = SEQ_TIPO_PRODACABADO
        objOPCarga.iOrd = 1
        objOPCarga.lItemOP = objItemOP.lNumIntDoc
        objOPCarga.sCodigoOP = objOP.sCodigo
        objOPCarga.iFilialEmpresa = objOP.iFilialEmpresa
        objOPCarga.sProduto = objItemOP.sProduto

        'Insere na Coleção
        colOPCargaAux.Add objOPCarga
        
    Next
    
    For iIndice = colOPCarga.Count To 1 Step -1
        colOPCarga.Remove iIndice
    Next
    
    For Each objOPCarga In colOPCargaAux
        colOPCarga.Add objOPCarga
    Next

    RelOPCarga_Espelho = SUCESSO
    
    Exit Function

Erro_RelOPCarga_Espelho:
    
    RelOPCarga_Espelho = gErr
        
    Select Case gErr
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select
    
    Exit Function

End Function

