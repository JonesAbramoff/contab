VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPCPGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'Dim giMRPIndice As Integer
'Dim giMRPSeq As Integer
'Dim gsMRPTexto As String

Public Function Maquina_Grava(ByVal objMaquinas As ClassMaquinas) As Long
'Grava a Máquina passada como parametro no BD

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Maquina_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134147

    lErro = CF("Maquinas_Grava_EmTrans", objMaquinas)
    If lErro <> SUCESSO Then gError 134148

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134149

    Maquina_Grava = SUCESSO

    Exit Function

Erro_Maquina_Grava:

    Maquina_Grava = gErr

    Select Case gErr

        Case 134147
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134148

        Case 134149
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152224)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Maquinas_Grava_EmTrans(ByVal objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim iFilial As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Maquinas_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134169
    Next

    'Busca a Máquina com mesmo Nome Reduzido
    lErro = Comando_Executar(alComando(0), "SELECT Codigo, FilialEmpresa FROM Maquinas WHERE NomeReduzido = ?", iCodigo, iFilial, objMaquinas.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 134170

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134171

    'Le a tabelaMaquinas com mesmo Codigo e Filial
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM Maquinas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 134172

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134173

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Se o NomeReduzido já existe
        If iCodigo <> 0 Then

            'Se o Código/FilialEmpresa referente ao NomeReduzido já existente forem diferentes da máquina a ser gravada, Erro.
            If objMaquinas.iCodigo <> iCodigo Or objMaquinas.iFilialEmpresa <> iFilial Then gError 134174

        End If

        objMaquinas.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 134175

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Maquinas SET NomeReduzido= ? , Descricao= ? , " & _
                "TempoMovimentacao= ? , TempoPreparacao= ? , TempoDescarga= ? , Recurso= ?, CustoHora = ?, Produto = ?, Peso = ?, Largura = ?, Comprimento = ?, Espessura = ? ", alComando(1), _
                objMaquinas.sNomeReduzido, objMaquinas.sDescricao, objMaquinas.dTempoMovimentacao, _
                objMaquinas.dTempoPreparacao, objMaquinas.dTempoDescarga, objMaquinas.iRecurso, objMaquinas.dCustoHora, objMaquinas.sProduto, objMaquinas.dPeso, objMaquinas.dLargura, objMaquinas.dComprimento, objMaquinas.dEspessura)
        If lErro <> AD_SQL_SUCESSO Then gError 134176
        
        lErro = Maquinas_ExcluiItens_EmTrans(objMaquinas)
        If lErro <> AD_SQL_SUCESSO Then gError 134177

    Else

        If iCodigo <> 0 Then gError 134178

        'Obter NumIntDoc
        'gera um NumIntDoc para a nova Máquina.
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_MAQUINA", lNumIntDoc)
        If lErro <> SUCESSO Then gError 134179

        objMaquinas.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(2), "INSERT INTO Maquinas(Codigo, FilialEmpresa, NumIntDoc, NomeReduzido, " & _
                "Descricao, TempoMovimentacao, TempoPreparacao, TempoDescarga, Recurso, CustoHora, Produto, Peso, Largura, Comprimento, Espessura) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objMaquinas.iCodigo, objMaquinas.iFilialEmpresa, objMaquinas.lNumIntDoc, objMaquinas.sNomeReduzido, _
                objMaquinas.sDescricao, objMaquinas.dTempoMovimentacao, objMaquinas.dTempoPreparacao, objMaquinas.dTempoDescarga, _
                objMaquinas.iRecurso, objMaquinas.dCustoHora, objMaquinas.sProduto, objMaquinas.dPeso, objMaquinas.dLargura, objMaquinas.dComprimento, objMaquinas.dEspessura)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134180

    End If

    lErro = Maquinas_GravaItens_EmTrans(objMaquinas)
    If lErro <> SUCESSO Then gError 134181

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Maquinas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Maquinas_Grava_EmTrans:

    Maquinas_Grava_EmTrans = gErr

    Select Case gErr

        Case 134169
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134172, 134173
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

        Case 134175
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_MAQUINAS", gErr)

        Case 134176
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_MAQUINAS", gErr)

        Case 134174, 134178
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_MAQUINA_EXISTENTE", gErr, objMaquinas.sNomeReduzido, iCodigo, iFilial)
        
        Case 134180
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MAQUINAS", gErr)

        Case 134170, 134171, 134177, 134179, 134181
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152225)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function Maquinas_GravaItens_EmTrans(ByVal objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim objMaquinasInsumos As ClassMaquinasInsumos
Dim objMaquinaOperadores As ClassMaquinaOperadores

On Error GoTo Erro_Maquinas_GravaItens_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134182
        
    'Para cada MaquinaItem na coleção
    For Each objMaquinasInsumos In objMaquinas.colProdutos
        
        objMaquinasInsumos.lNumIntDocMaq = objMaquinas.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO MaquinasInsumos (NumIntDocMaq, Produto, Quantidade, UMProduto, UMTempo) VALUES (?, ?, ?, ?, ?)", objMaquinasInsumos.lNumIntDocMaq, objMaquinasInsumos.sProduto, objMaquinasInsumos.dQuantidade, objMaquinasInsumos.sUMProduto, objMaquinasInsumos.sUMTempo)
        If lErro <> AD_SQL_SUCESSO Then gError 134183
                           
    Next
        
    'Para cada MaquinaOperadores na coleção
    For Each objMaquinaOperadores In objMaquinas.colTipoOperadores
        
        'Obter NumIntDoc
        'gera um NumIntDoc para a nova MaquinaOperadores
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_MAQUINAOPERADORES", lNumIntDoc)
        If lErro <> SUCESSO Then gError 134179
        
        objMaquinaOperadores.lNumIntDoc = lNumIntDoc
        objMaquinaOperadores.lNumIntDocMaq = objMaquinas.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO MaquinaOperadores (NumIntDoc, NumIntDocMaq, TipoMaoDeObra, Quantidade, PercentualUso) VALUES (?, ?, ?, ?, ?)", objMaquinaOperadores.lNumIntDoc, objMaquinaOperadores.lNumIntDocMaq, objMaquinaOperadores.iTipoMaoDeObra, objMaquinaOperadores.iQuantidade, objMaquinaOperadores.dPercentualUso)
        If lErro <> AD_SQL_SUCESSO Then gError 134183
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Maquinas_GravaItens_EmTrans = SUCESSO
    
    Exit Function

Erro_Maquinas_GravaItens_EmTrans:

    Maquinas_GravaItens_EmTrans = gErr

    Select Case gErr
    
        Case 134179

        Case 134182
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 134183
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152226)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Maquinas_Exclui_EmTrans(ByVal objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iFilialOP As Integer
Dim sCodOP As String
Dim lCodCT As Long
Dim iFilialCT As Integer
Dim iAux As Integer

On Error GoTo Erro_Maquinas_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134150
    Next

    'Le a tabelaMaquinas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Maquinas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 134151

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134152

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 134153

    sCodOP = String(STRING_ORDEM_DE_PRODUCAO, 0)

    'pesquisa se a Máquina já foi utilizada por alguma Ordem de Produção
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, FilialEmpresa FROM ItensOrdemProducao WHERE NumIntEquipamento = ?", sCodOP, iFilialOP, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134154

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134155

    'Se já foi utilizada em uma OP, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 134156

    'pesquisa se a Máquina já foi utilizada por algumna Ordem de Produção Baixada
    lErro = Comando_Executar(alComando(2), "SELECT Codigo, FilialEmpresa FROM ItensOrdemProducaoBaixadas WHERE NumIntEquipamento = ?", sCodOP, iFilialOP, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134157

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134158

    'se já foi utilizada em uma OP, mesmo que baixada, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 134159

    'pesquisa se a Maquina já foi utilizada por algum Centro de Trabalho
    lErro = Comando_Executar(alComando(3), "SELECT Codigo, FilialEmpresa FROM CentroDeTrabalho, CTMaquinas WHERE CentroDeTrabalho.NumIntDoc = CTMaquinas.NumIntDocCT AND CTMaquinas.NumIntDocMaq = ?", lCodCT, iFilialCT, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134273

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134274

    'se já foi utilizada em um CT, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 134275

    'pesquisa se a Maquina já foi utilizada por alguma TaxaDeProducao
    lErro = Comando_ExecutarPos(alComando(4), "SELECT 1 FROM TaxaDeProducao WHERE NumIntDocMaq = ? AND Ativo = ? ", 0, iAux, lNumIntDoc, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134276

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134277

    'se já foi utilizada em Taxas de Produção, torna-as INATIVAs.
    Do While lErro = AD_SQL_SUCESSO
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(4))
        If lErro <> SUCESSO Then gError 137178
        
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE TaxaDeProducao SET Produto= ? , NumIntDocCompet= ? , NumIntDocMaq= ? , Ativo= ? , DataDesativacao= ? ", alComando(4), _
                "", 0, 0, TAXA_INATIVA, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 137179

        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134278
    
    Loop

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 134160

    lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM Maquinas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 134161
    
    'Exclui os Produtos do Grid de Consumo, também
    lErro = Maquinas_ExcluiItens_EmTrans(objMaquinas)
    If lErro <> AD_SQL_SUCESSO Then gError 134162

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Maquinas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Maquinas_Exclui_EmTrans:

    Maquinas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 134150
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134151, 134152
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

        Case 134153
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADO", gErr, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)

        Case 134154, 134155
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", gErr)

        Case 134156, 134159
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_VINCULADA_OP", gErr, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa, sCodOP)

        Case 134157, 134158
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAOBAIXADAS", gErr)
        
        Case 134160
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_MAQUINAS", gErr)

        Case 134161
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MAQUINAS", gErr)

        Case 134162
        
        Case 134273, 134274
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)

        Case 134275
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_VINCULADA_CT", gErr, objMaquinas.iCodigo, lCodCT, objMaquinas.iFilialEmpresa)

        Case 134276, 134277, 134278
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)
        
        Case 137178
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TAXADEPRODUCAO", gErr)

        Case 137179
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TAXADEPRODUCAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152227)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Maquinas_ExcluiItens_EmTrans(ByVal objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDocMaq As Long

On Error GoTo Erro_Maquinas_ExcluiItens_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134163
    Next

    'Exclui MaquinasInsumos
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDocMaq FROM MaquinasInsumos WHERE NumIntDocMaq = ?", 0, lNumIntDocMaq, objMaquinas.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134164

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134165

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 134166
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM MaquinasInsumos", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 134167
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134168

    Loop

    'Exclui MaquinaOperadores
    lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDocMaq FROM MaquinaOperadores WHERE NumIntDocMaq = ?", 0, lNumIntDocMaq, objMaquinas.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134164

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134165

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(3))
        If lErro <> SUCESSO Then gError 134166
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM MaquinaOperadores", alComando(3))
        If lErro <> AD_SQL_SUCESSO Then gError 134167
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134168

    Loop

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Maquinas_ExcluiItens_EmTrans = SUCESSO

    Exit Function

Erro_Maquinas_ExcluiItens_EmTrans:

    Maquinas_ExcluiItens_EmTrans = gErr

    Select Case gErr

        Case 134163
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134164, 134165, 134168
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

        Case 134166
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_MAQUINAS", gErr)

        Case 134167
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152228)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Maquina_Exclui(ByVal objMaquinas As ClassMaquinas) As Long
'Exclui a Máquina a partir do Codigo e FilialEmpresa

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim bExcluiMaquina As Boolean
Dim lNumIntDoc As Long
Dim iAux As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lTransacao As Long

On Error GoTo Erro_Maquina_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137172
    Next

    bExcluiMaquina = True
    
    'Le a tabelaMaquinas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Maquinas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 137173

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137174

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 137175

    'pesquisa se a Maquina já foi utilizada por alguma TaxaDeProducao
    lErro = Comando_Executar(alComando(1), "SELECT 1 FROM TaxaDeProducao WHERE NumIntDocMaq = ? AND Ativo = ? ", iAux, lNumIntDoc, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 137176

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137177

    'se já foi utilizada em uma Taxa de Produção, pergunta se mesmo assim deseja excluir ...
    If lErro = AD_SQL_SUCESSO Then
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_MAQUINA_COM_TAXASVINC", objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    
        If vbMsgRes = vbNo Then
            bExcluiMaquina = False
        End If
    
    End If
    
    'Gera um Erro para ser tratado na rotina chamadora
    If Not bExcluiMaquina Then gError 137181

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134144

    lErro = CF("Maquinas_Exclui_EmTrans", objMaquinas)
    If lErro <> SUCESSO Then gError 134145

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134146

    Maquina_Exclui = SUCESSO

    Exit Function

Erro_Maquina_Exclui:

    Maquina_Exclui = gErr

    Select Case gErr

        Case 137172
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137173, 137174
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

        Case 137175
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADO", gErr, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)

        Case 137176, 137177
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)
            
        Case 137181
            'não vai excluir -> portanto não Limpa a Tela
        
        Case 134144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134145
            'erro tratado na rotina chamada

        Case 134146
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152229)

    End Select

    If lTransacao <> 0 Then
    
        'Desfaz Transação
        Call Transacao_Rollback

    End If

    Exit Function

End Function

Function Maquina_Automatico(iCodigo As Integer) As Long
'Obtem o próximo número de Máquina

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_Maquina_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "PCPConfig", "NUM_PROX_MAQUINA", "Maquinas", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 103130

    If lNumAuto > NUM_MAX_MAQUINAS Then gError 137065
    
    iCodigo = CInt(lNumAuto)

    Maquina_Automatico = SUCESSO

    Exit Function

Erro_Maquina_Automatico:

    Maquina_Automatico = gErr

   Select Case gErr
   
        Case 137065
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, lNumAuto)

        Case 103130
            'Erro já tratado
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152230)

    End Select

    Exit Function

End Function

''*** Daqui p/ baixo rotinas referente a Previsão de Vendas x Previsão de Consumo ******
''???? cadastrar esta função no dic
Public Function PrevVenda_ReqMat_Calcula(ByVal sProd_I As String, ByVal sProd_Fim As String, ByVal iMes As Integer, ByVal iAno As Integer, ByVal sPrevisao As String, ByVal sVersao As String) As Long
'Faz a filtragem dos Parâmetros que serão utilizados no calculo da Previsão de Consumo.

Dim lErro As Long
Dim alComando(0) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objPrevVMensal As ClassPrevVendaMensal
Dim tPrevVMensal As typePrevVendaMensal
Dim colMateriais As New Collection
Dim sSQL As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_PrevVenda_ReqMat_Calcula

    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 103210

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103187
    Next

    With tPrevVMensal

        .sCodigo = String(STRING_PREVVENDA_CODIGO, 0)
        .sProduto = String(STRING_PRODUTO, 0)

        sSQL = "SELECT Produto, SUM(Quantidade" & iMes & ") FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ? AND Ano = ? " & _
        "Group BY Produto"

        'Seleciona os dados de PrevVendaMensal de acordo com o Código, FilialEmpresa e Ano recebidos
        lErro = Comando_Executar(alComando(0), sSQL, _
         .sProduto, .dQuantidade1, sPrevisao, giFilialEmpresa, iAno)
        If lErro <> AD_SQL_SUCESSO Then gError 103193

    End With

    'Tenta buscar o primeiro elemento.
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103189

    Do While lErro = AD_SQL_SUCESSO
        
        Set objPrevVMensal = New ClassPrevVendaMensal

        objPrevVMensal.sProduto = tPrevVMensal.sProduto
        objPrevVMensal.dQuantidade1 = tPrevVMensal.dQuantidade1

        'Prepara a gravação dos elementos lidos e dispara a rotina de gravação em PrevVendaPrevConsumo.
        lErro = PrevVendaPrevConsumo_Grava(objPrevVMensal, iMes, iAno, sVersao, colMateriais)
        If lErro <> SUCESSO Then gError 103204
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103192
            
    Loop
    
    lErro = PrevVendaPrevConsumo_Grava_BD(colMateriais, iAno, iMes, sPrevisao)
    If lErro <> SUCESSO Then gError 103223

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 103217

    PrevVenda_ReqMat_Calcula = SUCESSO

    Exit Function

Erro_PrevVenda_ReqMat_Calcula:

    PrevVenda_ReqMat_Calcula = gErr

    Select Case gErr

        Case 103210
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103217
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 103189, 103192, 103193
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDAMENSAL", gErr, sPrevisao)

        Case 103187
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103204, 103223

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152231)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

'???? cadastrar esta função no dic
Public Function PrevVendaPrevConsumo_Grava(ByVal objPrevVMensal As ClassPrevVendaMensal, ByVal iMes As Integer, ByVal iAno As Integer, sVersao As String, ByVal colMateriais As Collection) As Long
'Prepara a gravação dos elementos lidos e dispara a rotina de gravação em PrevVendaPrevConsumo.

Dim lErro As Long
Dim objPrevConsumo As New ClassPrevVendaPrevConsumo
Dim objProdutoKit As New ClassProdutoKit
Dim objKit As ClassKit
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double

On Error GoTo Erro_PrevVendaPrevConsumo_Grava

    objProduto.sCodigo = objPrevVMensal.sProduto

    'Lê o produto para descobrir as unidades de medidas associadas
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 103206

    'Descobre o fator de conversao da UM de Venda p/UM de estoque do produto raiz
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFatorConversao)
    If lErro <> SUCESSO Then gError 103207

    'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
    objPrevConsumo.dQuantPrevInsumo = objPrevVMensal.dQuantidade1 * dFatorConversao
    objPrevConsumo.dQuantPrevVenda = objPrevConsumo.dQuantPrevInsumo
        
    'Carrega o obj com a Versão do produto
    objPrevConsumo.sVersao = sVersao
    objPrevConsumo.sProduto = objPrevVMensal.sProduto
    objPrevConsumo.sUMEstoque = objProduto.sSiglaUMEstoque

    lErro = Calcula_Insumos_PrevConsumo(colMateriais, objPrevConsumo, 0)
    If lErro <> SUCESSO Then gError 103209

    PrevVendaPrevConsumo_Grava = SUCESSO

    Exit Function

Erro_PrevVendaPrevConsumo_Grava:

    PrevVendaPrevConsumo_Grava = gErr

    Select Case gErr

        Case 103196, 103205, 103206, 103207, 103208, 103209

        Case 103195
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152232)

    End Select

    Exit Function

End Function

Public Function Calcula_Insumos_PrevConsumo(ByVal colMateriais As Collection, ByVal objPrevConsumo As ClassPrevVendaPrevConsumo, ByVal iNiveis As Integer) As Long
'Calcula de forma recursiva todos os componentes necessários para a fabricação de
'um Kit até último nivel.

Dim lErro As Long
Dim objPrevConsumo1 As ClassPrevVendaPrevConsumo
Dim objKit As New ClassKit
Dim objKitAux As New ClassKit
Dim bEncontrou As Boolean
Dim objProdutoKit As New ClassProdutoKit
Dim objProdutoKitAux As New ClassProdutoKit
Dim objProduto As New ClassProduto
Dim dQuantidadeFilho As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim iNiveisAux As Integer
Dim dFatorConversao As Double
Dim dQuantidade As Double
Dim dFatorConversaoProdutoRaiz As Double

On Error GoTo Erro_Calcula_Insumos_PrevConsumo

    
    'Se a quantidade de chegou a 1 Kg ou já desceu na recursividade 25 instâncias
    If (objPrevConsumo.dQuantPrevInsumo <= 1 And iNiveis > 0) Or iNiveis >= 20 Then
        Calcula_Insumos_PrevConsumo = SUCESSO
        Exit Function

    End If

    'verifica se o produto está na coleção de produtos
    For Each objPrevConsumo1 In colMateriais

        'se o Produto existe
        If objPrevConsumo1.sProduto = objPrevConsumo.sProduto Then

            objPrevConsumo1.dQuantPrevInsumo = objPrevConsumo1.dQuantPrevInsumo + objPrevConsumo.dQuantPrevInsumo
            objPrevConsumo1.dQuantPrevVenda = objPrevConsumo1.dQuantPrevVenda + objPrevConsumo.dQuantPrevVenda
            bEncontrou = True
            Exit For

        End If

    Next
    
    'Se ainda não existe na coleção de produtos
    If Not bEncontrou Then

        Set objPrevConsumo1 = objPrevConsumo
        
        objEstoqueProduto.sProduto = objPrevConsumo.sProduto
        
        'Le a qtd do produto em estoque
        lErro = CF("EstoqueProduto_ObterQTD_Filial", objEstoqueProduto, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 103277 Then gError 103220

        'se o produto realmente for da filial em questao
        If lErro = SUCESSO Then
        
            objPrevConsumo1.dQuantEstoque = objEstoqueProduto.dQuantDispNossa
            
            'Adiociona este produto em uma coleção global
            colMateriais.Add objPrevConsumo1

        End If

    End If

    

    objKit.sVersao = objPrevConsumo.sVersao
    objKit.sProdutoRaiz = objPrevConsumo.sProduto

    'tem que ler o proprio cara e os seus filhos
    lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
    If lErro <> SUCESSO And lErro <> 103203 Then gError 103221
    
    'Se o produto for um Kit, então leva a recursividade em diante
    If lErro = SUCESSO Then
    
        'percorre os elementos componentes do Kit Lido
        For Each objProdutoKit In objKit.colComponentes
        
            objProduto.sCodigo = objProdutoKit.sProduto
    
            'Lê o produto para descobrir as unidades de medidas associadas
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 103218
    
            'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
            If lErro <> SUCESSO Then gError 103219
        
            If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                
                dQuantidade = (objPrevConsumo.dQuantPrevInsumo / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                
            Else
                
                Set objPrevConsumo = New ClassPrevVendaPrevConsumo
        
                objPrevConsumo.sProduto = objProdutoKit.sProdutoRaiz
                objPrevConsumo.sVersao = objProdutoKit.sVersao
        
                'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
                objPrevConsumo.dQuantPrevInsumo = dQuantidade * (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda) * dFatorConversao)
                objPrevConsumo.sProduto = objProdutoKit.sProduto
                objPrevConsumo.sUMEstoque = objProduto.sSiglaUMEstoque
                objPrevConsumo.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                  
                'Calcula de maneira recursiva a quantidade de componentes necessários.
                lErro = Calcula_Insumos_PrevConsumo(colMateriais, objPrevConsumo, iNiveis + 1)
                If lErro <> SUCESSO Then gError 103222
        
            End If
        
        Next
    
    End If

    Calcula_Insumos_PrevConsumo = SUCESSO

    Exit Function

Erro_Calcula_Insumos_PrevConsumo:

    Calcula_Insumos_PrevConsumo = gErr

    Select Case gErr

        Case 103218 To 103222, 103224, 103225

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152233)

    End Select


    Exit Function

End Function

Public Function PrevVendaPrevConsumo_Grava_BD(ByVal colMateriais As Collection, ByVal iAno As Integer, ByVal iMes As Integer, sCodigo As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objPrevVendaPrevConsumo As New ClassPrevVendaPrevConsumo
Dim dQuantidade As Double
Dim dQuantPrevVenda As Double

On Error GoTo Erro_PrevVendaPrevConsumo_Grava_BD

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103211
    Next

    For Each objPrevVendaPrevConsumo In colMateriais
            
        'Verifica se o produto já existe
        lErro = Comando_ExecutarPos(alComando(0), "SELECT QuantPrevVenda FROM PrevVendaPrevConsumo WHERE Produto = ? AND FilialEmpresa = ? AND CodigoPrevVenda = ? AND Ano = ? AND Mes = ?", _
        0, dQuantPrevVenda, objPrevVendaPrevConsumo.sProduto, giFilialEmpresa, sCodigo, iAno, iMes)
        If lErro <> AD_SQL_SUCESSO Then gError 103212

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103213

        If lErro = AD_SQL_SUCESSO Then

            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PrevVendaPrevConsumo SET UMEstoque = ?, QuantPrevVenda = ?, QuantPrevInsumo = ?, QuantEstoque = ?, Data = ?", _
            alComando(0), objPrevVendaPrevConsumo.sUMEstoque, objPrevVendaPrevConsumo.dQuantPrevVenda, objPrevVendaPrevConsumo.dQuantPrevInsumo, objPrevVendaPrevConsumo.dQuantEstoque, gdtDataHoje)
            If lErro <> AD_SQL_SUCESSO Then gError 103214

        Else

          With objPrevVendaPrevConsumo

                lErro = Comando_Executar(alComando(0), "INSERT INTO PrevVendaPrevConsumo(FilialEmpresa, CodigoPrevVenda, Produto, Ano, Mes, UMEstoque, QuantPrevInsumo, QuantEstoque, QuantPrevVenda, Data) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                giFilialEmpresa, sCodigo, .sProduto, iAno, iMes, .sUMEstoque, .dQuantPrevInsumo, .dQuantEstoque, .dQuantPrevVenda, gdtDataHoje)
                If lErro <> AD_SQL_SUCESSO Then gError 103216

            End With

        End If

    Next

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    PrevVendaPrevConsumo_Grava_BD = SUCESSO

    Exit Function

Erro_PrevVendaPrevConsumo_Grava_BD:

    PrevVendaPrevConsumo_Grava_BD = gErr

    Select Case gErr

        Case 103211
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103212, 103213
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVENDAPREVCONSUMO", gErr)

        Case 103214
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_PREVENDAPREVCONSUMO", gErr)

        Case 103216
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PREVENDAPREVCONSUMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152234)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Public Function RelArvoreProduto_Grava(ByVal sProdutoDe As String, ByVal sProdutoAte As String, ByVal bSoPadrao As Boolean, lNumIntRel As Long, iMaiorNivel As Integer) As Long
'Grava a Árvore do Produto passado por objKit para relatório - Jorge Specian 04/03/2005

Dim lErro As Long
Dim lTransacao As Long
Dim sProdutoRaiz As String
Dim sVersao As String
Dim iSituacao As Integer
Dim objKit As ClassKit
Dim lComando As Long
Dim sComando_SQL As String
Dim iSeqPai As Integer
Dim iSeq As Integer
Dim iIdentacao As Integer
Dim sStringIdentacao As String
Dim dQtdePai As Double

On Error GoTo Erro_RelArvoreProduto_Grava

    GL_objMDIForm.MousePointer = vbHourglass

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134184

    'Iniciar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134185
    
    sProdutoRaiz = String(STRING_PRODUTO, 0)
    sVersao = String(STRING_KIT_VERSAO, 0)
        
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_ARVOREPRODUTO", lNumIntRel)
    If lErro <> SUCESSO Then gError 134186
    
    'Montar o SELECT ...
    Call RelArvoreProduto_Grava1(sProdutoDe, sProdutoAte, bSoPadrao, sComando_SQL, iSituacao)
    
    'Executar o SELECT por partes ...
    lErro = RelArvoreProduto_Grava2(lComando, sComando_SQL, sProdutoRaiz, sVersao, iSituacao, sProdutoDe, sProdutoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 134187

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134188

    If lErro = AD_SQL_SEM_DADOS Then gError 134189
        
    'Fazer enquanto houver Kits
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objKit = New ClassKit
        
        objKit.sProdutoRaiz = sProdutoRaiz
        objKit.sVersao = sVersao
        
        If objKit.sVersao = "" Then
        
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 134190
    
            If lErro = 106304 Then
                
                gError 134191
            
            End If

        End If
    
        lErro = RelArvoreProduto_Grava_ProdutoRaiz(objKit, dQtdePai, iSeqPai, lNumIntRel)
        If lErro <> SUCESSO Then gError 134192
        
        lErro = CF("RelArvoreProduto_Grava_EmTrans", objKit, dQtdePai, iSeqPai, iSeq, iIdentacao, sStringIdentacao, lNumIntRel, iMaiorNivel)
        If lErro <> SUCESSO Then gError 134193
        
        'Buscar o Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134194
        
        'zera Sequencia do Produto
        iSeq = 0
        
    Loop
    
    'Encerrar comando
    Call Comando_Fechar(lComando)

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134195
    
    GL_objMDIForm.MousePointer = vbDefault
    
    RelArvoreProduto_Grava = SUCESSO

    Exit Function

Erro_RelArvoreProduto_Grava:

    RelArvoreProduto_Grava = gErr

    Select Case gErr

        Case 134184
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134186, 134192, 134193
            'erro tratado nas rotinas chamadas
            
        Case 134187, 134188, 134190, 134191, 134194
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT_DE_ATE", gErr, sProdutoDe, sProdutoAte)

        Case 134189
            'erro tratado na rotina chamadora - não há dados para o relatório

        Case 134195
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152235)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Public Function RelArvoreProduto_Grava1(ByVal sProdutoDe As String, ByVal sProdutoAte As String, ByVal bSoPadrao As Boolean, sComando_SQL As String, iSituacao As Integer) As Long
'auxiliar a RelArvoreProduto_Grava
'monta o SELECT para obtencao dos itens do relatorio dinamicamente.

Dim lErro As Long

On Error GoTo Erro_RelArvoreProduto_Grava1

    sComando_SQL = "SELECT ProdutoRaiz, Versao FROM Kit "
    
    If bSoPadrao Then
        sComando_SQL = sComando_SQL & "WHERE Situacao = ? "
        iSituacao = KIT_SITUACAO_PADRAO
    Else
        sComando_SQL = sComando_SQL & "WHERE Situacao <> ? "
        iSituacao = KIT_SITUACAO_INATIVO
    End If
    
    If (sProdutoDe <> "") Then
    
        If (sProdutoAte <> "") Then
            sComando_SQL = sComando_SQL & "AND ProdutoRaiz >= ? AND ProdutoRaiz <= ? "
        Else
            sComando_SQL = sComando_SQL & "AND ProdutoRaiz >= ? "
        End If
        
    Else
    
        If (sProdutoAte <> "") Then
            sComando_SQL = sComando_SQL & "AND ProdutoRaiz <= ? "
        End If
        
    End If
    
    sComando_SQL = sComando_SQL & "ORDER BY ProdutoRaiz, Situacao, Versao"
    
    RelArvoreProduto_Grava1 = SUCESSO
    
    Exit Function

Erro_RelArvoreProduto_Grava1:

    RelArvoreProduto_Grava1 = Err

    Select Case Err
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152236)

    End Select

    Exit Function

End Function
    
Public Function RelArvoreProduto_Grava2(ByVal lComando As Long, ByVal sComando_SQL As String, vProdutoRaiz As Variant, vVersao As Variant, vSituacao As Variant, vProdutoDe As Variant, vProdutoAte As Variant)
'isola a preparacao da parte fixa do SELECT, e complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT p/obtencao dos itens

Dim ret As Integer, lErro As Long

On Error GoTo Erro_RelArvoreProduto_Grava2

    ret = Comando_PrepararInt(lComando, sComando_SQL)
    If (ret <> AD_SQL_SUCESSO) Then gError 134257

    ret = Comando_BindVarInt(lComando, vProdutoRaiz)
    If (ret <> AD_SQL_SUCESSO) Then gError 134258
    
    ret = Comando_BindVarInt(lComando, vVersao)
    If (ret <> AD_SQL_SUCESSO) Then gError 134259
    
    ret = Comando_BindVarInt(lComando, vSituacao)
    If (ret <> AD_SQL_SUCESSO) Then gError 134260

    'se produto inicial preenchido
    If (vProdutoDe <> "") Then
        ret = Comando_BindVarInt(lComando, vProdutoDe)
        If (ret <> AD_SQL_SUCESSO) Then gError 134261
    End If

    'se produto final preenchido
    If (vProdutoAte <> "") Then
        ret = Comando_BindVarInt(lComando, vProdutoAte)
        If (ret <> AD_SQL_SUCESSO) Then gError 134262
    End If

    ret = Comando_ExecutarInt(lComando)
    If (ret <> AD_SQL_SUCESSO) Then gError 134263

    RelArvoreProduto_Grava2 = SUCESSO

    Exit Function

Erro_RelArvoreProduto_Grava2:

    RelArvoreProduto_Grava2 = gErr

    Select Case gErr
        
        Case 134257 To 134263
            'erro a ser tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152237)

    End Select

    Exit Function

End Function

Public Function RelArvoreProduto_Grava_EmTrans(ByVal objKit As ClassKit, ByVal dQtdePai As Double, ByVal iSeqPai As Integer, iSeq As Integer, ByVal iIdentacao As Integer, ByVal sStringIdentacao As String, ByVal lNumIntRel As Long, iMaiorNivel As Integer) As Long
'Grava a Árvore do Produto passado por objKit para relatório - Jorge Specian 04/03/2005

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim tProdutoKit As typeProdutoKit
Dim sComando_SQL As String
Dim iClasseUM As Integer
Dim sDescricao As String
Dim iCompras As Integer
Dim sOrigem As String
Dim sNomeReduzido As String
Dim sSiglaUMEstoque As String
Dim sSiglaUM_Kit As String
Dim dQuantidadeKit As Double
Dim objKitIntermediario As ClassKit
Dim bEh_Outro_Kit As Boolean
Dim dFatorConversao As Double

On Error GoTo Erro_RelArvoreProduto_Grava_EmTrans

    'Se a quantidade já chegou a 0.0001 ou já desceu na recursividade n instâncias
    If (dQtdePai <= 0.0001 And iIdentacao > 0) Or iIdentacao + 1 >= 20 Then
    
        RelArvoreProduto_Grava_EmTrans = SUCESSO
        Exit Function

    End If

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134201
    Next
    
    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
    tProdutoKit.sVersaoKitComp = String(STRING_KIT_VERSAO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA_UM_BASE, 0)
    
    sComando_SQL = "SELECT ProdutoKit.PercentualPerda, ProdutoKit.ProdutoRaiz, ProdutoKit.Versao, ProdutoKit.Nivel, ProdutoKit.Seq, ProdutoKit.Produto, ProdutoKit.SeqPai, ProdutoKit.Quantidade, ProdutoKit.UnidadeMed, ProdutoKit.Composicao, ProdutoKit.VersaoKitComp, Produtos.Descricao, Produtos.NomeReduzido, Produtos.ClasseUM, Produtos.SiglaUMEstoque, Produtos.Compras FROM ProdutoKit, Produtos WHERE ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Versao = ? AND Produtos.Codigo = ProdutoKit.Produto ORDER BY ProdutoKit.Nivel, ProdutoKit.Seq"
    lErro = Comando_Executar(alComando(0), sComando_SQL, tProdutoKit.dPercentualPerda, tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, tProdutoKit.sVersaoKitComp, sDescricao, sNomeReduzido, iClasseUM, sSiglaUMEstoque, iCompras, objKit.sProdutoRaiz, objKit.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 134202

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134203

    If lErro = AD_SQL_SEM_DADOS Then gError 134210
        
    'Faço enquanto houver Componentes
    Do While lErro <> AD_SQL_SEM_DADOS
        
        bEh_Outro_Kit = False
        
        If iCompras = PRODUTO_COMPRAVEL Then
            sOrigem = INSUMO_COMPRADO
        Else
            sOrigem = INSUMO_PRODUZIDO
        End If
        
        'Se não for o ProtudoRaiz
        If tProdutoKit.iNivel > KIT_NIVEL_RAIZ Then
            
            'Descobre o fator de conversao da U.M. recebida p/U.M. de estoque do produto
            lErro = CF("UM_Conversao_Trans", iClasseUM, tProdutoKit.sUnidadeMed, sSiglaUMEstoque, dFatorConversao)
            If lErro <> SUCESSO Then gError 134204
            
            'Converter a quantidade pela Unidade do Kit x Produto
            dQuantidadeKit = dQtdePai * (tProdutoKit.dQuantidade / (1 - tProdutoKit.dPercentualPerda)) * dFatorConversao
            
            'e ajusto a U.M. conforme convertido
            sSiglaUM_Kit = sSiglaUMEstoque
            
            Set objKitIntermediario = New ClassKit
            
            objKitIntermediario.sProdutoRaiz = tProdutoKit.sProduto
            objKitIntermediario.sVersao = tProdutoKit.sVersaoKitComp
            
            'Leio o Kit para ver se é Intermediário
            If objKitIntermediario.sVersao = "" Then
        
                lErro = CF("Kit_Le_Padrao", objKitIntermediario)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 134205
        
                If lErro = SUCESSO Then
                   
                   bEh_Outro_Kit = True
                   
                End If
        
            Else
        
                lErro = CF("Kit_Le", objKitIntermediario)
                If lErro <> SUCESSO And lErro <> 21826 Then gError 134206
        
                If lErro = SUCESSO Then
                
                   bEh_Outro_Kit = True
                
                End If
        
            End If
                    
        Else
        
            'acrescento a perda se houver
            dQuantidadeKit = dQtdePai * (tProdutoKit.dQuantidade / (1 - tProdutoKit.dPercentualPerda))
            
            'e mantenho a mesma U.M. por ser nivel = zero
            sSiglaUM_Kit = tProdutoKit.sUnidadeMed
            
        End If
        
        'Se encontrou um Kit de Produto Intermédiário -> recursividade
        If bEh_Outro_Kit Then
                
            'Leio KitProduto novamente para Produto Intermediário - Chamo esta própria Rotina
            lErro = RelArvoreProduto_Grava_EmTrans(objKitIntermediario, dQuantidadeKit, iSeqPai, iSeq, iIdentacao, sStringIdentacao, lNumIntRel, iMaiorNivel)
            If lErro <> SUCESSO And lErro <> 134210 Then gError 134207
        
        Else
            
            'Incrementa a Sequencia
            iSeq = iSeq + 1
                                    
            'Grava na Tabela RelArvoreProdutoItens
            sComando_SQL = "INSERT INTO RelArvoreProdutoItens(NumIntRel, SeqPai, Seq, Produto, Descricao, Origem, NomeReduzido, Quantidade, UMedida, Identacao, StringIdentacao) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
            lErro = Comando_Executar(alComando(1), sComando_SQL, lNumIntRel, iSeqPai, iSeq, tProdutoKit.sProduto, sDescricao, sOrigem, sNomeReduzido, dQuantidadeKit, sSiglaUM_Kit, iIdentacao, sStringIdentacao)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134208
           
            If tProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                
                'Incrementa a Identação
                iIdentacao = iIdentacao + 1
                
                'Incrementa a String de Identação (...)
                sStringIdentacao = String((iIdentacao * NUMERO_DE_CARACTERES_POR_IDENTACAO), CARACTER_DA_STRING_DE_IDENTACAO)
                
                'E, conta os Níveis = Identação
                If iIdentacao > iMaiorNivel Then
                   
                   iMaiorNivel = iIdentacao
                
                End If
            
            End If
           
        End If
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134209
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelArvoreProduto_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RelArvoreProduto_Grava_EmTrans:

    RelArvoreProduto_Grava_EmTrans = gErr

    Select Case gErr

        Case 134201
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134202, 134203, 134209
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 134210
            'Não há ProdutoKit - Vamos prosseguir
            
        Case 134204, 134205, 134206, 134207
            'Erro tratado nas rotinas chamadas

        Case 134208
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_RELATORIO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152238)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelArvoreProduto_Grava_ProdutoRaiz(ByVal objKit As ClassKit, dQtdePai As Double, iSeqPai As Integer, ByVal lNumIntRel As Long) As Long
'Grava a Árvore do Produto passado por objKit para relatório - Jorge Specian 04/03/2005

Dim lErro As Long
Dim lComando As Long
Dim sComando_SQL As String
Dim objProdutoKit As ClassProdutoKit
Dim objProduto As ClassProduto
Dim sSiglaUM_Kit As String
Dim dFatorConversao As Double

On Error GoTo Erro_RelArvoreProduto_Grava_ProdutoRaiz

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134196
    
    Set objProdutoKit = New ClassProdutoKit
    
    objProdutoKit.sProdutoRaiz = objKit.sProdutoRaiz
    objProdutoKit.sVersao = objKit.sVersao
    
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO Then gError 134197
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objProdutoKit.sProdutoRaiz

    'Lê o produto para descobrir as unidades de medidas associadas
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 134198
    
    'Descobre o fator de conversao da U.M. recebida p/U.M. de estoque do produto
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
    If lErro <> SUCESSO Then gError 134199
    
    'Converter a quantidade pela Unidade do Kit x Produto
    dQtdePai = (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda)) * dFatorConversao
    
    'e ajusto a U.M. conforme convertido
    sSiglaUM_Kit = objProduto.sSiglaUMEstoque
       
    'Incrementa a Sequencia do Pai
    iSeqPai = iSeqPai + 1

    'Grava na Tabela RelArvoreProduto
    sComando_SQL = "INSERT INTO RelArvoreProduto(NumIntRel, Seq, ProdutoRaiz, Versao, Quantidade, UMedida) VALUES (?,?,?,?,?,?)"
    lErro = Comando_Executar(lComando, sComando_SQL, lNumIntRel, iSeqPai, objKit.sProdutoRaiz, objKit.sVersao, dQtdePai, sSiglaUM_Kit)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134200
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    RelArvoreProduto_Grava_ProdutoRaiz = SUCESSO

    Exit Function

Erro_RelArvoreProduto_Grava_ProdutoRaiz:

    RelArvoreProduto_Grava_ProdutoRaiz = gErr

    Select Case gErr

        Case 134196
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134197
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 134198, 134199
            'Erro tratado nas rotinas chamadas

        Case 134200
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_RELATORIO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152239)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CentrodeTrabalho_Exclui(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CentrodeTrabalho_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134420

    lErro = CF("CentrodeTrabalho_Exclui_EmTrans", objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134421

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134422

    CentrodeTrabalho_Exclui = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Exclui:

    CentrodeTrabalho_Exclui = gErr

    Select Case gErr

        Case 134420
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134421

        Case 134422
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152240)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CentrodeTrabalho_Exclui_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iSeq As Integer
Dim sProdutoRaiz As String
Dim sVersao As String
Dim sNomeReduzido As String

On Error GoTo Erro_CentrodeTrabalho_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134423
    Next
    
    sProdutoRaiz = String(STRING_PRODUTO, 0)
    sVersao = String(STRING_KIT_VERSAO, 0)
    sNomeReduzido = String(STRING_NOME_REDUZIDO, 0)

    'Le a tabelaCentrodeTrabalho
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CentrodeTrabalho WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 134424

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134425

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 134426

    'pesquisa se o Centro de Trabalho já foi utilizado por alguma Operação
    lErro = Comando_Executar(alComando(1), "SELECT RoteirosDeFabricacao.ProdutoRaiz, RoteirosDeFabricacao.Versao, Operacoes.Seq FROM RoteirosDeFabricacao, Operacoes WHERE RoteirosDeFabricacao.NumIntDoc = Operacoes.NumIntDocRotFabr AND Operacoes.NumIntDocCT = ?", sProdutoRaiz, sVersao, iSeq, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134427

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134428

    'se já foi utilizado em uma Operação, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 134429

    'pesquisa se o Centro de Trabalho é Padrão de alguma Competência
    lErro = Comando_Executar(alComando(3), "SELECT NomeReduzido FROM Competencias WHERE NumIntDocCT= ? ", sNomeReduzido, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137068

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137069

    'se é Padrão de alguma Competencia, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 137070

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 134430

    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM CentrodeTrabalho", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 134431

    lErro = CentrodeTrabalho_ExcluiMaquinas_EmTrans(objCentrodeTrabalho)
    If lErro <> AD_SQL_SUCESSO Then gError 134432

    lErro = CentrodeTrabalho_ExcluiCompetencias_EmTrans(objCentrodeTrabalho)
    If lErro <> AD_SQL_SUCESSO Then gError 134433

    lErro = CentrodeTrabalho_ExcluiTurnos_EmTrans(objCentrodeTrabalho)
    If lErro <> AD_SQL_SUCESSO Then gError 137195

    lErro = CentrodeTrabalho_ExcluiOperadores_EmTrans(objCentrodeTrabalho)
    If lErro <> AD_SQL_SUCESSO Then gError 139106

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CentrodeTrabalho_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Exclui_EmTrans:

    CentrodeTrabalho_Exclui_EmTrans = gErr

    Select Case gErr

        Case 134423
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134424, 134425
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)

        Case 134426
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO", gErr, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)

        Case 134427, 134428
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)
        
        Case 134429
            Call Rotina_Erro(vbOKOnly, "ERRO_CT_VINCULADO_ROTEIROS", gErr, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa, sProdutoRaiz, sVersao)
        
        Case 137068, 137069
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)
        
        Case 137070
            Call Rotina_Erro(vbOKOnly, "ERRO_CT_PADRAO_COMPETENCIA", gErr, sNomeReduzido)
                
        Case 134430
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CENTRODETRABALHO", gErr)

        Case 134431
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CENTRODETRABALHO", gErr)
            
        Case 134432, 134433, 137195, 139106
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152241)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CentrodeTrabalho_Grava(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CentrodeTrabalho_Grava

    'Verifica e emite aviso que este será o CT Padrão das Competências escolhidas sem CT Padrão
    lErro = CF("Competencias_VerificaCTPadrao", objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134923

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134398

    lErro = CF("CentrodeTrabalho_Grava_EmTrans", objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134399

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134400

    CentrodeTrabalho_Grava = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Grava:

    CentrodeTrabalho_Grava = gErr

    Select Case gErr

        Case 134398
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134399, 134923
            'erros tratados nas rotinas chamadas

        Case 134400
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152242)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CentrodeTrabalho_Grava_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lCodigo As Integer
Dim iFilial As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CentrodeTrabalho_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134401
    Next

    'Busca o Centro de Trabalho com mesmo Nome Reduzido
    lErro = Comando_Executar(alComando(0), "SELECT Codigo, FilialEmpresa FROM CentrodeTrabalho WHERE NomeReduzido = ?", lCodigo, iFilial, objCentrodeTrabalho.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 134402

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134403

    'Le a tabelaCentrodeTrabalho pelo Código
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM CentrodeTrabalho WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 134404

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134405

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Se o NomeReduzido já existe
        If lCodigo <> 0 Then

            'Se o Código/FilialEmpresa referente ao NomeReduzido já existente forem diferentes da máquina a ser gravada, Erro.
            If objCentrodeTrabalho.lCodigo <> lCodigo Or objCentrodeTrabalho.iFilialEmpresa <> iFilial Then gError 134406

        End If

        objCentrodeTrabalho.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 134407

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CentrodeTrabalho SET NomeReduzido= ? , Descricao= ? , CargaMin= ? , CargaMax= ? , Turnos= ? , HorasTurno= ? , DomingoUtil= ? , SegundaUtil= ? , TercaUtil= ? , QuartaUtil= ? , QuintaUtil= ? , SextaUtil= ? , SabadoUtil= ? ", alComando(1), _
                objCentrodeTrabalho.sNomeReduzido, objCentrodeTrabalho.sDescricao, objCentrodeTrabalho.dCargaMin, objCentrodeTrabalho.dCargaMax, objCentrodeTrabalho.iTurnos, objCentrodeTrabalho.dHorasTurno, objCentrodeTrabalho.iDiaisUteis(DOMINGO), objCentrodeTrabalho.iDiaisUteis(SEGUNDA), _
                objCentrodeTrabalho.iDiaisUteis(TERCA), objCentrodeTrabalho.iDiaisUteis(QUARTA), objCentrodeTrabalho.iDiaisUteis(QUINTA), objCentrodeTrabalho.iDiaisUteis(SEXTA), objCentrodeTrabalho.iDiaisUteis(SABADO))
        If lErro <> AD_SQL_SUCESSO Then gError 134408

        lErro = CentrodeTrabalho_ExcluiMaquinas_EmTrans(objCentrodeTrabalho)
        If lErro <> AD_SQL_SUCESSO Then gError 134409

        lErro = CentrodeTrabalho_ExcluiCompetencias_EmTrans(objCentrodeTrabalho, CTCOMPETENCIA_NAO_VALIDA_EXCLUSAO_TODAS)
        If lErro <> AD_SQL_SUCESSO Then gError 134410

        lErro = CentrodeTrabalho_ExcluiTurnos_EmTrans(objCentrodeTrabalho)
        If lErro <> AD_SQL_SUCESSO Then gError 137193
        
        lErro = CentrodeTrabalho_ExcluiOperadores_EmTrans(objCentrodeTrabalho)
        If lErro <> AD_SQL_SUCESSO Then gError 139087

    Else

        If lCodigo <> 0 Then gError 134411

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_CENTRODETRABALHO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 134412

        objCentrodeTrabalho.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(2), "INSERT INTO CentrodeTrabalho( NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, " & _
                "CargaMin, CargaMax, Turnos, HorasTurno, DomingoUtil, SegundaUtil, TercaUtil, QuartaUtil, QuintaUtil, SextaUtil, SabadoUtil) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objCentrodeTrabalho.lNumIntDoc, objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa, objCentrodeTrabalho.sNomeReduzido, objCentrodeTrabalho.sDescricao, _
                objCentrodeTrabalho.dCargaMin, objCentrodeTrabalho.dCargaMax, objCentrodeTrabalho.iTurnos, objCentrodeTrabalho.dHorasTurno, objCentrodeTrabalho.iDiaisUteis(DOMINGO), objCentrodeTrabalho.iDiaisUteis(SEGUNDA), _
                objCentrodeTrabalho.iDiaisUteis(TERCA), objCentrodeTrabalho.iDiaisUteis(QUARTA), objCentrodeTrabalho.iDiaisUteis(QUINTA), objCentrodeTrabalho.iDiaisUteis(SEXTA), objCentrodeTrabalho.iDiaisUteis(SABADO))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134413

    End If

    lErro = CentrodeTrabalho_GravaMaquinas_EmTrans(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134414

    lErro = CentrodeTrabalho_GravaCompetencias_EmTrans(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 134415
    
    lErro = CentrodeTrabalho_GravaTurnos_EmTrans(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 137194

    lErro = CentrodeTrabalho_GravaOperadores_EmTrans(objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 139088

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CentrodeTrabalho_Grava_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Grava_EmTrans:

    CentrodeTrabalho_Grava_EmTrans = gErr

    Select Case gErr

        Case 134401
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134402 To 134405
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)

        Case 134406, 134411
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_CENTRODETRABALHO_EXISTENTE", gErr, objCentrodeTrabalho.sNomeReduzido, lCodigo, iFilial)

        Case 134407
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CENTRODETRABALHO", gErr)

        Case 134408
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CENTRODETRABALHO", gErr)

        Case 134413
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CENTRODETRABALHO", gErr)
            
        Case 134409, 134410, 134412, 134414, 134415, 137193, 137194, 139087, 139088
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152243)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Competencias_Exclui(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim bExcluiCompetencia As Boolean
Dim lNumIntDoc As Long
Dim iAux As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lTransacao As Long

On Error GoTo Erro_Competencias_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137164
    Next

    bExcluiCompetencia = True
    
    'Le a tabelaCompetencias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Competencias WHERE Codigo= ? ", _
                0, lNumIntDoc, objCompetencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 137165

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137166

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 137167
    
    'pesquisa se a Competencia já foi utilizada por alguma TaxaDeProducao
    lErro = Comando_Executar(alComando(1), "SELECT 1 FROM TaxaDeProducao WHERE NumIntDocCompet = ? AND Ativo = ? ", iAux, lNumIntDoc, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 137168

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137169

    'se já foi utilizada em uma Taxa de Produção, pergunta se mesmo assim deseja excluir ...
    If lErro = AD_SQL_SUCESSO Then
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_COMP_COM_TAXA", objCompetencias.lCodigo)
    
        If vbMsgRes = vbNo Then
            bExcluiCompetencia = False
        End If
    
    End If

    'Gera um Erro para ser tratado na rotina chamadora
    If Not bExcluiCompetencia Then gError 137180
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134311

    lErro = CF("Competencias_Exclui_EmTrans", objCompetencias)
    If lErro <> SUCESSO Then gError 134312

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134313
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Competencias_Exclui = SUCESSO

    Exit Function

Erro_Competencias_Exclui:

    Competencias_Exclui = gErr

    Select Case gErr

        Case 137164
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137165, 137166
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 137167
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIAS_NAO_CADASTRADO", gErr, objCompetencias.lCodigo)

        Case 137168, 137169
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)
        
        Case 137180
            'não vai excluir -> portanto não exclui da ListView e nem Limpa a Tela

        Case 134311
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134312
            'erro tratado na rotina chamada

        Case 134313
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152244)

    End Select

    If lTransacao <> 0 Then
        
        'Desfaz Transação
        Call Transacao_Rollback
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Competencias_Exclui_EmTrans(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lCodCT As Long
Dim sCodOP As String
Dim iFilial As Integer
Dim iSeq As Integer
Dim sProduto As String
Dim sVersao As String
Dim iAux As Integer
Dim objCompetenciaPadrao As New ClassCompetencias

On Error GoTo Erro_Competencias_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134314
    Next

    sProduto = String(STRING_PRODUTO, 0)
    sVersao = String(STRING_KIT_VERSAO, 0)
    sCodOP = String(STRING_OPCODIGO, 0)

    'Le a tabelaCompetencias
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Competencias WHERE Codigo= ? ", _
                0, lNumIntDoc, objCompetencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 134315

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134316

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 134317

    'pesquisa se a Competencia já foi utilizada por algum Centro de Trabalho
    lErro = Comando_Executar(alComando(1), "SELECT CentroDeTrabalho.Codigo, CentroDeTrabalho.FilialEmpresa FROM CentroDeTrabalho, CTCompetencias WHERE CentroDeTrabalho.NumIntDoc = CTCompetencias.NumIntDocCT AND CTCompetencias.NumIntDocCompet = ?", lCodCT, iFilial, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134318

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134319

    'se já foi utilizada em um CT, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 134320

    'pesquisa se a Competencia já foi utilizada por alguma Operação
    lErro = Comando_Executar(alComando(2), "SELECT RoteirosDeFabricacao.ProdutoRaiz, RoteirosDeFabricacao.Versao, Operacoes.Seq FROM RoteirosDeFabricacao, Operacoes WHERE RoteirosDeFabricacao.NumIntDoc = Operacoes.NumIntDocRotFabr AND Operacoes.NumIntDocCompet = ?", sProduto, sVersao, iSeq, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134321

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134322

    'se já foi utilizada em uma Operação, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 134323

    'pesquisa se a Competencia já foi utilizada por alguma TaxaDeProducao
    lErro = Comando_ExecutarPos(alComando(3), "SELECT 1 FROM TaxaDeProducao WHERE NumIntDocCompet = ? AND Ativo = ? ", 0, iAux, lNumIntDoc, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134324

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134325

    'se já foi utilizada em Taxas de Produção, torna-as INATIVAS
    Do While lErro = AD_SQL_SUCESSO
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(3))
        If lErro <> SUCESSO Then gError 137170
        
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE TaxaDeProducao SET Produto= ? , NumIntDocCompet= ? , NumIntDocMaq= ? , Ativo= ? , DataDesativacao= ? ", alComando(3), _
                "", 0, 0, TAXA_INATIVA, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 137171

        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134326
    
    Loop

    'pesquisa se a Competencia já foi utilizada por alguma Operação em OPs
    lErro = Comando_Executar(alComando(5), "SELECT OrdensDeProducao.FilialEmpresa, OrdensDeProducao.Codigo, ItensOrdemProducao.Produto FROM OrdensDeProducao INNER JOIN ItensOrdemProducao ON OrdensDeProducao.FilialEmpresa = ItensOrdemProducao.FilialEmpresa AND OrdensDeProducao.Codigo = ItensOrdemProducao.Codigo INNER JOIN OrdemProducaoOperacoes ON ItensOrdemProducao.NumIntDoc = OrdemProducaoOperacoes.NumIntDocItemOP WHERE (OrdemProducaoOperacoes.NumIntDocCompet = ?)", iFilial, sCodOP, sProduto, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137073

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137074

    'se já foi utilizada em uma Operação em OPs, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 137075

    'pesquisa se a Competencia já foi utilizada por alguma Operação em OPs Baixadas
    lErro = Comando_Executar(alComando(6), "SELECT OrdensDeProducaoBaixadas.FilialEmpresa, OrdensDeProducaoBaixadas.Codigo, ItensOrdemProducaoBaixadas.Produto FROM OrdensDeProducaoBaixadas INNER JOIN ItensOrdemProducaoBaixadas ON OrdensDeProducaoBaixadas.FilialEmpresa = ItensOrdemProducaoBaixadas.FilialEmpresa AND OrdensDeProducaoBaixadas.Codigo = ItensOrdemProducaoBaixadas.Codigo INNER JOIN OrdemProducaoOperacoes ON ItensOrdemProducaoBaixadas.NumIntDoc = OrdemProducaoOperacoes.NumIntDocItemOP WHERE (OrdemProducaoOperacoes.NumIntDocCompet = ?)", iFilial, sCodOP, sProduto, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137076

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137077

    'se já foi utilizada em uma Operação em OPs Baixadas, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 137078

    'Le a Competencia Padrao
    lErro = CF("Competencias_Le_Padrao", objCompetenciaPadrao)
    If lErro <> SUCESSO And lErro <> 137844 Then gError 139264

    'Se encontrou a Padrão...
    If lErro = SUCESSO Then
    
        If objCompetencias.lCodigo = objCompetenciaPadrao.lCodigo Then gError 139265

    End If

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 134327

    lErro = Comando_ExecutarPos(alComando(7), "DELETE FROM Competencias", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 134328

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Competencias_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Competencias_Exclui_EmTrans:

    Competencias_Exclui_EmTrans = gErr

    Select Case gErr

        Case 134314
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134315, 134316
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 134317
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIAS_NAO_CADASTRADO", gErr, objCompetencias.lCodigo)

        Case 134318, 134319
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)
        
        Case 134320
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_VINCULADA_CT", gErr, objCompetencias.lCodigo, lCodCT, iFilial)
        
        Case 134321, 134322
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)
        
        Case 134323
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_VINCULADA_ROTEIROS", gErr, objCompetencias.lCodigo, sProduto, sVersao, iSeq)
        
        Case 134324, 134325, 134326
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 137170
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TAXADEPRODUCAO", gErr)

        Case 137171
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TAXADEPRODUCAO", gErr)
        
        Case 137073, 137074
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", gErr)

        Case 137075
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_VINCULADA_OP", gErr, objCompetencias.lCodigo, sCodOP, iFilial, sProduto)
        
        Case 137076, 137077
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOBAIXADA", gErr)

        Case 137078
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_VINCULADA_OPBAIXADA", gErr, objCompetencias.lCodigo, sCodOP, iFilial, sProduto)
        
        Case 134327
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_COMPETENCIAS", gErr)

        Case 134328
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMPETENCIAS", gErr)
            
        Case 139264
            'erro tratado na rotina chamada
            
        Case 139265
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMPETENCIA_PADRAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152245)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Competencias_Grava(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Competencias_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134296

    lErro = CF("Competencias_Grava_EmTrans", objCompetencias)
    If lErro <> SUCESSO Then gError 134297

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134298

    Competencias_Grava = SUCESSO

    Exit Function

Erro_Competencias_Grava:

    Competencias_Grava = gErr

    Select Case gErr

        Case 134296
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134297

        Case 134298
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152246)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Competencias_Grava_EmTrans(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim objCompetenciaPadrao As New ClassCompetencias
Dim bMudaPadrao As Boolean
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lCodigo As Integer

On Error GoTo Erro_Competencias_Grava_EmTrans

    'Le a Competencia Padrao
    lErro = CF("Competencias_Le_Padrao", objCompetenciaPadrao)
    If lErro <> SUCESSO And lErro <> 137844 Then gError 137845

    'Se não encontrou nenhuma Padrão...
    If lErro <> SUCESSO Then
    
        'Esta será a Padrão
        objCompetencias.iPadrao = MARCADO
    
    Else
    
        If objCompetencias.lCodigo = objCompetenciaPadrao.lCodigo Then
        
            'Esta continuará ser a Padrão
            objCompetencias.iPadrao = MARCADO
        
        Else
        
            'Se estou gravando como Padrão...
            If objCompetencias.iPadrao = MARCADO Then
            
                'Altero a que já era Padrão
                objCompetenciaPadrao.iPadrao = DESMARCADO
                
                bMudaPadrao = True
            
            End If
        
        End If
    
    End If

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134299
    Next

    'Le a tabelaCompetencias pelo NomeReduzido
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Competencias WHERE NomeReduzido= ? ", _
                0, lCodigo, objCompetencias.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 134300

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134301

    'Le a tabelaCompetencias pelo Código
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM Competencias WHERE Codigo= ? ", _
                0, lNumIntDoc, objCompetencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 134302

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134303

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Se o NomeReduzido já existe
        If lCodigo <> 0 Then

            'Se o Código referente ao NomeReduzido já existente forem diferentes da competencia a ser gravada, Erro.
            If objCompetencias.lCodigo <> lCodigo Then gError 134304

        End If

        objCompetencias.lNumIntDoc = lNumIntDoc
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 134305

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Competencias SET NomeReduzido= ?, Descricao= ?, NumIntDocCT= ?, Padrao= ?", alComando(1), objCompetencias.sNomeReduzido, objCompetencias.sDescricao, objCompetencias.lNumIntDocCT, objCompetencias.iPadrao)
        If lErro <> AD_SQL_SUCESSO Then gError 134306

    Else

        If lCodigo <> 0 Then gError 134307

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_COMPETENCIAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 134308

        objCompetencias.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(2), "INSERT INTO Competencias( NumIntDoc, Codigo, NomeReduzido, Descricao, NumIntDocCT, Padrao ) VALUES (?,?,?,?,?,?)", _
                objCompetencias.lNumIntDoc, objCompetencias.lCodigo, objCompetencias.sNomeReduzido, objCompetencias.sDescricao, objCompetencias.lNumIntDocCT, objCompetencias.iPadrao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134309

    End If
    
    If bMudaPadrao = True Then
    
        'Le a tabelaCompetencias pelo Código
        lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM Competencias WHERE Codigo= ? ", _
                    0, lNumIntDoc, objCompetenciaPadrao.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 139260
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 139261
    
        'Se não existir => Outro usuário pode já ter apagado
        If lErro <> AD_SQL_SUCESSO Then gError 137846
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(3))
        If lErro <> SUCESSO Then gError 139262

        lErro = Comando_ExecutarPos(alComando(4), "UPDATE Competencias SET Padrao= ?", alComando(3), objCompetenciaPadrao.iPadrao)
        If lErro <> AD_SQL_SUCESSO Then gError 139263
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Competencias_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Competencias_Grava_EmTrans:

    Competencias_Grava_EmTrans = gErr

    Select Case gErr

        Case 134299
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134300 To 134303, 139260, 139261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)
            
        Case 134304, 134307
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMERED_COMPETENCIA_EXISTENTE", gErr, objCompetencias.sNomeReduzido, lCodigo)

        Case 134305, 139262
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_COMPETENCIAS", gErr)

        Case 134306, 139263
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_COMPETENCIAS", gErr)

        Case 134309
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMPETENCIAS", gErr)
            
        Case 134308
            'erro tratado na rotina chamadora
            
        Case 137845, 137846
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152247)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Competencias_Grava_CTPadrao(ByVal objCTCompetencias As ClassCTCompetencias) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_Competencias_Grava_CTPadrao

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134927
    Next

    'Le a tabelaCompetencias pelo NumIntDoc
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM Competencias WHERE NumIntDoc= ? ", _
                0, iAux, objCTCompetencias.lNumIntDocCompet)
    If lErro <> AD_SQL_SUCESSO Then gError 134928

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134929

    'Se nao existir -> Erro
    If lErro <> AD_SQL_SUCESSO Then gError 134930
        
    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 134931

    lErro = Comando_ExecutarPos(alComando(1), "UPDATE Competencias SET NumIntDocCT= ?", alComando(0), objCTCompetencias.lNumIntDocCT)
    If lErro <> AD_SQL_SUCESSO Then gError 134932

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Competencias_Grava_CTPadrao = SUCESSO

    Exit Function

Erro_Competencias_Grava_CTPadrao:

    Competencias_Grava_CTPadrao = gErr

    Select Case gErr

        Case 134927
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134928, 134929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)
            
        Case 134930 'Sem dados - tratado na rotina chamadora
            
        Case 134931
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_COMPETENCIAS", gErr)

        Case 134932
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_COMPETENCIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152248)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TaxaDeProducao_Exclui(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TaxaDeProducao_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134531

    lErro = CF("TaxaDeProducao_Exclui_EmTrans", objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134532

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134533

    TaxaDeProducao_Exclui = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Exclui:

    TaxaDeProducao_Exclui = gErr

    Select Case gErr

        Case 134531
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134532
            'erro tratado na rotina chamada

        Case 134533
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152249)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TaxaDeProducao_Exclui_EmTrans(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iAux As Integer

On Error GoTo Erro_TaxaDeProducao_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134534
    Next

    'Le a tabela TaxaDeProducao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocMaq= ? AND NumIntDocCompet= ? AND Ativo = ? ", _
                0, lNumIntDoc, objTaxaDeProducao.sProduto, objTaxaDeProducao.lNumIntDocMaq, objTaxaDeProducao.lNumIntDocCompet, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134535

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134536

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 134537

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 134538
    
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE TaxaDeProducao SET Ativo= ? , DataDesativacao= ? ", alComando(0), _
            TAXA_INATIVA, gdtDataAtual)
    If lErro <> AD_SQL_SUCESSO Then gError 134539

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TaxaDeProducao_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Exclui_EmTrans:

    TaxaDeProducao_Exclui_EmTrans = gErr

    Select Case gErr

        Case 134534
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134535, 134536
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134537
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXADEPRODUCAO_NAO_CADASTRADA", gErr)

        Case 134538
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TAXADEPRODUCAO", gErr)

        Case 134539
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TAXADEPRODUCAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152250)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Public Function TaxaDeProducao_Grava(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TaxaDeProducao_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134521

    lErro = CF("TaxaDeProducao_Grava_EmTrans", objTaxaDeProducao)
    If lErro <> SUCESSO Then gError 134522

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134523

    TaxaDeProducao_Grava = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Grava:

    TaxaDeProducao_Grava = gErr

    Select Case gErr

        Case 134521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134522
            'erro tratado na rotina chamada

        Case 134523
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152251)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TaxaDeProducao_Grava_EmTrans(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TaxaDeProducao_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134524
    Next

    'Le a tabelaTaxaDeProducao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocMaq= ? AND NumIntDocCompet= ? AND Ativo= ? ", _
                0, iAux, objTaxaDeProducao.sProduto, objTaxaDeProducao.lNumIntDocMaq, objTaxaDeProducao.lNumIntDocCompet, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134525

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134526

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 134527

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TaxaDeProducao SET Ativo= ? , DataDesativacao= ? ", alComando(0), _
                TAXA_INATIVA, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 134528

    End If

    'Obter NumIntDoc
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_TAXADEPRODUCAO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 134529

    objTaxaDeProducao.lNumIntDoc = lNumIntDoc

    lErro = Comando_Executar(alComando(1), "INSERT INTO TaxaDeProducao(NumIntDoc, Produto, NumIntDocMaq, NumIntDocCompet, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, UMTempo, Data, Ativo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objTaxaDeProducao.lNumIntDoc, objTaxaDeProducao.sProduto, objTaxaDeProducao.lNumIntDocMaq, objTaxaDeProducao.lNumIntDocCompet, objTaxaDeProducao.dLoteMax, objTaxaDeProducao.dLoteMin, objTaxaDeProducao.dLotePadrao, objTaxaDeProducao.dTempoPreparacao, objTaxaDeProducao.dTempoMovimentacao, _
            objTaxaDeProducao.dTempoDescarga, objTaxaDeProducao.iTipo, objTaxaDeProducao.dQuantidade, objTaxaDeProducao.sUMProduto, objTaxaDeProducao.dTempoOperacao, objTaxaDeProducao.sUMTempo, objTaxaDeProducao.dtData, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134530

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TaxaDeProducao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Grava_EmTrans:

    TaxaDeProducao_Grava_EmTrans = gErr

    Select Case gErr

        Case 134524
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134525, 134526
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134527
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TAXADEPRODUCAO", gErr)

        Case 134528
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TAXADEPRODUCAO", gErr)

        Case 134529
            'erro tratado na rotina chamada

        Case 134530
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TAXADEPRODUCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152252)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Public Function RoteirosDeFabricacao_Grava(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim objKit As ClassKit
Dim vbMsgRes As VbMsgBoxResult
Dim lTransacao As Long

On Error GoTo Erro_RoteirosDeFabricacao_Grava
    
    Set objKit = New ClassKit
    
    objKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
    objKit.sVersao = objRoteirosDeFabricacao.sVersao
    
    'Verifica se existe o Kit
    lErro = CF("Kit_Le", objKit)
    If lErro <> SUCESSO And lErro <> 21826 Then gError 134826
    
    'Caso tenha encontrado ...
    If lErro = SUCESSO Then
    
        'Pergunta ao usuário se confirma a alteração do Kit
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ALTERACAO_DO_KIT", Trim(objRoteirosDeFabricacao.sProdutoRaiz), Trim(objRoteirosDeFabricacao.sVersao))
    
        If vbMsgRes = vbNo Then gError 134827
    
    Else
    
        'Se não tem vai criar um novo
        objKit.dtData = objRoteirosDeFabricacao.dtDataUltModificacao
        objKit.iSituacao = KIT_SITUACAO_PADRAO
        objKit.iVersaoFormPreco = objRoteirosDeFabricacao.iVersaoFormPreco
        
    End If
    
    'Cria ou recria os componentes do Kit
    lErro = RoteirosDeFabricacao_CriaProdutoKit(objRoteirosDeFabricacao, objKit)
    If lErro <> SUCESSO Then gError 134828
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134829
    
    lErro = CF("Kit_Grava_EmTrans", objKit)
    If lErro <> SUCESSO Then gError 134830

    lErro = CF("RoteirosDeFabricacao_Grava_EmTrans", objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134831

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134832

    RoteirosDeFabricacao_Grava = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Grava:

    RoteirosDeFabricacao_Grava = gErr

    Select Case gErr

        Case 134829
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134826, 134830, 134831
            'erros tratados nas rotinas chamadas
        
        Case 134827
            'usuário não aceitou alterar o Kit

        Case 134832
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152253)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RoteirosDeFabricacao_Grava_EmTrans(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteirosDeFabricacao_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134833
    Next

    'Le a tabela RoteirosDeFabricacao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteirosDeFabricacao WHERE ProdutoRaiz= ? AND Versao= ? ", _
                0, lNumIntDoc, objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 134834

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134835

    'Se existir
    If lErro = AD_SQL_SUCESSO Then
    
        objRoteirosDeFabricacao.lNumIntDoc = lNumIntDoc
        
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 134837
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RoteirosDeFabricacao SET ProdutoRaiz = ? , Versao = ? , Descricao = ? , DataCriacao = ? , DataUltModificacao = ? , Quantidade = ? , UM = ? , Autor = ?, NumMaxMaqPorOper = ? ", alComando(0), _
                objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao, objRoteirosDeFabricacao.sDescricao, _
                objRoteirosDeFabricacao.dtDataCriacao, objRoteirosDeFabricacao.dtDataUltModificacao, objRoteirosDeFabricacao.dQuantidade, objRoteirosDeFabricacao.sUM, objRoteirosDeFabricacao.sAutor, objRoteirosDeFabricacao.iNumMaxMaqPorOper)
        If lErro <> AD_SQL_SUCESSO Then gError 134838
    
        lErro = RoteirosDeFabricacao_ExcluiOperacoes_EmTrans(objRoteirosDeFabricacao)
        If lErro <> SUCESSO Then gError 134840
        
    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_ROTEIROSDEFABRICACAO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 134841
    
        objRoteirosDeFabricacao.lNumIntDoc = lNumIntDoc
    
        'E insere um novo
        lErro = Comando_Executar(alComando(2), "INSERT INTO RoteirosDeFabricacao( NumIntDoc, ProdutoRaiz, Versao, Descricao, " & _
                "DataCriacao, DataUltModificacao, Quantidade, UM, Autor, NumMaxMaqPorOper) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objRoteirosDeFabricacao.lNumIntDoc, objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao, objRoteirosDeFabricacao.sDescricao, _
                objRoteirosDeFabricacao.dtDataCriacao, objRoteirosDeFabricacao.dtDataUltModificacao, objRoteirosDeFabricacao.dQuantidade, objRoteirosDeFabricacao.sUM, objRoteirosDeFabricacao.sAutor, objRoteirosDeFabricacao.iNumMaxMaqPorOper)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134842
    
    End If
    
    lErro = RoteirosDeFabricacao_GravaOperacoes_EmTrans(objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134843

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteirosDeFabricacao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Grava_EmTrans:

    RoteirosDeFabricacao_Grava_EmTrans = gErr

    Select Case gErr

        Case 134833
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134834, 134835
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)
            
        Case 134840, 134841, 134843
            'erros tratados nas rotinas chamadas

        Case 134837
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSDEFABRICACAO", gErr)

        Case 134838
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ROTEIROSDEFABRICACAO", gErr)

        Case 134839
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSDEFABRICACAO", gErr)

        Case 134842
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ROTEIROSDEFABRICACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152254)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteirosDeFabricacao_GravaOperacoes_EmTrans(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objOperacoes As New ClassOperacoes
Dim objOperacoesAApagar As ClassOperacoes
Dim lNumIntDoc As Long
Dim bApagarOperacao As Boolean

On Error GoTo Erro_RoteirosDeFabricacao_GravaOperacoes_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134844
    Next
        
    'Para cada Operacoes na coleção
    For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes

        'Le a tabela Operacoes
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Operacoes WHERE NumIntDocRotFabr= ? AND Seq= ? ", _
                    0, lNumIntDoc, objRoteirosDeFabricacao.lNumIntDoc, objOperacoes.iSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 134845
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134846
    
        'Se existir => UPDATE, senão => INSERT
        If lErro = AD_SQL_SUCESSO Then
        
            objOperacoes.lNumIntDoc = lNumIntDoc
            objOperacoes.lNumIntDocRotFabr = objRoteirosDeFabricacao.lNumIntDoc
        
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then gError 134847
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE Operacoes SET Seq= ? , NumIntDocCompet= ? , NumIntDocCT= ? , Observacao= ? , " & _
                    "IgnoraTaxaProducao= ? , SeqPai= ? , SeqArvore= ? , Nivel= ? , PosicaoArvore= ?, NumMaxMaqPorOper= ?, NumRepeticoes = ? ", alComando(0), _
                    objOperacoes.iSeq, objOperacoes.lNumIntDocCompet, objOperacoes.lNumIntDocCT, objOperacoes.sObservacao, _
                    objOperacoes.iIgnoraTaxaProducao, objOperacoes.iSeqPai, objOperacoes.iSeqArvore, objOperacoes.iNivel, objOperacoes.iPosicaoArvore, objOperacoes.iNumMaxMaqPorOper, objOperacoes.iNumRepeticoes)
            If lErro <> AD_SQL_SUCESSO Then gError 134848
            
            lErro = RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans(objOperacoes)
            If lErro <> SUCESSO Then gError 134849
               
            lErro = RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans(objOperacoes)
            If lErro <> SUCESSO Then gError 134850
            
        Else
    
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_OPERACOES", lNumIntDoc)
            If lErro <> SUCESSO Then gError 134851
    
            objOperacoes.lNumIntDoc = lNumIntDoc
            objOperacoes.lNumIntDocRotFabr = objRoteirosDeFabricacao.lNumIntDoc
    
            lErro = Comando_Executar(alComando(1), "INSERT INTO Operacoes( NumIntDoc, NumIntDocRotFabr, Seq, NumIntDocCompet, NumIntDocCT, Observacao, " & _
                    "IgnoraTaxaProducao, SeqPai, SeqArvore, Nivel, PosicaoArvore, NumMaxMaqPorOper, NumRepeticoes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objOperacoes.lNumIntDoc, objOperacoes.lNumIntDocRotFabr, objOperacoes.iSeq, objOperacoes.lNumIntDocCompet, objOperacoes.lNumIntDocCT, objOperacoes.sObservacao, _
                    objOperacoes.iIgnoraTaxaProducao, objOperacoes.iSeqPai, objOperacoes.iSeqArvore, objOperacoes.iNivel, objOperacoes.iPosicaoArvore, objOperacoes.iNumMaxMaqPorOper, objOperacoes.iNumRepeticoes)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134852
    
        End If
        
        If objOperacoes.iIgnoraTaxaProducao = MARCADO Then
        
            lErro = RoteirosDeFabricacao_GravaOperacoesTempo_EmTrans(objOperacoes)
            If lErro <> SUCESSO Then gError 134853
        
        End If
        
        lErro = RoteirosDeFabricacao_GravaOperacaoInsumos_EmTrans(objOperacoes)
        If lErro <> SUCESSO Then gError 134854
        
    Next
    
    'Le a tabela Operacoes novamente para verificar se haverá deleções
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM Operacoes WHERE NumIntDocRotFabr= ? ", _
                0, lNumIntDoc, objRoteirosDeFabricacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134855
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134856

    'Faça enquanto existir
    Do While lErro = AD_SQL_SUCESSO
    
        bApagarOperacao = True
        
        'Para cada Operacoes na coleção
        For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes
        
            If objOperacoes.lNumIntDoc = lNumIntDoc Then
            
                 bApagarOperacao = False
                 Exit For
            
            End If
        
        Next
        
        If bApagarOperacao Then
            
            lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM Operacoes", alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 134857
            
            Set objOperacoesAApagar = New ClassOperacoes
            
            objOperacoesAApagar.lNumIntDoc = lNumIntDoc
            
            lErro = RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans(objOperacoesAApagar)
            If lErro <> SUCESSO Then gError 134858
        
            lErro = RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans(objOperacoesAApagar)
            If lErro <> SUCESSO Then gError 134859
                
        End If

        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134860

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RoteirosDeFabricacao_GravaOperacoes_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_GravaOperacoes_EmTrans:

    RoteirosDeFabricacao_GravaOperacoes_EmTrans = gErr

    Select Case gErr

        Case 134844
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134845, 134846, 134855, 134856, 134860
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOES", gErr)

        Case 134847
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OPERACOES", gErr)

        Case 134848
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_OPERACOES", gErr)
            
        Case 134849, 134850, 134851, 134853, 134854, 134858, 134859
            'erros tratados nas rotinas chamadas

        Case 134852
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OPERACOES", gErr)
            
        Case 134857
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPERACOES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152255)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RoteirosDeFabricacao_GravaOperacoesTempo_EmTrans(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOperacoesTempo As New ClassOperacoesTempo

On Error GoTo Erro_RoteirosDeFabricacao_GravaOperacoesTempo_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134861
        
    Set objOperacoesTempo = objOperacoes.objOperacoesTempo
        
    objOperacoesTempo.lNumIntDocOperacao = objOperacoes.lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO OperacoesTempo (NumIntDocOperacao, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, TempoOperacao, UMTempo, Tipo, NumIntDocMaq) VALUES (?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
                                        objOperacoesTempo.lNumIntDocOperacao, objOperacoesTempo.dLoteMax, objOperacoesTempo.dLoteMin, objOperacoesTempo.dLotePadrao, objOperacoesTempo.dTempoPreparacao, objOperacoesTempo.dTempoMovimentacao, objOperacoesTempo.dTempoDescarga, objOperacoesTempo.dTempoOperacao, objOperacoesTempo.sUMTempo, objOperacoesTempo.iTipo, objOperacoesTempo.lNumIntDocMaq)
    If lErro <> AD_SQL_SUCESSO Then gError 134862
                           
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    RoteirosDeFabricacao_GravaOperacoesTempo_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_GravaOperacoesTempo_EmTrans:

    RoteirosDeFabricacao_GravaOperacoesTempo_EmTrans = gErr

    Select Case gErr

        Case 134861
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 134862
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OPERACOESTEMPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152256)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RoteirosDeFabricacao_GravaOperacaoInsumos_EmTrans(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOperacaoInsumos As ClassOperacaoInsumos

On Error GoTo Erro_RoteirosDeFabricacao_GravaOperacaoInsumos_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134863
        
    'Para cada objOperacaoInsumos na coleção
    For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
        
        objOperacaoInsumos.lNumIntDocOper = objOperacoes.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO OperacaoInsumos (NumIntDocOper, Produto, Quantidade, UMProduto, VersaoKitComp, Composicao, PercentualPerda, CustoStandard) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", _
                                            objOperacaoInsumos.lNumIntDocOper, objOperacaoInsumos.sProduto, objOperacaoInsumos.dQuantidade, objOperacaoInsumos.sUMProduto, objOperacaoInsumos.sVersaoKitComp, objOperacaoInsumos.iComposicao, objOperacaoInsumos.dPercentualPerda, objOperacaoInsumos.dCustoStandard)
        If lErro <> AD_SQL_SUCESSO Then gError 134864
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    RoteirosDeFabricacao_GravaOperacaoInsumos_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_GravaOperacaoInsumos_EmTrans:

    RoteirosDeFabricacao_GravaOperacaoInsumos_EmTrans = gErr

    Select Case gErr

        Case 134863
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134864
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OPERACAOINSUMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152257)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function RoteirosDeFabricacao_Exclui(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RoteirosDeFabricacao_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 134865

    lErro = CF("RoteirosDeFabricacao_Exclui_EmTrans", objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134866

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 134867

    RoteirosDeFabricacao_Exclui = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Exclui:

    RoteirosDeFabricacao_Exclui = gErr

    Select Case gErr

        Case 134865
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 134866
            'erro tratado na função chamada

        Case 134867
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152258)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RoteirosDeFabricacao_Exclui_EmTrans(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RoteirosDeFabricacao_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134868
    Next

    'Le a tabelaRoteirosDeFabricacao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RoteirosDeFabricacao WHERE ProdutoRaiz= ?  AND Versao= ? ", _
                0, lNumIntDoc, objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 134869

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134870

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 134871
            
    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 134873
                
    objRoteirosDeFabricacao.lNumIntDoc = lNumIntDoc
    
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RoteirosDeFabricacao", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 134875
    
    lErro = RoteirosDeFabricacao_ExcluiOperacoes_EmTrans(objRoteirosDeFabricacao)
    If lErro <> SUCESSO Then gError 134876
            
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteirosDeFabricacao_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Exclui_EmTrans:

    RoteirosDeFabricacao_Exclui_EmTrans = gErr

    Select Case gErr

        Case 134868
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134869, 134870, 134877
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)

        Case 134871
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSDEFABRICACAO_NAO_CADASTRADO", gErr, objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao)

        Case 134876
            'erro tratado na rotina chamada
            
        Case 134873
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ROTEIROSDEFABRICACAO", gErr)
            
        Case 134875
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ROTEIROSDEFABRICACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152259)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RoteirosDeFabricacao_ExcluiOperacoes_EmTrans(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDocOper As Long
Dim objOperacoes As ClassOperacoes

On Error GoTo Erro_RoteirosDeFabricacao_ExcluiOperacoes_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134878
    Next

    'Le a Operacoes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Operacoes WHERE NumIntDocRotFabr= ? ", _
                0, lNumIntDocOper, objRoteirosDeFabricacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134879

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134880

    'Faz enquanto existir
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objOperacoes = New ClassOperacoes
        
        objOperacoes.lNumIntDoc = lNumIntDocOper

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 134881
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Operacoes", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 134882
        
        lErro = RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans(objOperacoes)
        If lErro <> SUCESSO Then gError 134883
    
        lErro = RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans(objOperacoes)
        If lErro <> SUCESSO Then gError 134884
    
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134885

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteirosDeFabricacao_ExcluiOperacoes_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_ExcluiOperacoes_EmTrans:

    RoteirosDeFabricacao_ExcluiOperacoes_EmTrans = gErr

    Select Case gErr

        Case 134878
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134879, 134880, 134885
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOES", gErr)

        Case 134881
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OPERACOES", gErr)

        Case 134882
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPERACOES", gErr)
            
        Case 134883, 134884
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152260)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134886
    Next
    
    'Le a OperacoesTempo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM OperacoesTempo WHERE NumIntDocOperacao= ? ", _
                0, iAux, objOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134887

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134888
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 134889
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM OperacoesTempo", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 134890
            
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans:

    RoteirosDeFabricacao_ExcluiOperacoesTempo_EmTrans = gErr

    Select Case gErr

        Case 134886
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134887, 134888
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOESTEMPO", gErr)

        Case 134889
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OPERACOESTEMPO", gErr)

        Case 134890
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPERACOESTEMPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152261)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134891
    Next
    
    'Le a OperacoesTempo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM OperacaoInsumos WHERE NumIntDocOper= ? ", _
                0, iAux, objOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134892

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134893
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 134894
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM OperacaoInsumos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 134895

        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134896

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans:

    RoteirosDeFabricacao_ExcluiOperacaoInsumos_EmTrans = gErr

    Select Case gErr

        Case 134891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134892, 134893, 134896
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACAOINSUMOS", gErr)

        Case 134894
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OPERACAOINSUMOS", gErr)

        Case 134895
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPERACAOINSUMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152262)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function RoteirosDeFabricacao_CriaProdutoKit(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao, objKit As ClassKit) As Long

Dim lErro As Long
Dim objOperacoes As New ClassOperacoes
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objProdutoKit As ClassProdutoKit
Dim objProdutoKit1 As New ClassProdutoKit
Dim iSeq As Integer
Dim bProdutoNovoNoKit As Boolean

On Error GoTo Erro_RoteirosDeFabricacao_CriaProdutoKit

    iSeq = 1
    
    Set objProdutoKit = New ClassProdutoKit
    
    'Inclui o próprio Kit como ProdutoKit na colecao
    objProdutoKit.dQuantidade = objRoteirosDeFabricacao.dQuantidade
    objProdutoKit.iComposicao = objRoteirosDeFabricacao.iComposicao
    objProdutoKit.iNivel = 0
    objProdutoKit.iSeq = iSeq
    objProdutoKit.iSeqPai = 0
    objProdutoKit.sProduto = objRoteirosDeFabricacao.sProdutoRaiz
    objProdutoKit.sUnidadeMed = objRoteirosDeFabricacao.sUM
    objProdutoKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
    objProdutoKit.sVersao = objRoteirosDeFabricacao.sVersao
    objProdutoKit.iPosicaoArvore = iSeq - 1
    objProdutoKit.dPercentualPerda = objRoteirosDeFabricacao.dPercentualPerda
    objProdutoKit.dCustoStandard = objRoteirosDeFabricacao.dCustoStandard
    objProdutoKit.sVersaoKitComp = ""
    
    objKit.colComponentes.Add objProdutoKit
        
    'Percorre as Operacoes ...
    For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes
    
        'Percorre os Insumos ...
        For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
        
            bProdutoNovoNoKit = True
        
            'Verifica se já tem igual na colecao e não é o nível zero
            For Each objProdutoKit1 In objKit.colComponentes
            
                'Se já tem soma a quantidade
                If objProdutoKit1.sProduto = objOperacaoInsumos.sProduto And _
                    objProdutoKit1.iComposicao = objOperacaoInsumos.iComposicao And _
                    objProdutoKit1.sUnidadeMed = objOperacaoInsumos.sUMProduto And _
                    objProdutoKit1.iNivel <> 0 Then
                    
                    objProdutoKit1.dQuantidade = objProdutoKit1.dQuantidade + objOperacaoInsumos.dQuantidade
                    
                    bProdutoNovoNoKit = False
                    
                End If
                
            Next
            
            If bProdutoNovoNoKit Then
                
                Set objProdutoKit = New ClassProdutoKit
                
                objProdutoKit.dQuantidade = objOperacaoInsumos.dQuantidade
                objProdutoKit.iComposicao = objOperacaoInsumos.iComposicao
                objProdutoKit.iNivel = 1
                objProdutoKit.iSeq = iSeq
                objProdutoKit.iSeqPai = 1
                objProdutoKit.sProduto = objOperacaoInsumos.sProduto
                objProdutoKit.sUnidadeMed = objOperacaoInsumos.sUMProduto
                objProdutoKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
                objProdutoKit.sVersao = objRoteirosDeFabricacao.sVersao
                objProdutoKit.iPosicaoArvore = iSeq - 1
                objProdutoKit.dPercentualPerda = objOperacaoInsumos.dPercentualPerda
                objProdutoKit.dCustoStandard = objOperacaoInsumos.dCustoStandard
                objProdutoKit.sVersaoKitComp = objOperacaoInsumos.sVersaoKitComp
                
                objKit.colComponentes.Add objProdutoKit
                
                iSeq = iSeq + 1
                            
            End If
            
        Next
    
    Next
    
    RoteirosDeFabricacao_CriaProdutoKit = SUCESSO
    
    Exit Function

Erro_RoteirosDeFabricacao_CriaProdutoKit:

    RoteirosDeFabricacao_CriaProdutoKit = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152263)
    
    End Select
    
    Exit Function

End Function

Function Competencias_Automatico(lCodigo As Long) As Long
'Obtem o próximo número de Competencia

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_Competencias_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_PROX_COMPETENCIA", "Competencias", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 134310

    lCodigo = lNumAuto

    Competencias_Automatico = SUCESSO

    Exit Function

Erro_Competencias_Automatico:

    Competencias_Automatico = gErr

   Select Case gErr

        'Erro já tratado
        Case 134310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152264)

    End Select

    Exit Function

End Function


Function CentroDeTrabalho_Automatico(lCodigo As Long) As Long
'Obtem o próximo número de CentrosDeTrabalho

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_CentroDeTrabalho_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_PROX_CENTRODETRABALHO", "CentroDeTrabalho", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 134419

    lCodigo = lNumAuto

    CentroDeTrabalho_Automatico = SUCESSO

    Exit Function

Erro_CentroDeTrabalho_Automatico:

    CentroDeTrabalho_Automatico = gErr

   Select Case gErr

        'Erro já tratado
        Case 134419

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152265)

    End Select

    Exit Function

End Function

Function CentrodeTrabalho_ExcluiMaquinas_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocMaq As Long
Dim objCTMaquinas As ClassCTMaquinas

On Error GoTo Erro_CentrodeTrabalho_ExcluiMaquinas_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134440
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDocMaq FROM CTMaquinas WHERE NumIntDocCT = ?", 0, lNumIntDocMaq, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134441

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134442

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 134443
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM CTMaquinas", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 134444
        
        'inicia o objCTMaquinas
        Set objCTMaquinas = New ClassCTMaquinas
        
        objCTMaquinas.lNumIntDocMaq = lNumIntDocMaq
        objCTMaquinas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
        
        'Exclui os Turnos da Maquina
        lErro = CentrodeTrabalho_ExcluiMaqTurnos_EmTrans(objCTMaquinas)
        If lErro <> AD_SQL_SUCESSO Then gError 137217
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134445

    Loop

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CentrodeTrabalho_ExcluiMaquinas_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_ExcluiMaquinas_EmTrans:

    CentrodeTrabalho_ExcluiMaquinas_EmTrans = gErr

    Select Case gErr

        Case 134440
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134441, 134442, 134445
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINAS", gErr)

        Case 134443
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQUINAS", gErr)

        Case 134444
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTMAQUINAS", gErr)
            
        Case 137217
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152266)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CentrodeTrabalho_ExcluiMaqTurnos_EmTrans(ByVal objCTMaquinas As ClassCTMaquinas) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CentrodeTrabalho_ExcluiMaqTurnos_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137218
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT 1 FROM CTMaquinasTurnos WHERE NumIntDocMaq= ? And NumIntDocCT= ?", 0, iAux, objCTMaquinas.lNumIntDocMaq, objCTMaquinas.lNumIntDocCT)
    If lErro <> AD_SQL_SUCESSO Then gError 137219

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137220

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 137221
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM CTMaquinasTurnos", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 137222
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137223
    
    Loop

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CentrodeTrabalho_ExcluiMaqTurnos_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_ExcluiMaqTurnos_EmTrans:

    CentrodeTrabalho_ExcluiMaqTurnos_EmTrans = gErr

    Select Case gErr

        Case 137218
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137219, 137220, 137223
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINASTURNOS", gErr)

        Case 137221
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQUINASTURNOS", gErr)

        Case 137222
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTMAQUINASTURNOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152267)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CentrodeTrabalho_ExcluiCompetencias_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho, Optional bExcluindoCT As Boolean = True) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDocCompet As Long
Dim objCTCompetencias As ClassCTCompetencias
Dim bConsisteExclusao As Boolean
Dim objCompetencias As ClassCompetencias
Dim iAux As Integer

On Error GoTo Erro_CentrodeTrabalho_ExcluiCompetencias_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134434
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDocCompet FROM CTCompetencias WHERE NumIntDocCT = ?", 0, lNumIntDocCompet, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134435

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134436

    Do While lErro <> AD_SQL_SEM_DADOS
    
        bConsisteExclusao = True
        
        If Not bExcluindoCT Then
        
            'Verifica se a Competência será incluída novamente - Alteração
            For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
                
                If lNumIntDocCompet = objCTCompetencias.lNumIntDocCompet Then
                   bConsisteExclusao = False
                   Exit For
                   
                End If
                                       
            Next
        
        End If
                
        If bConsisteExclusao = True Then
        
            'Verifica se o CT é Padrão na Competencia excluída
            Set objCompetencias = New ClassCompetencias
            
            objCompetencias.lNumIntDoc = lNumIntDocCompet
            
            'Lê a Competencia pelo NumIntDoc
            lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
            If lErro <> SUCESSO And lErro <> 134336 Then gError 134549
            
            'Se o CT é Padrão desta Competência, Erro.
            If objCompetencias.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc Then gError 134924
        
            'Verifica se o Centro de Trabalho e Competência já foi utilizado por alguma Operação
            lErro = Comando_Executar(alComando(2), "SELECT 1 FROM RoteirosDeFabricacao, Operacoes WHERE RoteirosDeFabricacao.NumIntDoc = Operacoes.NumIntDocRotFabr AND Operacoes.NumIntDocCT = ? AND Operacoes.NumIntDocCompet = ?", iAux, objCentrodeTrabalho.lNumIntDoc, lNumIntDocCompet)
            If lErro <> AD_SQL_SUCESSO Then gError 134544
        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134545
        
            'se o par já foi utilizado em uma Operação, Erro.
            If lErro = AD_SQL_SUCESSO Then gError 134546
                            
        End If

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 134437
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM CTCompetencias", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 134438
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134439

    Loop

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CentrodeTrabalho_ExcluiCompetencias_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_ExcluiCompetencias_EmTrans:

    CentrodeTrabalho_ExcluiCompetencias_EmTrans = gErr

    Select Case gErr

        Case 134434
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134435, 134436, 134439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTCOMPETENCIAS", gErr)

        Case 134437
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTCOMPETENCIAS", gErr)

        Case 134438
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTCOMPETENCIAS", gErr)
        
        Case 134544, 134545
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)
        
        Case 134546
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_VINCULADA_ROTEIROS", gErr, objCompetencias.lCodigo)

        Case 134549
            'erro tratado na rotina chamada
        
        Case 134924
            Call Rotina_Erro(vbOKOnly, "ERRO_CTPADRAO_DA_COMPETENCIA_EXCLUIDA", gErr, objCompetencias.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152268)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CentrodeTrabalho_GravaMaquinas_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim objTurno As New ClassTurno

On Error GoTo Erro_CentrodeTrabalho_GravaMaquinas_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134418
        
    'Para cada CTMaquina na coleção
    For Each objCTMaquinas In objCentrodeTrabalho.colMaquinas
        
        objCTMaquinas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO CTMaquinas (NumIntDocMaq, NumIntDocCT, Quantidade) VALUES (?, ?, ?)", objCTMaquinas.lNumIntDocMaq, objCTMaquinas.lNumIntDocCT, objCTMaquinas.iQuantidade)
        If lErro <> AD_SQL_SUCESSO Then gError 134419
        
        'Para cada Turno na coleção de CTMaquinas
        For Each objTurno In objCTMaquinas.colTurnos
        
            objTurno.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq
            objTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
            
            lErro = Comando_Executar(lComando, "INSERT INTO CTMaquinasTurnos (NumIntDocMaq, NumIntDocCT, Turno, HorasDomingo, HorasSegunda, HorasTerca, HorasQuarta, HorasQuinta, HorasSexta, HorasSabado) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
                                    objTurno.lNumIntDocMaq, objTurno.lNumIntDocCT, objTurno.iTurno, objTurno.dQtdHoras(DOMINGO), objTurno.dQtdHoras(SEGUNDA), objTurno.dQtdHoras(TERCA), objTurno.dQtdHoras(QUARTA), objTurno.dQtdHoras(QUINTA), objTurno.dQtdHoras(SEXTA), objTurno.dQtdHoras(SABADO))
            If lErro <> AD_SQL_SUCESSO Then gError 137216
        
        Next
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_GravaMaquinas_EmTrans = SUCESSO
    
    Exit Function

Erro_CentrodeTrabalho_GravaMaquinas_EmTrans:

    CentrodeTrabalho_GravaMaquinas_EmTrans = gErr

    Select Case gErr

        Case 134418
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 134419
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTMAQUINAS", gErr)
            
        Case 137216
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTMAQUINASTURNOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152269)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function CentrodeTrabalho_GravaCompetencias_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTCompetencias As ClassCTCompetencias
Dim objCompetencias As ClassCompetencias

On Error GoTo Erro_CentrodeTrabalho_GravaCompetencias_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134416
        
    'Para cada CTCompetencia na coleção
    For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
        objCTCompetencias.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO CTCompetencias (NumIntDocCT, NumIntDocCompet) VALUES (?, ?)", objCTCompetencias.lNumIntDocCT, objCTCompetencias.lNumIntDocCompet)
        If lErro <> AD_SQL_SUCESSO Then gError 134417
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objCTCompetencias.lNumIntDocCompet
        
        'Lê a Competencia pelo NumIntDoc
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 134925
        
        'Se o CT Padrão está vazio -> este será o Padrão
        If objCompetencias.lNumIntDocCT = 0 Then
        
            lErro = Competencias_Grava_CTPadrao(objCTCompetencias)
            If lErro <> SUCESSO Then gError 134926
        
        End If
        
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_GravaCompetencias_EmTrans = SUCESSO
    
    Exit Function

Erro_CentrodeTrabalho_GravaCompetencias_EmTrans:

    CentrodeTrabalho_GravaCompetencias_EmTrans = gErr

    Select Case gErr

        Case 134416
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 134417
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTCOMPETENCIAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152270)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TiposDeMaodeObra_Grava(ByVal objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TiposDeMaodeObra_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137577

    lErro = CF("TiposDeMaodeObra_Grava_EmTrans", objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137578

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137579

    TiposDeMaodeObra_Grava = SUCESSO

    Exit Function

Erro_TiposDeMaodeObra_Grava:

    TiposDeMaodeObra_Grava = gErr

    Select Case gErr

        Case 137577
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137578
            'erro tratado na rotina chamada

        Case 137579
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152271)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TiposDeMaodeObra_Grava_EmTrans(ByVal objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_TiposDeMaodeObra_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137580
    Next

    'Le a tabelaTiposDeMaodeObra
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM TiposDeMaodeObra WHERE Codigo= ? ", _
                0, iAux, objTiposDeMaodeObra.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 137581

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137582

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 137583

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposDeMaodeObra SET Descricao= ? , Observacao= ? , CustoHora= ?, Produto = ? ", alComando(0), _
                objTiposDeMaodeObra.sDescricao, objTiposDeMaodeObra.sObservacao, objTiposDeMaodeObra.dCustoHora, objTiposDeMaodeObra.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 137584

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO TiposDeMaodeObra( Codigo, Descricao, Observacao, CustoHora, Produto) VALUES (?,?,?,?,?)", _
                objTiposDeMaodeObra.iCodigo, objTiposDeMaodeObra.sDescricao, objTiposDeMaodeObra.sObservacao, objTiposDeMaodeObra.dCustoHora, objTiposDeMaodeObra.sProduto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137585

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TiposDeMaodeObra_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TiposDeMaodeObra_Grava_EmTrans:

    TiposDeMaodeObra_Grava_EmTrans = gErr

    Select Case gErr

        Case 137580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137581, 137582
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEMAODEOBRA", gErr)

        Case 137583
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TIPOSDEMAODEOBRA", gErr)

        Case 137584
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TIPOSDEMAODEOBRA", gErr)

        Case 137585
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOSDEMAODEOBRA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152272)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TiposDeMaodeObra_Exclui(ByVal objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TiposDeMaodeObra_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137586

    lErro = CF("TiposDeMaodeObra_Exclui_EmTrans", objTiposDeMaodeObra)
    If lErro <> SUCESSO Then gError 137587

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137588

    TiposDeMaodeObra_Exclui = SUCESSO

    Exit Function

Erro_TiposDeMaodeObra_Exclui:

    TiposDeMaodeObra_Exclui = gErr

    Select Case gErr

        Case 137586
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137587
            'erro tratado na rotina chamada

        Case 137588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152273)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TiposDeMaodeObra_Exclui_EmTrans(ByVal objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim iAux As Integer

On Error GoTo Erro_TiposDeMaodeObra_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137589
    Next

    'Le a tabelaTiposDeMaodeObra
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TiposDeMaodeObra WHERE Codigo= ? ", _
                0, iCodigo, objTiposDeMaodeObra.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 137590

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137591

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 137592
        
    'Lê MaquinaOperadores
    lErro = Comando_Executar(alComando(2), "SELECT 1 FROM MaquinaOperadores WHERE TipoMaoDeObra = ?", iAux, iCodigo)
    If lErro <> SUCESSO Then gError 137910
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137911

    'Se Encontrou algum ... Erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 137912

    'Lê CTOperadores
    lErro = Comando_Executar(alComando(3), "SELECT 1 FROM CTOperadores WHERE TipoMaoDeObra = ?", iAux, iCodigo)
    If lErro <> SUCESSO Then gError 139102
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 139103

    'Se Encontrou algum ... Erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 139104

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 137593

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TiposDeMaodeObra", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 137594

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TiposDeMaodeObra_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_TiposDeMaodeObra_Exclui_EmTrans:

    TiposDeMaodeObra_Exclui_EmTrans = gErr

    Select Case gErr

        Case 137589
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137590, 137591, 137910, 137911
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEMAODEOBRA", gErr)

        Case 137592
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)

        Case 137593
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TIPOSDEMAODEOBRA", gErr)

        Case 137594
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOSDEMAODEOBRA", gErr)
        
        Case 137912
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOMO_USADA_EM_MAQUINAS", gErr)
            
        Case 139102, 139103
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTOPERADORES", gErr)

        Case 139104
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOMO_USADA_EM_CT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152274)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
Function TiposDeMaodeObra_Automatico(iCodigo As Integer) As Long
'Obtem o próximo número de TiposDeMaodeObra

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_TiposDeMaodeObra_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_INT_PROX_TIPOSDEMAODEOBRA", "TiposDeMaodeObra", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 137575

    If lNumAuto > NUM_MAX_TIPOMO Then gError 137576
        
    iCodigo = CInt(lNumAuto)

    TiposDeMaodeObra_Automatico = SUCESSO

    Exit Function

Erro_TiposDeMaodeObra_Automatico:

    TiposDeMaodeObra_Automatico = gErr

    Select Case gErr
    
        Case 137575
            'Erro já tratado na rotina chamada
    
        Case 137576
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, lNumAuto)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152275)
    
    End Select
    
    Exit Function

End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function CTMaquinaProgDisponibilidade_Grava(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136574

    lErro = CF("CTMaquinaProgDisponibilidade_Grava_EmTrans", objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 136575

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136576

    CTMaquinaProgDisponibilidade_Grava = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Grava:

    CTMaquinaProgDisponibilidade_Grava = gErr

    Select Case gErr

        Case 136574
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 136575

        Case 136576
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152276)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgDisponibilidade_Grava_Periodo(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iPeriodoEmDias As Long
Dim iIndice As Integer
Dim objCTMaqProgDispUnitario As ClassCTMaqProgDisp

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Grava_Periodo

    dtDataInicial = objCTMaqProgDisp.dtDataDe
    dtDataFinal = objCTMaqProgDisp.dtDataAte
    
    'verifica qual o período
    iPeriodoEmDias = CInt(DateDiff("d", dtDataInicial, dtDataFinal))

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137240
    
    'para cada dia do período passado
    For iIndice = 0 To iPeriodoEmDias
    
        Set objCTMaqProgDispUnitario = New ClassCTMaqProgDisp
        
        With objCTMaqProgDispUnitario
        
            .lNumIntDocCT = objCTMaqProgDisp.lNumIntDocCT
            .lNumIntDocMaq = objCTMaqProgDisp.lNumIntDocMaq
            .iQuantidade = objCTMaqProgDisp.iQuantidade
            .sObservacao = objCTMaqProgDisp.sObservacao
            
            'se é a primeira data... não implementa nada
            If iIndice = 0 Then
                .dtData = objCTMaqProgDisp.dtData
            Else
                .dtData = DateAdd("d", iIndice, objCTMaqProgDisp.dtDataDe)
            End If
        
        End With
        
        lErro = CF("CTMaquinaProgDisponibilidade_Grava_EmTrans", objCTMaqProgDispUnitario)
        If lErro <> SUCESSO Then gError 137241
        
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137242

    CTMaquinaProgDisponibilidade_Grava_Periodo = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Grava_Periodo:

    CTMaquinaProgDisponibilidade_Grava_Periodo = gErr

    Select Case gErr

        Case 137240
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137241

        Case 137242
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152277)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgDisponibilidade_Grava_EmTrans(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136577
    Next

    'Le a tabelaCTMaquinaProgDisponibilidade
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CTMaquinaProgDisponibilidade WHERE NumIntDocCT= ?  AND NumIntDocMaq= ?  AND Data= ? ", _
                0, lNumIntDoc, objCTMaqProgDisp.lNumIntDocCT, objCTMaqProgDisp.lNumIntDocMaq, objCTMaqProgDisp.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 136578

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136579

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        objCTMaqProgDisp.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 136580
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE CTMaquinaProgDisponibilidade SET Quantidade= ? , Observacao= ? ", alComando(0), _
                objCTMaqProgDisp.iQuantidade, objCTMaqProgDisp.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 136581

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_CTMAQUINAPROGDISP", lNumIntDoc)
        If lErro <> SUCESSO Then gError 134308

        objCTMaqProgDisp.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO CTMaquinaProgDisponibilidade(NumIntDoc, NumIntDocCT, NumIntDocMaq, Data, Quantidade, " & _
                "Observacao) VALUES (?,?,?,?,?,?)", _
                objCTMaqProgDisp.lNumIntDoc, objCTMaqProgDisp.lNumIntDocCT, objCTMaqProgDisp.lNumIntDocMaq, objCTMaqProgDisp.dtData, objCTMaqProgDisp.iQuantidade, _
                objCTMaqProgDisp.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136582

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinaProgDisponibilidade_Grava_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Grava_EmTrans:

    CTMaquinaProgDisponibilidade_Grava_EmTrans = gErr

    Select Case gErr

        Case 136577
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136578, 136579
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQPROGDISP", gErr)

        Case 136580
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQPROGDISP", gErr)

        Case 136581
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CTMAQPROGDISP", gErr)

        Case 136582
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTMAQPROGDISP", gErr)
            
        Case 134308
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152278)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function CTMaquinaProgDisponibilidade_Exclui(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136565

    lErro = CF("CTMaquinaProgDisponibilidade_Exclui_EmTrans", objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 136566

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136567

    CTMaquinaProgDisponibilidade_Exclui = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Exclui:

    CTMaquinaProgDisponibilidade_Exclui = gErr

    Select Case gErr

        Case 136565
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 136566

        Case 136567
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152279)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgDisponibilidade_Exclui_Periodo(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iPeriodoEmDias As Long
Dim iIndice As Integer
Dim objCTMaqProgDispUnitario As ClassCTMaqProgDisp

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Exclui_Periodo

    dtDataInicial = objCTMaqProgDisp.dtDataDe
    dtDataFinal = objCTMaqProgDisp.dtDataAte
    
    'verifica qual o período
    iPeriodoEmDias = CInt(DateDiff("d", dtDataInicial, dtDataFinal))

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137236
    
    'para cada dia do período passado
    For iIndice = 0 To iPeriodoEmDias
    
        Set objCTMaqProgDispUnitario = New ClassCTMaqProgDisp
        
        With objCTMaqProgDispUnitario
        
            .lNumIntDocCT = objCTMaqProgDisp.lNumIntDocCT
            .lNumIntDocMaq = objCTMaqProgDisp.lNumIntDocMaq
            
            'se é a primeira data... não implementa nada
            If iIndice = 0 Then
                .dtData = objCTMaqProgDisp.dtData
            Else
                .dtData = DateAdd("d", iIndice, objCTMaqProgDisp.dtDataDe)
            End If
        
        End With
        
        lErro = CF("CTMaquinaProgDisponibilidade_Le", objCTMaqProgDispUnitario)
        If lErro <> SUCESSO And lErro <> 136564 Then gError 137237
        
        If lErro = SUCESSO Then

            lErro = CF("CTMaquinaProgDisponibilidade_Exclui_EmTrans", objCTMaqProgDispUnitario)
            If lErro <> SUCESSO Then gError 137238
        
        End If
        
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137239

    CTMaquinaProgDisponibilidade_Exclui_Periodo = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Exclui_Periodo:

    CTMaquinaProgDisponibilidade_Exclui_Periodo = gErr

    Select Case gErr

        Case 137236
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137237, 137238

        Case 137239
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152280)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function


Public Function CTMaquinaProgDisponibilidade_Exclui_EmTrans(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136568
    Next

    'Le a tabelaCTMaquinaProgDisponibilidade
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM CTMaquinaProgDisponibilidade WHERE NumIntDocCT= ?  AND NumIntDocMaq= ?  AND Data= ? ", _
                0, iAux, objCTMaqProgDisp.lNumIntDocCT, objCTMaqProgDisp.lNumIntDocMaq, objCTMaqProgDisp.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 136569

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136570

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 136571

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 136572

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CTMaquinaProgDisponibilidade", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 136573

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinaProgDisponibilidade_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Exclui_EmTrans:

    CTMaquinaProgDisponibilidade_Exclui_EmTrans = gErr

    Select Case gErr

        Case 136568
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136569, 136570
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQPROGDISP", gErr)

        Case 136571
            Call Rotina_Erro(vbOKOnly, "ERRO_CTMAQPROGDISP_NAO_CADASTRADO", gErr, objCTMaqProgDisp.lNumIntDocCT, objCTMaqProgDisp.lNumIntDocMaq, objCTMaqProgDisp.dtData)

        Case 136572
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQPROGDISP", gErr)

        Case 136573
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTMAQPROGDISP", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152281)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function CTMaquinaProgTurno_Grava(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTMaquinaProgTurno_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136714

    lErro = CF("CTMaquinaProgTurno_Grava_EmTrans", objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 136715

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136716

    CTMaquinaProgTurno_Grava = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Grava:

    CTMaquinaProgTurno_Grava = gErr

    Select Case gErr

        Case 136714
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 136715

        Case 136716
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152282)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgTurno_Grava_Periodo(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iPeriodoEmDias As Long
Dim iIndice As Integer
Dim objCTMaqProgTurnoUnitario As ClassCTMaqProgTurno

On Error GoTo Erro_CTMaquinaProgTurno_Grava_Periodo

    dtDataInicial = objCTMaqProgTurno.dtDataDe
    dtDataFinal = objCTMaqProgTurno.dtDataAte
    
    'verifica qual o período
    iPeriodoEmDias = CInt(DateDiff("d", dtDataInicial, dtDataFinal))

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137256

    'para cada dia do período passado
    For iIndice = 0 To iPeriodoEmDias
    
        Set objCTMaqProgTurnoUnitario = New ClassCTMaqProgTurno
        
        With objCTMaqProgTurnoUnitario
        
            .lNumIntDocCT = objCTMaqProgTurno.lNumIntDocCT
            .lNumIntDocMaq = objCTMaqProgTurno.lNumIntDocMaq
            .sObservacao = objCTMaqProgTurno.sObservacao
            Set .colTurnos = objCTMaqProgTurno.colTurnos
            
            'se é a primeira data... não implementa nada
            If iIndice = 0 Then
                .dtData = objCTMaqProgTurno.dtData
            Else
                .dtData = DateAdd("d", iIndice, objCTMaqProgTurno.dtDataDe)
            End If
        
        End With
        
        lErro = CF("CTMaquinaProgTurno_Grava_EmTrans", objCTMaqProgTurnoUnitario)
        If lErro <> SUCESSO Then gError 137257
        
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137258

    CTMaquinaProgTurno_Grava_Periodo = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Grava_Periodo:

    CTMaquinaProgTurno_Grava_Periodo = gErr

    Select Case gErr

        Case 137256
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137257

        Case 137258
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152283)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgTurno_Grava_EmTrans(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CTMaquinaProgTurno_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136718
    Next

    'Le a tabelaCTMaquinaProgTurno
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CTMaquinaProgTurno WHERE NumIntDocCT= ?  AND NumIntDocMaq= ?  AND Data= ? ", _
                0, lNumIntDoc, objCTMaqProgTurno.lNumIntDocCT, objCTMaqProgTurno.lNumIntDocMaq, objCTMaqProgTurno.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 136719

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136720

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then
    
        objCTMaqProgTurno.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 136721

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE CTMaquinaProgTurno SET Observacao= ? ", alComando(0), _
                objCTMaqProgTurno.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 136722
        
        lErro = CTMaquinaProgTurno_ExcluiItens_EmTrans(objCTMaqProgTurno)
        If lErro <> SUCESSO Then gError 137254

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_CTMAQUINAPROGTURNO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 136717

        objCTMaqProgTurno.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO CTMaquinaProgTurno(NumIntDoc, NumIntDocCT, NumIntDocMaq, Data, " & _
                "Observacao) VALUES (?,?,?,?,?)", _
                objCTMaqProgTurno.lNumIntDoc, objCTMaqProgTurno.lNumIntDocCT, objCTMaqProgTurno.lNumIntDocMaq, objCTMaqProgTurno.dtData, _
                objCTMaqProgTurno.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136723
        
    End If

    lErro = CTMaquinaProgTurno_GravaItens_EmTrans(objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137255

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinaProgTurno_Grava_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Grava_EmTrans:

    CTMaquinaProgTurno_Grava_EmTrans = gErr

    Select Case gErr

        Case 136718
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136719, 136720
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINAPROGTURNO", gErr)

        Case 136721
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQUINAPROGTURNO", gErr)

        Case 136722
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CTMAQUINAPROGTURNO", gErr)

        Case 136723
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTMAQUINAPROGTURNO", gErr)
            
        Case 137254, 137255

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152284)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CTMaquinaProgTurno_GravaItens_EmTrans(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTMaqProgTurnoItens As New ClassCTMaqProgTurnoItens

On Error GoTo Erro_CTMaquinaProgTurno_GravaItens_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137259
        
    'Para cada objTurno na coleção
    For Each objCTMaqProgTurnoItens In objCTMaqProgTurno.colTurnos
        
        objCTMaqProgTurnoItens.lNumIntDocCTMaqPTurno = objCTMaqProgTurno.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO CTMaquinaProgTurnoItens (NumIntDocCTMaqPTurno, Turno, Horas) VALUES (?, ?, ?)", _
                                objCTMaqProgTurnoItens.lNumIntDocCTMaqPTurno, objCTMaqProgTurnoItens.iTurno, objCTMaqProgTurnoItens.dHoras)
        If lErro <> AD_SQL_SUCESSO Then gError 137260
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    CTMaquinaProgTurno_GravaItens_EmTrans = SUCESSO
    
    Exit Function

Erro_CTMaquinaProgTurno_GravaItens_EmTrans:

    CTMaquinaProgTurno_GravaItens_EmTrans = gErr

    Select Case gErr

        Case 137259
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137260
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTMAQPROGTURNOITENS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152285)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function CTMaquinaProgTurno_Exclui(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTMaquinaProgTurno_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136705

    lErro = CF("CTMaquinaProgTurno_Exclui_EmTrans", objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 136706

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136707

    CTMaquinaProgTurno_Exclui = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Exclui:

    CTMaquinaProgTurno_Exclui = gErr

    Select Case gErr

        Case 136705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 136706

        Case 136707
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152286)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgTurno_Exclui_Periodo(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iPeriodoEmDias As Long
Dim iIndice As Integer
Dim objCTMaqProgTurnoUnitario As ClassCTMaqProgTurno

On Error GoTo Erro_CTMaquinaProgTurno_Exclui_Periodo

    dtDataInicial = objCTMaqProgTurno.dtDataDe
    dtDataFinal = objCTMaqProgTurno.dtDataAte
    
    'verifica qual o período
    iPeriodoEmDias = CInt(DateDiff("d", dtDataInicial, dtDataFinal))
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137244

    'para cada dia do período passado
    For iIndice = 0 To iPeriodoEmDias
    
        Set objCTMaqProgTurnoUnitario = New ClassCTMaqProgTurno
        
        With objCTMaqProgTurnoUnitario
        
            .lNumIntDocCT = objCTMaqProgTurno.lNumIntDocCT
            .lNumIntDocMaq = objCTMaqProgTurno.lNumIntDocMaq
            
            'se é a primeira data... não implementa nada
            If iIndice = 0 Then
                .dtData = objCTMaqProgTurno.dtData
            Else
                .dtData = DateAdd("d", iIndice, objCTMaqProgTurno.dtDataDe)
            End If
        
        End With
        
        lErro = CF("CTMaquinaProgTurno_Le", objCTMaqProgTurnoUnitario)
        If lErro <> SUCESSO And lErro <> 136704 Then gError 137245
        
        If lErro = SUCESSO Then
        
            lErro = CF("CTMaquinaProgTurno_Exclui_EmTrans", objCTMaqProgTurnoUnitario)
            If lErro <> SUCESSO Then gError 137246
        
        End If
        
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137247

    CTMaquinaProgTurno_Exclui_Periodo = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Exclui_Periodo:

    CTMaquinaProgTurno_Exclui_Periodo = gErr

    Select Case gErr

        Case 137244
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137245, 137246

        Case 137247
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152287)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinaProgTurno_Exclui_EmTrans(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CTMaquinaProgTurno_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136708
    Next

    'Le a tabelaCTMaquinaProgTurno
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CTMaquinaProgTurno WHERE NumIntDocCT= ?  AND NumIntDocMaq= ?  AND Data= ? ", _
                0, lNumIntDoc, objCTMaqProgTurno.lNumIntDocCT, objCTMaqProgTurno.lNumIntDocMaq, objCTMaqProgTurno.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 136709

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136710

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 136711
    
    objCTMaqProgTurno.lNumIntDoc = lNumIntDoc

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 136712

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CTMaquinaProgTurno", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 136713

    lErro = CTMaquinaProgTurno_ExcluiItens_EmTrans(objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137243

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinaProgTurno_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Exclui_EmTrans:

    CTMaquinaProgTurno_Exclui_EmTrans = gErr

    Select Case gErr

        Case 136708
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136709, 136710
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINAPROGTURNO", gErr)

        Case 136711
            Call Rotina_Erro(vbOKOnly, "ERRO_CTMAQUINAPROGTURNO_NAO_CADASTRADO", gErr, objCTMaqProgTurno.lNumIntDocCT, objCTMaqProgTurno.lNumIntDocMaq, objCTMaqProgTurno.dtData)

        Case 136712
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQUINAPROGTURNO", gErr)

        Case 136713
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTMAQUINAPROGTURNO", gErr)
            
        Case 137243

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152288)

        End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CTMaquinaProgTurno_ExcluiItens_EmTrans(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CTMaquinaProgTurno_ExcluiItens_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137248
    Next

    'Le a tabelaCTMaquinaProgTurnoItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM CTMaquinaProgTurnoItens WHERE NumIntDocCTMaqPTurno= ? ", _
                0, iAux, objCTMaqProgTurno.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137249

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137250

    'faz enquanto existir dados ...
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 137251
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CTMaquinaProgTurnoItens", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 137252

        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137253

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinaProgTurno_ExcluiItens_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_ExcluiItens_EmTrans:

    CTMaquinaProgTurno_ExcluiItens_EmTrans = gErr

    Select Case gErr

        Case 137248
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137249, 137250, 137253
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQPROGTURNOITENS", gErr)

        Case 137251
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQPROGTURNOITENS", gErr)

        Case 137252
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTMAQPROGTURNOITENS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152289)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CentrodeTrabalho_GravaTurnos_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objTurno As New ClassTurno

On Error GoTo Erro_CentrodeTrabalho_GravaTurnos_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137196
        
    'Para cada objTurno na coleção
    For Each objTurno In objCentrodeTrabalho.colTurnos
        
        objTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO CTTurnos (NumIntDocCT, Turno, HorasDomingo, HorasSegunda, HorasTerca, HorasQuarta, HorasQuinta, HorasSexta, HorasSabado) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", _
                                objTurno.lNumIntDocCT, objTurno.iTurno, objTurno.dQtdHoras(DOMINGO), objTurno.dQtdHoras(SEGUNDA), objTurno.dQtdHoras(TERCA), objTurno.dQtdHoras(QUARTA), objTurno.dQtdHoras(QUINTA), objTurno.dQtdHoras(SEXTA), objTurno.dQtdHoras(SABADO))
        If lErro <> AD_SQL_SUCESSO Then gError 137197
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_GravaTurnos_EmTrans = SUCESSO
    
    Exit Function

Erro_CentrodeTrabalho_GravaTurnos_EmTrans:

    CentrodeTrabalho_GravaTurnos_EmTrans = gErr

    Select Case gErr

        Case 137196
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 137197
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTTURNOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152290)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CentrodeTrabalho_ExcluiTurnos_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iTurno As Integer

On Error GoTo Erro_CentrodeTrabalho_ExcluiTurnos_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137198
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Turno FROM CTTurnos WHERE NumIntDocCT = ?", 0, iTurno, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137199

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137200

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 137201
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM CTTurnos", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 137202
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137203

    Loop

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CentrodeTrabalho_ExcluiTurnos_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_ExcluiTurnos_EmTrans:

    CentrodeTrabalho_ExcluiTurnos_EmTrans = gErr

    Select Case gErr

        Case 13198
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137199, 137200, 137203
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTTURNOS", gErr)

        Case 137201
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTTURNOS", gErr)

        Case 137202
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTTURNOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152291)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Inserido por Jorge Specian - Alterações ref. ao PCP ... (14/06/2005)
Function OrdemDeProducao_Gera_Operacoes(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim objItemOP As New ClassItemOP
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim iNivel As Integer, iNivelRoteiro As Integer, iSeqRoteiro As Integer, iSeq As Integer
Dim objKit As ClassKit

On Error GoTo Erro_OrdemDeProducao_Gera_Operacoes

    'se a data de emissão da OP for maior ou igual a data de inicio de operações MRP
    If objOrdemDeProducao.dtDataEmissao >= gobjEST.dtDataInicioMRP And gobjEST.dtDataInicioMRP <> DATA_NULA Then

        'Para cada item da OP
        For Each objItemOP In objOrdemDeProducao.colItens
        
            Set objKit = New ClassKit
    
            objKit.sProdutoRaiz = objItemOP.sProduto
        
            'Le as Versoes Ativas e a Padrao
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError ERRO_SEM_MENSAGEM
            
            If Len(Trim(objKit.sVersao)) > 0 Then
            
                Set objRoteirosDeFabricacao = New ClassRoteirosDeFabricacao
                
                objRoteirosDeFabricacao.sProdutoRaiz = objItemOP.sProduto
                objRoteirosDeFabricacao.sVersao = objItemOP.sVersao
                
                Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos
                
                objOrdemProducaoInsumos.sUMProduto = objItemOP.sSiglaUM
                objOrdemProducaoInsumos.dQuantidade = objItemOP.dQuantidade
                
                'Gera a árvore de operações
                lErro = OrdemDeProducao_Gera_Operacoes1(objRoteirosDeFabricacao, objOrdemProducaoInsumos, objItemOP, iNivel, iNivelRoteiro, iSeqRoteiro, iSeq)
                If lErro <> SUCESSO Then gError 137183
                
                'ajusta as sequencias dos itens gerados
                Call OrdemDeProducao_Gera_Operacoes2(objItemOP)
            
            End If
            
        Next
    
    End If

    OrdemDeProducao_Gera_Operacoes = SUCESSO
    
    Exit Function
    
Erro_OrdemDeProducao_Gera_Operacoes:

    OrdemDeProducao_Gera_Operacoes = gErr
    
    Select Case gErr
    
        Case 137183, 137186, ERRO_SEM_MENSAGEM
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152292)
    
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Gera_Operacoes1(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao, ByVal objOPInsumos As ClassOrdemProducaoInsumos, ByVal objItemOP As ClassItemOP, ByVal iNivel As Integer, ByVal iNivelRoteiro As Integer, iSeqRoteiro As Integer, iSeq As Integer)
'Monta a Árvore de Operações e é chamada recursivamente para os sub-produtos...
   
Dim lErro As Long
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim objOperacoes As New ClassOperacoes
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objProduto As ClassProduto
Dim dFatorUMProduto As Double
Dim dFatorUMRoteiro As Double
Dim dFatorQtdeOPInsumos As Double
Dim dQuantidade As Double
Dim objOperacoesTempo As ClassOperacoesTempo
Dim dTempoCorrigido As Double
Dim objRotFabr As ClassRoteirosDeFabricacao
Dim objKit As ClassKit
Dim iSeqRoteiroAux As Integer

On Error GoTo Erro_OrdemDeProducao_Gera_Operacoes1

    'Se a quantidade já chegou a 0.0001 ou já desceu na recursividade n instâncias
    If (objOPInsumos.dQuantidade <= 0.0001 And iNivel > 0) Or iNivel >= 20 Then
    
        OrdemDeProducao_Gera_Operacoes1 = SUCESSO
        Exit Function

    End If
    
    If Len(Trim(objRoteirosDeFabricacao.sVersao)) = 0 Then
    
        Set objKit = New ClassKit

        objKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
    
        'Le as Versoes Ativas e a Padrao
        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 139044
        
        objRoteirosDeFabricacao.sVersao = objKit.sVersao
    
    End If
    
    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 137186
    
    If lErro = SUCESSO Then
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objRoteirosDeFabricacao.sProdutoRaiz
        
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137187
        
        If lErro <> SUCESSO Then gError 137188
    
        'Descobre o fator de conversao da UM do Item da OP p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOPInsumos.sUMProduto, objProduto.sSiglaUMEstoque, dFatorUMProduto)
        If lErro <> SUCESSO Then gError 137189
    
        'Descobre o fator de conversao da UM do Roteiro p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objRoteirosDeFabricacao.sUM, objProduto.sSiglaUMEstoque, dFatorUMRoteiro)
        If lErro <> SUCESSO Then gError 137190
    
        dFatorQtdeOPInsumos = (objOPInsumos.dQuantidade * dFatorUMProduto) / (objRoteirosDeFabricacao.dQuantidade * dFatorUMRoteiro)
        
        iSeqRoteiro = iSeqRoteiro + 1
        iSeqRoteiroAux = iSeqRoteiro
        
        For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes
        
            iSeq = iSeq + 1
                        
            Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
            
            objOrdemProducaoOperacoes.lNumIntDocCompet = objOperacoes.lNumIntDocCompet
            objOrdemProducaoOperacoes.lNumIntDocCT = objOperacoes.lNumIntDocCT
            objOrdemProducaoOperacoes.sObservacao = objOperacoes.sObservacao
            objOrdemProducaoOperacoes.iNivel = iNivel + objOperacoes.iNivel
            objOrdemProducaoOperacoes.iSeq = iSeq
            objOrdemProducaoOperacoes.iNivelRoteiro = iNivelRoteiro
            objOrdemProducaoOperacoes.iSeqRoteiro = iSeqRoteiroAux
            objOrdemProducaoOperacoes.sProduto = objRoteirosDeFabricacao.sProdutoRaiz
            objOrdemProducaoOperacoes.sVersao = objRoteirosDeFabricacao.sVersao
            objOrdemProducaoOperacoes.dQuantidade = objOPInsumos.dQuantidade
            objOrdemProducaoOperacoes.sUMedida = objOPInsumos.sUMProduto
            objOrdemProducaoOperacoes.iIgnoraTaxaProducao = objOperacoes.iIgnoraTaxaProducao
            objOrdemProducaoOperacoes.iConsideraCarga = MARCADO
            objOrdemProducaoOperacoes.iOrigem = ORIGEM_ROTEIRO
            objOrdemProducaoOperacoes.iNumMaxMaqPorOper = objOperacoes.iNumMaxMaqPorOper
            objOrdemProducaoOperacoes.iNumRepeticoes = objOperacoes.iNumRepeticoes
            
            If objOperacoes.iIgnoraTaxaProducao = MARCADO Then
                                        
                Set objOperacoesTempo = New ClassOperacoesTempo
                
                objOperacoesTempo.dLoteMax = objOperacoes.objOperacoesTempo.dLoteMax
                objOperacoesTempo.dLoteMin = objOperacoes.objOperacoesTempo.dLoteMin
                objOperacoesTempo.dLotePadrao = objOperacoes.objOperacoesTempo.dLotePadrao
                objOperacoesTempo.dTempoDescarga = objOperacoes.objOperacoesTempo.dTempoDescarga
                objOperacoesTempo.dTempoMovimentacao = objOperacoes.objOperacoesTempo.dTempoMovimentacao
                If objOperacoes.objOperacoesTempo.iTipo <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
                    objOperacoesTempo.dTempoOperacao = objOperacoes.objOperacoesTempo.dTempoOperacao * dFatorQtdeOPInsumos
                Else
                    objOperacoesTempo.dTempoOperacao = objOperacoes.objOperacoesTempo.dTempoOperacao
                End If
                objOperacoesTempo.dTempoPreparacao = objOperacoes.objOperacoesTempo.dTempoPreparacao
                objOperacoesTempo.iTipo = objOperacoes.objOperacoesTempo.iTipo
                objOperacoesTempo.sUMTempo = objOperacoes.objOperacoesTempo.sUMTempo
                objOperacoesTempo.lNumIntDocMaq = objOperacoes.objOperacoesTempo.lNumIntDocMaq
                
                Set objOrdemProducaoOperacoes.objOperacoesTempo = objOperacoesTempo
            
            End If
                                            
            objItemOP.colOrdemProducaoOperacoes.Add objOrdemProducaoOperacoes
            
            For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
            
                Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos
                
                objOrdemProducaoInsumos.sProduto = objOperacaoInsumos.sProduto
                If objOperacaoInsumos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                    objOrdemProducaoInsumos.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQtdeOPInsumos
                Else
                    objOrdemProducaoInsumos.dQuantidade = objOperacaoInsumos.dQuantidade
                End If
                objOrdemProducaoInsumos.sUMProduto = objOperacaoInsumos.sUMProduto
                objOrdemProducaoInsumos.iComposicao = objOperacaoInsumos.iComposicao
                objOrdemProducaoInsumos.dPercentualPerda = objOperacaoInsumos.dPercentualPerda
                objOrdemProducaoInsumos.dCustoStandard = objOperacaoInsumos.dCustoStandard
                objOrdemProducaoInsumos.sVersaoKitComp = objOperacaoInsumos.sVersaoKitComp
            
                objOrdemProducaoOperacoes.colOPInsumos.Add objOrdemProducaoInsumos
                
                Set objRotFabr = New ClassRoteirosDeFabricacao
                
                objRotFabr.sProdutoRaiz = objOrdemProducaoInsumos.sProduto
                objRotFabr.sVersao = objOrdemProducaoInsumos.sVersaoKitComp
                
                lErro = OrdemDeProducao_Gera_Operacoes1(objRotFabr, objOrdemProducaoInsumos, objItemOP, iNivel + objOperacoes.iNivel + 1, iNivelRoteiro + 1, iSeqRoteiro, iSeq)
                If lErro <> SUCESSO Then gError 137191
            
            Next
            
        Next
        
    Else
    
        'Se não tem roteiro ... traz do Kit
        lErro = OrdemDeProducao_Gera_Operacoes3(objRoteirosDeFabricacao, objOPInsumos, objItemOP, iNivel, iNivelRoteiro, iSeqRoteiro, iSeq)
        If lErro <> SUCESSO Then gError 137710
            
    End If
    
    OrdemDeProducao_Gera_Operacoes1 = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Gera_Operacoes1:

    OrdemDeProducao_Gera_Operacoes1 = gErr
    
    Select Case gErr
    
        Case 137186, 137187, 137189 To 137191, 137710, 139044
            'erros tratados nas rotinas chamadas
            
        Case 137188
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objRoteirosDeFabricacao.sProdutoRaiz)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152293)
            
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Gera_Operacoes2(ByVal objItemOP As ClassItemOP) As Long
'Ajusta as sequencias...

Dim lErro As Long
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim iNivelAnterior As Integer
Dim aiSeqArvore() As Integer
Dim aiSeqPai() As Integer
Dim aiSeqRoteiroPai() As Integer
Dim iNivelRoteiroAnterior As Integer
Dim iOcorrencias As Integer
Dim iIndice As Integer

On Error GoTo Erro_OrdemDeProducao_Gera_Operacoes2

    'Número máximo de ocorrencias na coleção
    iOcorrencias = objItemOP.colOrdemProducaoOperacoes.Count

    'Redimensionar os arrays
    ReDim aiSeqArvore(iOcorrencias)
    ReDim aiSeqPai(iOcorrencias)
    ReDim aiSeqRoteiroPai(iOcorrencias)

    'Para cada Operação encontrada
    For Each objOrdemProducaoOperacoes In objItemOP.colOrdemProducaoOperacoes
    
        'Se o Nivel for zero ...
        If objOrdemProducaoOperacoes.iNivel = 0 Then
        
            aiSeqPai(objOrdemProducaoOperacoes.iNivel) = 0
        
        Else
        
            'se o nivel for diferente ...
            If iNivelAnterior <> objOrdemProducaoOperacoes.iNivel Then
            
                'se esta retornando a um nível mais anterior ...
                If iNivelAnterior > objOrdemProducaoOperacoes.iNivel Then
                
                    'zerar os próximos pais
                    For iIndice = objOrdemProducaoOperacoes.iNivel + 1 To iOcorrencias
                    
                        aiSeqPai(iIndice) = 0
                    
                    Next iIndice
                
                End If

                'se ainda não tem valor para o pai ...
                If aiSeqPai(objOrdemProducaoOperacoes.iNivel) = 0 Then
                
                    'atribuir ao valor do pai o do sequencial da árvore anterior
                    aiSeqPai(objOrdemProducaoOperacoes.iNivel) = objOrdemProducaoOperacoes.iSeq - 1
                
                End If
                
                'atualizar o nível
                iNivelAnterior = objOrdemProducaoOperacoes.iNivel
                                
            End If
        
        End If
        
        'Se o NivelRoteiro for zero ...
        If objOrdemProducaoOperacoes.iNivelRoteiro = 0 Then
        
            aiSeqRoteiroPai(objOrdemProducaoOperacoes.iNivelRoteiro) = 0
        
        Else
        
            'se o nivelRoteiro for diferente ...
            If iNivelRoteiroAnterior <> objOrdemProducaoOperacoes.iNivelRoteiro Then
            
                'se esta retornando a um nívelRoteiro mais anterior ...
                If iNivelRoteiroAnterior > objOrdemProducaoOperacoes.iNivelRoteiro Then
                
                    'zerar os próximos pais
                    For iIndice = objOrdemProducaoOperacoes.iNivelRoteiro + 1 To iOcorrencias
                    
                        aiSeqRoteiroPai(iIndice) = 0
                    
                    Next iIndice
                
                End If

                'se ainda não tem valor para o pai ...
                If aiSeqRoteiroPai(objOrdemProducaoOperacoes.iNivelRoteiro) = 0 Then
                
                    'atribuir ao valor do pai o do sequencial da árvore anterior
                    aiSeqRoteiroPai(objOrdemProducaoOperacoes.iNivelRoteiro) = objOrdemProducaoOperacoes.iSeqRoteiro - 1
                
                End If
                
                'atualizar o nível
                iNivelRoteiroAnterior = objOrdemProducaoOperacoes.iNivelRoteiro
                                
            End If
        
        End If

        'incrementar o sequencial do nível
        aiSeqArvore(objOrdemProducaoOperacoes.iNivel) = aiSeqArvore(objOrdemProducaoOperacoes.iNivel) + 1

        'ajustar as sequencias na coleção
        objOrdemProducaoOperacoes.iSeqArvore = aiSeqArvore(objOrdemProducaoOperacoes.iNivel)
        objOrdemProducaoOperacoes.iSeqPai = aiSeqPai(objOrdemProducaoOperacoes.iNivel)
        objOrdemProducaoOperacoes.iSeqRoteiroPai = aiSeqRoteiroPai(objOrdemProducaoOperacoes.iNivelRoteiro)
    
    Next

    OrdemDeProducao_Gera_Operacoes2 = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Gera_Operacoes2:

    OrdemDeProducao_Gera_Operacoes2 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152294)
    
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Gera_Operacoes3(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao, ByVal objOPInsumos As ClassOrdemProducaoInsumos, ByVal objItemOP As ClassItemOP, ByVal iNivel As Integer, ByVal iNivelRoteiro As Integer, iSeqRoteiro As Integer, iSeq As Integer)
    
Dim lErro As Long
Dim objProduto As ClassProduto
Dim objKit As ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim dFatorUMConv As Double
Dim dFatorQtdeOPInsumos As Double
Dim dQuantidade As Double
Dim objRotFabr As ClassRoteirosDeFabricacao
Dim objCompetenciaPadrao As New ClassCompetencias
Dim objProdutoKitAux As New ClassProdutoKit
Dim objKitAux As New ClassKit
Dim bExisteProduto As Boolean
Dim dFator As Double

On Error GoTo Erro_OrdemDeProducao_Gera_Operacoes3

    'Se a quantidade já chegou a 0.0001 ou já desceu na recursividade n instâncias
    If (objOPInsumos.dQuantidade <= 0.0001 And iNivel > 0) Or iNivel >= 20 Then
    
        OrdemDeProducao_Gera_Operacoes3 = SUCESSO
        Exit Function

    End If
        
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objRoteirosDeFabricacao.sProdutoRaiz
    
    'Le o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137882

    If lErro <> SUCESSO Then gError 137883
    
    'Se for Produto Produzivel
    If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
    
        'Le a Competencia Padrao
        lErro = CF("Competencias_Le_Padrao", objCompetenciaPadrao)
        If lErro <> SUCESSO And lErro <> 137844 Then gError 137849
        
        If lErro <> SUCESSO Then gError 137850
    
        Set objKit = New ClassKit
        
        objKit.sProdutoRaiz = objRoteirosDeFabricacao.sProdutoRaiz
        
        If Len(Trim(objRoteirosDeFabricacao.sVersao)) <> 0 Then
        
            objKit.sVersao = objRoteirosDeFabricacao.sVersao
        
        Else
        
            'Le as Versoes Ativas e a Padrao
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 137884
            
            If lErro <> SUCESSO Then gError 137885
        
        End If
        
        lErro = CF("Kit_Le_Componentes", objKit)
        If lErro <> SUCESSO And lErro <> 21831 Then gError 137115
    
        If lErro <> SUCESSO Then gError 137886
    
        '#############################################################################
        'Inserido por Wagner 08/08/2006
        For Each objProdutoKit In objKit.colComponentes
            
            bExisteProduto = False
            For Each objProdutoKitAux In objKitAux.colComponentes
        
                If objProdutoKitAux.sProduto = objProdutoKit.sProduto And objProdutoKitAux.iNivel <> KIT_NIVEL_RAIZ Then
                    If objProdutoKitAux.iComposicao <> objProdutoKit.iComposicao Or _
                        objProdutoKitAux.dCustoStandard <> objProdutoKit.dCustoStandard Or _
                        objProdutoKitAux.dPercentualPerda <> objProdutoKit.dPercentualPerda Or _
                        objProdutoKitAux.sVersaoKitComp <> objProdutoKit.sVersaoKitComp Then gError 181740
                    bExisteProduto = True
                    Exit For
                End If
        
            Next
            
            If Not bExisteProduto Then
            
                Set objProdutoKitAux = New ClassProdutoKit
                
                objProdutoKitAux.dCustoStandard = objProdutoKit.dCustoStandard
                objProdutoKitAux.dPercentualPerda = objProdutoKit.dPercentualPerda
                objProdutoKitAux.dQuantidade = objProdutoKit.dQuantidade
                objProdutoKitAux.iClasseUM = objProdutoKit.iClasseUM
                objProdutoKitAux.iComposicao = objProdutoKit.iComposicao
                objProdutoKitAux.iNivel = objProdutoKit.iNivel
                objProdutoKitAux.iPosicaoArvore = objProdutoKit.iPosicaoArvore
                objProdutoKitAux.iSeq = objProdutoKit.iSeq
                objProdutoKitAux.iSeqPai = objProdutoKit.iSeqPai
                objProdutoKitAux.sProduto = objProdutoKit.sProduto
                objProdutoKitAux.sProdutoDesc = objProdutoKit.sProdutoDesc
                objProdutoKitAux.sProdutoRaiz = objProdutoKit.sProdutoRaiz
                objProdutoKitAux.sUnidadeMed = objProdutoKit.sUnidadeMed
                objProdutoKitAux.sVersao = objProdutoKit.sVersao
                objProdutoKitAux.sVersaoKitComp = objProdutoKit.sVersaoKitComp
    
                objKitAux.colComponentes.Add objProdutoKitAux
            
            Else
            
                Set objProduto = New ClassProduto
            
                objProduto.sCodigo = objProdutoKit.sProduto
                
                'Lê o Produto Componente do Kit
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 181741
                
                'se o produto não estiver cadastrado... erro
                If lErro <> SUCESSO Then gError 181743
                   
                'Descobre o fator de conversao da UM do Roteiro p/UM do Kit do Produto Raiz
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProdutoKitAux.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 181742
                
                objProdutoKitAux.dQuantidade = objProdutoKitAux.dQuantidade + (objProdutoKit.dQuantidade * dFator)
       
            End If
            
        Next
        '#############################################################################
    
        dFatorQtdeOPInsumos = 1

        iSeqRoteiro = iSeqRoteiro + 1
    
        'Para cada ProdutoKit na coleção de componentes
        For Each objProdutoKit In objKitAux.colComponentes
        
            Set objProduto = New ClassProduto
        
            objProduto.sCodigo = objProdutoKit.sProduto
            
            'Lê o Produto Componente do Kit
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 139050
            
            If lErro <> SUCESSO Then gError 139051
            
            objProdutoKit.iClasseUM = objProduto.iClasseUM
        
            'Se for o ProdutoRaiz do Kit -> é uma Operação
            If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                       
                'Descobre o fator de conversao da UM do Item da OP p/UM do Kit do produto
                lErro = CF("UM_Conversao_Trans", objProdutoKit.iClasseUM, objOPInsumos.sUMProduto, objProdutoKit.sUnidadeMed, dFatorUMConv)
                If lErro <> SUCESSO Then gError 137118
                        
                dFatorQtdeOPInsumos = (objOPInsumos.dQuantidade * dFatorUMConv) / objProdutoKit.dQuantidade

                iSeq = iSeq + 1
    
                Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
    
                objOrdemProducaoOperacoes.lNumIntDocCompet = objCompetenciaPadrao.lNumIntDoc
                objOrdemProducaoOperacoes.lNumIntDocCT = objCompetenciaPadrao.lNumIntDocCT
    
                objOrdemProducaoOperacoes.iNivel = iNivel
                objOrdemProducaoOperacoes.iSeq = iSeq
                objOrdemProducaoOperacoes.iNivelRoteiro = iNivelRoteiro
                objOrdemProducaoOperacoes.iSeqRoteiro = iSeqRoteiro
                objOrdemProducaoOperacoes.sProduto = objRoteirosDeFabricacao.sProdutoRaiz
                objOrdemProducaoOperacoes.sVersao = objRoteirosDeFabricacao.sVersao
                objOrdemProducaoOperacoes.dQuantidade = objOPInsumos.dQuantidade
                objOrdemProducaoOperacoes.sUMedida = objOPInsumos.sUMProduto
                objOrdemProducaoOperacoes.iConsideraCarga = DESMARCADO
                objOrdemProducaoOperacoes.iOrigem = ORIGEM_KIT
                
                objItemOP.colOrdemProducaoOperacoes.Add objOrdemProducaoOperacoes
                
            Else
                
                'Senão são insumos desta operação
                Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos

                objOrdemProducaoInsumos.sProduto = objProdutoKit.sProduto
                If objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                    objOrdemProducaoInsumos.dQuantidade = objProdutoKit.dQuantidade * dFatorQtdeOPInsumos
                Else
                    objOrdemProducaoInsumos.dQuantidade = objProdutoKit.dQuantidade
                End If
                objOrdemProducaoInsumos.sUMProduto = objProdutoKit.sUnidadeMed
                objOrdemProducaoInsumos.iComposicao = objProdutoKit.iComposicao
                objOrdemProducaoInsumos.dPercentualPerda = objProdutoKit.dPercentualPerda
                objOrdemProducaoInsumos.dCustoStandard = objProdutoKit.dCustoStandard
                objOrdemProducaoInsumos.sVersaoKitComp = objProdutoKit.sVersaoKitComp

'                Set objProduto = New ClassProduto
'
'                objProduto.sCodigo = objOrdemProducaoInsumos.sProduto
'
'                'Le o produto
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 139050
'
'                If lErro <> SUCESSO Then gError 139051
                
                'Se for Produto Produzivel
                If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                
                    If Len(Trim(objOrdemProducaoInsumos.sVersaoKitComp)) = 0 Then
                    
                        Set objKit = New ClassKit
                    
                        objKit.sProdutoRaiz = objOrdemProducaoInsumos.sProduto
                    
                        'Le as Versoes Ativas e a Padrao
                        lErro = CF("Kit_Le_Padrao", objKit)
                        If lErro <> SUCESSO And lErro <> 106304 Then gError 139052
                        
                        If lErro <> SUCESSO Then gError 139053
                        
                        objOrdemProducaoInsumos.sVersaoKitComp = objKit.sVersao
                    
                    End If
                
                End If

                objOrdemProducaoOperacoes.colOPInsumos.Add objOrdemProducaoInsumos

                'Se for Produto Produzivel
                If objProduto.iCompras = PRODUTO_PRODUZIVEL Then

                    Set objRotFabr = New ClassRoteirosDeFabricacao
    
                    objRotFabr.sProdutoRaiz = objOrdemProducaoInsumos.sProduto
                    objRotFabr.sVersao = objOrdemProducaoInsumos.sVersaoKitComp
                    
                    'Chama novamente o Default pois pode ter Roteiro
                    lErro = OrdemDeProducao_Gera_Operacoes1(objRotFabr, objOrdemProducaoInsumos, objItemOP, iNivel + 1, iNivelRoteiro + 1, iSeqRoteiro, iSeq)
                    If lErro <> SUCESSO Then gError 137120

                End If

            End If
            
        Next

    End If
    
    OrdemDeProducao_Gera_Operacoes3 = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Gera_Operacoes3:

    OrdemDeProducao_Gera_Operacoes3 = gErr
    
    Select Case gErr
    
        Case 137115, 137116, 137118 To 137120, 137849, 137882, 137884, 139050, 139052, 181741, 181742
            'erros tratados nas rotinas chamadas
            
        Case 137850
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_PADRAO_NAO_CADASTRADA", gErr)
            
        Case 137883, 139051, 181743
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 137885, 139053
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz, objKit.sVersao)
        
        Case 137886
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOKIT_INEXISTENTE", gErr, objKit.sVersao, objKit.sProdutoRaiz)
            
        Case 181740
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOKIT_DIFERENTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152295)
            
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Grava_Operacoes_EmTrans(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objItemOP As New ClassItemOP
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim objOPOperacoesAApagar As New ClassOrdemProducaoOperacoes
Dim lNumIntDoc As Long
Dim bApagarOperacao As Boolean

On Error GoTo Erro_OrdemDeProducao_Grava_Operacoes_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134979
    Next
    
    'Inclui as Operações dos itens da OP
    For Each objItemOP In objOrdemDeProducao.colItens
    
        For Each objOrdemProducaoOperacoes In objItemOP.colOrdemProducaoOperacoes
        
            'Le a tabela OrdemProducaoOperacoes
            If objOrdemProducaoOperacoes.lNumIntDocItemOP = 0 Then
                lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM OrdemProducaoOperacoes WHERE NumIntDocItemOP= ? AND Seq= ? ", _
                            0, lNumIntDoc, objItemOP.lNumIntDoc, objOrdemProducaoOperacoes.iSeq)
            Else
                lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM OrdemProducaoOperacoes WHERE NumIntDocItemOP= ? AND Seq= ? ", _
                            0, lNumIntDoc, objOrdemProducaoOperacoes.lNumIntDocItemOP, objOrdemProducaoOperacoes.iSeq)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 134980
                   
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134981
        
            'Se existir => UPDATE, senão => INSERT
            If lErro = AD_SQL_SUCESSO Then
            
                objOrdemProducaoOperacoes.lNumIntDoc = lNumIntDoc
                If objOrdemProducaoOperacoes.lNumIntDocItemOP = 0 Then objOrdemProducaoOperacoes.lNumIntDocItemOP = objItemOP.lNumIntDoc
            
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(0))
                If lErro <> SUCESSO Then gError 134982
                
                With objOrdemProducaoOperacoes
                
                    lErro = Comando_ExecutarPos(alComando(1), "UPDATE OrdemProducaoOperacoes SET Seq= ? , NumIntDocCompet= ? , NumIntDocCT= ? , Observacao= ? , " & _
                            "SeqPai= ? , SeqArvore= ? , Nivel= ? , SeqRoteiroPai = ? , NivelRoteiro = ? , Produto = ? , Versao = ? , IgnoraTaxaProducao = ? , " & _
                            "ConsideraCarga = ? , Origem = ? , NumIntDocOperOrigem = ?, NumMaxMaqPorOper = ?, NumRepeticoes = ? ", alComando(0), _
                            .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                            .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiroPai, .iNivelRoteiro, _
                            .sProduto, .sVersao, .iIgnoraTaxaProducao, .iConsideraCarga, _
                            .iOrigem, .lNumIntDocOperOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes)
                    If lErro <> AD_SQL_SUCESSO Then gError 134983
                
                End With
                
                lErro = OrdemDeProducao_Exclui_Insumos_EmTrans(objOrdemProducaoOperacoes)
                If lErro <> SUCESSO Then gError 134984
                            
                lErro = OrdemDeProducao_Exclui_Tempo_EmTrans(objOrdemProducaoOperacoes)
                If lErro <> SUCESSO Then gError 137047
                
            Else
            
                'Obter NumIntDoc
                lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_ORDEMPRODUCAO_OPERACOES", lNumIntDoc)
                If lErro <> SUCESSO Then gError 134985
        
                objOrdemProducaoOperacoes.lNumIntDoc = lNumIntDoc
                objOrdemProducaoOperacoes.lNumIntDocItemOP = objItemOP.lNumIntDoc
                
                With objOrdemProducaoOperacoes
                
                    lErro = Comando_Executar(alComando(1), "INSERT INTO OrdemProducaoOperacoes( NumIntDoc, NumIntDocItemOP, Seq, NumIntDocCompet, NumIntDocCT, Observacao, " & _
                            "SeqPai, SeqArvore, Nivel, SeqRoteiro, SeqRoteiroPai, NivelRoteiro, Produto, Versao, IgnoraTaxaProducao, ConsideraCarga, Origem, NumIntDocOperOrigem, NumMaxMaqPorOper, NumRepeticoes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                            .lNumIntDoc, .lNumIntDocItemOP, .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                            .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, .iSeqRoteiroPai, .iNivelRoteiro, .sProduto, .sVersao, .iIgnoraTaxaProducao, .iConsideraCarga, .iOrigem, .lNumIntDocOperOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134986
                
                End With
                
            End If
            
            lErro = OrdemDeProducao_Grava_Insumos_EmTrans(objOrdemProducaoOperacoes)
            If lErro <> SUCESSO Then gError 134987
            
            '##############################
            'Inserido por Wagner 07/11/2005
            lErro = ItemOPOperacoesMaquinas_Exclui(objOrdemProducaoOperacoes)
            If lErro <> SUCESSO Then gError 140772
            
            lErro = ItemOPOperacoesMaquinas_Insere(objOrdemProducaoOperacoes)
            If lErro <> SUCESSO Then gError 140773
            '##############################
        
            If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
            
                lErro = OrdemDeProducao_Grava_Tempo_EmTrans(objOrdemProducaoOperacoes)
                If lErro <> SUCESSO Then gError 137048
            
            End If
            
        Next
        
        'Le a tabela OrdemProducaoOperacoes novamente para verificar se haverá deleções
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM OrdemProducaoOperacoes WHERE NumIntDocItemOP= ? ", _
                    0, lNumIntDoc, objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 134988
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134989
    
        'Faça enquanto existir
        Do While lErro = AD_SQL_SUCESSO
        
            bApagarOperacao = True
            
            'Para cada OrdemProducaoOperacoes na coleção
            For Each objOrdemProducaoOperacoes In objItemOP.colOrdemProducaoOperacoes
            
                If objOrdemProducaoOperacoes.lNumIntDoc = lNumIntDoc Then
                
                     bApagarOperacao = False
                     Exit For
                
                End If
            
            Next
            
            If bApagarOperacao Then
                
                lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM OrdemProducaoOperacoes", alComando(2))
                If lErro <> AD_SQL_SUCESSO Then gError 134990
                
                Set objOPOperacoesAApagar = New ClassOrdemProducaoOperacoes
                
                objOPOperacoesAApagar.lNumIntDoc = lNumIntDoc
                lErro = OrdemDeProducao_Exclui_Insumos_EmTrans(objOPOperacoesAApagar)
                If lErro <> SUCESSO Then gError 134991
                                        
                lErro = ItemOPOperacoesMaquinas_Exclui(objOPOperacoesAApagar)
                If lErro <> SUCESSO Then gError 140772
                    
            End If
    
            'vai para o próximo
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134992
    
        Loop
        
    Next
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OrdemDeProducao_Grava_Operacoes_EmTrans = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Grava_Operacoes_EmTrans:

    OrdemDeProducao_Grava_Operacoes_EmTrans = gErr
    
    Select Case gErr
        
        Case 134979
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134980, 134981, 134988, 134989, 134992
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OP_OPERACOES", gErr)

        Case 134982
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OP_OPERACOES", gErr)

        Case 134983
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_OP_OPERACOES", gErr)
            
        Case 134984, 134985, 134987, 134991, 137047, 137048, 140772, 140773
            'erros tratados nas rotinas chamadas

        Case 134986
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OP_OPERACOES", gErr)
            
        Case 134990
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OP_OPERACOES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152296)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function OrdemDeProducao_Exclui_Operacoes_EmTrans(ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim lNumIntDocOper As Long

On Error GoTo Erro_OrdemDeProducao_Exclui_Operacoes_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comandos
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134993
    Next
    
    'Le a tabela OrdemProducaoOperacoes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM OrdemProducaoOperacoes WHERE NumIntDocItemOP= ? ", _
                0, lNumIntDocOper, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134994

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134995

    'Faz enquanto existir
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
        
        objOrdemProducaoOperacoes.lNumIntDoc = lNumIntDocOper

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 134996
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM OrdemProducaoOperacoes", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 134997
        
        lErro = OrdemDeProducao_Exclui_Insumos_EmTrans(objOrdemProducaoOperacoes)
        If lErro <> SUCESSO Then gError 134998
    
        lErro = OrdemDeProducao_Exclui_Tempo_EmTrans(objOrdemProducaoOperacoes)
        If lErro <> SUCESSO Then gError 137049
        
        '##############################
        'Inserido por Wagner 07/11/2005
        lErro = ItemOPOperacoesMaquinas_Exclui(objOrdemProducaoOperacoes)
        If lErro <> SUCESSO Then gError 140738
        '##############################
        
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134999

    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    OrdemDeProducao_Exclui_Operacoes_EmTrans = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Exclui_Operacoes_EmTrans:

    OrdemDeProducao_Exclui_Operacoes_EmTrans = gErr
    
    Select Case gErr
        
        Case 134993
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134994, 134995, 134999
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OP_OPERACOES", gErr)

        Case 134996
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_OP_OPERACOES", gErr)

        Case 134997
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OP_OPERACOES", gErr)
            
        Case 134998, 137049, 140738
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152297)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function OrdemDeProducao_Exclui_Insumos_EmTrans(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_OrdemDeProducao_Exclui_Insumos_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137000
    Next
    
    'Le a OperacoesTempo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM OrdemProducaoInsumos WHERE NumIntDocOper= ? ", _
                0, iAux, objOrdemProducaoOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137001

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137002
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 137003
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM OrdemProducaoInsumos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 137004

        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137005

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Exclui_Insumos_EmTrans = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui_Insumos_EmTrans:

    OrdemDeProducao_Exclui_Insumos_EmTrans = gErr

    Select Case gErr

        Case 137000
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137001, 137002, 137005
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOINSUMOS", gErr)

        Case 137003
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ORDEMPRODUCAOINSUMOS", gErr)

        Case 137004
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORDEMPRODUCAOINSUMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152298)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function OrdemDeProducao_Grava_Insumos_EmTrans(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOrdemProducaoInsumos As New ClassOrdemProducaoInsumos

On Error GoTo Erro_OrdemDeProducao_Grava_Insumos_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137006
        
    'Para cada objOrdemProducaoInsumos na coleção
    For Each objOrdemProducaoInsumos In objOrdemProducaoOperacoes.colOPInsumos
        
        objOrdemProducaoInsumos.lNumIntDocOper = objOrdemProducaoOperacoes.lNumIntDoc

        With objOrdemProducaoInsumos
            
            lErro = Comando_Executar(lComando, "INSERT INTO OrdemProducaoInsumos (NumIntDocOper, Produto, Quantidade, UMProduto, VersaoKitComp, Composicao, PercentualPerda, CustoStandard) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", _
                                                .lNumIntDocOper, .sProduto, .dQuantidade, .sUMProduto, .sVersaoKitComp, .iComposicao, .dPercentualPerda, .dCustoStandard)
            If lErro <> AD_SQL_SUCESSO Then gError 137007
            
        End With
                           
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    OrdemDeProducao_Grava_Insumos_EmTrans = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava_Insumos_EmTrans:

    OrdemDeProducao_Grava_Insumos_EmTrans = gErr

    Select Case gErr

        Case 137006
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137007
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDEMPRODUCAOINSUMOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152299)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function OrdemDeProducao_Grava_Tempo_EmTrans(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOperacoesTempo As New ClassOperacoesTempo

On Error GoTo Erro_OrdemDeProducao_Grava_Tempo_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137050
        
    Set objOperacoesTempo = objOrdemProducaoOperacoes.objOperacoesTempo
        
    objOperacoesTempo.lNumIntDocOperacao = objOrdemProducaoOperacoes.lNumIntDoc

    lErro = Comando_Executar(lComando, "INSERT INTO OrdemProducaoTempo (NumIntDocOperacao, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, TempoOperacao, UMTempo, Tipo, NumIntDocMaq) VALUES (?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
                                        objOperacoesTempo.lNumIntDocOperacao, objOperacoesTempo.dLoteMax, objOperacoesTempo.dLoteMin, objOperacoesTempo.dLotePadrao, objOperacoesTempo.dTempoPreparacao, objOperacoesTempo.dTempoMovimentacao, objOperacoesTempo.dTempoDescarga, objOperacoesTempo.dTempoOperacao, objOperacoesTempo.sUMTempo, objOperacoesTempo.iTipo, objOperacoesTempo.lNumIntDocMaq)
    If lErro <> AD_SQL_SUCESSO Then gError 137051
                           
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    OrdemDeProducao_Grava_Tempo_EmTrans = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Grava_Tempo_EmTrans:

    OrdemDeProducao_Grava_Tempo_EmTrans = gErr

    Select Case gErr

        Case 137050
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 137051
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORDEMPRODUCAOTEMPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152300)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrdemDeProducao_Exclui_Tempo_EmTrans(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_OrdemDeProducao_Exclui_Tempo_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137052
    Next
    
    'Le a OperacoesTempo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM OrdemProducaoTempo WHERE NumIntDocOperacao= ? ", _
                0, iAux, objOrdemProducaoOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137053

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137054
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 137055
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM OrdemProducaoTempo", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 137056
            
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Exclui_Tempo_EmTrans = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Exclui_Tempo_EmTrans:

    OrdemDeProducao_Exclui_Tempo_EmTrans = gErr

    Select Case gErr

        Case 137052
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137053, 137054
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOTEMPO", gErr)

        Case 137055
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ORDEMPRODUCAOTEMPO", gErr)

        Case 137056
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORDEMPRODUCAOTEMPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152301)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'#########################################################
'MRP
'Inserido por Wagner
Public Function MRP_Gera(ByVal objPMP As ClassPMP) As Long
'Abre a Transação e Chama a Gera_EmTrans

Dim lErro As Long
Dim lTransacao As Long
Dim objPMPItens As ClassPMPItens
Dim vbResult As VbMsgBoxResult
Dim iCont As Integer
Dim bTemStatusNOK As Boolean

On Error GoTo Erro_MRP_Gera

    bTemStatusNOK = False

    'Verifica se todos os itens são válidos
    For Each objPMPItens In objPMP.colItens

        If objPMPItens.iStatus <> PO_STATUS_OK Then
            bTemStatusNOK = True
        Else
            iCont = iCont + 1
        End If

    Next

    If bTemStatusNOK Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_PMPITENS_NAO_OK")
        If vbResult = vbNo Then gError 136548
    End If

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 132450

    'Chama a Rotina que Gera dentro de transação
    lErro = MRP_Gera_EmTrans(objPMP)
    If lErro <> SUCESSO Then gError 132451

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 132452

    'Avisa sobre a conclusão da geração do Plano Mestre de Produção
    If objPMP.lCodGeracao <> 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_PMP_GERADA", objPMP.lCodGeracao)
    Else
        Call Rotina_Aviso(vbOKOnly, "AVISO_PMP_SEM_DADOS")
    End If

    MRP_Gera = SUCESSO

    Exit Function

Erro_MRP_Gera:

    MRP_Gera = gErr

    Select Case gErr

        Case 132450
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 132451 'Tratado na Rotina Chamada

        Case 136548

        Case 132452
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152302)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PMP_Acerta_OP_RC(ByVal objPMP As ClassPMP) As Long
'Acerta as datas e códigos das OPs e datas das requisições de compra
'com base nas informações do plano operacional
 
Dim lErro As Long
Dim objPMPItem As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objPOFilha As ClassPlanoOperacional
Dim objOP As ClassOrdemDeProducao
Dim objRC As ClassRequisicaoCompras
Dim objItemOP As ClassItemOP
 
On Error GoTo Erro_PMP_Acerta_OP_RC
 
    'Para cada item do Plano Mestre
    For Each objPMPItem In objPMP.colItens
    
        'Para cada etapa
        For Each objPO In objPMPItem.colPO
        
            'Para cada Requisição de compra da etapa
            For Each objRC In objPO.colRCFilhas
            
                'Acerta a data
                objRC.dtDataLimite = objPO.dtDataInicio
                objRC.sOPCodigo = objPO.sCodOPOrigem
            
            Next
        
            'Para cada SubOrdem de produção da etapa
            For Each objOP In objPO.colOPFilhas
                       
                objOP.sOPGeradora = objPO.sCodOPOrigem
                       
                'Para cada item da OP
                For Each objItemOP In objOP.colItens
                
                    'Para cada outra etapa
                    For Each objPOFilha In objPMPItem.colPO
                    
                        'Se for filha
                        If objPOFilha.lNumIntDocPOPai = objPO.lNumIntDoc Then
                        
                            'Se for a etapa que vai gerar esse item de OP
                            If objItemOP.sProduto = objPOFilha.sProduto Then
                            
                                'Pega o código que pode ter sido mudado
                                objOP.sCodigo = objPOFilha.sCodOPOrigem
            
                                'Acerta a data
                                objItemOP.dtDataFimProd = objPOFilha.dtDataFim
                                objItemOP.dtDataInicioProd = objPOFilha.dtDataInicio
            
                                Exit For
                                
                            End If
                        
                        End If
                    
                    Next
                
                    objItemOP.sCodigo = objOP.sCodigo
                                        
                Next
            
            Next
        
        Next
    
    Next
   
    PMP_Acerta_OP_RC = SUCESSO
 
    Exit Function
 
Erro_PMP_Acerta_OP_RC:
 
    PMP_Acerta_OP_RC = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152303)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Gera_EmTrans(ByVal objPMP As ClassPMP) As Long
'Gera o Plano Mestre, Suas Etapas, As SubOPs, as Requisições de compra e as
'alocações de máquinas
 
Dim lErro As Long
Dim objOP As ClassOrdemDeProducao
Dim objOPAux As ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim objItemOPPai As ClassItemOP
Dim objRC As ClassRequisicaoCompras
Dim objPMPItem As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim lCodigo As Long
Dim objOPOperacao As ClassOrdemProducaoOperacoes
 
On Error GoTo Erro_MRP_Gera_EmTrans
        
    'Acerta as informações que estão redudantes relativas as OPs e RCs
    lErro = PMP_Acerta_OP_RC(objPMP)
    If lErro <> SUCESSO Then gError 136318
    
    'Para cada item do Plano Mestre
    For Each objPMPItem In objPMP.colItens
    
        'Se a OP é planejada Grava
        If objPMPItem.objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA Then
        
            Set objOPAux = New ClassOrdemDeProducao
           
            'objPMPItem.objOP.colItens.Item(1).dtDataInicioProd = objPMPItem.colPO.Item(1).dtDataInicio
            'objPMPItem.objItemOP.dtDataFimProd = objPMPItem.colPO.Item(1).dtDataFim
            
            objPMPItem.objOP.colItens.Item(1).lNumIntDoc = 0
            
            objOPAux.sCodigo = objPMPItem.objOP.sCodigo
            objOPAux.iFilialEmpresa = objPMPItem.objOP.iFilialEmpresa

            lErro = CF("OrdemProducao_Le", objOPAux)
            If lErro <> SUCESSO And lErro <> 30368 Then gError 138210
            If lErro = SUCESSO Then gError 138211
            
            'Grava a OP
            lErro = CF("OrdemDeProducao_Grava_EmTrans", objPMPItem.objOP)
            If lErro <> SUCESSO Then gError 136429
            
            Set objPMPItem.objItemOP = objPMPItem.objOP.colItens.Item(1)
            Set objPMPItem.colPO.Item(1).objOP = objPMPItem.objOP
            
            'Acerta o NumIntDocOper das Etapas
            For Each objPO In objPMPItem.colPO
            
                For Each objOPOperacao In objPMPItem.objItemOP.colOrdemProducaoOperacoes
            
                    If objPO.iNivel = objOPOperacao.iNivel And objPO.iSeq = objOPOperacao.iSeqArvore And objOPOperacao.sProduto = objOPOperacao.sProduto Then
                        objPO.lNumIntDocOper = objOPOperacao.lNumIntDoc
                        Exit For
                    End If
            
                Next
            
            Next
                
        End If
    
        'Para cada etapa
        For Each objPO In objPMPItem.colPO
        
            'Para cada Requisição de compra da etapa
            For Each objRC In objPO.colRCFilhas
                
                'Chama Config_ObterAutomatico
                lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROXIMO_CODIGO_RC", "RequisicaoCompra", "Codigo", lCodigo)
                If lErro <> SUCESSO Then gError 136319
                
                objRC.lCodigo = lCodigo
                
                'Grava a nova ReqCompra gerada
                lErro = CF("RequisicaoCompra_Grava_EmTrans", objRC, True)
                If lErro <> SUCESSO Then gError 136320
            
            Next
        
            'Para cada Ordem de Produção da etapa
            For Each objOP In objPO.colOPFilhas
            
                'Acerta o NumIntDocPai
                For Each objItemOP In objOP.colItens
                    Exit For
                Next
                
                For Each objItemOPPai In objPO.objOP.colItens
                    If objItemOPPai.sProduto = objPO.sProduto Then
                        objItemOP.lNumIntDocPai = objItemOPPai.lNumIntDoc
                        Exit For
                    End If
                Next
                
                Set objOPAux = New ClassOrdemDeProducao
                                
                objOPAux.sCodigo = objOP.sCodigo
                objOPAux.iFilialEmpresa = objOP.iFilialEmpresa
    
                lErro = CF("OrdemProducao_Le", objOPAux)
                If lErro <> SUCESSO And lErro <> 30368 Then gError 138210
                If lErro = SUCESSO Then gError 138211
                
'                'gera as operações da OP filha
'                lErro = Gera_Operacoes_ItemOP(objPO, objPMPItem.objItemOP, objItemOP)
'                If lErro <> SUCESSO Then gError 137920
                                    
                'Grava a OP
                lErro = CF("OrdemDeProducao_Grava_EmTrans", objOP)
                If lErro <> SUCESSO Then gError 136321
                
            Next
        Next
    Next
    
    'Grava o Plano Mestre, os itens do Plano mestre e as Etapas (Plano Operacional)
    lErro = PlanoMestreProducao_Grava_EmTrans(objPMP)
    If lErro <> SUCESSO Then gError 136317
    
    MRP_Gera_EmTrans = SUCESSO
    
    Exit Function
 
Erro_MRP_Gera_EmTrans:
 
    MRP_Gera_EmTrans = gErr
 
    Select Case gErr
        
        Case 136317 To 136321, 138210, 137920, 136429
        
        Case 138211
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOP.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152304)
 
    End Select
 
    Exit Function
 
End Function

Public Function PlanoMestreProducao_Grava_EmTrans(ByVal objPMP As ClassPMP) As Long
'Grava o Plano Mestre de Produção
 
Dim lErro As Long
Dim lErro2 As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPMPItens As ClassPMPItens
Dim lCodGeracao As Long
 
On Error GoTo Erro_PlanoMestreProducao_Grava_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 132486
    Next
 
    'Verifica se já existe um Plano Mestre de produção
    lErro = Comando_Executar(alComando(1), "SELECT CodGeracao FROM PlanoMestreProducaoUltimo ", lCodGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 132487
    
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132488
    
    'Se era uma simulação, apaga
    If Len(Trim(objPMP.sVersao)) > 0 Then
        lErro2 = PMP_Exclui_EmTrans(objPMP)
        If lErro2 <> SUCESSO Then gError 136457
    End If
    
    'Se é um novo Plano Mestre de produção
    If lErro <> AD_SQL_SUCESSO Then
 
        'Obter o Código da geração novo
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_CODIGO_PMP", lCodGeracao)
        If lErro <> SUCESSO Then gError 132489
    
        objPMP.dtDataGeracao = gdtDataAtual
    
        'Insere o Plano Mestre
        lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoMestreProducao(CodGeracao, DataGeracao) VALUES (?,?) ", lCodGeracao, objPMP.dtDataGeracao)
        If lErro <> AD_SQL_SUCESSO Then gError 132490
 
    End If
 
    'Acerta as datas/prioridades e CTs das OPs com base no que foi simulado
    lErro = PMP_Acerta_OP(objPMP)
    If lErro <> SUCESSO Then gError 136482
 
    objPMP.lCodGeracao = lCodGeracao
 
    'Para cada item novo do Plano Mestre
    For Each objPMPItens In objPMP.colItens
    
        'Se for um item sem críticas
'        If objPMPItens.iStatus = PO_STATUS_OK Then
    
            'Obter NumIntDoc
            lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_PMP_ITENS", lNumIntDoc)
            If lErro <> SUCESSO Then gError 132491
        
            objPMPItens.lNumIntDoc = lNumIntDoc
            objPMPItens.lCodGeracao = objPMP.lCodGeracao
        
            'Insere os Itens do Plano Mestre
            lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoMestreProducaoItens( NumIntDoc, Produto, Versao, Quantidade, " & _
                    "UM, DataNecessidade, CodOPOrigem, CodGeracao, Cliente, FilialCliente, FilialEmpresa, Prioridade) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                  objPMPItens.lNumIntDoc, objPMPItens.sProduto, objPMPItens.sVersao, objPMPItens.dQuantidade, _
                  objPMPItens.sUM, objPMPItens.dtDataNecessidade, objPMPItens.sCodOPOrigem, objPMPItens.lCodGeracao, objPMPItens.lCliente, objPMPItens.iFilialCli, objPMPItens.iFilialEmpresa, objPMPItens.iPrioridade)
            If lErro <> AD_SQL_SUCESSO Then gError 132492
            
            'Insere as etapas
            lErro = PlanoOperacional_Insere_Col(objPMPItens)
            If lErro <> SUCESSO Then gError 132493
    
'        End If
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PlanoMestreProducao_Grava_EmTrans = SUCESSO
 
    Exit Function
 
Erro_PlanoMestreProducao_Grava_EmTrans:
 
    PlanoMestreProducao_Grava_EmTrans = gErr
 
    Select Case gErr
 
        Case 132486
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 132487, 132488
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREDAPRODUCAO", gErr)
 
        Case 132489, 132491, 132493, 136457, 136482
 
        Case 132490
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOMESTREDAPRODUCAO", gErr)
 
        Case 132492
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOMESTREDAPRODUCAOITENS", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152305)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Public Function PMP_Acerta_OP(ByVal objPMP As ClassPMP) As Long
'Atualiza as datas de Incío e Fim de Produção, bem como a prioridade data pelo MRP
 
Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objPMPItens As ClassPMPItens
Dim lNumIntDoc As Long
Dim objPO As New ClassPlanoOperacional
 
On Error GoTo Erro_PMP_Acerta_OP
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136483
    Next
 
    For Each objPMPItens In objPMP.colItens
    
        'Le o Item da OP
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensOrdemProducao WHERE NumIntDoc = ? ", 0, lNumIntDoc, objPMPItens.objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 136484
        
        'Busca o primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136485
        
        If lErro = SUCESSO Then
        
            'Acerta Data/Prioridade/Produz Logo
            lErro = Comando_ExecutarPos(alComando(0), "UPDATE ItensOrdemProducao SET DataInicioProd = ?,DataFimProd = ?, Prioridade = ?, ProduzLogo = ? ", alComando(1), objPMPItens.objItemOP.dtDataInicioProd, objPMPItens.objItemOP.dtDataFimProd, objPMPItens.iPrioridade, objPMPItens.iProduzLogo)
            If lErro <> AD_SQL_SUCESSO Then gError 136486
        
            'Para cada Plano Operacional deste Item
            For Each objPO In objPMPItens.colPO
                
                'Le a Operação da OP
                lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM OrdemProducaoOperacoes WHERE NumIntDoc = ? ", 0, lNumIntDoc, objPO.lNumIntDocOper)
                If lErro <> AD_SQL_SUCESSO Then gError 137894

                'Busca o primeiro registro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137895

                If lErro = SUCESSO Then

                    'Acerta o Centro de Trabalho
                    lErro = Comando_ExecutarPos(alComando(3), "UPDATE OrdemProducaoOperacoes SET NumIntDocCT = ? ", alComando(2), objPO.lNumIntDocCT)
                    If lErro <> AD_SQL_SUCESSO Then gError 137896

                End If
        
            Next
        
        End If
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PMP_Acerta_OP = SUCESSO
 
    Exit Function
 
Erro_PMP_Acerta_OP:
 
    PMP_Acerta_OP = gErr
 
    Select Case gErr
 
        Case 136483
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136484, 136485
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO", gErr, objPMPItens.objItemOP.sCodigo, objPMPItens.objItemOP.iFilialEmpresa, objPMPItens.objItemOP.iItem)
        
        Case 136486
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSORDENSDEPRODUCAO", gErr, objPMPItens.objItemOP.iItem, objPMPItens.objItemOP.sCodigo, objPMPItens.objItemOP.iFilialEmpresa)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152306)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Public Function PMP_Simulacao_Grava(ByVal objPMP As ClassPMP) As Long
'Abre a Transação e Chama a PMP_Simulacao_Grava_EmTrans
 
Dim lErro As Long
Dim lTransacao As Long
Dim objPMPItem As ClassPMPItens
Dim bAchou As Boolean
Dim vbmsgbox As VbMsgBoxResult
 
On Error GoTo Erro_PMP_Simulacao_Grava

    bAchou = False
      
    'Para cada item novo do Plano Mestre
    For Each objPMPItem In objPMP.colItens
    
        'Se a OP é planejada Grava
        If objPMPItem.objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA Then
            bAchou = True
            Exit For
        End If
      
    Next
    
    vbmsgbox = vbYes
    
    If bAchou Then
        vbmsgbox = Rotina_Aviso(vbYesNo, "AVISO_PMP_OP_PLANEJADA")
    End If
      
    If vbmsgbox = vbYes Then
      
       'Abertura de transação
       lTransacao = Transacao_Abrir()
       If lTransacao = 0 Then gError 136521
           
       'Chama a Rotina que Gera dentro de transação
       lErro = PMP_Simulacao_Grava_EmTrans(objPMP)
       If lErro <> SUCESSO Then gError 136522
    
       'Confirma a transação
       lErro = Transacao_Commit()
       If lErro <> AD_SQL_SUCESSO Then gError 136523
       
    End If
    
    PMP_Simulacao_Grava = SUCESSO
 
    Exit Function
 
Erro_PMP_Simulacao_Grava:
 
    PMP_Simulacao_Grava = gErr
 
    Select Case gErr
 
        Case 136521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 136522 'Tratado na Rotina Chamada
 
        Case 136523
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152307)
 
    End Select
 
    'Desfaz Transação
    Call Transacao_Rollback
 
    Exit Function
 
End Function
 
Public Function PMP_Simulacao_Grava_EmTrans(ByVal objPMP As ClassPMP) As Long
'Grava uma simulação do PlanoMestre ... grava também as requisições de compra que
'podem ser geradas pelo MRP
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objPMPItens As ClassPMPItens
Dim lCodGeracao As Long
Dim objPO As ClassPlanoOperacional
Dim objOPAux As ClassOrdemDeProducao
Dim objOPOperacao As ClassOrdemProducaoOperacoes
 
On Error GoTo Erro_PMP_Simulacao_Grava_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136487
    Next
 
    'Verifica se já existe um Plano Mestre de produção de simulação com a mesma versão
    lErro = Comando_Executar(alComando(1), "SELECT CodGeracao FROM PlanoMestreProducao WHERE Versao = ? ", lCodGeracao, objPMP.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 136488
    
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136489
    
    'Se a simulação já existe
    If lErro = AD_SQL_SUCESSO Then
 
        objPMP.lCodGeracao = lCodGeracao
 
        'Exclui a simulação
        lErro = PMP_Exclui_EmTrans(objPMP)
        If lErro <> SUCESSO Then gError 136430
 
    End If
 
    'Obter o Código da geração novo
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_CODIGO_PMP", lCodGeracao)
    If lErro <> SUCESSO Then gError 136490
 
    objPMP.dtDataGeracao = gdtDataAtual
 
    'Insere o Plano Mestre
    lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoMestreProducao(CodGeracao, DataGeracao, Versao) VALUES (?,?,?) ", lCodGeracao, objPMP.dtDataGeracao, objPMP.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 136491
 
    objPMP.lCodGeracao = lCodGeracao
 
    'Para cada item novo do Plano Mestre
    For Each objPMPItens In objPMP.colItens
    
        'Se a OP é planejada Grava
        If objPMPItens.objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA Then
        
            Set objOPAux = New ClassOrdemDeProducao
           
            objPMPItens.objOP.colItens.Item(1).dtDataInicioProd = objPMPItens.colPO.Item(1).dtDataInicio
            objPMPItens.objOP.colItens.Item(1).dtDataFimProd = objPMPItens.colPO.Item(1).dtDataFim
            
            objPMPItens.objOP.colItens.Item(1).lNumIntDoc = 0
            
            objOPAux.sCodigo = objPMPItens.objOP.sCodigo
            objOPAux.iFilialEmpresa = objPMPItens.objOP.iFilialEmpresa

            lErro = CF("OrdemProducao_Le", objOPAux)
            If lErro <> SUCESSO And lErro <> 30368 Then gError 138210
            If lErro = SUCESSO Then gError 138211
            
            'Grava a OP
            lErro = CF("OrdemDeProducao_Grava_EmTrans", objPMPItens.objOP)
            If lErro <> SUCESSO Then gError 136429
            
            Set objPMPItens.objItemOP = objPMPItens.objOP.colItens.Item(1)
            Set objPMPItens.colPO.Item(1).objOP = objPMPItens.objOP
            
            'Acerta o NumIntDocOper das Etapas
            For Each objPO In objPMPItens.colPO
            
                For Each objOPOperacao In objPMPItens.objItemOP.colOrdemProducaoOperacoes
            
                    If objPO.iNivel = objOPOperacao.iNivel And objPO.iSeq = objOPOperacao.iSeqArvore And objOPOperacao.sProduto = objOPOperacao.sProduto Then
                        objPO.lNumIntDocOper = objOPOperacao.lNumIntDoc
                        Exit For
                    End If
            
                Next
            
            Next
                
        End If
        
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_PMP_ITENS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 136492
    
        objPMPItens.lNumIntDoc = lNumIntDoc
        objPMPItens.lCodGeracao = objPMP.lCodGeracao
    
        'Insere os Itens do Plano Mestre
        lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoMestreProducaoItens( NumIntDoc, Produto, Versao, Quantidade, " & _
                "UM, DataNecessidade, CodOPOrigem, CodGeracao, Cliente, FilialCliente, FilialEmpresa, Prioridade) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
              objPMPItens.lNumIntDoc, objPMPItens.sProduto, objPMPItens.sVersao, objPMPItens.dQuantidade, _
              objPMPItens.sUM, objPMPItens.dtDataNecessidade, objPMPItens.sCodOPOrigem, objPMPItens.lCodGeracao, objPMPItens.lCliente, objPMPItens.iFilialCli, objPMPItens.iFilialEmpresa, objPMPItens.iPrioridade)
        If lErro <> AD_SQL_SUCESSO Then gError 136493
 
        'Insere a coleção de Etapas no banco
        lErro = PlanoOperacional_Insere_Col(objPMPItens)
        If lErro <> SUCESSO Then gError 136494
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
   PMP_Simulacao_Grava_EmTrans = SUCESSO
 
    Exit Function
 
Erro_PMP_Simulacao_Grava_EmTrans:
 
    PMP_Simulacao_Grava_EmTrans = gErr
 
    Select Case gErr
 
        Case 136487
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136488, 136489
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREDAPRODUCAO", gErr)
 
        Case 136490, 136492, 136494, 136430, 138210, 136429
 
        Case 136491
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOMESTREDAPRODUCAO", gErr)
 
        Case 136493
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOMESTREDAPRODUCAOITENS", gErr)
 
        Case 138211
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_JA_CADASTRADA", gErr, objOPAux.sCodigo)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152308)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Public Function PMP_Exclui_Simulacao(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PMP_Exclui_Simulacao

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 201181
    
    'Chama a Rotina que Exclui dentro de transação
    lErro = CF("PMP_Exclui_EmTrans", objPMP)
    If lErro <> SUCESSO Then gError 201182

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 201183
    
    PMP_Exclui_Simulacao = SUCESSO
    
    Exit Function
    
Erro_PMP_Exclui_Simulacao:

    PMP_Exclui_Simulacao = gErr

    Select Case gErr

        Case 201181
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 201182 'Tratado na Rotina Chamada

        Case 201183
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201184)

    End Select
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Function PMP_Exclui_EmTrans(ByVal objPMP As ClassPMP) As Long
'Exclui um Plano Mestre Simulado pela Versão passada em objPMP
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
 
On Error GoTo Erro_PMP_Exclui_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136431
    Next
    
    'Le a PlanoMestreProducao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM PlanoMestreProducao WHERE Versao = ? ", _
                0, iAux, objPMP.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 136432
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136433
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 136434
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoMestreProducao", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 136435
        
        'Apaga os Itens
        lErro = PMPItens_Exclui_EmTrans(objPMP)
        If lErro <> SUCESSO Then gError 136436
            
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PMP_Exclui_EmTrans = SUCESSO
 
    Exit Function
 
Erro_PMP_Exclui_EmTrans:
 
    PMP_Exclui_EmTrans = gErr
 
    Select Case gErr
 
        Case 136431
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136432, 136433
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAO", gErr)
 
        Case 136434
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOMESTREPRODUCAO", gErr)
 
        Case 136435
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOMESTREPRODUCAO", gErr)
            
        Case 136436
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152309)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Function PMPItens_Exclui_EmTrans(ByVal objPMP As ClassPMP) As Long
'Apaga todos itens de um Plano Mestre
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDocPMP As Integer
Dim objPMPItens As ClassPMPItens
 
On Error GoTo Erro_PMPItens_Exclui_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136437
    Next
    
    'Le a PlanoMestreProducaoItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PlanoMestreProducaoItens WHERE CodGeracao = ? ", _
                0, lNumIntDocPMP, objPMP.lCodGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 136438
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136439
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 136440
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoMestreProducaoItens ", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 136441
        
        Set objPMPItens = New ClassPMPItens
        
        objPMPItens.lNumIntDoc = lNumIntDocPMP
        
        'Apaga as etapas
        lErro = PO_Exclui_EmTrans(objPMPItens)
        If lErro <> SUCESSO Then gError 136442
            
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136443
            
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PMPItens_Exclui_EmTrans = SUCESSO
 
    Exit Function
 
Erro_PMPItens_Exclui_EmTrans:
 
    PMPItens_Exclui_EmTrans = gErr
 
    Select Case gErr
 
        Case 136437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136438, 136439, 136443
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAOITENS", gErr)
 
        Case 136440
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOMESTREPRODUCAOITENS", gErr)
 
        Case 136441
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOMESTREPRODUCAOITENS", gErr)
            
        Case 136442
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152310)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Function PO_Exclui_EmTrans(ByVal objPMPItem As ClassPMPItens) As Long
'Apagas as etapas de um item do Plano Mestre
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPO As ClassPlanoOperacional
Dim lNumIntDocPO As Integer
 
On Error GoTo Erro_PO_Exclui_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136444
    Next
    
    'Le a PlanoOperacional
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PlanoOperacional WHERE NumIntDocPMP = ? ", _
                0, lNumIntDocPO, objPMPItem.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 136445
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136446
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 136447
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoOperacional", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 136448
        
        Set objPO = New ClassPlanoOperacional
        
        objPO.lNumIntDoc = lNumIntDocPO
        
        'Apaga a alocação de máquinas
        lErro = POMaquinas_Exclui_EmTrans(objPO)
        If lErro <> SUCESSO Then gError 136449
        
        'Apaga as possíveis requisições de compras que seriam geradas
        lErro = POReqCompra_Exclui_EmTrans(objPO)
        If lErro <> SUCESSO Then gError 136519
            
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136450
            
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PO_Exclui_EmTrans = SUCESSO
 
    Exit Function
 
Erro_PO_Exclui_EmTrans:
 
    PO_Exclui_EmTrans = gErr
 
    Select Case gErr
 
        Case 136444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136445, 136446, 136450
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOOPERACIONAL", gErr)
 
        Case 136447
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOOPERACIONAL", gErr)
 
        Case 136448
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOOPERACIONAL", gErr)
            
        Case 136449, 136519
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152311)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Function POReqCompra_Exclui_EmTrans(ByVal objPO As ClassPlanoOperacional) As Long
'Apaga as requisições de compra que pertencem a simulação da etapa
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDocPO As Integer
 
On Error GoTo Erro_POReqCompra_Exclui_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136514
    Next
    
    'Le a PlanoOperacionalReqCompras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocPO FROM PlanoOperacionalReqCompras WHERE NumIntDocPO = ? ", _
                0, lNumIntDocPO, objPO.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 136515
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136516
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoOperacionalReqCompras", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 136517
                   
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136518
            
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    POReqCompra_Exclui_EmTrans = SUCESSO
 
    Exit Function
 
Erro_POReqCompra_Exclui_EmTrans:
 
    POReqCompra_Exclui_EmTrans = gErr
 
    Select Case gErr
 
        Case 136514
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136515, 136516, 136518
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POREQCOMPRAS", gErr)
 
        Case 136517
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_POREQCOMPRAS", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152312)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Function POMaquinas_Exclui_EmTrans(ByVal objPO As ClassPlanoOperacional) As Long
'Exclui a alocação de máquinas para a etapa
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer
 
On Error GoTo Erro_POMaquinas_Exclui_EmTrans
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136451
    Next
    
    'Le a PlanoOperacionalMaquinas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PlanoOperacionalMaquinas WHERE NumIntDocPO = ? ", _
                0, lNumIntDoc, objPO.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 136452
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 13653
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 136454
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoOperacionalMaquinas ", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 136455
                   
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136456
            
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    POMaquinas_Exclui_EmTrans = SUCESSO
 
    Exit Function
 
Erro_POMaquinas_Exclui_EmTrans:
 
    POMaquinas_Exclui_EmTrans = gErr
 
    Select Case gErr
 
        Case 136451
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136452, 136453, 136456
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POMAQUINAS", gErr)
 
        Case 136454
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_POMAQUINAS", gErr)
 
        Case 136455
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_POMAQUINAS", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152313)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Private Function PlanoOperacional_Obtem_PCPConfig(lComando As Long, lProxPO As Long) As Long
'Obtém o próximo código para lNumIntDoc do Plano Operacional
 
Dim lErro As Long, sConteudo As String
Dim iIndice As Integer
 
On Error GoTo Erro_PlanoOperacional_Obtem_PCPConfig
  
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136533
   
    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)
 
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM PCPConfig WHERE Codigo = ?", 0, sConteudo, "NUM_INT_PROX_PLANOOPERACIONAL")
    If lErro <> AD_SQL_SUCESSO Then gError 136530
 
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 136531
 
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 136532
 
    lProxPO = CLng(sConteudo)
 
    PlanoOperacional_Obtem_PCPConfig = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Obtem_PCPConfig:
 
    PlanoOperacional_Obtem_PCPConfig = gErr
 
    Select Case gErr
 
        Case 136530, 136531
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PCPCONFIG", gErr, "NUM_INT_PROX_PLANOOPERACIONAL")
 
        Case 136532
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PCPCONFIG", gErr, "NUM_INT_PROX_PLANOOPERACIONAL")
        
        Case 136533
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152314)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function PlanoOperacional_Insere_Col(ByVal objPMPItem As ClassPMPItens) As Long
'Recebe uma coleção de PlanoOperacionais e chama a rotina que grava um a um
 
Dim lErro As Long
Dim objPO As ClassPlanoOperacional
Dim lComando As Long
Dim lNumIntDocInicial As Long
Dim lNumIntDocAcerto As Long
Dim lNumIntDocMaior As Long
 
On Error GoTo Erro_PlanoOperacional_Insere_Col
 
    'Obter NumIntDoc
    lErro = PlanoOperacional_Obtem_PCPConfig(lComando, lNumIntDocAcerto)
    If lErro <> SUCESSO Then gError 136529
    
    lNumIntDocInicial = objPMPItem.colPO.Item(1).lNumIntDoc
 
    'Para cada Plano Operacional da Coleção => Chama a Gravação em transação
    For Each objPO In objPMPItem.colPO
 
        'O NumIntDoc possue somente um valor relativo a seqüência,
        'precisa ser acertado conforme o próximo do BD
        objPO.lNumIntDocPMP = objPMPItem.lNumIntDoc
        objPO.lNumIntDoc = objPO.lNumIntDoc + lNumIntDocAcerto - lNumIntDocInicial
        
        If objPO.lNumIntDocPOPai <> 0 Then
            objPO.lNumIntDocPOPai = objPO.lNumIntDocPOPai + lNumIntDocAcerto - lNumIntDocInicial
        End If
        
        If lNumIntDocMaior < objPO.lNumIntDoc Then
            lNumIntDocMaior = objPO.lNumIntDoc
        End If
 
        'Insere os Planos um a um
        lErro = CF("PlanoOperacional_Insere", objPO)
        If lErro <> SUCESSO Then gError 132485
        
    Next
    
    'Grava qual é o próximo NumIntDoc da Tabela PlanoOperacional
    lErro = PlanoOperacional_Acerta_PCPConfig(lComando, lNumIntDocMaior + 1)
    If lErro <> SUCESSO Then gError 136528
 
    PlanoOperacional_Insere_Col = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Insere_Col:
 
    PlanoOperacional_Insere_Col = gErr
 
    Select Case gErr
 
        Case 132485, 136356, 136528, 136529
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152315)
 
    End Select
 
    Exit Function
 
End Function
 
Private Function PlanoOperacional_Acerta_PCPConfig(lComando As Long, lProxPO As Long) As Long
'Atualiza numeração de lNumIntDoc da tabela PlanoOperacional
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
 
On Error GoTo Erro_PlanoOperacional_Acerta_PCPConfig
 
    'Abre novos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136526
    Next
 
    'Atualiza o registro
    lErro = Comando_ExecutarPos(alComando(0), "UPDATE PCPConfig SET Conteudo = ?", lComando, lProxPO)
    If lErro <> AD_SQL_SUCESSO Then gError 136527
 
    Call Comando_Fechar(lComando)
 
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PlanoOperacional_Acerta_PCPConfig = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Acerta_PCPConfig:
 
    PlanoOperacional_Acerta_PCPConfig = gErr
 
    Select Case gErr
 
        Case 136526
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136527
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PCPCONFIG", gErr, "NUM_INT_PROX_PLANOOPERACIONAL")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152316)
 
    End Select
 
    'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Public Function PlanoOperacional_Insere(ByVal objPlanoOperacional As ClassPlanoOperacional) As Long
'Insere uma etapa de um item do Plano Mestre de Produção e suas alocações de máquina e
'requisições de compra
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim objRC As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objTaxaDeProducao As ClassTaxaDeProducao
 
On Error GoTo Erro_PlanoOperacional_Insere
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 132494
    Next
 
    'Insere o Plano Operacional
    lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoOperacional( NumIntDoc, NumIntDocPMP, NumIntDocOper, Seq, CodOPOrigem, " & _
            "Produto, Versao, Quantidade, UM, NumIntDocCT, DataInicio, DataFim, NumIntDocPOPai, Nivel) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objPlanoOperacional.lNumIntDoc, objPlanoOperacional.lNumIntDocPMP, objPlanoOperacional.lNumIntDocOper, objPlanoOperacional.iSeq, objPlanoOperacional.sCodOPOrigem, _
            objPlanoOperacional.sProduto, objPlanoOperacional.sVersao, objPlanoOperacional.dQuantidade, objPlanoOperacional.sUM, objPlanoOperacional.lNumIntDocCT, _
            objPlanoOperacional.dtDataInicio, objPlanoOperacional.dtDataFim, objPlanoOperacional.lNumIntDocPOPai, objPlanoOperacional.iNivel)
    If lErro <> AD_SQL_SUCESSO Then gError 132495
    
    'Para cada Alocação de Máquinas no Plano Operacional
    For Each objPOMaquinas In objPlanoOperacional.colAlocacaoMaquinas
        
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_POMAQUINAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 132496
        
        objPOMaquinas.lNumIntDoc = lNumIntDoc
        objPOMaquinas.lNumIntDocPO = objPlanoOperacional.lNumIntDoc
        
        Set objTaxaDeProducao = objPOMaquinas.objTaxaProducao
        
        objPOMaquinas.lNumIntDocTxProd = objTaxaDeProducao.lNumIntDoc
    
        'Insere a alocação de máquinas no plano operacional
        With objPOMaquinas
            lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoOperacionalMaquinas (Data, HorasMaquinas,Quantidade,NumIntDoc,NumIntDocMaq,NumIntDocPO,NumIntDocTxProd) VALUES (?,?,?,?,?,?,?) ", _
            .dtData, .dHorasMaquina, .iQuantidade, .lNumIntDoc, .lNumIntDocMaq, .lNumIntDocPO, .lNumIntDocTxProd)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 132497
 
    Next
    
    'Para cada requisição de compra que pode ser gerada
    For Each objRC In objPlanoOperacional.colRCFilhas
    
        For Each objItemRC In objRC.colItens
        
            With objItemRC
                'Insere as requisições de compra relativas a etapa
                lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoOperacionalReqCompras (NumIntDocPO,Produto,Quantidade,UM) VALUES (?,?,?,?) ", _
                objPlanoOperacional.lNumIntDoc, .sProduto, .dQuantidade, .sUM)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 136520
        
        Next
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PlanoOperacional_Insere = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Insere:
 
    PlanoOperacional_Insere = gErr
 
    Select Case gErr
 
        Case 132494
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 132495
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOOPERACIONAL", gErr)
 
        Case 132496
 
        Case 132497
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOOPERACIONALMAQUINAS", gErr)
            
        Case 136520
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_POREQCOMPRAS", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152317)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Public Function MRP_Simula(ByVal objPMP As ClassPMP) As Long
'Gera o Plano Operacional
'Aloca as máquinas que estão sendo usadas pelo Plano Operacional
'Calcula as necessidades e datas de início e fim
 
Dim lErro As Long
Dim objPMPItens As ClassPMPItens
Dim objOPOperacao As ClassOrdemProducaoOperacoes
Dim objOPInsumos As ClassOrdemProducaoInsumos
Dim objPO As ClassPlanoOperacional
Dim objPOAux1 As ClassPlanoOperacional
Dim objPOAux2 As ClassPlanoOperacional
Dim objOP As ClassOrdemDeProducao
Dim colPO As Collection
Dim objItemOP As ClassItemOP
Dim iProximo As Integer
Dim colCampos As New Collection
Dim colSaida As New Collection
Dim lProxPO As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objProdutoFilial As ClassProdutoFilial
Dim dEstoqueProjetado As Double
Dim objProduto As ClassProduto
Dim dFatorConversao As Double
Dim dFatorAux As Double
Dim objOPTempoAux As ClassOperacoesTempo
Dim objOPOperacaoAux As ClassOrdemProducaoOperacoes
Dim objOPInsumosAux As ClassOrdemProducaoInsumos
Dim bFolha As Boolean
Dim objCT As ClassCentrodeTrabalho
Dim iIndice As Integer
Dim iNumTentativas As Integer
Dim iDiasDif As Integer
Dim dtDataMenor As Date
Dim objItemOPAux As ClassItemOP
Dim iNumLotes As Integer

On Error GoTo Erro_MRP_Simula

    'gsMRPTexto = "Indice;Seq;Produto;Data;QtdInsumo;CMR;CSR;VR;PR;CMP;CSP;VP;PP;QtdProjetada;EstoqueSeguranca;QtdAtual"
 
    iProximo = 0
        
    'Ordena pela prioridade
    colCampos.Add "iPrioridade"
    colCampos.Add "dtDataNecessidade"
   
    'Tem que ordenar pela data para poder pegar o estoque projetado corretamente
    lErro = Ordena_Colecao(objPMP.colItens, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 138105
    
    Set objPMP.colItens = colSaida
    
    'Roda duas vezes para primeiro calcular tudo em ordem e
    'depois refazer já considerando o cenário completo
    'Tem que rever todas as quantidades necessárias para ajustar as sobras geradas posteriormente as necessidades mas com data inferior
    'Tratar Lotes e Estoque de segurança -> Processar o estoque projetado de cada produto ajustando as quantidades necessárias, tempos, etc
    For iIndice = 1 To 2
'        giMRPIndice = iIndice
'        giMRPSeq = 0
        If iIndice = 1 Then
            For Each objPMPItens In objPMP.colItens
                objPMPItens.dQuantidade = objPMPItens.dQuantidadeNecess
                Set objPMPItens.colPOAux = New Collection
            Next
        Else
            'Adiciona o que foi calculado na primeira passada
            'para auxiliar no cálculo da projeção de estoque
            For Each objPMPItens In objPMP.colItens
                Set objPMPItens.colPOAux = New Collection
                For Each objPO In objPMPItens.colPO
                    objPMPItens.colPOAux.Add objPO
                Next
            Next
        End If
        For Each objPMPItens In objPMP.colItens
            Set objPMPItens.colPO = New Collection
        Next
            
        'Para cada item do plano mestre
        For Each objPMPItens In objPMP.colItens
                                
            Set objOPOperacao = New ClassOrdemProducaoOperacoes
            Set objOPInsumos = New ClassOrdemProducaoInsumos
            Set objPO = New ClassPlanoOperacional
            Set colPO = New Collection
                    
            objOPOperacao.iNivelRoteiro = 0
            objOPOperacao.sProduto = objPMPItens.sProduto
            objOPOperacao.sVersao = objPMPItens.sVersao
            objOPOperacao.lNumIntDocItemOP = objPMPItens.objItemOP.lNumIntDoc
            
            'Se a OP for gerada pelo MRP verifica a necessidade de quantidade para ela no dia
            If objPMPItens.objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA Then
            
                Set objEstoqueProduto = New ClassEstoqueProduto
                Set objProdutoFilial = New ClassProdutoFilial
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objPMPItens.sProduto
        
                'Le o produto para obter a ClasseUM
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 138440
        
                objEstoqueProduto.sProduto = objPMPItens.sProduto
                objEstoqueProduto.iFilialEmpresa = objPMPItens.iFilialEmpresa
            
                'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objPMPItens.sUM, dFatorConversao)
                If lErro <> SUCESSO Then gError 138441
            
                'Lê as quantidades do produtos projetada para data
                lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objPMPItens.dtDataNecessidade, dEstoqueProjetado, objPMP)
                If lErro <> SUCESSO And lErro <> 52979 Then gError 138442
                
'                If objPMPItens.dtDataEntregaPV <> DATA_NULA Then
                    'Se vai faltar mais que o necessário para atender o pedido de venda, então só atende a esse pedido de venda
                    'Se vai faltar menos produz só o que vai faltar.
'                    If objPMPItens.dQuantidade > -((dEstoqueProjetado - objProdutoFilial.dEstoqueSeguranca) * dFatorConversao) Then
'                        objPMPItens.dQuantidade = -((dEstoqueProjetado - objProdutoFilial.dEstoqueSeguranca) * dFatorConversao)
'                    End If
'                    objPMPItens.dQuantidade = -((dEstoqueProjetado - objProdutoFilial.dEstoqueSeguranca) * dFatorConversao)
'                Else
                'Se é para produzir com base na data inicial é porque não temn datra de entrega, e por isso
                'Não vai estar incluso no estoque projetado, logo, deve ser considerado nesse ponto
                    objPMPItens.dQuantidade = objPMPItens.dQuantidade - ((dEstoqueProjetado - objProdutoFilial.dEstoqueSeguranca) * dFatorConversao)
'                End If
            
                If objPMPItens.dQuantidade <= 0 Then
                    objPMPItens.dQuantidade = 0
                    objPMPItens.objOP.iStatusOP = ITEMOP_SITUACAO_DESAB
                Else
                    If Not (objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes Is Nothing) Then
                        If Not (objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes Is Nothing) Then
                            If Not (objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1) Is Nothing) Then
                                If Not (objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1).objOperacoesTempo Is Nothing) Then
                                    If objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1).objOperacoesTempo.dLoteMin > objPMPItens.dQuantidade Then
                                        If objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1).objOperacoesTempo.dLoteMin < objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1).objOperacoesTempo.dLotePadrao Then
                                            objPMPItens.dQuantidade = objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1).objOperacoesTempo.dLotePadrao
                                        Else
                                            objPMPItens.dQuantidade = objPMPItens.objOP.colItens.Item(1).colOrdemProducaoOperacoes.Item(1).objOperacoesTempo.dLoteMin
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                'Se a produção for em lotes ajusta para ser múltiplo do lote
                If objProdutoFilial.dLoteEconomico <> 0 Then
                    iNumLotes = Round(objPMPItens.dQuantidade / objProdutoFilial.dLoteEconomico, 0)
                    If (objProdutoFilial.dLoteEconomico * iNumLotes) < objPMPItens.dQuantidade Then iNumLotes = iNumLotes + 1
                    objPMPItens.dQuantidade = objProdutoFilial.dLoteEconomico * iNumLotes
                End If
                
                If Abs(objPMPItens.dQuantidade - objPMPItens.objOP.colItens.Item(1).dQuantidade) > QTDE_ESTOQUE_DELTA Then
    
                    dFatorAux = objPMPItens.dQuantidade / objPMPItens.objOP.colItens.Item(1).dQuantidade
        
                    objPMPItens.objOP.colItens.Item(1).dQuantidade = objPMPItens.dQuantidade
                    
                    For Each objOPOperacaoAux In objPMPItens.objItemOP.colOrdemProducaoOperacoes
                    
                        Set objOPTempoAux = objOPOperacao.objOperacoesTempo
                        
                        objOPTempoAux.dTempoOperacao = objOPTempoAux.dTempoOperacao * dFatorAux
                    
                        For Each objOPInsumosAux In objOPOperacaoAux.colOPInsumos
                        
                            objOPInsumosAux.dQuantidade = objOPInsumosAux.dQuantidade * dFatorAux
                        
                        Next
                    
                    Next
            
                End If
            
            End If
            
            objOPInsumos.sProduto = objPMPItens.sProduto
            objOPInsumos.sVersaoKitComp = objPMPItens.sVersao
            objOPInsumos.dQuantidade = objPMPItens.dQuantidade
            objOPInsumos.sUMProduto = objPMPItens.sUM
                            
            objPO.dtDataInicio = objPMPItens.dtDataNecessidade
            objPO.dQuantidade = objPMPItens.dQuantidade
            objPO.sUM = objPMPItens.sUM
            objPO.sProduto = objPMPItens.sProduto
            objPO.sVersao = objPMPItens.sVersao
            objPO.iFilialEmpresa = objPMPItens.iFilialEmpresa
            
            If objPMPItens.objItemOP.colOrdemProducaoOperacoes.Count > 0 Then
            
                objOPInsumos.lNumIntDocOper = objPMPItens.objItemOP.colOrdemProducaoOperacoes.Item(1).lNumIntDoc
                
                lProxPO = 0
                
                If objPMPItens.dQuantidade > 0 Then
                
                    'Gera o Plano Operacional ligado ao plano mestre com a devida alocação de máquinas
                    lErro = PlanoOperacional_Gera(objPMP, objPMPItens, objOPOperacao, objOPInsumos, objPO, objPMPItens.objOP, colPO, lProxPO)
                    If lErro <> SUCESSO Then gError 136322
                
                    Set objPMPItens.colPO = colPO
    
                    objPMPItens.lUltimoProxPO = lProxPO
    
                    '########################################################################
                    If objPMPItens.iProduzLogo = MARCADO Then
                    
                        'Le o item da OP para pegar a data inicial de produção original
                        Set objItemOPAux = New ClassItemOP
                        
                        objItemOPAux.iFilialEmpresa = objPMPItens.objItemOP.iFilialEmpresa
                        objItemOPAux.sCodigo = objPMPItens.objItemOP.sCodigo
                        objItemOPAux.sProduto = objPMPItens.objItemOP.sProduto
                        
                        'Se ainda não existe a ordem de produção
                        If objPMPItens.objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA Then
                            objItemOPAux.dtDataInicioProd = DateAdd("d", PRODUCAO_PV_DELAY, gdtDataAtual)
                        Else
                            lErro = CF("ItemOP_Le", objItemOPAux)
                            If lErro <> SUCESSO And lErro <> 34711 Then gError 177328
                            
                            If objItemOPAux.dtDataInicioProd <= gdtDataAtual Then objItemOPAux.dtDataInicioProd = DateAdd("d", PRODUCAO_PV_DELAY, gdtDataAtual)
                        End If
                            
                        'Obtém o início de produção calculado
                        dtDataMenor = DATA_NULA
        
                        For Each objPOAux2 In objPMPItens.colPO
                            If objPOAux2.dtDataInicio < dtDataMenor Or dtDataMenor = DATA_NULA Then
                                dtDataMenor = objPOAux2.dtDataInicio
                            End If
                        Next
                        
                        iNumTentativas = 0
                        
                        'Enquando não bater o início de produção calculado com o almejado e
                        'estiver dentro do limite de tentativas ...
                        Do While dtDataMenor <> objItemOPAux.dtDataInicioProd And (iNumTentativas < NUM_MAX_TENTATIVAS_PROD_DATA_INI)
                        
                            iNumTentativas = iNumTentativas + 1
            
                            'Obtém a etapa raiz
                            Set objPOAux1 = objPMPItens.colPO.Item(1)
            
                            'Calcula quanto falta para ficar com a data almejada
                            iDiasDif = DateDiff("d", dtDataMenor, objItemOPAux.dtDataInicioProd)
                            
                            objPOAux1.dtDataInicio = DateAdd("d", iDiasDif, objPOAux1.dtDataFim)
                            
                            Set colPO = New Collection
                            Set objPMPItens.colPO = New Collection
            
                            'Recalcula com a nova data
                            lErro = PlanoOperacional_Gera(objPMP, objPMPItens, objOPOperacao, objOPInsumos, objPOAux1, objPMPItens.objOP, colPO, lProxPO)
                            If lErro <> SUCESSO Then gError 177329
        
                            Set objPMPItens.colPO = colPO
            
                            objPMPItens.lUltimoProxPO = lProxPO
                        
                            'Obtém o início de produção calculado
                            dtDataMenor = DATA_NULA
            
                            For Each objPOAux2 In objPMPItens.colPO
                                If objPOAux2.dtDataInicio < dtDataMenor Or dtDataMenor = DATA_NULA Then
                                    dtDataMenor = objPOAux2.dtDataInicio
                                End If
                            Next
                        
                        Loop
                            
                    End If
                    '#########################################################
                    
                End If
                
                If objPMPItens.colPO.Count > 0 Then
                    
                    'Obtém o início de produção calculado
                    dtDataMenor = DATA_NULA
    
                    For Each objPOAux2 In objPMPItens.colPO
                        If objPOAux2.dtDataInicio < dtDataMenor Or dtDataMenor = DATA_NULA Then
                            dtDataMenor = objPOAux2.dtDataInicio
                        End If
                    Next
                            
                    objPMPItens.objItemOP.dtDataFimProd = objPMPItens.colPO.Item(1).dtDataFim
                    objPMPItens.objItemOP.dtDataInicioProd = dtDataMenor
                    
                End If
            
            End If
            
        Next
    
    Next
    
    objPMP.iUltimoProxOP = iProximo
        
    Set colCampos = New Collection
    Set colSaida = New Collection
    
    'Ordena pela prioridade
    colCampos.Add "iPrioridade"
    colCampos.Add "iOrdem"
   
    'Volta para Ordenação antiga
    lErro = Ordena_Colecao(objPMP.colItens, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 138106
    
    Set objPMP.colItens = colSaida
    
'    Open "C:\MRP.csv" For Output As #1
'    Print #1, gsMRPTexto
'    Close #1
    
    MRP_Simula = SUCESSO
    
    Exit Function
 
Erro_MRP_Simula:
 
    MRP_Simula = gErr
 
    Select Case gErr
 
        Case 136322, 138105, 138106, 138440, 138441, 138442, 177328, 177329
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152318)
 
    End Select
 
    Exit Function
 
End Function

Public Function MRP_Valida(ByVal objPMP As ClassPMP) As Long
'Faz as validações referentes a capacidade máquina e estoque necessário
'a execução de cada etapa
 
Dim lErro As Long
Dim objPMPItens As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim sMensagem As String
Dim sMensagemPMP As String
Dim objPMPBD As New ClassPMP
 
On Error GoTo Erro_MRP_Valida
        
    'Le o plano mestre já gravado
    lErro = CF("PlanoMestreProducao_Le", objPMPBD)
    If lErro <> SUCESSO And lErro <> 136298 Then gError 136333
        
    'Para cada item do Plano
    For Each objPMPItens In objPMP.colItens
       
        sMensagemPMP = ""
        
        objPMPItens.iStatus = PO_STATUS_OK
       
        'Para cada Etapa
        For Each objPO In objPMPItens.colPO
                
            sMensagem = ""
                
            objPO.iStatus = PO_STATUS_OK
                
            'Verifica se vai conseguir comprar os produtos que faltam a tempo
            lErro = MRP_Valida_Compras(objPMP, objPO, sMensagem)
            If lErro <> SUCESSO Then gError 136323
                
            'Verifica se vai conseguir produzir a tempo
            lErro = MRP_Valida_Producao(objPMP, objPO, sMensagem)
            If lErro <> SUCESSO Then gError 136324
                
            'Verifica se não tem algum CT sobrecarregado
            lErro = MRP_Valida_SobreCarga(objPMPBD, objPMP, objPMPItens, objPO, sMensagem)
            If lErro <> SUCESSO Then gError 136325
 
            'Se tiver erro no status do PO, o erro no coloca no status do PMP
            If objPO.iStatus <> PO_STATUS_OK Then
                sMensagemPMP = sMensagemPMP & sMensagem
            End If
            
            'Acumula o Pior Status dentre os dos PO para o PMP
            Select Case objPMPItens.iStatus
                            
                Case PO_STATUS_FALTAMATERIAL
                    If objPO.iStatus = PO_STATUS_SOBRECARGA Or objPO.iStatus = PO_STATUS_AMBOS Then
                        objPMPItens.iStatus = PO_STATUS_AMBOS
                    End If
                    
                Case PO_STATUS_SOBRECARGA
                    If objPO.iStatus = PO_STATUS_FALTAMATERIAL Or objPO.iStatus = PO_STATUS_AMBOS Then
                        objPMPItens.iStatus = PO_STATUS_AMBOS
                    End If
               
               Case PO_STATUS_AMBOS
               
               Case Else
                    objPMPItens.iStatus = objPO.iStatus
            
            End Select
 
            If Len(Trim(sMensagem)) = 0 Then Call MRP_Mensagem1(sMensagem)
            
            objPO.sDescErro = sMensagem
                    
        Next
        
        If Len(Trim(sMensagemPMP)) = 0 Then sMensagemPMP = sMensagem
        
        objPMPItens.sDescErro = sMensagemPMP
                        
    Next
    
    MRP_Valida = SUCESSO
    
    Exit Function
 
Erro_MRP_Valida:
 
    MRP_Valida = gErr
 
    Select Case gErr
 
        Case 136323 To 136325, 136333
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152319)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Valida_Compras(ByVal objPMP As ClassPMP, ByVal objPO As ClassPlanoOperacional, sMensagem As String) As Long
'Faz as validações referentes as requisições de compra
 
Dim lErro As Long
Dim objRC As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim dtData As Date
Dim iMaiorTempoRessuprimento As Integer
Dim objProdutoFilial As ClassProdutoFilial
Dim sProduto As String
 
On Error GoTo Erro_MRP_Valida_Compras
    
    'Para cada requisição de compra
    For Each objRC In objPO.colRCFilhas
    
        iMaiorTempoRessuprimento = 0
    
        'Para cada item da requisição
        For Each objItemRC In objRC.colItens
        
            Set objProdutoFilial = New ClassProdutoFilial
        
            objProdutoFilial.sProduto = objItemRC.sProduto
            objProdutoFilial.iFilialEmpresa = objPO.iFilialEmpresa
        
            'Le ProdutoFilial
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO And lErro <> 28261 Then gError 136326
        
            'Obtém o maior tempo de ressuprimenro dentre os itens
            If objProdutoFilial.iTempoRessup >= iMaiorTempoRessuprimento Then
                iMaiorTempoRessuprimento = objProdutoFilial.iTempoRessup
                sProduto = objProdutoFilial.sProduto
            End If
        
        Next
        
        'Calcula a data máxima para a requisição
        dtData = DateAdd("d", -1 * iMaiorTempoRessuprimento, objPO.dtDataInicio)
 
        'Se não dá mais tempo => Muda Status
        If dtData < gdtDataAtual Then
        
            If objPO.iStatus = PO_STATUS_SOBRECARGA Or objPO.iStatus = PO_STATUS_AMBOS Then
                objPO.iStatus = PO_STATUS_AMBOS
            Else
                objPO.iStatus = PO_STATUS_FALTAMATERIAL
            End If
            
            Call MRP_Mensagem3(sMensagem, objPO.iNivel, objPO.iSeq, sProduto, dtData)
        
        End If
    
    Next
    
    MRP_Valida_Compras = SUCESSO
    
    Exit Function
 
Erro_MRP_Valida_Compras:
 
    MRP_Valida_Compras = gErr
 
    Select Case gErr
 
        Case 136326
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152320)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Valida_Producao(ByVal objPMP As ClassPMP, ByVal objPO As ClassPlanoOperacional, sMensagem As String) As Long
'Faz as validações referentes a produção de subprodutos
 
Dim lErro As Long
 
On Error GoTo Erro_MRP_Valida_Producao
    
    'Se a data de início da produção já passou => Muda Status
    If objPO.dtDataInicio < gdtDataAtual Then
        
        If objPO.iStatus = PO_STATUS_SOBRECARGA Or objPO.iStatus = PO_STATUS_AMBOS Then
            objPO.iStatus = PO_STATUS_AMBOS
        Else
            objPO.iStatus = PO_STATUS_FALTAMATERIAL
        End If
    
        Call MRP_Mensagem5(sMensagem, objPO.sProduto, objPO.dtDataInicio)
    
    End If
 
    MRP_Valida_Producao = SUCESSO
    
    Exit Function
 
Erro_MRP_Valida_Producao:
 
    MRP_Valida_Producao = gErr
 
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152321)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Valida_SobreCarga(ByVal objPMPBD As ClassPMP, ByVal objPMP As ClassPMP, ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, sMensagem As String) As Long
'Verifica se tem algum CT sobrecarregado
 
Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMaquinas As ClassCTMaquinas
Dim iMaqQtdDia As Integer
Dim dMaqTempoDia As Double
Dim dtData As Date
Dim dUsoMaquina As Double
Dim colUsoOperadores As Collection
Dim colMaquinas As New Collection
Dim objMaquinaOper As ClassMaquinaOperadores
Dim objPOMaquinasAux As ClassPOMaquinas
Dim iIndice As Integer
Dim dOperTempoDia As Double
Dim objCTOperador As ClassCTOperadores
 
On Error GoTo Erro_MRP_Valida_SobreCarga
 
    objCT.lNumIntDoc = objPO.lNumIntDocCT
       
    'Le o centro de trabalho
    lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
    If lErro <> SUCESSO And lErro <> 137212 Then gError 136327
       
    'Para cada máquina alocada para o trabalho
    For Each objPOMaquinas In objPO.colAlocacaoMaquinas
        
        Set objCTMaquinas = New ClassCTMaquinas
         
        objCTMaquinas.lNumIntDocCT = objCT.lNumIntDoc
        objCTMaquinas.lNumIntDocMaq = objPOMaquinas.lNumIntDocMaq
     
        'Verifica a disponibilidade da máquina naquele dia
        lErro = Verifica_Maquinas_Disponiveis_Dia(objCTMaquinas, objPOMaquinas.dtData, objCT.iFilialEmpresa, iMaqQtdDia, dMaqTempoDia)
        If lErro <> SUCESSO Then gError 136328
        
        Set colUsoOperadores = New Collection
    
        For Each objMaquinaOper In objPOMaquinas.objMaquina.colTipoOperadores
            
            Set objPOMaquinasAux = New ClassPOMaquinas
            
            objPOMaquinasAux.lNumIntDocMaq = objMaquinaOper.lNumIntDocMaq
            
            colUsoOperadores.Add objPOMaquinasAux
        Next
    
        dUsoMaquina = 0
         
        'Verifica o que está sendo usado no próprio item do plano mestre
        lErro = MRP_Valida_SobreCarga_Mesmo_PMPItem(objPMPItem, objPO, objPOMaquinas, dUsoMaquina, colUsoOperadores)
        If lErro <> SUCESSO Then gError 136329
         
        'Verifica o que os outros itens do plano mestre estão usando
        lErro = MRP_Valida_SobreCarga_Outro_PMPItem(objPMP, objPMPItem, objPO, objPOMaquinas, dUsoMaquina, colUsoOperadores)
        If lErro <> SUCESSO Then gError 136330
         
        'Verifica o que os itens já gravados estão usando
        'Verifica o que os outros itens do plano mestre estão usando
        lErro = MRP_Valida_SobreCarga_Outro_PMPItem(objPMPBD, objPMPItem, objPO, objPOMaquinas, dUsoMaquina, colUsoOperadores)
        If lErro <> SUCESSO Then gError 136330
                             
        'se a capacidade do dia ficou estourada => Muda Status
        If dMaqTempoDia - dUsoMaquina < 0 Then
             
            If objPO.iStatus = PO_STATUS_FALTAMATERIAL Or objPO.iStatus = PO_STATUS_AMBOS Then
                objPO.iStatus = PO_STATUS_AMBOS
            Else
                objPO.iStatus = PO_STATUS_SOBRECARGA
            End If
         
            Call MRP_Mensagem2(sMensagem, objPO.iNivel, objPO.iSeq, objCT.lCodigo & SEPARADOR & objCT.sNomeReduzido, objPOMaquinas.objMaquina.iCodigo & SEPARADOR & objPOMaquinas.objMaquina.sNomeReduzido, objPOMaquinas.dtData)
         
        End If
        
        iIndice = 0
        
        For Each objMaquinaOper In objPOMaquinas.objMaquina.colTipoOperadores
            
            iIndice = iIndice + 1
            
            Set objPOMaquinasAux = colUsoOperadores.Item(iIndice)
            
            'Encontra o tipo de mão de obra no CT
            'Se não existisse teria dado erro antes
            For Each objCTOperador In objCT.colOperadores
                If objCTOperador.iCodTipoMO = objMaquinaOper.iTipoMaoDeObra Then
                    Exit For
                End If
            Next
        
            dOperTempoDia = objCTOperador.iQuantidade * (dMaqTempoDia / iMaqQtdDia)
            
            If dOperTempoDia - objPOMaquinasAux.dHorasMaquina < 0 Then
            
                If objPO.iStatus = PO_STATUS_FALTAMATERIAL Or objPO.iStatus = PO_STATUS_AMBOS Then
                    objPO.iStatus = PO_STATUS_AMBOS
                Else
                    objPO.iStatus = PO_STATUS_SOBRECARGA
                End If
            
                Call MRP_Mensagem4(sMensagem, objPO.iNivel, objPO.iSeq, objCT.lCodigo & SEPARADOR & objCT.sNomeReduzido, objMaquinaOper.iTipoMaoDeObra, objPOMaquinas.dtData)
            
            End If
        
        Next
        
    Next
    
    MRP_Valida_SobreCarga = SUCESSO
    
    Exit Function
 
Erro_MRP_Valida_SobreCarga:
 
    MRP_Valida_SobreCarga = gErr
 
    Select Case gErr
 
        Case 136327 To 136331
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152322)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Valida_SobreCarga_Mesmo_PMPItem(ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal objPOMaquinas As ClassPOMaquinas, dUsoMaquina As Double, ByVal colUsoOperadores As Collection) As Long
'Obtém o uso de uma máquina dentro do item do Plano Mestre passado
 
Dim lErro As Long
Dim iIndice As Integer
Dim objPOAux As ClassPlanoOperacional
Dim objPOMaquinasAux As ClassPOMaquinas
Dim objPOMaquinasUso As ClassPOMaquinas
Dim objMaquinaOper As ClassMaquinaOperadores
Dim objMaquinaOperAux As ClassMaquinaOperadores
Dim dHoras As Double
 
On Error GoTo Erro_MRP_Valida_SobreCarga_Mesmo_PMPItem
             
    'Para cada plano Operacional
    For Each objPOAux In objPMPItem.colPO
 
        'para cada máquina usada nessa etapa
        For Each objPOMaquinasAux In objPOAux.colAlocacaoMaquinas
                
            'Se ocorrer uso na mesma data
            If objPOMaquinasAux.dtData = objPOMaquinas.dtData And objPOAux.lNumIntDocCT = objPO.lNumIntDocCT Then
                    
                'Se for a mesma máquina
                If objPOMaquinasAux.lNumIntDocMaq = objPOMaquinas.lNumIntDocMaq Then
                    
                    dUsoMaquina = dUsoMaquina + objPOMaquinasAux.dHorasMaquina
        
                End If
                
                iIndice = 0
                
                For Each objMaquinaOper In objPOMaquinas.objMaquina.colTipoOperadores
                     
                    iIndice = iIndice + 1
                                         
                    For Each objMaquinaOperAux In objPOMaquinasAux.objMaquina.colTipoOperadores
                    
                        If objMaquinaOper.iTipoMaoDeObra = objMaquinaOperAux.iTipoMaoDeObra Then
                        
                            Set objPOMaquinasUso = colUsoOperadores.Item(iIndice)
                        
                            dHoras = objPOMaquinasUso.dHorasMaquina + (objMaquinaOperAux.iQuantidade * objMaquinaOperAux.dPercentualUso * objPOMaquinasAux.dHorasMaquina)
                        
                            objPOMaquinasUso.dHorasMaquina = dHoras
                        
                        End If
                    
                    Next
               
                Next
                
            End If
        
        Next
            
    Next
    
    MRP_Valida_SobreCarga_Mesmo_PMPItem = SUCESSO
    
    Exit Function
 
Erro_MRP_Valida_SobreCarga_Mesmo_PMPItem:
 
    MRP_Valida_SobreCarga_Mesmo_PMPItem = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152323)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Valida_SobreCarga_Outro_PMPItem(ByVal objPMP As ClassPMP, ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal objPOMaquinas As ClassPOMaquinas, dUsoMaquina As Double, ByVal colUsoOperadores As Collection) As Long
'Obtém o uso de uma máquina dentro dos item do Plano Mestre anteriores ao item passado
 
Dim lErro As Long
Dim objPMPItemAux As ClassPMPItens
 
On Error GoTo Erro_MRP_Valida_SobreCarga_Outro_PMPItem
    
    'Para cada item do Plano mestre
    For Each objPMPItemAux In objPMP.colItens
    
        If objPMPItemAux.objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
    
            'Se já chegou no item atual não pode mais verificar
            If objPMPItemAux.objItemOP.lNumIntDoc = objPMPItem.objItemOP.lNumIntDoc Then Exit For
        
            'Obtém o uso para cada item diferente do passado em objPMPItem
            lErro = MRP_Valida_SobreCarga_Mesmo_PMPItem(objPMPItemAux, objPO, objPOMaquinas, dUsoMaquina, colUsoOperadores)
            If lErro <> SUCESSO Then gError 136332
    
        End If
    
    Next
    
    MRP_Valida_SobreCarga_Outro_PMPItem = SUCESSO
    
    Exit Function
 
Erro_MRP_Valida_SobreCarga_Outro_PMPItem:
 
    MRP_Valida_SobreCarga_Outro_PMPItem = gErr
 
    Select Case gErr
 
        Case 136332
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152324)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Mensagem1(sMensagem As String) As Long
'Mensagem padrão para Status OK
 
Dim lErro As Long
 
On Error GoTo Erro_MRP_Mensagem1
    
    sMensagem = "OK. "
   
    MRP_Mensagem1 = SUCESSO
    
    Exit Function
 
Erro_MRP_Mensagem1:
 
   MRP_Mensagem1 = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152325)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Mensagem2(sMensagem As String, ByVal iNivel As Integer, ByVal iSeq As Integer, ByVal sCT As String, ByVal sMaquina As String, ByVal dtData As Date) As Long
'Mensagem padrão para Status de SobreCarga
 
Dim lErro As Long
 
On Error GoTo Erro_MRP_Mensagem2
    
    sMensagem = sMensagem & "Houve Sobrecarga no CT " & sCT & ", Máquina " & sMaquina & " no dia " & Format(dtData, "dd/mm/yyyy") & ". "
    
    MRP_Mensagem2 = SUCESSO
    
    Exit Function
 
Erro_MRP_Mensagem2:
 
   MRP_Mensagem2 = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152326)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Mensagem3(sMensagem As String, ByVal iNivel As Integer, ByVal iSeq As Integer, ByVal sProduto As String, ByVal dtData As Date) As Long
'Mensagem padrão para Status de Falta de Material quando não há tempo para comprar
 
Dim lErro As Long
Dim sProdutoMascarado As String
 
On Error GoTo Erro_MRP_Mensagem3
    
    sProdutoMascarado = String(STRING_PRODUTO, 0)
    
    lErro = Mascara_RetornaProdutoTela(sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 136555
    
    sMensagem = sMensagem & "A Data Limite da Req. de Compra do Produto " & Trim(sProdutoMascarado) & " era no dia " & Format(dtData, "dd/mm/yyyy") & ". "
    
    MRP_Mensagem3 = SUCESSO
    
    Exit Function
 
Erro_MRP_Mensagem3:
 
   MRP_Mensagem3 = gErr
 
    Select Case gErr
    
        Case 136555
  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152327)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Mensagem4(sMensagem As String, ByVal iNivel As Integer, ByVal iSeq As Integer, ByVal sCT As String, ByVal iTipoMaoDeObra As Integer, ByVal dtData As Date) As Long
'Mensagem alternativa para Status de Sobrecarga com Falta de material
 
Dim lErro As Long
 
On Error GoTo Erro_MRP_Mensagem4
    
    sMensagem = sMensagem & "Houve Sobrecarga no CT " & sCT & ", Operador " & CStr(iTipoMaoDeObra) & " no dia " & Format(dtData, "dd/mm/yyyy") & ". "
   
    MRP_Mensagem4 = SUCESSO
    
    Exit Function
 
Erro_MRP_Mensagem4:
 
   MRP_Mensagem4 = gErr
 
    Select Case gErr
  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152328)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function MRP_Mensagem5(sMensagem As String, ByVal sProduto As String, ByVal dtData As Date) As Long
'Mensagem padrão para status de Falta de Material quando não há possibilidade de produzir um subproduto a tempo
 
Dim lErro As Long
Dim sProdutoMascarado As String
 
On Error GoTo Erro_MRP_Mensagem5
    
    sProdutoMascarado = String(STRING_PRODUTO, 0)
    
    lErro = Mascara_RetornaProdutoTela(sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 138114
    
    sMensagem = sMensagem & "O Produto " & Trim(sProdutoMascarado) & " teria que ter começado a ser produzido no dia " & Format(dtData, "dd/mm/yyyy") & ". "
   
    MRP_Mensagem5 = SUCESSO
    
    Exit Function
 
Erro_MRP_Mensagem5:
 
   MRP_Mensagem5 = gErr
 
    Select Case gErr
        
        Case 138114
  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152329)
 
    End Select
 
    Exit Function
 
End Function
 
Private Sub PMP_Prepara_SQL(ByVal objMRPSeleciona As ClassMRPSeleciona, sSQL As String)
'Monta o SELECT que traz os dados necessários para geração do plano Mestre
 
Dim sSelect As String
Dim sWhere As String
Dim sWhereAux As String
Dim sFrom As String
Dim sOrderBy As String
Dim sSubSelect As String
 
On Error GoTo Erro_PMP_Prepara_SQL
    
    sSubSelect = "(SELECT PMPItens.CodOPOrigem " & _
                    "FROM PlanoMestreProducaoItens AS PMPItens, " & _
                            "PlanoMestreProducaoUltimo AS PMP " & _
                    "WHERE PMP.Codgeracao = PMPItens.CodGeracao AND " & _
                            "PMPItens.CodOPOrigem = OP.Codigo AND " & _
                            "PMPItens.FilialEmpresa = OP.FilialEmpresa AND " & _
                            "PMPItens.Produto = IOP.Produto) "
 
    sSelect = "SELECT OP.Codigo, " & _
                        "OP.FilialEmpresa, " & _
                        "IOP.Produto, " & _
                        "PV.Cliente, " & _
                        "IPV.DataEntrega "
 
    sFrom = "FROM      OrdensDeProducao AS OP, " & _
                        "ItensOrdemProducao AS IOP  LEFT OUTER JOIN " & _
                        "ItensPedidoDeVenda AS IPV ON " & _
                            "IPV.CodPedido = IOP.CodPedido AND " & _
                            "IOP.Produto = IPV.Produto LEFT OUTER JOIN " & _
                        "PedidosDeVenda AS PV ON " & _
                            "PV.Codigo = IPV.CodPedido AND " & _
                            "PV.FilialEmpresa = IPV.FilialEmpresa "
                    
    sWhere = "WHERE    OP.Codigo = IOP.Codigo AND " & _
                        "OP.OPGeradora =  '' AND " & _
                        "NOT EXISTS "
    
    sWhere = sWhere & sSubSelect
    
    sOrderBy = "ORDER BY IOP.Prioridade "
      
    If objMRPSeleciona.dtDataOPFinal <> DATA_NULA Then
        sWhereAux = sWhereAux & "AND OP.DataEmissao <= ? "
    End If
    If objMRPSeleciona.dtDataOPInicial <> DATA_NULA Then
        sWhereAux = sWhereAux & "AND OP.DataEmissao >= ? "
    End If
    
    If objMRPSeleciona.lClienteFinal <> 0 Then
        sWhereAux = sWhereAux & "AND PV.Cliente <= ? "
    End If
    If objMRPSeleciona.lClienteInicial <> 0 Then
        sWhereAux = sWhereAux & "AND PV.Cliente >= ? "
    End If
    
    If objMRPSeleciona.sOPFinal <> "" Then
        sWhereAux = sWhereAux & "AND OP.Codigo <= ? "
    End If
    If objMRPSeleciona.sOPInicial <> "" Then
        sWhereAux = sWhereAux & "AND OP.Codigo >= ? "
    End If
    
    If objMRPSeleciona.sProdutoFinal <> "" Then
        sWhereAux = sWhereAux & "AND IOP.Produto <= ? "
    End If
    If objMRPSeleciona.sProdutoInicial <> "" Then
        sWhereAux = sWhereAux & "AND IOP.Produto >= ? "
    End If
            
    If gobjEST.dtDataInicioMRP <> DATA_NULA Then
        sWhereAux = sWhereAux & "AND OP.DataEmissao >= ? "
    End If
            
    sWhere = sWhere & sWhereAux
    
    sSQL = sSelect & sFrom & sWhere & sOrderBy
 
    Exit Sub
 
Erro_PMP_Prepara_SQL:
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152330)
 
    End Select
 
    Exit Sub
 
End Sub

Private Function PMP_Prepara_Int(ByVal lComando As Long, vsCodigoOP As Variant, viFilialEmpresa As Variant, vsProduto As Variant, vlCliente As Variant, vdtDataNecessidade As Variant, ByVal vdtDataFim As Variant, ByVal vdtDataIni As Variant, ByVal vlClienteFim As Variant, ByVal vlClienteIni As Variant, ByVal vsCodOPFim As Variant, ByVal vsCodOPIni As Variant, ByVal vsProdutoFim As Variant, ByVal vsProdutoIni As Variant, ByVal sSQL As String) As Long
'Passa as variáveis de retorno e filtro do SELECT que monta os dados a serem faturados
 
Dim lErro As Long
Dim vdtDataInicioMRP As Variant 'Inserido por Jorge Specian
 
On Error GoTo Erro_PMP_Prepara_Int
 
    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136335
 
    lErro = Comando_BindVarInt(lComando, vsCodigoOP)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136336
    
    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136337
        
    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136338
   
    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136338
   
    lErro = Comando_BindVarInt(lComando, vdtDataNecessidade)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136338
   
    If vdtDataFim <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136339
    End If
    If vdtDataIni <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136340
    End If
 
    If vlClienteFim <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136341
    End If
    If vlClienteIni <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136342
    End If
 
    If vsCodOPFim <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCodOPFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136343
    End If
    If vsCodOPIni <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCodOPIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136344
    End If
 
    If vsProdutoFim <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdutoFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136345
    End If
    If vsProdutoIni <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdutoIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136346
    End If
    
    'Inserido por Jorge Specian
    '-----------------------------------------
    If gobjEST.dtDataInicioMRP <> DATA_NULA Then
        vdtDataInicioMRP = gobjEST.dtDataInicioMRP
        
        lErro = Comando_BindVarInt(lComando, vdtDataInicioMRP)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137699
    End If
    '-----------------------------------------
     
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136347
 
   PMP_Prepara_Int = SUCESSO
 
    Exit Function
 
Erro_PMP_Prepara_Int:
 
    PMP_Prepara_Int = gErr
 
    Select Case gErr
 
        Case 136335 To 136347, 137699  'Inserido por Jorge Specian
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152331)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function PMP_Gera(ByVal objMRPSeleciona As ClassMRPSeleciona, ByVal objPMP As ClassPMP) As Long
 
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim sCodigoOP As String
Dim iFilialEmpresa As Integer
Dim objOP As ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim objPMPItens As ClassPMPItens
Dim lCodigoGeracao As Long
Dim iIndice As Integer
Dim sProduto As String
Dim sSQL As String
Dim lCliente As Long
Dim dtDataNecessidade As Date
 
On Error GoTo Erro_PMP_Gera
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136348
    Next
    
    'Monta o Select
    Call PMP_Prepara_SQL(objMRPSeleciona, sSQL)
       
    sCodigoOP = String(STRING_ORDEM_DE_PRODUCAO, 0)
    sProduto = String(STRING_PRODUTO, 0)
    
    'Passa os parâmetros e executa o select
    lErro = PMP_Prepara_Int(alComando(1), sCodigoOP, iFilialEmpresa, sProduto, lCliente, dtDataNecessidade, objMRPSeleciona.dtDataOPFinal, objMRPSeleciona.dtDataOPInicial, objMRPSeleciona.lClienteFinal, objMRPSeleciona.lClienteInicial, objMRPSeleciona.sOPFinal, objMRPSeleciona.sOPInicial, objMRPSeleciona.sProdutoFinal, objMRPSeleciona.sProdutoInicial, sSQL)
    If lErro <> SUCESSO Then gError 136350
    
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136351
        
    iIndice = 0
        
    'Para cada OP válida
    Do While lErro = SUCESSO
    
        Set objOP = New ClassOrdemDeProducao
            
        objOP.sCodigo = sCodigoOP
        objOP.iFilialEmpresa = iFilialEmpresa
    
        'tenta ler a OP desejada
        lErro = CF("OrdemProducao_Le", objOP)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 136352
        If lErro = SUCESSO And objOP.iTipo <> OP_TIPO_OC Then
    
            'Le a ordem de produção associada ao PMP
            'Le as demais informações associadas a Ordem de Produção
            lErro = CF("OrdemDeProducao_Le_ComItens", objOP)
            If lErro <> SUCESSO And lErro <> 21960 Then gError 136353
        
            For Each objItemOP In objOP.colItens
            
                'Verifica se a situação do item está = NORMAL
                If objItemOP.iSituacao = ITEMOP_SITUACAO_NORMAL And objItemOP.sProduto = sProduto Then
           
                    Set objPMPItens = New ClassPMPItens
                    
                    iIndice = iIndice + 1
                   
                    objPMPItens.dQuantidade = objItemOP.dQuantidade
                    objPMPItens.dQuantidadeNecess = objItemOP.dQuantidade
                    
                    If dtDataNecessidade = DATA_NULA Then
                        objPMPItens.dtDataNecessidade = objItemOP.dtDataFimProd
                    Else
                        objPMPItens.dtDataNecessidade = dtDataNecessidade
                    End If
                    
                    objPMPItens.iPrioridade = objItemOP.iPrioridade
                    objPMPItens.iFilialEmpresa = objOP.iFilialEmpresa
                    objPMPItens.sCodOPOrigem = objOP.sCodigo
                    objPMPItens.sProduto = objItemOP.sProduto
                    objPMPItens.sUM = objItemOP.sSiglaUM
                    objPMPItens.sVersao = objItemOP.sVersao
                    objPMPItens.lCliente = lCliente
                    objPMPItens.iProduzLogo = objItemOP.iProduzLogo
                    
                    Set objPMPItens.objItemOP = objItemOP
                    Set objPMPItens.objOP = objOP
                    
                    objPMPItens.iOrdem = iIndice
                    
                    'Adiciona o Item ao plano mestre de produção
                    objPMP.colItens.Add objPMPItens
    
                End If
            
            Next
            
        End If
    
        'Busca o proximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136354
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PMP_Gera = SUCESSO
 
    Exit Function
 
Erro_PMP_Gera:
 
    PMP_Gera = gErr
 
    Select Case gErr
    
        Case 136348
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 136350, 136352, 136353
        
        Case 136351, 136354
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152332)
 
    End Select
 
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
 
End Function
 
Public Function PlanoOperacional_Gera(ByVal objPMP As ClassPMP, ByVal objPMPItens As ClassPMPItens, ByVal objOPOperacaoPai As ClassOrdemProducaoOperacoes, ByVal objOPInsumosOrigem As ClassOrdemProducaoInsumos, ByVal objPOPai As ClassPlanoOperacional, ByVal objOPOrigem As ClassOrdemDeProducao, ByVal colPO As Collection, lProxPO As Long) As Long
'Rotina recursiva para geração das etapas de um Item do Plano Mestre de produção
 
Dim lErro As Long
Dim colOperacoes As New Collection
Dim iIndice As Integer
Dim objOPOperacao As ClassOrdemProducaoOperacoes
Dim objPO As New ClassPlanoOperacional
 
On Error GoTo Erro_PlanoOperacional_Gera
    
    'Le as Operações da Ordem de Produção Roteiro a Roteiro
    lErro = CF("OrdemDeProducao_Le_Roteiro", objPMPItens.objItemOP, objOPOperacaoPai, objOPInsumosOrigem, colOperacoes)
    If lErro <> SUCESSO Then gError 136355
 
    'Para cada Operação
    For Each objOPOperacao In colOperacoes
    
        Set objPO = New ClassPlanoOperacional
            
        lErro = PlanoOperacional_Gera2(objPMP, objPMPItens, objOPOperacaoPai, objOPInsumosOrigem, objPOPai, objOPOrigem, colPO, objOPOperacao, colOperacoes, objPO, lProxPO, iIndice)
        If lErro <> SUCESSO Then gError 138315

    Next
 
    PlanoOperacional_Gera = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Gera:
 
    PlanoOperacional_Gera = gErr
 
    Select Case gErr
    
        Case 136355, 138315

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152333)
 
    End Select
 
    Exit Function
 
End Function

Public Function PlanoOperacional_Gera2(ByVal objPMP As ClassPMP, ByVal objPMPItens As ClassPMPItens, ByVal objOPOperacaoPai As ClassOrdemProducaoOperacoes, ByVal objOPInsumosOrigem As ClassOrdemProducaoInsumos, ByVal objPOPai As ClassPlanoOperacional, ByVal objOPOrigem As ClassOrdemDeProducao, ByVal colPO As Collection, ByVal objOPOperacao As ClassOrdemProducaoOperacoes, ByVal colOperacoes As Collection, ByVal objPO As ClassPlanoOperacional, lProxPO As Long, iIndice As Integer) As Long
'Rotina recursiva para geração das etapas de um Item do Plano Mestre de produção
 
Dim lErro As Long
Dim objOPOperacaoAux As ClassOrdemProducaoOperacoes
Dim objPOAux As ClassPlanoOperacional
Dim objOPInsumos As ClassOrdemProducaoInsumos
Dim objProduto As ClassProduto
Dim objEstoqueProduto As ClassEstoqueProduto
Dim bAchouProduto As Boolean
Dim dFatorConversao As Double
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objOP As ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim objNecesProdInfo As ClassNecesProdInfo
Dim objOPOperacaoFilho As ClassOrdemProducaoOperacoes
Dim lNumIntDoc As Long
Dim iAlmoxarifadoPadrao As Integer
Dim objPMPItensAux As ClassPMPItens
Dim dValorAux As Double
Dim dFatorAumentoLote As Double
Dim dEstoqueProjetado As Double
Dim objProdutoFilial As ClassProdutoFilial
Dim sCodigoOP As String
Dim objKit As ClassKit
Dim objPMPAux As ClassPMP, iNumLotes As Integer
Dim dQtdeAux As Double

On Error GoTo Erro_PlanoOperacional_Gera2
    
    lProxPO = lProxPO + 1
    
    'Obter NumIntDoc
    lNumIntDoc = lProxPO

    'Cria o Plano Operacional que é o detalhamento por etapas
    'de um item do plano mestre
    objPO.lNumIntDoc = lNumIntDoc
    objPO.dQuantidade = objOPOrigem.colItens(1).dQuantidade  'objOPInsumosOrigem.dQuantidade ... vai ser em cima da quantidade que vai ser produzida e não da quantidade consumida
    objPO.lNumIntDocCT = objOPOperacao.lNumIntDocCT
    objPO.lNumIntDocPMP = objPMPItens.lNumIntDoc
    objPO.lNumIntDocOper = objOPOperacao.lNumIntDoc
    objPO.sCodOPOrigem = objOPOrigem.sCodigo
    objPO.sProduto = objOPInsumosOrigem.sProduto
    objPO.sUM = objOPInsumosOrigem.sUMProduto
    objPO.sVersao = objOPInsumosOrigem.sVersaoKitComp
    objPO.iNivel = objOPOperacao.iNivel
    objPO.iSeq = objOPOperacao.iSeqArvore
    objPO.iFilialEmpresa = objPOPai.iFilialEmpresa
    objPO.iNumMaxMaqPorOper = objOPOperacao.iNumMaxMaqPorOper
    objPO.iNumRepeticoes = objOPOperacao.iNumRepeticoes
    
    'Encontra o Pai e a Data de Fim com base na data de início do Pai
    'Se é a primeira operação do roteiro
    If iIndice = 0 Then
        'Se for uma Operação com origem em outro roteiro
        objPO.lNumIntDocPOPai = objPOPai.lNumIntDoc
        objPO.dtDataFim = objPOPai.dtDataInicio
    Else
        'Se for uma operação com origem no mesmo roteiro
        For Each objOPOperacaoAux In colOperacoes
            If objOPOperacaoAux.iSeq = objOPOperacao.iSeqPai Then
                For Each objPOAux In colPO
                    If objPOAux.lNumIntDocOper = objOPOperacaoAux.lNumIntDoc Then
                        objPO.lNumIntDocPOPai = objPOAux.lNumIntDoc
                        objPO.dtDataFim = objPOAux.dtDataInicio
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    End If
    
    iIndice = iIndice + 1
    
    Set objPO.objOP = objOPOrigem
    Set objPMPItensAux = New ClassPMPItens
    
    For Each objPOAux In colPO
        objPMPItensAux.colPO.Add objPOAux
    Next
    
    'Guarda a quantidade
    dValorAux = objPO.dQuantidade
    
    Set objPMPAux = New ClassPMP
    
    For iIndice = 1 To objPMP.colItens.Count
        'Adiciona todos os outros itens da produção exceto o que está sendo calculado no momento
        'e posteriores
        If objPMPItens.sCodOPOrigem <> objPMP.colItens.Item(iIndice).sCodOPOrigem Or objPMPItens.sProduto <> objPMP.colItens.Item(iIndice).sProduto Then
            objPMPAux.colItens.Add objPMP.colItens.Item(iIndice)
            Exit For
        End If
    Next
    
    'Obtem as máquinas e tempos de produção para cada PlanoOperacional
    lErro = PlanoOperacional_Calcula_Tempos(objPMPAux, objPMPItensAux, objPO, objOPOperacao, MRP_ACERTA_POR_DATA_FIM)
    If lErro <> SUCESSO Then gError 136357

    'Se alterou a quantidade tem que guardar o fator de alteração para acertar os itens
    'dependentes (insumos)
    dFatorAumentoLote = objPO.dQuantidade / dValorAux

    colPO.Add objPO
    objPMPItensAux.colPO.Add objPO 'Adiciona inclusive o que ainda está sendo planejado pois o estoque projetado passou a retornar a quantidade projetada com o cenário completo
    
    'Acerta a quantidade de acordo com o lote
    For Each objItemOP In objOPOrigem.colItens
        If objItemOP.sProduto = objPO.sProduto Then
            objItemOP.dQuantidade = objPO.dQuantidade
            Exit For
        End If
    Next

    'Para cada insumo utilizado
    For Each objOPInsumos In objOPOperacao.colOPInsumos
            
        'Acerta os Insumos
        objOPInsumos.dQuantidade = dFatorAumentoLote * objOPInsumos.dQuantidade
            
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objOPInsumos.sProduto
         
        'Le o produto para obter a ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 136358
        
        If lErro <> SUCESSO Then gError 136359
        
        'Se o Produto for Comprável o produzível
        If (objProduto.iCompras = PRODUTO_COMPRAVEL) Or (objProduto.iCompras = PRODUTO_PRODUZIVEL) Then

            Set objEstoqueProduto = New ClassEstoqueProduto
            Set objNecesProdInfo = New ClassNecesProdInfo
            Set objProdutoFilial = New ClassProdutoFilial
    
            objEstoqueProduto.sProduto = objProduto.sCodigo
            objEstoqueProduto.iFilialEmpresa = objPO.iFilialEmpresa
            
            'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOPInsumos.sUMProduto, objProduto.sSiglaUMEstoque, dFatorConversao)
            If lErro <> SUCESSO Then gError 136361
            
'            'If InStr(1, objProduto.sCodigo, "0011167") <> 0 Then
'            If InStr(1, objProduto.sCodigo, "2143") <> 0 Then
'                lErro = lErro
'            End If
            
            'Lê as quantidades do produtos projetada para data
            Set objPMPAux = New ClassPMP
            
            For iIndice = 1 To objPMP.colItens.Count
                'Adiciona todos os outros itens da proção exceto o que está sendo calculado no momento
                'para esse vai pegar o que está sendo calculado e não o que já foi calculado
                If objPMPItens.sCodOPOrigem <> objPMP.colItens.Item(iIndice).sCodOPOrigem Or objPMPItens.sProduto <> objPMP.colItens.Item(iIndice).sProduto Then
                    objPMPAux.colItens.Add objPMP.colItens.Item(iIndice)
                End If
            Next
            objPMPAux.colItens.Add objPMPItensAux
            objPMPItensAux.iFilialEmpresa = objPO.iFilialEmpresa
            Set objPMPItensAux.objItemOP = objPMPItens.objItemOP
            Set objPMPItensAux.objOP = objPMPItens.objOP
            Set objPMPItensAux.colPOAux = objPMPItens.colPOAux
            
'            'gsMRPTexto = "Indice;Seq;Produto;Data;QtdInsumo;CMR;CSR;VR;PR;CMP;CSP;VP;PP;QtdProjetada;EstoqueSeguranca;QtdAtual"
'            giMRPSeq = giMRPSeq + 1
'            gsMRPTexto = gsMRPTexto & vbNewLine & CStr(giMRPIndice) & ";" & CStr(giMRPSeq) & ";" & objProduto.sCodigo & ";" & Format(objPO.dtDataInicio, "dd/mm/yyyy") & ";" & Formata_Estoque(objOPInsumos.dQuantidade)

            lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objPO.dtDataInicio, dEstoqueProjetado, objPMPAux)
            If lErro <> SUCESSO And lErro <> 52979 Then gError 136360
            
'            gsMRPTexto = gsMRPTexto & ";" & Formata_Estoque(dEstoqueProjetado) & ";" & Formata_Estoque(objProdutoFilial.dEstoqueSeguranca) & ";" & Formata_Estoque(objEstoqueProduto.dQuantDisponivel)
                                 
            'Considera como saldo disponível a quantidade disponível em
            'estoque menos o que for estoque de segurança projetado para data já considerando tudo
            'planejado para aquela data inclusive o consumo do próprio produto
            objNecesProdInfo.sProduto = objEstoqueProduto.sProduto
            objNecesProdInfo.dSaldoDisponivel = dEstoqueProjetado - objProdutoFilial.dEstoqueSeguranca
            
            'Código alterado para não mexer na quantidade consumida (estava afetando quando havia uma nova simulação)
            'Quantidade necessária para deixar o estoque somente com estoque de segurança
            'objOPInsumos.dQuantidade = -(objNecesProdInfo.dSaldoDisponivel / dFatorConversao) / (1 - objOPInsumos.dPercentualPerda)
            dQtdeAux = -(objNecesProdInfo.dSaldoDisponivel / dFatorConversao) / (1 - objOPInsumos.dPercentualPerda) 'Só vai mexer na quantidade se faltar estoque
            
'            'If InStr(1, objProduto.sCodigo, "0011167") <> 0 Then
'            If InStr(1, objProduto.sCodigo, "2143") <> 0 Then
'                lErro = lErro
'            End If
            
            'Se não tem disponível suficiente (precisa produzir)
            If dQtdeAux > 0 Then
            
                'Se a produção for em lotes ajusta para ser múltiplo do lote
                If objProdutoFilial.dLoteEconomico <> 0 Then
                    iNumLotes = Round(dQtdeAux / objProdutoFilial.dLoteEconomico, 0) 'Round(objOPInsumos.dQuantidade / objProdutoFilial.dLoteEconomico, 0)
                    If (objProdutoFilial.dLoteEconomico * iNumLotes) < dQtdeAux Then iNumLotes = iNumLotes + 1 'objOPInsumos.dQuantidade Then iNumLotes = iNumLotes + 1
                    dQtdeAux = objProdutoFilial.dLoteEconomico * iNumLotes 'objOPInsumos.dQuantidade = objProdutoFilial.dLoteEconomico * iNumLotes
                End If

                'Se o produto for produzível
                If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                    
                    Set objOP = New ClassOrdemDeProducao
                    Set objItemOP = New ClassItemOP
                    
                    'Se o insumo não tiver a versão tem que pegar a padrão
                    If Len(Trim(objOPInsumos.sVersaoKitComp)) = 0 Then
                                                        
                        Set objKit = New ClassKit
                
                        objKit.sProdutoRaiz = objOPInsumos.sProduto
                    
                        'Le as Versoes Ativas e a Padrao
                        lErro = CF("Kit_Le_Padrao", objKit)
                        If lErro <> SUCESSO And lErro <> 106304 Then gError 139045
                        
                        objOPInsumos.sVersaoKitComp = objKit.sVersao
                    
                    End If
                                        
                    'Pega um código automático para a OP
                    lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
                    If lErro <> SUCESSO Then gError 137936
                    
                    objOP.sCodigo = sCodigoOP
                    objOP.iFilialEmpresa = objOPOrigem.iFilialEmpresa
                    objOP.sOPGeradora = objOPOrigem.sCodigo
                    objOP.dtDataEmissao = gdtDataAtual
                    
                    lErro = CF("AlmoxarifadoPadrao_Le", objOP.iFilialEmpresa, objOPInsumos.sProduto, iAlmoxarifadoPadrao)
                    If lErro <> SUCESSO And lErro <> 23796 Then gError 136362
                    If iAlmoxarifadoPadrao = 0 Then gError 136363
                    
                    objItemOP.iItem = 1
                    objItemOP.dQuantidade = dQtdeAux 'objOPInsumos.dQuantidade
                    objItemOP.dtDataFimProd = objPO.dtDataInicio
                    objItemOP.iFilialEmpresa = objOPOrigem.iFilialEmpresa
                    objItemOP.iAlmoxarifado = iAlmoxarifadoPadrao
                    objItemOP.iDestinacao = ITEMOP_DESTINACAO_ESTOQUE
                    objItemOP.sProduto = objOPInsumos.sProduto
                    objItemOP.sSiglaUM = objOPInsumos.sUMProduto
                    objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
                    objItemOP.sVersao = objOPInsumos.sVersaoKitComp
                    objItemOP.sCodigo = objOP.sCodigo
                                            
                    objOP.colItens.Add objItemOP
                
                    objOP.iNumItens = objOP.colItens.Count
                    
                    'Guarda a SubOP no PlanoOperacional (Vai ser salvo depois)
                    objPO.colOPFilhas.Add objOP
                    
                    Set objOPOperacaoFilho = New ClassOrdemProducaoOperacoes
 
                    objOPOperacaoFilho.iNivelRoteiro = objOPOperacaoPai.iNivelRoteiro + 1
                    objOPOperacaoFilho.sProduto = objOPInsumos.sProduto
                    objOPOperacaoFilho.sVersao = objOPInsumos.sVersaoKitComp
                    objOPOperacaoFilho.lNumIntDocItemOP = objOPOperacaoPai.lNumIntDocItemOP
                    objOPOperacaoFilho.iSeqPai = objOPOperacao.iSeq

                    'Chama a função recursivamente
                    lErro = PlanoOperacional_Gera(objPMP, objPMPItens, objOPOperacaoFilho, objOPInsumos, objPO, objOP, colPO, lProxPO)
                    If lErro <> SUCESSO Then gError 136364
                    
                End If
                
                'Se o produto for comprável
                If objProduto.iCompras = PRODUTO_COMPRAVEL Then
                
                    Set objReqCompras = New ClassRequisicaoCompras
                    Set objItemRC = New ClassItemReqCompras
                                    
                    'Cria um ItemRC novo com o produto
                    lErro = CF("AlmoxarifadoPadrao_Le", objPO.iFilialEmpresa, objOPInsumos.sProduto, iAlmoxarifadoPadrao)
                    If lErro <> SUCESSO And lErro <> 23796 Then gError 136365
                                    
                    objItemRC.iItem = 1
                    
                    'Descobre o fator de conversao da UM de estoque para a de Compra
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOPInsumos.sUMProduto, objProduto.sSiglaUMCompra, dFatorConversao)
                    If lErro <> SUCESSO Then gError 136366
                    
                    objReqCompras.dtDataLimite = objPO.dtDataInicio 'Data Limite = Data que vai ser consumido o produto
                    objReqCompras.dtData = gdtDataAtual
                    objReqCompras.dtDataBaixa = DATA_NULA
                    objReqCompras.dtDataEnvio = DATA_NULA
                    objReqCompras.iFilialCompra = objPO.iFilialEmpresa
                    objReqCompras.iFilialDestino = objPO.iFilialEmpresa
                    objReqCompras.iFilialEmpresa = objPO.iFilialEmpresa
                    objReqCompras.iTipoDestino = TIPO_DESTINO_EMPRESA
                    
                    objReqCompras.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO
                    objReqCompras.sOPCodigo = objOPOrigem.sCodigo
            
                    objItemRC.dQuantidade = dQtdeAux * dFatorConversao 'objOPInsumos.dQuantidade * dFatorConversao
                    objItemRC.iAlmoxarifado = iAlmoxarifadoPadrao
                    objItemRC.sDescProduto = objProduto.sDescricao
                    objItemRC.sProduto = objOPInsumos.sProduto
                    objItemRC.sUM = objProduto.sSiglaUMCompra
        
                    objReqCompras.colItens.Add objItemRC
                
                    'Guarda a requisição de compra no Plano Operacional
                    objPO.colRCFilhas.Add objReqCompras
                End If
            End If
        End If
    Next
 
    PlanoOperacional_Gera2 = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Gera2:
 
    PlanoOperacional_Gera2 = gErr
 
    Select Case gErr
    
        Case 136355 To 136358, 136360 To 136362, 136364, 136366, 136816, 137936, 139045
        
        Case 136359
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
               
        Case 136363
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objOPInsumos.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152334)

    End Select

    Exit Function
 
End Function

Public Function PlanoOperacional_Calcula_Tempos(ByVal objPMP As ClassPMP, ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal objOPOperacao As ClassOrdemProducaoOperacoes, ByVal iFlag As Integer) As Long
'Obtém uma coleção de Máquinas aptas a fazer o serviço e uma coleção com suas
'respectivas taxas de produção
 
Dim lErro As Long
Dim objCT As ClassCentrodeTrabalho
Dim objCTMaquinas As ClassCTMaquinas
Dim objTaxaDeProducao As ClassTaxaDeProducao
Dim objOperacaoTempo As ClassOperacoesTempo
Dim colMaqAptas As New Collection
Dim colTaxaUsada As New Collection
Dim dMenorLoteMin As Double
Dim dMenorLotePadrao As Double
 
On Error GoTo Erro_PlanoOperacional_Calcula_Tempos
        
    Set objTaxaDeProducao = New ClassTaxaDeProducao
    Set objOperacaoTempo = objOPOperacao.objOperacoesTempo
        
    'se o CT não está preenchido... pode ser ORIGEM_KIT e
    'OP gerada por Pedido de Venda onde o usuário deixou em branco
    If objOPOperacao.lNumIntDocCT = 0 Then gError 137711
 
    Set objCT = New ClassCentrodeTrabalho
 
    objCT.lNumIntDoc = objOPOperacao.lNumIntDocCT
  
    'Le o CT e todos objs envolvidos
    lErro = CF("CentroDeTrabalho_Le_Completo", objCT)
    If lErro <> SUCESSO And lErro <> 137212 Then gError 136367
    
    'Se não encontrar => Erro
    If lErro <> SUCESSO Then gError 136368
    
    'Se não for considerar o cadastro da Taxa de produção
    If objOPOperacao.iIgnoraTaxaProducao = MARCADO Then
    
        'Marca a Taxa de produção da operação para todas máquinas do CT
        Set objTaxaDeProducao = New ClassTaxaDeProducao
        
        dMenorLoteMin = objOperacaoTempo.dLoteMin
        dMenorLotePadrao = objOperacaoTempo.dLotePadrao
        
        objTaxaDeProducao.dLoteMax = objOperacaoTempo.dLoteMax
        objTaxaDeProducao.dLoteMin = objOperacaoTempo.dLoteMin
        objTaxaDeProducao.dLotePadrao = objOperacaoTempo.dLotePadrao
        objTaxaDeProducao.dQuantidade = objPO.dQuantidade
        objTaxaDeProducao.dTempoDescarga = objOperacaoTempo.dTempoDescarga
        objTaxaDeProducao.dTempoMovimentacao = objOperacaoTempo.dTempoMovimentacao
        objTaxaDeProducao.dTempoOperacao = objOperacaoTempo.dTempoOperacao
        objTaxaDeProducao.dTempoPreparacao = objOperacaoTempo.dTempoPreparacao
        objTaxaDeProducao.iTipo = objOperacaoTempo.iTipo
        objTaxaDeProducao.sUMProduto = objPO.sUM
        objTaxaDeProducao.sUMTempo = objOperacaoTempo.sUMTempo
        objTaxaDeProducao.lNumIntDocCompet = objOPOperacao.lNumIntDocCompet
        objTaxaDeProducao.sProduto = objPO.sProduto
        
        'Coloca todas máquinas do CT como aptas e com a mesma taxa de produção
        For Each objCTMaquinas In objCT.colMaquinas
            
            If objOperacaoTempo.lNumIntDocMaq = 0 Or objOperacaoTempo.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq Then
            
                'Adiciona a Máquina a coleção de máquinas aptas
                colMaqAptas.Add objCTMaquinas
                
                'Adiciona a Taxa de produção que foi usada para o cálculo
                colTaxaUsada.Add objTaxaDeProducao
                
            End If
            
        Next
        
    Else
                
        'Para cada Tipo de Máquina do CT
        For Each objCTMaquinas In objCT.colMaquinas
        
            Set objTaxaDeProducao = New ClassTaxaDeProducao
            
            objTaxaDeProducao.sProduto = objPO.sProduto
            objTaxaDeProducao.lNumIntDocCompet = objOPOperacao.lNumIntDocCompet
            objTaxaDeProducao.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq
    
            'Busca a Taxa de Produção
            lErro = CF("TaxaDeProducao_Seleciona", objTaxaDeProducao)
            If lErro <> SUCESSO And lErro <> 134600 Then gError 136370
            
            'Se esse tipo de máquina está apto a fazer tal tarefa
            If lErro = SUCESSO Then
                            
                If (dMenorLoteMin = 0) Or (dMenorLoteMin > objTaxaDeProducao.dLoteMin And objTaxaDeProducao.dLoteMin <> 0) Then
                    dMenorLoteMin = objTaxaDeProducao.dLoteMin
                End If
                
                If (dMenorLotePadrao = 0) Or (dMenorLotePadrao > objTaxaDeProducao.dLotePadrao And objTaxaDeProducao.dLotePadrao <> 0) Then
                    dMenorLotePadrao = objTaxaDeProducao.dLotePadrao
                End If
                                           
                'Adiciona a Máquina a coleção de máquinas aptas
                colMaqAptas.Add objCTMaquinas
                
                'Adiciona a Taxa de produção que foi usada para o cálculo
                colTaxaUsada.Add objTaxaDeProducao
            
            End If
            
        Next
                
    End If
    
    'Se o que eu tenho que produzir é menor que o lote mínimo
    If (objPO.dQuantidade < dMenorLoteMin And dMenorLoteMin <> 0) And objPO.iNivel <> KIT_NIVEL_RAIZ Then
        'Se tiver lote padrão
        If dMenorLotePadrao > dMenorLoteMin Then
            objPO.dQuantidade = dMenorLotePadrao 'Produz o Padrão
        Else
            objPO.dQuantidade = dMenorLoteMin 'Produz o mínimo
        End If
    End If
    
    'Aloca as máquinas para o Plano Operacional
    lErro = PlanoOperacional_Calcula_Tempos2(objPMP, objPMPItem, objPO, colMaqAptas, colTaxaUsada, objCT, iFlag, objOPOperacao)
    If lErro <> SUCESSO Then gError 136371
       
    PlanoOperacional_Calcula_Tempos = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Calcula_Tempos:
 
    PlanoOperacional_Calcula_Tempos = gErr
 
    Select Case gErr
    
        Case 136367, 136369 To 136371
    
        Case 136368
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO2", gErr, objCT.lCodigo)
            
        Case 137711
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_INFORMADO_NA_OP", gErr, objOPOperacao.iNivel, objOPOperacao.iSeqArvore, Trim(objPO.sProduto), objPO.sCodOPOrigem)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152335)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function PlanoOperacional_Calcula_Tempos2(ByVal objPMP As ClassPMP, ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal colCTMaquinas As Collection, ByVal colTaxaDeProducao As Collection, ByVal objCT As ClassCentrodeTrabalho, ByVal iFlag As Integer, ByVal objOPOperacao As ClassOrdemProducaoOperacoes) As Long
'Descide sobre alocação do uso das máquinas no tempo que vai ser gasto
 
Dim lErro As Long
Dim bTemMaquinasAptas As Boolean
Dim objProduto As ClassProduto
Dim objCompetencias As ClassCompetencias
 
On Error GoTo Erro_PlanoOperacional_Calcula_Tempos2
    
    'Se for para Considerar a Carga de Maquinas (MRPII)...
    If objOPOperacao.iConsideraCarga = MARCADO Then
    
        bTemMaquinasAptas = True
        
        If colCTMaquinas.Count = 0 Then
            
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objPO.sProduto
            
            'Le o produto para obter o Nome Reduzido
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 139073
                        
            Set objCompetencias = New ClassCompetencias
            
            objCompetencias.lNumIntDoc = objOPOperacao.lNumIntDocCompet
            
            'Lê a Competencia pelo NumIntDoc
            lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
            If lErro <> SUCESSO And lErro <> 134336 Then gError 139074
            
            bTemMaquinasAptas = False
        
        End If
        
        If Not bTemMaquinasAptas Then
            gError 136372
        End If
        
        'Aloca as máquinas para o Plano Operacional
        lErro = PlanoOperacional_Calcula_Tempos4(objPMP, objPMPItem, objPO, colCTMaquinas, colTaxaDeProducao, objCT, iFlag)
        If lErro <> SUCESSO Then gError 137702
    
    Else
    
        'Não aloca as máquinas para o Plano Operacional (MRP Puro)
        lErro = PlanoOperacional_Calcula_Tempos5(objPMPItem, objPO, colCTMaquinas, colTaxaDeProducao, objCT, iFlag)
        If lErro <> SUCESSO Then gError 137703
    
    End If
    
    PlanoOperacional_Calcula_Tempos2 = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Calcula_Tempos2:
 
    PlanoOperacional_Calcula_Tempos2 = gErr
 
    Select Case gErr
    
        Case 136372
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_SEM_TAXA", gErr, objCompetencias.sNomeReduzido, Trim(objPO.sProduto) & SEPARADOR & objProduto.sNomeReduzido, objPO.iNivel, objPO.iSeq, objPO.sCodOPOrigem)
     
        Case 137702, 137703, 139073, 139074
            'erro tratado nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152336)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function PlanoOperacional_Calcula_Tempos3(ByVal objPO As ClassPlanoOperacional, ByVal objTaxaDeProducao As ClassTaxaDeProducao, iQtdTotalMaquinas As Integer, dTempoTotalOperacao As Double) As Long
'De acordo com o PlanoOperacional e a Taxa de produção, retorna o tempo total que
'será gasto na produção
 
Dim lErro As Long
Dim dFatorConversao As Double
Dim dFatorTempo As Double
Dim objProduto As New ClassProduto
Dim iFatorLote As Integer
 
On Error GoTo Erro_PlanoOperacional_Calcula_Tempos3
   
    'Descobre o fator de conversao da UM da necessidade informada p/UM para UM hora
    lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, objTaxaDeProducao.sUMTempo, "h", dFatorTempo)
    If lErro <> SUCESSO Then gError 136376
 
    'Se a composição não varia pela quantidade do produto
    If objTaxaDeProducao.iTipo <> ITEM_TIPO_TAXAPRODUCAO_VARIAVEL Then
        
        iFatorLote = 1
        
        If objTaxaDeProducao.dLoteMax <> 0 Then
            
            'Exemplo: Se tiver de produzir 20 e o Lote Máximo de é 15
            'Precisará de 2 Lotes ... de 10, logo o FatorLote = 2
            iFatorLote = Round((objPO.dQuantidade / objTaxaDeProducao.dLoteMax) + 0.49)
        End If
        
        'Devolve o número máximo de máquinas que ele pode usar com esse tamanho de lote
        iQtdTotalMaquinas = iFatorLote
        
        'Tempo Total de Operacao = Tempo de Operacao em Horas
        dTempoTotalOperacao = (objTaxaDeProducao.dTempoOperacao * dFatorTempo) * iFatorLote
    
    Else
 
        objProduto.sCodigo = objTaxaDeProducao.sProduto
        
        'Le o produto para obter a ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 136377
        
        If lErro <> SUCESSO Then gError 136378
 
        'Descobre o fator de conversao da UM da necessidade informada p/UM para UM hora
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objTaxaDeProducao.sUMProduto, objPO.sUM, dFatorConversao)
        If lErro <> SUCESSO Then gError 136379
 
        'Tempo Total de Operacao = Taxa producao * Quantidade a ser produzida
        'Taxa Producao = Quantidade do produto na Um a ser fabricada/ Tempo Operação em horas
        dTempoTotalOperacao = ((objTaxaDeProducao.dTempoOperacao * dFatorTempo) / (objTaxaDeProducao.dQuantidade * dFatorConversao)) * objPO.dQuantidade
    
    End If
    
    If gobjEST.iTemRepeticoesOper = MARCADO Then
        'Multiplica pelo número de repetições das operações
        dTempoTotalOperacao = dTempoTotalOperacao * objPO.iNumRepeticoes
    End If
           
    PlanoOperacional_Calcula_Tempos3 = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Calcula_Tempos3:
 
    PlanoOperacional_Calcula_Tempos3 = gErr
 
    Select Case gErr
    
        Case 136376, 136377, 136379
        
        Case 136378
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152337)
 
    End Select
 
    Exit Function
 
End Function

Public Function PlanoOperacional_Calcula_Tempos4(ByVal objPMP As ClassPMP, ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal colCTMaquinas As Collection, ByVal colTaxaDeProducao As Collection, ByVal objCT As ClassCentrodeTrabalho, ByVal iFlag As Integer) As Long
'Faz a alocação do uso das máquinas no tempo que vai ser gasto
 
Dim lErro As Long
Dim objCTMaquinas As ClassCTMaquinas
Dim objTaxa As ClassTaxaDeProducao
Dim iIndice As Integer
Dim dTempoTotalOperacao As Double
Dim objPOMaquinas As ClassPOMaquinas
Dim objPOMaquinasAux As ClassPOMaquinas
Dim dTempoDescarga As Double
Dim dTempoMovimentacao As Double
Dim dTempoPreparacao As Double
Dim dQtdProduzida As Double
Dim dtData As Date
Dim iQtdMaquinas As Integer
Dim dQtdHorasDia As Double
Dim iQtdTotalMaquinas As Integer
Dim dQtdTotalHorasDia As Double
Dim dCPDia As Double
Dim iMaxNumMaquinas As Integer
Dim dTempoQueFalta As Double
Dim dTempoDia As Double
Dim colPOMaquinas As New Collection
Dim dAcertoMedio As Double
Dim dResto As Double
Dim iQtdMaq As Integer
Dim dQtdHorasDiaUso As Double
Dim iFator As Integer
Dim iQtdMaquinas2 As Integer
Dim iCont As Integer
Dim iQtdMaqPodeUsarDia As Integer
Dim dHorasPodeUsarDia As Double
Dim dTempoFatorPreparacao As Double
Dim dTempoFatorMovimentacao As Double
Dim dTempoFatorDescarga As Double
Dim bAcerto As Boolean
Dim objMaquina As ClassMaquinas
Dim objMaquinaOperador As ClassMaquinaOperadores
Dim bAchou As Boolean
Dim objCTOperador As ClassCTOperadores
Dim dHorasNecessariaOper As Double
Dim dHorasDisponivelOper As Double
Dim dQtdNecessariaOper As Double
Dim dQtdHorasDiaUsoOper As Double
Dim dPercAux As Double
Dim colMaquinas As New Collection
Dim dQtdHorasPassadas As Double
 
On Error GoTo Erro_PlanoOperacional_Calcula_Tempos4
       
    bAcerto = False
       
    'AQUI SE TEM TODAS AS MÁQUINAS QUE TEM COMPETÊNCIA PARA FAZER DETERMINADA TAREFA
    'É NECESSÁRIO QUE DIA A DIA DE USO SE BUSQUE A QUANTIDADE REALMENTE DISPONIVEL E
    'A QUANTIDADE DE HORAS
    
    If iFlag = MRP_ACERTA_POR_DATA_FIM Then
        dtData = objPO.dtDataFim
        iFator = -1
    Else
        dtData = objPO.dtDataInicio
        iFator = 1
    End If
    
    'Cria uma coleção de máquinas para evitar leituras repetidas
    For Each objCTMaquinas In colCTMaquinas
    
        bAchou = False
                        
        'Verifica se já existe a máquina na coleção
        For Each objMaquina In colMaquinas
            If objCTMaquinas.lNumIntDocMaq = objMaquina.lNumIntDoc Then
                bAchou = True
                Exit For
            End If
        Next
    
        'Se não encontro lê
        If Not bAchou Then
        
            Set objMaquina = New ClassMaquinas
        
            objMaquina.lNumIntDoc = objCTMaquinas.lNumIntDocMaq
            
            'Le a Máquina
            lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
            If lErro <> SUCESSO And lErro <> 106353 Then gError 138513
            
            lErro = CF("Maquinas_Le_Itens", objMaquina)
            If lErro <> SUCESSO Then gError 138516
            
            colMaquinas.Add objMaquina
        
        End If
    
    Next
    
    'Enquanto não puder ter produzido tudo ou ainda faltar outros tempos (Preparação, Movimentação, etc ..), continua
    Do While (Abs(dQtdProduzida - objPO.dQuantidade) > QTDE_ESTOQUE_DELTA) Or (dTempoQueFalta > QTDE_ESTOQUE_DELTA)
   
        iCont = iCont + 1
        
        dTempoFatorPreparacao = 0
        dTempoFatorMovimentacao = 0
        dTempoFatorDescarga = 0
        
        'Se a produção está se estendendo muito dá erro para evitar Loop Infinito
        If iCont > NUM_MAX_DIAS_PRODUCAO Then
            gError 136815
        End If
   
        dCPDia = 0
        iQtdTotalMaquinas = 0
        dQtdTotalHorasDia = 0
        iIndice = 0
    
        'Para cada máquina apta busca a quantidade de máquinas e a disponibilidade do dia
        For Each objCTMaquinas In colCTMaquinas
        
            iIndice = iIndice + 1
        
            Set objTaxa = colTaxaDeProducao.Item(iIndice)
            
            'Busca o que há de disponível para o dia passado, retornando a quantidade de máquinas
            'e a quantidades de horas disponível daquelas máquinas, naquele dia
            lErro = Verifica_Maquinas_Disponiveis_Dia(objCTMaquinas, dtData, objPO.iFilialEmpresa, iQtdMaquinas, dQtdHorasDia)
            If lErro <> SUCESSO Then gError 136373
            
            dQtdHorasPassadas = 0
            lErro = Verifica_Tempo_Decorrido_Dia(objPMPItem, objPO, dtData, dQtdHorasPassadas)
            If lErro <> SUCESSO Then gError 136373
            
            'Busca o que já está se usando no dia
            lErro = Verifica_Uso_Maquina_Dia(objPMP, objPMPItem, objPO, objCTMaquinas, dtData, dQtdHorasDiaUso)
            If lErro <> SUCESSO Then gError 136374
            
            dHorasPodeUsarDia = dQtdHorasDia - (dQtdHorasPassadas * iQtdMaquinas) - dQtdHorasDiaUso
            If dHorasPodeUsarDia < 0 Then dHorasPodeUsarDia = 0
            
'            If dQtdHorasPassadas > dQtdHorasDiaUso Then
'                dHorasPodeUsarDia = dQtdHorasDia - (dQtdHorasPassadas * iQtdMaquinas)
'            Else
'                dHorasPodeUsarDia = dQtdHorasDia - dQtdHorasDiaUso
'            End If
            
            iQtdMaqPodeUsarDia = iQtdMaquinas
           
            'Guarda a quantidade de máquinas que pode ser alterada de acordo com
            'Lote ou Falta de Operadores
            iQtdMaquinas2 = iQtdMaqPodeUsarDia
           
            'Encontra a máquina
            For Each objMaquina In colMaquinas
                If objCTMaquinas.lNumIntDocMaq = objMaquina.lNumIntDoc Then
                    Exit For
                End If
            Next
                    
            'Para cada Operador necessário para o funcionamento da máquina
            For Each objMaquinaOperador In objMaquina.colTipoOperadores
                
                bAchou = False
                
                'Se encontrou no Centro de Trabalho
                For Each objCTOperador In objCT.colOperadores
                    If objCTOperador.iCodTipoMO = objMaquinaOperador.iTipoMaoDeObra Then
                        bAchou = True
                        Exit For
                    End If
                Next
                
                'Se não existe o Operador no CT => Erro
                If Not bAchou Then gError 138515
                
                'Se existem máquinas disponíveis para o dia encontra a disponibilidade do Operador
                'Turno Operador = Turno Máquina ... só é necessário acerto de quantidades
                If iQtdMaquinas > 0 Then
                    dHorasDisponivelOper = (dQtdHorasDia / iQtdMaquinas) * objCTOperador.iQuantidade
                Else
                    dHorasDisponivelOper = 0
                End If
                
                'Acha a quantidade de operadores necessária
                dQtdNecessariaOper = iQtdMaqPodeUsarDia * objMaquinaOperador.iQuantidade * objMaquinaOperador.dPercentualUso
                
                'Se existem máquinas disponíveis acha a quantidade de horas de trabalho para o Operador
                If iQtdMaqPodeUsarDia > 0 Then
                    dHorasNecessariaOper = dQtdNecessariaOper * dHorasPodeUsarDia / iQtdMaqPodeUsarDia
                End If
                                
                'Se o Operador tiver de trabalhar
                If dHorasNecessariaOper <> 0 Then
                
                    'Se eu preciso de mais operadores do que tenho no CT,
                    'Tem que reduzir a quantidade de máquinas e as horas de uso da máquina
                    If objCTOperador.iQuantidade < Fix(dQtdNecessariaOper) Then
                        iQtdMaquinas2 = Fix((objCTOperador.iQuantidade / dQtdNecessariaOper) * iQtdMaquinas2)
                        If iQtdMaquinas2 = 0 Then iQtdMaquinas2 = 1 '??? Jones 19/07/10
                    End If
                
                    'Verifica o que já está sendo usado por essa mão de obra
                    lErro = Verifica_Operador_Maquina_Dia(objPMPItem, objPO, objCTMaquinas, objMaquina, objMaquinaOperador, colMaquinas, dtData, dQtdHorasDiaUsoOper)
                    If lErro <> SUCESSO Then gError 138514
                    
                    'Se o que o tempo disponível menos o tempo em uso não for suficiente para utilizar todo tempo disponível da máquina
                    If (dHorasDisponivelOper - dQtdHorasDiaUsoOper - dHorasNecessariaOper) < 0 Then
                    
                        'Acerta a quantidades de horas que se pode usar da maquina
                        dPercAux = (dHorasDisponivelOper - dQtdHorasDiaUsoOper) / dHorasNecessariaOper
                        
                        dHorasPodeUsarDia = dHorasPodeUsarDia * dPercAux
                    
                    End If
                    
                End If
            
            Next
                                                         
            'retorna o tempo total que será gasto na produção (Tempo Operação)
            lErro = PlanoOperacional_Calcula_Tempos3(objPO, objTaxa, iQtdMaquinas2, dTempoTotalOperacao)
            If lErro <> SUCESSO Then gError 136375
             
            'Verifica se foi obrigado a reduzir a quantidade de máquinas por causa
            'do tempo ser fixo (iQtdMaquinas = Número de Lotes)
            If iQtdMaqPodeUsarDia > iQtdMaquinas2 Then
                iQtdMaqPodeUsarDia = iQtdMaquinas2
            End If
            
            If objPO.iNumMaxMaqPorOper < iQtdMaqPodeUsarDia And objPO.iNumMaxMaqPorOper <> 0 Then
                iQtdMaqPodeUsarDia = objPO.iNumMaxMaqPorOper
            End If
            
            'Acha as Horas disponíveis se levando em conta o número de máquinas
            'escolhido
            If iQtdMaquinas <> 0 Then
'                dHorasPodeUsarDia = (dHorasPodeUsarDia / iQtdMaquinas) * iQtdMaqPodeUsarDia
                If iQtdMaqPodeUsarDia <> iQtdMaquinas Then
                    'Se o tempos que me resta do dia é inferioa a quantidade total de horas reduz para o tempo que resta do dia
                    If dHorasPodeUsarDia - ((dQtdHorasDia - (dQtdHorasPassadas * iQtdMaquinas)) / iQtdMaqPodeUsarDia) > QTDE_ESTOQUE_DELTA Then
                        dHorasPodeUsarDia = ((dQtdHorasDia - (dQtdHorasPassadas * iQtdMaquinas)) / iQtdMaqPodeUsarDia)
                    End If
                End If
            Else
                dHorasPodeUsarDia = 0
            End If
            
'            'procura para ver se já existia alguma coisa para aquele dia alterado pelo usuário
'            For Each objPOMaquinasAux In objPO.colAlocacaoMaquinas
'
'                'Se for a mesma máquina no mesmo dia
'                If objPOMaquinasAux.dtData = dtData And objPOMaquinasAux.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq Then
'
'                    'Reduz o número de ,máquinas de acordo com o que estava antes
'                    If objPOMaquinasAux.iQuantidade < iQtdMaqPodeUsarDia Then
'                        dHorasPodeUsarDia = (dHorasPodeUsarDia / iQtdMaqPodeUsarDia) * objPOMaquinasAux.iQuantidade
'                        iQtdMaqPodeUsarDia = objPOMaquinasAux.iQuantidade
'                    End If
'                    If objPOMaquinasAux.dHorasMaquina < dHorasPodeUsarDia Then
'                        dHorasPodeUsarDia = objPOMaquinasAux.dHorasMaquina
'                    End If
'
'                End If
'            Next
                                                                       
            objCTMaquinas.iQuantidade = iQtdMaqPodeUsarDia
                                              
            iQtdTotalMaquinas = iQtdTotalMaquinas + iQtdMaqPodeUsarDia
            dQtdTotalHorasDia = dQtdTotalHorasDia + dHorasPodeUsarDia
 
            'Guarda o  número máximo de máquinas usadas para poder calcular o tempo de preparação que foi gasto
            If iQtdTotalMaquinas > iMaxNumMaquinas Then iMaxNumMaquinas = iQtdTotalMaquinas
 
            'Calcula os outros tempos com os fatores para poder achar a média ponderada
            dTempoFatorPreparacao = dTempoFatorPreparacao + (objTaxa.dTempoPreparacao * dHorasPodeUsarDia)
            dTempoFatorMovimentacao = dTempoFatorMovimentacao + (objTaxa.dTempoMovimentacao * dHorasPodeUsarDia)
            dTempoFatorDescarga = dTempoFatorDescarga + (objTaxa.dTempoDescarga * dHorasPodeUsarDia)
                       
            'A Capacidade de produção do dia é a Taxa média * a quantidade horas máquinas
            If dTempoTotalOperacao <> 0 Then
                dCPDia = dCPDia + (objPO.dQuantidade / dTempoTotalOperacao) * (dHorasPodeUsarDia)
            End If
            
            'Adiciona a alocação da máquina
            Set objPOMaquinas = New ClassPOMaquinas
            
            objPOMaquinas.dHorasMaquina = dHorasPodeUsarDia
            objPOMaquinas.dtData = dtData
            objPOMaquinas.iQuantidade = objCTMaquinas.iQuantidade
            objPOMaquinas.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq
            objPOMaquinas.lNumIntDocTxProd = objTaxa.lNumIntDoc
            
            Set objPOMaquinas.objTaxaProducao = objTaxa
            Set objPOMaquinas.objMaquina = objMaquina
            
            colPOMaquinas.Add objPOMaquinas
            
            If objPO.iNumMaxMaqPorOper <= iQtdTotalMaquinas And objPO.iNumMaxMaqPorOper <> 0 Then Exit For
            
        Next
       
        'Se com o que se pode produzir no dia + o que já foi produzido
        'for menor que o que se tem que produzir, então continua decrementando o dia, senão
        'acaba a produção relativa ao tempo de operação
        If objPO.dQuantidade - (dCPDia + dQtdProduzida) > QTDE_ESTOQUE_DELTA Then
            dQtdProduzida = dQtdProduzida + dCPDia
            dtData = DateAdd("d", iFator, dtData)
        Else
            
            'Média ponderada dos tempos das máquinas
            If dQtdTotalHorasDia <> 0 Then
                dTempoPreparacao = (dTempoFatorPreparacao / dQtdTotalHorasDia)
                dTempoMovimentacao = (dTempoFatorMovimentacao / dQtdTotalHorasDia)
                dTempoDescarga = (dTempoFatorDescarga / dQtdTotalHorasDia)
            End If
            
            If dQtdTotalHorasDia <> 0 And dCPDia <> 0 Then
                dTempoDia = dQtdTotalHorasDia - ((objPO.dQuantidade - dQtdProduzida) / (dCPDia / dQtdTotalHorasDia))
            
                dQtdProduzida = objPO.dQuantidade
                
                'Aqui o tempo de Operação já foi todo distribuido pelos dias
                'Agora falta os outros tempos
                If bAcerto Then
                    dTempoQueFalta = dTempoQueFalta - dTempoDia
                Else
                    dTempoQueFalta = (dTempoPreparacao * iMaxNumMaquinas) + dTempoMovimentacao + dTempoDescarga - dTempoDia
                    bAcerto = True
                End If
                
            End If
           
            'Se ainda estiver sobrando tempo tem que retirar do uso das máquinas
            If dTempoQueFalta > QTDE_ESTOQUE_DELTA Then
                dtData = DateAdd("d", iFator, dtData)
            Else
                'Se acabou, acerta a quantidade de Horas gastas no primeiro dia
                
                dResto = -dTempoQueFalta
                iQtdMaq = iQtdTotalMaquinas
                
                'Enquanto tiver sobra de tempo
                Do While dResto > QTDE_ESTOQUE_DELTA And iQtdMaq > 0
                
                    'Calcula quando vai ser a média de acordo com a
                    'quantidade de máquinas
                    dAcertoMedio = dResto / iQtdMaq
                    dResto = 0
                
                    'Para cada uso de máquina
                    For Each objPOMaquinas In colPOMaquinas
                    
                        'Se é no primeiro dia
                        If objPOMaquinas.dtData = dtData Then
                        
                            'Se está usando mais horas que o acerto, diminui o uso
                            If objPOMaquinas.dHorasMaquina >= (dAcertoMedio * objPOMaquinas.iQuantidade) Then
                            
                                objPOMaquinas.dHorasMaquina = objPOMaquinas.dHorasMaquina - (dAcertoMedio * objPOMaquinas.iQuantidade)
                            
                            Else
                            
                                If objPOMaquinas.dHorasMaquina <> 0 Then
                                    iQtdMaq = iQtdMaq - objPOMaquinas.iQuantidade
                                    dResto = dResto + (dAcertoMedio * objPOMaquinas.iQuantidade) - objPOMaquinas.dHorasMaquina
                                    objPOMaquinas.dHorasMaquina = 0
                                End If
                            
                            End If
                        
                        End If
                        
                    Next
            
                Loop
                
            End If
            
        End If
        
    Loop
    
    If iFlag = MRP_ACERTA_POR_DATA_FIM Then
        objPO.dtDataInicio = dtData
    Else
        objPO.dtDataFim = dtData
    End If
        
    'Limpa a coleção anterior
    Set objPO.colAlocacaoMaquinas = New Collection
        
    'Para cada uso de máquina
    For Each objPOMaquinas In colPOMaquinas
    
        'Se foi usada no dia
        If objPOMaquinas.dHorasMaquina > QTDE_ESTOQUE_DELTA Then
        
            objPO.colAlocacaoMaquinas.Add objPOMaquinas
        
        End If
        
    Next
           
    PlanoOperacional_Calcula_Tempos4 = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Calcula_Tempos4:
 
    PlanoOperacional_Calcula_Tempos4 = gErr
 
    Select Case gErr
    
        Case 136373 To 136375, 138513, 138514
        
        Case 136815
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAX_DIAS_PRODUCAO", gErr, NUM_MAX_DIAS_PRODUCAO)
 
        Case 138515
            Call Rotina_Erro(vbOKOnly, "ERRO_CT_SEM_OPERADOR", gErr, objCT.lCodigo, objCT.sNomeReduzido, objMaquinaOperador.iTipoMaoDeObra, objMaquina.iCodigo, objMaquina.sNomeReduzido)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152338)
 
    End Select
 
    Exit Function
 
End Function

Public Function PlanoOperacional_Calcula_Tempos5(ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal colCTMaquinas As Collection, ByVal colTaxaDeProducao As Collection, ByVal objCT As ClassCentrodeTrabalho, ByVal iFlag As Integer) As Long
'Calcula o tempo de duração de uma etapa apenas com base no que é cadastrado em produtos
 
Dim lErro As Long
Dim dtData As Date
Dim iFator As Integer
Dim objProduto As ClassProduto
Dim iDiasUteis As Integer
Dim iCont As Integer
Dim dQtdHorasDia As Double
Dim dQtdTotalHorasDia As Double
Dim objCTAux As ClassCentrodeTrabalho
 
On Error GoTo Erro_PlanoOperacional_Calcula_Tempos5

    If iFlag = MRP_ACERTA_POR_DATA_FIM Then
        dtData = objPO.dtDataFim
        iFator = -1
    Else
        dtData = objPO.dtDataInicio
        iFator = 1
    End If
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objPO.sProduto
    
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137116
    
    If lErro <> SUCESSO Then gError 137117
    
    iCont = 0
    iDiasUteis = 0
    
    'Enquanto não tiver tempo para produzir, avança na data
    Do While objProduto.iTempoProducao > iDiasUteis
    
        iCont = iCont + 1

        If iCont > NUM_MAX_DIAS_PRODUCAO Then gError 136815
        
        Set objCTAux = New ClassCentrodeTrabalho

        objCTAux.lNumIntDoc = objCT.lNumIntDoc
        
        'Verifica se o CT está disponivel nesta Data e a quantidade de horas
        lErro = Verifica_CT_Disponiveis_Dia(objCTAux, dtData, dQtdHorasDia)
        If lErro <> SUCESSO Then gError 136373
        
        If dQtdHorasDia > 0 Then
            iDiasUteis = iDiasUteis + 1
            dQtdTotalHorasDia = dQtdTotalHorasDia + dQtdHorasDia
        End If
        
        dtData = DateAdd("d", iFator, dtData)
            
    Loop
    
    If iFlag = MRP_ACERTA_POR_DATA_FIM Then
        objPO.dtDataInicio = dtData
    Else
        objPO.dtDataFim = dtData
    End If
    
    objPO.dTempoGasto = dQtdTotalHorasDia
    
    PlanoOperacional_Calcula_Tempos5 = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Calcula_Tempos5:
 
    PlanoOperacional_Calcula_Tempos5 = gErr
 
    Select Case gErr
    
        Case 136815
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAX_DIAS_PRODUCAO", gErr, NUM_MAX_DIAS_PRODUCAO)
        
        Case 137116, 136373
        
        Case 137117
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152339)
 
    End Select
 
    Exit Function
 
End Function

Public Function Verifica_Uso_Maquina_Dia(ByVal objPMP As ClassPMP, ByVal objPMPItens As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal objCTMaquinas As ClassCTMaquinas, ByVal dtData As Date, dQtdHorasDiaUso As Double) As Long
'Verifica o que já está sendo usada daquela máquina naquele dia para esse item (Capacidade Infinita)
 
Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim objPOAux As ClassPlanoOperacional
Dim bEhDescendente As Boolean
Dim objPMPItemAux As ClassPMPItens
 
On Error GoTo Erro_Verifica_Uso_Maquina_Dia
 
    dQtdHorasDiaUso = 0
 
    For Each objPOAux In objPMPItens.colPO
    
        If objPOAux.iNivel < objPO.iNivel Or (objPOAux.iNivel = objPO.iNivel And objPOAux.iSeq < objPO.iSeq) Then
    
            For Each objPOMaquinas In objPOAux.colAlocacaoMaquinas
        
                If objPOAux.lNumIntDocCT = objCTMaquinas.lNumIntDocCT And objPOMaquinas.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq And objPOMaquinas.dtData = dtData Then
                
                    'Não considera ascendentes pos o tempo decorrido do dia dela já está sendo abatido do
                    'tempo disponível independente de qual recurso está sendo usado
                    bEhDescendente = False
                    Call Verifica_Descendencia(objPMPItens, objPO, objPOAux.lNumIntDoc, bEhDescendente)
                
                    If Not bEhDescendente Then dQtdHorasDiaUso = dQtdHorasDiaUso + objPOMaquinas.dHorasMaquina
                
                End If
    
            Next
            
        End If
 
    Next
    
    For Each objPMPItemAux In objPMP.colItens
        For Each objPOAux In objPMPItemAux.colPO
            For Each objPOMaquinas In objPOAux.colAlocacaoMaquinas
                If objPOAux.lNumIntDocCT = objCTMaquinas.lNumIntDocCT And objPOMaquinas.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq And objPOMaquinas.dtData = dtData Then
                    dQtdHorasDiaUso = dQtdHorasDiaUso + objPOMaquinas.dHorasMaquina
                End If
            Next
        Next
    Next
              
    Verifica_Uso_Maquina_Dia = SUCESSO
 
    Exit Function
 
Erro_Verifica_Uso_Maquina_Dia:
 
    Verifica_Uso_Maquina_Dia = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152340)
 
    End Select
 
    Exit Function
 
End Function

Public Function Verifica_Descendencia(ByVal objPMPItens As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal lNumIntDocPO As Long, bEhDescendente As Boolean) As Long
 
Dim lErro As Long
Dim objPOPai As ClassPlanoOperacional
 
On Error GoTo Erro_Verifica_Descendencia

    If bEhDescendente Then Exit Function

    If objPO.lNumIntDocPOPai <> 0 Then
    
        For Each objPOPai In objPMPItens.colPO
        
            If objPOPai.lNumIntDoc = objPO.lNumIntDocPOPai Then
        
                If lNumIntDocPO = objPOPai.lNumIntDoc Then
                    bEhDescendente = True
                Else
                    Call Verifica_Descendencia(objPMPItens, objPOPai, lNumIntDocPO, bEhDescendente)
                End If
                
                Exit For
                
            End If
        
        Next
    
    End If
              
    Verifica_Descendencia = SUCESSO
 
    Exit Function
 
Erro_Verifica_Descendencia:
 
    Verifica_Descendencia = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152340)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function Verifica_Maquinas_Disponiveis_Dia(ByVal objCTMaquinas As ClassCTMaquinas, ByVal dtData As Date, ByVal iFilialEmpresa As Integer, iQtdMaquinas As Integer, dQtdHorasDia As Double) As Long
'Obtém quanto de tempo a máquina trabalha em um determinado dia
 
Dim lErro As Long
Dim objFeriado As New ClassFeriado
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMaquinasAux As New ClassCTMaquinas
Dim objTurno As New ClassTurno
Dim dHoras As Double
Dim iMaquinas As Integer
Dim objCTMaqProgDisp As New ClassCTMaqProgDisp
Dim objCTMaqProgTurno As New ClassCTMaqProgTurno
Dim objCTMaqProgTurnoItens As ClassCTMaqProgTurnoItens
 
On Error GoTo Erro_Verifica_Maquinas_Disponiveis_Dia
 
    dQtdHorasDia = 0
    iQtdMaquinas = 0
    
    objFeriado.dtData = dtData
    objFeriado.iFilialEmpresa = iFilialEmpresa
 
    lErro = CF("Feriado_Le", objFeriado)
    If lErro <> SUCESSO And lErro <> 43379 Then gError 136817
 
    'Se não é feriado
    If lErro <> SUCESSO Then
 
        objCT.lNumIntDoc = objCTMaquinas.lNumIntDocCT
 
        'Le o CT e todos objs descendentes
        lErro = CF("CentroDeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137212 Then gError 136818
 
        'Encontra a máquina em questão
        For Each objCTMaquinasAux In objCT.colMaquinas
 
            If objCTMaquinasAux.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq Then
                iMaquinas = objCTMaquinasAux.iQuantidade
                Exit For
            End If
 
        Next
 
        dHoras = 0
 
        'Para cada turno dessa máquina
        For Each objTurno In objCTMaquinasAux.colTurnos
            dHoras = dHoras + objTurno.dQtdHoras(Weekday(dtData))
        Next
 
        objCTMaqProgDisp.lNumIntDocCT = objCTMaquinasAux.lNumIntDocCT
        objCTMaqProgDisp.lNumIntDocMaq = objCTMaquinasAux.lNumIntDocMaq
        objCTMaqProgDisp.dtData = dtData
 
        'Busca por alterações programadas na disponibilidade
        lErro = CF("CTMaquinaProgDisponibilidade_Le", objCTMaqProgDisp)
        If lErro <> SUCESSO And lErro <> 136564 Then gError 136819
 
        'Se encontrou
        If lErro = SUCESSO Then
            iMaquinas = objCTMaqProgDisp.iQuantidade
        End If
 
        objCTMaqProgTurno.lNumIntDocCT = objCTMaquinasAux.lNumIntDocCT
        objCTMaqProgTurno.lNumIntDocMaq = objCTMaquinasAux.lNumIntDocMaq
        objCTMaqProgTurno.dtData = dtData
 
        'Busca por alterações programadas nos turnos
        lErro = CF("CTMaquinaProgTurno_Le", objCTMaqProgTurno)
        If lErro <> SUCESSO And lErro <> 136704 Then gError 136820
 
        'Se encontrou
        If lErro = SUCESSO Then
 
            dHoras = 0
 
            For Each objCTMaqProgTurnoItens In objCTMaqProgTurno.colTurnos
                dHoras = dHoras + objCTMaqProgTurnoItens.dHoras
            Next
 
        End If
 
    End If
    
    If dHoras = 0 Then iMaquinas = 0
           
    dQtdHorasDia = dHoras * iMaquinas
    iQtdMaquinas = iMaquinas
           
    Verifica_Maquinas_Disponiveis_Dia = SUCESSO
 
    Exit Function
 
Erro_Verifica_Maquinas_Disponiveis_Dia:
 
    Verifica_Maquinas_Disponiveis_Dia = gErr
 
    Select Case gErr
    
        Case 136617 To 136620
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152341)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function Verifica_CT_Disponiveis_Dia(ByVal objCT As ClassCentrodeTrabalho, ByVal dtData As Date, dQtdHorasDia As Double) As Long
'Verifica quantas horas o CT está disponível em uma data
 
Dim lErro As Long
Dim objFeriado As New ClassFeriado
Dim objTurno As New ClassTurno
Dim dHoras As Double
Dim objCTAux As New ClassCentrodeTrabalho
 
On Error GoTo Erro_Verifica_CT_Disponiveis_Dia
 
    dQtdHorasDia = 0
    
    objFeriado.dtData = dtData
    objFeriado.iFilialEmpresa = giFilialEmpresa
 
    lErro = CF("Feriado_Le", objFeriado)
    If lErro <> SUCESSO And lErro <> 43379 Then gError 137708
 
    'Se não é feriado
    If lErro <> SUCESSO Then
    
        objCTAux.lCodigo = objCT.lCodigo
        objCTAux.iFilialEmpresa = objCT.iFilialEmpresa
        objCTAux.sNomeReduzido = objCT.sNomeReduzido
        objCTAux.lNumIntDoc = objCT.lNumIntDoc
    
        'Lê o CTTurnos
        lErro = CF("CentrodeTrabalho_Le_CTTurnos", objCTAux)
        If lErro <> SUCESSO And lErro <> 137207 Then gError 137709
    
        dHoras = 0
 
        'Para cada turno desse CT
        For Each objTurno In objCTAux.colTurnos
            dHoras = dHoras + objTurno.dQtdHoras(Weekday(dtData))
        Next
 
    End If
           
    dQtdHorasDia = dHoras
           
    Verifica_CT_Disponiveis_Dia = SUCESSO
 
    Exit Function
 
Erro_Verifica_CT_Disponiveis_Dia:
 
    Verifica_CT_Disponiveis_Dia = gErr
 
    Select Case gErr
    
        Case 137708, 137709
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152342)
 
    End Select
 
    Exit Function
 
End Function
 
Public Function PlanoOperacional_AcertaDatas(ByVal objPMP As ClassPMP, ByVal objPMPItem As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal objCT As ClassCentrodeTrabalho, ByVal iFlag As Integer) As Long
'Acerta a data das etapas descendentes a etapa passada
 
Dim lErro As Long
Dim iIndice As Integer
Dim objPOAux As ClassPlanoOperacional
Dim colTaxas As New Collection
Dim colMaquinas As New Collection
Dim objCTMaquina As ClassCTMaquinas
Dim objPOMaquina As ClassPOMaquinas
Dim colAux As New Collection
Dim objOPOperacoes As ClassOrdemProducaoOperacoes
Dim objItemOP As ClassItemOP
Dim objOPInsumos As ClassOrdemProducaoInsumos
Dim objOP As ClassOrdemDeProducao
Dim objPOPai As New ClassPlanoOperacional
Dim lProxPO As Long
Dim colPO As New Collection
Dim objOPOperacoesPai As ClassOrdemProducaoOperacoes
Dim objPMPItemAux As ClassPMPItens
Dim dtDataAnterior As Date
Dim colOperacoes As New Collection
Dim bAchou  As Boolean
Dim lNumIntDocAnterior As Long
Dim objOPOperacaoAux As ClassOrdemProducaoOperacoes
Dim objPOFilha As ClassPlanoOperacional
Dim objPMPAux As ClassPMP

On Error GoTo Erro_PlanoOperacional_AcertaDatas
 
    Set objOPOperacoes = New ClassOrdemProducaoOperacoes
    
    'Para cada máquina alocada no plano operacional, Insere como apta
    For Each objPOMaquina In objPO.colAlocacaoMaquinas
 
        For Each objCTMaquina In colMaquinas
        
            bAchou = False
            If objPOMaquina.lNumIntDocMaq = objCTMaquina.lNumIntDocMaq Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
        
            Set objCTMaquina = New ClassCTMaquinas
    
            objCTMaquina.lNumIntDocCT = objPO.lNumIntDocCT
            objCTMaquina.lNumIntDocMaq = objPOMaquina.lNumIntDocMaq
    
            colMaquinas.Add objCTMaquina
            colTaxas.Add objPOMaquina.objTaxaProducao
            
        End If
 
    Next
  
    For Each objOPOperacoes In objPMPItem.objItemOP.colOrdemProducaoOperacoes
 
        If objOPOperacoes.lNumIntDoc = objPO.lNumIntDocOper Then
            Exit For
        End If
 
    Next
 
    'Se não mudou o CT recalcula a data final
    If objOPOperacoes.lNumIntDocCT = objPO.lNumIntDocCT Then

        Set objPMPAux = New ClassPMP
        
        For iIndice = 1 To objPMP.colItens.Count
            'Adiciona todos os outros itens da produção exceto o que está sendo calculado no momento
            'e posteriores
            If objPMPItem.sCodOPOrigem <> objPMP.colItens.Item(iIndice).sCodOPOrigem Or objPMPItem.sProduto <> objPMP.colItens.Item(iIndice).sProduto Then
                objPMPAux.colItens.Add objPMP.colItens.Item(iIndice)
                Exit For
            End If
        Next
    
        'Recalcula o Tempo de início e fim da etapa
        lErro = PlanoOperacional_Calcula_Tempos2(objPMPAux, objPMPItem, objPO, colMaquinas, colTaxas, objCT, iFlag, objOPOperacoes)
        If lErro <> SUCESSO Then gError 136380
    End If
 
    'Pega o Item de OP de origem
    Set objItemOP = objPMPItem.objItemOP
 
    'Se a Etapa não é a Ancestral (Roteiro Nivel 0 / OP principal)
    If objPO.sCodOPOrigem <> objPMPItem.sCodOPOrigem Or objPO.lNumIntDocPOPai <> 0 Then
 
        'Acha a etapa pai na coleção
        For Each objPOPai In objPMPItem.colPO
            If objPOPai.lNumIntDoc = objPO.lNumIntDocPOPai Then
                Exit For
            End If
        Next
 
        'Acha a Operação de Origem da etapa Pai
        For Each objOPOperacoesPai In objItemOP.colOrdemProducaoOperacoes
            If objOPOperacoesPai.lNumIntDoc = objPOPai.lNumIntDocOper Then
                Exit For
            End If
        Next
 
        'Verifica se é do mesmo roteiro
        If objOPOperacoesPai.iNivelRoteiro = objOPOperacoes.iNivelRoteiro And objOPOperacoesPai.iSeqRoteiro = objOPOperacoes.iSeqRoteiro Then
                                                    
            'Acha a primeira Operação do Roteiro
            For Each objOPOperacaoAux In objItemOP.colOrdemProducaoOperacoes
             
                 If objOPOperacoes.iNivelRoteiro = objOPOperacaoAux.iNivelRoteiro And objOPOperacoes.iSeqRoteiro = objOPOperacaoAux.iSeqRoteiro Then
                    Exit For
                 End If
                 
            Next
            
            'Acha a Operação Pai da Primeira Operação do roteiro
            For Each objOPOperacoesPai In objItemOP.colOrdemProducaoOperacoes
             
                 If objOPOperacoesPai.iSeq = objOPOperacaoAux.iSeqPai Then
                    Exit For
                 End If
            
            Next
        
        End If
                
        'Preenche a coleção de operações com os filhos de mesmo roteiro
        For Each objOPOperacaoAux In objItemOP.colOrdemProducaoOperacoes
             If objOPOperacoes.iNivelRoteiro = objOPOperacaoAux.iNivelRoteiro And objOPOperacoes.iSeqRoteiro = objOPOperacaoAux.iSeqRoteiro Then
                colOperacoes.Add objOPOperacaoAux
             End If
        Next
 
        'Obtem a OP associada a etapa pai
        Set objOP = objPOPai.objOP
 
        'Se não é o roteiro do produto acabado
        If Not (objOPOperacoesPai Is Nothing) Then
            'Acha o Insumo que originou a OP
            For Each objOPInsumos In objOPOperacoesPai.colOPInsumos
                If objOPInsumos.sProduto = objPO.sProduto Then
                    Exit For
                End If
            Next
        Else
        
           Set objOPInsumos = New ClassOrdemProducaoInsumos
           objOPInsumos.sProduto = objPMPItem.sProduto
           objOPInsumos.sVersaoKitComp = objPMPItem.sVersao
           objOPInsumos.dQuantidade = objPMPItem.dQuantidade
           objOPInsumos.sUMProduto = objPMPItem.sUM
        End If
 
        'Remove os as etapas alocadas posteriromente para simular o efeito do algoritmo
        lErro = PlanoOperacional_Remove_Descendentes(objPMPItem.colPO, colPO, objPO)
        If lErro <> SUCESSO Then gError 136382
 
    Else
 
        'Inicializa os objetos do zero
        Set objOPInsumos = New ClassOrdemProducaoInsumos
        Set objPOPai = New ClassPlanoOperacional
 
        Set objOP = objPMPItem.objOP
 
        objOPInsumos.sProduto = objPMPItem.sProduto
        objOPInsumos.sVersaoKitComp = objPMPItem.sVersao
        objOPInsumos.dQuantidade = objPMPItem.dQuantidade
        objOPInsumos.sUMProduto = objPMPItem.sUM
 
        objPOPai.dQuantidade = objPMPItem.dQuantidade
        objPOPai.sUM = objPMPItem.sUM
        objPOPai.sProduto = objPMPItem.sProduto
        objPOPai.sVersao = objPMPItem.sVersao
        objPOPai.iFilialEmpresa = objPMPItem.iFilialEmpresa
        objPOPai.dtDataInicio = DATA_NULA
 
        objOPInsumos.lNumIntDocOper = objPMPItem.objItemOP.colOrdemProducaoOperacoes.Item(1).lNumIntDoc
        
        'Preenche a coleção de operações com os filhos de mesmo roteiro
        For Each objOPOperacaoAux In objPMPItem.objItemOP.colOrdemProducaoOperacoes
             If objOPOperacoes.iNivelRoteiro = objOPOperacaoAux.iNivelRoteiro And objOPOperacoes.iSeqRoteiro = objOPOperacaoAux.iSeqRoteiro Then
                colOperacoes.Add objOPOperacaoAux
             End If
        Next
            
    End If
 
    lProxPO = objPMPItem.lUltimoProxPO
    
    'Valida a data de término da etapa filha
    If objPOPai.dtDataInicio <> DATA_NULA Then
        If objPOPai.dtDataInicio < objPO.dtDataFim Then gError 138251
    End If
 
    'Modifica os valores para gerar as etapas
    dtDataAnterior = objPOPai.dtDataInicio
    objPOPai.dtDataInicio = objPO.dtDataFim
 
    'Coloca a coleção de POs no Estado inicial e guarda valor anterior
    For Each objPMPItemAux In objPMP.colItens
        If objPMPItemAux.sCodOPOrigem = objPMPItem.sCodOPOrigem And objPMPItemAux.iFilialEmpresa = objPMPItem.iFilialEmpresa And objPMPItemAux.sProduto = objPMPItem.sProduto Then
            Set colAux = objPMPItemAux.colPO
            Set objPMPItemAux.colPO = New Collection 'É usado para leitura do Estoque
            Exit For
        End If
    Next
    
    Set objPOAux = New ClassPlanoOperacional
    objOPOperacoes.lNumIntDocCT = objPO.lNumIntDocCT
    
    Set objPOAux.colAlocacaoMaquinas = objPO.colAlocacaoMaquinas
    objPOAux.lNumIntDoc = objPO.lNumIntDoc
    objPOAux.dtDataFim = objPO.dtDataFim
    objPOAux.dtDataInicio = objPO.dtDataInicio
    objPOAux.lNumIntDocPOPai = objPO.lNumIntDocPOPai
    
    Set objOPOperacaoAux = New ClassOrdemProducaoOperacoes
    
    lErro = CF("OPOperacoes_Cria_Copia", objOPOperacoes, objOPOperacaoAux)
    If lErro <> SUCESSO Then gError 136381
    
    If objOPOperacoesPai Is Nothing Then
        Set objOPOperacoesPai = New ClassOrdemProducaoOperacoes
                 
        objOPOperacoesPai.iNivelRoteiro = 0
        objOPOperacoesPai.sProduto = objPMPItem.sProduto
        objOPOperacoesPai.sVersao = objPMPItem.sVersao
        objOPOperacoesPai.lNumIntDocItemOP = objPMPItem.objItemOP.lNumIntDoc

    End If
            
'    'Refaz a colPO
    lErro = PlanoOperacional_Gera2(objPMP, objPMPItem, objOPOperacoesPai, objOPInsumos, objPOPai, objPO.objOP, colPO, objOPOperacaoAux, colOperacoes, objPOAux, lProxPO, 0)
    If lErro <> SUCESSO Then gError 136381
    
    iIndice = 0
    
    For Each objOPOperacaoAux In colOperacoes
    
        Set objPOFilha = New ClassPlanoOperacional
        
        iIndice = iIndice + 1
    
        If objOPOperacoes.iSeq < objOPOperacaoAux.iSeq Then
            lErro = PlanoOperacional_Gera2(objPMP, objPMPItem, objOPOperacaoAux, objOPInsumos, objPO, objPO.objOP, colPO, objOPOperacaoAux, colOperacoes, objPOFilha, lProxPO, iIndice)
            If lErro <> SUCESSO Then gError 136381
        End If
    
    Next
 
    'Retorna aos valores iniciais
    objPOPai.dtDataInicio = dtDataAnterior
    Set objPMPItemAux.colPO = colAux
                          
    'Atualiza os descendentes de objPMPItem.colPO
    lErro = PlanoOperacional_Atualiza_Descendentes(objPMPItem.colPO, colPO, objPO)
    If lErro <> SUCESSO Then gError 136382
 
    'Atualiza a coleção do Item do Plano Mestre
    Set objPMPItem.colPO = colPO
 
    'Atualiza valores de próximo
    objPMPItem.lUltimoProxPO = lProxPO
 
    PlanoOperacional_AcertaDatas = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_AcertaDatas:
 
    PlanoOperacional_AcertaDatas = gErr
 
    Select Case gErr
 
        Case 136380, 136381, 136382
 
        Case 138251
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_ETAPA", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152343)
 
    End Select
 
    Exit Function
 
End Function

Public Function RelPreviaCargaCT_Prepara(ByVal objPMPTela As ClassPMP) As Long
'Monta a tabela temporária para o relatório de prévia de carga nos centros de trabalho
 
Dim lErro As Long
Dim objPMPBD As New ClassPMP
Dim colPMPItens As New Collection
Dim objPMPItem As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objPOMaquinas As ClassPOMaquinas
Dim objCT As ClassCentrodeTrabalho
Dim objCTMaquinas As ClassCTMaquinas
Dim dUsoMaquina As Double
Dim lNumIntRel As Long
Dim objCliente As ClassCliente
Dim objProduto As ClassProduto
Dim objMaquina As ClassMaquinas
Dim objRelatorio As New AdmRelatorio
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iSeqRel As Integer
Dim iQtdMaquinas As Integer
Dim dQtdHoras As Double
Dim objMaquinaOper As ClassMaquinaOperadores
Dim dOperDisp As Double
Dim dOperUso As Double
Dim objCTOperador As ClassCTOperadores
Dim objOperador As ClassTiposDeMaodeObra

On Error GoTo Erro_RelPreviaCargaCT_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136535
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136546
    Next
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_PREVIACARGACT", lNumIntRel)
    If lErro <> SUCESSO Then gError 136536
 
    'Le o Plano Mestre que já está gravado
    lErro = CF("PlanoMestreProducao_Le", objPMPBD)
    If lErro <> SUCESSO And lErro <> 136298 Then gError 136537
    
    'Adiciona os itens gravados
    For Each objPMPItem In objPMPBD.colItens
        colPMPItens.Add objPMPItem
    Next
    
    'Adiciona os novos itens
    For Each objPMPItem In objPMPTela.colItens
        colPMPItens.Add objPMPItem
    Next
    
    'Para cada item
    For Each objPMPItem In colPMPItens
    
        Set objCliente = New ClassCliente
 
        objCliente.lCodigo = objPMPItem.lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 136538
    
        'Para cada etapa
        For Each objPO In objPMPItem.colPO
        
            Set objProduto = New ClassProduto
            Set objCT = New ClassCentrodeTrabalho
            
            objProduto.sCodigo = objPO.sProduto
            
            'Le o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 136539
            
            objCT.lNumIntDoc = objPO.lNumIntDocCT
        
            'Le o centro de trabalho
            lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
            If lErro <> SUCESSO And lErro <> 137212 Then gError 136540
            
            'Para cada máquina alocada
            For Each objPOMaquinas In objPO.colAlocacaoMaquinas
            
                Set objMaquina = New ClassMaquinas
                Set objCTMaquinas = New ClassCTMaquinas
                
                objMaquina.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
                
                'Le a Máquina
                lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
                If lErro <> SUCESSO And lErro <> 106353 Then gError 136541
                           
                iSeqRel = iSeqRel + 1
                                
                objCTMaquinas.iQuantidade = objPOMaquinas.iQuantidade
                objCTMaquinas.lNumIntDocCT = objPO.lNumIntDocCT
                objCTMaquinas.lNumIntDocMaq = objPOMaquinas.lNumIntDocMaq
                
                'Verifica a disponibilidade da máquina naquele dia
                lErro = Verifica_Maquinas_Disponiveis_Dia(objCTMaquinas, objPOMaquinas.dtData, objCT.iFilialEmpresa, iQtdMaquinas, dQtdHoras)
                If lErro <> SUCESSO Then gError 136542
            
                'Se não tem cliente considera que é interno
                If objCliente.sNomeReduzido = "" Then objCliente.sNomeReduzido = "INTERNO"
            
                'Insere em RelPreviaCargaCT
                lErro = Comando_Executar(alComando(0), "INSERT INTO RelPreviaCargaCT (NumIntRel,SeqRel,CodigoOPOrigem, CodigoOP, CodCliente, NomeCliente, Produto, DescProduto, Etapa, Nivel, Seq, Ano, Mes, Dia, NumIntDocCT, NomeCT, NumIntDocMaq, NomeMaquina, PercUso, QtdMaquina, QtdTempo, UMTempo, Data) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                lNumIntRel, iSeqRel, objPMPItem.sCodOPOrigem, objPO.sCodOPOrigem, objPMPItem.lCliente, objCliente.sNomeReduzido, objPO.sProduto, objProduto.sDescricao, CStr(objPO.iNivel) & SEPARADOR & CStr(objPO.iSeq), objPO.iNivel, objPO.iSeq, Year(objPOMaquinas.dtData), Month(objPOMaquinas.dtData), Day(objPOMaquinas.dtData), objPO.lNumIntDocCT, objCT.sNomeReduzido, objPOMaquinas.lNumIntDocMaq, objMaquina.sNomeReduzido, objPOMaquinas.dHorasMaquina / dQtdHoras, objPOMaquinas.iQuantidade, objPOMaquinas.dHorasMaquina, "h", objPOMaquinas.dtData)
                If lErro <> SUCESSO Then gError 136543
                
                For Each objMaquinaOper In objPOMaquinas.objMaquina.colTipoOperadores
                
                    iSeqRel = iSeqRel + 1
                    
                    Set objOperador = New ClassTiposDeMaodeObra
                
                    'Obtém a informação do Operador cadastrada no CT
                    For Each objCTOperador In objCT.colOperadores
                        If objCTOperador.iCodTipoMO = objMaquinaOper.iTipoMaoDeObra Then
                            Exit For
                        End If
                    Next
                    
                    dOperDisp = objCTOperador.iQuantidade * (dQtdHoras / iQtdMaquinas)
                    dOperUso = objPOMaquinas.dHorasMaquina * objMaquinaOper.iQuantidade * objMaquinaOper.dPercentualUso
                    
                    objOperador.iCodigo = objMaquinaOper.iTipoMaoDeObra
                    
                    lErro = CF("TiposDeMaodeObra_Le", objOperador)
                    If lErro <> SUCESSO And lErro <> 137597 Then gError 138521
                
                    'Insere em RelPreviaCargaCT
                    lErro = Comando_Executar(alComando(0), "INSERT INTO RelPreviaCargaCT (NumIntRel,SeqRel,CodigoOPOrigem, CodigoOP, CodCliente, NomeCliente, Produto, DescProduto, Etapa, Nivel, Seq, Ano, Mes, Dia, NumIntDocCT, NomeCT, NumIntDocMaq, NomeMaquina, PercUso, QtdMaquina, QtdTempo, UMTempo, Data) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                    lNumIntRel, iSeqRel, objPMPItem.sCodOPOrigem, objPO.sCodOPOrigem, objPMPItem.lCliente, objCliente.sNomeReduzido, objPO.sProduto, objProduto.sDescricao, CStr(objPO.iNivel) & SEPARADOR & CStr(objPO.iSeq), objPO.iNivel, objPO.iSeq, Year(objPOMaquinas.dtData), Month(objPOMaquinas.dtData), Day(objPOMaquinas.dtData), objPO.lNumIntDocCT, objCT.sNomeReduzido, objPOMaquinas.lNumIntDocMaq, objOperador.sDescricao, dOperUso / dOperDisp, objPOMaquinas.iQuantidade, objPOMaquinas.dHorasMaquina, "h", objPOMaquinas.dtData)
                    If lErro <> SUCESSO Then gError 136543
                                
                Next
                       
            Next
        
        Next
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136544
 
    lErro = objRelatorio.ExecutarDireto("Prévia da Carga nos CTs", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 136545
 
    RelPreviaCargaCT_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelPreviaCargaCT_Prepara:
 
    RelPreviaCargaCT_Prepara = gErr
 
    Select Case gErr
 
        Case 136535
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 136536 To 136542, 138521
 
        Case 136543
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELPREVIACARGACT", gErr)
        
        Case 136544
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 136545
        
        Case 136546
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152344)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelErrosMRP_Prepara(ByVal objPMP As ClassPMP) As Long
'Monta a tabela temporária para o relatório de Críticas do MRP
 
Dim lErro As Long
Dim objPMPItem As New ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iSeqRel As Integer
Dim objCliente As ClassCliente
 
On Error GoTo Erro_RelErrosMRP_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136550
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136551
    Next
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_ERROSMRP", lNumIntRel)
    If lErro <> SUCESSO Then gError 136552
   
    'Para cada item
    For Each objPMPItem In objPMP.colItens
    
        Set objCliente = New ClassCliente
 
        objCliente.lCodigo = objPMPItem.lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 136553
    
        For Each objPO In objPMPItem.colPO
        
            If objPO.iStatus <> PO_STATUS_OK Then
            
                iSeqRel = iSeqRel + 1
                
                If objCliente.sNomeReduzido = "" Then objCliente.sNomeReduzido = "INTERNO"
                
                lErro = Comando_Executar(alComando(0), "INSERT INTO RelErrosMRP (NumIntRel,SeqRel,CodigoOPOrigem, CodigoOP, CodCliente, NomeCliente, Etapa, Nivel,Seq, Status, Mensagem ) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
                lNumIntRel, iSeqRel, objPMPItem.sCodOPOrigem, objPO.sCodOPOrigem, objPMPItem.lCliente, objCliente.sNomeReduzido, CStr(objPO.iNivel) & SEPARADOR & CStr(objPO.iSeq), objPO.iNivel, objPO.iSeq, objPO.iStatus, left(objPO.sDescErro, 255))
                If lErro <> SUCESSO Then gError 136554
            
            End If
        
        Next
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136555
 
    lErro = objRelatorio.ExecutarDireto("Erros no MRP", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 136556
 
    RelErrosMRP_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelErrosMRP_Prepara:
 
    RelErrosMRP_Prepara = gErr
 
    Select Case gErr
 
        Case 136550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 136551
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136552, 136553
 
        Case 136554
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELERROSMRP", gErr)
        
        Case 136555
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 136556
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152345)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelCriticasPMP_Prepara(ByVal objPMP As ClassPMP) As Long
'Monta a tabela temporária para o relatório de críticas do Plano Mestre de Produção
 
Dim lErro As Long
Dim lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPMPVazio As New ClassPMP
Dim objPMPItemVazio As New ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objPMPItem As ClassPMPItens
Dim objItemOP As ClassItemOP
Dim objOPOperacao As ClassOrdemProducaoOperacoes
Dim objOPInsumo As ClassOrdemProducaoInsumos
Dim dFatorMultiplicacao As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dEstoqueProjetado As Double
Dim sMensagem  As String
Dim iSeqRel As Integer
Dim objCliente As ClassCliente
Dim objCT As ClassCentrodeTrabalho
Dim objPOMaquinas As ClassPOMaquinas
Dim iMaqQtdDia As Integer
Dim dMaqTempoDia As Double
Dim objCTMaquinas As ClassCTMaquinas
Dim dUsoMaquina As Double
Dim objProdutoFilial As ClassProdutoFilial
Dim colUsoOperadores As Collection
Dim objMaquinaOper As ClassMaquinaOperadores
Dim objPOMaquinasAux As ClassPOMaquinas
 
On Error GoTo Erro_RelCriticasPMP_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 138107
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138108
    Next
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_CRITICASPMP", lNumIntRel)
    If lErro <> SUCESSO Then gError 138109
   
    For Each objPMPItem In objPMP.colItens
    
        If objPMPItem.objItemOP.iSituacao <> ITEMOP_SITUACAO_BAIXADA Then
    
            Set objCliente = New ClassCliente
    
            objCliente.lCodigo = objPMPItem.lCliente
           
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 138119
       
            Set objItemOP = objPMPItem.objItemOP
       
            iIndice = 0
            
            For Each objPO In objPMPItem.colPO
           
                objPO.iStatus = PO_STATUS_OK
                objPO.sDescErro = PO_STATUS_NOME_OK
                
                If objPO.iNivel = KIT_NIVEL_RAIZ Then
                
                    Set objEstoqueProduto = New ClassEstoqueProduto
                    Set objProdutoFilial = New ClassProdutoFilial
                     
                    'Nova Validação
                    objEstoqueProduto.sProduto = objPMPItem.sProduto
                    objEstoqueProduto.iFilialEmpresa = objPMPItem.iFilialEmpresa
                
                    'Le o estoque projetado para o dia
                    lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objPMPItem.dtDataNecessidade, dEstoqueProjetado)
                    If lErro <> SUCESSO Then gError 138113
                
                    'Se a projeção é negativa +> Erro
                    If dEstoqueProjetado < 0 Then
                                   
                        objPO.iStatus = PO_STATUS_FALTAMATERIAL
                        Call MRP_Mensagem6(sMensagem, objPMPItem.sProduto, objPMPItem.dtDataNecessidade)
                        
                    End If
            
                End If
           
                Set objCT = New ClassCentrodeTrabalho
                   
                objCT.lNumIntDoc = objPO.lNumIntDocCT
                  
                'Le o centro de trabalho
                lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
                If lErro <> SUCESSO And lErro <> 134590 Then gError 138120
    
                If objItemOP.dQuantidade <> 0 Then
                    dFatorMultiplicacao = objPO.dQuantidade / objItemOP.dQuantidade
                Else
                    dFatorMultiplicacao = 0
                End If
                
                'Para cada Operação da OP
                For Each objOPOperacao In objItemOP.colOrdemProducaoOperacoes
               
                   'Se ela Originou a etapa
                   If objPO.lNumIntDocOper = objOPOperacao.lNumIntDoc Then
                   
                       'Para cada insumo necessário
                       For Each objOPInsumo In objOPOperacao.colOPInsumos
                       
                            Set objEstoqueProduto = New ClassEstoqueProduto
                           
                            objEstoqueProduto.sProduto = objOPInsumo.sProduto
                            objEstoqueProduto.iFilialEmpresa = objPO.iFilialEmpresa
                           
                            Set objProdutoFilial = New ClassProdutoFilial
                       
                            'Le o estoque projetado para o dia
                            lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objPO.dtDataInicio, dEstoqueProjetado)
                            If lErro <> SUCESSO Then gError 138113
                       
                            'Se a projeção é negativa +> Erro
                            If dEstoqueProjetado < 0 Then
                                          
                                objPO.iStatus = PO_STATUS_FALTAMATERIAL
                                Call MRP_Mensagem6(sMensagem, objOPInsumo.sProduto, objPO.dtDataInicio)
                               
                            End If
                       
                        Next
                   
                    End If
                   
                Next
                   
                'Para cada máquina alocada para o trabalho
                For Each objPOMaquinas In objPO.colAlocacaoMaquinas
                   
                    Set objCTMaquinas = New ClassCTMaquinas
                    
                    objCTMaquinas.lNumIntDocCT = objCT.lNumIntDoc
                    objCTMaquinas.lNumIntDocMaq = objPOMaquinas.lNumIntDocMaq
                
                    'Verifica a disponibilidade da máquina naquele dia
                    lErro = Verifica_Maquinas_Disponiveis_Dia(objCTMaquinas, objPOMaquinas.dtData, objCT.iFilialEmpresa, iMaqQtdDia, dMaqTempoDia)
                    If lErro <> SUCESSO Then gError 138121
               
                    dUsoMaquina = 0
                    
                    Set colUsoOperadores = New Collection
               
                    For Each objMaquinaOper In objPOMaquinas.objMaquina.colTipoOperadores
                       
                        Set objPOMaquinasAux = New ClassPOMaquinas
                       
                        objPOMaquinasAux.lNumIntDocMaq = objMaquinaOper.lNumIntDocMaq
                       
                        colUsoOperadores.Add objPOMaquinasAux
                    Next
           
                    'Valida a carga
                    lErro = MRP_Valida_SobreCarga_Mesmo_PMPItem(objPMPItem, objPO, objPOMaquinas, dUsoMaquina, colUsoOperadores)
                    If lErro <> SUCESSO Then gError 138112
               
                    'Valida a carga
                    lErro = MRP_Valida_SobreCarga_Outro_PMPItem(objPMP, objPMPItem, objPO, objPOMaquinas, dUsoMaquina, colUsoOperadores)
                    If lErro <> SUCESSO Then gError 138112
               
                    'se a capacidade do dia ficou estourada => Muda Status
                    If dMaqTempoDia - dUsoMaquina < 0 Then
                       
                        If objPO.iStatus = PO_STATUS_FALTAMATERIAL Or objPO.iStatus = PO_STATUS_AMBOS Then
                            objPO.iStatus = PO_STATUS_AMBOS
                        Else
                            objPO.iStatus = PO_STATUS_SOBRECARGA
                        End If
                   
                        Call MRP_Mensagem2(sMensagem, objPO.iNivel, objPO.iSeq, objCT.lCodigo & SEPARADOR & objCT.sNomeReduzido, objPOMaquinas.objMaquina.iCodigo & SEPARADOR & objPOMaquinas.objMaquina.sNomeReduzido, objPOMaquinas.dtData)
                   
                    End If
                   
                Next
        
                'Essas validações não servem para o plano mestre e já existe
                'a verificação de saldo disponível para data
               'Verifica se vai conseguir comprar os produtos que faltam a tempo
'               lErro = MRP_Valida_Compras(objPMP, objPO, sMensagem)
'               If lErro <> SUCESSO Then gError 138213
                   
               'Verifica se vai conseguir produzir a tempo
'               lErro = MRP_Valida_Producao(objPMP, objPO, sMensagem)
'               If lErro <> SUCESSO Then gError 138214
        
               iSeqRel = iSeqRel + 1
               
               If objPO.iStatus <> PO_STATUS_OK Then
                   objPO.sDescErro = sMensagem
               End If
        
               lErro = Comando_Executar(alComando(0), "INSERT INTO RelCriticasPMP (NumIntRel,SeqRel,CodigoOPOrigem, CodigoOP, CodCliente, NomeCliente, Etapa, Nivel,Seq, Status, Mensagem ) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
               lNumIntRel, iSeqRel, objPMPItem.sCodOPOrigem, objPO.sCodOPOrigem, objPMPItem.lCliente, objCliente.sNomeReduzido, CStr(objPO.iNivel) & SEPARADOR & CStr(objPO.iSeq), objPO.iNivel, objPO.iSeq, objPO.iStatus, left(objPO.sDescErro, 255))
               If lErro <> SUCESSO Then gError 138118
        
               sMensagem = ""
        
            Next
           
        End If
        
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 138110
 
    lErro = objRelatorio.ExecutarDireto("Críticas PMP", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 138111
 
    RelCriticasPMP_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelCriticasPMP_Prepara:
 
    RelCriticasPMP_Prepara = gErr
 
    Select Case gErr
 
        Case 136556, 138119, 138120, 138121, 138213, 138214
 
        Case 138107
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 138108
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
  
        Case 138109, 138111, 138112, 138113
  
        Case 138110
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case 138118
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCRITICASPMP", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152346)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function MRP_Mensagem6(sMensagem As String, ByVal sProduto As String, ByVal dtData As Date) As Long
 
Dim lErro As Long
Dim sProdutoMascarado As String
 
On Error GoTo Erro_MRP_Mensagem6
    
    sProdutoMascarado = String(STRING_PRODUTO, 0)
    
    lErro = Mascara_RetornaProdutoTela(sProduto, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 138113
    
    sMensagem = sMensagem & "Não existe estoque suficiente do produto " & Trim(sProdutoMascarado) & " no dia " & Format(dtData, "dd/mm/yyyy") & ". "
   
    MRP_Mensagem6 = SUCESSO
    
    Exit Function
 
Erro_MRP_Mensagem6:
 
    MRP_Mensagem6 = gErr
 
    Select Case gErr
    
        Case 138113
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152347)
 
    End Select
 
    Exit Function
 
End Function

Public Function RelPreviaReqCompra_Prepara(ByVal objPMP As ClassPMP) As Long
'Monta a tabela temporária para o relatório de prévia requisições de compra gerada pelo MRP
 
Dim lErro As Long
Dim objPMPItem As New ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iSeqRel As Integer
Dim objCliente As ClassCliente
Dim objRC As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objProduto As ClassProduto
Dim objProdutoFilial As ClassProdutoFilial
 
On Error GoTo Erro_RelPreviaReqCompra_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136557
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136558
    Next
 
    'Acerta as informações das SubOps e Requisições de Compra
    lErro = PMP_Acerta_OP_RC(objPMP)
    If lErro <> SUCESSO Then gError 136559
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_PREVIAREQCOMPRA", lNumIntRel)
    If lErro <> SUCESSO Then gError 136560
   
    'Para cada item
    For Each objPMPItem In objPMP.colItens
    
        Set objCliente = New ClassCliente
 
        objCliente.lCodigo = objPMPItem.lCliente
        
        'Le o cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 136561
    
        'Para cada etapa
        For Each objPO In objPMPItem.colPO
                                  
            For Each objRC In objPO.colRCFilhas
        
                For Each objItemRC In objRC.colItens
                    
                    Set objProduto = New ClassProduto
                    
                    objProduto.sCodigo = objItemRC.sProduto
                    
                    'Le o produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 136562
                    
                    Set objProdutoFilial = New ClassProdutoFilial
                
                    objProdutoFilial.sProduto = objItemRC.sProduto
                    objProdutoFilial.iFilialEmpresa = objPO.iFilialEmpresa
                
                    'Le ProdutoFilial
                    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                    If lErro <> SUCESSO And lErro <> 28261 Then gError 136566
                    
                    iSeqRel = iSeqRel + 1
    
                    If objCliente.sNomeReduzido = "" Then objCliente.sNomeReduzido = "INTERNO"
    
                    lErro = Comando_Executar(alComando(0), "INSERT INTO RelPreviaReqCompra (NumIntRel,SeqRel,CodigoOPOrigem, CodigoOP, CodCliente, NomeCliente, Etapa, Nivel,Seq, Produto, DescProduto, Quantidade, UM, DataNecessidade, TempoRessuprimento, DataLimite) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                    lNumIntRel, iSeqRel, objPMPItem.sCodOPOrigem, objPO.sCodOPOrigem, objPMPItem.lCliente, objCliente.sNomeReduzido, CStr(objPO.iNivel) & SEPARADOR & CStr(objPO.iSeq), objPO.iNivel, objPO.iSeq, objProduto.sCodigo, objProduto.sDescricao, objItemRC.dQuantidade, objItemRC.sUM, objPO.dtDataInicio, objProdutoFilial.iTempoRessup, DateAdd("d", -1 * objProdutoFilial.iTempoRessup, objPO.dtDataInicio))
                    If lErro <> SUCESSO Then gError 136563
    
                Next
    
            Next
    
        Next
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136564
 
    lErro = objRelatorio.ExecutarDireto("Prévia de Requisições de Compra", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 136565
 
    RelPreviaReqCompra_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelPreviaReqCompra_Prepara:
 
    RelPreviaReqCompra_Prepara = gErr
 
    Select Case gErr
 
        Case 136557
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 136558
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136559 To 136562
 
        Case 136563
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELPREVIAREQCOMPRA", gErr)
        
        Case 136564
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 136565, 136566
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152348)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelEstoqueSimulacao_Prepara(ByVal objSimulacaoESTSeleciona As ClassSimulacaoESTSeleciona, lNumIntRel As Long) As Long
'Monta a tabela temporária para o relatório de projeção de estoque
 
Dim lErro As Long
Dim lTransacao As Long
Dim objProduto As ClassProduto
Dim colProduto As New Collection
Dim iSeq As Integer
Dim iSeqRel As Integer
Dim colSimulacaoESTItens As Collection

On Error GoTo Erro_RelEstoqueSimulacao_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136714
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_SIMULACAOEST", lNumIntRel)
    If lErro <> SUCESSO Then gError 136716
    
    'TEM QUE PEGAR OS PRODUTOS COM OS FILTROS
    lErro = CF("Produto_Le_Faixa", objSimulacaoESTSeleciona.sProdutoInicial, objSimulacaoESTSeleciona.sProdutoFinal, colProduto)
    If lErro <> SUCESSO Then gError 136717
        
    iSeqRel = 0
    
    For Each objProduto In colProduto
       
        iSeq = 0
        
        Set colSimulacaoESTItens = New Collection
        
        iSeqRel = iSeqRel + 1
        
        'Chama a função que Obtém e Grava os dados
        lErro = RelEstoqueSimulacao_Prepara_EmTrans(colSimulacaoESTItens, objProduto, objSimulacaoESTSeleciona.iFilialEmpresa, iSeqRel, iSeq, lNumIntRel)
        If lErro <> SUCESSO Then gError 138063
    
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136740
 
    RelEstoqueSimulacao_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelEstoqueSimulacao_Prepara:
 
    RelEstoqueSimulacao_Prepara = gErr
 
    Select Case gErr
 
        Case 136714
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 136716 To 136717, 138063

        Case 136740
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152349)
 
    End Select

    Call Transacao_Rollback
 
    Exit Function
 
End Function
 
Public Function RelEstoqueSimulacaoPMP_Prepara(ByVal objPMP As ClassPMP, ByVal iFilialEmpresa As Integer, lNumIntRel As Long) As Long
'Obtém e Grava as movimentações de estoque projetadas
 
Dim lErro As Long
Dim lTransacao As Long
Dim objProduto As ClassProduto
Dim colProduto As New Collection
Dim colSimulacaoESTItens As New Collection
Dim colSimulacaoESTItemTudo As New Collection
Dim iSeqRel As Integer
Dim iSeq As Integer
Dim iSeqRelAux As Integer
Dim iSeqAux As Integer
Dim objPMPItens As ClassPMPItens
Dim colcolSimulacaoESTItens As New Collection
Dim objPO As ClassPlanoOperacional
Dim objRC As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim bAchou As Boolean
Dim iIndice As Integer
Dim objSimulacaoESTItem As ClassSimulacaoESTItens
Dim objOPOperacao As ClassOrdemProducaoOperacoes
Dim objOPOperacaoAux As ClassOrdemProducaoOperacoes
Dim objOPOperacaoAux2 As ClassOrdemProducaoOperacoes
Dim objOPOperacaoAux3 As ClassOrdemProducaoOperacoes
Dim objOPInsumos As ClassOrdemProducaoInsumos
Dim objOPInsumoAux As ClassOrdemProducaoInsumos
Dim objItemOP As ClassItemOP
Dim dFatorConversao As Double
Dim dQuantidade As Double
Dim alComando(1 To 2) As Long
Dim dFatorMultiplicacao As Double
Dim objProdutoFilial As ClassProdutoFilial
Dim sCodigoOPAnt As String
Dim sProdutoAnt As String
Dim colCampos As Collection
Dim colSaida As Collection
 
On Error GoTo Erro_RelEstoqueSimulacaoPMP_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 138064
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138065
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_SIMULACAOEST", lNumIntRel)
    If lErro <> SUCESSO Then gError 138066
        
    iSeqRel = 0
        
    'Acerta as datas das OPs e RCs de acordo com as alterações do usuário
    lErro = PMP_Acerta_OP_RC(objPMP)
    If lErro <> SUCESSO Then gError 138074
        
    'Para cada item do Plano Mestre
    For Each objPMPItens In objPMP.colItens
    
        If objPMPItens.iFilialEmpresa = iFilialEmpresa Then
    
            Set objItemOP = objPMPItens.objItemOP
            
            iIndice = 0
                
            sCodigoOPAnt = ""
            sProdutoAnt = ""
            
            Set colCampos = New Collection
            Set colSaida = New Collection
            
            colCampos.Add "sCodOPOrigem"
            colCampos.Add "sProduto"
            colCampos.Add "iNivel"
            colCampos.Add "iSeq"
            
            Call Ordena_Colecao(objPMPItens.colPO, colSaida, colCampos)
            
            'Para cada Etapa
            For Each objPO In colSaida
        
                iIndice = iIndice + 1
                
                If sCodigoOPAnt <> objPO.sCodOPOrigem And sProdutoAnt <> objPO.sProduto Then
 
                    'Insere os Itens produzidos por cada etapa
                    Set objSimulacaoESTItem = New ClassSimulacaoESTItens
                    
                    objSimulacaoESTItem.dQuantidade = objPO.dQuantidade
                    objSimulacaoESTItem.dtData = objPO.dtDataFim
                    objSimulacaoESTItem.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAO
                    objSimulacaoESTItem.lNumIntRel = lNumIntRel
                    objSimulacaoESTItem.sProduto = objPO.sProduto
                    objSimulacaoESTItem.sUM = objPO.sUM
                    objSimulacaoESTItem.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_ENTRADA
                    objSimulacaoESTItem.iSimulacao = SIMULACAO_ESTOQUE_PROJETADO
                    
                    colSimulacaoESTItemTudo.Add objSimulacaoESTItem
                    
                    'Insere os Itens produzidos parar cada etapa
                    Set objSimulacaoESTItem = New ClassSimulacaoESTItens
                    
                    objSimulacaoESTItem.dQuantidade = objPO.dQuantidade
                    objSimulacaoESTItem.dtData = objPO.dtDataInicio
                    objSimulacaoESTItem.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAOINI
                    objSimulacaoESTItem.lNumIntRel = lNumIntRel
                    objSimulacaoESTItem.sProduto = objPO.sProduto
                    objSimulacaoESTItem.sUM = objPO.sUM
                    objSimulacaoESTItem.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA
                    objSimulacaoESTItem.iSimulacao = SIMULACAO_ESTOQUE_OUTROS
                    
                    colSimulacaoESTItemTudo.Add objSimulacaoESTItem
                    
                    sCodigoOPAnt = objPO.sCodOPOrigem
                    sProdutoAnt = objPO.sProduto
                
                End If
                
                Set objProduto = New ClassProduto
                             
                'Para cada Operação
                For Each objOPOperacao In objItemOP.colOrdemProducaoOperacoes
                               
                    'Se for a operação que originou a etapa
                    If objOPOperacao.lNumIntDoc = objPO.lNumIntDocOper Then
                                        
                        'Para cada insumo
                        For Each objOPInsumos In objOPOperacao.colOPInsumos
                                           
                            If objOPOperacao.iNivelRoteiro <> KIT_NIVEL_RAIZ Then
                        
                               'Tenta achar a Operação Pai
                                For Each objOPOperacaoAux In objItemOP.colOrdemProducaoOperacoes
                                
                                    If objOPOperacaoAux.iSeq = objOPOperacao.iSeqPai Then
                                    
'                                        'Tenta achar o Insumo que originou a OP pai
'                                        For Each objOPInsumoAux In objOPOperacaoAux.colOPInsumos
'
'                                            If objOPInsumoAux.sProduto = objOPOperacao.sProduto Then
'
'                                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objOPInsumoAux.sUMProduto, dFatorConversao)
'                                                If lErro <> SUCESSO Then gError 138103
'
'                                                dFatorMultiplicacao = (objPO.dQuantidade * dFatorConversao) / objOPInsumoAux.dQuantidade
'
'                                                Exit For
'
'                                            End If
'
'                                        Next

                                        'Verifica se é do mesmo roteiro
                                        If objOPOperacaoAux.iNivelRoteiro = objOPOperacao.iNivelRoteiro And objOPOperacaoAux.iSeqRoteiro = objOPOperacao.iSeqRoteiro Then
                                                                                    
                                            'Acha a primeira Operação do Roteiro
                                            For Each objOPOperacaoAux2 In objItemOP.colOrdemProducaoOperacoes
                                             
                                                 If objOPOperacao.iNivelRoteiro = objOPOperacaoAux2.iNivelRoteiro And objOPOperacao.iSeqRoteiro = objOPOperacaoAux2.iSeqRoteiro Then
                                                    Exit For
                                                 End If
                                                 
                                            Next
                                            
                                            'Acha a Operação Pai da Primeira Operação do roteiro
                                            For Each objOPOperacaoAux3 In objItemOP.colOrdemProducaoOperacoes
                                             
                                                 If objOPOperacaoAux3.iSeq = objOPOperacaoAux2.iSeqPai Then
                                                    Exit For
                                                 End If
                                            
                                            Next
                                            
                                        Else
                                        
                                            Set objOPOperacaoAux3 = objOPOperacaoAux
                                        
                                        End If
                                        
                                        'Tenta achar o Insumo que originou a OP pai
                                        For Each objOPInsumoAux In objOPOperacaoAux3.colOPInsumos
                                    
                                            If objOPInsumoAux.sProduto = objOPOperacao.sProduto Then
                        
                                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objOPInsumoAux.sUMProduto, dFatorConversao)
                                                If lErro <> SUCESSO Then gError 138103
                        
                                                dFatorMultiplicacao = (objPO.dQuantidade * dFatorConversao) / objOPInsumoAux.dQuantidade
    
                                                Exit For
                                                
                                            End If
                                    
                                        Next
                                    
                                        Exit For
                                    
                                    End If
                                
                                Next
                                
                            Else
                            
                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objItemOP.sSiglaUM, dFatorConversao)
                                If lErro <> SUCESSO Then gError 138103
                            
                                dFatorMultiplicacao = (objPO.dQuantidade * dFatorConversao) / objItemOP.dQuantidade
                            
                            End If
                        
                            'Insere os itens consumidos em cada etapa
                            Set objSimulacaoESTItem = New ClassSimulacaoESTItens
                            
                            objSimulacaoESTItem.dQuantidade = (-1 * objOPInsumos.dQuantidade * dFatorMultiplicacao) / (1 - objOPInsumos.dPercentualPerda)
                            objSimulacaoESTItem.dtData = objPO.dtDataInicio
                            objSimulacaoESTItem.iTipo = SIMULACAO_ESTOQUE_TIPO_COMSUMO
                            objSimulacaoESTItem.lNumIntRel = lNumIntRel
                            objSimulacaoESTItem.sProduto = objOPInsumos.sProduto
                            objSimulacaoESTItem.sUM = objOPInsumos.sUMProduto
                            objSimulacaoESTItem.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_SAIDA
                            objSimulacaoESTItem.iSimulacao = SIMULACAO_ESTOQUE_PROJETADO
                            
                            colSimulacaoESTItemTudo.Add objSimulacaoESTItem
                        
                        Next
                    
                    End If
                
                Next
                    
                For Each objRC In objPO.colRCFilhas
                
                    For Each objItemRC In objRC.colItens
                    
                        'Insere os itens comprados em cada etapa
                        Set objSimulacaoESTItem = New ClassSimulacaoESTItens
                         
                        objSimulacaoESTItem.dQuantidade = objItemRC.dQuantidade
                        objSimulacaoESTItem.dtData = objRC.dtDataLimite
                        objSimulacaoESTItem.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVCOMPRA
                        objSimulacaoESTItem.lNumIntRel = lNumIntRel
                        objSimulacaoESTItem.sProduto = objItemRC.sProduto
                        objSimulacaoESTItem.sUM = objItemRC.sUM
                        objSimulacaoESTItem.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_ENTRADA
                        objSimulacaoESTItem.iSimulacao = SIMULACAO_ESTOQUE_PROJETADO
                         
                        colSimulacaoESTItemTudo.Add objSimulacaoESTItem
                
                        'Insere os itens comprados em cada etapa
                        Set objSimulacaoESTItem = New ClassSimulacaoESTItens
                        Set objProdutoFilial = New ClassProdutoFilial
                    
                        objProdutoFilial.sProduto = objItemRC.sProduto
                        objProdutoFilial.iFilialEmpresa = objPO.iFilialEmpresa
                    
                        'Le ProdutoFilial
                        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                        If lErro <> SUCESSO And lErro <> 28261 Then gError 138178
                         
                        objSimulacaoESTItem.dQuantidade = objItemRC.dQuantidade
                        objSimulacaoESTItem.dtData = DateAdd("d", -1 * objProdutoFilial.iTempoRessup, objRC.dtDataLimite)
                        objSimulacaoESTItem.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVCOMPRAINI
                        objSimulacaoESTItem.lNumIntRel = lNumIntRel
                        objSimulacaoESTItem.sProduto = objItemRC.sProduto
                        objSimulacaoESTItem.sUM = objItemRC.sUM
                        objSimulacaoESTItem.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA
                        objSimulacaoESTItem.iSimulacao = SIMULACAO_ESTOQUE_OUTROS
                         
                        colSimulacaoESTItemTudo.Add objSimulacaoESTItem
                
                    Next
                
                Next
                
            Next
            
        End If
        
    Next
            
    'Para cada entrada ou saída projetada do estoque agrupa por produto
    For Each objSimulacaoESTItem In colSimulacaoESTItemTudo
        
       bAchou = False
        
        iIndice = 0
        'Verifica se o produto já está na coleção
        For Each objProduto In colProduto
            iIndice = iIndice + 1
            If objProduto.sCodigo = objSimulacaoESTItem.sProduto Then
                bAchou = True
                Exit For
            End If
        Next
                    
        'Se a coleção do Produto ainda não foi inicializada, inicializa
        If Not bAchou Then
        
            iSeqRel = iSeqRel + 1
            iSeq = 1
            
            iSeqRelAux = iSeqRel
            iSeqAux = iSeq
 
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objSimulacaoESTItem.sProduto
            
            'Lê o produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 138067
            
            colProduto.Add objProduto
            
            Set colSimulacaoESTItens = New Collection
            
            colcolSimulacaoESTItens.Add colSimulacaoESTItens
            
        Else
            'Se o produto já foi Inicializado, obtém a coleção e os sequenciais do produto corrente
            Set colSimulacaoESTItens = colcolSimulacaoESTItens.Item(iIndice)
            
            If colSimulacaoESTItens.Count > 0 Then
                iSeqAux = 1 + colcolSimulacaoESTItens.Item(iIndice).Item(colcolSimulacaoESTItens.Item(iIndice).Count).iSeq
            Else
                iSeqAux = 1
            End If
            
            iSeqRelAux = iIndice
        End If
        
        'Converte tudo para UM de estoque
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objSimulacaoESTItem.sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 138068
    
        objSimulacaoESTItem.iSeq = iSeqAux
        objSimulacaoESTItem.iSeqRel = iSeqRelAux
        objSimulacaoESTItem.dQuantidade = objSimulacaoESTItem.dQuantidade * dFatorConversao
    
        colSimulacaoESTItens.Add objSimulacaoESTItem
    
        Select Case objSimulacaoESTItem.iTipo
        
            Case SIMULACAO_ESTOQUE_TIPO_COMSUMO, SIMULACAO_ESTOQUE_TIPO_PREVVENDA
                dQuantidade = objSimulacaoESTItem.dQuantidade * -1
            Case Else
                dQuantidade = objSimulacaoESTItem.dQuantidade
        End Select
    
        'Insere na tabela temporária
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", objSimulacaoESTItem.lNumIntRel, objSimulacaoESTItem.iSeqRel, objSimulacaoESTItem.iSeq, objSimulacaoESTItem.dtData, dQuantidade, objSimulacaoESTItem.iTipo, objSimulacaoESTItem.iSimulacao, objSimulacaoESTItem.iTipoOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 138069
        
    Next
            
    iIndice = 0
            
    'Para cada produto usado
    For Each objProduto In colProduto
    
        iIndice = iIndice + 1
    
        'Obtem os sequenciais
        If colcolSimulacaoESTItens.Item(iIndice).Count > 0 Then
            iSeq = colcolSimulacaoESTItens.Item(iIndice).Item(colcolSimulacaoESTItens.Item(iIndice).Count).iSeq
        Else
            iSeq = 0
        End If
        
        iSeqRel = iIndice
    
        'Prepara as informações que já foram gravadas
        lErro = RelEstoqueSimulacao_Prepara_EmTrans(colcolSimulacaoESTItens.Item(iIndice), objProduto, iFilialEmpresa, iSeqRel, iSeq, lNumIntRel)
        If lErro <> SUCESSO Then gError 138070
    
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 138071
 
    RelEstoqueSimulacaoPMP_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelEstoqueSimulacaoPMP_Prepara:
 
    RelEstoqueSimulacaoPMP_Prepara = gErr
 
    Select Case gErr
  
        Case 138064
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 138065
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 138066 To 138068, 138070, 138074, 138178
 
        Case 138069
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELSIMULACAOEST", gErr)
 
        Case 138071
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152350)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelEstoqueSimulacao_Prepara_EmTrans(ByVal colSimulacaoESTItens As Collection, ByVal objProduto As ClassProduto, ByVal iFilialEmpresa As Integer, iSeqRel As Integer, iSeq As Integer, lNumIntRel As Long) As Long
'Obtém e Grava as movimentações de estoque que irão acontecer e a disponibilidade nesses dias
 
Dim lErro As Long
Dim alComando(0 To 8) As Long
Dim iIndice As Integer
Dim dQuantidade As Double
Dim dtData As Date
Dim sUM As String
Dim dFatorConversao As Double
Dim dEstoque As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim objSimulacaoESTItens As ClassSimulacaoESTItens
Dim objSimulacaoESTItensAux1 As ClassSimulacaoESTItens
Dim objSimulacaoESTItensAux2 As ClassSimulacaoESTItens
Dim dtDataAnterior As Date
Dim dSaldo As Double
Dim bAchou As Boolean
Dim colSaida As New Collection
Dim colCampos As New Collection
Dim dSaldoFalta As Double
Dim dSaldoFaltaAnterior As Double
Dim lNumIntPV As Long
Dim dQuantAux1 As Double
Dim dQuantAux2 As Double
Dim lCodEmpenho As Long
Dim lCodEmpenhoAnt As Long

On Error GoTo Erro_RelEstoqueSimulacao_Prepara_EmTrans
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136715
    Next
    
    'PARA CADA PRODUTO TENHO QUE INSERIR PARA CADA DATA
    
    '1 - QUANTIDADE EM ESTOQUE/ESTOQUE DE SEGURANÇA
    '2 - QUANTIDADE EMPENHADA (CONSUMIDO)
    '3 - QUANTIDADE DE PREVISÃO DE VENDA
    '3.1 - QUANTIDADE DE PREVISÃO DE VENDA COM VÁRIAS DATAS DE ENTREGA
    '4 - QUANTIDADE DE PREVISÃO DE PRODUÇÃO
    '5 - SALDO EM ESTOQUE
    '6 - SALDO DISPONÍVEL EM ESTOQUE
  
    'ESTOQUE
    objEstoqueProduto.sProduto = objProduto.sCodigo
    
    'Lê as quantidades do produtos nos almoxs da filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 52979 Then gError 136718
            
    objProdutoFilial.sProduto = objProduto.sCodigo
    objProdutoFilial.iFilialEmpresa = iFilialEmpresa

    'Le ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 136719
                   
    lErro = Comando_Executar(alComando(0), "INSERT INTO RelSimulacaoEST (NumIntRel,Seq, Produto, Descricao, UM, QuantidadeEstoque, EstoqueSeguranca) VALUES (?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, objEstoqueProduto.dQuantDisponivel, objProdutoFilial.dEstoqueSeguranca)
    If lErro <> AD_SQL_SUCESSO Then gError 136720
    'FIM ESTOQUE
    
    'CONSUMIDO
'    lErro = Comando_Executar(alComando(1), "SELECT E.Codigo, E.Quantidade, PO.DataInicio FROM Empenho AS E, ItensOrdemProducao AS IOP, PlanoOperacional AS PO, OrdemProducaoOperacoes AS O, OrdemProducaoInsumos AS OI " & _
'                                         "WHERE E.NumIntDocItemOP = IOP.NumIntDoc AND PO.DataInicio  >= ? AND E.Produto = ? AND E.FilialEmpresa = ? AND " & _
'                                         "O.NumIntDocItemOP = IOP.NumIntDoc AND O.Produto = IOP.Produto AND OI.NumIntDocOper = O.NumIntDoc AND PO.NUmIntDocOper = O.NumIntDoc AND OI.Produto = E.Produto AND " & _
'                                        "EXISTS (SELECT PMPItens.NumIntDoc FROM PlanoMestreProducaoUltimo AS PMP, PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMP.CodGeracao = PMPItens.CodGeracao AND PO.NumIntDocPMP = PMPItens.NumIntDoc AND IOP.Produto = PO.Produto AND IOP.Codigo = PO.CodOPOrigem) " & _
'                                        "UNION ALL SELECT E.Codigo, E.Quantidade, PO.DataInicio FROM Empenho AS E, ItensOrdemProducao AS IOP, PlanoOperacional AS PO, OrdemProducaoOperacoes AS O, OrdemProducaoInsumos AS OI " & _
'                                         "WHERE E.NumIntDocItemOP = IOP.NumIntDoc AND PO.DataInicio  >= ? AND E.Produto = ? AND E.FilialEmpresa = ? AND " & _
'                                         "O.NumIntDocItemOP = IOP.NumIntDoc AND O.Produto = IOP.Produto AND OI.NumIntDocOper = O.NumIntDoc AND PO.NUmIntDocOper = O.NumIntDocOperOrigem AND OI.Produto = E.Produto AND " & _
'                                        "EXISTS (SELECT PMPItens.NumIntDoc FROM PlanoMestreProducaoUltimo AS PMP, PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMP.CodGeracao = PMPItens.CodGeracao AND PO.NumIntDocPMP = PMPItens.NumIntDoc AND IOP.Produto = PO.Produto AND IOP.Codigo = PO.CodOPOrigem) ORDER BY E.Codigo ", _
'                                         lCodEmpenho, dQuantidade, dtData, gdtDataAtual, objProduto.sCodigo, iFilialEmpresa, gdtDataAtual, objProduto.sCodigo, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 136721
    lErro = Comando_Executar(alComando(1), "SELECT E.Codigo, (E.Quantidade - E.QuantidadeRequisitada), IOP.DataInicioProd " & _
                                        "FROM Empenho AS E, ItensOrdemProducao AS IOP " & _
                                        "WHERE E.NumIntDocItemOP = IOP.NumIntDoc AND IOP.DataInicioProd >= ? AND E.Produto = ? AND E.FilialEmpresa = ? AND " & _
                                        "EXISTS (SELECT PMPItens.NumIntDoc FROM PlanoMestreProducaoUltimo AS PMP, PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMP.CodGeracao = PMPItens.CodGeracao AND PO.NumIntDocPMP = PMPItens.NumIntDoc AND IOP.Produto = PO.Produto AND IOP.Codigo = PO.CodOPOrigem)", _
                                        lCodEmpenho, dQuantidade, dtData, gdtDataAtual, objProduto.sCodigo, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 136721
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136722
    
    sUM = objProduto.sSiglaUMEstoque
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Pula a quantidade caso exista mais de uma etapa do mesmo item usando o mesmo produto
        'em datas diferentes (não tem como desmembrar o consumo, apesar das requisições de compra
        'terem sido geradas separadamente com as datas corretas)
        If lCodEmpenhoAnt <> lCodEmpenho Then
    
            iSeq = iSeq + 1
            
            Set objSimulacaoESTItens = New ClassSimulacaoESTItens
            
            objSimulacaoESTItens.dQuantidade = dQuantidade * -1
            objSimulacaoESTItens.dtData = dtData
            objSimulacaoESTItens.iSeq = iSeq
            objSimulacaoESTItens.iSeqRel = iSeqRel
            objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_COMSUMO
            objSimulacaoESTItens.lNumIntRel = lNumIntRel
            objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_SAIDA
            objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_REAL
            
            colSimulacaoESTItens.Add objSimulacaoESTItens
        
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelSimulacaoESTItens (NumIntRel, SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, dtData, dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
            If lErro <> AD_SQL_SUCESSO Then gError 136723
            
            lCodEmpenhoAnt = lCodEmpenho
            
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136724
    
    Loop
    'FIM CONSUMO
       
    'VENDA
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT IPV.Quantidade - IPV.QuantReservada - IPV.QuantCancelada - IPV.QuantFaturada, IPV.UnidadeMed, IPV.DataEntrega " & _
                                        "FROM ItensPedidoDeVenda AS IPV " & _
                                        "WHERE IPV.DataEntrega >= ? AND IPV.Produto = ? AND IPV.FilialEmpresa = ? AND NOT EXISTS (SELECT IPVE.NumIntPV FROM ItensPVEntrega AS IPVE WHERE IPVE.NumIntItemPV = IPV.NumIntDoc) ", _
                                        dQuantidade, sUM, dtData, gdtDataAtual, objProduto.sCodigo, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 136725
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136726
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 136727
        
        dQuantidade = dQuantidade * dFatorConversao
        
        iSeq = iSeq + 1
    
        Set objSimulacaoESTItens = New ClassSimulacaoESTItens
        
        objSimulacaoESTItens.dQuantidade = dQuantidade * -1
        objSimulacaoESTItens.dtData = dtData
        objSimulacaoESTItens.iSeq = iSeq
        objSimulacaoESTItens.iSeqRel = iSeqRel
        objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVVENDA
        objSimulacaoESTItens.lNumIntRel = lNumIntRel
        objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_SAIDA
        objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_REAL
        
        colSimulacaoESTItens.Add objSimulacaoESTItens
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, dtData, dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 136728
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136729
    
    Loop
    'FIM VENDA
    
    'VENDA COM VÁRIAS DATAS DE ENTREGA
    sUM = String(STRING_UM_SIGLA, 0)
    
    'Le o item de entrega do pedido de venda
    lErro = Comando_Executar(alComando(6), "SELECT IPV.NumIntDoc, IPVE.Quantidade ,IPV.Quantidade - IPV.QuantReservada - IPV.QuantCancelada - IPV.QuantFaturada, IPV.UnidadeMed, IPVE.DataEntrega  " & _
                                        "FROM ItensPedidoDeVenda AS IPV, ItensPVEntrega AS IPVE " & _
                                        "WHERE IPVE.NumIntItemPV = IPV.NumIntDoc AND IPVE.DataEntrega >= ? AND IPV.Produto = ? AND IPV.FilialEmpresa = ? ", _
                                        lNumIntPV, dQuantidade, dQuantAux1, sUM, dtData, gdtDataAtual, objProduto.sCodigo, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 136725
    
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136726
   
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Le o que vai ter que ser entrega após o item lido acima
        lErro = Comando_Executar(alComando(7), "SELECT SUM(IPVE.Quantidade) " & _
                                            "FROM  ItensPVEntrega AS IPVE " & _
                                            "WHERE IPVE.NumIntItemPV = ? AND IPVE.DataEntrega > ? ", _
                                            dQuantAux2, lNumIntPV, dtData)
        If lErro <> SUCESSO Then gError 136725
        
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136726
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 136727
        
        'Se vai ser entrega mais do que precisa ser produzido é porque a
        'entrega atual usa o que já foi reservado, logo não consume do estoque
        If dQuantAux2 >= dQuantAux1 Then
            dQuantidade = 0
        Else
            'Se o que vai ser entregue é maior que o que falta nas entregas futuras
            If dQuantidade > dQuantAux1 - dQuantAux2 Then
                dQuantidade = dQuantAux1 - dQuantAux2
            End If
        End If
        
        dQuantidade = dQuantidade * dFatorConversao
        
        iSeq = iSeq + 1
    
        Set objSimulacaoESTItens = New ClassSimulacaoESTItens
        
        objSimulacaoESTItens.dQuantidade = dQuantidade * -1
        objSimulacaoESTItens.dtData = dtData
        objSimulacaoESTItens.iSeq = iSeq
        objSimulacaoESTItens.iSeqRel = iSeqRel
        objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVVENDA
        objSimulacaoESTItens.lNumIntRel = lNumIntRel
        objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_SAIDA
        objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_REAL
        
        colSimulacaoESTItens.Add objSimulacaoESTItens
        
        lErro = Comando_Executar(alComando(8), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, dtData, dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 136728
    
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136729
    
    Loop
    'FIM VENDA COM VÁRIAS DATAS DE ENTREGA

    
    'PRODUÇÃO
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(4), "SELECT IOP.Quantidade, IOP.SiglaUM, IOP.DataFimProd " & _
                                        "FROM ItensOrdemProducao AS IOP " & _
                                        "WHERE IOP.DataFimProd >= ? AND IOP.Produto = ? AND IOP.FilialEmpresa = ? AND " & _
                                        "EXISTS (SELECT PMPItens.NumIntDoc FROM PlanoMestreProducaoUltimo AS PMP, PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMP.CodGeracao = PMPItens.CodGeracao AND PO.NumIntDocPMP = PMPItens.NumIntDoc AND IOP.Produto = PO.Produto AND IOP.Codigo = PO.CodOPOrigem)", _
                                        dQuantidade, sUM, dtData, gdtDataAtual, objProduto.sCodigo, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 136730
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136731
    
    Do While lErro <> AD_SQL_SEM_DADOS
           
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 136732
        
        dQuantidade = dQuantidade * dFatorConversao
        
        iSeq = iSeq + 1
        
        Set objSimulacaoESTItens = New ClassSimulacaoESTItens
        
        objSimulacaoESTItens.dQuantidade = dQuantidade
        objSimulacaoESTItens.dtData = dtData
        objSimulacaoESTItens.iSeq = iSeq
        objSimulacaoESTItens.iSeqRel = iSeqRel
        objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAO
        objSimulacaoESTItens.lNumIntRel = lNumIntRel
        objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_ENTRADA
        objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_REAL
        
        colSimulacaoESTItens.Add objSimulacaoESTItens
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelSimulacaoESTItens (NumIntRel, SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, dtData, dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 136733
    
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136734
    
    Loop
    'FIM PRODUÇÃO
    
    'COMPRA
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(4), "SELECT IRC.Quantidade - IRC.QuantCancelada AS Quantidade, IRC.UM, R.DataLimite " & _
                                            "FROM RequisicaoCompra AS R, ItensReqCompra AS IRC " & _
                                            "WHERE R.NumIntDoc = IRC.ReqCompra AND IRC.StatusBaixa = ? AND IRC.Produto = ? AND R.DataLimite >= ? AND R.FilialEmpresa = ? ", _
                                            dQuantidade, sUM, dtData, 0, objProduto.sCodigo, gdtDataAtual, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 136735
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136736
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 136737
        
        dQuantidade = dQuantidade * dFatorConversao
        
        iSeq = iSeq + 1
        
        Set objSimulacaoESTItens = New ClassSimulacaoESTItens
        
        objSimulacaoESTItens.dQuantidade = dQuantidade
        objSimulacaoESTItens.dtData = dtData
        objSimulacaoESTItens.iSeq = iSeq
        objSimulacaoESTItens.iSeqRel = iSeqRel
        objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVCOMPRA
        objSimulacaoESTItens.lNumIntRel = lNumIntRel
        objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_ENTRADA
        objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_REAL
        
        colSimulacaoESTItens.Add objSimulacaoESTItens
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, dtData, dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 136738
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136739
    
    Loop
    'FIM COMPRA
                
    'SALDO ATUAL
    
    'Para cada Item
    For Each objSimulacaoESTItensAux1 In colSimulacaoESTItens
    
        'Se não fro um item de saldo
        If objSimulacaoESTItensAux1.iSimulacao <> SIMULACAO_ESTOQUE_CONSOLIDADO Then
    
            bAchou = False
            
            'Busca o Saldo
            For Each objSimulacaoESTItensAux2 In colSimulacaoESTItens
                If objSimulacaoESTItensAux2.dtData = objSimulacaoESTItensAux1.dtData And objSimulacaoESTItensAux2.iTipo = SIMULACAO_ESTOQUE_TIPO_SALDOATUAL Then
                    bAchou = True
                End If
            Next
            
            'Se ainda não tem saldo
            If Not bAchou Then
            
                dSaldo = objEstoqueProduto.dQuantDisponivel
            
                'Acumula o saldo anterior
                For Each objSimulacaoESTItensAux2 In colSimulacaoESTItens
                    If objSimulacaoESTItensAux2.dtData <= objSimulacaoESTItensAux1.dtData And objSimulacaoESTItensAux2.iTipoOrdem <> SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA Then
                        dSaldo = dSaldo + objSimulacaoESTItensAux2.dQuantidade
                    End If
                Next
                
                iSeq = iSeq + 1
                
                Set objSimulacaoESTItens = New ClassSimulacaoESTItens
                
                objSimulacaoESTItens.dQuantidade = dSaldo
                objSimulacaoESTItens.dtData = objSimulacaoESTItensAux1.dtData
                objSimulacaoESTItens.iSeq = iSeq
                objSimulacaoESTItens.iSeqRel = iSeqRel
                objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_SALDOATUAL
                objSimulacaoESTItens.lNumIntRel = lNumIntRel
                objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA
                objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_CONSOLIDADO
                
                colSimulacaoESTItens.Add objSimulacaoESTItens
                
                lErro = Comando_Executar(alComando(5), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, objSimulacaoESTItensAux1.dtData, dSaldo, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
                If lErro <> AD_SQL_SUCESSO Then gError 136738
                        
            End If
            
        End If
    
    Next
    'FIM SALDO ATUAL
    
    'SALDO LIVRE
    Set colCampos = New Collection
    Set colSaida = New Collection
    
    'Ordena pela prioridade
    colCampos.Add "dtData"
    
    lErro = Ordena_Colecao(colSimulacaoESTItens, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 138060
    
    iIndice = 0
    
    For Each objSimulacaoESTItens In colSaida
    
        iIndice = iIndice + 1
       
        objSimulacaoESTItens.iOrdem = colSaida.Count - iIndice
    
    Next
    
    Set colCampos = New Collection
    Set colSimulacaoESTItens = New Collection
    
    'Ordena pela ordem descrecente de data
    colCampos.Add "iOrdem"
    
    lErro = Ordena_Colecao(colSaida, colSimulacaoESTItens, colCampos)
    If lErro <> SUCESSO Then gError 138061
    
    If colSimulacaoESTItens.Count > 0 Then
        dtDataAnterior = colSimulacaoESTItens.Item(1).dtData
    
        dSaldoFaltaAnterior = 0
        dSaldo = 0
        
        'Para cada entrada/saída de estoque
        For Each objSimulacaoESTItensAux1 In colSimulacaoESTItens
       
            If objSimulacaoESTItensAux1.iTipo <> SIMULACAO_ESTOQUE_TIPO_SALDODISPONIVEL Then
       
                If dtDataAnterior <> objSimulacaoESTItensAux1.dtData Then
                   
                    iSeq = iSeq + 1
                    
                    Set objSimulacaoESTItens = New ClassSimulacaoESTItens
                    
                    If dSaldoFaltaAnterior < 0 Then dSaldoFaltaAnterior = 0
                    
                    objSimulacaoESTItens.dQuantidade = dSaldo - dSaldoFaltaAnterior
                    objSimulacaoESTItens.dtData = dtDataAnterior
                    objSimulacaoESTItens.iSeq = iSeq
                    objSimulacaoESTItens.iSeqRel = iSeqRel
                    objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_SALDODISPONIVEL
                    objSimulacaoESTItens.lNumIntRel = lNumIntRel
                    objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA
                    objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_CONSOLIDADO
                    
                    colSimulacaoESTItens.Add objSimulacaoESTItens
                    
                    lErro = Comando_Executar(alComando(5), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, objSimulacaoESTItens.dtData, objSimulacaoESTItens.dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
                    If lErro <> AD_SQL_SUCESSO Then gError 138062
        
                    dSaldo = 0
                    dSaldoFaltaAnterior = dSaldoFalta
                    dtDataAnterior = objSimulacaoESTItensAux1.dtData
                
                End If
                
                Select Case objSimulacaoESTItensAux1.iTipo
                    
                    Case SIMULACAO_ESTOQUE_TIPO_SALDOATUAL
                        dSaldo = objSimulacaoESTItensAux1.dQuantidade
                        
                    Case Else
                    
                        If objSimulacaoESTItensAux1.iTipoOrdem <> SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA Then
                            dSaldoFalta = dSaldoFalta - objSimulacaoESTItensAux1.dQuantidade
                        End If
                        
                End Select
                                           
            End If
    
        Next
        
        iSeq = iSeq + 1
        
        Set objSimulacaoESTItens = New ClassSimulacaoESTItens
        
        If dSaldoFaltaAnterior < 0 Then dSaldoFaltaAnterior = 0
        
        objSimulacaoESTItens.dQuantidade = dSaldo - dSaldoFaltaAnterior
        objSimulacaoESTItens.dtData = dtDataAnterior
        objSimulacaoESTItens.iSeq = iSeq
        objSimulacaoESTItens.iSeqRel = iSeqRel
        objSimulacaoESTItens.iTipo = SIMULACAO_ESTOQUE_TIPO_SALDODISPONIVEL
        objSimulacaoESTItens.lNumIntRel = lNumIntRel
        objSimulacaoESTItens.iTipoOrdem = SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA
        objSimulacaoESTItens.iSimulacao = SIMULACAO_ESTOQUE_CONSOLIDADO
        
        colSimulacaoESTItens.Add objSimulacaoESTItens
        
        lErro = Comando_Executar(alComando(5), "INSERT INTO RelSimulacaoESTItens (NumIntRel,SeqRel, Seq, Data, Quantidade, Tipo, Simulacao, TipoOrdem) VALUES (?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRel, iSeq, objSimulacaoESTItens.dtData, objSimulacaoESTItens.dQuantidade, objSimulacaoESTItens.iTipo, objSimulacaoESTItens.iSimulacao, objSimulacaoESTItens.iTipoOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 138062
        
    End If
    'FIM SALDO LIVRE
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelEstoqueSimulacao_Prepara_EmTrans = SUCESSO
    
    Exit Function

Erro_RelEstoqueSimulacao_Prepara_EmTrans:
 
    RelEstoqueSimulacao_Prepara_EmTrans = gErr
 
    Select Case gErr
        
        Case 136715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136718 To 136719, 136727, 136732, 136737, 138060, 138061
        
        Case 136720, 136723, 136728, 136733, 136738, 138062
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELSIMULACAOEST", gErr)
        
        Case 136721, 136722, 136724
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO2", gErr)
    
        Case 136725, 136726, 136729
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
            
        Case 136730, 136731, 136734
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO3", gErr)
            
        Case 136735, 136736, 136739
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152351)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
 
Public Function RelCalendarioMaquina_Prepara(ByVal objCTMaquina As ClassCTMaquinas) As Long
'Monta a tabela temporária para o relatório de calendário de máquina
 
Dim lErro As Long
Dim lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dtDataFim As Date
Dim dtDataIni As Date
Dim dtData As Date
Dim objFeriado As ClassFeriado
Dim objMaquina As New ClassMaquinas
Dim objCT As New ClassCentrodeTrabalho
Dim iDiaSemana As Integer
Dim iQtdMaq(1 To 7) As Integer
Dim dQtdHoras(1 To 7) As Double
Dim iFeriado(1 To 7) As Integer
Dim dtDataAux(1 To 7) As Date
Dim sNomeFeriado(1 To 7) As String
Dim bTrocouSemana As Boolean
Dim bTrocouMes As Boolean
Dim bUltima As Boolean
Dim sMes As String
Dim iMes As Integer
Dim sMesNome As String
Dim iSeq As Integer
 
On Error GoTo Erro_RelCalendarioMaquina_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 136900
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136901
    Next
    
    'Começa no dia primeiro do mês e termina um ano depois
    dtDataIni = StrParaDate("01/" & Month(gdtDataAtual) & "/" & Year(gdtDataAtual))
    dtDataFim = DateAdd("yyyy", 1, dtDataIni)
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_CTMAQDISP", lNumIntRel)
    If lErro <> SUCESSO Then gError 136902
    
    objMaquina.lNumIntDoc = objCTMaquina.lNumIntDocMaq
    
    'Le a máquina
    lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
    If lErro <> SUCESSO And lErro <> 106353 Then gError 136903
   
    objCT.lNumIntDoc = objCTMaquina.lNumIntDocCT
    
    'Le o Centro de Trabalho
    lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
    If lErro <> SUCESSO And lErro <> 134590 Then gError 136904
   
    'Insere em RelCTMaqDisp
    lErro = Comando_Executar(alComando(0), "INSERT INTO RelCTMaqDisp (NumIntRel,CodCT, NomeCT,CodMaq,NomeMaq) VALUES (?,?,?,?,?)", lNumIntRel, objCT.lCodigo, objCT.sNomeReduzido, objMaquina.iCodigo, objMaquina.sNomeReduzido)
    If lErro <> SUCESSO Then gError 136905
   
    dtData = dtDataIni
    bTrocouSemana = False
    bTrocouMes = False
    bUltima = False
    
    'Zera array
    For iIndice = 1 To 7
        iQtdMaq(iIndice) = 0
        dQtdHoras(iIndice) = 0
        iFeriado(iIndice) = 2
        sNomeFeriado(iIndice) = ""
        dtDataAux(iIndice) = dtData
    Next
            
    Do While dtData < dtDataFim
       
        iDiaSemana = Weekday(dtData)
        dtDataAux(iDiaSemana) = dtData
        iFeriado(iDiaSemana) = 0
        
        iMes = Month(dtData)
        
        Call MesNome(iMes, sMesNome)
        
        If iMes >= 10 Then
            sMes = CStr(iMes) & SEPARADOR & sMesNome
        Else
            sMes = "0" & CStr(iMes) & SEPARADOR & sMesNome
        End If
    
        Set objFeriado = New ClassFeriado
    
        objFeriado.dtData = dtData
        objFeriado.iFilialEmpresa = objCT.iFilialEmpresa
        
        lErro = CF("Feriado_Le", objFeriado)
        If lErro <> SUCESSO And lErro <> 43379 Then gError 136906
    
        If lErro = SUCESSO Then
            iFeriado(iDiaSemana) = 1
            sNomeFeriado(iDiaSemana) = objFeriado.sDescricao
        End If
        
        'Obtem a quantidade de máquinas e a quantidade de horas disponíveis
        lErro = Verifica_Maquinas_Disponiveis_Dia(objCTMaquina, dtData, objCT.iFilialEmpresa, iQtdMaq(iDiaSemana), dQtdHoras(iDiaSemana))
        If lErro <> SUCESSO Then gError 136907
           
        'Incrementa a data
        dtData = DateAdd("d", 1, dtData)
    
        'Se é domingo então mudou a semana
        If Weekday(dtData) = vbSunday Then
            bTrocouSemana = True
        Else
            bTrocouSemana = False
        End If
        
        'Se é dia primeiro, então mudou o mês
        If Day(dtData) = 1 Then
            bTrocouMes = True
        Else
            bTrocouMes = False
        End If
         
        'Se mudou o mês ou mudou a semana ou é o último dia
        If bTrocouMes Or bTrocouSemana Or bUltima Then
        
            iSeq = iSeq + 1
       
            'Insere em RelCTMaqDisp
            lErro = Comando_Executar(alComando(0), "INSERT INTO RelCTMaqDispItens (NumIntRel,Seq,Mes," & _
                                                "Data1,Feriado1,NomeFeriado1,QtdMaq1,Horas1," & _
                                                "Data2,Feriado2,NomeFeriado2,QtdMaq2,Horas2," & _
                                                "Data3,Feriado3,NomeFeriado3,QtdMaq3,Horas3," & _
                                                "Data4,Feriado4,NomeFeriado4,QtdMaq4,Horas4," & _
                                                "Data5,Feriado5,NomeFeriado5,QtdMaq5,Horas5," & _
                                                "Data6,Feriado6,NomeFeriado6,QtdMaq6,Horas6," & _
                                                "Data7,Feriado7,NomeFeriado7,QtdMaq7,Horas7) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntRel, iSeq, sMes, _
                                                dtDataAux(1), iFeriado(1), sNomeFeriado(1), iQtdMaq(1), dQtdHoras(1), _
                                                dtDataAux(2), iFeriado(2), sNomeFeriado(2), iQtdMaq(2), dQtdHoras(2), _
                                                dtDataAux(3), iFeriado(3), sNomeFeriado(3), iQtdMaq(3), dQtdHoras(3), _
                                                dtDataAux(4), iFeriado(4), sNomeFeriado(4), iQtdMaq(4), dQtdHoras(4), _
                                                dtDataAux(5), iFeriado(5), sNomeFeriado(5), iQtdMaq(5), dQtdHoras(5), _
                                                dtDataAux(6), iFeriado(6), sNomeFeriado(6), iQtdMaq(6), dQtdHoras(6), _
                                                dtDataAux(7), iFeriado(7), sNomeFeriado(7), iQtdMaq(7), dQtdHoras(7))
            If lErro <> SUCESSO Then gError 136908
 
            'Zera array
            For iIndice = 1 To 7
                iQtdMaq(iIndice) = 0
                dQtdHoras(iIndice) = 0
                iFeriado(iIndice) = 2
                sNomeFeriado(iIndice) = ""
                dtDataAux(iIndice) = dtData
            Next
        
        End If
        
        'Se for o últmo dia força a gravação
        If dtData = dtDataFim Then bUltima = True
    
    Loop
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 136909
 
    lErro = objRelatorio.ExecutarDireto("Caledário - Máquina", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 136910
 
    RelCalendarioMaquina_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelCalendarioMaquina_Prepara:
 
    RelCalendarioMaquina_Prepara = gErr
 
    Select Case gErr
 
        Case 136900
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 136901
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 136902 To 136904, 136906 To 136907, 136910
 
        Case 136905
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCTMAQDISP", gErr)
        
        Case 136908
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCTMAQDISPITENS", gErr)
        
        Case 136909
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152352)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function ApontamentoProducao_Grava(ByVal colPO As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objPO As New ClassPlanoOperacional
Dim objApontamentoProducao As ClassApontamentoProducao
Dim objApontamentoAnterior As ClassApontamentoProducao

On Error GoTo Erro_ApontamentoProducao_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137723
    
    'Para cada plano operacional na colecao
    For Each objPO In colPO
    
        'setar o Apontamento de Produção
        Set objApontamentoProducao = objPO.objApontamento
        
        Set objApontamentoAnterior = New ClassApontamentoProducao
        
        objApontamentoAnterior.lNumIntDocPO = objApontamentoProducao.lNumIntDocPO
        
        'Le o Apontamento para ver se existe pelo NumIntDocPO
        lErro = CF("ApontamentoProducao_Le", objApontamentoAnterior)
        If lErro <> SUCESSO And lErro <> 137717 Then gError 137724
        
        'Se existe... Apaga
        If lErro = SUCESSO Then
        
            lErro = ApontamentoProducao_Exclui_EmTrans(objApontamentoProducao)
            If lErro <> SUCESSO Then gError 137725
        
        End If
        
        'e Se tem Apontamento... Insere
        If objPO.iTemApontamento Then

            lErro = CF("ApontamentoProducao_Grava_EmTrans", objApontamentoProducao)
            If lErro <> SUCESSO Then gError 137726
        
        End If
    
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137727

    ApontamentoProducao_Grava = SUCESSO

    Exit Function

Erro_ApontamentoProducao_Grava:

    ApontamentoProducao_Grava = gErr

    Select Case gErr

        Case 137723
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137724, 137725, 137726

        Case 137727
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152353)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function ApontamentoProducao_Grava_EmTrans(ByVal objApontamentoProducao As ClassApontamentoProducao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_ApontamentoProducao_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137824
    Next

    'Le a tabelaApontamentoProducao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocPO FROM ApontamentoProducao WHERE NumIntDocPO= ? ", _
                0, lNumIntDoc, objApontamentoProducao.lNumIntDocPO)
    If lErro <> AD_SQL_SUCESSO Then gError 137825

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137826

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 137827
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ApontamentoProducao SET Concluido= ? , PercConcluido= ? , Quantidade= ? , Data= ? , Observacao= ? ", alComando(0), _
                objApontamentoProducao.iConcluido, objApontamentoProducao.dPercConcluido, objApontamentoProducao.dQuantidade, objApontamentoProducao.dtData, objApontamentoProducao.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 137828

    Else

        lErro = Comando_Executar(alComando(1), "INSERT INTO ApontamentoProducao(NumIntDocPO, Concluido, PercConcluido, Quantidade, Data, Observacao) VALUES (?,?,?,?,?,?)", _
                objApontamentoProducao.lNumIntDocPO, objApontamentoProducao.iConcluido, objApontamentoProducao.dPercConcluido, objApontamentoProducao.dQuantidade, objApontamentoProducao.dtData, objApontamentoProducao.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137829

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ApontamentoProducao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_ApontamentoProducao_Grava_EmTrans:

    ApontamentoProducao_Grava_EmTrans = gErr

    Select Case gErr

        Case 137824
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137825, 137826
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTAMENTOPRODUCAO", gErr)

        Case 137827
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTAMENTOPRODUCAO", gErr)

        Case 137828
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_APONTAMENTOPRODUCAO", gErr)

        Case 137829
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_APONTAMENTOPRODUCAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152354)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ApontamentoProducao_Exclui(ByVal objApontamentoProducao As ClassApontamentoProducao) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ApontamentoProducao_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 137830

    lErro = CF("ApontamentoProducao_Exclui_EmTrans", objApontamentoProducao)
    If lErro <> SUCESSO Then gError 137831

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 137832

    ApontamentoProducao_Exclui = SUCESSO

    Exit Function

Erro_ApontamentoProducao_Exclui:

    ApontamentoProducao_Exclui = gErr

    Select Case gErr

        Case 137830
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 137831

        Case 137832
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152355)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function ApontamentoProducao_Exclui_EmTrans(ByVal objApontamentoProducao As ClassApontamentoProducao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_ApontamentoProducao_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 137833
    Next

    'Le a tabelaApontamentoProducao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM ApontamentoProducao WHERE NumIntDocPO= ?  ", _
                0, iAux, objApontamentoProducao.lNumIntDocPO)
    If lErro <> AD_SQL_SUCESSO Then gError 137834

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137835

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 137836

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 137837

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ApontamentoProducao", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 137838

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ApontamentoProducao_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_ApontamentoProducao_Exclui_EmTrans:

    ApontamentoProducao_Exclui_EmTrans = gErr

    Select Case gErr

        Case 137833
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137834, 137835
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTAMENTOPRODUCAO", gErr)

        Case 137836
            Call Rotina_Erro(vbOKOnly, "ERRO_APONTAMENTOPRODUCAO_NAO_CADASTRADO", gErr, objApontamentoProducao.lNumIntDocPO)

        Case 137837
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_APONTAMENTOPRODUCAO", gErr)

        Case 137838
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APONTAMENTOPRODUCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152356)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PMP_Exclui_EmTrans2(ByVal objPMP As ClassPMP) As Long
'Exclui pelo código de geração um plano mestre gerado
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
 
On Error GoTo Erro_PMP_Exclui_EmTrans2
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138015
    Next
    
    'Le a PlanoMestreProducao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM PlanoMestreProducao WHERE CodGeracao = ? ", _
                0, iAux, objPMP.lCodGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 138016
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138017
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 138018
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoMestreProducao", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 138019
        
        lErro = PMPItens_Exclui_EmTrans2(objPMP)
        If lErro <> SUCESSO Then gError 138020
            
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PMP_Exclui_EmTrans2 = SUCESSO
 
    Exit Function
 
Erro_PMP_Exclui_EmTrans2:
 
    PMP_Exclui_EmTrans2 = gErr
 
    Select Case gErr
 
        Case 138015
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 138016, 138017
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAO", gErr)
 
        Case 138018
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOMESTREPRODUCAO", gErr)
 
        Case 138019
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOMESTREPRODUCAO", gErr)
            
        Case 138020
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152357)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function PMP_Exclui(ByVal objPMP As ClassPMP) As Long
'Abre a Transação e Chama a PMP_Exclui_EmTrans
 
Dim lErro As Long
Dim objPMPItens As New ClassPMPItens
Dim objPO As New ClassPlanoOperacional
Dim objApontamentoProducao As ClassApontamentoProducao
Dim bAchou As Boolean
Dim bAchouApontamento As Boolean
Dim bProcedeExclusao As Boolean
Dim lTransacao As Long
Dim vbMsgRes As VbMsgBoxResult
 
On Error GoTo Erro_PMP_Exclui

    bProcedeExclusao = True
    
    bAchou = False
    bAchouApontamento = False
        
    For Each objPMPItens In objPMP.colItens

        For Each objPO In objPMPItens.colPO
    
            Set objApontamentoProducao = New ClassApontamentoProducao
            
            objApontamentoProducao.lNumIntDocPO = objPO.lNumIntDoc
            
            lErro = CF("ApontamentoProducao_Le", objApontamentoProducao)
            If lErro <> SUCESSO And lErro <> 137717 Then gError 137909
            
            If lErro = SUCESSO Then
                bAchouApontamento = True
                bAchou = True
                Exit For
            End If
        
        Next
        
        If bAchou Then Exit For
    
    Next
    
    If bAchouApontamento Then
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PMPITENS_COM_APONTAMENTO")
    
        If vbMsgRes = vbNo Then
            bProcedeExclusao = False
        End If
    
    End If
            
    If bProcedeExclusao Then

        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 138021
        
        'Chama a Rotina que Exclui dentro de transação
        lErro = CF("PMP_Exclui_EmTrans2", objPMP)
        If lErro <> SUCESSO Then gError 138022
    
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 138023
    
    End If
    
    PMP_Exclui = SUCESSO
 
    Exit Function
 
Erro_PMP_Exclui:
 
    PMP_Exclui = gErr
 
    Select Case gErr
    
        Case 137909
            'erro tratado nas rotina chamada
 
        Case 138021
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 138022 'Tratado na Rotina Chamada

        Case 138023
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152358)
 
    End Select
 
    'Desfaz Transação
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Function PMPItens_Exclui_EmTrans2(ByVal objPMP As ClassPMP) As Long
'Exclui os itens geradas de um plano mestre
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPMPItens As ClassPMPItens
Dim tPMPItem As typePMPItens
 
On Error GoTo Erro_PMPItens_Exclui_EmTrans2
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138024
    Next
    
    tPMPItem.sCodOPOrigem = String(STRING_OPCODIGO, 0)
    tPMPItem.sProduto = String(STRING_PRODUTO, 0)
    
    'Le a PlanoMestreProducaoItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, CodOPOrigem, Produto, FilialEmpresa FROM PlanoMestreProducaoItens WHERE CodGeracao = ? ", _
                0, tPMPItem.lNumIntDoc, tPMPItem.sCodOPOrigem, tPMPItem.sProduto, tPMPItem.iFilialEmpresa, objPMP.lCodGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 138025
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138026
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 138027
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoMestreProducaoItens ", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 138028
        
        Set objPMPItens = New ClassPMPItens
        
        objPMPItens.sCodOPOrigem = tPMPItem.sCodOPOrigem
        objPMPItens.iFilialEmpresa = tPMPItem.iFilialEmpresa
        objPMPItens.sProduto = tPMPItem.sProduto
        objPMPItens.lNumIntDoc = tPMPItem.lNumIntDoc
        
        If Len(objPMP.sVersao) = 0 Then
                
            lErro = PO_Exclui_EmTrans2(objPMPItens)
            If lErro <> SUCESSO Then gError 138029
        
        Else
        
            lErro = PO_Exclui_EmTrans(objPMPItens)
            If lErro <> SUCESSO Then gError 137919
        
        End If
            
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138030
            
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PMPItens_Exclui_EmTrans2 = SUCESSO
 
    Exit Function
 
Erro_PMPItens_Exclui_EmTrans2:
 
    PMPItens_Exclui_EmTrans2 = gErr
 
    Select Case gErr
 
        Case 138024
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 138025, 138025, 138030
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAOITENS", gErr)
 
        Case 138027
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOMESTREPRODUCAOITENS", gErr)
 
        Case 138028
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOMESTREPRODUCAOITENS", gErr)
            
        Case 138029, 137919
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152359)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Function PO_Exclui_EmTrans2(ByVal objPMPItem As ClassPMPItens) As Long
'Exclui as etapas do plano mestre juntamente com as OPs e RCs geradas pelo MRP
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPO As ClassPlanoOperacional
Dim tPO As typePlanoOperacional
Dim objRC As ClassRequisicaoCompras
Dim objOP As ClassOrdemDeProducao
Dim lNumIntDocRC As Long
Dim objApontamentoProducao As ClassApontamentoProducao
Dim iAux As Integer
Dim colPO As New Collection
 
On Error GoTo Erro_PO_Exclui_EmTrans2
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138030
    Next
    
    tPO.sCodOPOrigem = String(STRING_OPCODIGO, 0)
    tPO.sProduto = String(STRING_PRODUTO, 0)
    
    'Le a PlanoOperacional
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, CodOpOrigem, Produto, Nivel FROM PlanoOperacional WHERE NumIntDocPMP = ? ", _
                0, tPO.lNumIntDoc, tPO.sCodOPOrigem, tPO.sProduto, tPO.iNivel, objPMPItem.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 138031
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138032
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 138033
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoOperacional", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 138034
        
        Set objPO = New ClassPlanoOperacional
        
        objPO.lNumIntDoc = tPO.lNumIntDoc
        objPO.sCodOPOrigem = tPO.sCodOPOrigem
        
        colPO.Add objPO
    
        'Busca Proximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138041
    
    Loop
    
    For Each objPO In colPO
    
        lErro = POMaquinas_Exclui_EmTrans(objPO)
        If lErro <> SUCESSO Then gError 138035
        
        lErro = POReqCompra_Exclui_EmTrans(objPO)
        If lErro <> SUCESSO Then gError 138208
        
        If objPO.sCodOPOrigem = objPMPItem.sCodOPOrigem Then
        
            'Verifica se a OP tem Reqs associadas a ela
            lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM RequisicaoCompra WHERE OPCodigo = ? AND FilialEmpresa = ?", lNumIntDocRC, objPO.sCodOPOrigem, objPMPItem.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 138036
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138037
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                Set objRC = New ClassRequisicaoCompras
            
                objRC.lNumIntDoc = lNumIntDocRC
                
                lErro = CF("RequisicaoCompra_Exclui_EmTrans", objRC)
                If lErro <> SUCESSO Then gError 138038
            
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138039
            
            Loop
            
        Else
        
            'Exclui as sub OPs geradas pelo sistema
            Set objOP = New ClassOrdemDeProducao
            
            objOP.iFilialEmpresa = objPMPItem.iFilialEmpresa
            objOP.sCodigo = objPO.sCodOPOrigem
            
            lErro = CF("OrdemDeProducao_Exclui_EmTrans", objOP)
            If lErro <> SUCESSO And lErro <> 21936 Then gError 138040
                    
        End If
        
        Set objApontamentoProducao = New ClassApontamentoProducao
        
        objApontamentoProducao.lNumIntDocPO = objPO.lNumIntDoc
        
        'Le a tabelaApontamentoProducao
        lErro = Comando_ExecutarPos(alComando(1), "SELECT 1 FROM ApontamentoProducao WHERE NumIntDocPO= ?  ", _
                    0, iAux, objApontamentoProducao.lNumIntDocPO)
        If lErro <> AD_SQL_SUCESSO Then gError 137904
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137905
    
        'Se existir ... Exclui
        If lErro <> AD_SQL_SEM_DADOS Then
            
            lErro = CF("ApontamentoProducao_Exclui_EmTrans", objApontamentoProducao)
            If lErro <> SUCESSO Then gError 137906
            
        End If
                        
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PO_Exclui_EmTrans2 = SUCESSO
 
    Exit Function
 
Erro_PO_Exclui_EmTrans2:
 
    PO_Exclui_EmTrans2 = gErr
 
    Select Case gErr
    
        Case 137904, 137905
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTAMENTOPRODUCAO", gErr)
 
        Case 138030
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 138031, 138032, 138041
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOOPERACIONAL", gErr)
 
        Case 138033
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOOPERACIONAL", gErr)
 
        Case 138034
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOOPERACIONAL", gErr)
            
        Case 138035, 138038, 138040, 138208, 137906
        
        Case 138036, 138037, 138039
             Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152360)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function PMP_Grava(ByVal objPMP As ClassPMP) As Long
'Abre a Transação e Chama a PMP_Grava_EmTrans
 
Dim lErro As Long
Dim objPMPItens As ClassPMPItens
Dim objPMPBD As New ClassPMP
Dim objPMPItemBD As ClassPMPItens
Dim colItensExcluidos As Collection
Dim objPO As New ClassPlanoOperacional
Dim objApontamentoProducao As ClassApontamentoProducao
Dim bAchou As Boolean
Dim bAchouApontamento As Boolean
Dim bProcedeGravacao As Boolean
Dim lTransacao As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_PMP_Grava

    bProcedeGravacao = True

    lErro = CF("PlanoMestreProducao_Le", objPMPBD)
    If lErro <> SUCESSO And lErro <> 136298 Then gError 137907
    
    'Para cada item do Plano Mestre
    For Each objPMPItemBD In objPMPBD.colItens
    
        bAchou = False
    
        For Each objPMPItens In objPMP.colItens
            If objPMPItens.sCodOPOrigem = objPMPItemBD.sCodOPOrigem And objPMPItens.sProduto = objPMPItemBD.sProduto Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            bAchouApontamento = False
       
            For Each objPO In objPMPItemBD.colPO
        
                Set objApontamentoProducao = New ClassApontamentoProducao
                
                objApontamentoProducao.lNumIntDocPO = objPO.lNumIntDoc
                
                lErro = CF("ApontamentoProducao_Le", objApontamentoProducao)
                If lErro <> SUCESSO And lErro <> 137717 Then gError 137908
                
                If lErro = SUCESSO Then
                    bAchouApontamento = True
                    Exit For
                End If
            
            Next
            
            If bAchouApontamento Then
            
                'Pergunta ao usuário se confirma a exclusão
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PMPITENS_COM_APONTAMENTO")
            
                If vbMsgRes = vbNo Then
                    bProcedeGravacao = False
                End If
                Exit For
            
            End If
            
        End If
    
    Next

    If bProcedeGravacao Then
    
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 138042
            
        'Chama a Rotina que Gera dentro de transação
        lErro = PMP_Grava_EmTrans(objPMP, objPMPBD)
        If lErro <> SUCESSO Then gError 138043
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 138044
    
    End If
    
    PMP_Grava = SUCESSO
 
    Exit Function
 
Erro_PMP_Grava:
 
    PMP_Grava = gErr
 
    Select Case gErr
 
        Case 138042
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 137907, 137908
            'erros tratados nas rotinas chamadas
            
        Case 138043 'Tratado na Rotina Chamada

        Case 138044
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152361)
 
    End Select
 
    'Desfaz Transação
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function PMP_Grava_EmTrans(ByVal objPMP As ClassPMP, ByVal objPMPBD As ClassPMP) As Long
'Grava um Plano Mestre no BD a partir de um já existente
'Exclui os itens que estão no BD e não estão no obj passado
 
Dim lErro As Long
Dim objPMPItens As ClassPMPItens
'Dim objPMPBD As New ClassPMP
Dim objPMPItemBD As ClassPMPItens
Dim bAchou As Boolean
 
On Error GoTo Erro_PMP_Grava_EmTrans
    
'    lErro = CF("PlanoMestreProducao_Le", objPMPBD)
'    If lErro <> SUCESSO And lErro <> 136298 Then gError 138046

    'Para cada item novo do Plano Mestre
    For Each objPMPItemBD In objPMPBD.colItens
    
        bAchou = False
    
        For Each objPMPItens In objPMP.colItens
            If objPMPItens.sCodOPOrigem = objPMPItemBD.sCodOPOrigem And objPMPItens.sProduto = objPMPItemBD.sProduto Then
                bAchou = True
                Exit For
            End If
        Next
        
        If Not bAchou Then
            lErro = PMPItens_Exclui_EmTrans3(objPMPItemBD)
            If lErro <> SUCESSO Then gError 138047
        End If
    
    Next
  
    PMP_Grava_EmTrans = SUCESSO
 
    Exit Function
 
Erro_PMP_Grava_EmTrans:
 
    PMP_Grava_EmTrans = gErr
 
    Select Case gErr
 
        Case 138046, 138047
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152362)
 
    End Select
 
    Exit Function
 
End Function

Function PMPItens_Exclui_EmTrans3(ByVal objPMPItem As ClassPMPItens) As Long
'Exclui um item do plano mestre gerado
 
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_PMPItens_Exclui_EmTrans3
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138052
    Next
    
    'Le a PlanoMestreProducaoItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PlanoMestreProducaoItens WHERE NumIntDoc = ? ", _
                0, lNumIntDoc, objPMPItem.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 138053
 
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138054
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 138055
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoMestreProducaoItens ", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 138056
                       
        'Exclui as etapas geradas, excluindo também as SubOps e as ReqCompras
        lErro = PO_Exclui_EmTrans2(objPMPItem)
        If lErro <> SUCESSO Then gError 138057
                        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    PMPItens_Exclui_EmTrans3 = SUCESSO
 
    Exit Function
 
Erro_PMPItens_Exclui_EmTrans3:
 
    PMPItens_Exclui_EmTrans3 = gErr
 
    Select Case gErr
 
        Case 138052
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 138053, 138054
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAOITENS", gErr)
 
        Case 138055
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOMESTREPRODUCAOITENS", gErr)
 
        Case 138056
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOMESTREPRODUCAOITENS", gErr)
            
        Case 138057
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152363)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function PlanoOperacional_Remove_Descendentes(ByVal colPO As Collection, ByVal colPONovo As Collection, ByVal objPOPai As ClassPlanoOperacional) As Long
'Gera uma nova coleção sem os descendentes do objPO passado
 
Dim lErro As Long
Dim objPOAux As ClassPlanoOperacional
 
On Error GoTo Erro_PlanoOperacional_Remove_Descendentes
 
    'Para cada Item da Coleção
    For Each objPOAux In colPO
        
        If objPOAux.iSeq = objPOPai.iSeq And objPOAux.iNivel = objPOPai.iNivel Then
            Exit For
        End If
                
        colPONovo.Add objPOAux

    Next
 
    PlanoOperacional_Remove_Descendentes = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Remove_Descendentes:
 
    PlanoOperacional_Remove_Descendentes = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152364)
 
    End Select
 
    Exit Function
 
End Function

Public Function RelProducaoComEstoque_Prepara(ByVal objProdEstSeleciona As ClassProdEstSeleciona, lNumIntRel As Long) As Long
'Monta a tabela temporária para o relatório sde Capacidade com o que se tem em estoque
 
Dim lErro As Long
Dim lTransacao As Long
Dim objProduto As ClassProduto
Dim colProduto As New Collection
Dim iSeq As Integer
 
On Error GoTo Erro_RelProducaoComEstoque_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 138165
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_PRODEST", lNumIntRel)
    If lErro <> SUCESSO Then gError 138166
    
    'TEM QUE PEGAR OS PRODUTOS COM OS FILTROS
    lErro = CF("Produto_Le_Faixa", objProdEstSeleciona.sProdutoInicial, objProdEstSeleciona.sProdutoFinal, colProduto)
    If lErro <> SUCESSO Then gError 138167
    
    For Each objProduto In colProduto
    
        If objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And objProduto.iAtivo = PRODUTO_ATIVO Then
        
            iSeq = iSeq + 1
        
            lErro = RelProducaoComEstoque_Prepara_EmTrans(objProdEstSeleciona, objProduto, lNumIntRel, iSeq)
            If lErro <> SUCESSO Then gError 138168
            
        End If
    
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 138169
 
    RelProducaoComEstoque_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelProducaoComEstoque_Prepara:
 
    RelProducaoComEstoque_Prepara = gErr
 
    Select Case gErr
 
        Case 138165
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 138166 To 138168
 
        Case 138169
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152365)
 
    End Select
 
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelProducaoComEstoque_Prepara_EmTrans(ByVal objProdEstSeleciona As ClassProdEstSeleciona, ByVal objProduto As ClassProduto, ByVal lNumIntRel As Long, ByVal iSeq As Integer) As Long
'Monta a tabela temporária para o relatório sde Capacidade com o que se tem em estoque
 
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim dMaximoProducao As Double
Dim objKit As New ClassKit
Dim objProdutoKit As New ClassProdutoKit
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim dEstoqueProjetado As Double
 
On Error GoTo Erro_RelProducaoComEstoque_Prepara_EmTrans
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138192
    Next
    
'    objEstoqueProduto.sProduto = objProduto.sCodigo
'    objEstoqueProduto.iFilialEmpresa = objProdEstSeleciona.iFilialEmpresa
'
'    'Lê as quantidades do produtos projatada para data
'    lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objProdEstSeleciona.dtData, dEstoqueProjetado)
'    If lErro <> SUCESSO Then gError 138204
    
    objKit.sProdutoRaiz = objProduto.sCodigo
    
    'Le para ver se é um Novo Kit
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 138201
    
    objProdutoKit.dQuantidade = 1
    objProdutoKit.sUnidadeMed = objProduto.sSiglaUMEstoque
    
    If lErro = SUCESSO Then
       
        objEstoqueProduto.sProduto = objProduto.sCodigo
        objEstoqueProduto.iFilialEmpresa = objProdEstSeleciona.iFilialEmpresa
                    
        'Lê as quantidades do produtos projatada para data
        lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objProdEstSeleciona.dtData, dEstoqueProjetado)
        If lErro <> SUCESSO Then gError 138204
       
        lErro = RelProducaoComEstoque_Prepara_Producao(objProdEstSeleciona, objProdutoKit, objKit, dMaximoProducao, 1)
        If lErro <> SUCESSO Then gError 138202
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelProdEst (NumIntRel, Seq, Produto, Descricao, UMEstoque, EstoqueAtual, EstoqueProjDisp, CapacidadeProd, DataInicioOP)VALUES (?,?,?,?,?,?,?,?,?)", lNumIntRel, iSeq, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, objEstoqueProduto.dQuantDisponivel, dEstoqueProjetado, dMaximoProducao, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 138203
   
    End If
   
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelProducaoComEstoque_Prepara_EmTrans = SUCESSO
    
    Exit Function
 
Erro_RelProducaoComEstoque_Prepara_EmTrans:
 
    RelProducaoComEstoque_Prepara_EmTrans = gErr
 
    Select Case gErr
        
        Case 138192
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 138201 To 138204
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152366)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function RelProducaoComEstoque_Prepara_Producao(ByVal objProdEstSeleciona As ClassProdEstSeleciona, ByVal objProdutoKitPai As ClassProdutoKit, ByVal objKit As ClassKit, dMaximo As Double, ByVal ilvFunc As Integer) As Long
'Retorno o que se pode produzir a mais do Pai com os filhos recursivamente
 
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objKitFilho As ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim dMaximoFilho As Double
Dim dMaximoAux As Double
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objProdutoFilial As ClassProdutoFilial
Dim dEstoqueProjetado As Double
Dim dFatorConversao As Double
Dim dFatorConversaoKit As Double
Dim objProduto As ClassProduto
Dim dFatorKit As Double
 
On Error GoTo Erro_RelProducaoComEstoque_Prepara_Producao

    If ilvFunc >= 20 Or objProdutoKitPai.dQuantidade <= 0.0001 Then
        RelProducaoComEstoque_Prepara_Producao = SUCESSO
        Exit Function
    End If
 
    'tem que ler o proprio cara e os seus filhos
    lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
    If lErro <> SUCESSO And lErro <> 103203 Then gError 138193
 
    'percorre os elementos componentes do Kit Lido
    For Each objProdutoKit In objKit.colComponentes
 
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objProdutoKit.sProduto
 
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 138194
 
        If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
        
            'Descobre o fator de conversao da UM
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProdutoKitPai.sUnidadeMed, dFatorConversaoKit)
            If lErro <> SUCESSO Then gError 138195
        
            dFatorKit = (objProdutoKitPai.dQuantidade * dFatorConversaoKit) / objProdutoKit.dQuantidade
        
        Else
        
            objProdutoKit.dQuantidade = objProdutoKit.dQuantidade * dFatorKit
    
            If objProdutoKit.dQuantidade > 0.0001 Then
    
                Set objKitFilho = New ClassKit
                Set objEstoqueProduto = New ClassEstoqueProduto
                Set objProdutoFilial = New ClassProdutoFilial
                
                objEstoqueProduto.sProduto = objProdutoKit.sProduto
                objEstoqueProduto.iFilialEmpresa = objProdEstSeleciona.iFilialEmpresa
                            
                dEstoqueProjetado = 0
                            
                'Lê as quantidades do produtos projatada para data
                lErro = CF("EstoqueDisponivelProjetado_Le", objEstoqueProduto, objProdutoFilial, objProdEstSeleciona.dtData, dEstoqueProjetado)
                If lErro <> SUCESSO Then gError 138196
            
                'Descobre o fator de conversao da UM de Venda p/UM de estoque do produto raiz
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
                If lErro <> SUCESSO Then gError 138197
        
                dMaximoAux = dEstoqueProjetado / ((objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda)) * dFatorConversao)
        
                objKitFilho.sProdutoRaiz = objProdutoKit.sProduto
                
                'Le para ver se é um Novo Kit
                lErro = CF("Kit_Le_Padrao", objKitFilho)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 138198
                
                'Se o produto for um Kit, então leva a recursividade em diante
                If lErro = SUCESSO Then
                            
                    lErro = RelProducaoComEstoque_Prepara_Producao(objProdEstSeleciona, objProdutoKit, objKitFilho, dMaximoFilho, ilvFunc + 1)
                    If lErro <> SUCESSO Then gError 138200
                    
                    'Soma o que se tem com o que se pode produzir usando os filhos
                    dMaximoAux = dMaximoAux + dMaximoFilho
                                       
                End If
                
                If dMaximo = 0 Then dMaximo = dMaximoAux
                
                'Pega o máximo que se pode produzir do produto pai com os filhos
                If dMaximo > dMaximoAux Then dMaximo = dMaximoAux
                
            End If
 
        End If
 
    Next
    
    RelProducaoComEstoque_Prepara_Producao = SUCESSO
 
    Exit Function
 
Erro_RelProducaoComEstoque_Prepara_Producao:
 
    RelProducaoComEstoque_Prepara_Producao = gErr
 
    Select Case gErr
        
        Case 138193 To 138200
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152367)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function PlanoOperacional_Atualiza_Descendentes(ByVal colPO As Collection, ByVal colPONovo As Collection, ByVal objPO As ClassPlanoOperacional) As Long
'Gera uma nova coleção com a atualização das etapas que não deviam ter sido modificadas
 
Dim lErro As Long
Dim iIndice As Integer
Dim objPOFilho As ClassPlanoOperacional
Dim objPOAux As ClassPlanoOperacional
Dim bDescendente As Boolean
Dim bAchou As Boolean
Dim bAchou2 As Boolean
 
On Error GoTo Erro_PlanoOperacional_Atualiza_Descendentes
 
    iIndice = 0
    bAchou = False
 
    'Para cada Item da Coleo
    For Each objPOFilho In colPO
    
        iIndice = iIndice + 1
    
        bDescendente = False
        
        'Se não for a própria etapa
        If objPO.lNumIntDocOper <> objPOFilho.lNumIntDocOper Then
        
            'Verifica se  descentente da etapa que foi mexida
            lErro = PlanoOperacional_Verifica_Descendencia(colPO, objPO, objPOFilho, bDescendente)
            If lErro <> SUCESSO Then gError 138125
            
            bAchou2 = False
            
            For Each objPOAux In colPONovo
                
                If objPOAux.lNumIntDocOper = objPOFilho.lNumIntDocOper Then
                
                    If Not bDescendente Then
                        Set objPOAux.colAlocacaoMaquinas = objPOFilho.colAlocacaoMaquinas
                        Set objPOAux.colOPFilhas = objPOFilho.colOPFilhas
                        Set objPOAux.colRCFilhas = objPOFilho.colRCFilhas
                        objPOAux.dQuantidade = objPOFilho.dQuantidade
                        objPOAux.dtDataFim = objPOFilho.dtDataFim
                        objPOAux.dtDataInicio = objPOFilho.dtDataInicio
                        objPOAux.lNumIntDoc = objPOFilho.lNumIntDoc
                        objPOAux.lNumIntDocPOPai = objPOFilho.lNumIntDocPOPai
                    End If
                
                    bAchou2 = True
                    Exit For
                End If
            
            Next
                
            'Se no for descendente e for posterior na coleção
            If Not bDescendente And bAchou Then
                        
                If Not bAchou2 Then
                    colPONovo.Add objPOFilho
                End If
                
            End If
                        
        Else
            bAchou = True
        End If
            
    Next
 
    PlanoOperacional_Atualiza_Descendentes = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Atualiza_Descendentes:
 
    PlanoOperacional_Atualiza_Descendentes = gErr
 
    Select Case gErr
 
        Case 138125
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152368)
 
    End Select
 
    Exit Function
 
End Function

Public Function PlanoOperacional_Verifica_Descendencia(ByVal colPO As Collection, ByVal objPO As ClassPlanoOperacional, ByVal objPOFilho As ClassPlanoOperacional, bDescendente As Boolean) As Long
'Verifica se uma etapa é descendente de outra
 
Dim lErro As Long
Dim objPOPai As ClassPlanoOperacional
 
On Error GoTo Erro_PlanoOperacional_Verifica_Descendencia

    If objPO.lNumIntDoc = objPOFilho.lNumIntDocPOPai Then
    
        bDescendente = True
    
    Else
    
        For Each objPOPai In colPO
        
            If objPOPai.lNumIntDoc = objPOFilho.lNumIntDocPOPai Then
            
                'Verifica se é descentente da etapa que foi mexida
                lErro = PlanoOperacional_Verifica_Descendencia(colPO, objPO, objPOPai, bDescendente)
                If lErro <> SUCESSO Then gError 138250
            
                Exit For
            
            End If
        
        Next
    
    End If
 
    PlanoOperacional_Verifica_Descendencia = SUCESSO
 
    Exit Function
 
Erro_PlanoOperacional_Verifica_Descendencia:
 
    PlanoOperacional_Verifica_Descendencia = gErr
 
    Select Case gErr
 
        Case 138250
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152369)
 
    End Select
 
    Exit Function
 
End Function
 
'------------------------------------------------------------------
Function Gera_Operacoes_ItemOP(ByVal objPO As ClassPlanoOperacional, ByVal objItemOPOrigem As ClassItemOP, ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim colOPOperacoes As Collection
Dim objOPOperacoes As New ClassOrdemProducaoOperacoes
Dim dFatorQtde As Double

On Error GoTo Erro_Gera_Operacoes_ItemOP

    'Calcula fator de correção da quantidade em virtude da quantidade a ser produzida
    dFatorQtde = objPO.dQuantidade / objItemOPOrigem.dQuantidade
    
    Set colOPOperacoes = New Collection
    
    'gera as operações da OP filha conforme a OP de Origem
    lErro = Gera_Operacoes_ItemOP1(objPO, objItemOPOrigem, objItemOP, colOPOperacoes, dFatorQtde, 0, 0, 0, 0)
    If lErro <> SUCESSO Then gError 137933
    
    For Each objOPOperacoes In colOPOperacoes

        objItemOP.colOrdemProducaoOperacoes.Add objOPOperacoes
        
        'ajusta as sequencias dos itens gerados
        Call OrdemDeProducao_Gera_Operacoes2(objItemOP)
    
    Next
                        
    Gera_Operacoes_ItemOP = SUCESSO
 
    Exit Function
 
Erro_Gera_Operacoes_ItemOP:
 
    Gera_Operacoes_ItemOP = gErr
 
    Select Case gErr
    
        Case 137933
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152370)
 
    End Select
 
    Exit Function

End Function

Function Gera_Operacoes_ItemOP1(ByVal objPO As ClassPlanoOperacional, ByVal objItemOPOrigem As ClassItemOP, ByVal objItemOP As ClassItemOP, colOPOperacoes As Collection, ByVal dFatorQtde As Double, ByVal iNivel As Integer, ByVal iNivelRoteiro, iSeqRoteiro As Integer, iSeq As Integer) As Long

Dim lErro As Long
Dim objOPOperacoes As New ClassOrdemProducaoOperacoes
Dim objOPOperacoesAux As New ClassOrdemProducaoOperacoes
Dim objOperacoesTempo As ClassOperacoesTempo
Dim objOPInsumos As New ClassOrdemProducaoInsumos
Dim objOPInsumosAux As New ClassOrdemProducaoInsumos
Dim objItemOPNovo As ClassItemOP

On Error GoTo Erro_Gera_Operacoes_ItemOP1
    
    iSeqRoteiro = iSeqRoteiro + 1
    
    'Para cada operação da OP Origem
    For Each objOPOperacoes In objItemOPOrigem.colOrdemProducaoOperacoes
    
        'Verifica se é o Produto que está gerando a OP Nova
        If objOPOperacoes.sProduto = objItemOP.sProduto And Trim(objOPOperacoes.sVersao) = Trim(objItemOP.sVersao) Then
            
            'Cria um novo objOPOperacoes
            Set objOPOperacoesAux = New ClassOrdemProducaoOperacoes
            
            iSeq = iSeq + 1
            
            'Copia os dados de objOPOperacoes Original
            objOPOperacoesAux.lNumIntDocCompet = objOPOperacoes.lNumIntDocCompet
            objOPOperacoesAux.lNumIntDocCT = objOPOperacoes.lNumIntDocCT
            objOPOperacoesAux.sObservacao = objOPOperacoes.sObservacao
            objOPOperacoesAux.sProduto = objOPOperacoes.sProduto
            objOPOperacoesAux.sVersao = objOPOperacoes.sVersao
            objOPOperacoesAux.dQuantidade = objOPOperacoes.dQuantidade * dFatorQtde
            objOPOperacoesAux.sUMedida = objOPOperacoes.sUMedida
            objOPOperacoesAux.iIgnoraTaxaProducao = objOPOperacoes.iIgnoraTaxaProducao
            objOPOperacoesAux.iConsideraCarga = objOPOperacoes.iConsideraCarga
            objOPOperacoesAux.iOrigem = objOPOperacoes.iOrigem
            
            'Ajusta as sequencias internas
            objOPOperacoesAux.iNivel = iNivel
            objOPOperacoesAux.iSeq = iSeq
            objOPOperacoesAux.iNivelRoteiro = iNivelRoteiro
            objOPOperacoesAux.iSeqRoteiro = iSeqRoteiro
            
            'e a referencia com a Operação da OP de Origem
            objOPOperacoesAux.lNumIntDocOperOrigem = objOPOperacoes.lNumIntDoc
                        
            'Inclui na Coleção a Operação
            colOPOperacoes.Add objOPOperacoesAux
            
            'Se tiver taxa de produção específica
            If objOPOperacoes.iIgnoraTaxaProducao = MARCADO Then
                                        
                Set objOperacoesTempo = New ClassOperacoesTempo
                
                'copia os dados para cálculo da taxa
                objOperacoesTempo.dLoteMax = objOPOperacoes.objOperacoesTempo.dLoteMax
                objOperacoesTempo.dLoteMin = objOPOperacoes.objOperacoesTempo.dLoteMin
                objOperacoesTempo.dLotePadrao = objOPOperacoes.objOperacoesTempo.dLotePadrao
                objOperacoesTempo.dTempoDescarga = objOPOperacoes.objOperacoesTempo.dTempoDescarga
                objOperacoesTempo.dTempoMovimentacao = objOPOperacoes.objOperacoesTempo.dTempoMovimentacao
                If objOPOperacoes.objOperacoesTempo.iTipo <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
                    objOperacoesTempo.dTempoOperacao = objOPOperacoes.objOperacoesTempo.dTempoOperacao * dFatorQtde
                Else
                    objOperacoesTempo.dTempoOperacao = objOPOperacoes.objOperacoesTempo.dTempoOperacao
                End If
                objOperacoesTempo.dTempoPreparacao = objOPOperacoes.objOperacoesTempo.dTempoPreparacao
                objOperacoesTempo.iTipo = objOPOperacoes.objOperacoesTempo.iTipo
                objOperacoesTempo.sUMTempo = objOPOperacoes.objOperacoesTempo.sUMTempo
                objOperacoesTempo.lNumIntDocMaq = objOPOperacoes.objOperacoesTempo.lNumIntDocMaq
                
                Set objOPOperacoesAux.objOperacoesTempo = objOperacoesTempo
            
            End If
            
            'Para cada Insumo desta Operação
            For Each objOPInsumos In objOPOperacoes.colOPInsumos
                        
                Set objOPInsumosAux = New ClassOrdemProducaoInsumos
                
                'Copia os dados de Insumos da Operação da OP Original
                objOPInsumosAux.sProduto = objOPInsumos.sProduto
                If objOPInsumos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                    objOPInsumosAux.dQuantidade = objOPInsumos.dQuantidade * dFatorQtde
                Else
                    objOPInsumosAux.dQuantidade = objOPInsumos.dQuantidade
                End If
                objOPInsumosAux.sUMProduto = objOPInsumos.sUMProduto
                objOPInsumosAux.iComposicao = objOPInsumos.iComposicao
                objOPInsumosAux.dPercentualPerda = objOPInsumos.dPercentualPerda
                objOPInsumosAux.dCustoStandard = objOPInsumos.dCustoStandard
                objOPInsumosAux.sVersaoKitComp = objOPInsumos.sVersaoKitComp
            
                objOPOperacoesAux.colOPInsumos.Add objOPInsumosAux
            
                'Inicializa um novo Item de OP
                Set objItemOPNovo = New ClassItemOP
            
                objItemOPNovo.sProduto = objOPInsumos.sProduto
                objItemOPNovo.sVersao = objOPInsumos.sVersaoKitComp
            
                'e Pesquisa novamente suas Operações - chamada recursiva desta mesma função
                lErro = Gera_Operacoes_ItemOP1(objPO, objItemOPOrigem, objItemOPNovo, colOPOperacoes, dFatorQtde, iNivel + 1, iNivelRoteiro + 1, iSeqRoteiro, iSeq)
                If lErro <> SUCESSO Then gError 137932
            
            Next
            
            'encontrando o produto sai...
            Exit For
            
        End If
        
    Next
    
    Gera_Operacoes_ItemOP1 = SUCESSO
 
    Exit Function
 
Erro_Gera_Operacoes_ItemOP1:
 
    Gera_Operacoes_ItemOP1 = gErr
 
    Select Case gErr
    
        Case 137932
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152371)
 
    End Select
 
    Exit Function

End Function

Public Function RelRoteiroFabricacao_Prepara(ByVal objRoteiroPai As ClassRoteirosDeFabricacao, lNumIntRel As Long) As Long
'Monta a tabela temporária para o relatório de Roteiro de Fabricação
 
Dim lErro As Long
Dim lTransacao As Long
 
On Error GoTo Erro_RelRoteiroFabricacao_Prepara
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 138265
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_ROTEIRO", lNumIntRel)
    If lErro <> SUCESSO Then gError 138267
    
    'Montra em objRoteiroPai a coleção de SubRoteiros
    lErro = RelRoteiroFabricacao_Prepara_EmTrans(objRoteiroPai, lNumIntRel)
    If lErro <> SUCESSO Then gError 138266
     
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 138269
 
    RelRoteiroFabricacao_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelRoteiroFabricacao_Prepara:
 
    RelRoteiroFabricacao_Prepara = gErr
 
    Select Case gErr
    
        Case 138265
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 138266 To 138267
        
        Case 138269
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152372)
 
    End Select
 
    'Desfaz Transação
    Call Transacao_Rollback
 
    Exit Function
 
End Function


Public Function RelRoteiroFabricacao_Prepara_EmTrans(ByVal objRoteiroPai As ClassRoteirosDeFabricacao, ByVal lNumIntRel As Long) As Long
'Monta a tabela temporária para o relatório de Roteiro de Fabricação
 
Dim lErro As Long
Dim objRelRoteiro As New ClassRelRoteiro
Dim objRelRoteiroP As New ClassRelRoteiro
Dim objProduto As New ClassProduto
 
On Error GoTo Erro_RelRoteiroFabricacao_Prepara_EmTrans
        
    objProduto.sCodigo = objRoteiroPai.sProdutoRaiz
    
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138270
        
    'Le o Roteiro de Fabricação
    lErro = CF("RoteirosDeFabricacao_Le", objRoteiroPai)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 138271

    Set objRelRoteiro.objRoteiro = objRoteiroPai
    Set objRelRoteiroP.objRoteiro = objRoteiroPai
    
    Set objRelRoteiro.objProduto = objProduto
    Set objRelRoteiroP.objProduto = objProduto
        
    objRelRoteiro.iNivel = 0
    objRelRoteiro.dFatorMultiplicacao = 1

    objRelRoteiroP.iNivel = 0
    objRelRoteiroP.dFatorMultiplicacao = 1

    'Adiciona o Roteiro principal entre os subRoteiros
    objRelRoteiroP.colRoteirosFilhos.Add objRelRoteiro
    
    'Le os SubRoteiros em recursividade
    lErro = RelRoteiroFabricacao_Prepara_EmTrans2(objRelRoteiroP, objRelRoteiroP.colRoteirosFilhos)
    If lErro <> SUCESSO Then gError 138272
     
    'Faz a Gravação da Estrutura de Roteiros
    lErro = RelRoteiroFabricacao_Prepara_EmTrans3(objRelRoteiroP, lNumIntRel, 1)
    If lErro <> SUCESSO Then gError 138268
     
    RelRoteiroFabricacao_Prepara_EmTrans = SUCESSO
 
    Exit Function
 
Erro_RelRoteiroFabricacao_Prepara_EmTrans:
 
    RelRoteiroFabricacao_Prepara_EmTrans = gErr
 
    Select Case gErr
    
        Case 138270 To 138272, 138268
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152373)
 
    End Select
 
    Exit Function
 
End Function

Public Function RelRoteiroFabricacao_Prepara_EmTrans2(ByVal objRelRoteiro As ClassRelRoteiro, ByVal colRoteiros As Collection) As Long
'Monta a tabela temporária para o relatório de Roteiro de Fabricação
 
Dim lErro As Long
Dim objOperacao As ClassOperacoes
Dim objInsumo As ClassOperacaoInsumos
Dim objRoteiroFilho As ClassRoteirosDeFabricacao
Dim objProduto As ClassProduto
Dim objKit As ClassKit
Dim objRelRoteiroFilho As ClassRelRoteiro
Dim dFatorConversao As Double
 
On Error GoTo Erro_RelRoteiroFabricacao_Prepara_EmTrans2
    
    'Para cada Operação
    For Each objOperacao In objRelRoteiro.objRoteiro.colOperacoes
    
        'Para cada insumo utilizado na operação
        For Each objInsumo In objOperacao.colOperacaoInsumos
        
            Set objProduto = New ClassProduto
            Set objKit = New ClassKit
            Set objRelRoteiroFilho = New ClassRelRoteiro
            Set objRoteiroFilho = New ClassRoteirosDeFabricacao
            
            objProduto.sCodigo = objInsumo.sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 138273
        
            objKit.sProdutoRaiz = objInsumo.sProduto
            objKit.sVersao = objInsumo.sVersaoKitComp
        
            'Se o produto for produzível
            If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
            
                'Se a versão não estiver preenchida
                If objKit.sVersao = "" Then
                    
                    'Le a versão padrão do Kit
                    lErro = CF("Kit_Le_Padrao", objKit)
                    If lErro <> SUCESSO And lErro <> 106304 Then gError 138274
            
                    'Se não tiver versão padrão => Erro
                    If lErro <> SUCESSO Then gError 138275
            
                Else
                
                    'Le o Kit
                    lErro = CF("Kit_Le", objKit)
                    If lErro <> SUCESSO And lErro <> 21826 Then gError 138276
                
                    'Se não existir => Erro
                    If lErro <> SUCESSO Then gError 138277
                
                End If
                
                objRoteiroFilho.sProdutoRaiz = objKit.sProdutoRaiz
                objRoteiroFilho.sVersao = objKit.sVersao
            
                'Le o Roteiro de Fabricação
                lErro = CF("RoteirosDeFabricacao_Le", objRoteiroFilho)
                If lErro <> SUCESSO And lErro <> 134617 Then gError 138278
                
                'Se não encontrar => Erro
                If lErro <> SUCESSO Then gError 138279
                
                'Descobre o fator de conversao da UM do Insumo para o roteiro
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objInsumo.sUMProduto, objRoteiroFilho.sUM, dFatorConversao)
                If lErro <> SUCESSO Then gError 138280
                
                objRelRoteiroFilho.iNivel = objRelRoteiro.iNivel + 1
                objRelRoteiroFilho.dFatorMultiplicacao = objRelRoteiro.dFatorMultiplicacao * (((objInsumo.dQuantidade / (1 - objInsumo.dPercentualPerda)) * dFatorConversao) / objRoteiroFilho.dQuantidade)
                
                Set objRelRoteiroFilho.objProduto = objProduto
                Set objRelRoteiroFilho.objRoteiro = objRoteiroFilho
                
                colRoteiros.Add objRelRoteiroFilho
                
                'Le os descendetes em recursividade
                lErro = RelRoteiroFabricacao_Prepara_EmTrans2(objRelRoteiroFilho, colRoteiros)
                If lErro <> SUCESSO Then gError 138281
            
            End If
        
        Next
        
    Next
 
    RelRoteiroFabricacao_Prepara_EmTrans2 = SUCESSO
 
    Exit Function
 
Erro_RelRoteiroFabricacao_Prepara_EmTrans2:
 
    RelRoteiroFabricacao_Prepara_EmTrans2 = gErr
 
    Select Case gErr
 
        Case 138273, 138274, 138276, 138278, 138280, 138281
        
        Case 138275
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
            
        Case 138276
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_INEXISTENTE", gErr, objKit.sVersao, objKit.sProdutoRaiz)
 
        Case 138279
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSDEFABRICACAO_NAO_CADASTRADO", gErr, objRoteiroFilho.sProdutoRaiz, objRoteiroFilho.sVersao)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152374)
 
    End Select
 
    Exit Function
 
End Function


Public Function RelRoteiroFabricacao_Prepara_EmTrans3(ByVal objRelRoteiro As ClassRelRoteiro, ByVal lNumIntRel As Long, ByVal iSeqRot As Integer) As Long
'Grava a tabela temporária para o relatório de Roteiro de Fabricação
 
Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim objCompetencia As ClassCompetencias
Dim objCT As ClassCentrodeTrabalho
Dim objRelRoteiroFilho As ClassRelRoteiro
Dim objOperacao As ClassOperacoes
Dim iSeqOper As Integer
Dim iSeqIns As Integer
Dim objInsumo As ClassOperacaoInsumos
Dim objProduto As ClassProduto
Dim sIdentacao As String
Dim sCompra As String
 
On Error GoTo Erro_RelRoteiroFabricacao_Prepara_EmTrans3
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138282
    Next
    
    'Insere o Cabeçalho (Roteiro Origem)
    lErro = Comando_Executar(alComando(1), "INSERT INTO RelRoteiro (NumIntRel, Seq, Produto, Versao, DescProd, DescRot, Data, Quantidade, UM, NumIntDocRot ) VALUES (?,?,?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRot, objRelRoteiro.objProduto.sCodigo, objRelRoteiro.objRoteiro.sVersao, objRelRoteiro.objProduto.sDescricao, objRelRoteiro.objRoteiro.sDescricao, objRelRoteiro.objRoteiro.dtDataCriacao, objRelRoteiro.objRoteiro.dQuantidade, objRelRoteiro.objRoteiro.sUM, objRelRoteiro.objRoteiro.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 138283
    
    'Inicia o sequencial de Operações
    iSeqOper = 0

    'Para cada item na coleção de Roteiros
    For Each objRelRoteiroFilho In objRelRoteiro.colRoteirosFilhos
    
        'Acerta a quantidade do roteiro filho confomr possível diferença entre Insumos/Roteiros
        objRelRoteiroFilho.objRoteiro.dQuantidade = objRelRoteiroFilho.objRoteiro.dQuantidade * objRelRoteiroFilho.dFatorMultiplicacao
    
        'Para cada Operação do Roteiro
        For Each objOperacao In objRelRoteiroFilho.objRoteiro.colOperacoes
               
            'Incrementa o sequencial de Operações e zera o Sequencial de Insumos
            iSeqOper = iSeqOper + 1
            iSeqIns = 0
        
            Set objCompetencia = New ClassCompetencias
            Set objCT = New ClassCentrodeTrabalho
            
            objCompetencia.lNumIntDoc = objOperacao.lNumIntDocCompet
            
            'Lê a Competencia pelo NumIntDoc
            lErro = CF("Competencias_Le_NumIntDoc", objCompetencia)
            If lErro <> SUCESSO And lErro <> 134336 Then gError 138284
    
            objCT.lNumIntDoc = objOperacao.lNumIntDocCT
    
            'Le o centro de trabalho
            lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 138285
    
            'Encontra a identação com abse no vível da Operação + o Nivel do Roteiro
            sIdentacao = String(((objOperacao.iNivel + objRelRoteiroFilho.iNivel) * NUMERO_DE_CARACTERES_POR_IDENTACAO), CARACTER_DA_STRING_DE_IDENTACAO)
        
            'Insere as Operações
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelRoteiroOperacoes (NumIntRel, SeqRot, Seq, Produto, Versao, Descricao,Quantidade, UM, CT,NomeCT, Competencia, NomeCompetencia, Identacao, NumIntDocOper ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRot, iSeqOper, objRelRoteiroFilho.objProduto.sCodigo, objRelRoteiroFilho.objRoteiro.sVersao, objRelRoteiroFilho.objProduto.sDescricao, objRelRoteiroFilho.objRoteiro.dQuantidade, objRelRoteiroFilho.objRoteiro.sUM, objCT.lCodigo, objCT.sNomeReduzido, objCompetencia.lCodigo, objCompetencia.sNomeReduzido, sIdentacao, objOperacao.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 138286
    
            'Para cada Insumo dessa Operação
            For Each objInsumo In objOperacao.colOperacaoInsumos
            
                'Acerta a quantidade do Insumo conforme acerto anterior no roteiro
                objInsumo.dQuantidade = objInsumo.dQuantidade * objRelRoteiroFilho.dFatorMultiplicacao
            
                'Incrementa sequencial de Insumos
                iSeqIns = iSeqIns + 1
            
                Set objProduto = New ClassProduto
            
                objProduto.sCodigo = objInsumo.sProduto
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 138287
            
                If objProduto.iCompras = PRODUTO_COMPRAVEL Then
                    sCompra = "COMPRADO"
                Else
                    sCompra = "PRODUZIDO"
                End If
           
                'Insere os Insumos
                lErro = Comando_Executar(alComando(3), "INSERT INTO RelRoteiroInsumos (NumIntRel, SeqRot, SeqOper, Seq, Produto, Versao, Descricao, Quantidade, UM, Compra ) VALUES (?,?,?,?,?,?,?,?,?,?)", lNumIntRel, iSeqRot, iSeqOper, iSeqIns, objInsumo.sProduto, objInsumo.sVersaoKitComp, objProduto.sDescricao, objInsumo.dQuantidade, objInsumo.sUMProduto, sCompra)
                If lErro <> AD_SQL_SUCESSO Then gError 138288
    
            Next
    
        Next
    
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelRoteiroFabricacao_Prepara_EmTrans3 = SUCESSO
 
    Exit Function
 
Erro_RelRoteiroFabricacao_Prepara_EmTrans3:
 
    RelRoteiroFabricacao_Prepara_EmTrans3 = gErr
 
    Select Case gErr
    
        Case 138282
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 138283, 138286, 138288
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELROTEIRO", gErr)
        
        Case 138284, 138285, 138287
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152375)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
 
End Function

Function CusteioRoteiro_Automatico(lCodigo As Long) As Long
'Obtem o próximo número de CentrosDeTrabalho

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_CusteioRoteiro_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "PCPConfig", "NUM_PROX_CUSTEIOROTEIRO", "CusteioRoteiroFabricacao", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 137938

    lCodigo = lNumAuto

    CusteioRoteiro_Automatico = SUCESSO

    Exit Function

Erro_CusteioRoteiro_Automatico:

    CusteioRoteiro_Automatico = gErr

    Select Case gErr
    
         'Erro já tratado
         Case 137938
    
         Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152376)
    
     End Select
    
     Exit Function

End Function

Function CentrodeTrabalho_GravaOperadores_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTOperadores As New ClassCTOperadores
Dim objTurno As New ClassTurno

On Error GoTo Erro_CentrodeTrabalho_GravaOperadores_EmTrans

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 139089
        
    'Para cada CTMaquina na coleção
    For Each objCTOperadores In objCentrodeTrabalho.colOperadores
        
        objCTOperadores.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc

        lErro = Comando_Executar(lComando, "INSERT INTO CTOperadores (TipoMaoDeObra, NumIntDocCT, Quantidade) VALUES (?, ?, ?)", objCTOperadores.iCodTipoMO, objCTOperadores.lNumIntDocCT, objCTOperadores.iQuantidade)
        If lErro <> AD_SQL_SUCESSO Then gError 139090
        
    Next
        
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_GravaOperadores_EmTrans = SUCESSO
    
    Exit Function

Erro_CentrodeTrabalho_GravaOperadores_EmTrans:

    CentrodeTrabalho_GravaOperadores_EmTrans = gErr

    Select Case gErr

        Case 139089
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 139090
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTOPERADORES", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152387)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CentrodeTrabalho_ExcluiOperadores_EmTrans(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iCodTipoMO As Integer

On Error GoTo Erro_CentrodeTrabalho_ExcluiOperadores_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 139091
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT TipoMaoDeObra FROM CTOperadores WHERE NumIntDocCT = ?", 0, iCodTipoMO, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 139092

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 139093

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 139094
        
        'Apaga o registro
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM CTOperadores", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 139095
                
        'vai para o próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 139096

    Loop

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CentrodeTrabalho_ExcluiOperadores_EmTrans = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_ExcluiOperadores_EmTrans:

    CentrodeTrabalho_ExcluiOperadores_EmTrans = gErr

    Select Case gErr

        Case 139091
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 139092, 139093, 139096
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTOPERADORES", gErr)

        Case 139094
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTOPERADORES", gErr)

        Case 139095
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTOPERADORES", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152388)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Verifica_Operador_Maquina_Dia(ByVal objPMPItens As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal objCTMaquinas As ClassCTMaquinas, ByVal objMaquina As ClassMaquinas, ByVal objMaquinaOperador As ClassMaquinaOperadores, ByVal colMaquinas As Collection, ByVal dtData As Date, dQtdHorasDiaUso As Double) As Long
'Verifica o que já está sendo usada daquela máquina naquele dia para esse item (Capacidade Infinita)
 
Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim objMaquinaAux As ClassMaquinas
Dim objMaquinaOperadoresAux As ClassMaquinaOperadores
Dim objPOAux As ClassPlanoOperacional
Dim bAchou As Boolean
 
On Error GoTo Erro_Verifica_Operador_Maquina_Dia
 
    dQtdHorasDiaUso = 0
     
    For Each objPOAux In objPMPItens.colPO
    
        'Para cada máquina usada na etapa
        For Each objPOMaquinas In objPOAux.colAlocacaoMaquinas
    
            'Se tiver sido usada no mesmo dia e for do mesmo CT
            If objPOAux.lNumIntDocCT = objCTMaquinas.lNumIntDocCT And objPOMaquinas.dtData = dtData Then
                                
                bAchou = False
                                
                'Verifica se já existe a máquina na coleção para evitar leituras repetidas
                For Each objMaquinaAux In colMaquinas
                    If objPOMaquinas.lNumIntDocMaq = objMaquinaAux.lNumIntDoc Then
                        bAchou = True
                        Exit For
                    End If
                Next
            
                'Se não encontro lê
                If Not bAchou Then
                
                    Set objMaquinaAux = New ClassMaquinas
                
                    objMaquinaAux.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
                    
                    'Le a Máquina
                    lErro = CF("Maquinas_Le_NumIntDoc", objMaquinaAux)
                    If lErro <> SUCESSO And lErro <> 106353 Then gError 138512
                    
                    lErro = CF("Maquinas_Le_Itens", objMaquinaAux)
                    If lErro <> SUCESSO Then gError 138517
                    
                    colMaquinas.Add objMaquinaAux
                
                End If
                
                'Para cada Operador da máquina
                For Each objMaquinaOperadoresAux In objMaquinaAux.colTipoOperadores
                
                    'Se for o mesmo tipo de mão de Obra
                    If objMaquinaOperador.iTipoMaoDeObra = objMaquinaOperadoresAux.iTipoMaoDeObra Then
                
                        dQtdHorasDiaUso = dQtdHorasDiaUso + (objPOMaquinas.dHorasMaquina * objMaquinaOperador.iQuantidade * objMaquinaOperador.dPercentualUso)
                
                    End If
                
                Next
            
            End If

        Next
        
    Next
           
    Verifica_Operador_Maquina_Dia = SUCESSO
 
    Exit Function
 
Erro_Verifica_Operador_Maquina_Dia:
 
    Verifica_Operador_Maquina_Dia = gErr
 
    Select Case gErr
    
        Case 138512, 138517
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152389)
 
    End Select
 
    Exit Function
 
End Function

'##########################################################################
'Inserido por Wagner 07/11/2005
'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function CTMaquinasParadas_Grava(ByVal objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTMaquinasParadas_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 140714

    lErro = CF("CTMaquinasParadas_Grava_EmTrans", objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140715

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 140716

    CTMaquinasParadas_Grava = SUCESSO

    Exit Function

Erro_CTMaquinasParadas_Grava:

    CTMaquinasParadas_Grava = gErr

    Select Case gErr

        Case 140714
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 140715

        Case 140716
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152390)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinasParadas_Grava_EmTrans(ByVal objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CTMaquinasParadas_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140718
    Next

    'Le a tabelaCTMaquinasParadas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM CTMaquinasParadas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, iAux, objCTMaquinasParadas.lCodigo, objCTMaquinasParadas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 140719

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140720

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 140721

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE CTMaquinasParadas SET Data= ? , NumIntDocCT= ? , NumIntDocMaq= ? , Tipo= ? , " & _
                "Horas= ? , QtdMaquinas= ?, oBSERVACAO = ? ", alComando(0), _
                objCTMaquinasParadas.dtData, objCTMaquinasParadas.lNumIntDocCT, objCTMaquinasParadas.lNumIntDocMaq, objCTMaquinasParadas.iTipo, _
                objCTMaquinasParadas.dHoras, objCTMaquinasParadas.iQtdMaquinas, objCTMaquinasParadas.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 140722

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_CTMAQUINASPARADAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 140717

        objCTMaquinasParadas.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO CTMaquinasParadas( NumIntDoc, Codigo, FilialEmpresa, Data, " & _
                "NumIntDocCT, NumIntDocMaq, Tipo, Horas, QtdMaquinas, Observacao) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objCTMaquinasParadas.lNumIntDoc, objCTMaquinasParadas.lCodigo, objCTMaquinasParadas.iFilialEmpresa, objCTMaquinasParadas.dtData, _
                objCTMaquinasParadas.lNumIntDocCT, objCTMaquinasParadas.lNumIntDocMaq, objCTMaquinasParadas.iTipo, objCTMaquinasParadas.dHoras, objCTMaquinasParadas.iQtdMaquinas, objCTMaquinasParadas.sObservacao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140723

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinasParadas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinasParadas_Grava_EmTrans:

    CTMaquinasParadas_Grava_EmTrans = gErr

    Select Case gErr

        Case 140717

        Case 140718
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140719, 140720
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINASPARADAS", gErr)

        Case 140721
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQUINASPARADAS", gErr)

        Case 140722
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CTMAQUINASPARADAS", gErr)

        Case 140723
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTMAQUINASPARADAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152391)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function CTMaquinasParadas_Exclui(ByVal objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTMaquinasParadas_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 140705

    lErro = CF("CTMaquinasParadas_Exclui_EmTrans", objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140706

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 140707

    CTMaquinasParadas_Exclui = SUCESSO

    Exit Function

Erro_CTMaquinasParadas_Exclui:

    CTMaquinasParadas_Exclui = gErr

    Select Case gErr

        Case 140705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 140706

        Case 140707
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152392)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function CTMaquinasParadas_Exclui_EmTrans(ByVal objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CTMaquinasParadas_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140708
    Next

    'Le a tabelaCTMaquinasParadas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM CTMaquinasParadas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                0, iAux, objCTMaquinasParadas.lCodigo, objCTMaquinasParadas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 140709

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140710

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 140711

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 140712

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CTMaquinasParadas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 140713

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTMaquinasParadas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_CTMaquinasParadas_Exclui_EmTrans:

    CTMaquinasParadas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 140708
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140709, 140710
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINASPARADAS", gErr)

        Case 140711
            Call Rotina_Erro(vbOKOnly, "ERRO_CTMAQUINASPARADAS_NAO_CADASTRADO", gErr, objCTMaquinasParadas.lCodigo, objCTMaquinasParadas.iFilialEmpresa)

        Case 140712
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CTMAQUINASPARADAS", gErr)

        Case 140713
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CTMAQUINASPARADAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152393)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ItemOPOperacoesMaquinas_Insere(ByVal objOPOperacao As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long
Dim objItemOPOperacoesMaquinas As ClassItemOPOperacoesMaquinas

On Error GoTo Erro_ItemOPOperacoesMaquinas_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140741
    Next

    For Each objItemOPOperacoesMaquinas In objOPOperacao.colUsoMaquinas

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_ITEMOPOPERACOESMAQUINAS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 140740

        objItemOPOperacoesMaquinas.lNumIntDoc = lNumIntDoc
        
        If objOPOperacao.lNumIntDocOperOrigem = 0 Then
            objItemOPOperacoesMaquinas.lNumIntDocOper = objOPOperacao.lNumIntDoc
        Else
            objItemOPOperacoesMaquinas.lNumIntDocOper = objOPOperacao.lNumIntDocOperOrigem
        End If

        lErro = Comando_Executar(alComando(1), "INSERT INTO ItemOPOperacoesMaquinas( NumIntDoc,  NumIntDocOper, " & _
                "NumIntDocMaq, Data, Horas, Quantidade) VALUES (?,?,?,?,?,?)", _
                objItemOPOperacoesMaquinas.lNumIntDoc, objItemOPOperacoesMaquinas.lNumIntDocOper, _
                objItemOPOperacoesMaquinas.lNumIntDocMaq, objItemOPOperacoesMaquinas.dtData, objItemOPOperacoesMaquinas.dHoras, objItemOPOperacoesMaquinas.iQuantidade)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140746

        lErro = ItemOPOperacoesMO_Insere(objItemOPOperacoesMaquinas)
        If lErro <> SUCESSO Then gError 140745
    
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPOperacoesMaquinas_Insere = SUCESSO

    Exit Function

Erro_ItemOPOperacoesMaquinas_Insere:

    ItemOPOperacoesMaquinas_Insere = gErr

    Select Case gErr

        Case 140740

        Case 140741
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140746
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMOPOPERACOESMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152394)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function ItemOPOperacoesMaquinas_Exclui(ByVal objOPOperacao As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocOrigem As Long
Dim objItemOPOperacoesMaquinas As ClassItemOPOperacoesMaquinas

On Error GoTo Erro_ItemOPOperacoesMaquinas_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140731
    Next
    
    If objOPOperacao.lNumIntDocOperOrigem = 0 Then
        lNumIntDocOrigem = objOPOperacao.lNumIntDoc
    Else
        lNumIntDocOrigem = objOPOperacao.lNumIntDocOperOrigem
    End If

    'Le a tabela ItemOPOperacoesMaquinas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItemOPOperacoesMaquinas WHERE NumIntDocOper= ? ", _
                0, lNumIntDoc, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 140732

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140733

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 140735
        
        Set objItemOPOperacoesMaquinas = New ClassItemOPOperacoesMaquinas
        
        objItemOPOperacoesMaquinas.lNumIntDoc = lNumIntDoc
        
        lErro = ItemOPOperacoesMO_Exclui(objItemOPOperacoesMaquinas)
        If lErro <> SUCESSO Then gError 140737
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItemOPOperacoesMaquinas", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 140736
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140734
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPOperacoesMaquinas_Exclui = SUCESSO

    Exit Function

Erro_ItemOPOperacoesMaquinas_Exclui:

    ItemOPOperacoesMaquinas_Exclui = gErr

    Select Case gErr

        Case 140731
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140732, 140733, 140734
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPOPERACOESMAQUINAS", gErr)

        Case 140735
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ITEMOPOPERACOESMAQUINAS", gErr)

        Case 140736
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOPOPERACOESMAQUINAS", gErr)
            
        Case 140737

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152395)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function ItemOPOperacoesMO_Insere(ByVal objItemOPOperacoesMaquinas As ClassItemOPOperacoesMaquinas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long
Dim objItemOPOperacoesMO As ClassItemOPOperacoesMO

On Error GoTo Erro_ItemOPOperacoesMO_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140764
    Next

    For Each objItemOPOperacoesMO In objItemOPOperacoesMaquinas.colMO

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_INT_PROX_ITEMOPOPERACOESMO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 140763

        objItemOPOperacoesMO.lNumIntDoc = lNumIntDoc
        objItemOPOperacoesMO.lNumIntDocItemOPMaq = objItemOPOperacoesMaquinas.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO ItemOPOperacoesMO( NumIntDoc, NumIntDocItemOPMaq, TipoMO, Horas, Quantidade) VALUES (?,?,?,?,?)", _
                objItemOPOperacoesMO.lNumIntDoc, objItemOPOperacoesMO.lNumIntDocItemOPMaq, objItemOPOperacoesMO.iTipoMO, objItemOPOperacoesMO.dHoras, objItemOPOperacoesMO.iQuantidade)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140769

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPOperacoesMO_Insere = SUCESSO

    Exit Function

Erro_ItemOPOperacoesMO_Insere:

    ItemOPOperacoesMO_Insere = gErr

    Select Case gErr

        Case 140763

        Case 140764
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140769
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMOPOPERACOESMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152396)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function ItemOPOperacoesMO_Exclui(ByVal objItemOPOperacoesMaquinas As ClassItemOPOperacoesMaquinas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_ItemOPOperacoesMO_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140754
    Next

    'Le a tabelaItemOPOperacoesMO
    lErro = Comando_ExecutarPos(alComando(0), "SELECT 1 FROM ItemOPOperacoesMO WHERE NumIntDocItemOPMaq= ? ", _
                0, iAux, objItemOPOperacoesMaquinas.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 140755

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140756

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 140758
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItemOPOperacoesMO", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 140759
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140757
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemOPOperacoesMO_Exclui = SUCESSO

    Exit Function

Erro_ItemOPOperacoesMO_Exclui:

    ItemOPOperacoesMO_Exclui = gErr

    Select Case gErr

        Case 140754
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140755, 140756, 140757
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPOPERACOESMO", gErr)

        Case 140758
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ITEMOPOPERACOESMO", gErr)

        Case 140759
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMOPOPERACOESMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152397)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CTMaquinasParadas_Automatico(lCodigo As Long) As Long
'Obtem o próximo número de CentrosDeTrabalho

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_CTMaquinasParadas_Automatico

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "PCPConfig", "NUM_PROX_CTMAQUINASPARADAS_COD", "CTMaquinasParadas", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then gError 140971

    lCodigo = lNumAuto

    CTMaquinasParadas_Automatico = SUCESSO

    Exit Function

Erro_CTMaquinasParadas_Automatico:

    CTMaquinasParadas_Automatico = gErr

   Select Case gErr

        'Erro já tratado
        Case 140971

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152398)

    End Select

    Exit Function

End Function
'##########################################################################

Public Function RelUsoMaquina_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal iCodMaqIni As Integer, ByVal iCodMaqFim As Integer, ByVal lCodCTIni As Long, ByVal lCodCTFim As Long, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long
'Monta a tabela temporária para o relatório de calendário de máquina
 
Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dtData As Date
Dim colCTMaquinas As New Collection
Dim objCTMaquina As ClassCTMaquinas
Dim objMaquina As ClassMaquinas
Dim objCT As ClassCentrodeTrabalho
 
On Error GoTo Erro_RelUsoMaquina_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 141109
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141110
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_USOMAQUINAS", lNumIntRel)
    If lErro <> SUCESSO Then gError 141111
    
    lErro = CF("CTMaquina_Le_Faixa", iFilialEmpresa, iCodMaqIni, iCodMaqFim, lCodCTIni, lCodCTFim, colCTMaquinas)
    If lErro <> SUCESSO Then gError 141112
    
    For Each objCTMaquina In colCTMaquinas
    
        Set objMaquina = New ClassMaquinas
        Set objCT = New ClassCentrodeTrabalho
   
        objMaquina.lNumIntDoc = objCTMaquina.lNumIntDocMaq
         
        'Le a máquina
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 141113
        
        objCT.lNumIntDoc = objCTMaquina.lNumIntDocCT
         
        'Le o Centro de Trabalho
        lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 141114
        
        dtData = dtDataIni
        
        Do While dtData <= dtDataFim
        
            lErro = RelUsoMaquina_Prepara2(alComando(1), alComando(0), lNumIntRel, objMaquina, objCT, dtData)
            If lErro <> SUCESSO Then gError 141115
            
            dtData = DateAdd("d", 1, dtData)
            
        Loop
    
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 141116
 
    RelUsoMaquina_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelUsoMaquina_Prepara:
 
    RelUsoMaquina_Prepara = gErr
 
    Select Case gErr
 
        Case 141109
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 141110
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 141111 To 141115
 
        Case 141116
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152399)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelUsoMaquina_Prepara2(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lNumIntRel As Integer, ByVal objMaquina As ClassMaquinas, ByVal objCT As ClassCentrodeTrabalho, ByVal dtData As Date) As Long
'Monta a tabela temporária para o relatório de calendário de máquina
 
Dim lErro As Long
Dim objCTMaquina As New ClassCTMaquinas
Dim dHorasUsoPrevisto As Double
Dim dHorasUsoReal As Double
Dim dHorasParada As Double
Dim iQtdDisponivel As Integer
Dim dHorasDisponivel As Double
 
On Error GoTo Erro_RelUsoMaquina_Prepara2
    
    objCTMaquina.lNumIntDocCT = objCT.lNumIntDoc
    objCTMaquina.lNumIntDocMaq = objMaquina.lNumIntDoc
    
    'Obtem a quantidade de máquinas e a quantidade de horas disponíveis
    lErro = Verifica_Maquinas_Disponiveis_Dia(objCTMaquina, dtData, objCT.iFilialEmpresa, iQtdDisponivel, dHorasDisponivel)
    If lErro <> SUCESSO Then gError 141124
    
    'Obtem a quantidade de máquinas prevista para uso no dia
    lErro = Verifica_Uso_Maquina_Dia_Previsto(lComando1, objCTMaquina, dtData, dHorasUsoPrevisto)
    If lErro <> SUCESSO Then gError 141125
    
    'Obtem a quantidade de máquinas usadas no dia
    lErro = Verifica_Uso_Maquina_Dia_Real(lComando1, objCTMaquina, dtData, dHorasUsoReal)
    If lErro <> SUCESSO Then gError 141126
    
    'Obtem a quantidade de máquinas paradas no dia
    lErro = Verifica_Maquina_Parada_Dia(lComando1, objCTMaquina, dtData, dHorasParada)
    If lErro <> SUCESSO Then gError 141127
    
    lErro = Comando_Executar(lComando2, "INSERT INTO RelOPUsoMaquina( NumIntRel, NumIntDocCT, NumIntDocMaq, Data, CodCT, NomeCT, CodMaq, NomeMaq, QtdeDisponivel, HorasUsoReal, HorasUsoPrevisto, HorasDisponivel, HorasParada, HorasOcisosa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                        lNumIntRel, objCT.lNumIntDoc, objMaquina.lNumIntDoc, dtData, objCT.lCodigo, objCT.sNomeReduzido, objMaquina.iCodigo, objMaquina.sNomeReduzido, iQtdDisponivel, dHorasUsoReal, dHorasUsoPrevisto, dHorasDisponivel, dHorasParada, dHorasDisponivel - dHorasParada - dHorasUsoReal)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141128
    
    RelUsoMaquina_Prepara2 = SUCESSO
 
    Exit Function
 
Erro_RelUsoMaquina_Prepara2:
 
    RelUsoMaquina_Prepara2 = gErr
 
    Select Case gErr
    
        Case 141124 To 141128

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152400)
 
    End Select
 
    Exit Function
 
End Function

Public Function Verifica_Uso_Maquina_Dia_Previsto(ByVal lComando As Long, ByVal objCTMaquina As ClassCTMaquinas, ByVal dtData As Date, dHorasUsoPrevisto As Double) As Long

Dim lErro As Long

On Error GoTo Erro_Verifica_Uso_Maquina_Dia_Previsto

    dHorasUsoPrevisto = 0

    'Le a tabelaPlanoOperacionalMaquinas
    lErro = Comando_Executar(lComando, "SELECT SUM(POM.HorasMaquinas) FROM PlanoOperacionalMaquinas AS POM, PlanoOperacional AS PO WHERE PO.NumIntDoc = POM.NumIntDocPO AND PO.NumIntDocCT = ? AND POM.NumIntDocMaq = ? AND POM.Data = ?", _
                                    dHorasUsoPrevisto, objCTMaquina.lNumIntDocCT, objCTMaquina.lNumIntDocMaq, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 141118

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141119

    Verifica_Uso_Maquina_Dia_Previsto = SUCESSO

    Exit Function

Erro_Verifica_Uso_Maquina_Dia_Previsto:

    Verifica_Uso_Maquina_Dia_Previsto = gErr

    Select Case gErr

        Case 141118, 141119
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOOPERACIONAMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152401)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Verifica_Uso_Maquina_Dia_Real(ByVal lComando As Long, ByVal objCTMaquina As ClassCTMaquinas, ByVal dtData As Date, dHorasUsoReal As Double) As Long

Dim lErro As Long

On Error GoTo Erro_Verifica_Uso_Maquina_Dia_Real

    dHorasUsoReal = 0

    'Le a tabelaPlanoOperacionalMaquinas
    lErro = Comando_Executar(lComando, "SELECT SUM(Horas) FROM ItemOPOperacoesMaquinas AS IOPOM, OrdemProducaoOperacoes AS IOPO WHERE IOPOM.NumIntDocOper = IOPO.NumIntDoc AND IOPO.NumIntDocCT = ? AND IOPOM.NumIntDocMaq = ? AND IOPOM.Data = ?", _
                                    dHorasUsoReal, objCTMaquina.lNumIntDocCT, objCTMaquina.lNumIntDocMaq, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 141120

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141121

    Verifica_Uso_Maquina_Dia_Real = SUCESSO

    Exit Function

Erro_Verifica_Uso_Maquina_Dia_Real:

    Verifica_Uso_Maquina_Dia_Real = gErr

    Select Case gErr

        Case 141120, 141121
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPOPERACOESMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152402)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Verifica_Maquina_Parada_Dia(ByVal lComando As Long, ByVal objCTMaquina As ClassCTMaquinas, ByVal dtData As Date, dHorasParada As Double) As Long

Dim lErro As Long

On Error GoTo Erro_Verifica_Maquina_Parada_Dia

    dHorasParada = 0

    'Le a tabelaPlanoOperacionalMaquinas
    lErro = Comando_Executar(lComando, "SELECT SUM(CTM.Horas) FROM CTMaquinasParadas AS CTM WHERE CTM.NumIntDocCT = ? AND CTM.NumIntDocMaq = ? AND CTM.Data = ? ", _
                                    dHorasParada, objCTMaquina.lNumIntDocCT, objCTMaquina.lNumIntDocMaq, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 141122

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141123

    Verifica_Maquina_Parada_Dia = SUCESSO

    Exit Function

Erro_Verifica_Maquina_Parada_Dia:

    Verifica_Maquina_Parada_Dia = gErr

    Select Case gErr

        Case 141122, 141123
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPOPERACOESMAQUINAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152403)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'########################################################################

'########################################################################
'Inserido por Wagner 06/02/2006
Public Function RelNecessidadeRoteiro_Prepara(lNumIntRel As Long, ByVal objRoteiroNecess As ClassRoteiroNecessidade, Optional bRecursivo As Boolean = False) As Long
'Monta a tabela temporária para o relatório de críticas do Plano Mestre de Produção
 
Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim objRoteiro As New ClassRoteirosDeFabricacao
Dim objRoteiroInsumosMaq As ClassRoteiroInsumosMaq
Dim objRoteiroMaoDeObra As ClassRoteiroMaoDeObra
Dim objRoteiroMP As ClassRoteiroMP
Dim objProduto As ClassProduto
Dim objMaodeObra As ClassTiposDeMaodeObra
Dim iSeq As Integer
Dim sSQL As String
Dim iFilialEmpresa As Integer, dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double, dCusto As Double
Dim objRoteiroNecessAux As ClassRoteiroNecessidade
Dim objRoteiroMPAux As ClassRoteiroMP
Dim dCustoAux As Double
Dim bTrue As Boolean

On Error GoTo Erro_RelNecessidadeRoteiro_Prepara
 
    GL_objMDIForm.MousePointer = vbHourglass
    
    bTrue = True
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 141667
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141668
    Next
    
    iFilialEmpresa = objRoteiroNecess.iFilialEmpresa
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    'obter parametros para calculo de precos
    lErro = CF("CalcMP_ObterValores", iFilialEmpresa, dTaxaDescPadrao, dTaxaValFut, dDiasValFut)
    If lErro <> SUCESSO Then gError 130488
    
    objRoteiro.sProdutoRaiz = objRoteiroNecess.sProdutoRaiz
    objRoteiro.sVersao = objRoteiroNecess.sVersao
 
    lErro = CF("RoteirosDeFabricacao_Le", objRoteiro)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 141669
    
    Set objProduto = New ClassProduto
 
    objProduto.sCodigo = objRoteiroNecess.sProdutoRaiz
 
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141670
 
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_NECESSROT", lNumIntRel)
    If lErro <> SUCESSO Then gError 141671
           
    lErro = Comando_Executar(alComando(0), "INSERT INTO RelNecessidadeRoteiro (NumIntRel, Produto, DescProd, Versao, DescRoteiro, UM, Quantidade) VALUES (?,?,?,?,?,?,?) ", _
    lNumIntRel, objProduto.sCodigo, objProduto.sDescricao, objRoteiroNecess.sVersao, objRoteiro.sDescricao, objRoteiroNecess.sUM, objRoteiroNecess.dQuantidade)
    If lErro <> SUCESSO Then gError 141672

    lErro = CF("NecessidadeRoteiro_Le", objRoteiroNecess, bRecursivo)
    If lErro <> SUCESSO Then gError 141673
    
    sSQL = "INSERT INTO RelNecessidadeRoteiroItens (NumIntRel, Seq, Tipo, Produto, DescProd, Compravel, UM, Quantidade, MaoDeObra, DescMO, CustoUnitario) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
    
    For Each objRoteiroInsumosMaq In objRoteiroNecess.colInsumosMaquina
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objRoteiroInsumosMaq.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141674
        
        iSeq = iSeq + 1
    
        dCusto = objRoteiroInsumosMaq.dCustoUnitario
        If dCusto = 0 Then
        
            'obter o seu custo em custembmp.
            lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objProduto.sCodigo, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCusto, alComando(3))
            If lErro <> SUCESSO Then gError 130489
        
        End If
        
        lErro = Comando_Executar(alComando(0), sSQL, _
            lNumIntRel, iSeq, 1, objProduto.sCodigo, objProduto.sDescricao, objProduto.iCompras, objRoteiroInsumosMaq.sUM, objRoteiroInsumosMaq.dQuantidade, 0, "", dCusto)
        If lErro <> SUCESSO Then gError 141675
        
    Next
     
    For Each objRoteiroMaoDeObra In objRoteiroNecess.colMaoDeObra
    
        Set objMaodeObra = New ClassTiposDeMaodeObra
    
        objMaodeObra.iCodigo = objRoteiroMaoDeObra.iCodMO
    
        lErro = CF("TiposDeMaodeObra_Le", objMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 141676
        
        iSeq = iSeq + 1
        
        lErro = Comando_Executar(alComando(0), sSQL, _
        lNumIntRel, iSeq, 2, "", "", 0, objRoteiroMaoDeObra.sUM, objRoteiroMaoDeObra.dHoras, objRoteiroMaoDeObra.iCodMO, objMaodeObra.sDescricao, objRoteiroMaoDeObra.dCustoUnitario)
        If lErro <> SUCESSO Then gError 141677
        
    Next
    
    For Each objRoteiroMP In objRoteiroNecess.colMP
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objRoteiroMP.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141678
        
        iSeq = iSeq + 1
                
        dCusto = objRoteiroMP.dCustoUnitario
        If dCusto = 0 Then
            
            If objProduto.iCompras = PRODUTO_COMPRAVEL Then
            
                'obter o seu custo em custembmp.
                lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objProduto.sCodigo, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCusto, alComando(2))
                If lErro <> SUCESSO Then gError 130490
                
            Else
            
                '#####################################################
                'Alteração para pegar o custo de produtos intermediários
                Set objRoteiroNecessAux = New ClassRoteiroNecessidade
                
                objRoteiroNecessAux.iFilialEmpresa = iFilialEmpresa
                objRoteiroNecessAux.dQuantidade = 1 'objRoteiroMP.dQuantidade
                objRoteiroNecessAux.sProdutoRaiz = objRoteiroMP.sProduto
                objRoteiroNecessAux.sUM = objRoteiroMP.sUM
                objRoteiroNecessAux.sVersao = objRoteiroMP.sVersao

                lErro = CF("NecessidadeRoteiro_Le", objRoteiroNecessAux, bTrue)
                If lErro <> SUCESSO Then gError 180268
                
                For Each objRoteiroInsumosMaq In objRoteiroNecessAux.colInsumosMaquina

                    dCustoAux = 0
                    
                    'obter o seu custo em custembmp.
                    lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objRoteiroInsumosMaq.sProduto, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCustoAux, alComando(3))
                    If lErro <> SUCESSO Then gError 180269
                    
                    dCusto = dCusto + (dCustoAux * objRoteiroInsumosMaq.dQuantidade)

                Next

                For Each objRoteiroMaoDeObra In objRoteiroNecessAux.colMaoDeObra

                    dCusto = dCusto + objRoteiroMaoDeObra.dCustoTotal

                Next
                
                For Each objRoteiroMPAux In objRoteiroNecessAux.colMP

                    dCustoAux = 0
                    
                    'obter o seu custo em custembmp.
                    lErro = CF("CustoDiretoProd_ObtemValor", iFilialEmpresa, objRoteiroMPAux.sProduto, dTaxaDescPadrao, dTaxaValFut, dDiasValFut, dCustoAux, alComando(4))
                    If lErro <> SUCESSO Then gError 180270
                    
                    dCusto = dCusto + (dCustoAux * objRoteiroMPAux.dQuantidade)

                Next
                '#####################################################
            
            End If
            
        End If
        
        lErro = Comando_Executar(alComando(0), sSQL, _
        lNumIntRel, iSeq, 3, objProduto.sCodigo, objProduto.sDescricao, objProduto.iCompras, objRoteiroMP.sUM, objRoteiroMP.dQuantidade, 0, "", dCusto)
        If lErro <> SUCESSO Then gError 141679
        
    Next
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 141680
 
    GL_objMDIForm.MousePointer = vbDefault
 
    RelNecessidadeRoteiro_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelNecessidadeRoteiro_Prepara:
 
    GL_objMDIForm.MousePointer = vbDefault
 
    RelNecessidadeRoteiro_Prepara = gErr
 
    Select Case gErr
 
        Case 141667
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 141668
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
  
        Case 141669, 141670, 141671, 141673, 141674, 141676, 141678, 130488, 130489, 130490, 180268, 180269, 180270
                
        Case 141672, 141675, 141677, 141679
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELNECESSROT", gErr)
 
        Case 141680
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152404)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Call Transacao_Rollback
 
    Exit Function
 
End Function
'########################################################################


Public Function Verifica_Tempo_Decorrido_Dia(ByVal objPMPItens As ClassPMPItens, ByVal objPO As ClassPlanoOperacional, ByVal dtData As Date, dQtdHoras As Double) As Long
'Verifica o que já está sendo usada daquela máquina naquele dia para esse item (Capacidade Infinita)
 
Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim objPOPai As ClassPlanoOperacional
Dim dMaiorQtdHoras As Double
Dim dQtdHorasAux As Double
 
On Error GoTo Erro_Verifica_Tempo_Decorrido_Dia

    If objPO.lNumIntDocPOPai <> 0 Then
    
        For Each objPOPai In objPMPItens.colPO
        
            If objPOPai.lNumIntDoc = objPO.lNumIntDocPOPai Then
        
                dMaiorQtdHoras = 0
                For Each objPOMaquinas In objPOPai.colAlocacaoMaquinas
                
                    If objPOMaquinas.dtData = dtData Then
                    
                        dQtdHorasAux = objPOMaquinas.dHorasMaquina / objPOMaquinas.iQuantidade
                        
                        If dQtdHorasAux > dMaiorQtdHoras Then dMaiorQtdHoras = dQtdHorasAux
                    
                    End If
                
                Next
                
                dQtdHoras = dQtdHoras + dMaiorQtdHoras
                
                Call Verifica_Tempo_Decorrido_Dia(objPMPItens, objPOPai, dtData, dQtdHoras)
                
                Exit For
                
            End If
        
        Next
    
    End If
              
    Verifica_Tempo_Decorrido_Dia = SUCESSO
 
    Exit Function
 
Erro_Verifica_Tempo_Decorrido_Dia:
 
    Verifica_Tempo_Decorrido_Dia = gErr
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152340)
 
    End Select
 
    Exit Function
 
End Function

Function RelPCPCompConsumo_Prepara(ByVal objRelCompConsParam As ClassRelCompConsParam) As Long

Dim lErro As Long, lNumIntRel As Long
Dim lTransacao As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer, iSequencial As Integer
Dim sSel1 As String, lUltNumIntDocItemOP As Long, sSiglaUMEstoque As String
Dim lNumIntDocItemOP As Long, sCodigoOP As String, sProduto As String, sVersao As String, dtData As Date, sDescricao As String
Dim objRelCompCons As ClassRelCompCons, objRelCompConsItem As ClassRelCompConsItem, iClasseUM As Integer

On Error GoTo Erro_RelPCPCompConsumo_Prepara
 
    GL_objMDIForm.MousePointer = vbHourglass

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184145
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184146
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_COMPCONSUMO", lNumIntRel)
    If lErro <> SUCESSO Then gError 184147
    
    objRelCompConsParam.lNumIntRel = lNumIntRel
    
    'procurar entradas de producao ligadas a itens de op baixadas no periodo informado
    sSel1 = "SELECT ItensOrdemProducaoBaixadas.NumIntDoc, ItensOrdemProducaoBaixadas.Codigo, ItensOrdemProducaoBaixadas.Produto, ItensOrdemProducaoBaixadas.Versao, MovimentoEstoque.Data, Produtos.Descricao, Produtos.ClasseUM, Produtos.SiglaUMEstoque FROM MovimentoEstoque, ItensOrdemProducaoBaixadas, OrdensDeProducaoBaixadas, Produtos WHERE Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND TipoMov IN (?,?) AND Data BETWEEN ? AND ? AND MovimentoEstoque.NumIntDocOrigem = ItensOrdemProducaoBaixadas.NumIntDoc AND ItensOrdemProducaoBaixadas.FilialEmpresa = OrdensDeProducaoBaixadas.FilialEmpresa AND ItensOrdemProducaoBaixadas.Codigo = OrdensDeProducaoBaixadas.Codigo AND MovimentoEstoque.NumIntDocEst = 0 ORDER BY CodigoOP ASC, MovimentoEstoque.Produto ASC, MovimentoEstoque.Data DESC"
    
    'ler
    sCodigoOP = String(STRING_OPCODIGO, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sVersao = String(STRING_KIT_VERSAO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(alComando(0), sSel1, lNumIntDocItemOP, sCodigoOP, sProduto, sVersao, dtData, sDescricao, iClasseUM, sSiglaUMEstoque, _
        objRelCompConsParam.iFilialEmpresa, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3, objRelCompConsParam.dtDataFimProdDe, objRelCompConsParam.dtDataFimProdAte)
    If lErro <> AD_SQL_SUCESSO Then gError 184148
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184149
    
    'enquanto houverem producoes
    Do While lErro <> AD_SQL_SEM_DADOS
       
        'se op e produto dentro da faixa
        If ((sCodigoOP <= objRelCompConsParam.sOPCodigoAte Or objRelCompConsParam.sOPCodigoAte = "") And _
            (sCodigoOP >= objRelCompConsParam.sOPCodigoDe Or objRelCompConsParam.sOPCodigoDe = "") And _
            (sProduto <= objRelCompConsParam.sProdFim Or objRelCompConsParam.sProdFim = "") And _
            (sProduto >= objRelCompConsParam.sProdIni Or objRelCompConsParam.sProdIni = "")) Then
            
            'se trocou item op entao processá-lo
            If lUltNumIntDocItemOP <> lNumIntDocItemOP Then
            
                Set objRelCompCons = New ClassRelCompCons
                Set objRelCompCons.objParam = objRelCompConsParam
                
                With objRelCompCons
                    .lNumIntDocItemOP = lNumIntDocItemOP
                    .sCodigoOP = sCodigoOP
                    .sProdutoPai = sProduto
                    .sProdutoPaiDescricao = sDescricao
                    .dtDataFinalProducao = dtData
                    .sKitVersao = sVersao
                    .sUMProdutoPai = sSiglaUMEstoque
                    .iClasseUMPai = iClasseUM
                End With
                
                'obtem consumos reais
                lErro = RelPCPCompConsumo_Prepara2(objRelCompCons, alComando(2))
                If lErro <> SUCESSO Then gError 184150
                
                'obtem consumos previstos (fora embalagens)
                lErro = RelPCPCompConsumo_Prepara1(objRelCompCons, alComando(3))
                If lErro <> SUCESSO Then gError 184151
                
                'obter consumo previsto de embalagens
                If objRelCompCons.objParam.iIgnorarEmbalagem = 0 Then
                    lErro = RelPCPCompConsumo_Prepara3(objRelCompCons, alComando(4))
                    If lErro <> SUCESSO Then gError 184152
                End If
                
                'grava registro "pai" do item da op
                With objRelCompCons
                    lErro = Comando_Executar(alComando(1), "INSERT INTO RelPCPCompConsumo ( NumIntRel, FilialEmpresa, CodigoOP, ProdutoPai, ProdutoPaiDescricao, DataFinalProducao, UMProdutoPai, QtdeProduzida, KitVersao, KitPerda, CustoContabil ) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                        lNumIntRel, objRelCompConsParam.iFilialEmpresa, .sCodigoOP, .sProdutoPai, .sProdutoPaiDescricao, .dtDataFinalProducao, .sUMProdutoPai, .dQtdeProduzida, .sKitVersao, .dKitPerda, .dCustoContabil)
                End With
                
                If lErro <> AD_SQL_SUCESSO Then gError 184153
                
                iSequencial = 900
                
                'grava registros "filhos" do item da op
                For Each objRelCompConsItem In objRelCompCons.colItens
                
                    If objRelCompConsItem.iSequencial = 0 Then objRelCompConsItem.iSequencial = 999
                    
                    With objRelCompConsItem
                        lErro = Comando_Executar(alComando(2), "INSERT INTO RelPCPCompConsumoItem (NumIntRel, FilialEmpresa, CodigoOP, ProdutoPai, ProdutoFilho, ProdutoFilhoDescricao, UMProdutoFilho, QtdeConsumida, KitVersao, KitPerda, KitQtdePrevista, CustoContabil, Sequencial) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                            lNumIntRel, objRelCompConsParam.iFilialEmpresa, objRelCompCons.sCodigoOP, objRelCompCons.sProdutoPai, .sProdutoFilho, .sProdutoFilhoDescricao, .sUMProdutoFilho, .dQtdeConsumida, .sKitVersao, .dKitPerda, .dKitQtdePrevista, .dCustoContabil, .iSequencial)
                    End With
                    If lErro <> SUCESSO Then gError 184157
                    
                Next
                
                lUltNumIntDocItemOP = lNumIntDocItemOP
            
            End If
            
        End If
                
        'ler a proxima
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184154
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184155
 
    GL_objMDIForm.MousePointer = vbDefault
 
    RelPCPCompConsumo_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelPCPCompConsumo_Prepara:
 
    GL_objMDIForm.MousePointer = vbDefault
 
    RelPCPCompConsumo_Prepara = gErr
 
    Select Case gErr
 
        Case 184145
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 184146
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
  
        Case 184147, 184150, 184151, 184152
                
        Case 184153, 184157
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELPCPCOMPCONSUMO", gErr)
 
        Case 184148, 184149, 184154
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELPCPCOMPCONS", gErr)
        
        Case 184155
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184156)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    Call Transacao_Rollback
 
    Exit Function
 
End Function

Function RelPCPCompConsumo_Prepara1(ByVal objRelCompCons As ClassRelCompCons, lComando As Long) As Long
'obtem consumo previsto pelo kit (menos de embalagem)

Dim lErro As Long, objRelCompConsItem As ClassRelCompConsItem, dQuantidadeBase As Double, dPercentualPerdaBase As Double, iNivel As Integer, sSiglaUMEstoquePai As String, iSequencial As Integer
Dim sProduto As String, dQuantidade As Double, sUnidadeMed As String, iComposicao As Integer, dPercentualPerda As Double, sDescricao As String, sSiglaUMEstoque As String, iClasseUM As Integer, iNatureza As Integer
Dim dQuantidadeProprio As Double, bAchou As Boolean, objRelCompConsItemNovo As ClassRelCompConsItem, dFatorConv As Double, iClasseUMPai As Integer, sUnidadeMedPai As String, dFator As Double

On Error GoTo Erro_RelPCPCompConsumo_Prepara1
    
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
        
    'pesquisar em produtokit no nivel 0 e 1, por ordem de nivel
    lErro = Comando_Executar(lComando, "SELECT ProdutoKit.Seq, ProdutoKit.Nivel, Produto, Quantidade, UnidadeMed, Composicao, PercentualPerda, Produtos.Descricao, Produtos.SiglaUMEstoque, Produtos.ClasseUM, Natureza FROM ProdutoKit, Produtos WHERE ProdutoKit.Produto = Produtos.Codigo AND ProdutoRaiz = ? AND Versao = ? AND ProdutoKit.Nivel < 2 AND Produtos.ControleEstoque <> ? ORDER BY ProdutoKit.Nivel, ProdutoKit.Seq", _
        iSequencial, iNivel, sProduto, dQuantidade, sUnidadeMed, iComposicao, dPercentualPerda, sDescricao, sSiglaUMEstoque, iClasseUM, iNatureza, objRelCompCons.sProdutoPai, objRelCompCons.sKitVersao, PRODUTO_CONTROLE_SEM_ESTOQUE)
    If lErro <> SUCESSO Then gError 184134
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184135
        
    If lErro <> AD_SQL_SEM_DADOS Then
            
        'ler
        Do While lErro <> AD_SQL_SEM_DADOS
            
            'Faz a Conversao de unidades de medida
            lErro = CF("UM_Conversao_Trans", iClasseUM, sUnidadeMed, sSiglaUMEstoque, dFatorConv)
            If lErro <> SUCESSO Then gError 184136
        
            dQuantidade = dFatorConv * dQuantidade
                
            'se o nivel for 0 entao
            If iNivel = 0 Then
                
                dQuantidadeBase = dQuantidade
                
                objRelCompCons.dKitPerda = dPercentualPerda
                
                lErro = CF("CustoProd_AjustarPerda", sProduto, "", objRelCompCons.objParam.iFilialEmpresa, FORMACAO_PRECO_REL_COMP_CONSUMO, dPercentualPerda)
                If lErro <> SUCESSO Then gError 184137
                
                dPercentualPerdaBase = dPercentualPerda
                iClasseUMPai = iClasseUM
                sSiglaUMEstoquePai = sSiglaUMEstoque
                sUnidadeMedPai = sUnidadeMed
                
            Else
                
                'ajusta qtde pela perda
                If dPercentualPerda <> 0 And dPercentualPerda <> 1 And gobjEST.iUsaPerdaInsumosKit Then
                    dQuantidade = dQuantidade / (1 - dPercentualPerda)
                End If
                
                'pesquisa produto na colecao
                bAchou = False
                For Each objRelCompConsItem In objRelCompCons.colItens
                    If sProduto = objRelCompConsItem.sProdutoFilho Then
                        objRelCompConsItem.dKitQtdePrevista = objRelCompConsItem.dKitQtdePrevista + dQuantidade
                        If objRelCompConsItem.iSequencial = 0 Then objRelCompConsItem.iSequencial = iSequencial
                        bAchou = True
                        Exit For
                    End If
                Next
                
                If bAchou = False Then
                    Set objRelCompConsItemNovo = New ClassRelCompConsItem
                    With objRelCompConsItemNovo
                        .sProdutoFilho = sProduto
                        .sProdutoFilhoDescricao = sDescricao
                        .sUMProdutoFilho = sSiglaUMEstoque
                        .dKitPerda = dPercentualPerda
                        .dKitQtdePrevista = dQuantidade
                        .iSequencial = iSequencial
                    End With
                    objRelCompCons.colItens.Add objRelCompConsItemNovo
                End If
                        
            End If
            
            'ler o proximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184138
            
        Loop
        
        'multiplicar pela qtde produzida proporcional ao kit
        dFator = objRelCompCons.dQtdeProduzida / ((1 - dPercentualPerdaBase) * dQuantidadeBase)
        For Each objRelCompConsItem In objRelCompCons.colItens
            
            objRelCompConsItem.dKitQtdePrevista = objRelCompConsItem.dKitQtdePrevista * dFator
            
            If objRelCompConsItem.sProdutoFilho = objRelCompCons.sProdutoPai Then dQuantidadeProprio = objRelCompConsItem.dKitQtdePrevista
            
        Next
        
        'se usou do proprio produto na producao
        If objRelCompCons.dQtdeReqPropProd <> 0 Then
        
            If objRelCompCons.dQtdeReqPropProd > dQuantidadeProprio Then
                dFator = 1 - ((objRelCompCons.dQtdeReqPropProd - dQuantidadeProprio) / objRelCompCons.dQtdeProduzida)
            Else
                dFator = objRelCompCons.dQtdeProduzida / (objRelCompCons.dQtdeProduzida - (dQuantidadeProprio - objRelCompCons.dQtdeReqPropProd))
            End If
        
            For Each objRelCompConsItem In objRelCompCons.colItens
                
                If objRelCompConsItem.sProdutoFilho <> objRelCompCons.sProdutoPai Then
                    objRelCompConsItem.dKitQtdePrevista = objRelCompConsItem.dKitQtdePrevista * dFator
                Else
                    objRelCompConsItem.dKitQtdePrevista = objRelCompCons.dQtdeReqPropProd
                End If
                
            Next
        
        End If
        
    End If
        
    RelPCPCompConsumo_Prepara1 = SUCESSO
 
    Exit Function
 
Erro_RelPCPCompConsumo_Prepara1:
  
    RelPCPCompConsumo_Prepara1 = gErr
 
    Select Case gErr
 
        Case 184136, 184137
        
        Case 184134, 184135, 184138
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELPCPCOMPCONS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184139)
 
    End Select
 
    Exit Function

End Function

Function RelPCPCompConsumo_Prepara2(ByVal objRelCompCons As ClassRelCompCons, lComando As Long) As Long
'obtem consumo real pelos movimentos de estoque

Dim lErro As Long, dFatorConv As Double, sSel As String, bAchou As Boolean
Dim objRelCompConsItem As ClassRelCompConsItem, objRelCompConsItemNovo As ClassRelCompConsItem
Dim dCusto As Double, iTipoMov As Integer, dQuantidade As Double, sSiglaUM As String, sSiglaUMEstoque As String, iClasseUM As Integer, sProduto As String, sDescricao As String

On Error GoTo Erro_RelPCPCompConsumo_Prepara2
    
    sSel = "SELECT Custo, TipoMov, Quantidade, SiglaUM, Produtos.SiglaUMEstoque, Produtos.ClasseUM, Produtos.Codigo, Produtos.Descricao FROM MovimentoEstoque, Produtos WHERE MovimentoEstoque.NumIntDocEst = 0 AND MovimentoEstoque.TipoNumIntDocOrigem = 5 AND MovimentoEstoque.NumIntDocOrigem = ? AND MovimentoEstoque.Produto = Produtos.Codigo AND Produtos.ControleEstoque <> ?"
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    'pesquisar movimentos ligados à producao ordenados por produto
    lErro = Comando_Executar(lComando, sSel, dCusto, iTipoMov, dQuantidade, sSiglaUM, sSiglaUMEstoque, iClasseUM, sProduto, sDescricao, objRelCompCons.lNumIntDocItemOP, PRODUTO_CONTROLE_SEM_ESTOQUE)
    If lErro <> SUCESSO Then gError 184128
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184129
        
    'ler
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Faz a Conversao de unidades de medida
        lErro = CF("UM_Conversao_Rel", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFatorConv)
        If lErro <> SUCESSO Then gError 184130
        
        dQuantidade = dQuantidade * dFatorConv
        
        Select Case iTipoMov
        
            Case MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3
                objRelCompCons.dQtdeProduzida = objRelCompCons.dQtdeProduzida + dQuantidade
                objRelCompCons.dCustoContabil = objRelCompCons.dCustoContabil + dCusto
            
            Case Else
                
                If iTipoMov = MOV_EST_DEV_MATERIAL_PRODUCAO Then
                    dQuantidade = -dQuantidade
                    dCusto = -dCusto
                End If
                
                If sProduto = objRelCompCons.sProdutoPai Then objRelCompCons.dQtdeReqPropProd = objRelCompCons.dQtdeReqPropProd + dQuantidade
                
                'pesquisa produto na colecao
                bAchou = False
                For Each objRelCompConsItem In objRelCompCons.colItens
                    If sProduto = objRelCompConsItem.sProdutoFilho Then
                        objRelCompConsItem.dQtdeConsumida = objRelCompConsItem.dQtdeConsumida + dQuantidade
                        objRelCompConsItem.dCustoContabil = objRelCompConsItem.dCustoContabil + dCusto
                        bAchou = True
                        Exit For
                    End If
                Next
                
                If bAchou = False Then
                    Set objRelCompConsItemNovo = New ClassRelCompConsItem
                    With objRelCompConsItemNovo
                        .dCustoContabil = dCusto
                        .dQtdeConsumida = dQuantidade
                        .sProdutoFilho = sProduto
                        .sProdutoFilhoDescricao = sDescricao
                        .sUMProdutoFilho = sSiglaUMEstoque
                    End With
                    objRelCompCons.colItens.Add objRelCompConsItemNovo
                End If
            
        End Select
        
        'ler o proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184131
        
    Loop
    
    RelPCPCompConsumo_Prepara2 = SUCESSO
 
    Exit Function
 
Erro_RelPCPCompConsumo_Prepara2:
  
    RelPCPCompConsumo_Prepara2 = gErr
 
    Select Case gErr
 
        Case 184130
        
        Case 184128, 184129, 184131
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELPCPCOMPCONS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184132)
 
    End Select
 
    Exit Function

End Function

Function RelPCPCompConsumo_Prepara3(ByVal objRelCompCons As ClassRelCompCons, lComando As Long) As Long
'obtem consumo de embalagens previsto pelo kit

Dim lErro As Long, dFatorConv As Double, sSel As String, dQtdeEmb As Double
Dim bAchou As Boolean, objRelCompConsItemNovo As ClassRelCompConsItem, objRelCompConsItem As ClassRelCompConsItem
Dim sProduto As String, sDescricao As String, iClasseUM As Integer, sSiglaUMEstoque As String, sUMEmbalagem As String, dCapacidade As Double

On Error GoTo Erro_RelPCPCompConsumo_Prepara3
    
    sSel = "SELECT Produtos.Codigo, Produtos.Descricao, Produtos.ClasseUM, Produtos.SiglaUMEstoque, PE.UMEmbalagem, PE.Capacidade FROM ProdutoEmbalagem PE, Embalagens, Produtos WHERE Embalagens.Produto = Produtos.Codigo AND PE.Embalagem = Embalagens.Codigo AND PE.Produto = ? AND Padrao = 1"
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sUMEmbalagem = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(lComando, sSel, sProduto, sDescricao, iClasseUM, sSiglaUMEstoque, sUMEmbalagem, dCapacidade, objRelCompCons.sProdutoPai)
    If lErro <> SUCESSO Then gError 184141
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184142
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
         'Faz a Conversao de unidades de medida
        lErro = CF("UM_Conversao_Trans", objRelCompCons.iClasseUMPai, objRelCompCons.sUMProdutoPai, sUMEmbalagem, dFatorConv)
        If lErro <> SUCESSO Then gError 184143
        
        'Transfere para o objProdutoKitProdutosTEMP os dados da embalagem
        dQtdeEmb = (objRelCompCons.dQtdeProduzida * dFatorConv) / dCapacidade
        
        'Arredonda para cima a divisao de Produto/Embalagem
        dQtdeEmb = Arredonda_ParaCima(dQtdeEmb)
   
        'pesquisa produto na colecao
        bAchou = False
        For Each objRelCompConsItem In objRelCompCons.colItens
            If sProduto = objRelCompConsItem.sProdutoFilho Then
                objRelCompConsItem.dKitQtdePrevista = objRelCompConsItem.dKitQtdePrevista + dQtdeEmb
                bAchou = True
                Exit For
            End If
        Next
        
        If bAchou = False Then
            Set objRelCompConsItemNovo = New ClassRelCompConsItem
            With objRelCompConsItemNovo
                .sProdutoFilho = sProduto
                .sProdutoFilhoDescricao = sDescricao
                .sUMProdutoFilho = sSiglaUMEstoque
                .dKitPerda = 0
                .dKitQtdePrevista = dQtdeEmb
            End With
            objRelCompCons.colItens.Add objRelCompConsItemNovo
        End If
    
    End If
    
    RelPCPCompConsumo_Prepara3 = SUCESSO
   
    Exit Function
 
Erro_RelPCPCompConsumo_Prepara3:
  
    RelPCPCompConsumo_Prepara3 = gErr
 
    Select Case gErr
 
        Case 184143
        
        Case 184141, 184142
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELPCPCOMPCONS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184144)
 
    End Select
 
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function Certificados_Grava(ByVal objCertificado As ClassCertificados) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Certificados_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213290

    lErro = CF("Certificados_Grava_EmTrans", objCertificado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213291

    Certificados_Grava = SUCESSO

    Exit Function

Erro_Certificados_Grava:

    Certificados_Grava = gErr

    Select Case gErr

        Case 213290
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213291
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213292)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Certificados_Grava_EmTrans(ByVal objCertificado As ClassCertificados) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lCodAux As Long

On Error GoTo Erro_Certificados_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213293
    Next

    'Verifica se tem repetição de siglas
    lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Certificados WHERE Sigla= ? AND Codigo <> ? ", _
                lCodAux, objCertificado.sSigla, objCertificado.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213314

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213315
    
    If lErro = AD_SQL_SUCESSO Then gError 213316

    'Le a tabelaCertificados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Certificados WHERE Codigo= ? ", _
                0, lCodAux, objCertificado.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213294

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213295

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 213296

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Certificados SET Descricao= ? , Sigla= ? , Validade= ? ", alComando(0), _
                objCertificado.sDescricao, objCertificado.sSigla, objCertificado.lValidade)
        If lErro <> AD_SQL_SUCESSO Then gError 213297

    Else

        lErro = Comando_Executar(alComando(1), "INSERT INTO Certificados( Codigo, Descricao, Sigla, Validade) VALUES (?,?,?,?)", _
                objCertificado.lCodigo, objCertificado.sDescricao, objCertificado.sSigla, objCertificado.lValidade)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213297

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Certificados_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Certificados_Grava_EmTrans:

    Certificados_Grava_EmTrans = gErr

    Select Case gErr

        Case 213293
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213294, 213295, 213314, 213315
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CERTIFICADOS", gErr)

        Case 213296
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CERTIFICADOS", gErr)

        Case 213297
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CERTIFICADOS", gErr)

        Case 213298
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CERTIFICADOS", gErr)

        Case 213316
            Call Rotina_Erro(vbOKOnly, "ERRO_CERTIFICADOS_SIGLA_JA_EXISTE", gErr, objCertificado.sSigla, lCodAux)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213299)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function Certificados_Exclui(ByVal objCertificado As ClassCertificados) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Certificados_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213300

    lErro = CF("Certificados_Exclui_EmTrans", objCertificado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213301

    Certificados_Exclui = SUCESSO

    Exit Function

Erro_Certificados_Exclui:

    Certificados_Exclui = gErr

    Select Case gErr

        Case 213300
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213301
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213302)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Certificados_Exclui_EmTrans(ByVal objCertificado As ClassCertificados) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodAux As Long

On Error GoTo Erro_Certificados_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213303
    Next

    'Le a tabelaCertificados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Certificados WHERE Codigo= ? ", _
                0, lCodAux, objCertificado.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213304

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213305

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 213306

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 213307
    
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_CERTIFICADOS, objCertificado.lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Certificados", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 213308

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Certificados_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Certificados_Exclui_EmTrans:

    Certificados_Exclui_EmTrans = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 213303
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213304, 213305
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CERTIFICADOS", gErr)

        Case 213306
            Call Rotina_Erro(vbOKOnly, "ERRO_CERTIFICADOS_NAO_CADASTRADO", gErr, objCertificado.lCodigo)

        Case 213307
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CERTIFICADOS", gErr)

        Case 213308
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CERTIFICADOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213309)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function Cursos_Grava(ByVal objCurso As ClassCursos) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Cursos_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213317

    lErro = CF("Cursos_Grava_EmTrans", objCurso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213318

    Cursos_Grava = SUCESSO

    Exit Function

Erro_Cursos_Grava:

    Cursos_Grava = gErr

    Select Case gErr

        Case 213317
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213318
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213319)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Cursos_Grava_EmTrans(ByVal objCurso As ClassCursos) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objMO As ClassCursoMO
Dim objCertificado As ClassCursoCertificados
Dim iSeq As Integer

On Error GoTo Erro_Cursos_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213320
    Next

    'Le a tabelaCursos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Cursos WHERE Codigo= ? AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objCurso.lCodigo, objCurso.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213321

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213322

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        objCurso.lNumIntDoc = lNumIntDoc

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 213323

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Cursos SET Detalhamento= ? , Responsavel= ? , DataInicio= ? , DataConclusao= ? ", alComando(0), _
                objCurso.sDetalhamento, objCurso.sResponsavel, objCurso.dtDataInicio, objCurso.dtDataConclusao)
        If lErro <> AD_SQL_SUCESSO Then gError 213324
        
        lErro = Comando_Executar(alComando(2), "DELETE CursoCertificados WHERE NumIntDocCurso = ?", objCurso.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213325

        lErro = Comando_Executar(alComando(3), "DELETE CursoMO WHERE NumIntDocCurso = ?", objCurso.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213326
    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_CURSOS", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objCurso.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO Cursos( NumIntDoc, Codigo, FilialEmpresa, Detalhamento, " & _
                "Responsavel, DataInicio, DataConclusao) VALUES (?,?,?,?,?,?,?)", _
                objCurso.lNumIntDoc, objCurso.lCodigo, objCurso.iFilialEmpresa, objCurso.sDetalhamento, _
                objCurso.sResponsavel, objCurso.dtDataInicio, objCurso.dtDataConclusao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213327

    End If
    
    For Each objCertificado In objCurso.colCertificados
    
        lErro = Comando_Executar(alComando(3), "INSERT INTO CursoCertificados(NumIntDocCurso,CodCertificado) VALUES (?,?)", _
                objCurso.lNumIntDoc, objCertificado.lCodCertificado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213328
    
    Next
    
    iSeq = 0
    For Each objMO In objCurso.colMOCursos
    
        iSeq = iSeq + 1
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO CursoMO(NumIntDocCurso,Seq,CodMO,Avaliacao,Aprovado) VALUES (?,?,?,?,?)", _
                objCurso.lNumIntDoc, iSeq, objMO.iCodMO, objMO.sAvaliacao, objMO.iAprovado)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213329
    
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cursos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Cursos_Grava_EmTrans:

    Cursos_Grava_EmTrans = gErr

    Select Case gErr

        Case 213320
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213321, 213322
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CURSOS", gErr)

        Case 213323
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CURSOS", gErr)

        Case 213324
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CURSOS", gErr)

        Case 213325
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_CURSOCERTIFICADOS", gErr)

        Case 213326
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_CURSOMO", gErr)

        Case 213327
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CURSOS", gErr)

        Case 213328
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CURSOCERTIFICADOS", gErr)

        Case 213329
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CURSOMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213330)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function Cursos_Exclui(ByVal objCurso As ClassCursos) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Cursos_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213331

    lErro = CF("Cursos_Exclui_EmTrans", objCurso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213332

    Cursos_Exclui = SUCESSO

    Exit Function

Erro_Cursos_Exclui:

    Cursos_Exclui = gErr

    Select Case gErr

        Case 213331
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213332
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213333)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function Cursos_Exclui_EmTrans(ByVal objCurso As ClassCursos) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Integer

On Error GoTo Erro_Cursos_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213334
    Next

    'Le a tabelaCursos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Cursos WHERE Codigo= ? AND FilialEmpresa= ? ", _
                0, lNumIntDoc, objCurso.lCodigo, objCurso.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213335

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213336

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 213337

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 213338

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Cursos", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 213339
    
    lErro = Comando_Executar(alComando(2), "DELETE CursoCertificados WHERE NumIntDocCurso = ?", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213340

    lErro = Comando_Executar(alComando(3), "DELETE CursoMO WHERE NumIntDocCurso = ?", lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213341

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cursos_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Cursos_Exclui_EmTrans:

    Cursos_Exclui_EmTrans = gErr

    Select Case gErr

        Case 213334
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213335, 213336
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CURSOS", gErr)

        Case 213337
            Call Rotina_Erro(vbOKOnly, "ERRO_CURSOS_NAO_CADASTRADO", gErr, objCurso.lCodigo, objCurso.iFilialEmpresa)

        Case 213338
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CURSOS", gErr)

        Case 213339
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CURSOS", gErr)

        Case 213340
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CURSOCERTIFICADOS", gErr)

        Case 213341
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CURSOMO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213342)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


