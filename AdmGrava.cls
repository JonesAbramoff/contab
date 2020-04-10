VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Function BrowseUsuario_Atualiza(ByVal objBrowseUsuario As AdmBrowseUsuario) As Long
'grava na tabela BrowseUsuario os campos contidos em objBrowseUsuario

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim sNomeTela As String, sCodUsuario As String

On Error GoTo Erro_BrowseUsuario_Atualiza

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 91294
    Next

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 91295
    
    sNomeTela = String(STRING_NOME_TELA, 0)
    sCodUsuario = String(STRING_CODUSUARIO, 0)

    'verifica o BrowseUsuario está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NomeTela, CodUsuario FROM BrowseUsuario WHERE NomeTela=? AND CodUsuario=?", 0, sNomeTela, sCodUsuario, objBrowseUsuario.sNomeTela, objBrowseUsuario.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 91296
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91297

    'se BrowseUsuario estiver cadastrado ==> é uma alteração
    If lErro = SUCESSO Then
        
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 91298
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE BrowseUsuario SET Topo = ?, Esquerda = ?, Largura = ?, Altura = ? ", alComando(0), objBrowseUsuario.lTopo, objBrowseUsuario.lEsquerda, objBrowseUsuario.lLargura, objBrowseUsuario.lAltura)
        If lErro <> AD_SQL_SUCESSO Then gError 91299
    
    Else
    
        'se BrowseUsuario não estiver cadastrado ==> é uma inclusão
        lErro = Comando_Executar(alComando(2), "INSERT INTO BrowseUsuario (NomeTela, CodUsuario, Topo, Esquerda, Largura, Altura) VALUES (?, ?, ?, ?, ?, ?)", objBrowseUsuario.sNomeTela, objBrowseUsuario.sCodUsuario, objBrowseUsuario.lTopo, objBrowseUsuario.lEsquerda, objBrowseUsuario.lLargura, objBrowseUsuario.lAltura)
        If lErro <> AD_SQL_SUCESSO Then gError 91300
        
    End If

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 91301

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   BrowseUsuario_Atualiza = SUCESSO
   
   Exit Function
   
Erro_BrowseUsuario_Atualiza:

   BrowseUsuario_Atualiza = gErr

   Select Case gErr

        Case 91294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 91295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 91296, 91297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIO", gErr)
      
        Case 91298
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_BROWSEUSUARIO", gErr)

        Case 91299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BROWSEUSUARIO", gErr)

        Case 91300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEUSUARIO", gErr)

        Case 91301
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142374)

   End Select
   
   'Call Transacao_Rollback
   Call Transacao_RollbackDic

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function BrowseUsuarioCampo_Grava(ByVal colBrowseUsuarioCampo As Collection) As Long
'exclui as informacoes antigas para o usuario x tela em questao e grava as novas
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lComando1 As Long
Dim lErro As Long
Dim sNome As String
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
    
On Error GoTo Erro_BrowseUsuarioCampo_Grava
    
    Set objBrowseUsuarioCampo = colBrowseUsuarioCampo.Item(1)
    
    lComando1 = 0
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 9086
    
    sNome = String(STRING_NOME_TELA, 0)
    
    'tenta fazer o lock da tela em questão
    lErro = Comando_ExecutarLockado(lComando1, "SELECT Nome FROM Telas WHERE Nome=?", sNome, objBrowseUsuarioCampo.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 9087
    
    'le a tela em questão
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9088
    
    'Código comentado por Wagner - Tranferido para a Exclusão
'    lErro = Comando_LockShared(lComando1)
'    If lErro <> AD_SQL_SUCESSO Then Error 9089
    
    'exclui os campos desta tela x usuario
    lErro = BrowseUsuarioCampo_Exclui_BD(objBrowseUsuarioCampo.sNomeTela, objBrowseUsuarioCampo.sCodUsuario)
    If lErro Then Error 9091

    lErro = BrowseUsuarioCampo_Grava_BD(colBrowseUsuarioCampo)
    If lErro Then Error 9092
    
    lErro = Comando_Fechar(lComando1)
    
    BrowseUsuarioCampo_Grava = SUCESSO

    Exit Function

Erro_BrowseUsuarioCampo_Grava:
    
    BrowseUsuarioCampo_Grava = Err
    
    Select Case Err
    
        Case 9086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9087, 9088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAS", Err)
            
'        Case 9089
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TELAS", Err)
                   
        Case 9091, 9092
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142375)
            
    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function BrowseUsuarioCampo_Grava_BD(ByVal colBrowseUsuarioCampo As Collection) As Long
'grava na tabela BrowseUsuarioCampo os campos contidos na coleção
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

On Error GoTo Erro_BrowseUsuarioCampo_Grava_BD

    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9093
    
    'Inicia Loop de GRAVACAO
    For Each objBrowseUsuarioCampo In colBrowseUsuarioCampo
            
        'Faz inserção do campo na Tabela
        lErro = Comando_Executar(lComando, "INSERT INTO BrowseUsuarioCampo (NomeTela,CodUsuario,NomeArq,Nome,PosicaoTela,Titulo, Largura) VALUES (?,?,?,?,?,?,?)", objBrowseUsuarioCampo.sNomeTela, objBrowseUsuarioCampo.sCodUsuario, objBrowseUsuarioCampo.sNomeArq, objBrowseUsuarioCampo.sNome, objBrowseUsuarioCampo.iPosicaoTela, objBrowseUsuarioCampo.sTitulo, objBrowseUsuarioCampo.lLargura)
        If lErro <> AD_SQL_SUCESSO Then Error 9094
        
    Next

    lErro = Comando_Fechar(lComando)
    
    BrowseUsuarioCampo_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioCampo_Grava_BD:

    BrowseUsuarioCampo_Grava_BD = Err
    
    Select Case Err
    
        Case 9093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEUSUARIOCAMPO", Err)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142376)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function BrowseUsuarioCampo_Exclui_BD(ByVal sNomeTela As String, ByVal sCodUsuario As String)
'exclui da tabela de BrowseUsuarioCampo os registros referentes a tela x usuario em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim sNome As String

On Error GoTo Erro_BrowseUsuarioCampo_Exclui_BD

    lComando = 0
    lComando1 = 0

    'Abre comando para pesquisar os campos
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 9095
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 9096
    
    sNome = String(STRING_NOME_CAMPO, 0)
    
    'Pesquisa a existencia dos registros da tela x usuario em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Nome FROM BrowseUsuarioCampo WHERE NomeTela=? AND CodUsuario=?", 0, sNome, sNomeTela, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 9097
     
    'Lê o primeiro campo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9098

    'Inicia Loop de EXCLUSAO
    Do While lErro <> AD_SQL_SEM_DADOS
     
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9089
        
        'Exclui o Campo da  Tabela de BrowseUsuarioCampo
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BrowseUsuarioCampo", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9099
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9100
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    BrowseUsuarioCampo_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioCampo_Exclui_BD:

    BrowseUsuarioCampo_Exclui_BD = Err
    
    Select Case Err
    
        Case 9095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9097, 9098, 9100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOCAMPO", Err)
           
        Case 9099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO", Err)
            
        Case 9089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_BROWSEUSUARIOCAMPO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142377)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function BrowseUsuarioOrdenacao_Grava(ByVal objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'grava a informacao de ordenacao para o usuario x tela em questão
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim sNomeTela As String

On Error GoTo Erro_BrowseUsuarioOrdenacao_Grava

    lComando1 = 0
    lComando2 = 0

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 9102

    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando2 = 0 Then Error 9103

    sNomeTela = String(STRING_NOME_TELA, 0)

    'pesquisa no BD as informacoes de ordenacao
    lErro = Comando_ExecutarPos(lComando1, "SELECT NomeTela FROM BrowseUsuarioOrdenacao WHERE NomeTela=? AND CodUsuario=?", 0, sNomeTela, objBrowseUsuarioOrdenacao.sNomeTela, objBrowseUsuarioOrdenacao.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 9105

    'le a ordenacao, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9106

    'encontrou a ordenacao ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        'atualiza a ordenação
        lErro = Comando_ExecutarPos(lComando2, "UPDATE BrowseUsuarioOrdenacao SET Indice=?, SelecaoSQL1=?, SelecaoSQL1Usuario=?, NomeIndice=?", lComando1, objBrowseUsuarioOrdenacao.iIndice, objBrowseUsuarioOrdenacao.sSelecaoSQL1, objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario, objBrowseUsuarioOrdenacao.sNomeIndice)
        If lErro <> AD_SQL_SUCESSO Then Error 9107

    Else
        'não encontrou a ordenação ==>  é uma inserção de ordenação
        lErro = Comando_Executar(lComando2, "INSERT INTO BrowseUsuarioOrdenacao (NomeTela, CodUsuario, Indice, SelecaoSQL1, SelecaoSQL1Usuario, NomeIndice) VALUES (?,?,?,?,?,?)", objBrowseUsuarioOrdenacao.sNomeTela, objBrowseUsuarioOrdenacao.sCodUsuario, objBrowseUsuarioOrdenacao.iIndice, objBrowseUsuarioOrdenacao.sSelecaoSQL1, objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario, objBrowseUsuarioOrdenacao.sNomeIndice)
        If lErro <> AD_SQL_SUCESSO Then Error 9108

    End If

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)

    BrowseUsuarioOrdenacao_Grava = SUCESSO

    Exit Function

Erro_BrowseUsuarioOrdenacao_Grava:

    BrowseUsuarioOrdenacao_Grava = Err
    
    Select Case Err

        Case 9102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9105, 9106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOORDENACAO", Err, objBrowseUsuarioOrdenacao.sNomeTela, objBrowseUsuarioOrdenacao.sCodUsuario)
            
        Case 9107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BROWSEUSUARIOORDENACAO", Err, objBrowseUsuarioOrdenacao.sNomeTela, objBrowseUsuarioOrdenacao.sCodUsuario)
            
        Case 9108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEUSUARIOORDENACAO", Err, objBrowseUsuarioOrdenacao.sNomeTela, objBrowseUsuarioOrdenacao.sCodUsuario)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142378)

    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Relatorio_Grava(ByVal objRelatorio As AdmRelatorio, iOperacao As Integer, ByVal sModulo As String) As Long
'grava o Relatorio no banco de dados

Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim lErro As Long, iOrigem As Integer

On Error GoTo Erro_Relatorio_Grava
    
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0
    lTransacao = 0
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 12024

    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando2 = 0 Then Error 12025
    
    lComando3 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando3 = 0 Then Error 12026

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 12027

    'pesquisa no BD o Relatorio em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Origem FROM Relatorios WHERE CodRel = ?", 0, iOrigem, objRelatorio.sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 12028

    'le o Relatorio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12029

    'encontrou o Relatorio
    If lErro = AD_SQL_SUCESSO Then
    
        If iOrigem = REL_ORIGEM_FORPRINT Then Error 59298
        
        iOperacao = MODIFICACAO
    
        'atualiza o Relatorio
        lErro = Comando_ExecutarPos(lComando3, "UPDATE Relatorios SET NomeTsk = ?, Autor = ? , Descricao = ? , TelaObterParam = ?, Landscape = ?", lComando1, objRelatorio.sNomeTsk, objRelatorio.sAutor, objRelatorio.sDescricao, objRelatorio.sTelaAuxiliar, objRelatorio.iLandscape)
        If lErro <> AD_SQL_SUCESSO Then Error 12032

    Else
        
        'não encontrou o Relatorio
        'insere relatorio em Relatorios , GrupoRelatorios e RelatoriosModulo no BD
        iOperacao = GRAVACAO
        lErro = Comando_Executar(lComando2, "INSERT INTO Relatorios (CodRel, NomeTsk, Autor, Descricao, TelaObterParam, Landscape, Origem) VALUES (?,?,?,?,?,?,?)", objRelatorio.sCodRel, objRelatorio.sNomeTsk, objRelatorio.sAutor, objRelatorio.sDescricao, objRelatorio.sTelaAuxiliar, objRelatorio.iLandscape, REL_ORIGEM_USUARIO)
        If lErro <> AD_SQL_SUCESSO Then Error 12033
        
        lErro = Comando_Executar(lComando2, "INSERT INTO GrupoRelatorios (CodGrupo, CodRel, TipoDeAcesso, NomeTskCustomizado) SELECT CodGrupo, ?, ?, ? FROM GruposDeUsuarios", objRelatorio.sCodRel, COM_ACESSO, "")
        If lErro <> AD_SQL_SUCESSO Then Error 12034
        
        lErro = Comando_Executar(lComando2, "INSERT INTO RelatoriosModulo (Modulo, CodRel) VALUES (?,?)", sModulo, objRelatorio.sCodRel)
        If lErro <> AD_SQL_SUCESSO Then Error 12035
        
    End If

    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 12034

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando3)
    
    Relatorio_Grava = SUCESSO

    Exit Function

Erro_Relatorio_Grava:

    Relatorio_Grava = Err
    
    Select Case Err '??? falta acertar

        Case 59298
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_ORIGINAL_INALTERAVEL", Err)
        
        Case 12024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142379)

        Case 12025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142380)
            
        Case 12026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142381)
            
        Case 12027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142382)
            
        Case 12028, 12029, 12032, 12033, 12034, 12035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142383)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142384)
                        
    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Relatorio_Exclui(ByVal objRelatorio As AdmRelatorio, ByVal sModulo As String) As Long
'Deleta o Relatorio no banco de dados

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long, iOrigem As Integer

On Error GoTo Erro_Relatorio_Exclui
             
    lComando1 = 0
    lComando2 = 0
    lTransacao = 0
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Or lComando2 = 0 Then Error 12055

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 12056

    'Pesquisa a existencia
    lErro = Comando_ExecutarPos(lComando1, "SELECT Origem FROM Relatorios WHERE CodRel=?", 0, iOrigem, objRelatorio.sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 7064
     
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7065
    If lErro <> AD_SQL_SUCESSO Then Error 59300

    If iOrigem = REL_ORIGEM_FORPRINT Then Error 59299
        
    'deleta o Relatorio da tabela Relatorios
    lErro = Comando_ExecutarPos(lComando2, "Delete FROM Relatorios", lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 12057
    
    'deleta o Relatorio da tabela GrupoRelatorios
    lErro = Comando_Executar(lComando2, "DELETE FROM GrupoRelatorios WHERE CodRel = ? ", objRelatorio.sCodRel) 'Alterado por Wagner
    If lErro <> AD_SQL_SUCESSO Then Error 12058
    
    'deleta o Relatorio da tabela RelatoriosModulo
    lErro = Comando_Executar(lComando2, "DELETE FROM RelatoriosModulo WHERE CodRel = ? ", objRelatorio.sCodRel) 'Alterado por Wagner
    If lErro <> AD_SQL_SUCESSO Then Error 12059
        
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 12060
    lTransacao = 0
    
    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)
        
    Relatorio_Exclui = SUCESSO

    Exit Function

Erro_Relatorio_Exclui:

    Relatorio_Exclui = Err

    Select Case Err '??? falta acertar

        Case 59299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_ORIGINAL_INALTERAVEL", Err)
        
        Case 59300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_NAO_CADASTRADO", Err, objRelatorio.sCodRel)
        
        Case 7064
        
        Case 7065
        
        Case 12055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142385)

        Case 12056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142386)
            
        Case 12057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142387)
            
        Case 12058, 12059, 12060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142388)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142389)
                        
    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function RelOpcoes_Grava(ByVal objRelOpcoes As AdmRelOpcoes, iResultado As Integer) As Long
'inclui ou modifica um registro na tabela RelatorioOpcoes
'iResultado retorna GRAVACAO ou MODIFICACAO

Dim lErro As Long, iFim_de_Arquivo As Integer
Dim lComando As Long, lComando1 As Long
Dim lTransacao As Long
Dim sArqC_Ant As String
Dim sTempCod As String, sTempNome As String, sTempArq As String

On Error GoTo Erro_RelatorioOpcoes_Gravar

    sArqC_Ant = String(STRING_NOME_ARQ_COMPLETO, 0)
    sTempArq = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    sTempCod = objRelOpcoes.sCodRel
    sTempNome = objRelOpcoes.sNome
    
    lComando = 0
    lComando1 = 0
    lTransacao = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 13003

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 13004

    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 13005

    'verifica se a opção já existe em RelatorioOpcoes
    lErro = Comando_ExecutarPos(lComando, "SELECT ArqParametros FROM RelatorioOpcoes WHERE CodRel = ? AND Nome = ?", 0, sArqC_Ant, sTempCod, sTempNome)
    If lErro <> AD_SQL_SUCESSO Then Error 13006
    
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_ERRO Then Error 13007
    
    lErro = objRelOpcoes.Salvar()
    
    sTempArq = objRelOpcoes.sArqParam

    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then
    
        iResultado = GRAVACAO

        'não existe, então incluir
        lErro = Comando_Executar(lComando, "INSERT INTO RelatorioOpcoes (CodRel,Nome,ArqParametros) VALUES (?,?,?)", sTempCod, sTempNome, sTempArq)
        If lErro <> AD_SQL_SUCESSO Then Error 13008
    
    Else
    
        iResultado = MODIFICACAO
        
        'existe, então alterar
        lErro = Comando_ExecutarPos(lComando1, "UPDATE RelatorioOpcoes SET ArqParametros = ?", lComando, objRelOpcoes.sArqParam)
        If lErro <> AD_SQL_SUCESSO Then Error 13009
    
    End If
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 13010
    
    'apagar arquivo C anterior
    'If iResultado = MODIFICACAO Then Kill sArqC_Ant
        
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    RelOpcoes_Grava = SUCESSO
    
    Exit Function
    
Erro_RelatorioOpcoes_Gravar:

    RelOpcoes_Grava = Err
    
    Select Case Err

        Case 13003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 13005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 13006, 13007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOOPCOES", Err)
            
        Case 13008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELATORIOOPCOES", Err)
            
        Case 13009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELATORIOOPCOES", Err)
                   
        Case 13010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142390)

    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Function RelOpcoes_Exclui(ByVal objRelOpcoes As AdmRelOpcoes) As Long
'exclui da tabela RelatorioOpcoes o registro referente a opção em questão

Dim sArqParam As String
Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_RelOpcoesExclui

    sArqParam = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    lComando = 0
    lComando1 = 0
    lTransacao = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 13018
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 13019

    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 13020
    
    'Exclui os Parametros da Opcao passada
    lErro = RelOpcoes_Exclui_Parametros(objRelOpcoes)
    If lErro <> SUCESSO Then Error 61272
    
    'pesquisa a existencia do registro com o código do relatório (CodRel) e o nome das opções ( Nome ) na tabela RelatorioOpcoes
    lErro = Comando_ExecutarPos(lComando, "SELECT ArqParametros FROM RelatorioOpcoes WHERE CodRel = ? AND Nome = ?", 0, sArqParam, objRelOpcoes.sCodRel, objRelOpcoes.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 13021

    'le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 13022

    'lock no registro
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 13023

    'exclui o registro da tabela RelatorioOpcoes
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RelatorioOpcoes", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 13024
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 13025

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    'apaga arquivo C de parâmetros do disco
    'Kill sArqParam
    
    RelOpcoes_Exclui = SUCESSO

    Exit Function

Erro_RelOpcoesExclui:

    RelOpcoes_Exclui = Err
    
    Select Case Err

        Case 13018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 13019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                        
        Case 13020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                        
        Case 13021, 13022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOOPCOES", Err)
                        
        Case 13023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELATORIOOPCOES", Err)
                        
        Case 13024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELATORIOOPCOES", Err)
                        
        Case 13025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 61272 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142391)

    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function RelOpcoes_Exclui_Parametros(ByVal objRelOpcoes As AdmRelOpcoes) As Long
'exclui da tabela RelatorioParametros os Parametros referente a opção em questão

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sOpcao As String

On Error GoTo Erro_RelOpcoes_Exclui_Parametros
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 61253
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 61254
    
    sOpcao = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    
    'pesquisa a existencia do registro com o código do relatório (CodRel) e o nome das opções ( Nome ) na tabela RelatorioParametros
    lErro = Comando_ExecutarPos(lComando, "SELECT Nome FROM RelatorioParametros WHERE CodRel = ? AND Nome = ?", 0, sOpcao, objRelOpcoes.sCodRel, objRelOpcoes.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 61255

    'le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61256

    'lock no registro
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61257
    
    Do While lErro = AD_SQL_SUCESSO
    
        'exclui o registro da tabela RelatorioOpcoes
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RelatorioParametros", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 61258
        
        'lê a primeira opção de relatório
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61273

    Loop
    
    Call objRelOpcoes.Excluir
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    RelOpcoes_Exclui_Parametros = SUCESSO

    Exit Function

Erro_RelOpcoes_Exclui_Parametros:

    RelOpcoes_Exclui_Parametros = Err
    
    Select Case Err

        Case 61253, 61254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 61255, 61256, 61273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOPARAMETROS", Err)
                        
        Case 61257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELATORIOPARAMETROS", Err)
                        
        Case 61258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELATORIOPARAMETROS", Err)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142392)

    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

'Function Usuario_Grava1(objUsuario As ClassUsuario) As Long
''inclui ou modifica um registro na tabela Usuario
''IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO
'
'Dim lErro As Long
'Dim lComando As Long, lComando1 As Long, lTransacao As Long
'Dim iLote As Integer
'
'On Error GoTo Erro_Usuario_Grava1
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 39704
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then Error 39705
'
'    'verifica se o usuario está cadastrado
'    lErro = Comando_ExecutarPos(lComando, "SELECT Lote FROM Usuario WHERE CodUsuario = ?", 0, iLote, objUsuario.sCodUsuario)
'    If lErro <> AD_SQL_SUCESSO Then Error 39706
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39707
'
'    If lErro = SUCESSO Then
'
'        'existe, então alterar
'        lErro = Comando_ExecutarPos(lComando1, "UPDATE Usuario SET Lote = ?", lComando, objUsuario.iLote)
'        If lErro <> AD_SQL_SUCESSO Then Error 39708
'
'    Else
'
'        'não existe, então incluir
'        lErro = Comando_Executar(lComando, "INSERT INTO Usuario (CodUsuario,Lote) VALUES (?,?)", objUsuario.sCodUsuario, objUsuario.iLote)
'        If lErro <> AD_SQL_SUCESSO Then Error 39709
'
'
'    End If
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Usuario_Grava1 = SUCESSO
'
'    Exit Function
'
'Erro_Usuario_Grava1:
'
'    Usuario_Grava1 = Err
'
'    Select Case Err
'
'        Case 39704, 39705
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 39706, 39707
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case 39708
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case 39709
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142393)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function

Function Config_ObterAutomatico(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, lNumAuto As Long, Optional ByVal sCampoFilial As String = "FilialEmpresa") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim lNumeroBD As Long
Dim lTransacao As Long
Dim sComandoSQL As String
Dim sConteudo As String, iDependeFilial As Integer

On Error GoTo Erro_Config_ObterAutomatico
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 25192
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 25193
    
    sConteudo = String(STRING_BUFFER_MAX_TEXTO, 0)
    'Monta comando SQL de leitura na tabela de configuração
    sComandoSQL = "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?"

    'Pesquisa o número automático que está no BD para giFilialEmpresa
    lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 25194

    'Le o registro com o número automático
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25195
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa = EMPRESA_TODA Then Error 25202
    
    'se o codigo depende da filial
    iDependeFilial = 0
    
    'Se não achar pesquisa o número automático que está no BD para EMPRESA_TODA
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa <> EMPRESA_TODA Then
    
        lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 25203
    
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 25204
        
    Else
    
        If giFilialEmpresa <> EMPRESA_TODA Then iDependeFilial = 1
        
    End If
    
    'Locka o registro
    lErro = Comando_LockExclusive(lComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 25196

    lNumAuto = CLng(sConteudo)

    'Verifica se número automático ainda não foi utilizado
    
    If iDependeFilial = 0 Then
    
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, lNumeroBD, lNumAuto)
        If lErro <> AD_SQL_SUCESSO Then Error 25197
    
    Else
    
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampoFilial & " = ? AND " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, lNumeroBD, giFilialEmpresa, lNumAuto)
        If lErro <> AD_SQL_SUCESSO Then Error 32212
        
    End If
    
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25198

    'pesquisa o proximo número automático disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        Select Case lNumeroBD
            Case Is > lNumAuto
                Exit Do
            Case Is = lNumAuto
                lNumAuto = lNumAuto + 1
        End Select

        lErro = Comando_BuscarProximo(lComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25199

    Loop

    sConteudo = CStr(lNumAuto + 1)
    
    'Monta Comando SQL de atualizar número automático na tabela configuração
    sComandoSQL = "UPDATE " & sTabelaConfig & " SET Conteudo = ?"
    
    'Atualiza o numero automático no BD
    lErro = Comando_ExecutarPos(lComando(3), sComandoSQL, lComando(1), sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 25200

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 25201

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_ObterAutomatico = SUCESSO

    Exit Function

Erro_Config_ObterAutomatico:

    Config_ObterAutomatico = Err

    Select Case Err

    Case 25192
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 25193
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

    Case 25194, 25197, 25203, 32212
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaConfig)

    Case 25195, 25202, 25204
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25196
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25198, 25199
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaObjeto)
    
    Case 25200
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25201
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142394)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function

Function Config_ObterAutomatico_EmTrans(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, lNumAuto As Long, Optional ByVal sCampoFilial As String = "FilialEmpresa") As Long
'CHAMADA EM TRANSACAO
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim lNumeroBD As Long
Dim lTransacao As Long
Dim sComandoSQL As String
Dim sConteudo As String, iDependeFilial As Integer

On Error GoTo Erro_Config_ObterAutomatico_EmTrans
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 25456
    Next

    sConteudo = String(STRING_BUFFER_MAX_TEXTO, 0)
    'Monta comando SQL de leitura na tabela de configuração
    sComandoSQL = "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?"

    'Pesquisa o número automático que está no BD para giFilialEmpresa
    lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 25446

    'Le o registro com o número automático
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25447
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa = EMPRESA_TODA Then Error 25448
    
    'se o codigo depende da filial
    iDependeFilial = 0
    
    'Se não achar pesquisa o número automático que está no BD para EMPRESA_TODA
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa <> EMPRESA_TODA Then
    
        lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 25449
    
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 25450
        
    Else
    
        If giFilialEmpresa <> EMPRESA_TODA Then iDependeFilial = 1
        
    End If
    
    'Locka o registro
    lErro = Comando_LockExclusive(lComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 25451

    lNumAuto = CLng(sConteudo)

    If iDependeFilial = 0 Then

        'Verifica se número automático ainda não foi utilizado
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, lNumeroBD, lNumAuto)
        If lErro <> AD_SQL_SUCESSO Then Error 25452

    Else
    
        'Verifica se número automático ainda não foi utilizado
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampoFilial & " = ? AND " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, lNumeroBD, giFilialEmpresa, lNumAuto)
        If lErro <> AD_SQL_SUCESSO Then Error 20847
    
    End If

    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25453

    'pesquisa o proximo número automático disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        Select Case lNumeroBD
            Case Is > lNumAuto
                Exit Do
            Case Is = lNumAuto
                lNumAuto = lNumAuto + 1
        End Select

        lErro = Comando_BuscarProximo(lComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25454

    Loop

    sConteudo = CStr(lNumAuto + 1)
    
    'Monta Comando SQL de atualizar número automático na tabela configuração
    sComandoSQL = "UPDATE " & sTabelaConfig & " SET Conteudo = ?"
    
    'Atualiza o numero automático no BD
    lErro = Comando_ExecutarPos(lComando(3), sComandoSQL, lComando(1), sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 25455

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_ObterAutomatico_EmTrans = SUCESSO

    Exit Function

Erro_Config_ObterAutomatico_EmTrans:

    Config_ObterAutomatico_EmTrans = Err

    Select Case Err

    Case 20847, 25446, 25452, 25449
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaConfig)

    Case 25456
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 25447, 25448, 25450
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25451
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25453, 25454
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaObjeto)
    
    Case 25455
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142395)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function

Function Config_ObterNumInt(ByVal sTabelaConfig As String, ByVal sCodigo As String, lNumInt As Long, Optional ByVal iQtde As Integer = 1, Optional ByVal iFilialEmpresa As Integer = EMPRESA_TODA) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComandoSel As Long, lComandoMod As Long
Dim sConteudo As String

On Error GoTo Erro_Config_ObterNumInt

    If iQtde < 0 Then Error 32287
    
    sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)

    lComandoSel = Comando_Abrir()
    If lComandoSel = 0 Then Error 19301

    lComandoMod = Comando_Abrir()
    If lComandoMod = 0 Then Error 19302

    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 19303

    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19304

    If lErro = AD_SQL_SEM_DADOS Then Error 20737
    
    lErro = Comando_LockExclusive(lComandoSel)
    If lErro <> AD_SQL_SUCESSO Then Error 32281
    
    lNumInt = CLng(sConteudo)
    sConteudo = CStr(lNumInt + iQtde)

    lErro = Comando_ExecutarPos(lComandoMod, "UPDATE " & sTabelaConfig & " SET Conteudo = ?", lComandoSel, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 19305

    lErro = Comando_Fechar(lComandoSel)
    lErro = Comando_Fechar(lComandoMod)

    Config_ObterNumInt = SUCESSO

    Exit Function

Erro_Config_ObterNumInt:

    Config_ObterNumInt = Err

    Select Case Err

        Case 19301, 19302
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 19303, 19304
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

        Case 19305
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CPRCONFIG", Err, sCodigo)

        Case 20737
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFIG_NAO_CADASTRADO", Err, sCodigo, sTabelaConfig)
            
        Case 32281
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)
        
        Case 32287
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_INVALIDA_NUMINT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142396)

    End Select

    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoMod)

    Exit Function

End Function

Function Config_ObterNumInt2(ByVal lComandoSel As Long, ByVal lComandoMod As Long, ByVal sTabelaConfig As String, ByVal sCodigo As String, lNumInt As Long, Optional ByVal iQtde As Integer = 1, Optional ByVal iFilialEmpresa As Integer = EMPRESA_TODA) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Config_ObterNumInt2

    If iQtde < 0 Then Error 32287
    
    sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 19303

    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19304

    If lErro = AD_SQL_SEM_DADOS Then Error 20737
    
    lErro = Comando_LockExclusive(lComandoSel)
    If lErro <> AD_SQL_SUCESSO Then Error 32281
    
    lNumInt = CLng(sConteudo)
    sConteudo = CStr(lNumInt + iQtde)

    lErro = Comando_ExecutarPos(lComandoMod, "UPDATE " & sTabelaConfig & " SET Conteudo = ?", lComandoSel, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 19305

    Config_ObterNumInt2 = SUCESSO

    Exit Function

Erro_Config_ObterNumInt2:

    Config_ObterNumInt2 = Err

    Select Case Err

        Case 19301, 19302
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 19303, 19304
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

        Case 19305
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CPRCONFIG", Err, sCodigo)

        Case 20737
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFIG_NAO_CADASTRADO", Err, sCodigo, sTabelaConfig)
            
        Case 32281
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)
        
        Case 32287
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_INVALIDA_NUMINT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142396)

    End Select

    Exit Function

End Function

Function Config_Obter_Inteiro_Automatico(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, iNumAuto As Integer) As Long
'retorna em iNumAuto(inteiro), o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lTransacao As Long, lErro As Long

On Error GoTo Erro_Config_Obter_Inteiro_Automatico
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 46515

    lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, iNumAuto)
    If lErro <> SUCESSO Then gError 130746

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 46526

    Config_Obter_Inteiro_Automatico = SUCESSO

    Exit Function

Erro_Config_Obter_Inteiro_Automatico:

    Config_Obter_Inteiro_Automatico = gErr

    Select Case gErr

        Case 130746
        
        Case 46515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 46526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130745)

    End Select

    Call Transacao_Rollback
    
End Function

Function Paises_Grava(ByVal objPais As ClassPais) As Long
'Grava o Pais no Banco de Dados

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Paises_Grava
    
    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lTransacao = 0
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49740
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 47856

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 47857

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 47858
    
    'Pesquisa no Banco de Dados o Pais em questão
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Paises WHERE Nome = ? AND Codigo <> ?", iCodigo, objPais.sNome, objPais.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49741
    
    'Se estiver no Banco de Dados
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49742

    'Encontrou o Pais com o Nome em questão ----> Erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 49743
    
    'Pesquisa no Banco de Dados o Pais em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM Paises WHERE Codigo=?", 0, iCodigo, objPais.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 47859

    'Lê os dados do Pais, se estiver no Banco de Dados
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 47860

    'Encontrou o Pais, então isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        'Atualiza os dados do Pais
        lErro = Comando_ExecutarPos(lComando2, "UPDATE Paises SET Nome=?, CodBacen = ?", lComando1, objPais.sNome, objPais.iCodBacen)
        If lErro <> AD_SQL_SUCESSO Then Error 47861

    'Não encontrou o Pais, então isto é uma inserção
    Else

        'Insere o Pais
        lErro = Comando_Executar(lComando2, "INSERT INTO Paises (Codigo,Nome,CodBacen) VALUES (?,?,?)", objPais.iCodigo, objPais.sNome, objPais.iCodBacen)
        If lErro <> AD_SQL_SUCESSO Then Error 47862

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 47863

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    Paises_Grava = SUCESSO

    Exit Function

Erro_Paises_Grava:

    Paises_Grava = Err

    Select Case Err

        Case 47856, 47857, 49740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 47858
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 47859, 47860, 49741, 49742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_PAISES", Err, objPais.iCodigo)

        Case 47861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELA_PAISES", Err, objPais.iCodigo)

        Case 47862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PAISES", Err, objPais.iCodigo)

        Case 47863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 49743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAISES_NOME_DUPLICADO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142398)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Paises_Exclui(ByVal iCodigo As Integer) As Long
'Exclui o Pais do Banco de Dados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim sNome As String
Dim lEndereco As Long
Dim iRegiao As Integer

On Error GoTo Erro_Paises_Exclui

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 47864

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 47865

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 43487
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 49744
    
    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 47866

    sNome = String(STRING_PAISES_NOME, 0)

    'Seleciona os Países que possuem o Código em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Nome FROM Paises WHERE Codigo = ?", 0, sNome, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 47867

    'Lê o primeiro País da seleção feita
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 47868

    'Se não existir País com o código em questão --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 47869

    'Loca o registro que guarda o Código do País
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 47870
    
    'Verifica se o País foi utilizado em alguma Regiao de Venda
    lErro = Comando_Executar(lComando3, "SELECT Codigo FROM RegioesVendas WHERE CodigoPais = ?", iRegiao, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49745

    'Lê o primeiro País da seleção feita
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49746

    'Se encontrou País associado ao Endereço --> Erro
    If lErro = AD_SQL_SUCESSO Then Error 49747
    
    'Verifica se País foi utilizado em algum Endereço
    lErro = Comando_Executar(lComando2, "SELECT Codigo FROM Enderecos WHERE CodigoPais = ?", lEndereco, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 43492

    'Lê o primeiro País da seleção feita
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43493

    'Se encontrou País associado ao Endereço --> Erro
    If lErro = AD_SQL_SUCESSO Then Error 43494
    
    'Exclui o País do Banco de Dados
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Paises", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 47871

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 47872

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Paises_Exclui = SUCESSO

    Exit Function

Erro_Paises_Exclui:

    Paises_Exclui = Err

    Select Case Err

        Case 43487, 47864, 47865, 49744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43492, 43493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", Err)

        Case 43494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_ASSOCIADO_ENDERECO", Err, iCodigo)

        Case 47866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 47867, 47868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_PAISES", Err, iCodigo)

        Case 47869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, iCodigo)

        Case 47870
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CODIGO_PAIS", Err, iCodigo)

        Case 47871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PAISES", Err, iCodigo)

        Case 47872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 49745, 49746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS2", Err)
               
        Case 49747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_ASSOCIADO_REGIOESVENDAS", Err, iCodigo)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142399)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Feriado_Grava(ByVal objFeriado As ClassFeriado) As Long
'Grava Feriado

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sDescricao As String

On Error GoTo Erro_Feriado_Grava

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43384

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 43389

    'Inicializa transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 43385
    
    sDescricao = String(STRING_FERIADO, 0)

    'Tenta ler o Feriado
    lErro = Comando_ExecutarPos(lComando, "SELECT Descricao FROM Feriados WHERE FilialEmpresa = ? AND Data = ?", 0, sDescricao, objFeriado.iFilialEmpresa, objFeriado.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 43386

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 43387
    
    If lErro = AD_SQL_SUCESSO Then
    
        'Se achou verifica se foi mudada a Descrição
        If sDescricao <> objFeriado.sDescricao Then
        
            'Lock da tabela de Feriados
            lErro = Comando_LockExclusive(lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 43388
            
            'Atualiza a Descrição
            lErro = Comando_ExecutarPos(lComando1, "UPDATE Feriados SET Descricao = ?", lComando, objFeriado.sDescricao)
            If lErro <> AD_SQL_SUCESSO Then Error 43390
            
        End If
    
    Else
    
        'Se não encontrou insere um novo registro na tabela de Feriados
        lErro = Comando_Executar(lComando, "INSERT INTO Feriados (FilialEmpresa, Data, Descricao) VALUES (?,?,?)", objFeriado.iFilialEmpresa, objFeriado.dtData, objFeriado.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then Error 43391
    
    End If

    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 43392

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Feriado_Grava = SUCESSO

    Exit Function

Erro_Feriado_Grava:

    Feriado_Grava = Err

    Select Case Err

        Case 43384, 43389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 43386, 43387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err, objFeriado.iFilialEmpresa, objFeriado.dtData)

        Case 43388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FERIADOS", Err, objFeriado.dtData, objFeriado.iFilialEmpresa)

        Case 43390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FERIADOS", Err, objFeriado.dtData, objFeriado.iFilialEmpresa)

        Case 43391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FERIADOS", Err, objFeriado.dtData, objFeriado.iFilialEmpresa)

        Case 43392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142400)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Feriado_Exclui(ByVal objFeriado As ClassFeriado) As Long
'Exclui Feriado

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sDescricao As String

On Error GoTo Erro_Feriado_Exclui
    
    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43397
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 43403
    
    'Inicializa transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 43398
    
    sDescricao = String(STRING_FERIADO, 0)
    
    'Tenta ler o Feriado
    lErro = Comando_ExecutarPos(lComando, "SELECT Descricao FROM Feriados WHERE FilialEmpresa = ? AND Data = ?", 0, sDescricao, objFeriado.iFilialEmpresa, objFeriado.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 43399
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43400
    
    'Se não encontrou o Feriado --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43401
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> SUCESSO Then Error 43402
    
    'Exclui o Feriado do BD
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Feriados", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43404
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 43405
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Feriado_Exclui = SUCESSO
    
    Exit Function
    
Erro_Feriado_Exclui:

    Feriado_Exclui = Err
    
    Select Case Err
        
        Case 43397, 43403
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 43398
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 43399, 43400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FERIADOS", Err, objFeriado.iFilialEmpresa, objFeriado.dtData)
        
        Case 43401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FERIADO_NAO_CADASTRADO", Err, objFeriado.dtData, objFeriado.iFilialEmpresa)
                    
        Case 43402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FERIADOS", Err, objFeriado.dtData, objFeriado.iFilialEmpresa)
                    
        Case 43404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FERIADO", Err, objFeriado.dtData, objFeriado.iFilialEmpresa)
                    
        Case 43405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142401)
            
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function ModuloFilEmp_Atualiza_Configurado(ByVal lCodEmpresa As Long, ByVal iFilialEmpresa As Integer, ByVal colModuloFilEmp As Collection) As Long
'atualiza a flag configurado de todos os modulos (sigla) passados em colMOduloFilEmp para a Empresa e Filial em questão
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim sSiglaModulo As String
Dim objModuloFilEmp As ClassModuloFilEmp
Dim vSiglaModulo As Variant
    
On Error GoTo Erro_ModuloFilEmp_Atualiza_Configurado

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 44949

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 44950

    sSiglaModulo = String(STRING_MODULO_SIGLA, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT SiglaModulo FROM ModuloFilEmp  WHERE CodEmpresa=? AND CodFilial=? AND Configurado = ?", 0, sSiglaModulo, lCodEmpresa, iFilialEmpresa, NAO_CONFIGURADO)
    If lErro <> AD_SQL_SUCESSO Then Error 44951
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44952
    
    Do While lErro = AD_SQL_SUCESSO
    
        For Each vSiglaModulo In colModuloFilEmp
        
            If vSiglaModulo = sSiglaModulo Then
                
                lErro = Comando_ExecutarPos(lComando1, "UPDATE ModuloFilEmp SET Configurado = ?", lComando, CONFIGURADO)
                If lErro <> AD_SQL_SUCESSO Then Error 44953
    
                Exit For
    
            End If
    
        Next
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44954
               
    Loop
               
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ModuloFilEmp_Atualiza_Configurado = SUCESSO
    
    Exit Function
    
Erro_ModuloFilEmp_Atualiza_Configurado:

    ModuloFilEmp_Atualiza_Configurado = Err

    Select Case Err
    
        Case 44949, 44950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 44951, 44952, 44954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP_DIC", Err)
        
        Case 44953
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MODULOFILEMP_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142402)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Relatorio_AtualizarArqReimp(ByVal sCodRel As String, ByVal sNomeArqReimp As String) As Long
'atualiza nome do arquivo de reimpressao no cadastro de relatorios do dicionario de dados

Dim lErro As Long, alComando(1 To 2) As Long, lTransacao As Long, iIndice As Integer
Dim sNomeArqReimpAtual As String

On Error GoTo Erro_Relatorio_AtualizarArqReimp

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 19422

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then Error 19423
        
    Next
        
    sNomeArqReimpAtual = String(255, 0)
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NomeArqReimp FROM Relatorios WHERE CodRel = ?", 0, sNomeArqReimpAtual, sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 19424
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19425
    
    'se nao encontrou registro para o relatorio
    If lErro <> AD_SQL_SUCESSO Then Error 19426
    
    'se trocou o nome do arquivo de reimpressao
    If sNomeArqReimpAtual <> sNomeArqReimp Then
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Relatorios SET NomeArqReimp = ?", alComando(1), sNomeArqReimp)
        If lErro <> AD_SQL_SUCESSO Then Error 19427
        
        'Confirma transação
        lErro = Transacao_CommitDic
        If lErro <> AD_SQL_SUCESSO Then Error 19428
        
    Else
    
        Call Transacao_RollbackDic
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Relatorio_AtualizarArqReimp = SUCESSO
     
    Exit Function
    
Erro_Relatorio_AtualizarArqReimp:

    Relatorio_AtualizarArqReimp = Err
     
    Select Case Err
          
        Case 19422
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 19423
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 19424, 19425
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CADREL", Err)
        
        Case 19426
             lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_NAO_CADASTRADO", Err, sCodRel)
        
        Case 19427
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELATORIO", Err, sCodRel)
        
        Case 19428
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142403)
     
    End Select
     
    Call Transacao_RollbackDic
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

'Function Usuario_Grava_Trans(objUsuario As ClassUsuario) As Long
''inclui ou modifica um registro na tabela Usuario
'
'Dim lErro As Long
'Dim lComando As Long, lComando1 As Long, lTransacao As Long
'Dim iLote As Integer
'
'On Error GoTo Erro_Usuario_Grava_Trans
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 55975
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then Error 55976
'
''   Está impedindo a gravação de Notas Fiscais a Pagar
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then Error 55977
'
'    'verifica se o usuario está cadastrado
'    lErro = Comando_ExecutarPos(lComando, "SELECT Lote FROM Usuario WHERE CodUsuario = ?", 0, iLote, objUsuario.sCodUsuario)
'    If lErro <> AD_SQL_SUCESSO Then Error 55978
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55979
'
'    If lErro = SUCESSO Then
'
'        'existe, então alterar
'        lErro = Comando_ExecutarPos(lComando1, "UPDATE Usuario SET Lote = ?", lComando, objUsuario.iLote)
'        If lErro <> AD_SQL_SUCESSO Then Error 55980
'
'    Else
'
'        'não existe, então incluir
'        lErro = Comando_Executar(lComando, "INSERT INTO Usuario (CodUsuario,Lote) VALUES (?,?)", objUsuario.sCodUsuario, objUsuario.iLote)
'        If lErro <> AD_SQL_SUCESSO Then Error 55981
'
'
'    End If
'
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then Error 55982
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Usuario_Grava_Trans = SUCESSO
'
'    Exit Function
'
'Erro_Usuario_Grava_Trans:
'
'    Usuario_Grava_Trans = Err
'
'    Select Case Err
'
'        Case 55975, 55976
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 55977
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
'
'        Case 55978, 55979
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case 55980
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case 55981
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUARIO", Err, objUsuario.sCodUsuario)
'
'        Case 55982
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142404)
'
'    End Select
'
'    Call Transacao_Rollback
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function

Function BrowseConfiguracoes_Grava(ByVal objBrowse As AdmBrowse, ByVal objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'grava as configurações do browser alteradas pelo usuario

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_BrowseConfiguracoes_Grava

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 9090

    lErro = CF("BrowseUsuarioOrdenacao_Grava", objBrowseUsuarioOrdenacao)
    If lErro <> SUCESSO Then Error 9199

    If objBrowse.iAlterado = 1 Then
        lErro = CF("BrowseUsuarioCampo_Grava", objBrowse.colBrowseUsuarioCampo)
        If lErro <> SUCESSO Then Error 9200
    End If
    
    If objBrowse.iAlteradoOrdenacao = 1 Then
        lErro = BrowseIndiceUsuario_Grava(objBrowse.colBrowseIndiceUsuario, objBrowse.sCodUsuario, objBrowse.objForm.Name)
        If lErro <> SUCESSO Then Error 60829
    End If
    
    lErro = CF("BrowseExcel_Grava", objBrowse, objBrowse.sOpcao, objBrowse.objBrowseExcel)
    If lErro <> SUCESSO Then Error 9199

    'Confirma transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 9093
    
    BrowseConfiguracoes_Grava = SUCESSO
    
    Exit Function

Erro_BrowseConfiguracoes_Grava:

    BrowseConfiguracoes_Grava = Err
    
    Select Case Err
    
        Case 9090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                    
        Case 9093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
         Case 9199, 9200, 60829
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142405)

    End Select

    Call Transacao_RollbackDic
    
    Exit Function

End Function

Private Function BrowseIndiceUsuario_Grava(ByVal colBrowseIndiceUsuario As Collection, ByVal sCodUsuario As String, ByVal sNomeTela As String) As Long
'exclui as informacoes antigas dos indices de ordenacao do usuario e grava as novas
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lComando1 As Long
Dim lErro As Long
Dim sNome As String
    
On Error GoTo Erro_BrowseIndiceUsuario_Grava
    
    lComando1 = 0
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 60833
    
    sNome = String(STRING_NOME_TELA, 0)
    
    'tenta fazer o lock da tela em questão
    lErro = Comando_ExecutarLockado(lComando1, "SELECT Nome FROM Telas WHERE Nome=?", sNome, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 60834
    
    'le a tela em questão
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 60835
    
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 60836
    
    'exclui os indices referentes a tela x usuario
    lErro = BrowseIndiceUsuario_Exclui_BD(sNomeTela, sCodUsuario)
    If lErro <> SUCESSO Then Error 60837

    lErro = BrowseIndiceUsuario_Grava_BD(colBrowseIndiceUsuario, sCodUsuario)
    If lErro <> SUCESSO Then Error 60838
    
    Call Comando_Fechar(lComando1)
    
    BrowseIndiceUsuario_Grava = SUCESSO

    Exit Function

Erro_BrowseIndiceUsuario_Grava:
    
    BrowseIndiceUsuario_Grava = Err
    
    Select Case Err
    
        Case 60833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 60834, 60835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAS", Err)
            
        Case 60836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TELAS", Err)
                   
        Case 60837, 60838
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142406)
            
    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Private Function BrowseIndiceUsuario_Grava_BD(ByVal colBrowseIndiceUsuario As Collection, ByVal sCodUsuario As String) As Long
'grava na tabela BrowseIndiceUsuario os campos contidos na coleção
'IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim objBrowseIndice As AdmBrowseIndice

On Error GoTo Erro_BrowseIndiceUsuario_Grava_BD

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 60839
    
    'Inicia Loop de GRAVACAO
    For Each objBrowseIndice In colBrowseIndiceUsuario
            
        'Faz inserção do campo na Tabela
        lErro = Comando_Executar(lComando, "INSERT INTO BrowseIndiceUsuario (NomeTela,CodUsuario,Indice,NomeIndice,OrdenacaoSQL,SelecaoSQL) VALUES (?,?,?,?,?,?)", objBrowseIndice.sNomeTela, sCodUsuario, objBrowseIndice.iIndice, objBrowseIndice.sNomeIndice, objBrowseIndice.sOrdenacaoSQL, objBrowseIndice.sSelecaoSQL)
        If lErro <> AD_SQL_SUCESSO Then Error 60840
        
    Next

    Call Comando_Fechar(lComando)
    
    BrowseIndiceUsuario_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_BrowseIndiceUsuario_Grava_BD:

    BrowseIndiceUsuario_Grava_BD = Err
    
    Select Case Err
    
        Case 60839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 60840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEINDICEUSUARIO", Err)
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142407)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function BrowseIndiceUsuario_Exclui_BD(ByVal sNomeTela As String, ByVal sCodUsuario As String) As Long
'exclui da tabela de BrowseIndiceUsuario os registros referentes a tela x usuario em questão
'IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_BrowseIndiceUsuario_Exclui_BD

    lComando = 0
    lComando1 = 0

    'Abre comando para pesquisar os campos
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 60841
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 60842
    
    'Pesquisa a existencia dos registros da tela x usuario em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Indice FROM BrowseIndiceUsuario WHERE NomeTela=? AND CodUsuario=?", 0, iIndice, sNomeTela, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 60843
     
    'Lê o primeiro campo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60844

    'Inicia Loop de EXCLUSAO
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Exclui o Registro da  Tabela de BrowseIndiceUsuario
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BrowseIndiceUsuario", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 60845
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60846
    
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    BrowseIndiceUsuario_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_BrowseIndiceUsuario_Exclui_BD:

    BrowseIndiceUsuario_Exclui_BD = Err
    
    Select Case Err
    
        Case 60841, 60842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 60843, 60844, 60846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEINDICEUSUARIO", Err)
           
        Case 60845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEINDICEUSUARIO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142408)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function EdicaoTela_Grava() As Long

Dim lErro As Long
Dim objEdicaoTela_Tela As ClassEdicaoTela_Tela
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sTabIndex1 As String
Dim sTabIndex2 As String
Dim sTabIndex3 As String
Dim sTabIndex4 As String

On Error GoTo Erro_EdicaoTela_Grava

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 64023

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 64024

    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 64025
        
    For Each objEdicaoTela_Tela In gcolEdicaoTela
                       
        For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
           
            lErro = Comando_Executar(lComando1, "INSERT INTO EdicaoTela (NomeTela, NomeControle, Titulo, Visivel, Largura, Altura, Esquerda, Topo, TabStop, Container,IndiceContainer, TabIndex) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", objEdicaoTela_Controle.sNomeTela, objEdicaoTela_Controle.sNomeControle, objEdicaoTela_Controle.sTitulo, objEdicaoTela_Controle.iVisivel, objEdicaoTela_Controle.iLargura, objEdicaoTela_Controle.iAltura, objEdicaoTela_Controle.iEsquerda, objEdicaoTela_Controle.iTopo, objEdicaoTela_Controle.iTabStop, objEdicaoTela_Controle.sContainer, objEdicaoTela_Controle.iIndiceContainer, objEdicaoTela_Controle.iTabIndex)
            If lErro <> AD_SQL_SUCESSO Then Error 64027
        Next
        
        If Len(Trim(objEdicaoTela_Tela.sTabIndex)) > 0 Then
            If Len(Trim(objEdicaoTela_Tela.sTabIndex)) > 255 Then
                sTabIndex1 = left(objEdicaoTela_Tela.sTabIndex, 255)
                sTabIndex2 = Mid(objEdicaoTela_Tela.sTabIndex, 256, 255)
                sTabIndex3 = Mid(objEdicaoTela_Tela.sTabIndex, 511, 255)
                sTabIndex4 = Mid(objEdicaoTela_Tela.sTabIndex, 766, 255)
            Else
                sTabIndex1 = objEdicaoTela_Tela.sTabIndex
            End If
    
            lErro = Comando_Executar(lComando1, "INSERT INTO TabIndex (NomeTela, TabIndex1, TabIndex2, TabIndex3, TabIndex4) VALUES (?,?,?,?,?)", objEdicaoTela_Tela.sNomeTela, sTabIndex1, sTabIndex2, sTabIndex3, sTabIndex4, , objEdicaoTela_Controle.sGrupoUsuarios)
            If lErro <> AD_SQL_SUCESSO Then Error 64033
        End If
        
    Next
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 64028
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    EdicaoTela_Grava = SUCESSO
    
    Exit Function
    
Erro_EdicaoTela_Grava:

    EdicaoTela_Grava = Err
    
    Select Case Err

        Case 64023, 64024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 64026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EDICAOTELA", Err)
        
        Case 64027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EDICAOTELA", Err)
        
        Case 64028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 64032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EDICAOTELA", Err)
        
        Case 64033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EDICAOTELA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142409)

    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

'Inserido por Wagner
Function EdicaoTela_Grava1(ByVal objEdicaoTela_Tela As ClassEdicaoTela_Tela) As Long

Dim lErro As Long
Dim objEdicaoTela_Controle As ClassEdicaoTela_Controle
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sTabIndex1 As String
Dim sTabIndex2 As String
Dim sTabIndex3 As String
Dim sTabIndex4 As String

On Error GoTo Erro_EdicaoTela_Grava1

    If objEdicaoTela_Tela.colEdicaoTela_Controle Is Nothing Then Exit Function

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 64023

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 64024

    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 64025
        
    lErro = Comando_Executar(lComando, "DELETE FROM EdicaoTela WHERE GrupoUsuarios = ? AND NomeTela = ? ", objEdicaoTela_Tela.sGrupoUsuarios, objEdicaoTela_Tela.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 64026

    lErro = Comando_Executar(lComando, "DELETE FROM TabIndex WHERE GrupoUsuarios = ? AND NomeTela = ? ", objEdicaoTela_Tela.sGrupoUsuarios, objEdicaoTela_Tela.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then Error 64032
        
    For Each objEdicaoTela_Controle In objEdicaoTela_Tela.colEdicaoTela_Controle
       
        lErro = Comando_Executar(lComando1, "INSERT INTO EdicaoTela (NomeTela, NomeControle, Titulo, Visivel, Largura, Altura, Esquerda, Topo, TabStop, Container,IndiceContainer, TabIndex, GrupoUsuarios, zOrder, Habilitado) VALUES (?,?,?,?,?,?,?,?,?,?,?,?, ?, ?,?)", objEdicaoTela_Controle.sNomeTela, objEdicaoTela_Controle.sNomeControle, objEdicaoTela_Controle.sTitulo, objEdicaoTela_Controle.iVisivel, objEdicaoTela_Controle.iLargura, objEdicaoTela_Controle.iAltura, objEdicaoTela_Controle.iEsquerda, objEdicaoTela_Controle.iTopo, objEdicaoTela_Controle.iTabStop, objEdicaoTela_Controle.sContainer, objEdicaoTela_Controle.iIndiceContainer, objEdicaoTela_Controle.iTabIndex, objEdicaoTela_Tela.sGrupoUsuarios, objEdicaoTela_Controle.izOrder, objEdicaoTela_Controle.iHabilitado)
        If lErro <> AD_SQL_SUCESSO Then Error 64027
    Next
    
    If Len(Trim(objEdicaoTela_Tela.sTabIndex)) > 0 Then
        If Len(Trim(objEdicaoTela_Tela.sTabIndex)) > 255 Then
            sTabIndex1 = left(objEdicaoTela_Tela.sTabIndex, 255)
            sTabIndex2 = Mid(objEdicaoTela_Tela.sTabIndex, 256, 255)
            sTabIndex3 = Mid(objEdicaoTela_Tela.sTabIndex, 511, 255)
            sTabIndex4 = Mid(objEdicaoTela_Tela.sTabIndex, 766, 255)
        Else
            sTabIndex1 = objEdicaoTela_Tela.sTabIndex
        End If

        lErro = Comando_Executar(lComando1, "INSERT INTO TabIndex (NomeTela, TabIndex1, TabIndex2, TabIndex3, TabIndex4, GrupoUsuarios) VALUES (?,?,?,?,?,?)", objEdicaoTela_Tela.sNomeTela, sTabIndex1, sTabIndex2, sTabIndex3, sTabIndex4, objEdicaoTela_Tela.sGrupoUsuarios)
        If lErro <> AD_SQL_SUCESSO Then Error 64033
    End If
        
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 64028
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    EdicaoTela_Grava1 = SUCESSO
    
    Exit Function
    
Erro_EdicaoTela_Grava1:

    EdicaoTela_Grava1 = Err
    
    Select Case Err

        Case 64023, 64024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 64026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EDICAOTELA", Err)
        
        Case 64027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EDICAOTELA", Err)
        
        Case 64028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 64032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EDICAOTELA", Err)
        
        Case 64033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EDICAOTELA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142410)

    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

'Inserido por Wagner
Function EdicaoTela_Exclui(ByVal objEdicaoTela_Tela As ClassEdicaoTela_Tela, colGrupoUsu As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim objGrupoUsu As ClassGrupoUsuarios

On Error GoTo Erro_EdicaoTela_Exclui

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 64023

    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 64025
        
    For Each objGrupoUsu In colGrupoUsu
        
        lErro = Comando_Executar(lComando, "DELETE FROM EdicaoTela WHERE NomeTela = ? AND GrupoUsuarios = ? ", objEdicaoTela_Tela.sNomeTela, objGrupoUsu.sCodGrupo)
        If lErro <> AD_SQL_SUCESSO Then Error 64026
    
        lErro = Comando_Executar(lComando, "DELETE FROM TabIndex WHERE NomeTela = ? AND GrupoUsuarios = ? ", objEdicaoTela_Tela.sNomeTela, objGrupoUsu.sCodGrupo)
        If lErro <> AD_SQL_SUCESSO Then Error 64032
       
    Next
       
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 64028
    
    lErro = Comando_Fechar(lComando)
    
    EdicaoTela_Exclui = SUCESSO
    
    Exit Function
    
Erro_EdicaoTela_Exclui:

    EdicaoTela_Exclui = Err
    
    Select Case Err

        Case 64023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 64026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EDICAOTELA", Err)
                
        Case 64028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 64032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EDICAOTELA", Err)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142411)

    End Select
    
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelOpcoes_Grava_Parametros(ByVal sCodRel As String, ByVal sNome As String, ByVal colParametros As Collection) As Long
'Lê os parametros do Relatório para Opcão
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim objParametro As New AdmRelParametro
Dim sOpcao As String
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_RelOpcoes_Grava_Parametros
    
    'Abrir comando com o Dicionário de Dados
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 61259
    
    'Abrir comando com o Dicionário de Dados
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then Error 61260
    
    sOpcao = String(STRING_CODIGO_OPCAO_RELATORIO, 0)
    
    'verifica se a opção já existe em RelatorioParametro
    lErro = Comando_ExecutarPos(lComando, "SELECT Nome FROM RelatorioParametros WHERE Nome = ? AND CodRel = ?", 0, sOpcao, sNome, sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 61261
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61262
    
    'Se encontrou o registro
    If lErro = AD_SQL_SUCESSO Then
    
        'Faz Lock nos Registros
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61263
        
        'Se existir exclui todas as opções
        Do While lErro = AD_SQL_SUCESSO
            
            lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RelatorioParametros", lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 61264
    
            'Busca o primeiro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61265
    
        Loop
        
    End If
    
    'Inclui todas as Opções
    For Each objParametro In colParametros
                
        'não existe, então incluir
        lErro = Comando_Executar(lComando1, "INSERT INTO RelatorioParametros (CodRel, Nome,Parametro,ParamValor) VALUES (?,?,?,?)", sCodRel, sNome, objParametro.sParametro, objParametro.sParamValor)
        If lErro <> AD_SQL_SUCESSO Then Error 61266
    
    Next
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    RelOpcoes_Grava_Parametros = SUCESSO
    
    Exit Function
    
Erro_RelOpcoes_Grava_Parametros:

    RelOpcoes_Grava_Parametros = Err
    
    Select Case Err

        Case 61259, 61260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 61261, 61262, 61265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIOPARAMETROS", Err)
                        
        Case 61263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELATORIOPARAMETROS", Err)
                        
        Case 61264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELATORIOPARAMETROS", Err)
                        
        Case 61266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELATORIOPARAMETROS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142412)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
        
End Function

Public Function ObjetosBD_Atualiza(ByVal objObjetosBD As ClassObjetoBD) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long, lTransacao As Long
Dim iAvisaSobrePosicao As Integer

On Error GoTo Erro_ObjetosBD_Atualiza

    'Abertura de comandos
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 80436
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 80437
           
    'Abre transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 80453
    
    'Faz a leitura da tabela
    lErro = Comando_ExecutarPos(lComando, "SELECT AvisaSobreposicao FROM ObjetosBD WHERE ClasseObjeto = ?", 0, iAvisaSobrePosicao, objObjetosBD.sClasseObjeto)
    If lErro <> AD_SQL_SUCESSO Then gError 80438

    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80439
    
    'Faz Atualização na tabela
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ObjetosBD SET AvisaSobrePosicao = ?", lComando, objObjetosBD.iAvisaSobrePosicao)
    If lErro <> AD_SQL_SUCESSO Then gError 80440

    Call Transacao_CommitDic
    
    'Fecha comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando)

    ObjetosBD_Atualiza = SUCESSO

    Exit Function
    
Erro_ObjetosBD_Atualiza:

    ObjetosBD_Atualiza = gErr

    Select Case gErr
    
        Case 80436, 80437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)

        Case 80438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBJETOBD", gErr, Error)
            
        Case 80439
            'Erro tratado na rotina chamadora
        
        Case 80440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_OBJETOSBD", gErr, Error, objObjetosBD.iAvisaSobrePosicao)

        Case 80453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr, Error)
    
        Case Else
            lErro = Rotina_Erro(ERRO_FORNECIDO_PELO_VB, "gErr", Error, 142413)
    
    End Select
    
    'Fecha comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando)
    
    Call Transacao_RollbackDic
    
    Exit Function
    
End Function

Function Config_ObterNumInt_Trans(ByVal sTabelaConfig As String, ByVal sCodigo As String, lNumInt As Long, Optional ByVal iQtde As Integer = 1, Optional ByVal iFilialEmpresa As Integer = EMPRESA_TODA) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado

Dim lErro As Long
Dim lComandoSel As Long, lComandoMod As Long
Dim sConteudo As String
Dim lTransacao As Long

On Error GoTo Erro_Config_ObterNumInt_Trans

    If iQtde < 0 Then gError 106589
    
    sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92994

    lComandoSel = Comando_Abrir()
    If lComandoSel = 0 Then gError 92995

    lComandoMod = Comando_Abrir()
    If lComandoMod = 0 Then gError 92996

    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 92997

    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92998

    If lErro = AD_SQL_SEM_DADOS Then gError 92999
    
    lErro = Comando_LockExclusive(lComandoSel)
    If lErro <> AD_SQL_SUCESSO Then gError 105000
    
    lNumInt = CLng(sConteudo)
    sConteudo = CStr(lNumInt + iQtde)

    lErro = Comando_ExecutarPos(lComandoMod, "UPDATE " & sTabelaConfig & " SET Conteudo = ?", lComandoSel, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then gError 105001

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 105002

    lErro = Comando_Fechar(lComandoSel)
    lErro = Comando_Fechar(lComandoMod)

    Config_ObterNumInt_Trans = SUCESSO

    Exit Function

Erro_Config_ObterNumInt_Trans:

    Config_ObterNumInt_Trans = gErr

    Select Case gErr

        Case 92994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92995, 92996
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92997, 92998
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", gErr, sTabelaConfig, sCodigo)

        Case 92999
            Call Rotina_Erro(vbOKOnly, "ERRO_CONFIG_NAO_CADASTRADO", gErr, sCodigo, sTabelaConfig)
            
        Case 105000
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", gErr, sTabelaConfig, sCodigo)
        
        Case 105001
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CPRCONFIG", gErr, sCodigo)

        Case 105002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 106589
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_INVALIDA_NUMINT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142414)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoMod)

    Exit Function

End Function

Function Config_Instalacao_Filial(ByVal iFilialEmpresa As Integer, ByVal colModuloFilEmp As Collection) As Long
'pesquisa nas tabelas config (MATConfig, ESTConfig, LojaConfig, etc) por registros que tenham o campo PorFilial = 1
'isto indicará que deverá ser gerado uma cópia do registro por filial com o conteudo do campo ConteudoInicial

Dim lErro As Long
Dim sCodigo As String
Dim sDescricao As String
Dim iTipo As Integer
Dim alComando(3) As Long
Dim sConteudoInicial As String
Dim iIndice As Integer
Dim sNomeTabela As String
Dim vSiglaModulo As Variant
Dim lTransacao As Long

On Error GoTo Erro_Config_Instalacao_Filial

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 110027

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 110019
    Next
        
    If giFilialEmpresa <> 1 Then
    
        For Each vSiglaModulo In colModuloFilEmp
            
            sNomeTabela = String(STRING_NOMETABELA, 0)
        
            'Lê em TabelasConfig as tabelas que deverão ser pesquisadas por modulo para verificar se tem registros com o campo PorFilial = POR_FILIAL
            lErro = Comando_Executar(alComando(0), "SELECT NomeTabela FROM TabelasConfig WHERE SiglaModulo = ?", sNomeTabela, vSiglaModulo)
            If lErro <> AD_SQL_SUCESSO Then gError 110020
        
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110021
        
            Do While lErro = AD_SQL_SUCESSO
        
                sCodigo = String(STRING_CONFIG_CODIGO, 0)
                sDescricao = String(STRING_CONFIG_DESCRICAO, 0)
                sConteudoInicial = String(STRING_CONTEUDO, 0)
        
                'Lê em MatConfig os registros que devem ser inicializados porfilial
                lErro = Comando_Executar(alComando(1), "SELECT Codigo, FilialEmpresa,Descricao,Tipo, ConteudoInicial FROM " & sNomeTabela & " T1 WHERE PorFilial = 1 AND FilialEmpresa = 1 AND NOT EXISTS ( SELECT * FROM " & sNomeTabela & " T2 WHERE FilialEmpresa = ? AND T2.Codigo = T1.Codigo)", sCodigo, iFilialEmpresa, sDescricao, iTipo, sConteudoInicial, giFilialEmpresa)
                If lErro <> AD_SQL_SUCESSO Then gError 110022
        
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110023
        
                Do While lErro = AD_SQL_SUCESSO
                
                    'Inclui registro referente a filial no config selecionado
                    lErro = Comando_Executar(alComando(2), "INSERT INTO " & sNomeTabela & " (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo,ConteudoInicial, PorFilial) VALUES (?,?,?,?,?,?,?)", sCodigo, giFilialEmpresa, sDescricao, iTipo, sConteudoInicial, sConteudoInicial, 1)
                    If lErro <> AD_SQL_SUCESSO Then gError 110024
                
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110025
                
                Loop
                
                lErro = Comando_BuscarProximo(alComando(0))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110026
        
            Loop
    
        Next
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 110028

    Config_Instalacao_Filial = SUCESSO
     
    Exit Function
    
Erro_Config_Instalacao_Filial:

    Config_Instalacao_Filial = gErr
     
    Select Case gErr
          
        Case 110019
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 110020, 110021, 110026
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASCONFIG", gErr)

        Case 110022, 110023, 110025
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQCONFIG", gErr, sNomeTabela)
        
        Case 110024
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQCONFIG1", gErr, sNomeTabela)
            
        Case 110027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 110028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142415)

    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
            
    Exit Function

End Function

Function TelasOpcao_Grava(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long
'Função responsável por fazer a gravação de uma opção em TelasOpcoes e TelasOpcoesValores

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim sNomeTela As String

On Error GoTo Erro_TelasOpcao_Grava
    
    If Len(Trim(objTelasOpcoes.objComboOpcoes.Text)) = 0 Then gError 103284
    
    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103230
    Next

    'Abre a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 103231
        
    objTelasOpcoes.sNomeOpcao = objTelasOpcoes.objComboOpcoes.Text
        
    sNomeTela = String(STRING_TELASCAMPOS_NOMETELA, 0)

    'verifica se a opção já existe em TelasOpcoes
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NomeTela FROM TelasOpcoes WHERE NomeTela=? AND TituloTela=? AND NomeOpcao=?", 0, sNomeTela, objTelasOpcoes.sNomeTela, objTelasOpcoes.sTituloTela, objTelasOpcoes.sNomeOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 103232

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103233

    'se a opção ainda não existir em TelasOpcoes
    If lErro <> SUCESSO Then

        lErro = Comando_Executar(alComando(1), "INSERT INTO TelasOpcoes(NomeTela, TituloTela, NomeOpcao) VALUES (?,?,?)", objTelasOpcoes.sNomeTela, objTelasOpcoes.sTituloTela, objTelasOpcoes.sNomeOpcao)
        If lErro <> AD_SQL_SUCESSO Then gError 103234
        
        objTelasOpcoes.objComboOpcoes.AddItem objTelasOpcoes.sNomeOpcao
        
    'se a opção já existir em TelasOpcoes
    Else

        'Exclui em TelasOpcoesValores a opção em questão
        lErro = CF("TelasOpcoesValores_Exclui_EmTrans", objTelasOpcoes)
        If lErro <> SUCESSO Then gError 103235

    End If

    'Grava o Layout da tela para a opção em questão
    lErro = CF("TelasOpcoesValores_Grava_EmTrans", objTelasOpcoes)
    If lErro <> SUCESSO Then gError 103236

    objTelasOpcoes.objTela.BotaoLimpar_Click

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 103237

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TelasOpcao_Grava = SUCESSO

    Exit Function

Erro_TelasOpcao_Grava:

   TelasOpcao_Grava = gErr

   Select Case gErr

        Case 103235, 103236

        Case 103230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103232, 103233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELASOPCOES", gErr)

        Case 103234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TELASOPCOES", gErr, objTelasOpcoes.sNomeOpcao)

        Case 103237
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 103284
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCAO_NAO_PREENCHIDA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142416)

   End Select

   'Call Transacao_Rollback
   Call Transacao_RollbackDic

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TelasOpcoesValores_Grava_EmTrans(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long
'Grava em TelasOpcoesValores o layout da tela

Dim lErro As Long
Dim objControle As Object
Dim objTelasCampos As ClassTelasCampos
Dim colTelasCampos As New Collection
Dim sConteudo As String
Dim alComando(0) As Long
Dim iIndice As Integer

On Error GoTo Erro_TelasOpcoesValores_Grava_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103238
    Next

    lErro = CF("TelasCampos_Le_Todos", colTelasCampos, objTelasOpcoes)
    If lErro <> SUCESSO Then gError 103239
    
    If colTelasCampos.Count = 0 Then gError 103250
    
    For Each objTelasCampos In colTelasCampos

        Set objControle = objTelasOpcoes.objTela.Controls(objTelasCampos.sNomeCampo)

        Select Case TypeName(objControle)

            Case "MaskEdBox", "TextBox", "ComboBox"

                sConteudo = objControle.Text

            Case "CheckBox"

                If objControle.Value <> DESMARCADO Then

                    sConteudo = objControle.Value

                End If

            Case "OptionButton"

                If objControle.Value <> False Then

                    sConteudo = True

                End If

        End Select

        If Len(Trim(sConteudo)) <> 0 Then

            lErro = Comando_Executar(alComando(0), "INSERT INTO TelasOpcoesValores(TituloTela, NomeTela, NomeOpcao, NomeCampo, ValorCampo)VALUES(?,?,?,?,?)", _
            objTelasOpcoes.sTituloTela, objTelasOpcoes.sNomeTela, objTelasOpcoes.sNomeOpcao, objTelasCampos.sNomeCampo, sConteudo)
            If lErro <> AD_SQL_SUCESSO Then gError 103240
            
            sConteudo = ""
            
        End If

    Next

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   TelasOpcoesValores_Grava_EmTrans = SUCESSO

   Exit Function

Erro_TelasOpcoesValores_Grava_EmTrans:

   TelasOpcoesValores_Grava_EmTrans = gErr

    Select Case gErr

        Case 103238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103239

        Case 103240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELASCAMPOS_NAO_CADASTRADA", gErr, objTelasCampos.sTituloTela)
        
        Case 103250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TELASOPCOESVALORES", gErr, objTelasOpcoes.sNomeOpcao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142417)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Opcao_Exclui(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long
'Inicia o processo de exclusão da opção

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer

On Error GoTo Erro_Opcao_Exclui

    If Len(Trim(objTelasOpcoes.objComboOpcoes.Text)) = 0 Then gError 103285
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 103253
    
    objTelasOpcoes.sNomeOpcao = objTelasOpcoes.objComboOpcoes.Text
    
    'Exclui os registros referentes em telasopcoes
    lErro = CF("TelasOpcoes_Exclui_EmTrans", objTelasOpcoes)
    If lErro <> SUCESSO Then gError 103254

    'Exclui os registros referentes em telasopcoesvalores
    lErro = CF("TelasOpcoesValores_Exclui_EmTrans", objTelasOpcoes)
    If lErro <> SUCESSO Then gError 103255
    
    For iIndice = 0 To objTelasOpcoes.objComboOpcoes.ListCount - 1
        
        If objTelasOpcoes.objComboOpcoes.List(iIndice) = objTelasOpcoes.sNomeOpcao Then
        
            objTelasOpcoes.objComboOpcoes.RemoveItem iIndice
            Exit For
            
        End If
        
    Next
    
    Call objTelasOpcoes.objTela.BotaoLimpar_Click
    
    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 103256

    Opcao_Exclui = SUCESSO

    Exit Function

Erro_Opcao_Exclui:

    Opcao_Exclui = gErr

    Select Case gErr

        Case 103253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103254, 103255

        Case 103256
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 103285
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCAO_NAO_PREENCHIDA", gErr)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142418)

    End Select

    'Call Transacao_Rollback
    Call Transacao_RollbackDic

    Exit Function

End Function

Function TelasOpcoes_Exclui_EmTrans(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim sNomeTela As String

On Error GoTo Erro_TelasOpcoes_Exclui_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103278
    Next

    sNomeTela = String(STRING_TELASCAMPOS_NOMETELA, 0)

    lErro = Comando_ExecutarPos(alComando(0), "SELECT NomeTela FROM TelasOpcoes WHERE NomeTela = ? AND TituloTela = ? AND NomeOpcao=?", _
    0, sNomeTela, objTelasOpcoes.sNomeTela, objTelasOpcoes.sTituloTela, objTelasOpcoes.sNomeOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 103257

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103258

    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TelasOpcoes", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 103259

    End If

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   TelasOpcoes_Exclui_EmTrans = SUCESSO

   Exit Function

Erro_TelasOpcoes_Exclui_EmTrans:

   TelasOpcoes_Exclui_EmTrans = gErr

    Select Case gErr

        Case 103278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103257, 103258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELASOPCOES", gErr)

        Case 103259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TELASOPCOES", gErr, objTelasOpcoes.sNomeOpcao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142419)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TelasOpcoesValores_Exclui_EmTrans(ByVal objTelasOpcoes As ClassTelasOpcoes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim sNomeCampo As String

On Error GoTo Erro_TelasOpcoesValores_Exclui_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 103279
    Next

    sNomeCampo = String(STRING_TELASCAMPOS_NOMECAMPO, 0)

    lErro = Comando_ExecutarPos(alComando(0), "SELECT NomeCampo FROM TelasOpcoesValores WHERE NomeTela = ? AND TituloTela = ? AND NomeOpcao=?", _
    0, sNomeCampo, objTelasOpcoes.sNomeTela, objTelasOpcoes.sTituloTela, objTelasOpcoes.sNomeOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 103280

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103281

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TelasOpcoesValores", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 103282

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103283

    Loop

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   TelasOpcoesValores_Exclui_EmTrans = SUCESSO

   Exit Function

Erro_TelasOpcoesValores_Exclui_EmTrans:

   TelasOpcoesValores_Exclui_EmTrans = gErr

    Select Case gErr

        Case 103279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103280, 103281, 103283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELASOPCOESVALORES", gErr)

        Case 103282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TELASOPCOESVALORES", gErr, objTelasOpcoes.sNomeOpcao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142420)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CotacaoMoeda_Exclui(ByVal objCotacaoMoeda As ClassCotacaoMoeda) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim dValor As Double

On Error GoTo Erro_CotacaoMoeda_Exclui

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80252
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 80253

    'Abre Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 80254

    'Faz Selecao do registro a ser excluido
    lErro = Comando_ExecutarPos(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", 0, dValor, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 80255

    'Tenta encontrar o registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80256

    'Se o registro não existe erro
    If lErro = AD_SQL_SEM_DADOS Then gError 80257

    'Faz lock
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 80258
    
    'Executa exclusão
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM CotacoesMoeda", lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 80259

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 80260

    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    CotacaoMoeda_Exclui = SUCESSO
    
    Exit Function
    
Erro_CotacaoMoeda_Exclui:

    CotacaoMoeda_Exclui = gErr
    
    Select Case gErr
    
        Case 80252, 80253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 80254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 80255, 80256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
            
        Case 80257 'Erro a Cotação não existe
            
        Case 80258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COTACOESMOEDA", gErr)
            
        Case 80259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)

        Case 80260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142421)
            
    End Select
    
    Call Transacao_Rollback
    
    'Fecha comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando)
    
End Function

Public Function CotacaoMoeda_Grava(ByVal objCotacaoMoeda As ClassCotacaoMoeda) As Long
'Função de Gravação

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim dValor As Double

On Error GoTo Erro_CotacaoMoeda_Grava

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80241

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 80242

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 80243

    '???? MELHORAR Comentário. É claro que um Select faz seleção - Ok
    'Faz seleção do campo valor passando data e moeda como paramêtros
    lErro = Comando_ExecutarPos(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", 0, dValor, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 80244

    'Tenta encontrar o registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80245

    
    If lErro = AD_SQL_SUCESSO Then
    '====> Encontrou registro, é uma atualização
    
        lErro = Comando_ExecutarPos(lComando1, "UPDATE CotacoesMoeda SET Valor = ? ", lComando, objCotacaoMoeda.dValor)
        If lErro <> AD_SQL_SUCESSO Then gError 80246
    
    Else
    '====> Não encontrou registro, é uma inclusão
        
        lErro = Comando_Executar(lComando1, "INSERT INTO CotacoesMoeda (Data, Valor, Moeda) VALUES (?, ?, ?)", objCotacaoMoeda.dtData, objCotacaoMoeda.dValor, objCotacaoMoeda.iMoeda)
        If lErro <> AD_SQL_SUCESSO Then gError 80247

    End If
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 80248
    
    'Fecha Comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando)
    
    CotacaoMoeda_Grava = SUCESSO
    
    Exit Function
    
Erro_CotacaoMoeda_Grava:

    CotacaoMoeda_Grava = gErr
    
    Select Case gErr
    
        Case 80241, 80242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 80243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 80244, 80245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr, objCotacaoMoeda.dValor)
        
        Case 80246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.dValor)
        
        Case 80247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.dValor)
            
        Case 80248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142422)
    
    End Select
    
    Call Transacao_Rollback
    
    'Fecha Comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando)
    
End Function

Public Function Moedas_Exclui(objMoedas As ClassMoedas) As Long

Dim lErro As Long
Dim alComando(1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Moedas_Exclui

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 108826
    Next
    
    'Abre Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 108827

    'Prepara a busca do Registro a ser excluído
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Moedas WHERE Codigo = ?", 0, iCodigo, objMoedas.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108828

    'Busca o Registro a ser excluído
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108829

    'Se nao Encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108830

    'Faz lock
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 108831
    
    'Verifica a existencia (uso) da moeda ...
    lErro = CF("Moeda_Critica_Exclusao", objMoedas)
    If lErro <> SUCESSO Then gError 108814
    
    'Executa exclusão
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Moedas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 108832

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 108833

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Moedas_Exclui = SUCESSO
    
    Exit Function
    
Erro_Moedas_Exclui:

    Moedas_Exclui = gErr
    
    Select Case gErr
    
        Case 108814
    
        Case 108826
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 108827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 108828, 108829
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)
            
        Case 108830
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_INEXISTENTE", gErr, objMoedas.iCodigo)
            
        Case 108831
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOEDAS", gErr)
            
        Case 108832
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MOEDAS", gErr, objMoedas.iCodigo)

        Case 108833
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142423)
            
    End Select
    
    Call Transacao_Rollback
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Public Function Moedas_Grava(objMoedas As ClassMoedas) As Long
'Função de Gravação das Moedas (Indexadores)

Dim lErro As Long
Dim alComando(1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Moedas_Grava

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 108806
    Next
    
    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 108807
    
    'Verifica a existencia da moeda com mesmo nome
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Moedas WHERE Nome=? And Codigo <> ?", 0, iCodigo, objMoedas.sNome, objMoedas.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108808

    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108809
    'Se encontrou => Erro
    If lErro = AD_SQL_SUCESSO Then gError 108842
    
    'Verifica a existencia da moeda
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Moedas WHERE Codigo=?", 0, iCodigo, objMoedas.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108808

    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108809

    'Se Encontrou registro => é uma atualização
    If lErro = AD_SQL_SUCESSO Then
    
        '##################################
        'Alterado por Wagner
        'Inserido ValoresEmPercentual
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Moedas SET Nome=?, Simbolo=?, ValoresEmPercentual = ? ", alComando(0), objMoedas.sNome, objMoedas.sSimbolo, objMoedas.iPercentual)
        If lErro <> AD_SQL_SUCESSO Then gError 108810
        '##################################
    
    'Não encontrou registro => é uma inclusão
    Else
        
        '##################################
        'Alterado por Wagner
        'Inserido ValoresEmPercentual
        lErro = Comando_Executar(alComando(1), "INSERT INTO Moedas (Codigo, Nome, Simbolo, ValoresEmPercentual) VALUES (?, ?, ?, ?)", objMoedas.iCodigo, objMoedas.sNome, objMoedas.sSimbolo, objMoedas.iPercentual)
        If lErro <> AD_SQL_SUCESSO Then gError 108811
        '##################################

    End If
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 108812
    
    'Fecha Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Moedas_Grava = SUCESSO
    
    Exit Function
    
Erro_Moedas_Grava:

    Moedas_Grava = gErr
    
    Select Case gErr
    
        Case 108806
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 108807
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 108808, 108809
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)
        
        Case 108810
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOEDAS", gErr, objMoedas.iCodigo)
        
        Case 108811
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MOEDAS", gErr)
            
        Case 108812
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 108842
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NOME_EXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142424)
    
    End Select
    
    Call Transacao_Rollback
    
    'Fecha Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Public Function Anotacoes_Exclui(ByVal objAnotacoes As ClassAnotacoes) As Long
'exclui os registros da tabela anotacoes através dos dados passados como parametro

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_Anotacoes_Exclui
    
    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 119825

    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 119826
    Next

    'procura o numintdoc em anotacoes a partir dos dados passados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Anotacoes WHERE Origem = ? AND ID = ? AND Titulo = ? AND DataAlteracao = ?", 0, lNumIntDoc, objAnotacoes.iTipoDocOrigem, objAnotacoes.sID, objAnotacoes.sTitulo, objAnotacoes.dtDataAlteracao)
    If lErro <> AD_SQL_SUCESSO Then gError 119827

    'busca o 1º registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119828

    'sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119829

    'faz um lock exclusivo do registro
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 119830

    'Exclui o registro encontrado na tabela anotacoes
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Anotacoes", alComando(0))
    If lErro <> SUCESSO Then gError 119831
                
    'exclui o registro em anotacoeslinha a partir do NumIntDoc obtido
    lErro = CF("AnotacoesLinha_Exclui_Em_Trans", lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 119837 Then gError 119832
        
    'não achou o texto referente a anotacao
    If lErro = 119837 Then gError 119860
        
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 119833

    Anotacoes_Exclui = SUCESSO

    Exit Function

Erro_Anotacoes_Exclui:

    Anotacoes_Exclui = gErr

    Select Case gErr
    
        Case 119832
    
        Case 119825
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_ABRIR", gErr)
    
        Case 119826
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)
    
        Case 119827, 119828
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOES", gErr)
    
        Case 119829
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOTACAO_NAO_ENCONTRADA", gErr, objAnotacoes.sID)
    
        Case 119830
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXCLUSIVE", gErr)
            
        Case 119831
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ANOTACAO", gErr)
       
        Case 119860
            Call Rotina_Erro(vbOKOnly, "ERRO_ANOTACOESLINHA_NAO_ENCONTRADA", gErr, objAnotacoes.sID)
       
        Case 119833
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142425)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Anotacoes_Grava(ByVal objAnotacoes As ClassAnotacoes) As Long
'grava a anotação nas tabelas anotacoes e anotacoeslinha

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim lIndiceTexto As Long
Dim objAnotacoesLinha As ClassAnotacoesLinha
Dim iSeq As Integer, sTextoAux As String

On Error GoTo Erro_Anotacoes_Grava
    
    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 119813

    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 119814
    Next
                
    'busca os registros da tabela anotacoes a partir dos demais dados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Anotacoes WHERE DataAlteracao = ? AND Origem = ? AND ID = ? AND Titulo = ?", 0, lNumIntDoc, objAnotacoes.dtDataAlteracao, objAnotacoes.iTipoDocOrigem, objAnotacoes.sID, objAnotacoes.sTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 119815

    'busca o 1ºregistro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119816
        
    'se não encontrou registros
    If lErro = AD_SQL_SEM_DADOS Then
    
        'obtem o num. do documento
        lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_ANOTACAO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 119859
    
        'coloca em objanotacoes o novo numintdoc
        objAnotacoes.lNumIntDoc = lNumIntDoc
    
        'insere os novos registros na tabela anotacoes
        lErro = Comando_Executar(alComando(1), "INSERT INTO Anotacoes(NumIntDoc, Origem, ID, Titulo, DataAlteracao) VALUES (?,?,?,?,?)", _
        objAnotacoes.lNumIntDoc, objAnotacoes.iTipoDocOrigem, objAnotacoes.sID, objAnotacoes.sTitulo, gdtDataHoje)
        If lErro <> AD_SQL_SUCESSO Then gError 119817
         
    'senão
    Else

        'faz um lock exclusivo p/ garantir o update
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 119928
    
        'atualiza a data do registro em anotacoes
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Anotacoes SET DataAlteracao = ?", alComando(0), gdtDataHoje)
        If lErro <> SUCESSO Then gError 119819
            
        'exclui o texto antigo na tabela anotacoeslinha a partir do NumIntDoc p/ ser gravado depois
        lErro = CF("AnotacoesLinha_Exclui_Em_Trans", lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 119837 Then gError 119820

    End If
    
    sTextoAux = objAnotacoes.sTextoCompleto

    'inicia a etapa de gravação em anotacoeslinha
    'de 1 até o total de caracteres guardados em objanotacoes.sTextoCompleto
    'For lIndiceTexto = 1 To Len(objAnotacoes.sTextoCompleto)
    Do While Len(Trim(sTextoAux)) > 0
        
        'instancia o obj
        Set objAnotacoesLinha = New ClassAnotacoesLinha
        
        'guarda o numintdoc referente
        objAnotacoesLinha.lNumIntDoc = lNumIntDoc
            
       'pega a partir da posição Indice, 255 caracteres
        'objAnotacoesLinha.sTexto = Mid(objAnotacoes.sTextoCompleto, lIndiceTexto, 255)
        objAnotacoesLinha.sTexto = left(sTextoAux, 255)
        
        'atualiza a seq.
        iSeq = iSeq + 1

        'diz qual é a sequencia p/ poder "remontar" o texto novamente (ordenar qnd ler)
        objAnotacoesLinha.iSeq = iSeq

        'atualiza o Indice c/ + 255 posições
        'lIndiceTexto = lIndiceTexto + 255
        
        If Len(sTextoAux) > 255 Then
            sTextoAux = Mid(sTextoAux, 256)
        Else
            sTextoAux = ""
        End If

        'grava a linha c/ um "pedaço" do texto em AnotacoesLinha
        lErro = CF("AnotacoesLinha_Grava_Em_Trans", objAnotacoesLinha)
        If lErro <> SUCESSO Then gError 119821

    Loop
    'Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 119822

    Anotacoes_Grava = SUCESSO

    Exit Function

Erro_Anotacoes_Grava:

    Anotacoes_Grava = gErr

    Select Case gErr
        
        Case 119818, 119820, 119821, 119859
        
        Case 119813
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_ABRIR", gErr)
        
        Case 119814
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)
                    
        Case 119815, 119816
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOES", gErr)
                    
        Case 119928
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXCLUSIVE", gErr)
                    
        Case 119817
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ANOTACOES", gErr)
        
        Case 119819
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ANOTACOES", gErr)
        
        Case 119822
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142426)
                
    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
        
    Exit Function

End Function

Public Function AnotacoesLinha_Exclui_Em_Trans(lNumIntDoc As Long) As Long
'exclui os "pedaços" do texto em anotacoeslinha
'apenas em TRANSACAO ABERTA

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iSeq As Integer

On Error GoTo Erro_AnotacoesLinha_Exclui_Em_Trans

    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 119834
    Next
        
    'procura o registro em anotacoesliha a partir do NumIntDoc passado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Seq FROM AnotacoesLinha WHERE NumIntDoc = ?", 0, iSeq, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 119835

    'busca o 1º registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119836

    'sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119837

    'enquanto houver "pedaços" de textos referentes ao NumIntRel
    Do While lErro = SUCESSO
    
        'faz um lock exclusivo do registro
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 119838
    
        'Exclui o registro encontrado
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM AnotacoesLinha", alComando(0))
        If lErro <> SUCESSO Then gError 119839
                    
        'busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119840
        
    Loop
        
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    AnotacoesLinha_Exclui_Em_Trans = SUCESSO

    Exit Function

Erro_AnotacoesLinha_Exclui_Em_Trans:

    AnotacoesLinha_Exclui_Em_Trans = gErr

    Select Case gErr
    
        Case 119834
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)
    
        Case 119835, 119836, 119840
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANOTACOESLINHA", gErr)
    
        Case 119837 'sem dados
    
        Case 119838
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXCLUSIVE", gErr)
            
        Case 119839
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ANOTACOESLINHA", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142427)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function AnotacoesLinha_Grava_Em_Trans(ByVal objAnotacoesLinha As ClassAnotacoesLinha) As Long '(ByVal lNumIntDoc As Long, colAnotacoesLinha As Collection) As Long
'grava o texto "desmembrado" em anotacoeslinha
'apenas em TRANSACAO ABERTA

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_AnotacoesLinha_Grava_Em_Trans

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 119823
        
    'insere o novo registro na tabela anotacoeslinha (que foi passado como parametro)
    lErro = Comando_Executar(lComando, "INSERT INTO AnotacoesLinha(NumIntDoc, Seq, Texto) VALUES (?,?,?)", objAnotacoesLinha.lNumIntDoc, objAnotacoesLinha.iSeq, objAnotacoesLinha.sTexto)
    If lErro <> AD_SQL_SUCESSO Then gError 119824
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    AnotacoesLinha_Grava_Em_Trans = SUCESSO

    Exit Function

Erro_AnotacoesLinha_Grava_Em_Trans:

    AnotacoesLinha_Grava_Em_Trans = gErr

    Select Case gErr
        
        Case 119823
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 119824
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ANOTACOESLINHA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142428)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Cidade_Grava(ByVal objCidades As ClassCidades) As Long
'Realiza a gravação do Cadastro no BD

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Cidade_Grava

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 125022

    lErro = CF("Cidade_Grava_EmTrans", objCidades)
    If lErro <> SUCESSO Then gError 130009
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125029
    
    Cidade_Grava = SUCESSO
    
    Exit Function
    
Erro_Cidade_Grava:

    Cidade_Grava = gErr
    
    Select Case gErr
    
        Case 125022
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130009
        
        Case 125029
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142429)

    End Select

    Call Transacao_Rollback
    
    Exit Function
    
End Function

Public Function Cidade_Exclui(ByVal objCidades As ClassCidades) As Long
'Exclui do Banco de Dados o registro com o código passado como parâmetro

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim lTransacao As Long
Dim sDescricao As String
Dim lCodigo As Long

On Error GoTo Erro_Cidade_Exclui

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 125030

    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 125031
    Next
    
    sDescricao = String(STRING_CIDADE, 0)
    
    'Verifica se o Código já existe no Banco de Dados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, Descricao FROM Cidades WHERE Codigo=?", 0, iCodigo, sDescricao, objCidades.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 125032
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125033
    
    'Se não existe --> ERRO
    If lErro <> SUCESSO Then gError 125034
    
     'Verifica se o endereco existe, e se existir carrega seus dados em objEndereco
    lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Enderecos WHERE Cidade = ?", lCodigo, sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 207378

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207379

    'se há a cidade em algum Endereco ==> nao pode excluir
    If lErro = SUCESSO Then gError 207380
    
    'Realiza a locação do Código para poder exclui-lo
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 125035
    
    'Exclui o Código do Banco de Dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Cidades", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 125036
    
    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125037
    
    Cidade_Exclui = SUCESSO
    
    Exit Function
    
Erro_Cidade_Exclui:
    
    Cidade_Exclui = gErr
    
    Select Case gErr
    
        Case 125030
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 125031
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 125032, 125033
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr, objCidades.iCodigo)
        
        Case 125034
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA", gErr, objCidades.iCodigo)
        
        Case 125035
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CIDADES", gErr, objCidades.iCodigo)
        
        Case 125036
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CIDADES", gErr, objCidades.iCodigo)
        
        Case 125037
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 207378, 207379
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)
        
        Case 207380
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_EM_ENDERECOS", gErr, sDescricao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142430)

    End Select

    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Cidade_Grava_EmTrans(ByVal objCidades As ClassCidades) As Long
'Realiza a gravação do Cadastro no BD

Dim lErro As Long
Dim alComando(1 To 6) As Long
Dim iCodigo As Integer
Dim iIndice As Integer, lEndereco As Long
Dim sCidadeBD As String

On Error GoTo Erro_Cidade_Grava_EmTrans

    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 125023
    Next
    
    If objCidades.iCodigo = 0 Then
        lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigo)
        If lErro <> SUCESSO Then gError 125053
        objCidades.iCodigo = iCodigo
    End If
    
    'verifica se a cidade já está cadastrada com outro nome
    lErro = Comando_Executar(alComando(4), "SELECT Codigo FROM Cidades WHERE Codigo <> ?  AND Descricao = ?", iCodigo, objCidades.iCodigo, objCidades.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 124077
    
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124078
    If lErro = AD_SQL_SUCESSO Then gError 124079
    
    sCidadeBD = String(STRING_CIDADE, 0)
    
    'Verifica se o Código já existe no Banco de Dados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Descricao FROM Cidades WHERE Codigo=?", 0, sCidadeBD, objCidades.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 125024
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125025
    
    'Se retornou SUCESSO então realiza a atualização do registro
    If lErro = AD_SQL_SUCESSO Then
    
        'Loca o Código para realizar a modificação
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 125026
        
        'Realiza o UPDATE através do Código
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Cidades SET Descricao = ?, CodIBGE = ?", alComando(1), objCidades.sDescricao, objCidades.sCodIBGE)
        If lErro <> AD_SQL_SUCESSO Then gError 125027
        
        'se o nome da cidade foi alterado
        If objCidades.sDescricao <> sCidadeBD Then
                    
            lErro = Comando_ExecutarPos(alComando(5), "SELECT Codigo FROM Enderecos WHERE Cidade = ?", 0, lEndereco, sCidadeBD)
            If lErro <> AD_SQL_SUCESSO Then gError 124080
            
            lErro = Comando_BuscarProximo(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124081
        
            Do While lErro = AD_SQL_SUCESSO
            
                lErro = Comando_ExecutarPos(alComando(6), "UPDATE Enderecos SET Cidade = ?", alComando(5), objCidades.sDescricao)
                If lErro <> AD_SQL_SUCESSO Then gError 124082
            
                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124083
        
            Loop
        
        End If
    
    'Se não achou o registro então
    Else
        
        'Insere um novo registro no Banco
        lErro = Comando_Executar(alComando(3), "INSERT INTO Cidades (Codigo, Descricao, CodIBGE) VALUES (?, ?, ?)", objCidades.iCodigo, objCidades.sDescricao, objCidades.sCodIBGE)
        If lErro <> AD_SQL_SUCESSO Then gError 125028
        
    End If
    
    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cidade_Grava_EmTrans = SUCESSO
    
    Exit Function
    
Erro_Cidade_Grava_EmTrans:

    Cidade_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case 125023
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 125024, 125025, 124077, 124078
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CIDADES", gErr, objCidades.iCodigo)
            
        Case 125026
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CIDADES", gErr, objCidades.iCodigo)
        
        Case 125027
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CIDADES", gErr, objCidades.iCodigo)
        
        Case 125028
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CIDADES", gErr, objCidades.iCodigo)
            
        Case 125053
        
        Case 124079
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_OUTRO_CODIGO", gErr, iCodigo)
        
        Case 124080, 124081, 124082
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ENDERECOS", gErr)
        
        Case 124083
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ENDERECO", gErr, lEndereco)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142431)

    End Select

    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function


'#####################################################################################
Function Config_ObterAutomatico2(ByVal sTabelaObjeto As String, ByVal sCampo As String, lNumAuto As Long, Optional sFiltro As String = "") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto não incrementando o Config
'sTabelaObjeto é a Tabela do objeto
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.
'sFiltro vai ser usado para limitar a busca dentro da Tabela do objeto

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim lNumeroBD As Long
Dim sComandoSQL As String

On Error GoTo Erro_Config_ObterAutomatico2
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 132173
    Next
    
    If Len(Trim(sFiltro)) <> 0 Then sFiltro = " WHERE " & sFiltro
    
    'Monta comando SQL leitura na tabela do objeto
    sComandoSQL = "SELECT MAX(" & sCampo & ") FROM " & sTabelaObjeto & sFiltro

    'Pesquisa as numerações do objeto maiores ou iguais ao número automático
    lErro = Comando_Executar(lComando(1), sComandoSQL, lNumeroBD)
    If lErro <> AD_SQL_SUCESSO Then gError 132174
 
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132175

    lNumAuto = lNumeroBD + 1

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_ObterAutomatico2 = SUCESSO

    Exit Function

Erro_Config_ObterAutomatico2:

    Config_ObterAutomatico2 = gErr

    Select Case gErr

        Case 132173
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 132174, 132175
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabelaObjeto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142432)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function

Function Config_ObterAutomatico4(ByVal sTabelaObjeto As String, ByVal sCampo As String, lNumAuto As Double, Optional sFiltro As String = "") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto não incrementando o Config
'sTabelaObjeto é a Tabela do objeto
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.
'sFiltro vai ser usado para limitar a busca dentro da Tabela do objeto

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim lNumeroBD As Double
Dim sComandoSQL As String

On Error GoTo Erro_Config_ObterAutomatico4
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 132173
    Next
    
    If Len(Trim(sFiltro)) <> 0 Then sFiltro = " WHERE " & sFiltro
    
    'Monta comando SQL leitura na tabela do objeto
    sComandoSQL = "SELECT MAX(" & sCampo & ") FROM " & sTabelaObjeto & sFiltro

    'Pesquisa as numerações do objeto maiores ou iguais ao número automático
    lErro = Comando_Executar(lComando(1), sComandoSQL, lNumeroBD)
    If lErro <> AD_SQL_SUCESSO Then gError 132174
 
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132175

    lNumAuto = lNumeroBD + 1

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_ObterAutomatico4 = SUCESSO

    Exit Function

Erro_Config_ObterAutomatico4:

    Config_ObterAutomatico4 = gErr

    Select Case gErr

        Case 132173
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 132174, 132175
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabelaObjeto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142432)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function
'#########################################################################################

Function Config_Obter_Txt_Automatico(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, lNumAuto As Long, Optional ByVal sCampoFilial As String = "FilialEmpresa") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Config_Obter_Txt_Automatico
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 25193
    
    lErro = CF("Config_Obter_Txt_Automatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lNumAuto, sCampoFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 32304
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 25201

    Config_Obter_Txt_Automatico = SUCESSO

    Exit Function

Erro_Config_Obter_Txt_Automatico:

    Config_Obter_Txt_Automatico = Err

    Select Case Err

    Case 25193
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

    Case 32304
    
    Case 25201
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142433)

    End Select

    Call Transacao_Rollback
    
End Function

Function Config_Obter_Txt_Automatico_EmTrans(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, lNumAuto As Long, Optional ByVal sCampoFilial As String = "FilialEmpresa") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim lNumeroBD As Long, sNumeroBD As String
Dim sComandoSQL As String
Dim sConteudo As String, iDependeFilial As Integer

On Error GoTo Erro_Config_Obter_Txt_Automatico_EmTrans
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 25192
    Next

    sConteudo = String(STRING_BUFFER_MAX_TEXTO, 0)
    'Monta comando SQL de leitura na tabela de configuração
    sComandoSQL = "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?"

    'Pesquisa o número automático que está no BD para giFilialEmpresa
    lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 25194

    'Le o registro com o número automático
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25195
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa = EMPRESA_TODA Then Error 25202
    
    'se o codigo depende da filial
    iDependeFilial = 0
    
    'Se não achar pesquisa o número automático que está no BD para EMPRESA_TODA
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa <> EMPRESA_TODA Then
    
        lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 25203
    
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 25204
        
    Else
    
        If giFilialEmpresa <> EMPRESA_TODA Then iDependeFilial = 1
        
    End If
    
    'Locka o registro
    lErro = Comando_LockExclusive(lComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 25196

    lNumAuto = CLng(sConteudo)

    'Verifica se número automático ainda não foi utilizado
    
    sNumeroBD = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    If iDependeFilial = 0 Then
    
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, sNumeroBD, CStr(lNumAuto))
        If lErro <> AD_SQL_SUCESSO Then Error 25197
    
    Else
    
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampoFilial & " = ? AND " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, sNumeroBD, giFilialEmpresa, CStr(lNumAuto))
        If lErro <> AD_SQL_SUCESSO Then Error 32212
        
    End If
    
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25198

    'pesquisa o proximo número automático disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        If Not IsNumeric(sNumeroBD) Then Error 32305
        
        lNumeroBD = StrParaLong(sNumeroBD)
        
        Select Case lNumeroBD
            Case Is > lNumAuto
                Exit Do
            Case Is = lNumAuto
                lNumAuto = lNumAuto + 1
        End Select

        lErro = Comando_BuscarProximo(lComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25199

    Loop

    sConteudo = CStr(lNumAuto + 1)
    
    'Monta Comando SQL de atualizar número automático na tabela configuração
    sComandoSQL = "UPDATE " & sTabelaConfig & " SET Conteudo = ?"
    
    'Atualiza o numero automático no BD
    lErro = Comando_ExecutarPos(lComando(3), sComandoSQL, lComando(1), sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 25200

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_Obter_Txt_Automatico_EmTrans = SUCESSO

    Exit Function

Erro_Config_Obter_Txt_Automatico_EmTrans:

    Config_Obter_Txt_Automatico_EmTrans = Err

    Select Case Err

    Case 32305
        Call Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_NUMERICA", Err)
    
    Case 25192
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 25194, 25197, 25203, 32212
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaConfig)

    Case 25195, 25202, 25204
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25196
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25198, 25199
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaObjeto)
    
    Case 25200
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25201
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142434)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function


'#######################################################################
'Inserido por Wagner 30/09/05
Function Modulos_Filial_Gravar_Registro(ByVal iFilialEmpresa As Integer) As Long
'Tem que ser chamado dentro de transação
'Pega os dados da Filial Matriz e replica para a filial passada desde que
'o registro ainda não exista

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim sCodigo As String
Dim sDescricao As String
Dim iTipo As Integer
Dim sConteudo As String
Dim sConteudoInicial As String
Dim asModulo(1 To 13) As String
Dim iNumTabelas As Integer
Dim iAux As Integer
Dim sTabela As String

On Error GoTo Erro_Modulos_Filial_Gravar_Registro

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 140420
    Next
    
    'Tabelas de Config que vão ser atualizadas (Nome - Config)
    asModulo(1) = "CTB"
    asModulo(2) = "COMPRAS"
    asModulo(3) = "CP"
    asModulo(4) = "CPR"
    asModulo(5) = "CR"
    asModulo(6) = "CRFAT"
    asModulo(7) = "EST"
    asModulo(8) = "FAT"
    asModulo(9) = "FIS"
    asModulo(10) = "LOJA"
    asModulo(11) = "MAT"
    asModulo(12) = "PCP"
    asModulo(13) = "TES"
    
    iNumTabelas = 13
    
    For iIndice = 1 To iNumTabelas
    
        sTabela = asModulo(iIndice) & "Config"
    
        sCodigo = String(STRING_CONFIG_CODIGO, 0)
        sDescricao = String(STRING_CONFIG_DESCRICAO, 0)
        sConteudo = String(STRING_CONTEUDO, 0)
        sConteudoInicial = String(STRING_CONTEUDO, 0)
    
        'Le todos os registros na FILIAL_MATRIZ
        lErro = Comando_Executar(alComando(0), "SELECT Codigo, Descricao, ConteudoInicial, Tipo, Conteudo FROM " & sTabela & " WHERE FilialEmpresa =? ", sCodigo, sDescricao, sConteudoInicial, iTipo, sConteudo, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 140421
    
        'Busca o Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140422
    
        Do While lErro = AD_SQL_SUCESSO
    
            'Le o código da filial corrente para ver se existe
            lErro = Comando_Executar(alComando(1), "SELECT Tipo FROM " & sTabela & " WHERE FilialEmpresa =? AND Codigo = ? ", iAux, iFilialEmpresa, sCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 140423
        
            'Busca o Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140424
    
            'Nao encontrou o registro => Insere
            If lErro = AD_SQL_SEM_DADOS Then
            
                If Len(Trim(sConteudoInicial)) = 0 Then sConteudoInicial = sConteudo
            
                lErro = Comando_Executar(alComando(2), "INSERT INTO " & sTabela & " (FilialEmpresa, Codigo, Descricao, Tipo, Conteudo, ConteudoInicial) VALUES (?,?,?,?,?,?)", _
                    iFilialEmpresa, sCodigo, sDescricao, iTipo, sConteudoInicial, sConteudoInicial)
                If lErro <> AD_SQL_SUCESSO Then gError 140425
    
            End If
            
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140426
            
        Loop
        
    Next

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Modulos_Filial_Gravar_Registro = SUCESSO

    Exit Function

Erro_Modulos_Filial_Gravar_Registro:

    Modulos_Filial_Gravar_Registro = gErr

    Select Case gErr

        Case 140420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140421 To 140424, 140426
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_" & UCase(sTabela), gErr, sCodigo)

        Case 140425
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_" & UCase(sTabela), gErr, sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142435)

    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function
'#######################################################################
'#######################################################################
'Inserido por Wagner 31/01/2006
Function Valida_Controle() As Long

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim sPath As String
Dim sPathAux As String
Dim sCaminhoIni As String
Dim sCaminhoBD As String
Dim sDescricao As String
Dim lTypedrive As Long
Dim sComputador As String
Dim sConteudo As String

On Error GoTo Erro_Valida_Controle

    'Abre o comando
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 141618

    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando2 = 0 Then gError 141618

    lComando3 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando3 = 0 Then gError 141618

    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 141619

    sPath = String(STRING_CONTROLE_CONTEUDO, 0)
    sCaminhoIni = String(STRING_CONTROLE_CONTEUDO, 0)
    sCaminhoBD = String(STRING_CONTROLE_CONTEUDO, 0)
    sComputador = String(512, 0)
    
    Call GetPrivateProfileString("Geral", "LockFile", "", sPath, 128, "ADM100.INI")
    sPath = Replace(sPath, Chr(0), "")

    'Lê o Path do arquivo de Lock
    lErro = Comando_Executar(lComando1, "SELECT Conteudo FROM Controle WHERE Codigo = ? ", sCaminhoBD, CONTROLE_PATH_LOCK)
    If lErro <> AD_SQL_SUCESSO Then gError 141620

    'Busca o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141621

    If Len(Trim(sPath)) = 0 Then
        sCaminhoIni = sPath
        gError 141623
    End If
    
    sPathAux = sPath
        
    If InStr(1, sPathAux, ":") = 2 Then
        Call Converte_Letra_para_UNC(sPathAux, sCaminhoIni)
    Else
        sCaminhoIni = sPathAux
    End If
     
    If InStr(1, sCaminhoIni, "\") <> 1 Then
    
        Call GetComputerName(sComputador, Len(sComputador))
        
        sComputador = Replace(sComputador, Chr(0), "")
        
        sCaminhoIni = "\\" & sComputador & "\" & Replace(sCaminhoIni, ":", "")
    
    End If

    'Se não achou => Insere o Path do usuário atual
    If lErro = AD_SQL_SEM_DADOS Then
    
        sDescricao = "Path do AdLock"
    
        lErro = Comando_Executar(lComando2, "INSERT INTO Controle (Codigo, Descricao, Conteudo) VALUES (?,?,?)", CONTROLE_PATH_LOCK, sDescricao, sCaminhoIni)
        If lErro <> AD_SQL_SUCESSO Then gError 141622
    
    Else
    
        If UCase(sCaminhoIni) <> UCase(sCaminhoBD) Then gError 141623
    
    End If
        
    sConteudo = String(STRING_CONTROLE_CONTEUDO, 0)
        
    'Lê a indicacao se usa o Browse_Posicao como antigamente ou o novo
    lErro = Comando_Executar(lComando3, "SELECT Conteudo FROM Controle WHERE Codigo = ? ", sConteudo, CONTROLE_BROWSE_POSICAO_ANTIGO)
    If lErro <> AD_SQL_SUCESSO Then gError 141620

    'Busca o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141621
        
    If lErro = AD_SQL_SUCESSO Then giBrowsePosicaoAntigo = StrParaInt(sConteudo)
        
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 141624
    
    Valida_Controle = SUCESSO
    
    Exit Function

Erro_Valida_Controle:

    Valida_Controle = gErr

    Select Case gErr
   
        Case 141618
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 141619
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
       
        Case 141620, 141621
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTROLE", gErr)
            
        Case 141622
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTROLE", gErr)
        
        Case 141623
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMINHOS_ADLOCK_DIFERENTES", gErr, sCaminhoIni, sCaminhoBD)
        
        Case 141624
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142436)
        
    End Select
    
    Call Transacao_RollbackDic
   
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function
'#######################################################################

Public Function Exclusao_Valida(ByVal sCodigo As String, ByVal vValorCampo As Variant) As Long

Dim lErro As Long
Dim lComando As Long
Dim colValidaExclusoes As New Collection
Dim objValidaExclusoes As ClassValidaExclusoes
Dim sSelect As String
Dim vValorRetorno As Variant
Dim vValorCampoMascarado As Variant
Dim vValorRetornoMascarado As Variant
Dim sGenero As String

On Error GoTo Erro_Exclusao_Valida

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 177935
    
    lErro = CF("ValidaExclusoes_Le", sCodigo, colValidaExclusoes)
    If lErro <> SUCESSO Then gError 177936
    
    For Each objValidaExclusoes In colValidaExclusoes
    
        If objValidaExclusoes.iGeneroMsgErro = VALIDAEXCLUSAO_GENERO_MASCULINO Then
            sGenero = "o"
        Else
            sGenero = "a"
        End If
    
        lErro = CF("Exclusao_Valida_Mascara", 0, objValidaExclusoes.iSubTipoCampoProc, vValorCampo, vValorCampoMascarado)
        If lErro <> SUCESSO Then gError 177949

        sSelect = "SELECT " & objValidaExclusoes.sCampoLer & " FROM  " & objValidaExclusoes.sTabela & " WHERE " & objValidaExclusoes.sCampo & " = ?"
    
        lErro = Exclusao_Valida_AlocaBuffer(objValidaExclusoes, vValorRetorno)
        If lErro <> SUCESSO Then gError 177937
    
        'Le a tabelaValidaExclusoes
        lErro = Comando_Executar(lComando, sSelect, vValorRetorno, vValorCampo)
        If lErro <> AD_SQL_SUCESSO Then gError 177938
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177939
    
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = CF("Exclusao_Valida_Mascara", objValidaExclusoes.iTipoCampoLer, objValidaExclusoes.iSubTipoCampoLer, vValorRetorno, vValorRetornoMascarado)
            If lErro <> SUCESSO Then gError 177949

            gError 177940
        End If
        
    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exclusao_Valida = SUCESSO

    Exit Function

Erro_Exclusao_Valida:

    Exclusao_Valida = gErr

    Select Case gErr

        Case 177935
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 177936, 177937, 177949

        Case 177938 To 177939
            Call Rotina_Erro(vbOKOnly, objValidaExclusoes.sMsgErroLer, gErr)

        Case 177940
            Call Rotina_Erro(vbOKOnly, "ERRO_EXLUSAO_VALIDA", gErr, objValidaExclusoes.sMsgErro1, vValorCampoMascarado, sGenero, sGenero, objValidaExclusoes.sMsgErro2, vValorRetornoMascarado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177941)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Exclusao_Valida_AlocaBuffer(ByVal objValidaExclusoes As ClassValidaExclusoes, vValorRetorno As Variant) As Long

On Error GoTo Erro_Exclusao_Valida_AlocaBuffer

    Select Case objValidaExclusoes.iTipoCampoLer

        Case ADM_TIPO_SMALLINT
            vValorRetorno = CInt(0)

        Case ADM_TIPO_INTEGER
            vValorRetorno = CLng(0)

        Case ADM_TIPO_DOUBLE
            vValorRetorno = CDbl(0)

        Case ADM_TIPO_VARCHAR
            If objValidaExclusoes.iTamanhoCampoLer < 500 Then
                vValorRetorno = String(500, 0)
            Else
                vValorRetorno = String(objValidaExclusoes.iTamanhoCampoLer, 0)
            End If

        Case ADM_TIPO_DATE
            vValorRetorno = CDate("1/1/1997")

        Case Else
            gError 177942

    End Select
    
    Exclusao_Valida_AlocaBuffer = SUCESSO

    Exit Function

Erro_Exclusao_Valida_AlocaBuffer:

    Exclusao_Valida_AlocaBuffer = gErr

    Select Case gErr

        Case 177942
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", gErr, objValidaExclusoes.iTipoCampoLer)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177943)

    End Select
    
    Exit Function

End Function

Public Function Exclusao_Valida_Mascara(ByVal iTipo As Integer, ByVal iSubTipo As Integer, ByVal vValor As Variant, vValorMascarado As Variant, Optional ByVal bExcel As Boolean = False, Optional ByVal lComando1 As Long, Optional ByVal lComando2 As Long, Optional ByVal iExercicio As Integer = 0) As Long

Dim lErro As Long
Dim sConta As String
Dim sContaMascarada As String
Dim sCcl As String
Dim sCclMascarado As String
Dim sProduto As String
Dim sProdutoMascarado As String
Dim sCgc As String
Dim sDocOrigem As String
Dim iIndice As Integer
Dim iExercicioAux As Integer
Dim iPeriodo As Integer
Dim sNomeExterno As String

On Error GoTo Erro_Exclusao_Valida_Mascara

    Select Case iTipo

        Case ADM_TIPO_DOUBLE
        
            Select Case iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    vValorMascarado = Format(vValor, "Percent")
                        
                Case ADM_SUBTIPO_HORA
                    vValorMascarado = CDate(vValor)
                    
                Case Else
                    vValorMascarado = Format(vValor, "Standard")
                    
            End Select

        Case ADM_TIPO_DATE
           
            Select Case iSubTipo
    
                Case Else
                
                    vValorMascarado = Format(vValor, "dd/mm/yyyy")

            End Select
    
        Case ADM_TIPO_SMALLINT
    
            Select Case iSubTipo
            
                Case ADM_SUBTIPO_EXERCICIO
                
                    iExercicioAux = vValor
                
                    'le o exercicio do BD e exibe o nome externo
                    lErro = CF("Exercicio_Le_NomeExterno", lComando1, iExercicioAux, sNomeExterno)
                    If lErro <> SUCESSO Then gError 9261
                    
                    vValorMascarado = sNomeExterno
                
                Case ADM_SUBTIPO_PERIODO
                
                    If iExercicio <> 0 Then
                                    
                        iPeriodo = vValor
    
                        'le o periodo do BD e exibe o nome externo
                        lErro = CF("Periodo_Le_NomeExterno", lComando2, iExercicio, iPeriodo, sNomeExterno)
                        If lErro <> SUCESSO Then gError 9269
                                                
                        vValorMascarado = sNomeExterno
                        
                    Else

                        vValorMascarado = vValor
                        
                    End If
        
                Case Else
                    vValorMascarado = CStr(vValor)
                    
            End Select
            
        Case ADM_TIPO_INTEGER
        
            Select Case iSubTipo
        
                Case ADM_SUBTIPO_TIPO_DOC_PROJETO
                
                    lErro = CF("Projeto_Le_DocInfo", vValor, sDocOrigem)
                    If lErro <> SUCESSO Then gError 182052
                    
                    vValorMascarado = sDocOrigem
                    
                Case Else
                    vValorMascarado = CStr(vValor)
    
            End Select
                
        Case Else
            
            Select Case iSubTipo
        
                Case ADM_SUBTIPO_PERCENTUAL
                    vValorMascarado = Format(vValor, "Percent")
                        
                Case ADM_SUBTIPO_HORA
                    vValorMascarado = CDate(vValor)
       
                Case ADM_SUBTIPO_CONTA
                
                    If Len(vValor) > 0 Then
                
                        sConta = vValor
                    
                        lErro = CF("Mascara_MascararConta", sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 177950
                        
                    Else
                
                        sContaMascarada = ""
                        
                    End If
                
                    vValorMascarado = sContaMascarada
        
                Case ADM_SUBTIPO_CCL
            
                    If Len(vValor) > 0 Then
                    
                        sCcl = vValor
        
                        lErro = CF("Mascara_MascararCcl", sCcl, sCclMascarado)
                        If lErro <> SUCESSO Then gError 177951
                        
                    Else
                    
                        sCclMascarado = ""
                        
                    End If
                
                    vValorMascarado = sCclMascarado
                    
                    
                Case ADM_SUBTIPO_PRODUTO
                
                    If Len(vValor) > 0 Then
                    
                        sProduto = vValor
                        
                        lErro = Mascara_MascararProduto(sProduto, sProdutoMascarado)
                        If lErro <> SUCESSO Then gError 177952
                        
                        If bExcel Then
                            If IsNumeric(sProduto) Then
                                sProdutoMascarado = "'" & sProdutoMascarado
                            End If
                        End If
                        
                    Else
                
                        sProdutoMascarado = ""
                        
                    End If
                
                    vValorMascarado = sProdutoMascarado

                Case ADM_SUBTIPO_CGC
                
                    Select Case Len(Trim(vValor))
                    
                        Case STRING_CPF 'CPF
                                        
                            sCgc = Format(vValor, "000\.000\.000-00; ; ; ")
                    
                        Case STRING_CGC 'CGC
                    
                            sCgc = Format(vValor, "00\.000\.000\/0000-00; ; ; ")
                            
                    End Select
                    
                    vValorMascarado = sCgc
                
                Case Else
                
                    vValorMascarado = CStr(vValor)
                    
            End Select
            
    End Select
    
    Exclusao_Valida_Mascara = SUCESSO
    
    Exit Function
    
Erro_Exclusao_Valida_Mascara:

    Exclusao_Valida_Mascara = gErr

    Select Case gErr
        
        Case 177950
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", gErr, CStr(vValor))
       
        Case 177951
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, CStr(vValor))
                   
        Case 177952
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, sProduto)
            
        Case 182052
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177949)

    End Select

    Exit Function

End Function

Function RegraWFW_Grava(ByVal colRegraWFW As Collection, ByVal iAtivoWFW As Integer, ByVal sUsuario As String) As Long
'grava as regras associados ao modulo/transacao passados como parametro.

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objRegraWFW As ClassRegraWFW
Dim lTransacao As Long
Dim iWorkFlowAtivo As Integer
Dim lTransacaoDic As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_RegraWFW_Grava
    
    If colRegraWFW.Count > 0 Then
    
        'abre a transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 178196
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 178091
        Next
    
        Set objRegraWFW = colRegraWFW.Item(1)
    
        lErro = RegraWFW_Exclui_Trans(objRegraWFW)
        If lErro <> SUCESSO Then gError 178099
        
        For Each objRegraWFW In colRegraWFW
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RegraWFW (Modulo, Transacao, Item, CodUsuario, Regra, TipoBloqueio, EmailPara, EmailAssunto, EmailMsg, AvisoMsg, LogDoc, LogMsg, RelModulo, RelNome, RelOpcao, BrowseModulo, BrowseNome, BrowseOpcao, RelPorEmail, RelSel, RelAnexo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem, objRegraWFW.sUsuario, objRegraWFW.sRegra, objRegraWFW.iTipoBloqueio, objRegraWFW.sEmailPara, objRegraWFW.sEmailAssunto, objRegraWFW.sEmailMsg, _
                objRegraWFW.sAvisoMsg, objRegraWFW.sLogDoc, objRegraWFW.sLogMsg, objRegraWFW.sRelModulo, objRegraWFW.sRelNome, objRegraWFW.sRelOpcao, objRegraWFW.sBrowseModulo, objRegraWFW.sBrowseNome, objRegraWFW.sBrowseOpcao, objRegraWFW.iRelPorEmail, objRegraWFW.sRelSel, objRegraWFW.sRelAnexo)
            If lErro <> AD_SQL_SUCESSO Then gError 178095
        
            For iIndice = 1 To objRegraWFW.colUsuarios.Count
                
                sUsuario = objRegraWFW.colUsuarios(iIndice)
                
                lErro = Comando_Executar(alComando(2), "INSERT INTO AvisoUsuWFW (Modulo, Transacao, Item, CodUsuarioOrigem, CodUsuario) VALUES (?,?,?,?,?)", _
                objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem, objRegraWFW.sUsuario, sUsuario)
                If lErro <> AD_SQL_SUCESSO Then gError 178094
        
            Next
        
        Next
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 178198
        
    End If
    
    objUsuarios.sCodUsuario = sUsuario
    
    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO And lErro <> 40832 Then gError 178411

    If lErro <> AD_SQL_SUCESSO Then gError 178413
    
    If objUsuarios.iWorkFlowAtivo <> iAtivoWFW Then
    
        lTransacaoDic = Transacao_AbrirDic
        If lTransacaoDic = 0 Then gError 178414
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
            If alComando(iIndice) = 0 Then gError 178412
        Next
    
        lErro = Comando_Executar(alComando(1), "UPDATE Usuarios SET WorkflowAtivo = ? WHERE CodUsuario = ?", iAtivoWFW, sUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 178415

        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next

        'confirma a transação
        lErro = Transacao_CommitDic
        If lErro <> AD_SQL_SUCESSO Then gError 178416

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RegraWFW_Grava = SUCESSO
    
    Exit Function

Erro_RegraWFW_Grava:

    RegraWFW_Grava = gErr
    
    Select Case gErr
                
        Case 178091, 178412
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178094
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_AVISOUSUWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem)
                  
        Case 178095
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_REGRAWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem)
                  
        Case 178099, 178411
                  
        Case 178196, 178414
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                  
        Case 178198, 178416
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case 178413
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, sUsuario)
                  
        Case 178415
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIO", gErr, sUsuario)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178096)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
    Call Transacao_RollbackDic
        
    Exit Function

End Function

Function RegraWFW_Exclui(ByVal objRegraWFW As ClassRegraWFW) As Long
'exclui as regras associados ao modulo/transacao passados como parametro.

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lTransacao As Long
    
On Error GoTo Erro_RegraWFW_Exclui
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178197
    
    lErro = RegraWFW_Exclui_Trans(objRegraWFW)
    If lErro <> SUCESSO Then gError 178200
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 178199
    
    RegraWFW_Exclui = SUCESSO
    
    Exit Function

Erro_RegraWFW_Exclui:

    RegraWFW_Exclui = gErr
    
    Select Case gErr
                
        Case 178197
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                  
        Case 178199
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case 178200
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178096)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
        
    Exit Function

End Function
                    
Function RegraWFW_Exclui_Trans(ByVal objRegraWFW As ClassRegraWFW) As Long
'exclui as regras associados ao modulo/transacao passados como parametro.

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lTransacao As Long
    
On Error GoTo Erro_RegraWFW_Exclui_Trans
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178098
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM AvisoUsuWFW WHERE Modulo = ? AND Transacao = ? AND (CodUsuarioOrigem = ? OR CodUsuarioOrigem = ?)", objRegraWFW.sModulo, objRegraWFW.iTransacao, gsUsuario, "")
    If lErro <> AD_SQL_SUCESSO Then gError 178092
    
    lErro = Comando_Executar(alComando(2), "DELETE FROM RegraWFW WHERE Modulo = ? AND Transacao = ? AND (CodUsuario = ? OR CodUsuario = ?)", objRegraWFW.sModulo, objRegraWFW.iTransacao, gsUsuario, "")
    If lErro <> AD_SQL_SUCESSO Then gError 178093
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RegraWFW_Exclui_Trans = SUCESSO
    
    Exit Function

Erro_RegraWFW_Exclui_Trans:

    RegraWFW_Exclui_Trans = gErr
    
    Select Case gErr
                
        Case 178098
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178092
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_AVISOUSUWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao)
                  
        Case 178093
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGRAWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178096)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function AvisoWFW_Grava(ByVal colUsuario As Collection, ByVal sMsg As String, ByVal iTransacao As Integer) As Long
'grava o aviso passado como parametro

Dim lErro As Long
Dim lTransacao As Long
    
On Error GoTo Erro_AvisoWFW_Grava
    
    If colUsuario.Count > 0 Then
    
        'abre a transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 178407
    
        lErro = CF("AvisoWFW_Grava_EmTrans", colUsuario, sMsg, iTransacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Confirma a transação
        lErro = Transacao_Commit
        If lErro <> AD_SQL_SUCESSO Then gError 178408
    
    End If
    
    AvisoWFW_Grava = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Grava:

    AvisoWFW_Grava = gErr
    
    Select Case gErr
                
        Case 178407
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                  
        Case 178408
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178169)
        
    End Select
        
    Call Transacao_Rollback
        
    Exit Function

End Function
                    
Function AvisoWFW_Grava_EmTrans(ByVal colUsuario As Collection, ByVal sMsg As String, ByVal iTransacao As Integer) As Long
'grava o aviso passado como parametro

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim vUsuario As Variant
Dim dHora As Double
    
On Error GoTo Erro_AvisoWFW_Grava_EmTrans
  
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178166
    Next

    lErro = CF("Config_ObterNumInt", "ADMConfig", "NUM_PROX_AVISOWFW", lNumIntDoc, colUsuario.Count)
    If lErro <> SUCESSO Then gError 178167

    For Each vUsuario In colUsuario
    
        dHora = CDbl(Time)
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO AvisoWFW (NumIntDoc, Msg, Data, Hora, Usuario, Transacao, UsuarioOrig) VALUES (?,?,?,?,?,?,?)", _
        lNumIntDoc, sMsg, Date, dHora, vUsuario, iTransacao, gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 178168
    
        lNumIntDoc = lNumIntDoc + 1
    
    Next
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    AvisoWFW_Grava_EmTrans = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Grava_EmTrans:

    AvisoWFW_Grava_EmTrans = gErr
    
    Select Case gErr
                
        Case 178166
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178167
                  
        Case 178168
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_AVISOWFW", gErr, lNumIntDoc)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178169)
        
    End Select
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function
                    
Function LogWFW_Grava(ByVal sDoc As String, ByVal sMsg As String, ByVal iTransacao As Integer) As Long
'grava o log passado como parametro

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim dHora As Double
Dim lTransacao As Long
    
On Error GoTo Erro_LogWFW_Grava
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178233
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178406
    Next

    lErro = CF("Config_ObterNumInt", "ADMConfig", "NUM_PROX_LOGWFW", lNumIntDoc)
    If lErro <> SUCESSO Then gError 178171

    dHora = CDbl(Time)

    lErro = Comando_Executar(alComando(1), "INSERT INTO LogWFW (NumIntDoc, Doc, Msg, Data, Hora, Usuario, Transacao) VALUES (?,?,?,?,?,?,?)", _
    lNumIntDoc, sDoc, sMsg, Date, dHora, gsUsuario, iTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 178172
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 178409
    
    LogWFW_Grava = SUCESSO
    
    Exit Function

Erro_LogWFW_Grava:

    LogWFW_Grava = gErr
    
    Select Case gErr
                
        Case 178406
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178171
                  
        Case 178172
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_LOGWFW", gErr, lNumIntDoc)
                  
        Case 178233
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                  
        Case 178409
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178173)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function AvisoWFW_Atualiza_DataHoraUlt(ByVal colAvisoWFW As Collection) As Long
'grava a data e hora atual em datault e horault no aviso passado como parametro

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objAvisoWFW As ClassAvisoWFW
    
    
On Error GoTo Erro_AvisoWFW_Atualiza_DataHoraUlt
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178233
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178234
    Next
    
    For Each objAvisoWFW In colAvisoWFW
    
        lErro = Comando_Executar(alComando(1), "UPDATE AvisoWFW SET DataUltAviso = ? , HoraUltAviso = ? WHERE NumIntDoc = ?", _
        Date, CDbl(Time), objAvisoWFW.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 178235
        
    Next
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 178236
    
    AvisoWFW_Atualiza_DataHoraUlt = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Atualiza_DataHoraUlt:

    AvisoWFW_Atualiza_DataHoraUlt = gErr
    
    Select Case gErr
                
        Case 178233
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 178234
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178235
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_AVISOWFW", gErr, objAvisoWFW.lNumIntDoc)
                  
        Case 178236
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178237)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
        
    Exit Function

End Function

'Function WorkFlow_Trata_Transacao(ByVal sSigla As String, ByVal objTela As Object, ByVal sCodUsuario As String, Optional ByVal objContexto As Object) As Long
''grava a data e hora atual em datault e horault no aviso passado como parametro
'
'Dim lErro As Long
'Dim alComando(1 To 1) As Long
'Dim lTransacao As Long
'Dim iIndice As Integer
'Dim objAvisoWFW As ClassAvisoWFW
'
'
'On Error GoTo Erro_WorkFlow_Trata_Transacao
'
'    'abre a transacao
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 178388
'
'    lErro = WorkFlow_Trata_Transacao_Trans(sSigla, objTela, sCodUsuario, objContexto)
'    If lErro <> SUCESSO Then gError 178389
'
'    'Confirma a transação
'    lErro = Transacao_Commit
'    If lErro <> AD_SQL_SUCESSO Then gError 178390
'
'    WorkFlow_Trata_Transacao = SUCESSO
'
'    Exit Function
'
'Erro_WorkFlow_Trata_Transacao:
'
'    WorkFlow_Trata_Transacao = gErr
'
'    Select Case gErr
'
'        Case 178388
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 178389
'
'        Case 178390
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178391)
'
'    End Select
'
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function

Function WorkFlow_Trata_Transacao(ByVal sSigla As String, ByVal objTela As Object, ByVal sCodUsuario As String, Optional ByVal objContexto As Object, Optional ByVal iCodigoTransacao As Integer = 0) As Long

Dim lComando As Long
Dim lErro As Long
Dim colRegraWFW As New Collection
Dim objRegraWFW As ClassRegraWFW
Dim objRegraWFW1 As ClassRegraWFW
Dim colRegraWFW1 As New Collection
Dim iIndice As Integer
Dim objEmail As Object
Dim lTransacao As Long
Dim iIncremento As Integer
Dim objRelatorio As New AdmRelatorio
Dim colSelecao As New Collection
Dim objUsuarios As New ClassUsuarios
Dim objBloqueio As Object, sRelSelecao As String

On Error GoTo Erro_WorkFlow_Trata_Transacao

    objUsuarios.sCodUsuario = sCodUsuario

    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO Then gError 178419

    If objUsuarios.iWorkFlowAtivo = WORKFLOW_ATIVO And Not (objTela Is Nothing) Then

        If iCodigoTransacao <> 0 Then
            lErro = CF("RegraWFW_Le_Transacao1", sSigla, objTela.Name, sCodUsuario, colRegraWFW1, iCodigoTransacao)
        Else
            lErro = CF("RegraWFW_Le_Transacao1", sSigla, objTela.Name, sCodUsuario, colRegraWFW1)
        End If
        If lErro <> SUCESSO Then gError 178174
    
        lErro = CF("Executa_ExpressaoWFW", sSigla, objTela, colRegraWFW, colRegraWFW1, objContexto)
        If lErro <> SUCESSO Then gError 178175
        
        For iIndice = 1 To colRegraWFW1.Count
        
            Set objRegraWFW = colRegraWFW.Item(iIndice + iIncremento)
            Set objRegraWFW1 = colRegraWFW1.Item(iIndice)
            
            objRegraWFW1.dRegraRet = objRegraWFW.dRegraRet
        
            If objRegraWFW.dRegraRet <> 0 Then
                
                If Len(Trim(objRegraWFW1.sAvisoMsg)) > 0 Or Len(Trim(objRegraWFW1.sEmailPara)) > 0 Or Len(Trim(objRegraWFW1.sEmailAssunto)) > 0 Or Len(Trim(objRegraWFW1.sEmailMsg)) > 0 Or Len(Trim(objRegraWFW1.sLogDoc)) > 0 Or Len(Trim(objRegraWFW1.sLogMsg)) > 0 Then
                
                    iIncremento = iIncremento + 1
                    
                    Set objRegraWFW = colRegraWFW.Item(iIndice + iIncremento)
                
                    If Len(objRegraWFW.sEmailParaRet) > 0 And objRegraWFW1.iRelPorEmail = 0 Then
                    
                        Set objEmail = CreateObject("ADRel2.ClassInterfaceAdRel2")
                        Call CF("EmailConfig_Grava_INI")
                        lErro = objEmail.Email_Enviar(objRegraWFW.sEmailParaRet, objRegraWFW.sEmailAssuntoRet, objRegraWFW.sEmailMsgRet, "", "", "", "", "")
                        If lErro <> SUCESSO Then gError 178176
                    
                    End If
                    
                    If Len(objRegraWFW.sAvisoMsgRet) > 0 Then
                
                            'gravar o aviso
                            lErro = CF("AvisoWFW_Grava", objRegraWFW1.colUsuarios, objRegraWFW.sAvisoMsgRet, objRegraWFW1.iTransacao)
                            If lErro <> SUCESSO Then gError 178177
                    
                    End If
                    
                    If Len(objRegraWFW.sLogDocRet) > 0 Or Len(objRegraWFW.sLogMsgRet) > 0 Then
                    
                        'gravar o log
                        lErro = CF("LogWFW_Grava", objRegraWFW.sLogDocRet, objRegraWFW.sLogMsgRet, objRegraWFW1.iTransacao)
                        If lErro <> SUCESSO Then gError 178178
                    
                    End If
            
                End If
                        
                If Len(Trim(objRegraWFW1.sRelNome)) > 0 Then
                
                    If Len(Trim(objRegraWFW1.sRelOpcao)) > 0 Then
                        objRelatorio.sOpcao = objRegraWFW1.sRelOpcao
                    Else
                        'serve para indicar que esta chamando para Workflow, ou seja, vai executar direto sem passar pela tela de opcao
                        objRelatorio.sOpcao = "?*?"
                    End If
                    
                    If Len(Trim(objRegraWFW.sEmailParaRet)) > 0 And objRegraWFW1.iRelPorEmail <> 0 Then
                    
                        If Len(Trim(objRegraWFW.sRelAnexoRet)) <> 0 Then
                            Call objRelatorio.ExecutarDiretoEmail(objRegraWFW1.sRelNome, objRegraWFW.sRelSelRet, 0, "", "TTO_EMAIL", objRegraWFW.sEmailParaRet, "TSUBJECT", objRegraWFW.sEmailAssuntoRet, "TMENSAGEM", objRegraWFW.sEmailMsgRet, "TALIASATTACH", objRegraWFW.sRelAnexoRet)
                        Else
                            Call objRelatorio.ExecutarDiretoEmail(objRegraWFW1.sRelNome, objRegraWFW.sRelSelRet, 0, "", "TTO_EMAIL", objRegraWFW.sEmailParaRet, "TSUBJECT", objRegraWFW.sEmailAssuntoRet, "TMENSAGEM", objRegraWFW.sEmailMsgRet)
                        End If
                        
                    Else
                    
                        objRelatorio.Rel_Menu_Executar (objRegraWFW1.sRelNome)
                
                    End If
                    
                End If
            
                If Len(Trim(objRegraWFW1.sBrowseNome)) > 0 Then
                
                    Call Chama_Tela(objRegraWFW1.sBrowseNome, colSelecao, , , , , objRegraWFW1.sBrowseOpcao)
                
                End If
            
                If objRegraWFW1.iTransacao = TRANSACAOWFW_PEDIDO_VENDA Then
                    If objRegraWFW1.iTipoBloqueio > 0 Then
                    
                        Set objBloqueio = CreateObject("GlobaisMAT.ClassBloqueioPV")
                    
                        Set objContexto.colBloqueiosPV = CreateObject("GlobaisMAT.ColBloqueioPV")
                    
                        objContexto.colBloqueiosPV.Add objContexto.iFilialEmpresa, objContexto.lCodigo, 0, objRegraWFW1.iTipoBloqueio, gsUsuario, "BLOQUEIO AUTOMÁTICO", gdtDataHoje, "", "", DATA_NULA, ""
                        
                        lErro = CF("BloqueiosPV_Grava_Trans", objContexto)
                        If lErro <> SUCESSO Then gError 178466
                        
                    
                    End If
                
                End If
        
            End If
        
        Next
        
        'atualiza a data e hora da ultima execucao verdadeira das regras
        lErro = RegraWFW_Atualiza_DataHora(colRegraWFW1)
        If lErro <> SUCESSO Then gError 178372
    
    End If
    
    WorkFlow_Trata_Transacao = SUCESSO

    Exit Function

Erro_WorkFlow_Trata_Transacao:

    WorkFlow_Trata_Transacao = gErr
    
    Select Case gErr
    
        Case 178174 To 178178, 178372, 178419, 178466, ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178145)
        
    End Select
    
    Exit Function

End Function

Function AvisoWFW_Exclui(ByVal colNumIntDoc As Collection) As Long
'exclui uma colecao de avisos

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim lTransacao As Long
Dim vNumIntDoc As Variant
Dim iIndice As Integer
    
On Error GoTo Erro_AvisoWFW_Exclui
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178207
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178208
    Next
    
    For Each vNumIntDoc In colNumIntDoc
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM AvisoWFW WHERE NumIntDoc = ?", _
        vNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 178209
        
    Next
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 178210
    
    AvisoWFW_Exclui = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Exclui:

    AvisoWFW_Exclui = gErr
    
    Select Case gErr
                
        Case 178207
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 178208
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178209
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_AVISOWFW", gErr, vNumIntDoc)
                  
        Case 178210
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178169)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
        
    Exit Function

End Function

Function AvisoWFW_Atualiza_Intervalo(ByVal colNumIntDoc As Collection, ByVal dIntervalo As Double, ByVal iUMIntervalo As Integer) As Long
'grava a data e hora atual em datault e horault no aviso passado como parametro

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim vNumIntDoc As Variant
    
    
On Error GoTo Erro_AvisoWFW_Atualiza_Intervalo
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178228
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178229
    Next
    
    For Each vNumIntDoc In colNumIntDoc
    
        lErro = Comando_Executar(alComando(1), "UPDATE AvisoWFW SET Intervalo = ? , UMIntervalo = ? WHERE NumIntDoc = ?", _
        dIntervalo, iUMIntervalo, vNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 178230
        
    Next
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 178231
    
    AvisoWFW_Atualiza_Intervalo = SUCESSO
    
    Exit Function

Erro_AvisoWFW_Atualiza_Intervalo:

    AvisoWFW_Atualiza_Intervalo = gErr
    
    Select Case gErr
                
        Case 178228
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 178229
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178230
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_AVISOWFW", gErr, vNumIntDoc)
                  
        Case 178231
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178232)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Call Transacao_Rollback
        
    Exit Function

End Function

Function BrowseOpcao_Grava_EmTrans(ByVal sOpcao As String, ByVal objBrowseUsuario As AdmBrowseUsuario) As Long
'grava na tabela BrowseOpcao os campos contidos em objBrowseUsuario

Dim lErro As Long
Dim alComando(0 To 0) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim sNomeTela As String

On Error GoTo Erro_BrowseOpcao_Grava_EmTrans

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 178268
    Next

    lErro = BrowseOpcao_Exclui_EmTrans(sOpcao, objBrowseUsuario.sNomeTela)
    If lErro <> SUCESSO Then gError 178281
    
    'se BrowseUsuario não estiver cadastrado ==> é uma inclusão
    lErro = Comando_Executar(alComando(0), "INSERT INTO BrowseOpcao (NomeTela, Opcao, Topo, Esquerda, Largura, Altura) VALUES (?, ?, ?, ?, ?, ?)", objBrowseUsuario.sNomeTela, sOpcao, objBrowseUsuario.lTopo, objBrowseUsuario.lEsquerda, objBrowseUsuario.lLargura, objBrowseUsuario.lAltura)
    If lErro <> AD_SQL_SUCESSO Then gError 178273
        
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   BrowseOpcao_Grava_EmTrans = SUCESSO
   
   Exit Function
   
Erro_BrowseOpcao_Grava_EmTrans:

   BrowseOpcao_Grava_EmTrans = gErr

   Select Case gErr

        Case 178268
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178281
        
        Case 178273
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEOPCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178274)

   End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function BrowseOpcao_Exclui_EmTrans(ByVal sOpcao As String, ByVal sNomeTela As String) As Long
'grava na tabela BrowseOpcao os campos contidos em objBrowseUsuario

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim sNomeTela1 As String

On Error GoTo Erro_BrowseOpcao_Exclui_EmTrans

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 178275
    Next

    sNomeTela1 = String(STRING_NOME_TELA, 0)

    'verifica o BrowseUsuario está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NomeTela FROM BrowseOpcao WHERE NomeTela=? AND Opcao=?", 0, sNomeTela1, sNomeTela, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 178276
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178277

    'se BrowseUsuario estiver cadastrado ==> é uma alteração
    If lErro = SUCESSO Then
        
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 178278
        
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM BrowseOpcao", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 178279
    
    End If

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   BrowseOpcao_Exclui_EmTrans = SUCESSO
   
   Exit Function
   
Erro_BrowseOpcao_Exclui_EmTrans:

   BrowseOpcao_Exclui_EmTrans = gErr

   Select Case gErr

        Case 178275
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 178276, 178277
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAO", gErr)
      
        Case 178278
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_BROWSEOPCAO", gErr)

        Case 178279
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEOPCAO", gErr)
                                                                     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178280)

   End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function BrowseOpcaoCampo_Grava_EmTrans(ByVal sOpcao As String, ByVal colBrowseUsuarioCampo As Collection) As Long
'exclui as informacoes antigas para o usuario x tela em questao e grava as novas
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lComando1 As Long
Dim lErro As Long
Dim sNome As String
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo
    
On Error GoTo Erro_BrowseOpcaoCampo_Grava_EmTrans
    
    Set objBrowseUsuarioCampo = colBrowseUsuarioCampo.Item(1)
    
    lComando1 = 0
    
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 178282
    
    sNome = String(STRING_NOME_TELA, 0)
    
    'tenta fazer o lock da tela em questão
    lErro = Comando_ExecutarLockado(lComando1, "SELECT Nome FROM Telas WHERE Nome=?", sNome, objBrowseUsuarioCampo.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 178283
    
    'le a tela em questão
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 178284
    
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 178285
    
    'exclui os campos desta tela x usuario
    lErro = BrowseOpcaoCampo_Exclui_BD_EmTrans(sOpcao, objBrowseUsuarioCampo.sNomeTela)
    If lErro Then gError 178286

    lErro = BrowseOpcaoCampo_Grava_BD_EmTrans(sOpcao, colBrowseUsuarioCampo)
    If lErro Then gError 178287
    
    lErro = Comando_Fechar(lComando1)
    
    BrowseOpcaoCampo_Grava_EmTrans = SUCESSO

    Exit Function

Erro_BrowseOpcaoCampo_Grava_EmTrans:
    
    BrowseOpcaoCampo_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case 178282
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 178283, 178284
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAS", gErr)
            
        Case 178285
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TELAS", gErr)
                   
        Case 178286, 178287
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178288)
            
    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function BrowseOpcaoCampo_Grava_BD_EmTrans(ByVal sOpcao As String, ByVal colBrowseUsuarioCampo As Collection) As Long
'grava na tabela BrowseOpcaoCampo os campos contidos na coleção
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim objBrowseUsuarioCampo As AdmBrowseUsuarioCampo

On Error GoTo Erro_BrowseOpcaoCampo_Grava_BD_EmTrans

    lComando = 0
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 178289
    
    'Inicia Loop de GRAVACAO
    For Each objBrowseUsuarioCampo In colBrowseUsuarioCampo
            
        'Faz inserção do campo na Tabela
        lErro = Comando_Executar(lComando, "INSERT INTO BrowseOpcaoCampo (NomeTela,Opcao,NomeArq,Nome,PosicaoTela,Titulo, Largura) VALUES (?,?,?,?,?,?,?)", objBrowseUsuarioCampo.sNomeTela, sOpcao, objBrowseUsuarioCampo.sNomeArq, objBrowseUsuarioCampo.sNome, objBrowseUsuarioCampo.iPosicaoTela, objBrowseUsuarioCampo.sTitulo, objBrowseUsuarioCampo.lLargura)
        If lErro <> AD_SQL_SUCESSO Then gError 178290
        
    Next

    Call Comando_Fechar(lComando)
    
    BrowseOpcaoCampo_Grava_BD_EmTrans = SUCESSO
    
    Exit Function
    
Erro_BrowseOpcaoCampo_Grava_BD_EmTrans:

    BrowseOpcaoCampo_Grava_BD_EmTrans = gErr
    
    Select Case gErr
    
        Case 178289
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 178290
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEOPCAOCAMPO", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178291)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function BrowseOpcaoCampo_Exclui_BD_EmTrans(ByVal sOpcao As String, ByVal sNomeTela As String) As Long
'exclui da tabela de BrowseOpcaoCampo os registros referentes a tela x Opcao em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim sNome As String

On Error GoTo Erro_BrowseOpcaoCampo_Exclui_BD_EmTrans

    lComando = 0
    lComando1 = 0

    'Abre comando para pesquisar os campos
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 178292
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 178293
    
    sNome = String(STRING_NOME_CAMPO, 0)
    
    'Pesquisa a existencia dos registros da tela x usuario em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Nome FROM BrowseOpcaoCampo WHERE NomeTela=? AND Opcao=?", 0, sNome, sNomeTela, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 178294
     
    'Lê o primeiro campo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178295

    'Inicia Loop de EXCLUSAO
    Do While lErro <> AD_SQL_SEM_DADOS
        
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 178308
        
        'Exclui o Campo da  Tabela de BrowseOpcaoCampo
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM BrowseOpcaoCampo", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 178296
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178297
    
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    BrowseOpcaoCampo_Exclui_BD_EmTrans = SUCESSO
    
    Exit Function
    
Erro_BrowseOpcaoCampo_Exclui_BD_EmTrans:

    BrowseOpcaoCampo_Exclui_BD_EmTrans = gErr
    
    Select Case gErr
    
        Case 178292, 178293
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178294, 178295, 178297
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAOCAMPO", gErr)
           
        Case 178296
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEOPCAOCAMPO", gErr)
            
        Case 178308
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_BROWSEOPCAOCAMPO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178298)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function BrowseOpcaoOrdenacao_Grava_EmTrans(ByVal sOpcao As String, ByVal objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao) As Long
'grava a informacao de ordenacao para a opcao x tela em questão
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long

On Error GoTo Erro_BrowseOpcaoOrdenacao_Grava_EmTrans

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 178299

    lErro = BrowseOpcaoOrdenacao_Exclui_EmTrans(sOpcao, objBrowseUsuarioOrdenacao.sNomeTela)

    'não encontrou a ordenação ==>  é uma inserção de ordenação
    lErro = Comando_Executar(lComando1, "INSERT INTO BrowseOpcaoOrdenacao (NomeTela, Opcao, Indice, SelecaoSQL1, SelecaoSQL1Usuario, NomeIndice) VALUES (?,?,?,?,?,?)", objBrowseUsuarioOrdenacao.sNomeTela, sOpcao, objBrowseUsuarioOrdenacao.iIndice, objBrowseUsuarioOrdenacao.sSelecaoSQL1, objBrowseUsuarioOrdenacao.sSelecaoSQL1Usuario, objBrowseUsuarioOrdenacao.sNomeIndice)
    If lErro <> AD_SQL_SUCESSO Then gError 178304

    Call Comando_Fechar(lComando1)

    BrowseOpcaoOrdenacao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_BrowseOpcaoOrdenacao_Grava_EmTrans:

    BrowseOpcaoOrdenacao_Grava_EmTrans = gErr
    
    Select Case gErr

        Case 178299
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178304
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEOPCAOORDENACAO", gErr, objBrowseUsuarioOrdenacao.sNomeTela, sOpcao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178305)

    End Select
    
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function BrowseOpcaoOrdenacao_Exclui_EmTrans(ByVal sOpcao As String, ByVal sNomeTela As String) As Long
'exclui a informacao de ordenacao para a opcao x tela em questão
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim sNomeTela1 As String

On Error GoTo Erro_BrowseOpcaoOrdenacao_Exclui_EmTrans

    lComando1 = 0
    lComando2 = 0

    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 178306

    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando2 = 0 Then gError 178300

    sNomeTela1 = String(STRING_NOME_TELA, 0)

    'pesquisa no BD as informacoes de ordenacao
    lErro = Comando_ExecutarPos(lComando1, "SELECT NomeTela FROM BrowseOpcaoOrdenacao WHERE NomeTela=? AND Opcao=?", 0, sNomeTela1, sNomeTela, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 178301

    'le a ordenacao, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178302

    'encontrou a ordenacao ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 178309

        'atualiza a ordenação
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM BrowseOpcaoOrdenacao", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 178303

    End If

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)

    BrowseOpcaoOrdenacao_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_BrowseOpcaoOrdenacao_Exclui_EmTrans:

    BrowseOpcaoOrdenacao_Exclui_EmTrans = gErr
    
    Select Case gErr

        Case 178300, 178306
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178301, 178302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEOPCAOORDENACAO", gErr, sNomeTela, sOpcao)
            
        Case 178303
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEOPCAOORDENACAO", gErr, sNomeTela, sOpcao)
            
        Case 178309
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_BROWSEOPCAOORDENACAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178307)

    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function BrowseOpcao_Grava(ByVal sOpcao As String, ByVal objBrowse As AdmBrowse, ByVal objBrowseUsuarioOrdenacao As AdmBrowseUsuarioOrdenacao, ByVal objBrowseUsuario As AdmBrowseUsuario) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_BrowseOpcao_Grava

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 178312

    If Len(Trim(objBrowseUsuario.sNomeTela)) > 0 Then

        lErro = BrowseOpcao_Grava_EmTrans(sOpcao, objBrowseUsuario)
        If lErro <> SUCESSO Then gError 178313

    End If
    
    lErro = BrowseOpcaoCampo_Grava_EmTrans(sOpcao, objBrowse.colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then gError 178314
    
    lErro = BrowseOpcaoOrdenacao_Grava_EmTrans(sOpcao, objBrowseUsuarioOrdenacao)
    If lErro <> SUCESSO Then gError 178315
    
    lErro = CF("BrowseExcel_Grava", objBrowse, sOpcao, objBrowse.objBrowseExcel)
    If lErro <> SUCESSO Then gError 178315
    
    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 178316
    
    BrowseOpcao_Grava = SUCESSO
    
    Exit Function
    
Erro_BrowseOpcao_Grava:

    BrowseOpcao_Grava = gErr
    
    Select Case gErr

        Case 178312
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 178313 To 178315

        Case 178316
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178317)

    End Select
    
    Call Transacao_RollbackDic
    
    Exit Function
    
End Function

Function BrowseOpcao_Exclui(ByVal sOpcao As String, ByVal sNomeTela As String) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_BrowseOpcao_Exclui

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 178318

    lErro = BrowseOpcao_Exclui_EmTrans(sOpcao, sNomeTela)
    If lErro <> SUCESSO Then gError 178319
    
    lErro = BrowseOpcaoCampo_Exclui_BD_EmTrans(sOpcao, sNomeTela)
    If lErro <> SUCESSO Then gError 178320
    
    lErro = BrowseOpcaoOrdenacao_Exclui_EmTrans(sOpcao, sNomeTela)
    If lErro <> SUCESSO Then gError 178321
    
    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 178322
    
    BrowseOpcao_Exclui = SUCESSO
    
    Exit Function
    
Erro_BrowseOpcao_Exclui:

    BrowseOpcao_Exclui = gErr
    
    Select Case gErr

        Case 178318
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 178319 To 178321

        Case 178322
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178323)

    End Select
    
    Call Transacao_RollbackDic
    
    Exit Function
    
End Function

Function RegraWFW_Atualiza_DataHora(ByVal colRegraWFW As Collection) As Long
'atualiza a data e hora da ultima execucao verdadeira das regras

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objRegraWFW As ClassRegraWFW
Dim sUsuario As String
Dim lTransacao As Long
    
On Error GoTo Erro_RegraWFW_Atualiza_DataHora
    
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 178367

    lErro = RegraWFW_Atualiza_DataHora_Trans(colRegraWFW)
    If lErro <> SUCESSO Then gError 178392
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 178370
    
    RegraWFW_Atualiza_DataHora = SUCESSO
    
    Exit Function

Erro_RegraWFW_Atualiza_DataHora:

    RegraWFW_Atualiza_DataHora = gErr
    
    Select Case gErr
                
        Case 178367
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                  
        Case 178370
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                  
        Case 178392
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178371)
        
    End Select
        
    Call Transacao_Rollback
        
    Exit Function

End Function

Function RegraWFW_Atualiza_DataHora_Trans(ByVal colRegraWFW As Collection) As Long
'atualiza a data e hora da ultima execucao verdadeira das regras

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objRegraWFW As ClassRegraWFW
Dim sUsuario As String
Dim lTransacao As Long
    
On Error GoTo Erro_RegraWFW_Atualiza_DataHora_Trans
    
    If colRegraWFW.Count > 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 178368
        Next
    
        For Each objRegraWFW In colRegraWFW
        
            If objRegraWFW.dRegraRet <> 0 Then
        
                lErro = Comando_Executar(alComando(1), "UPDATE RegraWFW SET DataUltExec = ? , HoraUltExec = ? WHERE Modulo = ? AND Transacao = ? AND Item = ? And CodUsuario = ?", _
                Date, CDbl(Time), objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem, objRegraWFW.sUsuario)
                If lErro <> AD_SQL_SUCESSO Then gError 178369
        
            End If
        
        Next
        
        'Fecha comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
    
    RegraWFW_Atualiza_DataHora_Trans = SUCESSO
    
    Exit Function

Erro_RegraWFW_Atualiza_DataHora_Trans:

    RegraWFW_Atualiza_DataHora_Trans = gErr
    
    Select Case gErr
                
        Case 178368
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178369
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REGRAWFW", gErr, objRegraWFW.sModulo, objRegraWFW.iTransacao, objRegraWFW.iItem, objRegraWFW.sUsuario)
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178371)
        
    End Select
        
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function Config_Obter_Inteiro_Automatico_EmTrans(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, iNumAuto As Integer) As Long
'retorna em iNumAuto(inteiro), o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim iNumeroBD As Integer
Dim sComandoSQL As String
Dim sConteudo As String

On Error GoTo Erro_Config_Obter_Inteiro_Automatico_EmTrans
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 46514
    Next

    sConteudo = String(STRING_BUFFER_MAX_TEXTO, 0)
    'Monta comando SQL de leitura na tabela de configuração
    sComandoSQL = "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?"

    'Pesquisa o número automático que está no BD para giFilialEmpresa
    lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 46516

    'Le o registro com o número automático
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46519
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa = EMPRESA_TODA Then Error 46520
    
    'Se não achar pesquisa o número automático que está no BD para EMPRESA_TODA
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa <> EMPRESA_TODA Then
    
        lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 46518
    
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 46521
        
    End If
    
    'Locka o registro
    lErro = Comando_LockExclusive(lComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 46522

    iNumAuto = CInt(sConteudo)

    'Verifica se número automático ainda não foi utilizado
    'Monta comando SQL leitura na tabela do objeto
    sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampo & " >= ? ORDER BY " & sCampo

    'Pesquisa as numerações do objeto maiores ou iguais ao número automático
    lErro = Comando_Executar(lComando(2), sComandoSQL, iNumeroBD, iNumAuto)
    If lErro <> AD_SQL_SUCESSO Then Error 46517

    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46523

    'pesquisa o proximo número automático disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        Select Case iNumeroBD
            Case Is > iNumAuto
                Exit Do
            Case Is = iNumAuto
                iNumAuto = iNumAuto + 1
        End Select

        lErro = Comando_BuscarProximo(lComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46524

    Loop

    sConteudo = CStr(iNumAuto + 1)
    
    'Monta Comando SQL de atualizar número automático na tabela configuração
    sComandoSQL = "UPDATE " & sTabelaConfig & " SET Conteudo = ?"
    
    'Atualiza o numero automático no BD
    lErro = Comando_ExecutarPos(lComando(3), sComandoSQL, lComando(1), sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 46525

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_Obter_Inteiro_Automatico_EmTrans = SUCESSO

    Exit Function

Erro_Config_Obter_Inteiro_Automatico_EmTrans:

    Config_Obter_Inteiro_Automatico_EmTrans = Err

    Select Case Err

    Case 46514
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 46516, 46517, 46518
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaConfig)

    Case 46519, 46520, 46521
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 46522
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 46523, 46524
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaObjeto)
    
    Case 46525
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142397)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function

Function Config_ObterAutomatico3(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, dNumAuto As Double, Optional ByVal sCampoFilial As String = "FilialEmpresa") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Config_ObterAutomatico3
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 25193
    
    lErro = Config_ObterAutomatico3_EmTrans(sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, dNumAuto, sCampoFilial)
    If lErro <> SUCESSO Then Error 25194
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 25201

    Config_ObterAutomatico3 = SUCESSO

    Exit Function

Erro_Config_ObterAutomatico3:

    Config_ObterAutomatico3 = Err

    Select Case Err

    Case 25194
    
    Case 25193
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

    Case 25201
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142394)

    End Select

    Call Transacao_Rollback
    
End Function

Function Config_ObterAutomatico3_EmTrans(ByVal sTabelaConfig As String, ByVal sCodigo As String, ByVal sTabelaObjeto As String, ByVal sCampo As String, dNumAuto As Double, Optional ByVal sCampoFilial As String = "FilialEmpresa") As Long
'retorna em lNumAuto o número automático disponivel do próximo objeto
'o objeto é identificado por sCodigo
'sTabelaObjeto é a Tabela do objeto
'sTabelaConfig é a Tabela de configuração do módulo
'sCampo é o campo na Tabela do objeto que armazena o número/código do objeto.

Dim lComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim dNumeroBD As Double
Dim sComandoSQL As String
Dim sConteudo As String, iDependeFilial As Integer

On Error GoTo Erro_Config_ObterAutomatico3_EmTrans
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 25192
    Next
    
    sConteudo = String(STRING_BUFFER_MAX_TEXTO, 0)
    'Monta comando SQL de leitura na tabela de configuração
    sComandoSQL = "SELECT Conteudo FROM " & sTabelaConfig & " WHERE Codigo = ? AND FilialEmpresa = ?"

    'Pesquisa o número automático que está no BD para giFilialEmpresa
    lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 25194

    'Le o registro com o número automático
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25195
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa = EMPRESA_TODA Then Error 25202
    
    'se o codigo depende da filial
    iDependeFilial = 0
    
    'Se não achar pesquisa o número automático que está no BD para EMPRESA_TODA
    If lErro = AD_SQL_SEM_DADOS And giFilialEmpresa <> EMPRESA_TODA Then
    
        lErro = Comando_ExecutarPos(lComando(1), sComandoSQL, 0, sConteudo, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then Error 25203
    
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 25204
        
    Else
    
        If giFilialEmpresa <> EMPRESA_TODA Then iDependeFilial = 1
        
    End If
    
    'Locka o registro
    lErro = Comando_LockExclusive(lComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 25196

    dNumAuto = StrParaDbl(sConteudo)

    'Verifica se número automático ainda não foi utilizado
    
    If iDependeFilial = 0 Then
    
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, dNumeroBD, dNumAuto)
        If lErro <> AD_SQL_SUCESSO Then Error 25197
    
    Else
    
        'Monta comando SQL leitura na tabela do objeto
        sComandoSQL = "SELECT " & sCampo & " FROM " & sTabelaObjeto & " WHERE " & sCampoFilial & " = ? AND " & sCampo & " >= ? ORDER BY " & sCampo
    
        'Pesquisa as numerações do objeto maiores ou iguais ao número automático
        lErro = Comando_Executar(lComando(2), sComandoSQL, dNumeroBD, giFilialEmpresa, dNumAuto)
        If lErro <> AD_SQL_SUCESSO Then Error 32212
        
    End If
    
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25198

    'pesquisa o proximo número automático disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        Select Case Round(dNumeroBD)
            Case Is > Round(dNumAuto)
                Exit Do
            Case Is = Round(dNumAuto)
                dNumAuto = dNumAuto + 1
        End Select

        lErro = Comando_BuscarProximo(lComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25199

    Loop

    sConteudo = CStr(Round(dNumAuto + 1))
    
    'Monta Comando SQL de atualizar número automático na tabela configuração
    sComandoSQL = "UPDATE " & sTabelaConfig & " SET Conteudo = ?"
    
    'Atualiza o numero automático no BD
    lErro = Comando_ExecutarPos(lComando(3), sComandoSQL, lComando(1), sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 25200
    
    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next
    
    Config_ObterAutomatico3_EmTrans = SUCESSO

    Exit Function

Erro_Config_ObterAutomatico3_EmTrans:

    Config_ObterAutomatico3_EmTrans = Err

    Select Case Err

    Case 25192
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 25193
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

    Case 25194, 25197, 25203, 32212
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaConfig)

    Case 25195, 25202, 25204
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25196
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25198, 25199
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", Err, sTabelaObjeto)
    
    Case 25200
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELA_CONFIG", Err, sTabelaConfig, sCodigo)

    Case 25201
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142394)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next
    
End Function

Function BrowseExcel_Grava(ByVal objBrowse As AdmBrowse, ByVal sOpcao As String, ByVal objBrowseExcel As AdmBrowseExcel) As Long

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim objBrowseExcelAux As AdmBrowseExcelAux


On Error GoTo Erro_BrowseExcel_Grava

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 202211
    Next
     
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Grafico FROM BrowseExcel WHERE NomeTela = ? AND Opcao = ?", _
    0, iAux, objBrowse.objForm.Name, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 202212
         
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202213
    
    If lErro <> AD_SQL_SEM_DADOS Then
    
        lErro = BrowseExcelCampos_Exclui(alComando(3), alComando(4), objBrowse, sOpcao)
        If lErro <> SUCESSO Then gError 202214
        
        lErro = BrowseExcelFormulas_Exclui(alComando(5), alComando(6), objBrowse, sOpcao)
        If lErro <> SUCESSO Then gError 202215
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE BrowseExcel SET Titulo = ?,TabelaDinamica = ?,Grafico = ?,TipoGrafico = ?, Formato = ?, LocalizacaoCSV = ?, Arquivo = ? ", _
        alComando(1), objBrowseExcel.sTitulo, objBrowseExcel.iTabelaDinamica, objBrowseExcel.iIncluirGrafico, objBrowseExcel.iTipoGrafico, objBrowseExcel.iFormato, objBrowseExcel.sLocalizacaoCSV, objBrowseExcel.sArquivo)
        If lErro <> AD_SQL_SUCESSO Then gError 202216
    
    Else
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO BrowseExcel(NomeTela, Opcao, Titulo,TabelaDinamica,Grafico,TipoGrafico, Formato, LocalizacaoCSV, Arquivo) VALUES (?,?,?,?,?,?,?,?,?)", _
                objBrowse.objForm.Name, sOpcao, objBrowseExcel.sTitulo, objBrowseExcel.iTabelaDinamica, objBrowseExcel.iIncluirGrafico, objBrowseExcel.iTipoGrafico, objBrowseExcel.iFormato, objBrowseExcel.sLocalizacaoCSV, objBrowseExcel.sArquivo)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202217
    
    End If
    
    lErro = BrowseExcelCampos_Insere(alComando(7), objBrowse, sOpcao, objBrowseExcel)
    If lErro <> SUCESSO Then gError 202218
    
    lErro = BrowseExcelFormulas_Insere(alComando(8), objBrowse, sOpcao, objBrowseExcel)
    If lErro <> SUCESSO Then gError 202219
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

   BrowseExcel_Grava = SUCESSO
   
   Exit Function
   
Erro_BrowseExcel_Grava:

   BrowseExcel_Grava = gErr

    Select Case gErr

        Case 202211
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 202212, 202213
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEEXCEL", gErr)
            
        Case 202214, 202215, 202218, 202219
        
        Case 202216
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_BROWSEEXCEL", gErr)
            
        Case 202217
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEEXCEL", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202220)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function BrowseExcelFormulas_Insere(ByVal lComando As Long, ByVal objBrowse As AdmBrowse, ByVal sOpcao As String, ByVal objBrowseExcel As AdmBrowseExcel) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objBrowseExcelAux As AdmBrowseExcelAux

On Error GoTo Erro_BrowseExcelFormulas_Insere

    iIndice = 0
    For Each objBrowseExcelAux In objBrowseExcel.colFormulas
        iIndice = iIndice + 1
        lErro = Comando_Executar(lComando, "INSERT INTO BrowseExcelFormulas(NomeTela, Opcao, Campo, Formula, Seq) VALUES (?,?,?,?,?)", _
                objBrowse.objForm.Name, sOpcao, objBrowseExcelAux.sCampo, objBrowseExcelAux.iFormula, iIndice)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202198

    Next

    BrowseExcelFormulas_Insere = SUCESSO

    Exit Function

Erro_BrowseExcelFormulas_Insere:

    BrowseExcelFormulas_Insere = gErr

    Select Case gErr

        Case 202198
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEEXCEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202199)

    End Select

    Exit Function

End Function

Private Function BrowseExcelCampos_Insere(ByVal lComando As Long, ByVal objBrowse As AdmBrowse, ByVal sOpcao As String, ByVal objBrowseExcel As AdmBrowseExcel) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objBrowseExcelAux As AdmBrowseExcelAux

On Error GoTo Erro_BrowseExcelCampos_Insere

    iIndice = 0
    For Each objBrowseExcelAux In objBrowseExcel.colCampos
        iIndice = iIndice + 1
        lErro = Comando_Executar(lComando, "INSERT INTO BrowseExcelCampos(NomeTela, Opcao, Campo, Formula, Posicao, Seq) VALUES (?,?,?,?,?,?)", _
                objBrowse.objForm.Name, sOpcao, objBrowseExcelAux.sCampo, objBrowseExcelAux.iFormula, objBrowseExcelAux.iPosicao, iIndice)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202200

    Next

    BrowseExcelCampos_Insere = SUCESSO

    Exit Function

Erro_BrowseExcelCampos_Insere:

    BrowseExcelCampos_Insere = gErr

    Select Case gErr

        Case 202200
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BROWSEEXCEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202201)

    End Select

    Exit Function

End Function

Public Function BrowseExcelFormulas_Exclui(ByVal lComando1 As Long, lComando2 As Long, ByVal objBrowse As AdmBrowse, ByVal sOpcao As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_BrowseExcelFormulas_Exclui

    'Le a tabela BrowseExcelFormulas
    lErro = Comando_ExecutarPos(lComando1, "SELECT Seq FROM BrowseExcelFormulas WHERE NomeTela= ? AND Opcao = ? ", _
                0, iAux, objBrowse.objForm.Name, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 202201

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202202

    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM BrowseExcelFormulas", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 202203

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202204

    Loop

    BrowseExcelFormulas_Exclui = SUCESSO

    Exit Function

Erro_BrowseExcelFormulas_Exclui:

    BrowseExcelFormulas_Exclui = gErr

    Select Case gErr

        Case 202201, 202202, 202204
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEEXCEL", gErr)

        Case 202203
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEEXCEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202205)

    End Select

    Exit Function

End Function

Public Function BrowseExcelCampos_Exclui(ByVal lComando1 As Long, lComando2 As Long, ByVal objBrowse As AdmBrowse, ByVal sOpcao As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_BrowseExcelCampos_Exclui

    'Le a tabela BrowseExcelCampos
    lErro = Comando_ExecutarPos(lComando1, "SELECT Seq FROM BrowseExcelCampos WHERE NomeTela= ? AND Opcao = ? ", _
                0, iAux, objBrowse.objForm.Name, sOpcao)
    If lErro <> AD_SQL_SUCESSO Then gError 202206

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202207

    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM BrowseExcelCampos", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 202208

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202209

    Loop

    BrowseExcelCampos_Exclui = SUCESSO

    Exit Function

Erro_BrowseExcelCampos_Exclui:

    BrowseExcelCampos_Exclui = gErr

    Select Case gErr

        Case 202206, 202207, 202209
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEEXCEL", gErr)

        Case 202208
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEEXCEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202210)

    End Select

    Exit Function

End Function

Function Tamanhos_Grava(ByVal sStringCampo As String, iTamanho As Integer) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iTamanhoAux As Integer
Dim lTransacao As Long

On Error GoTo Erro_Tamanhos_Grava

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 202225

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 202226
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT Tamanho FROM Tamanhos WHERE StringCampo = ?", _
    iTamanhoAux, sStringCampo)
    If lErro <> AD_SQL_SUCESSO Then gError 202227
         
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202228
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO Tamanhos(StringCampo, Tamanho) VALUES (?,?)", _
                sStringCampo, iTamanho)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202229
    
    End If
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 202230

   Tamanhos_Grava = SUCESSO
   
   Exit Function
   
Erro_Tamanhos_Grava:

   Tamanhos_Grava = gErr

    Select Case gErr

        Case 202225
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 202226
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 202227 To 202228
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAMANHOS", gErr)
            
        Case 202229
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TAMANHOS", gErr)
    
        Case 202230
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202231)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_RollbackDic
    
    Exit Function

End Function

Public Function SpedCadastro_Exclui(ByVal iFilialEmpresa As Integer, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_SpedCadastro_Exclui
    
    sCodigo = String(STRING_SPEDCADASTRO_CODIGOINST, 0)

    'Le a tabelaSpedCadastro
    lErro = Comando_ExecutarPos(lComando1, "SELECT CodigoInst FROM SpedCadastro WHERE FilialEmpresa= ? ", _
                0, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 202574

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202575

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 202576
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM SpedCadastro", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 202577
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202578
        
    Loop

    SpedCadastro_Exclui = SUCESSO

    Exit Function

Erro_SpedCadastro_Exclui:

    SpedCadastro_Exclui = gErr

    Select Case gErr

        Case 202573
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202574, 202575, 202578
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SPEDCADASTRO", gErr)

        Case 202576
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_SPEDCADASTRO", gErr)

        Case 202577
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SPEDCADASTRO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202579)

    End Select

    Exit Function

End Function

Public Function SpedCadastro_Insere(ByVal objFilial As AdmFiliais, ByVal lComando As Long) As Long

Dim lErro As Long
Dim objFilialAdmCad As AdmFiliaisAdmCad

On Error GoTo Erro_SpedCadastro_Insere

    For Each objFilialAdmCad In objFilial.colAdmCad
   
        objFilialAdmCad.iFilialEmpresa = objFilial.iCodFilial

        lErro = Comando_Executar(lComando, "INSERT INTO SpedCadastro( FilialEmpresa, CodigoInst, CodigoCad) VALUES (?,?,?)", _
                objFilialAdmCad.iFilialEmpresa, objFilialAdmCad.sCodigo, objFilialAdmCad.sInscricao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202580

    Next

    SpedCadastro_Insere = SUCESSO

    Exit Function

Erro_SpedCadastro_Insere:

    SpedCadastro_Insere = gErr

    Select Case gErr

        Case 202580
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SPEDCADASTRO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202581)

    End Select

    Exit Function

End Function

Public Function ContribuinteSubstituto_Exclui(ByVal iFilialEmpresa As Integer, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim sUF As String

On Error GoTo Erro_ContribuinteSubstituto_Exclui
    
    sUF = String(STRING_ESTADO, 0)

    'Le a tabelaContribuinteSubstituto
    lErro = Comando_ExecutarPos(lComando1, "SELECT UF FROM ContribuinteSubstituto WHERE FilialEmpresa= ? ", _
                0, sUF, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 202574

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202575

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 202576
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ContribuinteSubstituto", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 202577
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202578
        
    Loop

    ContribuinteSubstituto_Exclui = SUCESSO

    Exit Function

Erro_ContribuinteSubstituto_Exclui:

    ContribuinteSubstituto_Exclui = gErr

    Select Case gErr

        Case 202573
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202574, 202575, 202578
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRIBUINTESUBSTITUTO", gErr)

        Case 202576
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CONTRIBUINTESUBSTITUTO", gErr)

        Case 202577
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTRIBUINTESUBSTITUTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202579)

    End Select

    Exit Function

End Function

Public Function ContribuinteSubstituto_Insere(ByVal objFilial As AdmFiliais, ByVal lComando As Long) As Long

Dim lErro As Long
Dim objFilialContrib As AdmFiliaisContrib

On Error GoTo Erro_ContribuinteSubstituto_Insere

    For Each objFilialContrib In objFilial.colContrib
   
        objFilialContrib.iFilialEmpresa = objFilial.iCodFilial

        lErro = Comando_Executar(lComando, "INSERT INTO ContribuinteSubstituto( FilialEmpresa, UF, InscricaoEstadual) VALUES (?,?,?)", _
                objFilialContrib.iFilialEmpresa, objFilialContrib.sUF, objFilialContrib.sInscricao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202580

    Next

    ContribuinteSubstituto_Insere = SUCESSO

    Exit Function

Erro_ContribuinteSubstituto_Insere:

    ContribuinteSubstituto_Insere = gErr

    Select Case gErr

        Case 202580
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTRIBUINTESUBSTITUTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202581)

    End Select

    Exit Function

End Function

Function Executa_Rotinas_Inicializacao() As Long

Dim lErro As Long
Dim sAux As String

On Error GoTo Erro_Executa_Rotinas_Inicializacao

    lErro = CF("FiliaisClientes_Desativa_Sem_Compra")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("Config_Le", "AdmConfig", "TELAS_TAMANHO_VARIAVEL", EMPRESA_TODA, sAux)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    giTelaTamanhoVariavel = sAux
    
    Executa_Rotinas_Inicializacao = SUCESSO
   
    Exit Function
   
Erro_Executa_Rotinas_Inicializacao:

    Executa_Rotinas_Inicializacao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208380)

   End Select
    
    Exit Function

End Function

Function Config_Grava(ByVal sTabela As String, ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, ByVal sConteudo As String) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer, iTipo As Integer
Dim sDescricao As String

On Error GoTo Erro_Config_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 208975

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208976
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM " & sTabela & " WHERE Codigo = ? AND FilialEmpresa = ?", 0, iTipo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 208977
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208978
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS And iFilialEmpresa = EMPRESA_TODA Then gError 208979
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        sDescricao = String(STRING_MAXIMO, 0)
    
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Tipo, Descricao FROM " & sTabela & " WHERE Codigo = ? AND FilialEmpresa = ?", 0, iTipo, sDescricao, sCodigo, EMPRESA_TODA)
        If lErro <> AD_SQL_SUCESSO Then gError 208977
        
        'Lê o primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208978
    
        'Se não encontrou o registro
        If lErro = AD_SQL_SEM_DADOS Then gError 208979
    
        lErro = Comando_Executar(alComando(4), "INSERT INTO  " & sTabela & " (Codigo, Descricao, FilialEmpresa, Tipo, Conteudo) VALUES (?,?,?,?,?)", sCodigo, sDescricao, iFilialEmpresa, iTipo, sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 208980
        
    Else
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE  " & sTabela & " SET Conteudo = ?", alComando(1), sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 208980
        
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 208981
          
    Config_Grava = SUCESSO
    
    Exit Function
    
Erro_Config_Grava:
    
    Config_Grava = gErr
    
    Select Case gErr
    
        Case 208975
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 208976
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208977 To 208978
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_" & UCase(sTabela), gErr)

        Case 208979
            Call Rotina_Erro(vbOKOnly, "ERRO_" & UCase(sTabela) & "_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)

        Case 208980
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_" & UCase(sTabela), gErr, sCodigo, iFilialEmpresa)

        Case 208981
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208982)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function BackupConfig_Grava(ByVal objBackupConfig As ClassBackupConfig) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_BackupConfig_Grava

    'Abertura de transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 208998

    lErro = CF("BackupConfig_Grava_EmTrans", objBackupConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 208999

    BackupConfig_Grava = SUCESSO

    Exit Function

Erro_BackupConfig_Grava:

    BackupConfig_Grava = gErr

    Select Case gErr

        Case 208998
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 208999
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209000)

    End Select

    'Desfaz Transação
    Call Transacao_RollbackDic

    Exit Function

End Function

Public Function BackupConfig_Grava_EmTrans(ByVal objBackupConfig As ClassBackupConfig) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dtData As Date

On Error GoTo Erro_BackupConfig_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 208991
    Next

    'Le a tabelaBackupConfig
    lErro = Comando_ExecutarPos(alComando(0), "SELECT DataUltBkp FROM BackupConfig WHERE Codigo= ? ", _
                0, dtData, objBackupConfig.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208992

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208993

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 208994

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE BackupConfig SET Descricao= ? , Habilitado= ? , DataInicio= ? , Hora= ? , " & _
                "RepetirDias= ? , Diretorio= ? , IncluirDataNomeArq= ? , DataProxBkp= ?, Compactar= ? , TransfFTP= ? , FTPURL= ? , FTPUsu= ? , FTPSenha= ? , FTPDir= ?, DirDownload = ?   ", alComando(0), _
                objBackupConfig.sDescricao, objBackupConfig.iHabilitado, objBackupConfig.dtDataInicio, objBackupConfig.dHora, _
                objBackupConfig.iRepetirDias, objBackupConfig.sDiretorio, objBackupConfig.iIncluirDataNomeArq, objBackupConfig.dtDataProxBkp, objBackupConfig.iCompactar, objBackupConfig.iTransfFTP, objBackupConfig.sFTPURL, objBackupConfig.sFTPUsu, objBackupConfig.sFTPSenha, objBackupConfig.sFTPDir, objBackupConfig.sDirDownload)
        If lErro <> AD_SQL_SUCESSO Then gError 208995

    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO BackupConfig( Codigo, Descricao, Habilitado, DataInicio, " & _
                "Hora, RepetirDias, Diretorio, IncluirDataNomeArq, DataUltBkp, " & _
                "DataProxBkp, Compactar, TransfFTP, FTPURL, FTPUsu, FTPSenha, FTPDir, DirDownload) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objBackupConfig.lCodigo, objBackupConfig.sDescricao, objBackupConfig.iHabilitado, objBackupConfig.dtDataInicio, _
                objBackupConfig.dHora, objBackupConfig.iRepetirDias, objBackupConfig.sDiretorio, objBackupConfig.iIncluirDataNomeArq, objBackupConfig.dtDataUltBkp, _
                objBackupConfig.dtDataProxBkp, objBackupConfig.iCompactar, objBackupConfig.iTransfFTP, objBackupConfig.sFTPURL, objBackupConfig.sFTPUsu, objBackupConfig.sFTPSenha, objBackupConfig.sFTPDir, objBackupConfig.sDirDownload)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208996

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    BackupConfig_Grava_EmTrans = SUCESSO

    Exit Function

Erro_BackupConfig_Grava_EmTrans:

    BackupConfig_Grava_EmTrans = gErr

    Select Case gErr

        Case 208991
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208992, 208993
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPCONFIG", gErr)

        Case 208994
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_BACKUPCONFIG", gErr)

        Case 208995
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_BACKUPCONFIG", gErr)

        Case 208996
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BACKUPCONFIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208997)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Backup_Executa() As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long, iLiberado As Integer
Dim iIndice As Integer, iRepetirDias As Integer, lTransacao As Long
Dim dtData As Date, dHora As Double, dHoraAtual As Double
Dim objBackupConfig As ClassBackupConfig, dtProxBkp As Date
Dim sDiretorio As String, iIncluirDataNomeArq As Integer, lCodigo As Long
Dim bExecutar As Boolean, lConexaoExt As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer

On Error GoTo Erro_Backup_Executa

    sDSN = "DSN=SGEDic;UID=admin;PWD=cacareco;"
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoExt = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoExt = 0 Then gError ERRO_SEM_MENSAGEM

    lErro = Backup_Token(BACKUP_TOKEN_ACAO_OBTER, iLiberado, lConexaoExt)
    If lErro = SUCESSO And iLiberado = MARCADO Then
                
        dHoraAtual = CDbl(Time)
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_AbrirExt(lConexaoExt)
            If alComando(iIndice) = 0 Then gError 209010
        Next
        
        sDiretorio = String(STRING_MAXIMO, 0)
        
        'Le os backups a realizar
        lErro = Comando_Executar(alComando(0), "SELECT Codigo, DataProxBkp, Hora, RepetirDias, Diretorio ,IncluirDataNomeArq FROM BackupConfig WHERE Habilitado <> 0 AND DataProxBkp <> ? AND DataProxBkp <= ? ", _
                    lCodigo, dtProxBkp, dHora, iRepetirDias, sDiretorio, iIncluirDataNomeArq, DATA_NULA, Date)
        If lErro <> AD_SQL_SUCESSO Then gError 209011
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209012
        
        bExecutar = False
        Do While lErro <> AD_SQL_SEM_DADOS
        
            If dtProxBkp < Date Or (dtProxBkp = Date And dHora <= dHoraAtual) Then
                bExecutar = True
                Exit Do
            End If
                
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209014
        
        Loop
        
        If bExecutar Then
        
            Set objBackupConfig = New ClassBackupConfig
        
            objBackupConfig.lCodigo = lCodigo
                
            lErro = CF("BackupConfig_Le", objBackupConfig)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Backup_Executa2(objBackupConfig, lConexaoExt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            dtProxBkp = Date + iRepetirDias
    
            lTransacao = Transacao_AbrirExt(lConexaoExt)
            If lTransacao = 0 Then gError 209024
    
            lErro = Comando_Executar(alComando(1), "UPDATE BackupConfig SET DataUltBkp = ?, DataProxBkp = ? WHERE Codigo = ? ", Date, dtProxBkp, lCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 209013
            
            lErro = Transacao_CommitExt(lTransacao)
            If lErro <> AD_SQL_SUCESSO Then gError 209025
            
        End If
                        
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
        lErro = Backup_Token(BACKUP_TOKEN_ACAO_LIBERAR, iLiberado, lConexaoExt)
        If lErro <> SUCESSO Then gError 209023
        
    End If

    lErro = Backup_Trata_Log(lConexaoExt)
    If lErro <> SUCESSO Then gError 209023
    
    Call Conexao_FecharExt(lConexaoExt)
    
    Backup_Executa = SUCESSO

    Exit Function

Erro_Backup_Executa:

    Backup_Executa = gErr

    Select Case gErr

        Case 209010
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209011, 209012, 209014
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPCONFIG", gErr)

        Case 209013
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_BACKUPCONFIG", gErr)
            
        Case 209023
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_TOKEN_LIBERAR", gErr)
            
        Case 209024
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 209025
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209015)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_RollbackExt(lTransacao)
       
    If iLiberado = MARCADO Then
        Call Backup_Token(BACKUP_TOKEN_ACAO_LIBERAR, iLiberado, lConexaoExt)
    End If

    Call Conexao_FecharExt(lConexaoExt)

    Exit Function

End Function

Public Function Backup_Executa_Direto(ByVal lCodigo As Long, Optional ByVal bApenasBkpeRestoreEmpTeste As Boolean = False) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long, iLiberado As Integer
Dim iIndice As Integer, iRepetirDias As Integer, lTransacao As Long
Dim dtData As Date, dHora As Double, dHoraAtual As Double
Dim objBackupConfig As New ClassBackupConfig, dtProxBkp As Date
Dim sDiretorio As String, iIncluirDataNomeArq As Integer
Dim bExecutar As Boolean, lConexaoExt As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer

On Error GoTo Erro_Backup_Executa_Direto

    sDSN = "DSN=SGEDic;UID=admin;PWD=cacareco;"
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoExt = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoExt = 0 Then gError ERRO_SEM_MENSAGEM

    lErro = Backup_Token(BACKUP_TOKEN_ACAO_OBTER, iLiberado, lConexaoExt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
    dHoraAtual = CDbl(Time)

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(lConexaoExt)
        If alComando(iIndice) = 0 Then gError 209010
    Next
    
    sDiretorio = String(STRING_MAXIMO, 0)
    
    'Le os backups a realizar
    lErro = Comando_Executar(alComando(0), "SELECT DataProxBkp, Hora, RepetirDias, Diretorio ,IncluirDataNomeArq FROM BackupConfig WHERE Codigo = ? ", _
                dtProxBkp, dHora, iRepetirDias, sDiretorio, iIncluirDataNomeArq, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 209011

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209012
    
    If lErro <> AD_SQL_SUCESSO Then gError 211711 'Código não cadastrado
    
    objBackupConfig.lCodigo = lCodigo
    
    lErro = CF("BackupConfig_Le", objBackupConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If bApenasBkpeRestoreEmpTeste Then
        objBackupConfig.iTransfFTP = 0
        objBackupConfig.iCompactar = 0
        objBackupConfig.bApenasBkpeRestoreEmpTeste = bApenasBkpeRestoreEmpTeste
    End If

    lErro = Backup_Executa2(objBackupConfig, lConexaoExt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Not bApenasBkpeRestoreEmpTeste Then
        dtProxBkp = Date + iRepetirDias
    
        lTransacao = Transacao_AbrirExt(lConexaoExt)
        If lTransacao = 0 Then gError 209024
    
        lErro = Comando_Executar(alComando(1), "UPDATE BackupConfig SET DataUltBkp = ?, DataProxBkp = ? WHERE Codigo = ? ", Date, dtProxBkp, lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 209013
        
        lErro = Transacao_CommitExt(lTransacao)
        If lErro <> AD_SQL_SUCESSO Then gError 209025
    End If
                    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    lErro = Backup_Token(BACKUP_TOKEN_ACAO_LIBERAR, iLiberado, lConexaoExt)
    If lErro <> SUCESSO Then gError 209023
    
    lErro = Backup_Trata_Log(lConexaoExt, objBackupConfig.bApenasBkpeRestoreEmpTeste)
    If lErro <> SUCESSO Then gError 209023
    
    Call Conexao_FecharExt(lConexaoExt)

    Backup_Executa_Direto = SUCESSO

    Exit Function

Erro_Backup_Executa_Direto:

    Backup_Executa_Direto = gErr

    Select Case gErr

        Case 209010
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209011, 209012, 209014
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPCONFIG", gErr)

        Case 209013
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_BACKUPCONFIG", gErr)
            
        Case 209023
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_TOKEN_LIBERAR", gErr)
            
        Case 209024
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 209025
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 211711
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUPCONFIG_NAO_CADASTRADO", gErr, lCodigo)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209015)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_RollbackExt(lTransacao)
    
    If iLiberado = MARCADO Then
        Call Backup_Token(BACKUP_TOKEN_ACAO_LIBERAR, iLiberado, lConexaoExt)
    End If

    Call Conexao_FecharExt(lConexaoExt)

    Exit Function

End Function

Private Function Arquivo_Exclui(ByVal sNomeArq As String) As Long
On Error GoTo Erro_Arquivo_Exclui
    Call Kill(sNomeArq)
    Exit Function
Erro_Arquivo_Exclui:
    Exit Function
End Function

Private Function Backup_Executa2(ByVal objBackupConfig As ClassBackupConfig, ByVal lConexaoExt As Long) As Long

Dim lErro As Long, lTransacaoDic As Long
Dim alComando(1 To 3) As Long, sComando As String
Dim iIndice As Integer, lEmpresaRestore As Long
Dim sNome As String, sNomeArq As String, sBDTesteParaRestore As String
Dim objFrmBackup As New FrmBackup
Dim sHora As String, sHoraNomeArq As String, dtData As Date, sArq As String
Dim asArquivos(1 To 50, 1 To 4) As String, iQtdArqs As Integer

On Error GoTo Erro_Backup_Executa2

    Call objFrmBackup.Abrir

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(lConexaoExt)
        If alComando(iIndice) = 0 Then gError 209016
    Next
    
    sHora = Format(Now, "HH:MM:SS")
    sHoraNomeArq = Format(Now, "hhmm")
    dtData = Date
    
    sNome = String(STRING_MAXIMO, 0)
     
    If Not objBackupConfig.bApenasBkpeRestoreEmpTeste Then
        lErro = Comando_Executar(alComando(1), "SELECT Nome FROM BDsInfo ORDER BY Nome", sNome)
    Else
        sBDTesteParaRestore = String(STRING_MAXIMO, 0)
        lErro = Comando_Executar(alComando(1), "SELECT Nome, BDTesteParaRestore, EmpresaTesteParaRestore FROM BDsInfo WHERE BDTesteParaRestore <> '' ORDER BY Nome", sNome, sBDTesteParaRestore, lEmpresaRestore)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 209017
         
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209018
    
    If lErro = AD_SQL_SEM_DADOS Then gError 209019
    
    iQtdArqs = 0
    Do While lErro <> AD_SQL_SEM_DADOS
    
        iQtdArqs = iQtdArqs + 1
    
        If objBackupConfig.iIncluirDataNomeArq = MARCADO Then
            sArq = sNome & "_" & Format(dtData, "yyyymmdd") & "_" & sHoraNomeArq & ".bak"
        Else
            sArq = sNome & ".bak"
        End If
        sNomeArq = objBackupConfig.sDiretorio & sArq
        
        Call Arquivo_Exclui(sNomeArq)
    
        If Not objBackupConfig.bApenasBkpeRestoreEmpTeste Then
            sComando = "EXEC Executar_Backup '" & sNome & "','" & sNomeArq & "'"
        Else
            If lEmpresaRestore = 0 Or lEmpresaRestore = glEmpresa Then gError 213171
            
            sComando = "EXEC Executar_Backup_Com_Restore '" & sNome & "','" & sNomeArq & "','" & sBDTesteParaRestore & "'"
        End If
        
        lErro = Comando_Executar(alComando(2), sComando)
        If lErro <> AD_SQL_SUCESSO Then gError 209020
        
        asArquivos(iQtdArqs, 1) = sNome
        asArquivos(iQtdArqs, 2) = sArq
        asArquivos(iQtdArqs, 3) = sBDTesteParaRestore
        asArquivos(iQtdArqs, 4) = lEmpresaRestore
        
'        lErro = Comando_Executar(alComando(3), "INSERT INTO BackupLog (Usuario, Data, Hora, NomeBD, NomeDir, NomeArq, Compactar, TransfFTP, CodBkp) VALUES (?,?,?,?,?,?,?,?,?)", gsUsuario, dtData, sHora, sNome, objBackupConfig.sDiretorio, sArq, objBackupConfig.iCompactar, objBackupConfig.iTransfFTP, objBackupConfig.lCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 209017
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209021
    
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
       
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexaoExt)
    If lTransacaoDic = 0 Then gError 211743
    
    alComando(3) = Comando_AbrirExt(lConexaoExt)
    If alComando(3) = 0 Then gError 209016

    For iIndice = 1 To iQtdArqs
    
        sNome = asArquivos(iIndice, 1)
        sArq = asArquivos(iIndice, 2)
        sBDTesteParaRestore = asArquivos(iIndice, 3)
        lEmpresaRestore = asArquivos(iIndice, 4)

        lErro = Comando_Executar(alComando(3), "INSERT INTO BackupLog (Usuario, Data, Hora, NomeBD, NomeDir, NomeArq, Compactar, TransfFTP, CodBkp, ComRestore, BDRestore, EmpresaRestore, ConfigurarRestore) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", gsUsuario, dtData, sHora, sNome, objBackupConfig.sDiretorio, sArq, objBackupConfig.iCompactar, objBackupConfig.iTransfFTP, objBackupConfig.lCodigo, IIf(objBackupConfig.bApenasBkpeRestoreEmpTeste, 1, 0), sBDTesteParaRestore, lEmpresaRestore, IIf(objBackupConfig.bApenasBkpeRestoreEmpTeste, 1, 0))
        If lErro <> AD_SQL_SUCESSO Then gError 209017
        
    Next
    
    Call Comando_Fechar(alComando(3))
        
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then gError 211754

    Call objFrmBackup.Fechar

    Backup_Executa2 = SUCESSO
   
   Exit Function
   
Erro_Backup_Executa2:

   Backup_Executa2 = gErr

    Select Case gErr

        Case 209016
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 209017, 209018, 209021
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BDSINFO", gErr)
            
        Case 209019
            Call Rotina_Erro(vbOKOnly, "ERRO_BDSINFO_NAO_CADASTRADO", gErr)

        Case 209020
            Call Rotina_Erro(vbOKOnly, "ERRO_BACKUP_BDSINFO", gErr, sComando)

        Case 211743
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211754
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 213171
            Call Rotina_Erro(vbOKOnly, "ERRO_BKP_COM_RESTORE_CONFIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209022)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_RollbackExt(lTransacaoDic)
    
    Call objFrmBackup.Fechar
    
    Exit Function

End Function

Private Function Backup_Token(ByVal iAcao As Integer, iLiberado As Integer, ByVal lConexaoExt As Long) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lTransacao As Long
Dim sConteudo As String
Dim sConteudoNovo As String

On Error GoTo Erro_Backup_Token

    iLiberado = DESMARCADO

    lTransacao = Transacao_AbrirExt(lConexaoExt)
    If lTransacao = 0 Then gError ERRO_SEM_MENSAGEM

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(lConexaoExt)
        If alComando(iIndice) = 0 Then gError ERRO_SEM_MENSAGEM
    Next
       
    sConteudo = String(STRING_CONTROLE_CONTEUDO, 0)

    'Lê a marcação para ver se já existe alguém tentando executar o backup
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM Controle WHERE Codigo = ? ", 0, sConteudo, CONTROLE_TOKEN_BACKUP)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Busca o Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    'Se não achou => Insere o Path do usuário atual
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If iAcao = BACKUP_TOKEN_ACAO_LIBERAR Then
        sConteudoNovo = ""
        iLiberado = MARCADO
    Else
        sConteudoNovo = Format(Date, "dd/mm/yyyy")
        If StrParaDate(sConteudo) <> Date Then iLiberado = MARCADO
    End If
    
    If iLiberado = MARCADO Then
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Controle SET Conteudo = ?", alComando(1), sConteudoNovo)
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'confirma a transação
    lErro = Transacao_CommitExt(lTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Backup_Token = SUCESSO
    
    Exit Function

Erro_Backup_Token:

    Backup_Token = gErr
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_RollbackExt(lTransacao)
   
    Exit Function

End Function

Function UsuConfig_Grava(ByVal sCodUsuario As String, ByVal lCodConfig As String, ByVal sValor As String) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, lAux As Long

On Error GoTo Erro_UsuConfig_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211168

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211169
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodConfig FROM UsuConfigValor WHERE CodUsuario = ? AND CodConfig = ?", 0, lAux, sCodUsuario, lCodConfig)
    If lErro <> AD_SQL_SUCESSO Then gError 211170
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211171
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO UsuConfigValor (CodUsuario, CodConfig, Valor) VALUES (?,?,?)", sCodUsuario, lCodConfig, sValor)
        If lErro <> AD_SQL_SUCESSO Then gError 211172
        
    Else
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE UsuConfigValor SET Valor = ?", alComando(1), sValor)
        If lErro <> AD_SQL_SUCESSO Then gError 211173
        
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211174
          
    UsuConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_UsuConfig_Grava:
    
    UsuConfig_Grava = gErr
    
    Select Case gErr
    
        Case 211168
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211169
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211170, 211171
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA1", gErr, "UsuConfigValor")

        Case 211172
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TABELA1", gErr, "UsuConfigValor")

        Case 211173
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TABELA1", gErr, "UsuConfigValor")

        Case 211174
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211175)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function Usuario_Configura_Aviso_NFe(ByVal sCodUsuario As String, ByVal lCodConfig As Long) As Long

Dim lErro As Long, sValor As String
Dim objRegraWFW As New ClassRegraWFW, colRegraWFW As New Collection
Dim objRegraWFWAux As ClassRegraWFW
Dim vbResult As VbMsgBoxResult, sTexto As String
Dim iSeq As Integer, sMnemonico As String, sMnemonicoAux As String
Dim bAchou As String

On Error GoTo Erro_Usuario_Configura_Aviso_NFe

    'Le a configuração do usuário para verificar se já foi indicado se ele quer ou não receber os avisos
    lErro = CF("UsuConfig_Le", sCodUsuario, lCodConfig, sValor)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se ele ainda não escolheu
    If StrParaInt(sValor) = DESMARCADO Then
    
        Select Case lCodConfig
        
            Case USUCONFIG_AVISO_NFE_SEM_AUTO
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_USUCONFIG_AVISO_NFE_SEM_AUTO")
                sMnemonico = "ListaNFeNaoAuto"
                sMnemonicoAux = "(14)"
                sTexto = "As seguintes NFs não foram autorizadas: "
                
            Case USUCONFIG_AVISO_CANC_NFE_SEM_HOM
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_USUCONFIG_AVISO_CANC_NFE_SEM_HOM")
                sMnemonico = "ListaCancNFeNaoHom"
                sMnemonicoAux = "(14)"
                sTexto = "Os seguintes cancelamentos não foram homologados: "
        
            Case USUCONFIG_AVISO_CANC_NFE_HOM
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_USUCONFIG_AVISO_CANC_NFE_HOM")
                sMnemonico = "ListaCancNFeHom"
                sMnemonicoAux = "(14)"
                sTexto = "As seguintes nfes estão canceladas na receita mas não no Corporator: "
        
        End Select
        
        'Se ele quiser receber o aviso lê as regras existentes e adiciona a nova
        If vbResult = vbYes Then

            objRegraWFW.iTransacao = 9999
            objRegraWFW.sUsuario = sCodUsuario
            objRegraWFW.sModulo = MODULO_ADM
        
            lErro = CF("RegraWFW_Le_Transacao", objRegraWFW, colRegraWFW)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            bAchou = False
            For Each objRegraWFWAux In colRegraWFW
                iSeq = objRegraWFWAux.iItem
                If InStr(1, objRegraWFWAux.sAvisoMsg, sMnemonico) <> 0 Then bAchou = True
                
                lErro = CF("AvisoUsuWFW_Le_Item", objRegraWFWAux)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Next
            If bAchou Then
                'Se já existe um aviso cadastrado informa e não altera as regras, apenas marca que já foi configurado para não exibir a mensagem de novo
                Call Rotina_Aviso(vbOKOnly, "AVISO_REGRAWFW_JA_CONFIGURADA")
            Else
                objRegraWFW.iItem = iSeq + 1
                objRegraWFW.sRegra = "Saida_Sistema E " & sMnemonico & sMnemonicoAux & " <> """" "
                objRegraWFW.sAvisoMsg = """" & sTexto & """" & " + " & sMnemonico & sMnemonicoAux
                objRegraWFW.colUsuarios.Add sCodUsuario
                colRegraWFW.Add objRegraWFW
                
                lErro = CF("RegraWFW_Grava", colRegraWFW, MARCADO, sCodUsuario)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
        End If
        
        'Marca que já houve a configuração escolhida pelo usuário
        lErro = CF("UsuConfig_Grava", sCodUsuario, lCodConfig, CStr(MARCADO))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
          
    Usuario_Configura_Aviso_NFe = SUCESSO
    
    Exit Function
    
Erro_Usuario_Configura_Aviso_NFe:
    
    Usuario_Configura_Aviso_NFe = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211176)
            
    End Select
    
    Exit Function
    
End Function

Function TelaUsuario_Grava(ByVal objTelaUsuario As ClassTelaUsuario) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, lAux As Long
Dim lLargura As Long
Dim lAltura As Long

On Error GoTo Erro_TelaUsuario_Grava
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 210540

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 210541
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Largura, Altura FROM TelaUsuario WHERE CodUsuario = ? AND NomeTela = ?", 0, lLargura, lAltura, objTelaUsuario.sCodUsuario, objTelaUsuario.sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 210542
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210543
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO TelaUsuario (CodUsuario, NomeTela, Topo, Esquerda, Largura, Altura) VALUES (?,?,?,?,?,?)", objTelaUsuario.sCodUsuario, objTelaUsuario.sNomeTela, 0, 0, objTelaUsuario.lLargura, objTelaUsuario.lAltura)
        If lErro <> AD_SQL_SUCESSO Then gError 210544
        
    Else
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TelaUsuario SET Largura = ?, Altura = ?", alComando(1), objTelaUsuario.lLargura, objTelaUsuario.lAltura)
        If lErro <> AD_SQL_SUCESSO Then gError 210545
        
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 210546
    
          
    TelaUsuario_Grava = SUCESSO
    
    Exit Function
    
Erro_TelaUsuario_Grava:
    
    TelaUsuario_Grava = gErr
    
    Select Case gErr
    
        Case 210540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 210541
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 210542, 210543
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELAUSUARIO", gErr, objTelaUsuario.sNomeTela, objTelaUsuario.sCodUsuario)

        Case 210544
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TABELA1", gErr, "TelaUsuario")

        Case 210545
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TABELA1", gErr, "TelaUsuario")

        Case 210546
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210547)
            
    End Select
    
    Call Transacao_RollbackDic
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function Arquivamento_Executa(ByRef objFrmAguarde As ClassFrmAguarde, ByRef objFrmAguardeTela As Object) As Long

Dim iIndice As Integer, lErro As Long
Dim lTransacaoDic As Long
Dim alComandoDic(0 To 100) As Long
Dim lTransacao As Long
Dim alComando(0 To 2000) As Long
Dim sComandoSQL As String, iDic As Integer, iSeq As Integer, asCmdSQL(1 To 20) As String
Dim iProxCmd As Integer, iProxCmdDic As Integer, lComando As Long
Dim colBDs As New Collection, vValor As Variant, sBD As String, iFinaliza As Integer
Dim lCount As Long

On Error GoTo Erro_Arquivamento_Executa

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211742

    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then gError 211743

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 211744
    Next
    
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then gError 211745
    Next
    
    iProxCmdDic = 1
    'Le os bancos que devem fazer o truncate do log
    sBD = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(alComandoDic(iProxCmdDic), "SELECT Nome FROM BDsInfo", sBD)
    If lErro <> AD_SQL_SUCESSO Then gError 211746
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(iProxCmdDic))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211747

    Do While lErro = AD_SQL_SUCESSO

        colBDs.Add sBD

        lErro = Comando_BuscarProximo(alComandoDic(iProxCmdDic))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211748

    Loop
    
    iProxCmdDic = iProxCmdDic + 1
    lErro = Comando_Executar(alComandoDic(iProxCmdDic), "SELECT COUNT(*) FROM Arquivamento", lCount)
    If lErro <> AD_SQL_SUCESSO Then gError 211749
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(iProxCmdDic))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211750
    
    objFrmAguarde.iTotalItens = lCount
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)
    
    'Le os comandos SQLs a serem executados de limpeza/transferência
    For iIndice = 1 To 20
        asCmdSQL(iIndice) = String(STRING_MAXIMO, 0)
    Next
    sComandoSQL = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(alComandoDic(0), "SELECT Seq, SUBSTRING(ComandoSQL,1,250),SUBSTRING(ComandoSQL,251,250),SUBSTRING(ComandoSQL,501,250),SUBSTRING(ComandoSQL,751,250),SUBSTRING(ComandoSQL,1001,250),SUBSTRING(ComandoSQL,1251,250),SUBSTRING(ComandoSQL,1501,250),SUBSTRING(ComandoSQL,1751,250),SUBSTRING(ComandoSQL,2001,250),SUBSTRING(ComandoSQL,2251,250),SUBSTRING(ComandoSQL,2501,250),SUBSTRING(ComandoSQL,2751,250),SUBSTRING(ComandoSQL,3001,250),SUBSTRING(ComandoSQL,3251,250),SUBSTRING(ComandoSQL,3501,250),SUBSTRING(ComandoSQL,3751,250),SUBSTRING(ComandoSQL,4001,250),SUBSTRING(ComandoSQL,4251,250),SUBSTRING(ComandoSQL,4501,250),SUBSTRING(ComandoSQL,4751,250), Dic, Finaliza FROM Arquivamento ORDER BY Seq", _
                                            iSeq, asCmdSQL(1), asCmdSQL(2), asCmdSQL(3), asCmdSQL(4), asCmdSQL(5), asCmdSQL(6), asCmdSQL(7), asCmdSQL(8), asCmdSQL(9), asCmdSQL(10), asCmdSQL(11), asCmdSQL(12), asCmdSQL(13), asCmdSQL(14), asCmdSQL(15), asCmdSQL(16), asCmdSQL(17), asCmdSQL(18), asCmdSQL(19), asCmdSQL(20), iDic, iFinaliza)
    If lErro <> AD_SQL_SUCESSO Then gError 211749
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211750
    
    Do While lErro = AD_SQL_SUCESSO
    
        sComandoSQL = ""
        For iIndice = 1 To 20
            sComandoSQL = sComandoSQL & asCmdSQL(iIndice)
        Next
    
        If iDic = MARCADO Then
            iProxCmdDic = iProxCmdDic + 1
            lComando = alComandoDic(iProxCmdDic)
        Else
            iProxCmd = iProxCmd + 1
            lComando = alComando(iProxCmd)
        End If
        
        'Apaga os registros desnecessários e transfere registros antigos
        lErro = Comando_Executar(lComando, sComandoSQL)
        'If lErro <> AD_SQL_SUCESSO Then gError 211751 'COMENTAR APÓS TESTES
        
        If iFinaliza = MARCADO Then Exit Do
        
        Call objFrmAguardeTela.ProcessouItem
        If objFrmAguarde.iCancelar = MARCADO Then Exit Do
    
        lErro = Comando_BuscarProximo(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211752
    
    Loop

    Call objFrmAguardeTela.Fechar
    If objFrmAguarde.iCancelar = MARCADO Then gError ERRO_SEM_MENSAGEM
    
    'Acerta tabelas customizadas dos clientes (se necessário)
    lErro = CF("Arquivamento_Executa_Cust")
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha o comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 211753
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 211754
    
'    alComandoDic(0) = Comando_AbrirExt(GL_lConexaoDic)
'    If alComandoDic(0) = 0 Then gError 211755
'
'    For Each vValor In colBDs
'        lErro = Comando_Executar(alComandoDic(0), "DBCC SHRINKDATABASE (" & vValor & ", TRUNCATEONLY);")
'        If lErro <> AD_SQL_SUCESSO Then gError 211756 'COMENTAR APÓS TESTES
'    Next
'
'    Call Comando_Fechar(alComandoDic(0))

    Arquivamento_Executa = SUCESSO
   
    Exit Function
   
Erro_Arquivamento_Executa:

    Arquivamento_Executa = gErr

    Select Case gErr

        Case 211742, 211743
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211744, 211745, 211755
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211746 To 211748
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BDSINFO", gErr)

        Case 211749, 211750, 211752
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVAMENTO", gErr)

        Case 211751
            Call Rotina_Erro(vbOKOnly, "ERRO_EXEC_ARQUIVAMENTO", gErr, iSeq, sComandoSQL)

        Case 211753, 211754
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 211756
            Call Rotina_Erro(vbOKOnly, "ERRO_SHRINKDATABASE", gErr, vValor)
         
        Case ERRO_SEM_MENSAGEM
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211757)

    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    Call Transacao_Rollback
    Call Transacao_RollbackDic
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If
    
    Exit Function

End Function

Public Function Arquivamento_Executa_Cust() As Long
    Arquivamento_Executa_Cust = SUCESSO
End Function

Private Function Backup_Trata_Log(ByVal lConexaoExt As Long, Optional ByVal bApenasBkpeRestoreEmpTeste As Boolean = False) As Long

Dim lErro As Long
Dim objFrmBackup As New FrmBackup

On Error GoTo Erro_Backup_Trata_Log

    Call objFrmBackup.Abrir

    If Not bApenasBkpeRestoreEmpTeste Then
    
        lErro = Backup_Trata_Zip(lConexaoExt)
        If lErro <> SUCESSO Then Call Rotina_Aviso(vbOKOnly, "AVISO_ERRO_AO_COMPACTAR_O_BACKUP", lErro)
        
        lErro = Backup_Trata_Upload(lConexaoExt)
        If lErro <> SUCESSO Then Call Rotina_Aviso(vbOKOnly, "AVISO_ERRO_AO_FAZER_O_UPLOAD_DO_BACKUP", lErro)
        
        lErro = Backup_Trata_Limpa_FTP(lConexaoExt)
        If lErro <> SUCESSO Then Call Rotina_Aviso(vbOKOnly, "AVISO_ERRO_AO_APAGAR_BACKUPS_ANTIGOS_FTP", lErro)
     
    Else
    
        lErro = Backup_Trata_Configura_Restore(lConexaoExt)
        If lErro <> SUCESSO Then gError 213172 'Call Rotina_Aviso(vbOKOnly, "AVISO_ERRO_AO_CONFIGURAR_BD_RESTAURADO", lErro)
    
    End If
     
    Call objFrmBackup.Fechar
    
    Backup_Trata_Log = SUCESSO
   
    Exit Function
   
Erro_Backup_Trata_Log:

    Backup_Trata_Log = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
         
        Case 213172
            Call Rotina_Aviso(vbOKOnly, "AVISO_ERRO_AO_CONFIGURAR_BD_RESTAURADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211926)

    End Select
    
    Call objFrmBackup.Fechar
    
    Exit Function
    
End Function

Private Function Backup_Trata_Zip(ByVal lConexaoExt As Long) As Long

Dim iIndice As Integer, lErro As Long
Dim lTransacaoDic As Long
Dim alComandoDic(0 To 2) As Long
Dim sUsuario As String, dtData As Date, sHora As String, sNomeArq As String, sNomeBd As String, sDir As String
Dim bDirTempCriado As Boolean, bZipTempCriado As Boolean, sDirTemp As String, sArqTemp As String
Dim iPos As Integer, dtDataZip As Date, sHoraZip As String

On Error GoTo Erro_Backup_Trata_Zip

    bDirTempCriado = False
    bZipTempCriado = False

    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexaoExt)
    If lTransacaoDic = 0 Then gError 211927

    'Abre o comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(lConexaoExt)
        If alComandoDic(iIndice) = 0 Then gError 211928
    Next
    
    dtDataZip = Date
    sHoraZip = Format(Now, "HH:MM:SS")
    
    sUsuario = String(STRING_MAXIMO, 0)
    sHora = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComandoDic(0), "SELECT DISTINCT Usuario, Data, Hora FROM BackupLog WHERE Compactar = 1 AND NomeArqZip = ''", sUsuario, dtData, sHora)
    If lErro <> AD_SQL_SUCESSO Then gError 211929
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211930
    
    Do While lErro = AD_SQL_SUCESSO
    
        sDirTemp = ""
                
        sNomeArq = String(STRING_MAXIMO, 0)
        sNomeBd = String(STRING_MAXIMO, 0)
        sDir = String(STRING_MAXIMO, 0)
        
        lErro = Comando_ExecutarPos(alComandoDic(1), "SELECT NomeDir, NomeArq, NomeBD FROM BackupLog WHERE Usuario = ? AND Data = ? AND Hora = ? ", 0, sDir, sNomeArq, sNomeBd, sUsuario, dtData, sHora)
        If lErro <> AD_SQL_SUCESSO Then gError 211931
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211932
        
        Do While lErro = AD_SQL_SUCESSO
        
            If sDirTemp = "" Then
            
                sDirTemp = "BKP_CORP_BD_" & Format(dtDataZip, "YYYYMMDD") & "_" & Replace(sHoraZip, ":", "") & "\"
                
                lErro = Zip_Cria_Diretorio(sDir & sDirTemp)
                If lErro <> SUCESSO Then gError 211933
                bDirTempCriado = True

            End If
        
            lErro = Zip_Verifica_Existencia_Arquivo(sDir & sNomeArq)
            If lErro <> SUCESSO Then gError 211934
        
            lErro = Zip_Copia_Arquivo(sDir & sNomeArq, sDir & sDirTemp & sNomeArq)
            If lErro <> SUCESSO Then gError 211935
        
            lErro = Comando_BuscarProximo(alComandoDic(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211936
        
        Loop
                
        lErro = Zip_Compacta(sDir & Replace(sDirTemp, "\", "") & ".zip", sDir & sDirTemp & "*.*")
        If lErro <> SUCESSO Then gError 211937
        bZipTempCriado = True
        
        lErro = Zip_Exclui_Diretorio(sDir & sDirTemp)
        If lErro <> SUCESSO Then gError 211938
        bDirTempCriado = False
        
        lErro = Comando_Executar(alComandoDic(2), "UPDATE BackupLog SET NomeArqZip = ?, DataZip = ?, HoraZip = ? WHERE Usuario = ? AND Data = ? AND Hora = ? ", Replace(sDirTemp, "\", "") & ".zip", dtDataZip, sHoraZip, sUsuario, dtData, sHora)
        If lErro <> AD_SQL_SUCESSO Then gError 211939
    
        lErro = Comando_BuscarProximo(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211940
    
    Loop
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then gError 211941

    Backup_Trata_Zip = SUCESSO
   
    Exit Function
   
Erro_Backup_Trata_Zip:

    Backup_Trata_Zip = gErr

    Select Case gErr

        Case 211927 To 211941
         
        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211942)

    End Select
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    Call Transacao_RollbackExt(lTransacaoDic)
    
    If bZipTempCriado Then Call Zip_Exclui_Arquivo(sDir & sDirTemp & ".zip")
    If bDirTempCriado Then Call Zip_Exclui_Diretorio(sDir & sDirTemp)
    
    Exit Function
    
End Function

Private Function Backup_Trata_Upload(ByVal lConexaoExt As Long) As Long

Dim iIndice As Integer, lErro As Long
Dim lTransacaoDic As Long
Dim alComandoDic(0 To 2) As Long
Dim sNomeZip As String, sNomeArqFTP As String, sNomeDir As String
Dim sFTPURL As String, sFTPUsu As String, sFTPSenha As String, sFTPDir As String
Dim objFrmFTP As New FrmFTP, sUsuario, dtData As Date, sHora As String

On Error GoTo Erro_Backup_Trata_Upload

    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexaoExt)
    If lTransacaoDic = 0 Then gError 211943

    'Abre o comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(lConexaoExt)
        If alComandoDic(iIndice) = 0 Then gError 211944
    Next
    
    sNomeZip = String(STRING_MAXIMO, 0)
    sNomeArqFTP = String(STRING_MAXIMO, 0)
    sFTPURL = String(STRING_MAXIMO, 0)
    sFTPUsu = String(STRING_MAXIMO, 0)
    sFTPSenha = String(STRING_MAXIMO, 0)
    sFTPDir = String(STRING_MAXIMO, 0)
    sNomeDir = String(STRING_MAXIMO, 0)
    sUsuario = String(STRING_MAXIMO, 0)
    sHora = String(STRING_MAXIMO, 0)
        
    'Lê só o último backup zipado, só ele deverá ser transferido (a princípio)
    lErro = Comando_Executar(alComandoDic(0), "SELECT L.Usuario, L.Data, L.Hora, L.NomeDir, L.NomeArqZip, L.NomeArqFTP, B.FTPURL, B.FTPUsu, B.FTPSenha, B.FTPDir FROM BackupLog AS L, BackupConfig AS B WHERE L.CodBkp = Codigo AND L.TransfFTP = 1 AND L.NomeArqZip <> '' ORDER BY L.Data DESC, L.Hora DESC", sUsuario, dtData, sHora, sNomeDir, sNomeZip, sNomeArqFTP, sFTPURL, sFTPUsu, sFTPSenha, sFTPDir)
    If lErro <> AD_SQL_SUCESSO Then gError 211945
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211946
    
    If lErro = AD_SQL_SUCESSO Then
    
        'Se ainda não foi para o ftp
        If Len(Trim(sNomeArqFTP)) = 0 Then
        
            objFrmFTP.Inet1.AccessType = icUseDefault
            objFrmFTP.Inet1.URL = sFTPURL
            objFrmFTP.Inet1.UserName = sFTPUsu
            objFrmFTP.Inet1.Password = sFTPSenha
            
            If right(sFTPDir, 1) <> "/" Then sFTPDir = sFTPDir & "/"
            
            lErro = objFrmFTP.Fazer_Upload(sNomeDir & sNomeZip, sFTPDir & sNomeZip)
            If lErro <> SUCESSO Then gError 211947
     
            lErro = Comando_Executar(alComandoDic(1), "UPDATE BackupLog SET ApagarArqFTP = 1 WHERE NomeArqFTP <> '' AND ArqFTPApagado = 0 ")
            If lErro <> AD_SQL_SUCESSO Then gError 211948
     
            lErro = Comando_Executar(alComandoDic(2), "UPDATE BackupLog SET FTP = ?, NomeArqFTP = ?, DataFTP = ?, HoraFTP = ?, ApagarArqFTP = ? WHERE Usuario = ? AND Data = ? AND Hora = ? ", sFTPURL, sFTPDir & sNomeZip, Date, Format(Now, "HH:MM:SS"), DESMARCADO, sUsuario, dtData, sHora)
            If lErro <> AD_SQL_SUCESSO Then gError 211948
     
        End If
            
    End If
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then gError 211949

    Set objFrmFTP = Nothing

    Backup_Trata_Upload = SUCESSO
   
    Exit Function
   
Erro_Backup_Trata_Upload:

    Backup_Trata_Upload = gErr

    Select Case gErr

        Case 211943 To 211949
         
        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211950)

    End Select
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    Call Transacao_RollbackExt(lTransacaoDic)
    
    Set objFrmFTP = Nothing
    
    Exit Function
    
End Function

Private Function Backup_Trata_Limpa_FTP(ByVal lConexaoExt As Long) As Long

Dim iIndice As Integer, lErro As Long
Dim lTransacaoDic As Long
Dim alComandoDic(0 To 2) As Long
Dim sNomeArqFTP As String, sNomeArqFTPAnt As String
Dim sFTPURL As String, sFTPUsu As String, sFTPSenha As String
Dim objFrmFTP As New FrmFTP, sUsuario, dtData As Date, sHora As String

On Error GoTo Erro_Backup_Trata_Limpa_FTP

    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexaoExt)
    If lTransacaoDic = 0 Then gError 211951

    'Abre o comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(lConexaoExt)
        If alComandoDic(iIndice) = 0 Then gError 211952
    Next
    
    sNomeArqFTP = String(STRING_MAXIMO, 0)
    sFTPURL = String(STRING_MAXIMO, 0)
    sFTPUsu = String(STRING_MAXIMO, 0)
    sFTPSenha = String(STRING_MAXIMO, 0)
    sUsuario = String(STRING_MAXIMO, 0)
    sHora = String(STRING_MAXIMO, 0)
        
    'Lê só os FTPs que devem ser pagados e ainda não o foram
    lErro = Comando_Executar(alComandoDic(0), "SELECT L.Usuario, L.Data, L.Hora, L.NomeArqFTP, B.FTPURL, B.FTPUsu, B.FTPSenha FROM BackupLog AS L, BackupConfig AS B WHERE L.CodBkp = Codigo AND L.ApagarArqFTP = 1 AND L.NomeArqFTP <> '' AND L.ArqFTPApagado = 0 ORDER BY L.NomeArqFTP ", sUsuario, dtData, sHora, sNomeArqFTP, sFTPURL, sFTPUsu, sFTPSenha)
    If lErro <> AD_SQL_SUCESSO Then gError 211953
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211954
    
    Do While lErro = AD_SQL_SUCESSO
    
        If sNomeArqFTPAnt <> sNomeArqFTP Then
        
            sNomeArqFTPAnt = sNomeArqFTP
        
            objFrmFTP.Inet1.AccessType = icUseDefault
            objFrmFTP.Inet1.URL = sFTPURL
            objFrmFTP.Inet1.UserName = sFTPUsu
            objFrmFTP.Inet1.Password = sFTPSenha
            
            lErro = objFrmFTP.Apagar_Arquivo_FTP(sNomeArqFTP)
            If lErro <> SUCESSO Then gError 211955
     
            lErro = Comando_Executar(alComandoDic(1), "UPDATE BackupLog SET ArqFTPApagado = ?, DataDEL = ?, HoraDEL = ? WHERE Usuario = ? AND Data = ? AND Hora = ? ", MARCADO, Date, Format(Now, "HH:MM:SS"), sUsuario, dtData, sHora)
            If lErro <> AD_SQL_SUCESSO Then gError 211956
     
        End If
            
        lErro = Comando_BuscarProximo(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211957
        
    Loop
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then gError 211958

    Set objFrmFTP = Nothing

    Backup_Trata_Limpa_FTP = SUCESSO
   
    Exit Function
   
Erro_Backup_Trata_Limpa_FTP:

    Backup_Trata_Limpa_FTP = gErr

    Select Case gErr

        Case 211951 To 211958
         
        Case Else
            'Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211959)

    End Select
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    Call Transacao_RollbackExt(lTransacaoDic)
    
    Set objFrmFTP = Nothing
    
    Exit Function
    
End Function

Public Function Backup_Trata_Download(ByVal sDir As String) As Long

Dim iIndice As Integer, lErro As Long
Dim lTransacaoDic As Long
Dim alComandoDic(0 To 2) As Long
Dim sNomeArqFTP As String, sNomeZip As String
Dim sFTPURL As String, sFTPUsu As String, sFTPSenha As String
Dim objFrmFTP As New FrmFTP, sUsuario, dtData As Date, sHora As String

On Error GoTo Erro_Backup_Trata_Download

    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then gError 211960

    'Abre o comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then gError 211961
    Next
    
    sNomeArqFTP = String(STRING_MAXIMO, 0)
    sFTPURL = String(STRING_MAXIMO, 0)
    sFTPUsu = String(STRING_MAXIMO, 0)
    sFTPSenha = String(STRING_MAXIMO, 0)
    sUsuario = String(STRING_MAXIMO, 0)
    sHora = String(STRING_MAXIMO, 0)
    sNomeZip = String(STRING_MAXIMO, 0)
        
    'Lê só os FTPs que devem ser pagados e ainda não o foram
    lErro = Comando_Executar(alComandoDic(0), "SELECT L.Usuario, L.Data, L.Hora, L.NomeArqFTP, L.NomeArqZip, B.FTPURL, B.FTPUsu, B.FTPSenha FROM BackupLog AS L, BackupConfig AS B WHERE L.CodBkp = Codigo AND L.NomeArqFTP <> '' AND L.ArqFTPApagado = 0 ORDER BY L.DataFTP DESC ", sUsuario, dtData, sHora, sNomeArqFTP, sNomeZip, sFTPURL, sFTPUsu, sFTPSenha)
    If lErro <> AD_SQL_SUCESSO Then gError 211962
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211963
    
    If lErro <> AD_SQL_SUCESSO Then gError 211964
        
    objFrmFTP.Inet1.AccessType = icUseDefault
    objFrmFTP.Inet1.URL = sFTPURL
    objFrmFTP.Inet1.UserName = sFTPUsu
    objFrmFTP.Inet1.Password = sFTPSenha
    
    lErro = objFrmFTP.Fazer_Download(sNomeArqFTP, sDir & sNomeZip)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_Executar(alComandoDic(1), "UPDATE BackupLog SET NomeDirDownload = ?, DataDownload = ?, HoraDownload = ? WHERE Usuario = ? AND Data = ? AND Hora = ? ", sDir, Date, Format(Now, "HH:MM:SS"), sUsuario, dtData, sHora)
    If lErro <> AD_SQL_SUCESSO Then gError 211965
        
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 211966

    Set objFrmFTP = Nothing

    Backup_Trata_Download = SUCESSO
   
    Exit Function
   
Erro_Backup_Trata_Download:

    Backup_Trata_Download = gErr

    Select Case gErr

        Case 211960
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 211961
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211962, 211963
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPLOG", gErr)
            
        Case 211964
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMARQ_BKP_ENCONTRADO_FTP", gErr)
            
        Case 211965
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_BACKUPLOG", gErr)

        Case 211966
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211967)

    End Select
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    Call Transacao_RollbackDic
    
    Set objFrmFTP = Nothing
    
    Exit Function
    
End Function

Public Function Anexos_Grava(ByVal objAnexos As ClassAnexos) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objAnexosArq As ClassAnexosArq
Dim iIndice As Integer

On Error GoTo Erro_Anexos_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213150
    Next

    lErro = CF("Anexos_Exclui", objAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objAnexosArq In objAnexos.colArq

        lErro = Comando_Executar(alComando(1), "INSERT INTO Anexos( TipoDoc, NumIntDoc, Seq, Arquivo, Descricao) VALUES (?,?,?,?,?)", _
                objAnexos.iTipoDoc, objAnexos.lNumIntDoc, objAnexosArq.iSeq, objAnexosArq.sArquivo, objAnexosArq.sDescricao)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213151

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Anexos_Grava = SUCESSO

    Exit Function

Erro_Anexos_Grava:

    Anexos_Grava = gErr

    Select Case gErr

        Case 213150
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 213151
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ANEXOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213152)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Anexos_Exclui(ByVal objAnexos As ClassAnexos) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Anexos_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213153
    Next

    'Le a tabelaAnexos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Anexos WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, lNumIntDoc, objAnexos.iTipoDoc, objAnexos.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 213154

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213155

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 213156
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Anexos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 213157
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213158
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Anexos_Exclui = SUCESSO

    Exit Function

Erro_Anexos_Exclui:

    Anexos_Exclui = gErr

    Select Case gErr

        Case 213153
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213154, 213155, 213158
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ANEXOS", gErr)

        Case 213156
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_ANEXOS", gErr)

        Case 213157
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ANEXOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213158)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function IdiomaTextos_Grava(ByVal objIdiomaTela As ClassIdiomaTela) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objTexto As ClassIdiomaTabCampoTexto
Dim objTab As ClassIdiomaTab
Dim objCampo As ClassIdiomaTabCampo
Dim sTextoAux As String, iSeq As Integer
Dim sTexto As String

On Error GoTo Erro_IdiomaTextos_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213159
    Next
    
    For Each objTab In objIdiomaTela.colTabelas
    
        If objTab.sChaveDocS <> "" Or objTab.lChaveDocL <> 0 Or objTab.iChaveDocI <> 0 Then
        
            For Each objCampo In objTab.colCampos
            
                'APAGA DADOS DESSA TABELA/CAMPO GRAVADOS
                lErro = Comando_Executar(alComando(0), "DELETE FROM IdiomaTextos WHERE NomeTabela= ? AND NomeCampo= ? AND ChaveDocS = ? AND ChaveDocL = ? AND ChaveDocI = ? ", _
                        objTab.sNomeTabela, objCampo.sNomeCampo, objTab.sChaveDocS, objTab.lChaveDocL, objTab.iChaveDocI)
                If lErro <> AD_SQL_SUCESSO Then gError 213160
        
                For Each objTexto In objCampo.colTextos
                
                    sTextoAux = objTexto.sTexto
                    iSeq = 0
                    Do While sTextoAux <> ""
                    
                        sTexto = left(sTextoAux, 250)
                        iSeq = iSeq + 1
                        If Len(sTextoAux) > 250 Then
                            sTextoAux = Mid(sTextoAux, 251)
                        Else
                            sTextoAux = ""
                        End If
                    
                        lErro = Comando_Executar(alComando(1), "INSERT INTO IdiomaTextos( NomeTabela, NomeCampo, ChaveDocS, ChaveDocL, ChaveDocI, " & _
                                "Idioma, Seq, Texto) VALUES (?,?,?,?,?,?,?,?)", _
                                objTab.sNomeTabela, objCampo.sNomeCampo, objTab.sChaveDocS, objTab.lChaveDocL, objTab.iChaveDocI, _
                                objTexto.iIdioma, iSeq, sTexto)
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213161
                        
                    Loop
        
                Next
                
            Next
            
        End If
        
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    IdiomaTextos_Grava = SUCESSO

    Exit Function

Erro_IdiomaTextos_Grava:

    IdiomaTextos_Grava = gErr

    Select Case gErr

        Case 213159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213160
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_IDIOMATEXTOS", gErr)

        Case 213161
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IDIOMATEXTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213162)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function IdiomaTextos_Exclui(ByVal sNomeTabela As String, ByVal sChaveS As String, ByVal lChaveL As Long, ByVal iChaveI As Integer) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_IdiomaTextos_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213163
    Next

    'Le a tabelaIdiomaTextos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Idioma FROM IdiomaTextos WHERE NomeTabela= ? AND ChaveDocS = ? AND ChaveDocL = ? AND ChaveDocI = ? ", _
                0, iAux, sNomeTabela, sChaveS, lChaveL, iChaveI)
    If lErro <> AD_SQL_SUCESSO Then gError 213164

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213165

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 213166
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM IdiomaTextos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 213167
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213168
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    IdiomaTextos_Exclui = SUCESSO

    Exit Function

Erro_IdiomaTextos_Exclui:

    IdiomaTextos_Exclui = gErr

    Select Case gErr

        Case 213163
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213164, 213165, 213168
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IDIOMATEXTOS", gErr)

        Case 213166
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_IDIOMATEXTOS", gErr)

        Case 213167
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_IDIOMATEXTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213169)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Backup_Trata_Configura_Restore(ByVal lConexaoDicExt As Long) As Long

Dim iIndice As Integer, lErro As Long, sComandoSQL As String
Dim lTransacaoDic As Long, lConexaoExt As Long, lTransacao As Long
Dim alComandoDic(0 To 2) As Long, lComando As Long
Dim lEmpresaRestore As Long, asCmdSQL(1 To 20) As String
Dim sDSN As String, iLenDSN As Integer, sParamOut As String, iLenParamOut As Integer

On Error GoTo Erro_Backup_Trata_Configura_Restore

    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexaoDicExt)
    If lTransacaoDic = 0 Then gError 213173

    'Abre o comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(lConexaoDicExt)
        If alComandoDic(iIndice) = 0 Then gError 213174
    Next
    
    lErro = Comando_Executar(alComandoDic(0), "SELECT DISTINCT EmpresaRestore FROM BackupLog WHERE ConfigurarRestore = 1", lEmpresaRestore)
    If lErro <> AD_SQL_SUCESSO Then gError 213175
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213176
    
    Do While lErro = AD_SQL_SUCESSO
    
        sDSN = String(STRING_MAXIMO, 0)
    
        lErro = Comando_Executar(alComandoDic(1), "SELECT StringConexao FROM Empresas WHERE Codigo = ?", sDSN, lEmpresaRestore)
        If lErro <> AD_SQL_SUCESSO Then gError 213177
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213178
    
        If lErro <> AD_SQL_SUCESSO Then gError 213179
        
        iLenDSN = Len(sDSN)
        sParamOut = String(1024, 0)
        iLenParamOut = 1024
        lConexaoExt = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
        If lConexaoExt = 0 Then gError ERRO_SEM_MENSAGEM
        
        lTransacao = Transacao_AbrirExt(lConexaoExt)
        If lTransacao = 0 Then gError 213180
        
        'Le os comandos SQLs a serem executados de limpeza/transferência
        For iIndice = 1 To 20
            asCmdSQL(iIndice) = String(STRING_MAXIMO, 0)
        Next
        sComandoSQL = ""
        lErro = Comando_Executar(alComandoDic(2), "SELECT SUBSTRING(ComandoSQL,1,250),SUBSTRING(ComandoSQL,251,250),SUBSTRING(ComandoSQL,501,250),SUBSTRING(ComandoSQL,751,250),SUBSTRING(ComandoSQL,1001,250),SUBSTRING(ComandoSQL,1251,250),SUBSTRING(ComandoSQL,1501,250),SUBSTRING(ComandoSQL,1751,250),SUBSTRING(ComandoSQL,2001,250),SUBSTRING(ComandoSQL,2251,250),SUBSTRING(ComandoSQL,2501,250),SUBSTRING(ComandoSQL,2751,250),SUBSTRING(ComandoSQL,3001,250),SUBSTRING(ComandoSQL,3251,250),SUBSTRING(ComandoSQL,3501,250),SUBSTRING(ComandoSQL,3751,250),SUBSTRING(ComandoSQL,4001,250),SUBSTRING(ComandoSQL,4251,250),SUBSTRING(ComandoSQL,4501,250),SUBSTRING(ComandoSQL,4751,250) FROM Scripts WHERE Tipo = ? AND Dic = 0 ORDER BY Seq", _
                                                asCmdSQL(1), asCmdSQL(2), asCmdSQL(3), asCmdSQL(4), asCmdSQL(5), asCmdSQL(6), asCmdSQL(7), asCmdSQL(8), asCmdSQL(9), asCmdSQL(10), asCmdSQL(11), asCmdSQL(12), asCmdSQL(13), asCmdSQL(14), asCmdSQL(15), asCmdSQL(16), asCmdSQL(17), asCmdSQL(18), asCmdSQL(19), asCmdSQL(20), TIPO_SCRIPT_RESTORE_BD_TESTE)
        If lErro <> AD_SQL_SUCESSO Then gError 213181
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213182
        
        Do While lErro = AD_SQL_SUCESSO
        
            sComandoSQL = ""
            For iIndice = 1 To 20
                sComandoSQL = sComandoSQL & asCmdSQL(iIndice)
            Next
            
            lComando = Comando_AbrirExt(lConexaoExt)
            If lComando = 0 Then gError 213183
            
            'Apaga os registros desnecessários e transfere registros antigos
            lErro = Comando_Executar(lComando, sComandoSQL)
            If lErro <> AD_SQL_SUCESSO Then gError 213184
        
            Call Comando_Fechar(lComando)
        
            lErro = Comando_BuscarProximo(alComandoDic(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213185
        
        Loop
        
        lErro = Transacao_CommitExt(lTransacao)
        If lErro <> AD_SQL_SUCESSO Then gError 213186
    
        Call Conexao_FecharExt(lConexaoExt)
        
        lErro = Comando_Executar(alComandoDic(2), "UPDATE BackupLog SET ConfigurarRestore = 0 WHERE EmpresaRestore = ? AND ConfigurarRestore = 1 ", lEmpresaRestore)
        If lErro <> AD_SQL_SUCESSO Then gError 213187
    
        lErro = Comando_BuscarProximo(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213188
    
    Loop
   
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then gError 213189

    Backup_Trata_Configura_Restore = SUCESSO
   
    Exit Function
   
Erro_Backup_Trata_Configura_Restore:

    Backup_Trata_Configura_Restore = gErr

    Select Case gErr

        Case 213173, 213180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213174, 213183
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
         
        Case 213175, 213176, 213188
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BACKUPLOG", gErr)
            
        Case 213177, 213178
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", gErr, lEmpresaRestore)
            
        Case 213179
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_CADASTRADA", gErr, lEmpresaRestore)
        
        Case 213181, 213182, 213185
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SCRIPTS", gErr, TIPO_SCRIPT_RESTORE_BD_TESTE)
        
        Case 213184
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL)
        
        Case 213187
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_BACKUPLOG", gErr)
         
        Case 213186, 213189
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213190)

    End Select
   
    Call Comando_Fechar(lComando)
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    Call Transacao_RollbackExt(lTransacao)
    Call Transacao_RollbackExt(lTransacaoDic)
    
    Call Conexao_FecharExt(lConexaoExt)

    Exit Function
    
End Function

Public Function ConfigOutros_Grava(ByVal colConfigs As Collection) As Long
'Realiza a gravação do Cadastro no BD

Dim lErro As Long
Dim lTransacao As Long
Dim objUsu As New ClassUsuarios

On Error GoTo Erro_ConfigOutros_Grava

    objUsu.sCodUsuario = gsUsuario
    
    lErro = CF("Usuarios_Le", objUsu)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If UCase(objUsu.sCodGrupo) <> UCase("supervisor") Then gError 213396

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213367

    lErro = CF("ConfigOutros_Grava_EmTrans", colConfigs)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213368
    
    ConfigOutros_Grava = SUCESSO
    
    Exit Function
    
Erro_ConfigOutros_Grava:

    ConfigOutros_Grava = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 213367
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 213368
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 213396
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERAR_CONFIGOUTROS_NAO_SUPER", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213369)

    End Select

    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function ConfigOutros_Grava_EmTrans(ByVal colConfigs As Collection) As Long

Dim lErro As Long
Dim objConfig As ClassConfigOutros

On Error GoTo Erro_ConfigOutros_Grava_EmTrans
    
    For Each objConfig In colConfigs
    
        'Se houve alteração
        If UCase(Trim(objConfig.sConteudo)) <> UCase(Trim(objConfig.sConteudoNovo)) Then
        
            'grava na tabela
            lErro = CF("Config_Altera", objConfig.sTabela, objConfig.sCodigo, objConfig.iFilialEmpresa, objConfig.sConteudoNovo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            'acerta o gobj
            If Len(Trim(objConfig.sNomeObj)) > 0 Then
                lErro = CF("Config_Altera_Obj", objConfig)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
        End If
    
    Next
    
    ConfigOutros_Grava_EmTrans = SUCESSO

    Exit Function

Erro_ConfigOutros_Grava_EmTrans:

    ConfigOutros_Grava_EmTrans = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213370)

    End Select
    
End Function

Function Config_Altera(ByVal sTabela As String, ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, ByVal sConteudo As String) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iTipo As Integer

On Error GoTo Erro_Config_Altera
    
    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213371
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM " & sTabela & " WHERE Codigo = ? AND FilialEmpresa = ?", 0, iTipo, sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213372
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213373
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 213374
        
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE  " & sTabela & " SET Conteudo = ?", alComando(1), sConteudo)
    If lErro <> AD_SQL_SUCESSO Then gError 213375
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Config_Altera = SUCESSO
    
    Exit Function
    
Erro_Config_Altera:
    
    Config_Altera = gErr
    
    Select Case gErr
    
        Case 213371
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213372 To 213373
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_" & UCase(sTabela), gErr)

        Case 213374
            Call Rotina_Erro(vbOKOnly, "ERRO_" & UCase(sTabela) & "_INEXISTENTE", gErr, sCodigo, iFilialEmpresa)

        Case 213375
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_" & UCase(sTabela), gErr, sCodigo, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213376)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function Avisos_Grava(ByVal colAvisos As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Avisos_Grava

    'Abertura de transação
    lTransacao = Transacao_AbrirDic()
    If lTransacao = 0 Then gError 213524

    lErro = CF("Avisos_Grava_EmTrans", colAvisos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_CommitDic()
    If lErro <> AD_SQL_SUCESSO Then gError 213525

    Avisos_Grava = SUCESSO

    Exit Function

Erro_Avisos_Grava:

    Avisos_Grava = gErr

    Select Case gErr

        Case 213524
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213525
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213526)

    End Select

    'Desfaz Transação
    Call Transacao_RollbackDic

    Exit Function

End Function

Public Function Avisos_Grava_EmTrans(ByVal colAvisos As Collection) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer, iLidoAnt As Integer
Dim objAviso As ClassAvisos
Dim dtDataLidoAnt As Date, dHoraLidoAnt As Double

On Error GoTo Erro_Avisos_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 213527
    Next
    
    For Each objAviso In colAvisos
    
        'Le a tabelaAvisosUsu
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Lido, DataLido, HoraLido FROM AvisosUsu WHERE CodAviso = ? AND CodUsuario = ? ", _
                    0, iLidoAnt, dtDataLidoAnt, dHoraLidoAnt, objAviso.lCodigo, gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 213528
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213529
    
        'Se existir => UPDATE, senão => INSERT
        If lErro = AD_SQL_SUCESSO Then
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE AvisosUsu SET Lido= ? , Excluido= ?, DataLido = ?, HoraLido = ?, DataExcluido = ?, HoraExcluido = ? ", alComando(0), objAviso.iLido, objAviso.iExcluido, IIf(objAviso.iLido = MARCADO And iLidoAnt = DESMARCADO, Date, dtDataLidoAnt), IIf(objAviso.iLido = MARCADO And iLidoAnt = DESMARCADO, CDbl(Time), dHoraLidoAnt), IIf(objAviso.iExcluido = MARCADO, Date, DATA_NULA), IIf(objAviso.iExcluido = MARCADO, CDbl(Time), 0))
            If lErro <> AD_SQL_SUCESSO Then gError 213530
    
        Else
    
            lErro = Comando_Executar(alComando(2), "INSERT INTO AvisosUsu( CodAviso, CodUsuario, Lido, Excluido, DataAberto, HoraAberto, DataLido, HoraLido, DataExcluido, HoraExcluido) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                    objAviso.lCodigo, gsUsuario, objAviso.iLido, objAviso.iExcluido, Date, CDbl(Time), IIf(objAviso.iLido = MARCADO, Date, DATA_NULA), IIf(objAviso.iLido = MARCADO, CDbl(Time), 0), IIf(objAviso.iExcluido = MARCADO, Date, DATA_NULA), IIf(objAviso.iExcluido = MARCADO, CDbl(Time), 0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213531
    
        End If
        
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Avisos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Avisos_Grava_EmTrans:

    Avisos_Grava_EmTrans = gErr

    Select Case gErr

        Case 213527
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213528 To 213531
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_AVISOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213532)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TelaControlLargUsu_Grava(ByVal sNomeTela As String, ByVal iFator As Integer) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, iFatorAux As Integer

On Error GoTo Erro_TelaControlLargUsu_Grava
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 213707

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then gError 213708
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Fator FROM TelaControlLargUsu WHERE NomeTela = ? AND Usuario = ?", 0, iFatorAux, sNomeTela, gsUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 213709
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213710
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO TelaControlLargUsu (NomeTela, Usuario, Fator) VALUES (?,?,?)", sNomeTela, gsUsuario, iFator)
        If lErro <> AD_SQL_SUCESSO Then gError 213711
        
    Else
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TelaControlLargUsu SET Fator = ?", alComando(1), iFator)
        If lErro <> AD_SQL_SUCESSO Then gError 213712
        
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 213713
          
    TelaControlLargUsu_Grava = SUCESSO
    
    Exit Function
    
Erro_TelaControlLargUsu_Grava:
    
    TelaControlLargUsu_Grava = gErr
    
    Select Case gErr
    
        Case 213707
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213708
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213709, 213710
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "TelaControlLargUsu com NomeTela = " & sNomeTela & " e Usuário = " & gsUsuario)

        Case 213711
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TABELA1", gErr, "TelaControlLargUsu")

        Case 213712
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TABELA1", gErr, "TelaControlLargUsu")

        Case 213713
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213714)
            
    End Select
    
    Call Transacao_RollbackDic
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Sub Trata_Tamanho_Grid(objGridInt As AdmGrid, Optional ByVal iDiminui As Integer = 0, Optional ByVal iFatorAtual As Integer = -1)

Dim lErro As Long
Dim colControles As New AdmColCampoIndiceValor
Dim objControle As AdmCampoIndiceValor
Dim iFator As Integer, objTelaUC As Object
Dim bEhGridBaseTela As Boolean

On Error GoTo Erro_Trata_Tamanho_Grid

    bEhGridBaseTela = False

    If (objGridInt.iFatorTam <> 0 Or iDiminui = DESMARCADO) And (objGridInt.iFatorTam < 20000 Or iDiminui = MARCADO) Then

        'Le os controles que devem ser aumentados
        lErro = CF("TelaControlLarg_Le", objGridInt.objForm.Name, colControles)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Verifica se está sendo chamado do grid base para aumento da tela (ignora nos demais)
        For Each objControle In colControles
            If objControle.vValor = MARCADO And UCase(objControle.sCampo) = UCase(objGridInt.objGrid.Name) Then
                bEhGridBaseTela = True
                Exit For
            End If
        Next
        
        If bEhGridBaseTela Then
            
            'Se vai aumentar pelo + ou -
            If iFatorAtual = -1 Then
                If iDiminui = MARCADO Then
                    iFator = -500
                Else
                    iFator = 500
                End If
                objGridInt.iFatorTam = objGridInt.iFatorTam + iFator
            Else
            'Se vai aumentar pelo que está gravado
                iFator = iFatorAtual - objGridInt.iFatorTam
                objGridInt.iFatorTam = iFatorAtual
            End If
            
            If iFator <> 0 Then
                
                'Ajusta todos os controles
                Set objTelaUC = objGridInt.objForm.objUserControl
                    
                For Each objControle In colControles
                    If objControle.iIndice <> -1 Then
                        objTelaUC.Controls(objControle.sCampo)(objControle.iIndice).Width = objTelaUC.Controls(objControle.sCampo)(objControle.iIndice).Width + iFator
                    Else
                        objTelaUC.Controls(objControle.sCampo).Width = objTelaUC.Controls(objControle.sCampo).Width + iFator
                    End If
                Next
                
                If objTelaUC.Width < objTelaUC.Parent.Width Then objTelaUC.Parent!Tela.Width = objTelaUC.Parent.Width
                
                'Grava a informação para o usuário
                lErro = CF("TelaControlLargUsu_Grava", objGridInt.objForm.Name, objGridInt.iFatorTam)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
            
        End If
        
    End If

    Exit Sub

Erro_Trata_Tamanho_Grid:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213715)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function FilialEmpresa_Replica_PreCadastros_Matriz(ByVal iFilialEmp As Integer) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim sSQL As String

On Error GoTo Erro_FilialEmpresa_Replica_PreCadastros_Matriz

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216166
    Next
    
    'REDES
    sSQL = "INSERT INTO Redes (Codigo,FilialEmpresa,Nome,FilialCli,Cliente,Ativo,DataLog,DirImportacaoExtrato,Bandeira,CodConta,Estabelecimento) "
    sSQL = sSQL & "(SELECT Codigo,?,Nome,0,0,Ativo,DataLog,DirImportacaoExtrato,Bandeira,0,'' FROM Redes WHERE FilialEmpresa = ?) "
   
    lErro = Comando_Executar(alComando(0), sSQL, iFilialEmp, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 216167

    'ADMMEIOPAGTO
    sSQL = "INSERT INTO AdmMeioPagto (Codigo,FilialEmpresa,Nome,Rede,TaxaVista,TaxaParcelado,TipoMeioPagto,DefasagemPagtoVista,ContaCorrenteInterna,GeraTituloRec,Endereco,CodBanco,Agencia,Lote,Ativo,DataLog,ProxCodParcelamento) "
    sSQL = sSQL & "(SELECT Codigo,?,Nome,Rede,TaxaVista,TaxaParcelado,TipoMeioPagto,DefasagemPagtoVista,0,GeraTituloRec,Endereco,CodBanco,Agencia,Lote,Ativo,DataLog,0 FROM AdmMeioPagto WHERE FilialEmpresa = ?) "
   
    lErro = Comando_Executar(alComando(1), sSQL, iFilialEmp, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 216168

    'ADMMEIOPAGTOCONDPAGTO
    sSQL = "INSERT INTO AdmMeioPagtoCondPagto (AdmMeioPagto,FilialEmpresa,Parcelamento,NomeParcelamento,NumParcelas,ParcelasRecebto,Taxa,ValorMinimo,Desconto,JurosParcelamento,Juros,Saldo,Ativo,DataLog,PreDatado) "
    sSQL = sSQL & "(SELECT AdmMeioPagto,?,Parcelamento,NomeParcelamento,NumParcelas,ParcelasRecebto,Taxa,ValorMinimo,Desconto,JurosParcelamento,Juros,0,Ativo,DataLog,PreDatado FROM AdmMeioPagtoCondPagto WHERE FilialEmpresa = ?) "
   
    lErro = Comando_Executar(alComando(2), sSQL, iFilialEmp, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 216169
    
    'ADMMEIOPAGTOPARCELAS
    sSQL = "INSERT INTO AdmMeioPagtoParcelas (AdmMeioPagto,FilialEmpresa,Parcelamento,Parcela,PercRecebimento,IntervaloRecebto) "
    sSQL = sSQL & "(SELECT AdmMeioPagto,?,Parcelamento,Parcela,PercRecebimento,IntervaloRecebto FROM AdmMeioPagtoParcelas WHERE FilialEmpresa = ?) "
   
    lErro = Comando_Executar(alComando(3), sSQL, iFilialEmp, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 216170
    
    'CAIXA
    sSQL = "INSERT INTO Caixa (Codigo,FilialEmpresa,NomeReduzido,Descricao,Status,Teclado,Ativo,POS,SessaoStatus) "
    sSQL = sSQL & "VALUES (1,?,'Caixa 1','Caixa 1',1,1,1,1,2) "
   
    lErro = Comando_Executar(alComando(4), sSQL, iFilialEmp)
    If lErro <> AD_SQL_SUCESSO Then gError 216171
    
    'IMPRESSORAECF
    sSQL = "INSERT INTO ImpressoraECF (Codigo,FilialEmpresa,CodModelo) "
    sSQL = sSQL & "VALUES (1,?,9) "
   
    lErro = Comando_Executar(alComando(5), sSQL, iFilialEmp)
    If lErro <> AD_SQL_SUCESSO Then gError 216172
    
    'ECF
    sSQL = "INSERT INTO ECF (Codigo,FilialEmpresa,Caixa,ImpressoraECF,HorarioVerao,Ativo) "
    sSQL = sSQL & "VALUES (1,?,1,1,0,1) "
   
    lErro = Comando_Executar(alComando(6), sSQL, iFilialEmp)
    If lErro <> AD_SQL_SUCESSO Then gError 216173
    
    'TECLADO
    sSQL = "INSERT INTO Teclado (FilialEmpresa,Codigo, Descricao) "
    sSQL = sSQL & "(SELECT ?,Codigo, Descricao FROM Teclado WHERE FilialEmpresa = ?) "
   
    lErro = Comando_Executar(alComando(7), sSQL, iFilialEmp, FILIAL_MATRIZ)
    If lErro <> AD_SQL_SUCESSO Then gError 216174
    
    'Fecha Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    FilialEmpresa_Replica_PreCadastros_Matriz = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresa_Replica_PreCadastros_Matriz:

    FilialEmpresa_Replica_PreCadastros_Matriz = gErr
    
    Select Case gErr
    
        Case 216166
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 216167 To 216174 'ERRO_EXECUCAO_COMANDO_SQL
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216175)
    
    End Select
        
    'Fecha Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

