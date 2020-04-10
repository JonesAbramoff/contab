VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ServicoItemServico_Grava(ByVal colServItemServ As Collection) As Long
'IMPORTANTE: Esta função exclui as relacoes já existentes no BD

Dim lErro As Long
Dim lComando As Long
Dim lTransacao As Long
Dim iCodigo As Integer
Dim objProduto As New ClassProduto
Dim objItemServico As New ClassItemServico
Dim iIndice As Integer
Dim objServItemServ As New ClassServItemServ

On Error GoTo Erro_ServicoItemServico_Grava

    'Abrindo Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97512

    'Abrindo Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97513
    
    'Le e Faz o Lock do Produto a partir do código
    objProduto.sCodigo = colServItemServ.Item(1).sProduto
    
    lErro = CF("Produto_Lock", objProduto)
    If lErro <> SUCESSO Then gError 97515
    
    For iIndice = 1 To colServItemServ.Count
    
        'Le e faz o lock dos itens de servicos associados ao Servico
        objItemServico.iCodigo = colServItemServ.Item(iIndice).iCodItemServico
        lErro = CF("ItemServico_Lock", objItemServico)
        If lErro <> SUCESSO Then gError 97522
        
    Next
               
    'Exclui todas as associacoes de Servico X Item de Servico
    objServItemServ.sProduto = objProduto.sCodigo
    
    lErro = ServicoItemServico_Exclui_EmTrans(objServItemServ)
    If lErro <> SUCESSO And lErro <> 97532 Then gError 97536
               
    'Inclui na Tabela ServicoItemServico a nova relacao de Servico X Item de Servico
    For iIndice = 1 To colServItemServ.Count
    
        lErro = Comando_Executar(lComando, "INSERT INTO ServicoItemServico (Produto, CodItemServico, Ordem) VALUES (?,?,?)", colServItemServ.Item(iIndice).sProduto, colServItemServ.Item(iIndice).iCodItemServico, colServItemServ.Item(iIndice).iOrdem)
        If lErro <> AD_SQL_SUCESSO Then gError 97537
    
    Next
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97516

    'fechando comando
    Call Comando_Fechar(lComando)
    
    ServicoItemServico_Grava = SUCESSO
    
    Exit Function
        
Erro_ServicoItemServico_Grava:
        
    ServicoItemServico_Grava = gErr
        
    Select Case gErr
    
        Case 97512
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 97513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97515, 97522, 97536
        
        Case 97516
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case 97537
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SERVICOITEMSERVICO", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Volta os dados anteriores a transacao
    Call Transacao_Rollback
    
    'fechando comandos
    Call Comando_Fechar(lComando)
       
    Exit Function

End Function

Function ServicoItemServico_Exclui(ByVal objServItemServ As ClassServItemServ) As Long
'Exclui a Relacao de Servico X Item de servico a partir do codigo passado no obj

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ServicoItemServico_Exclui
    
    'Abrindo Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97527

    'Faz a exclusao da associacao Servico X Item de Servico
    lErro = ServicoItemServico_Exclui_EmTrans(objServItemServ)
    If lErro <> SUCESSO And lErro <> 97532 Then gError 97529
    
    'Se nao achou => erro
    If lErro = 97532 Then gError 97557
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97528

    ServicoItemServico_Exclui = SUCESSO
    
    Exit Function

Erro_ServicoItemServico_Exclui:

    ServicoItemServico_Exclui = gErr
    
    Select Case gErr
    
        Case 97527
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 97528
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 97529
        
        Case 97557
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICOITEMSERVICO_NAO_ENCONTRADO", gErr, objServItemServ.sProduto)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Volta os dados anteriores a transacao
    Call Transacao_Rollback
    
End Function

Function ServicoItemServico_Exclui_EmTrans(ByVal objServItemServ As ClassServItemServ) As Long
'IMPORTANTE: TEM QUE ESTAR EM TRANSACAO

Dim alComando(1) As Long
Dim lErro As Long
Dim sSelect As String
Dim sDelete As String
Dim tServItemServ As typeServItemServ
Dim iIndice As Integer

On Error GoTo Erro_ServicoItemServico_Exclui_EmTrans

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97513
    Next
    
    'Expressoes para acesso ao BD
    sSelect = "SELECT Produto, CodItemServico, Ordem FROM ServicoItemServico WHERE Produto = ?"
    sDelete = "DELETE FROM ServicoItemServico"
    
    'Inicializa a string para guardar o servico
    tServItemServ.sProduto = String(STRING_PRODUTO, 0)
    
    'Lê na tabela ServicoItemServico o servico com código passado em objSerItemServ
    lErro = Comando_ExecutarPos(alComando(0), sSelect, 0, tServItemServ.sProduto, tServItemServ.iCodItemServico, tServItemServ.iOrdem, objServItemServ.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 97530
        
    'Buscar o Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97531

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97532
    
    'Para cada item de servico relacionado ao servico
    Do While lErro = AD_SQL_SUCESSO
    
        'Faz o lock exclusivo do Servico
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 97533
                        
        'Exclui a relação do servico com o item de servico
        lErro = Comando_ExecutarPos(alComando(1), sDelete, alComando(0))
        If lErro <> SUCESSO Then gError 97534
        
        'Buscar o Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97535

    Loop
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ServicoItemServico_Exclui_EmTrans = SUCESSO
    
    Exit Function
    
Erro_ServicoItemServico_Exclui_EmTrans:

    ServicoItemServico_Exclui_EmTrans = gErr

    Select Case gErr
    
        Case 97513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97530, 97531, 97535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERVICOITEMSERVICO", gErr)
        
        Case 97532 'Erro tratado na rotina chamadora
                    
        Case 97533
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERVICOITEMSERVICO", gErr)
            
        Case 97534
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SERVICOITEMSERVICO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

Function ItemServico_Grava(objItemServico As ClassItemServico) As Long

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim iCod As Integer

On Error GoTo Erro_ItemServico_Grava

    'Abrindo Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97037

    'Abrindo Comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 97038
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 97039
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 98342
    
    'Verifica no Banco de dados se Já existe Item de Serviço com o Código diferente e Descricao iqual
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM ItemServico WHERE Codigo<>? and Descricao = ?", 0, iCod, objItemServico.iCodigo, objItemServico.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 98343

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98344

    'Se Existe... ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 98345
    
    'Pesquisa no BD o item
    lErro = Comando_ExecutarPos(lComando2, "SELECT codigo FROM ItemServico WHERE codigo = ? ", 0, iCod, objItemServico.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97042
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97043

    If lErro = AD_SQL_SEM_DADOS Then
        'Inclusao --> nao encontrou
        lErro = Comando_Executar(lComando3, "INSERT INTO ItemServico (Codigo,Descricao,DataInicioFim) values (?,?,?)", objItemServico.iCodigo, objItemServico.sDescricao, objItemServico.iTipoData)
        If lErro <> AD_SQL_SUCESSO Then gError 97044
    Else
        'Alteracao --> encontrou
        lErro = Comando_ExecutarPos(lComando3, "UPDATE ItemServico Set Descricao = ?, DataInicioFim = ?", lComando2, objItemServico.sDescricao, objItemServico.iTipoData)
        If lErro <> AD_SQL_SUCESSO Then gError 97045
    End If

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97046

    'fechando comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    ItemServico_Grava = SUCESSO
    
    Exit Function
        
Erro_ItemServico_Grava:
        
    ItemServico_Grava = gErr
        
    Select Case gErr
    
        Case 97037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 97038, 97039, 98342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97042, 97043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO", gErr, objItemServico.iCodigo)
        
        Case 97044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMSERVICO", gErr, objItemServico.iCodigo)
            
        Case 97045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITEMSERVICO", gErr, objItemServico.iCodigo)
    
        Case 97046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 98343, 98344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO1", gErr)
        
        Case 98345
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_JA_EXISTENTE1", gErr, iCod, objItemServico.sDescricao)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Call Transacao_Rollback
    
    'fechando conexoes
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function ItemServico_Exclui(objItemServico As ClassItemServico) As Long

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lTransacao As Long
Dim iCod As Integer

On Error GoTo Erro_ItemServico_Exclui

    'Abrindo Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97055

    'Abrindo Comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 97056
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 97057
        
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 98329
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then gError 98330
    
    'selecionar item no bd
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM ItemServico WHERE Codigo =? ", 0, iCod, objItemServico.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 97058

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97059
    
    'se nao estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 97060

    'Lock do Item
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 97061
    
    'Verifica se o Item de Serviço está associado a algum ServiçoItemServiço
    lErro = Comando_Executar(lComando2, "SELECT CodItemServico FROM ServicoItemServico WHERE CodItemServico= ?", iCod, objItemServico.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98331

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98332
    
    'Se estiver --> Erro.
    If lErro = AD_SQL_SUCESSO Then gError 98333
        
    'Verifica se o Item de Serviço está associado a algum Comprovante
    lErro = Comando_Executar(lComando3, "SELECT CodItemServico FROM CompServItemGR WHERE CodItemServico= ?", iCod, objItemServico.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98334

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98335
    
    'Se estiver --> Erro.
    If lErro = AD_SQL_SUCESSO Then gError 98336
    
    'Exclui o item
    lErro = Comando_ExecutarPos(lComando4, "DELETE FROM ItemServico", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 97069
        
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97068
    
    'fechando conexoes
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    ItemServico_Exclui = SUCESSO
    
    Exit Function

Erro_ItemServico_Exclui:
        
    ItemServico_Exclui = gErr
        
    Select Case gErr
    
        Case 97055
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 97056, 97057, 98329, 98330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97058, 97059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMSERVICO", gErr, objItemServico.iCodigo)
        
        Case 97060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_NAO_CADASTRADO", gErr, objItemServico.iCodigo)
            
        Case 97061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMSERVICO", gErr, objItemServico.iCodigo)
        
        Case 97068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                      
        Case 97069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMSERVICO", gErr, objItemServico.iCodigo)
        
        Case 98331, 98332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERVICOITEMSERVICO", gErr)
        
        Case 98333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_ASSOCIADO_SERVICOTEMSERVICO", gErr, iCod)
        
        Case 98334, 98335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
        
        Case 98336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMSERVICO_ASSOCIADO_COMPSERV", gErr, iCod)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Call Transacao_Rollback
    
    'fechando conexoes
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
       
    Exit Function

End Function

Function TipoEmbalagem_Grava(objTipoEmbalagem As ClassTipoEmbalagem) As Long
'Grava o Tipo de Carga no banco de dados

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iTipo As Integer
Dim sDescricao As String
Dim iIndice As Integer

On Error GoTo Erro_TipoEmbalagem_Grava
    
    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96551
    Next
                
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 96520
                
    'Verifica no Banco de dados se Já existe Tipo de Embalagem com o Código diferente e Descricao iqual
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Tipo FROM TipoEmbalagem WHERE Tipo<>? and Descricao = ?", 0, iTipo, objTipoEmbalagem.iTipo, objTipoEmbalagem.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 98377

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98378

    'Se Existe... ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 98379
    
    sDescricao = String(STRING_TIPOEMBALAGEM_DESCRICAO, 0)
    
    'Verifica se o Código do Tipo de Embalagem já existe
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM TipoEmbalagem WHERE Tipo = ?", 0, iTipo, objTipoEmbalagem.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 96521
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96522
    
    'Se não existe o código...
    If lErro = AD_SQL_SEM_DADOS Then
            
        'Insere no BD um Tipo de Embalagem novo
        lErro = Comando_Executar(alComando(2), "INSERT INTO TipoEmbalagem ( Tipo, Descricao) VALUES (?,?)", objTipoEmbalagem.iTipo, objTipoEmbalagem.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 96523
    
    'Se já existe o código no BD...
    Else
        
        'Altera um Tipo de Embalagem já existente no BD
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TipoEmbalagem SET Tipo = ? , Descricao = ? ", alComando(1), objTipoEmbalagem.iTipo, objTipoEmbalagem.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 96524

    End If

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96525
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoEmbalagem_Grava = SUCESSO
          
    Exit Function
    
Erro_TipoEmbalagem_Grava:

    TipoEmbalagem_Grava = gErr
    
    Select Case gErr

        Case 96520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 96521, 96522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
                    
        Case 96523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
            
        Case 96524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
        
        Case 96525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 96551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98377, 98378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM1", gErr)
                    
        Case 98379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_JA_EXISTENTE2", gErr, iTipo, objTipoEmbalagem.sDescricao)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoEmbalagem_Exclui(objTipoEmbalagem As ClassTipoEmbalagem) As Long
'Exclui o Tipo Carga

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim lTransacao As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim lNumero As Long
Dim lCodigo As Long

On Error GoTo Erro_TipoEmbalagem_Exclui
    
    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96553

    Next
    
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 96531
    
    'Verifica se existe na tabela TipoEmbalagem o Tipo com o código passado
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM TipoEmbalagem WHERE Tipo = ?", 0, iTipo, objTipoEmbalagem.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 96532
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96533
    
    If lErro = AD_SQL_SEM_DADOS Then gError 96661
    
    'Faz lock na tabela TipoEmbalagem
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96534
    
    'Verifica se existe na tabela Cotacao o Tipo com o código passado
    lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM CotacaoGR WHERE TipoEmbalagem = ?", lCodigo, objTipoEmbalagem.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 96535

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96536

    'Se existir, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 96545

    'Verifica se existe na tabela SolicitacaoServico o Tipo com o código passado
    lErro = Comando_Executar(alComando(3), "SELECT Numero FROM SolicitacaoServico WHERE TipoEmbalagem = ?", lNumero, objTipoEmbalagem.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 96537

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96538

    'Se existir, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 96539
             
    'Exclui o Tipo de Carga com o Código passado
    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM TipoEmbalagem", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96541

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96542

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TipoEmbalagem_Exclui = SUCESSO
          
    Exit Function
    
Erro_TipoEmbalagem_Exclui:

    TipoEmbalagem_Exclui = gErr
    
    Select Case gErr

        Case 96531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96532, 96533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
            
        Case 96534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
        
        Case 96535, 96536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO2", gErr)
            
        Case 96537, 96538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO1", gErr)
            
        Case 96539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOEMBALAGEM_ASSOCIADO_SOLICITACAOSERVICO", gErr, lNumero)
        
        Case 96541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOEMBALAGEM", gErr, objTipoEmbalagem.iTipo)
        
        Case 96542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 96545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOEMBALAGEM_ASSOCIADO_COTACAO", gErr, lCodigo)
                        
        Case 96553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 96661
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objTipoEmbalagem.iTipo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TipoContainer_Grava(objTipoContainer As ClassTipoContainer) As Long

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim iTipo As Integer

On Error GoTo Erro_TipoContainer_Grava

    'Abrindo Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97065

    'Abrindo Comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 97066
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 97067
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 98341
    
    'Verifica no Banco de dados se Já existe TipoContainer com o Código diferente e Descricao iqual
    lErro = Comando_ExecutarPos(lComando1, "SELECT Tipo FROM TipoContainer WHERE Tipo<>? and Descricao = ?", 0, iTipo, objTipoContainer.iTipo, objTipoContainer.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 98338

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98339

    'Se Existe... ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 98340
    
    'Pesquisa no BD o tipo
    lErro = Comando_ExecutarPos(lComando2, "SELECT tipo FROM TipoContainer WHERE tipo = ? ", 0, iTipo, objTipoContainer.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 97068
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97069

    If lErro = AD_SQL_SEM_DADOS Then
        'Inclusao --> nao encontrou
        lErro = Comando_Executar(lComando3, "INSERT INTO TipoContainer (Tipo,Descricao,Valor,ISO) values (?,?,?,?)", objTipoContainer.iTipo, objTipoContainer.sDescricao, objTipoContainer.dValor, objTipoContainer.dIso)
        If lErro <> AD_SQL_SUCESSO Then gError 97070
    Else
        'Alteracao --> encontrou
        lErro = Comando_ExecutarPos(lComando3, "UPDATE TipoContainer Set Descricao = ?, Valor = ?,ISO=?", lComando2, objTipoContainer.sDescricao, objTipoContainer.dValor, objTipoContainer.dIso)
        If lErro <> AD_SQL_SUCESSO Then gError 97071
    End If

    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97072

    'fechando comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    TipoContainer_Grava = SUCESSO
    
    Exit Function
        
Erro_TipoContainer_Grava:
        
    TipoContainer_Grava = gErr
        
    Select Case gErr
    
        Case 97065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 97066, 97067, 98341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97068, 97069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER", gErr, objTipoContainer.iTipo)
        
        Case 97070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOCONTAINER", gErr, objTipoContainer.iTipo)
            
        Case 97071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPOCONTAINER", gErr, objTipoContainer.iTipo)
    
        Case 97072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 98338, 98339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER1", gErr)
            
        Case 98340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_JA_EXISTENTE", gErr, iTipo, objTipoContainer.sDescricao)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Call Transacao_Rollback
    
    'fechando conexoes
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function TipoContainer_Exclui(objTipoContainer As ClassTipoContainer) As Long

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim iCod As Integer

On Error GoTo Erro_TipoContainer_Exclui

    'Abrindo Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97074

    'Abrindo Comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 97075
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 97076
    
    'selecionar item no bd
    lErro = Comando_ExecutarPos(lComando1, "SELECT Tipo FROM TipoContainer WHERE Tipo =? ", 0, iCod, objTipoContainer.iTipo)
    If lErro <> AD_SQL_SUCESSO Then gError 97077

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97078
    
    'se nao estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 97079

    'Lock do Item
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 97080
    
    'VER SE EXISTE FOREIGN KEYS REFERENCIANDO O REGISTRO EM QUESTAO
    'SE HOUVER, ERRO!
        
    'Exclui o item
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TipoContainer", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 97081
        
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97082
    
    'fechando conexoes
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    TipoContainer_Exclui = SUCESSO
    
    Exit Function

Erro_TipoContainer_Exclui:
        
    TipoContainer_Exclui = gErr
        
    Select Case gErr
    
        Case 97074
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 97075, 97076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97077, 97078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCONTAINER", gErr, objTipoContainer.iTipo)
        
        Case 97079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCONTAINER_NAO_CADASTRADO", gErr, objTipoContainer.iTipo)
            
        Case 97080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOCONTAINER", gErr, objTipoContainer.iTipo)
        
        Case 97081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                      
        Case 97082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOCONTAINER", gErr, objTipoContainer.iTipo)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Call Transacao_Rollback
    
    'fechando conexoes
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
       
    Exit Function

End Function

Function OrigemDestino_Grava(objOrigemDestino As ClassOrigemDestino) As Long
'Grava a OrigemDestino no banco de dados

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_OrigemDestino_Grava

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96582
    
    Next
    
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 96583

    'Verifica se os campos OrigemDestino com código diferente já existem no BD
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM OrigemDestino WHERE Codigo <> ? AND (OrigemDestino = ? AND UF = ?)", 0, iCodigo, objOrigemDestino.iCodigo, objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 98384

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98385
    
    'Se existe --> erro.
    If lErro = AD_SQL_SUCESSO Then gError 96586
    
    'Verifica se os campos OrigemDestino e Código já existem no BD
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM OrigemDestino WHERE Codigo = ? OR (OrigemDestino = ? AND UF = ?)", 0, iCodigo, objOrigemDestino.iCodigo, objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 96584

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96585
    
    'Se existe...
    If lErro = AD_SQL_SUCESSO Then
        
        'Altera uma OrigemDestino já existente no BD
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE OrigemDestino SET OrigemDestino = ? ,UF = ?", alComando(2), objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
        If lErro <> AD_SQL_SUCESSO Then gError 96587
                                   
    Else

        'Insere no BD uma OrigemDestino nova
        lErro = Comando_Executar(alComando(3), "INSERT INTO OrigemDestino ( Codigo,OrigemDestino, UF) VALUES (?,?,?)", objOrigemDestino.iCodigo, objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)
        If lErro <> AD_SQL_SUCESSO Then gError 96588

    End If

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96589
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    OrigemDestino_Grava = SUCESSO

    Exit Function

Erro_OrigemDestino_Grava:

    OrigemDestino_Grava = gErr

    Select Case gErr
        
        Case 96582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96584, 96585, 98384, 98385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)
       
        Case 96586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEMDESTINO_JA_EXISTENTE", gErr, objOrigemDestino.sOrigemDestino, objOrigemDestino.sUF)

        Case 96587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)

        Case 96588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)
            
        Case 96589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
    
         Call Comando_Fechar(alComando(iIndice))
    
    Next

    Exit Function

End Function

Function OrigemDestino_Exclui(objOrigemDestino As ClassOrigemDestino) As Long
'Exclui a OrigemDestino

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim sOrigemDestino As String
Dim lCodigo As Long
Dim dtData As Date

On Error GoTo Erro_OrigemDestino_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96597
    Next

    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 96598
    
    'Inicializa a variável
    sOrigemDestino = String(STRING_ORIGEMDESTINO_ORIGEMDESTINO, 0)

    'Verifica se existe na tabela OrigemDestino com o código passado
    lErro = Comando_ExecutarPos(alComando(1), "SELECT OrigemDestino FROM OrigemDestino WHERE Codigo = ?", 0, sOrigemDestino, objOrigemDestino.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96599

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96600
    
    'Se não encontrar --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 96662

    'Faz lock na tabela OrigemDestino
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96601

    'Verifica se existe na tabela TabPreco registros relacionados com o código passado
    lErro = Comando_Executar(alComando(2), "SELECT Codigo, DataVigencia FROM TabPrecoGR WHERE Origem = ? OR Destino=?", lCodigo, dtData, objOrigemDestino.iCodigo, objOrigemDestino.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96602

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96603

    'Se existir, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 96604
    
    'Exclui a OrigemDestino com o Código passado
    lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM OrigemDestino", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96605

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96606

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrigemDestino_Exclui = SUCESSO

    Exit Function

Erro_OrigemDestino_Exclui:

    OrigemDestino_Exclui = gErr

    Select Case gErr

        Case 96597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 96598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96599, 96600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)

        Case 96601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)

        Case 96602, 96603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO1", gErr)
      
        Case 96604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEMDESTINO_ASSOCIADA_TABPRECO", gErr, lCodigo, dtData)

        Case 96605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORIGEMDESTINO", gErr, objOrigemDestino.iCodigo)

        Case 96606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 96662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objOrigemDestino.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ProgNavio_Exclui(objProgNavio As ClassProgNavio) As Long
'Exclui a ProgNavio

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim iIndice As Integer
Dim lNumero As Long

On Error GoTo Erro_ProgNavio_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96644
    Next

    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 96645

    'Verifica se existe na tabela ProgNavio o código passado
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM ProgNavio WHERE Codigo = ?", 0, lCodigo, objProgNavio.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96646

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96647
    
    'Se não encontrar --> erro.
    If lErro = AD_SQL_SEM_DADOS Then gError 96660

    'Faz lock na tabela ProgNavio
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96648

'    'Verifica se existe na tabela CompServ registros relacionados com o código passado
'    lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM CompServGR WHERE CodProgNavio = ?", lCodigo, objProgNavio.lCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 96649
'
'    lErro = Comando_BuscarPrimeiro(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96650
'
'    'Se existir, Erro.
'    If lErro = AD_SQL_SUCESSO Then gError 96651
'
    'Verifica se existe na tabela SolicitacaoServico registros relacionados com o código passado
    lErro = Comando_Executar(alComando(3), "SELECT Numero FROM SolicitacaoServico WHERE CodProgNavio = ?", lNumero, objProgNavio.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98301

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98302

    'Se existir, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 98303
    
    'Exclui a ProgNavio com o Código passado
    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ProgNavio", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96652

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96653

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProgNavio_Exclui = SUCESSO

    Exit Function

Erro_ProgNavio_Exclui:

    ProgNavio_Exclui = gErr

    Select Case gErr

        Case 96644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 96645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96646, 96647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROGNAVIO", gErr, objProgNavio.lCodigo)

        Case 96648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PROGNAVIO", gErr, objProgNavio.lCodigo)

        Case 96649, 96650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
      
        Case 96651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROGNAVIO_ASSOCIADA_COMPSERV", gErr, lCodigo)
        
        Case 96652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PROGNAVIO", gErr, objProgNavio.lCodigo)

        Case 96653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 96660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_ENCONTRADO", gErr, objProgNavio.lCodigo)
                 
        Case 98301, 98302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO1", gErr)
      
        Case 98303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROGNAVIO_ASSOCIADA_SOLICITACAO", gErr, lNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ProgNavio_Grava(objProgNavio As ClassProgNavio) As Long
'Grava ProgNavio no banco de dados

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim iIndice As Integer
Dim vi As String
Dim co As Long

On Error GoTo Erro_ProgNavio_Grava

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96631

    Next
    
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 96632
   
    'Verifica se existe uma viagem com código diferente
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM ProgNavio WHERE Codigo <> ? AND Viagem = ?", 0, lCodigo, objProgNavio.lCodigo, objProgNavio.sViagem)
    If lErro <> AD_SQL_SUCESSO Then gError 98382
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98383
    
    'Se Nome Viagem já existente em registro diferente --> erro
    If lErro = AD_SQL_SUCESSO Then gError 96635
   
    'Verifica se o campo Código já existem no BD
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM ProgNavio WHERE Codigo = ? ", 0, lCodigo, objProgNavio.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96633

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96634
    
    'Se não existe --> erro.
    If lErro = AD_SQL_SUCESSO Then
                    
        'Altera uma Programação de Navio já existente no BD
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ProgNavio SET Navio = ?,Terminal = ?, Armador = ?, AgMaritima = ?, Viagem = ?, Observacao= ?, DataChegada = ?, HoraChegada = ?, DataDeadLine = ?, HoraDeadLine =?", alComando(2), objProgNavio.sNavio, objProgNavio.sTerminal, objProgNavio.sArmador, objProgNavio.sAgMaritima, objProgNavio.sViagem, objProgNavio.sObservacao, objProgNavio.dtDataChegada, CDbl(objProgNavio.dtHoraChegada), objProgNavio.dtDataDeadLine, CDbl(objProgNavio.dtHoraDeadLine))
        If lErro <> AD_SQL_SUCESSO Then gError 96636

    Else

        'Insere no BD uma Programação de Navio nova
        lErro = Comando_Executar(alComando(3), "INSERT INTO ProgNavio ( Codigo, Navio, Terminal, Armador, AgMaritima, Viagem, Observacao, DataChegada, HoraChegada, DataDeadLine, HoraDeadLine) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objProgNavio.lCodigo, objProgNavio.sNavio, objProgNavio.sTerminal, objProgNavio.sArmador, objProgNavio.sAgMaritima, objProgNavio.sViagem, objProgNavio.sObservacao, objProgNavio.dtDataChegada, CDbl(objProgNavio.dtHoraChegada), objProgNavio.dtDataDeadLine, CDbl(objProgNavio.dtHoraDeadLine))
        If lErro <> AD_SQL_SUCESSO Then gError 96637

    End If

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96638

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    ProgNavio_Grava = SUCESSO

    Exit Function

Erro_ProgNavio_Grava:

    ProgNavio_Grava = gErr

    Select Case gErr

        Case 96631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96633, 96634, 98382, 98383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PROGNAVIO", gErr, objProgNavio.lCodigo)
        
        Case 96635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VIAGEM_JA_EXISTENTE", gErr, objProgNavio.sViagem)

        Case 96636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_PROGNAVIO", gErr, objProgNavio.lCodigo)

        Case 96637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PROGNAVIO", gErr, objProgNavio.lCodigo)

        Case 96638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Despachante_Grava(objDespachante As ClassDespachante) As Long
'Grava o Despachante

Dim lErro As Long
Dim lErro_Despachante As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lEndereco As Long
Dim iCodigo As Integer
Dim sNome As String
Dim sNomeReduzido As String
Dim sCGC As String

On Error GoTo Erro_Despachante_Grava

    'Abertura dos Comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96710
    Next

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 96711
    
    sNome = String(STRING_DESPACHANTE_NOME, 0)
    sNomeReduzido = String(STRING_DESPACHANTE_NOMEREDUZIDO, 0)
    sCGC = String(STRING_DESPACHANTE_CGC, 0)
    
    'Verifica se os campos Nome ou NomeReduzido ou CPF com Código diferente já existem no BD
    lErro = Comando_Executar(alComando(0), "SELECT CGC, Nome, NomeReduzido FROM Despachante WHERE Codigo <> ? and ( Nome = ? or NomeReduzido = ? or CGC = ?)", sCGC, sNome, sNomeReduzido, objDespachante.iCodigo, objDespachante.sNome, objDespachante.sNomeReduzido, objDespachante.sCGC)
    If lErro <> AD_SQL_SUCESSO Then gError 96712

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96713
        
    'Se encontrou --> Nome ou NomeReduzido ou CPF já existente
    If lErro = AD_SQL_SUCESSO Then
        If UCase(objDespachante.sNome) = UCase(sNome) Then
            gError 96714
        Else
            If UCase(objDespachante.sNomeReduzido) = UCase(sNomeReduzido) Then
                gError 96723
            Else 'CPF
                gError 96754
            End If
        End If
    End If
    
    'Verifica se o Código já existe no BD
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Endereco FROM Despachante WHERE Codigo = ?", 0, lEndereco, objDespachante.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96761

    lErro_Despachante = Comando_BuscarPrimeiro(alComando(1))
    If lErro_Despachante <> AD_SQL_SUCESSO And lErro_Despachante <> AD_SQL_SEM_DADOS Then gError 96762
        
    'Verifica se a variavel lEndereco contem informações
    If lEndereco = 0 Then
        lErro = CF("Endereco_Automatico", lEndereco)
        If lErro <> SUCESSO Then gError 96715
    End If
    
    'Carrega os dados em objDespachante
    objDespachante.objEndereco.lCodigo = lEndereco
    objDespachante.lEndereco = lEndereco

    lErro = CF("Endereco_Grava", objDespachante.objEndereco)
    If lErro <> SUCESSO Then gError 96716
            
    If lErro_Despachante = AD_SQL_SUCESSO Then
                      
        'Altera um Despachante já existente no BD
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Despachante SET CGC = ?, Nome = ?, NomeReduzido = ?, Endereco = ? ", alComando(1), objDespachante.sCGC, objDespachante.sNome, objDespachante.sNomeReduzido, objDespachante.lEndereco)
        If lErro <> AD_SQL_SUCESSO Then gError 96717
                                   
        lErro = DespachanteContatos_Exclui(objDespachante)
        If lErro <> SUCESSO Then gError 96718

        'Após a exclusão,grava novos Contatos do Despachante no BD
        lErro = DespachanteContatos_Inclui(objDespachante)
        If lErro <> SUCESSO Then gError 96719
                           
    Else
        
        'Insere no BD um Despachante novo
        lErro = Comando_Executar(alComando(2), "INSERT INTO Despachante ( Codigo, CGC, Nome, NomeReduzido, Endereco) VALUES (?,?,?,?,?)", objDespachante.iCodigo, objDespachante.sCGC, objDespachante.sNome, objDespachante.sNomeReduzido, objDespachante.lEndereco)
        If lErro <> AD_SQL_SUCESSO Then gError 96720
        
        'Grava os Contatos do Despachante no BD
        lErro = DespachanteContatos_Inclui(objDespachante)
        If lErro <> SUCESSO Then gError 96721

    End If
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96722

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Despachante_Grava = SUCESSO

    Exit Function

Erro_Despachante_Grava:

    Despachante_Grava = gErr

    Select Case gErr

        Case 96710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96712, 96713, 96761, 96762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESPACHANTE", gErr, objDespachante.iCodigo)
        
        Case 96714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_JA_EXISTENTE", gErr, objDespachante.sNome)
            
        Case 96715, 96716, 96718, 96719, 96721
            'Erro tratado na rotina chamadora
        
        Case 96717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_DESPACHANTE", gErr, objDespachante.iCodigo)
        
        Case 96720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DESPACHANTE", gErr, objDespachante.iCodigo)
            
        Case 96722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 96723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_JA_EXISTENTE", gErr, objDespachante.sNomeReduzido)
        
        Case 96754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_JA_EXISTENTE", gErr, objDespachante.sCGC)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function DespachanteContatos_Exclui(objDespachante As ClassDespachante) As Long
'Exclui os Itens do Contato do Despachante

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim lTransacao As Long

On Error GoTo Erro_DespachanteContatos_Exclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96724
    Next
        
    'Lê e Faz "lock" Contato
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Contato WHERE TipoNumIntDocOrigem = ? and NumIntDocOrigem = ?", 0, lCodigo, DESPACHANTE, objDespachante.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96725

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96726

    Do While lErro = AD_SQL_SUCESSO

        'Faz "LockExclusive" de Contato
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 96727

        'Exclui o Contato do banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Contato", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 96728

        'Busca próximo Contato
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96729

    Loop
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    DespachanteContatos_Exclui = SUCESSO

    Exit Function

Erro_DespachanteContatos_Exclui:

    DespachanteContatos_Exclui = gErr

    Select Case gErr

        Case 96724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96725, 96726, 96729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATO", gErr)

        Case 96727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTATO", gErr, objDespachante.iCodigo)

        Case 96728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTATO", gErr, objDespachante.iCodigo)
            
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function DespachanteContatos_Inclui(objDespachante As ClassDespachante) As Long
'Insere os Itens do Despachante no BD

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objContato As ClassContato
Dim iIndice As Integer
Dim lCodigo As Long
Dim lTransacao As Long

On Error GoTo Erro_DespachanteContatos_Inclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96730
    Next
    
    'Para cada item do Despachante
    For Each objContato In objDespachante.colContato
               
        lErro = CF("Config_ObterAutomatico_EmTrans", "CRFatConfig", "NUM_PROX_CONTATO", "Contato", "NumIntDoc", lCodigo)
        If lErro <> SUCESSO Then gError 96731

        objContato.lNumIntDoc = lCodigo
        objContato.lNumIntDocOrigem = objDespachante.iCodigo
                
        'Grava o Contato do Despachante no BD
        lErro = Comando_Executar(alComando(1), "INSERT INTO Contato (NumIntDoc, TipoNumIntDocOrigem, NumIntDocOrigem, Contato, Fax, Email, Telefone, Setor) VALUES (?,?,?,?,?,?,?,?)", objContato.lNumIntDoc, DESPACHANTE, objContato.lNumIntDocOrigem, objContato.sContato, objContato.sFax, objContato.sEmail, objContato.sTelefone, objContato.sSetor)
        If lErro <> AD_SQL_SUCESSO Then gError 96732
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    DespachanteContatos_Inclui = SUCESSO

    Exit Function

Erro_DespachanteContatos_Inclui:

    DespachanteContatos_Inclui = gErr

    Select Case gErr
        
        Case 96730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96731

        Case 96732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTATO", gErr, objDespachante.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Despachante_Exclui(objDespachante As ClassDespachante) As Long
'Exclui o Despachante

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lEndereco As Long
Dim lCodigo As Long
Dim iCodigo As Integer
Dim lNumero As Long

On Error GoTo Erro_Despachante_Exclui

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96734
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 96735
    
    'Seleciona Despachante com o código passado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo, Endereco FROM Despachante WHERE Codigo = ? ", 0, iCodigo, lEndereco, objDespachante.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96736

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96737
    
    If lErro = AD_SQL_SEM_DADOS Then gError 96738

    'Lock do Despachante
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 96739
    
    'Verifica se existe na tabela SolicitacaoServico registros relacionados com o código passado
    lErro = Comando_Executar(alComando(2), "SELECT Numero FROM SolicitacaoServico WHERE Despachante = ?", lNumero, objDespachante.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96743

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96744

    'Se existir, Erro.
    If lErro = AD_SQL_SUCESSO Then gError 96745
    
    'Exclui todos os contatos do despachante a ser excluído
    lErro = DespachanteContatos_Exclui(objDespachante)
    If lErro <> SUCESSO Then gError 96746
    
    'Verifica se a variavel lEndereco contem informações
    If lEndereco > 0 Then
        lErro = CF("Endereco_Exclui", lEndereco)
        If lErro <> SUCESSO Then gError 96747
    End If

    'Exclui o Despachante
    lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM Despachante", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 96748

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96749

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Despachante_Exclui = SUCESSO

    Exit Function

Erro_Despachante_Exclui:

    Despachante_Exclui = gErr

    Select Case gErr
    
        Case 96734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96736, 96737
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESPACHANTE", gErr, objDespachante.iCodigo)
        
        Case 96738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESPACHANTE_NAO_CADASTRADO", gErr, objDespachante.iCodigo)

        Case 96739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_DESPACHANTE", gErr, objDespachante.iCodigo)
        
        Case 96743, 96744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO1", gErr)
        
        Case 96745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESPACHANTE_ASSOCIADO_SOLICITACAOSERVICO", gErr, lNumero)

        Case 96746, 96747
        
        Case 96748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DESPACHANTE", gErr, objDespachante.iCodigo)

        Case 96749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
    
Function TabPreco_Exclui(objTabPreco As ClassTabPreco) As Long
'Exclui o TabPreco

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim dtData As Date
Dim lCodigo As Long
Dim lNum As Long

On Error GoTo Erro_TabPreco_Exclui

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96940
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 96941
    
    'Seleciona TabPreco com o código passado e data de vigência
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TabPrecoGR WHERE Codigo = ? and DataVigencia = ?", 0, lCodigo, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 96942

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96943
    
    'Se não emcontrar --> erro.
    If lErro = AD_SQL_SEM_DADOS Then gError 96944

    'Lock da TabPreco
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 96945
    
    'Seleciona TabPreco com data menor que a data de vigência
    'Se existir --> tabela de preço passada pode ser excluída
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo ,DataVigencia FROM TabPrecoGR WHERE Codigo = ? and DataVigencia < ?", 0, lCodigo, dtData, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 98009

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98010
    
    'Se não encontrar tabela com data de vigência menor ...
    If lErro = AD_SQL_SEM_DADOS Then

        'Seleciona TabPreco com data maior que a data de vigência
        lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo ,DataVigencia FROM TabPrecoGR WHERE Codigo = ? and DataVigencia > ? ORDER BY DataVigencia", 0, lCodigo, dtData, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 98011
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98012
        
        'Se não encontrar tabela com data de vigência maior ...
        If lErro = AD_SQL_SEM_DADOS Then
                
            'Verifica se existe uma solicitação associada a tabela de preço na vigência passada
            lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM SolicitacaoServico WHERE CodTabPreco = ? AND DataEmissao >= ?", 0, lNum, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
            If lErro <> AD_SQL_SUCESSO Then gError 98013
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98014
            
            'Se existe --> não pode excluir a tabela de preço
            If lErro = AD_SQL_SUCESSO Then gError 98015
            
        'Se encontrar tabela com data de vigência maior
        Else
            
            'Verifica se existe uma solicitação associada a tabela de preço entre a menor e maior data de vigência da  vigência passada
            lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM SolicitacaoServico WHERE CodTabPreco = ? AND DataEmissao >= ? AND DataEmissao < ?", 0, lNum, objTabPreco.lCodigo, objTabPreco.dtDataVigencia, dtData)
            If lErro <> AD_SQL_SUCESSO Then gError 98016
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98017
            
            'Se encontrar a tabela entre esse espaço de tempo --> não pode excluir a tabela
            If lErro = AD_SQL_SUCESSO Then gError 98018
            
        End If
    End If
    
    'Exclui todos os itens associados a tabela de preço
    lErro = TabPrecoItens_Exclui_Todas(objTabPreco)
    If lErro <> SUCESSO Then gError 96951
    
    'Exclui a TabPreco
    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM TabPrecoGR", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 96954
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96955

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TabPreco_Exclui = SUCESSO

    Exit Function

Erro_TabPreco_Exclui:

    TabPreco_Exclui = gErr

    Select Case gErr
    
        Case 96940
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96942, 96943, 98009, 98010, 98011, 98012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objTabPreco.lCodigo)
        
        Case 96944
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_CADASTRADA", gErr, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)

        Case 96945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABPRECO", gErr, objTabPreco.lCodigo)
        
        Case 96951
                             
        Case 96954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TABPRECO", gErr, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)

        Case 96955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 98013, 98014, 98016, 98017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAOSERVICO1", gErr)
        
        Case 98015, 98018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_ASSOCIADA_SOLICITACAOSERVICO", gErr, lNum)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function TabPreco_Grava(objTabPreco As ClassTabPreco) As Long
'Grava os dados recebidos em TabPreco

Dim lErro As Long
Dim alComando(1 To 7) As Long
Dim lCodigo As Long
Dim iCodigo As Integer
Dim lTransacao As Long
Dim iIndice As Integer
Dim objTabPrecoItens As ClassTabPrecoItens
Dim iOrigem As Integer
Dim iDestino As Integer
Dim lCliente As Long
Dim lNumero As Long
Dim iFilial As Integer

On Error GoTo Erro_TabPreco_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 96905

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96906
    Next
    
    'Ler e lock Cliente
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM Clientes WHERE Codigo=?", 0, lCodigo, objTabPreco.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 96907

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96908
    
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 96909
    
    'Ler e lock  Origem
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM OrigemDestino WHERE Codigo=?", 0, iCodigo, objTabPreco.iOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 96910

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96911
    
    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 96912
    
    'Ler e lock Destino
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Codigo FROM OrigemDestino WHERE Codigo=?", 0, iCodigo, objTabPreco.iDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 96913

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96914
    
    lErro = Comando_LockShared(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 96915
    
    'Ler e faz lock em todos os produtos passados na coleção
    For Each objTabPrecoItens In objTabPreco.colTabPrecoItens
        
        'Seleciona cada um dos produtos passados
        lErro = Comando_ExecutarPos(alComando(4), "SELECT Tipo FROM Produtos WHERE Codigo=?", 0, iCodigo, objTabPrecoItens.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 96916
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96917
        
        lErro = Comando_LockShared(alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 96918
        
    Next
    
    'Verifica no Banco de dados se Já existe tabela com o Código recebido no obj
    lErro = Comando_ExecutarPos(alComando(5), "SELECT Cliente, Origem, Destino FROM TabPrecoGR WHERE Codigo=?", 0, lCliente, iOrigem, iDestino, objTabPreco.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 96919

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96920

    'Se Existe...
    If lErro = AD_SQL_SUCESSO Then
        
        'Verifica se o Cliente, a origem e o destino são iguais, senão -->erro
        If objTabPreco.lCliente <> lCliente Or objTabPreco.iDestino <> iDestino Or objTabPreco.iOrigem <> iOrigem Then gError 98068

        'Verifica no Banco de dados se Já existe tabela com o Código e a DataVigencia recebidos no obj
        lErro = Comando_ExecutarPos(alComando(6), "SELECT Codigo FROM TabPrecoGR WHERE Codigo=? and DataVigencia = ?", 0, lCodigo, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 98066
    
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98067
    
        'Se Existe ==> Atualizar os dados
        If lErro = AD_SQL_SUCESSO Then
                 
            'Altera uma tabela de preço já existente
            lErro = Comando_ExecutarPos(alComando(7), "UPDATE TabPrecoGR SET Pedagio = ?, AdValoren = ?, Observacao=?", alComando(6), objTabPreco.dPedagio, objTabPreco.dAdValoren, objTabPreco.sObservacao)
            If lErro <> AD_SQL_SUCESSO Then gError 96921
            
            'Verificar se existe Solicitação de servico relacionados a esta Tabela de Preço
            lErro = Comando_Executar(alComando(7), "SELECT FilialEmpresa, Numero FROM SolicitacaoServico WHERE CodTabPreco=? ", iFilial, lNumero, objTabPreco.lCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 98349
            
            lErro = Comando_BuscarPrimeiro(alComando(7))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98350
                            
            'Se existe pelo menos uma solicitação associada a esta tabela de preço
            If lErro = AD_SQL_SUCESSO Then
                
                'Verifica cada produto da tabela de preço para ver se está relacionado a alguma solicitação
                lErro = CF("Verifica_Servico_TabPreco", objTabPreco)
                If lErro <> SUCESSO Then gError 98351
                
            'Caso não ache a solicitação associada
            ElseIf lErro = AD_SQL_SEM_DADOS Then
                
                'Exclui todos os itens associados a tabela
                lErro = TabPrecoItens_Exclui_Todas(objTabPreco)
                If lErro <> SUCESSO Then gError 96922
                
                'Inclui os novos produtos passados no grid
                lErro = TabPrecoItens_Inclui(objTabPreco.colTabPrecoItens)
                If lErro <> SUCESSO Then gError 96923
        
            End If
            
        Else
            'Se não existe, cria uma nova tabela com a Data de Vigência em questão
            lErro = Comando_Executar(alComando(7), "INSERT INTO TabPrecoGR (Codigo,DataVigencia, Cliente, Origem, Destino, ADValoren, Pedagio, Observacao) VALUES (?,?,?,?,?,?,?,?)", _
            objTabPreco.lCodigo, objTabPreco.dtDataVigencia, objTabPreco.lCliente, objTabPreco.iOrigem, objTabPreco.iDestino, objTabPreco.dAdValoren, objTabPreco.dPedagio, objTabPreco.sObservacao)
            If lErro <> AD_SQL_SUCESSO Then gError 96924
                                                            
            'Inclui os produtos passados no grid
            lErro = TabPrecoItens_Inclui(objTabPreco.colTabPrecoItens)
            If lErro <> SUCESSO Then gError 96925
                
        End If
        
    'Caso não encontrou o código passado da tabela
    ElseIf lErro = AD_SQL_SEM_DADOS Then
        
        'Insere a nova tabela de preço
        lErro = Comando_Executar(alComando(6), "INSERT INTO TabPrecoGR (Codigo,DataVigencia, Cliente, Origem, Destino, ADValoren, Pedagio, Observacao) VALUES (?,?,?,?,?,?,?,?)", _
        objTabPreco.lCodigo, objTabPreco.dtDataVigencia, objTabPreco.lCliente, objTabPreco.iOrigem, objTabPreco.iDestino, objTabPreco.dAdValoren, objTabPreco.dPedagio, objTabPreco.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 96926
                                                              
        'Inclui os produtos passados no grid
        lErro = TabPrecoItens_Inclui(objTabPreco.colTabPrecoItens)
        If lErro <> SUCESSO Then gError 96927
                                                  
    End If

    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 96928

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
        
    TabPreco_Grava = SUCESSO

    Exit Function

Erro_TabPreco_Grava:

    TabPreco_Grava = gErr

    Select Case gErr

        Case 96905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 96906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96907, 96908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr, objTabPreco.lCliente)

        Case 96909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", gErr, objTabPreco.lCliente)
        
        Case 96910, 96911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, objTabPreco.iOrigem)

        Case 96912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORIGEMDESTINO", gErr, objTabPreco.iOrigem)
        
        Case 96913, 96914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEMDESTINO", gErr, objTabPreco.iDestino)

        Case 96915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ORIGEMDESTINO", gErr, objTabPreco.iDestino)
        
        Case 96916, 96917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objTabPrecoItens.sProduto)

        Case 96918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", gErr)
        
        Case 96919, 96920, 98066, 98067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECO", gErr, objTabPreco.lCodigo)
        
        Case 96921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABPRECO", gErr, objTabPreco.lCodigo)

        Case 96922, 96923, 96925, 96927, 98351
        
        Case 96924, 96926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABPRECO", gErr, objTabPreco.lCodigo)
        
        Case 96928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case 98068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_ALTERADO", gErr, objTabPreco.lCodigo)
            
        Case 98349, 98350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAO", gErr, iFilial, lNumero)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TabPrecoItens_Exclui_Todas(objTabPreco As ClassTabPreco) As Long
'Exclui todos os Itens da Tabela de Preços
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim lTransacao As Long

On Error GoTo Erro_TabPrecoItens_Exclui_Todas

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96987
    Next
        
    'Lê e Faz "lock" de TabPrecoItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodTabela FROM TabPrecoItensGR WHERE CodTabela = ? and DataVigencia = ?", 0, lCodigo, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 96988

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96989

    Do While lErro = AD_SQL_SUCESSO

        'Faz "LockExclusive" de Contato
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 96990

        'Exclui o TabPrecoItens do banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TabPrecoItensGR", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 96991

        'Busca próximo Item
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96992

    Loop
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TabPrecoItens_Exclui_Todas = SUCESSO

    Exit Function

Erro_TabPrecoItens_Exclui_Todas:

    TabPrecoItens_Exclui_Todas = gErr

    Select Case gErr

        Case 96987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96988, 96989, 96992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENS", gErr, objTabPreco.lCodigo)

        Case 96990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABPRECOITENS", gErr, objTabPreco.lCodigo)

        Case 96991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TABPRECOITENS", gErr, objTabPreco.lCodigo)
            
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TabPrecoItens_Exclui(ByVal ColServico As Collection, ByVal objTabPreco As ClassTabPreco) As Long
'Exclui os Itens da tabela TabPrecoItens no BD recebidos na coleção
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objServico As ClassServico
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_TabPrecoItens_Exclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98368
    Next
    
    'Para cada item da Coleção a ser excluído
    For Each objServico In ColServico
    
        'Lê e Faz "lock" da TabPrecoItens relacionado ao serviço
        lErro = Comando_ExecutarPos(alComando(0), "SELECT CodTabela FROM TabPrecoItensGR WHERE Produto = ? AND CodTabela = ? AND DataVigencia =?", 0, lCodigo, objServico.sProduto, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 98369
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98370
        
        'Faz "LockExclusive" de TabPrecoItensGR
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98371

        'Exclui o Item na tabela de TabPrecoItensGR no BD
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TabPrecoItensGR", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98372
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TabPrecoItens_Exclui = SUCESSO

    Exit Function

Erro_TabPrecoItens_Exclui:

    TabPrecoItens_Exclui = gErr

    Select Case gErr
        
        Case 98368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98369, 98370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABPRECOITENS", gErr, objTabPreco.lCodigo)

        Case 98371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABPRECOITENS", gErr, objTabPreco.lCodigo)

        Case 98372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TABPRECOITENS", gErr, objTabPreco.lCodigo, objTabPreco.dtDataVigencia)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TabPrecoItens_Inclui(ByVal ColServico As Collection) As Long
'Insere os Itens da tabela de preços no BD
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objTabPrecoItens As ClassTabPrecoItens
Dim iIndice As Integer
Dim lCodigo As Long
Dim lTransacao As Long

On Error GoTo Erro_TabPrecoItens_Inclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 96993
    Next
    
    'Para cada item da Tabela de Preços
    For Each objTabPrecoItens In ColServico
             
        'Grava o Item na tabela de preços no BD
        lErro = Comando_Executar(alComando(1), "INSERT INTO TabPrecoItensGR (CodTabela, DataVigencia, Produto, Preco) VALUES (?,?,?,?)", objTabPrecoItens.lCodTabela, objTabPrecoItens.dtDataVigencia, objTabPrecoItens.sProduto, objTabPrecoItens.dPreco)
        If lErro <> AD_SQL_SUCESSO Then gError 96994
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TabPrecoItens_Inclui = SUCESSO

    Exit Function

Erro_TabPrecoItens_Inclui:

    TabPrecoItens_Inclui = gErr

    Select Case gErr
        
        Case 96993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABPRECOITENS", gErr, objTabPrecoItens.lCodTabela)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Documento_Exclui(objDocumento As ClassDocumento) As Long
'Exclui o Documento

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim iCodigo As Integer

On Error GoTo Erro_Documento_Exclui

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98045
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98046
    
    'Seleciona Documento com o código passado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Documento WHERE Codigo = ? ", 0, iCodigo, objDocumento.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98047

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98048
    
    'Se não encontrar --> erro.
    If lErro = AD_SQL_SEM_DADOS Then gError 98049

    'Lock do Documento
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98050
    
    'Verifica se o Documento está associado a algum comprovante de Serviço
    lErro = Comando_Executar(alComando(1), "SELECT CompServGR.Codigo FROM CompServItemGR, CompServGR WHERE (DocIntTipo= ? or DocExtTipo = ?) AND CompServItemGR.NumIntDocOrigem = CompServGR.NumIntDoc", lCodigo, iCodigo, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98051

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98052
    
    'Se estiver --> Erro.
    If lErro = AD_SQL_SUCESSO Then gError 98053
                   
    'Exclui o Documento
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Documento", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98054
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98055

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

     Documento_Exclui = SUCESSO

    Exit Function

Erro_Documento_Exclui:

    Documento_Exclui = gErr

    Select Case gErr
    
        Case 98945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 98047, 98048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCUMENTO", gErr, objDocumento.iCodigo)
        
        Case 98049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_CADASTRADA", gErr, objDocumento.iCodigo)

        Case 98050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_DOCUMENTO", gErr, objDocumento.iCodigo)
        
        Case 98051, 98052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
                
        Case 98053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_ASSOCIADO_COMPSERV", gErr, lCodigo)
            
        Case 98954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DOCUMENTO", gErr, objDocumento.iCodigo)

        Case 98955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Documento_Grava(objDocumento As ClassDocumento) As Long
'Grava os dados recebidos em Documento

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iCodigo As Integer
Dim lTransacao As Long
Dim iIndice As Integer
Dim sNome As String

On Error GoTo Erro_Documento_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98056

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98057
    Next
            
    'Verifica no Banco de dados se Já existe Documento com o Código diferente e Nome reduzido iqual
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Documento WHERE Codigo<>? and NomeReduzido = ?", 0, iCodigo, objDocumento.iCodigo, objDocumento.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 98058

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98059

    'Se Existe... ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 98061
        
    'Verifica no Banco de dados se Já existe Documento com o Código
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM Documento WHERE Codigo=? ", 0, iCodigo, objDocumento.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 98064

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98065

    'Se não Existe...
    If lErro = AD_SQL_SEM_DADOS Then
        
        'cria um novo Documento.
        lErro = Comando_Executar(alComando(2), "INSERT INTO Documento ( Codigo, Descricao, NomeReduzido, TipoDoc,Documento ) VALUES (?,?,?,?,?)", objDocumento.iCodigo, objDocumento.sDescricao, objDocumento.sNomeReduzido, objDocumento.iTipoDoc, objDocumento.sDocumento)
        If lErro <> AD_SQL_SUCESSO Then gError 98060
                                                
    Else
        'altera um documento já existente.
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Documento SET NomeReduzido = ?, Descricao = ?, TipoDoc=?,Documento = ?", alComando(1), objDocumento.sNomeReduzido, objDocumento.sDescricao, objDocumento.iTipoDoc, objDocumento.sDocumento)
        If lErro <> AD_SQL_SUCESSO Then gError 98062
        
    End If
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98063

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
        
    Documento_Grava = SUCESSO

    Exit Function

Erro_Documento_Grava:

    Documento_Grava = gErr

    Select Case gErr

        Case 98056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 98057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98058, 98059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCUMENTO1", gErr)
        
        Case 98064, 98065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCUMENTO", gErr, objDocumento.iCodigo)

        Case 98060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DOCUMENTO", gErr, objDocumento.iCodigo)
        
        Case 98061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_JA_EXISTENTE", gErr, iCodigo, objDocumento.sNomeReduzido)
            
        Case 98062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_DOCUMENTO", gErr, objDocumento.iCodigo)

        Case 98063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Cotacao_Insere(ByVal objCotacao As ClassCotacaoGR) As Long
'Faz a inclusao da Cotacao

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim sCamposCotacaoGR As String
Dim sObsDestOrigem1 As String
Dim sObsDestOrigem2 As String
Dim sObsDestOrigem3 As String
Dim sObsDestOrigem4 As String

On Error GoTo Erro_Cotacao_Insere
        
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97317
    'Cyntia
    sCamposCotacaoGR = "NumIntDoc,FilialEmpresa,Codigo,Data,TipoOperacao,Cliente,Envio," _
    & "EnvioComplemento,CodVendedor,Indicacao,Observacao,DataPrevInicio," _
    & "TipoEmbalagem,Ajudantes,Carga,CargaPorConta,Descarga,DescargaPorConta,Ova,OvaPorConta," _
    & "Desova,DesovaPorConta,CargaSolta,DescCargaSolta,CondicaoPagto,Situacao,Justificativa,ObsResultado,ValorMercadoria,Atendente, Responsavel,ObsDestOrigem1, ObsDestOrigem2, ObsDestOrigem3, ObsDestOrigem4"
    
    sObsDestOrigem1 = Mid(objCotacao.sObsDestOrigem, 1, 250)
    sObsDestOrigem2 = Mid(objCotacao.sObsDestOrigem, 251, 250)
    sObsDestOrigem3 = Mid(objCotacao.sObsDestOrigem, 501, 250)
    sObsDestOrigem4 = Mid(objCotacao.sObsDestOrigem, 751, 250)
    
    'Pegar o numintdoc
    lErro = CF("Config_ObterNumInt", "FatConfig", "NUMINTDOC_PROX_COTACAO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 97274
    
    'Cyntia
    'insere o registro na tabela CotacaoGR
    lErro = Comando_Executar(lComando, "INSERT INTO CotacaoGR (" & sCamposCotacaoGR & ") values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
    lNumIntDoc, objCotacao.iFilialEmpresa, objCotacao.lCodigo, objCotacao.dtData, objCotacao.iTipoOperacao, objCotacao.sCliente, objCotacao.iEnvio, objCotacao.sEnvioComplemento, _
    objCotacao.iCodVendedor, objCotacao.sIndicacao, objCotacao.sObservacao, objCotacao.dtDataPrevInicio, objCotacao.iTipoEmbalagem, objCotacao.iAjudantes, objCotacao.iCarga, _
    objCotacao.iCargaPorConta, objCotacao.iDesCarga, objCotacao.iDesCargaPorConta, objCotacao.iOva, objCotacao.iOvaPorConta, objCotacao.iDesova, objCotacao.iDesovaPorConta, objCotacao.iCargaSolta, _
    objCotacao.sDescCargaSolta, objCotacao.iCondicaoPagto, objCotacao.iSituacao, objCotacao.iJustificativa, objCotacao.sObsResultado, objCotacao.dValorMerc, objCotacao.sAtendente, objCotacao.sResponsavel, sObsDestOrigem1, sObsDestOrigem2, sObsDestOrigem3, sObsDestOrigem4)
    
    If lErro <> AD_SQL_SUCESSO Then gError 97280

    objCotacao.lNumIntDoc = lNumIntDoc
       
    'fechar comando
    Call Comando_Fechar(lComando)
    
    Cotacao_Insere = SUCESSO

    Exit Function

Erro_Cotacao_Insere:

    Cotacao_Insere = gErr
    
    Select Case gErr
    
        Case 97274
        
        Case 97280
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAO1", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
        
        Case 97317
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select
    
    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_Insere_Contato(ByVal objCotacao As ClassCotacaoGR) As Long
'Faz a inclusao dos Contatos relacionados com a cotacao

Dim lErro As Long
Dim lComando As Long
Dim objContato As ClassContato
Dim lNumIntDoc As Long
        
On Error GoTo Erro_Cotacao_Insere_Contato

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97318
    
    For Each objContato In objCotacao.colContato
    
        'Pegar o numintdoc do contato
        lErro = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_CONTATO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 97275
        
        lErro = Comando_Executar(lComando, "INSERT INTO Contato (NumIntDoc,TipoNumIntDocOrigem,NumIntDocOrigem,Contato,Fax,Email,Telefone,Setor) values (?,?,?,?,?,?,?,?)", _
        lNumIntDoc, TIPO_DOC_ORIGEM_COTACAO, objCotacao.lNumIntDoc, objContato.sContato, objContato.sFax, objContato.sEmail, objContato.sTelefone, objContato.sSetor)
    
        If lErro <> AD_SQL_SUCESSO Then gError 97282
    
    Next

    'fechar comando
    Call Comando_Fechar(lComando)
    
    Cotacao_Insere_Contato = SUCESSO

    Exit Function

Erro_Cotacao_Insere_Contato:

    Cotacao_Insere_Contato = gErr
    
    Select Case gErr

        Case 97275
                      
        Case 97282
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTATO1", gErr)
                    
        Case 97318
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select

    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_Insere_CotacaoContainer(ByVal objCotacao As ClassCotacaoGR) As Long
'Faz a inclusao das CotacoesContainer relacionadas com a cotacao

Dim lErro As Long
Dim lComando As Long
Dim objCotacaoContainer As ClassCotacaoContainer
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Insere_CotacaoContainer

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97319

    For Each objCotacaoContainer In objCotacao.colCotacaoContainer
        
        'Pegar o numintdoc da cotacaocontainer
        lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_COTACAOCONTAINER", lNumIntDoc)
        If lErro <> SUCESSO Then gError 97277
                           
        lErro = Comando_Executar(lComando, "INSERT INTO CotacaoContainer (NumIntDoc,NumIntDocOrigem,TipoContainer,Quantidade) values (?,?,?,?)", lNumIntDoc, objCotacao.lNumIntDoc, objCotacaoContainer.iTipoContainer, objCotacaoContainer.iQuantidade)
        If lErro <> AD_SQL_SUCESSO Then gError 97284
        
    Next

    'fechar comando
    Call Comando_Fechar(lComando)
    
    Cotacao_Insere_CotacaoContainer = SUCESSO

    Exit Function

Erro_Cotacao_Insere_CotacaoContainer:

    Cotacao_Insere_CotacaoContainer = gErr
    
    Select Case gErr

        Case 97277
        
        Case 97284
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTACAOCONTAINER", gErr)
        
        Case 97319
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select
    
    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_Insere_CotacaoServico(ByVal objCotacao As ClassCotacaoGR) As Long
'Faz a inclusao das CotacoesServico relacionadas com a cotacao

Dim lErro As Long
Dim lComando As Long
Dim objCotacaoServico As ClassCotacaoServico
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Insere_CotacaoServico

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97320

    For Each objCotacaoServico In objCotacao.colCotacaoServico
        
        'Pegar o numintdoc da cotacaoservico
        lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_COTACAOSERVICO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 97278
                            
        lErro = Comando_Executar(lComando, "INSERT INTO CotacaoServico (NumIntDoc,NumIntDocOrigem,Produto,Quantidade,PrecoUnitario,AdValoren,Pedagio,CodigoOrigem,CodigoDestino) values (?,?,?,?,?,?,?,?,?)", lNumIntDoc, objCotacao.lNumIntDoc, objCotacaoServico.sProduto, objCotacaoServico.dQuantidade, objCotacaoServico.dPrecoUnitario, objCotacaoServico.dAdValoren, objCotacaoServico.dPedagio, objCotacaoServico.iOrigem, objCotacaoServico.iDestino)
    
        If lErro <> AD_SQL_SUCESSO Then gError 97285
    
    Next
    
    'fechar comando
    Call Comando_Fechar(lComando)
    
    Cotacao_Insere_CotacaoServico = SUCESSO
    
    Exit Function

Erro_Cotacao_Insere_CotacaoServico:

    Cotacao_Insere_CotacaoServico = gErr
    
    Select Case gErr

        Case 97278

        Case 97285
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOSERVICO", gErr)
            
        Case 97320
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select

    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_Insere_CotacaoOrigemDestino(ByVal objCotacao As ClassCotacaoGR) As Long
'Faz a inclusao das CotacoesOrigemDestino relacionadas com a cotacao

Dim lErro As Long
Dim lComando As Long
Dim objCotacaoDestOrigem As ClassCotacaoOrigemDestino
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Insere_CotacaoOrigemDestino

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97321

    For Each objCotacaoDestOrigem In objCotacao.colCotacaoOrigemDestino
        
        'Pegar o numintdoc da cotacaoorigemdestino
        lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_COTACAOORIGEMDESTINO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 97276
                            
        lErro = Comando_Executar(lComando, "INSERT INTO CotacaoOrigemDestino (NumIntDoc,NumIntDocOrigem,Servico,Origem,Destino) values (?,?,?,?,?)", lNumIntDoc, objCotacao.lNumIntDoc, objCotacaoDestOrigem.sServico, objCotacaoDestOrigem.sOrigem, objCotacaoDestOrigem.sDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 97286
    
    Next
    
    'fechar comando
    Call Comando_Fechar(lComando)
    
    Cotacao_Insere_CotacaoOrigemDestino = SUCESSO
    
    Exit Function

Erro_Cotacao_Insere_CotacaoOrigemDestino:

    Cotacao_Insere_CotacaoOrigemDestino = gErr
    
    Select Case gErr

        Case 97276
        
        Case 97286
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOORIGEMDESTINO", gErr)
        
        Case 97321
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select
    
    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_ExcluiGR(ByVal objCotacao As ClassCotacaoGR) As Long
'Exclui a Cotacao

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim tCotacao As typeCotacaoGR

On Error GoTo Erro_Cotacao_ExcluiGR

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97220
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97221

    'Le a Cotacao passada como parametro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Codigo, FilialEmpresa FROM CotacaoGR WHERE Codigo=? AND FilialEmpresa=? ", 0, tCotacao.lNumIntDoc, tCotacao.lCodigo, tCotacao.iFilialEmpresa, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 97226
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 97227
    
    'Lock Exclusive da Cotacao
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 97224
    
    'Exclui os registros que apontam para a cotacao em questao nas
    'tabelas em que os mesmos possam existir
    lErro = Cotacao_Exclui_Relacionamentos(objCotacao)
    If lErro <> SUCESSO Then gError 97225
    
    'Exclui a Cotacao
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CotacaoGR", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 97228

    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 97229

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cotacao_ExcluiGR = SUCESSO
    
    Exit Function
    
Erro_Cotacao_ExcluiGR:

    Cotacao_ExcluiGR = gErr
    
    Select Case gErr
    
        Case 97220
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97221
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 97224
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COTACAO1", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
        
        Case 97226, 97227
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO1", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
        
        Case 97225
            
        Case 97228
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAO1", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
            
        Case 97229
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function Cotacao_Exclui_Contato(ByVal objCotacao As ClassCotacaoGR) As Long
'Exclui os Contatos relacionados com a Cotacao em questao
'IMPORTANTE: Chamar esta funçao em transacao

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Exclui_Contato

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97231
    Next
    
    'Busca Contato associado a Cotacao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocOrigem FROM Contato WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem=? ", 0, lNumIntDoc, TIPO_DOC_ORIGEM_COTACAO, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97232
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97233
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui Contato ligado a Cotacao
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Contato", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 97235

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97234

    Loop

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Cotacao_Exclui_Contato = SUCESSO
    
    Exit Function

Erro_Cotacao_Exclui_Contato:

    Cotacao_Exclui_Contato = gErr
    
    Select Case gErr
    
        Case 97231
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97232, 97233, 97234
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATO", gErr)
        
        Case 97235
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTATO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Cotacao_Exclui_CotacaoContainer(ByVal objCotacao As ClassCotacaoGR) As Long
'Exclui as CotacoesContainer relacionadas com a Cotacao em questao
'IMPORTANTE: Chamar esta funçao em transacao

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Exclui_CotacaoContainer

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97367
    Next
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocOrigem FROM CotacaoContainer WHERE NumIntDocOrigem=? ", 0, lNumIntDoc, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97236
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97237
        
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui cotacaocontainer ligado a Cotacao
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CotacaoContainer", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 97238

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97239

    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cotacao_Exclui_CotacaoContainer = SUCESSO
    
    Exit Function

Erro_Cotacao_Exclui_CotacaoContainer:

    Cotacao_Exclui_CotacaoContainer = gErr
    
    Select Case gErr
    
        Case 97236, 97237, 97239
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOCONTAINER", gErr)
            
        Case 97238
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOCONTAINER", gErr)
        
        Case 97367
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Cotacao_Exclui_CotacaoOrigemDestino(ByVal objCotacao As ClassCotacaoGR) As Long
'Exclui as CotacoesOrigemDestino relacionadas com a Cotacao em questao
'IMPORTANTE: Chamar esta funçao em transacao

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Exclui_CotacaoOrigemDestino

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97368
    Next
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocOrigem FROM CotacaoOrigemDestino WHERE NumIntDocOrigem=? ", 0, lNumIntDoc, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97240
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97241
        
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui CotacoesOrigemDestino ligado a Cotacao
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CotacaoOrigemDestino", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 97242

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97243

    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cotacao_Exclui_CotacaoOrigemDestino = SUCESSO
    
    Exit Function

Erro_Cotacao_Exclui_CotacaoOrigemDestino:

    Cotacao_Exclui_CotacaoOrigemDestino = gErr
    
    Select Case gErr
    
       Case 97240, 97241, 97243
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOORIGEMDESTINO", gErr)
            
        Case 97242
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOORIGEMDESTINO", gErr)
        
        Case 97368
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Cotacao_Exclui_CotacaoServico(ByVal objCotacao As ClassCotacaoGR) As Long
'Exclui as CotacoesServicos relacionadas com a Cotacao em questao
'IMPORTANTE: Chamar esta funçao em transacao

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Cotacao_Exclui_CotacaoServico

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97370
    Next
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocOrigem FROM CotacaoServico WHERE NumIntDocOrigem=? ", 0, lNumIntDoc, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 97244
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97245
        
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui Servicos ligado a Cotacao
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CotacaoServico", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 97246

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97247

    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cotacao_Exclui_CotacaoServico = SUCESSO
    
    Exit Function

Erro_Cotacao_Exclui_CotacaoServico:

    Cotacao_Exclui_CotacaoServico = gErr
    
    Select Case gErr
    
        Case 97244, 97245, 97247
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOSERVICO", gErr)
            
        Case 97246
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOSERVICO", gErr)
        
        Case 97370
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Cotacao_GravaGR(ByVal objCotacao As ClassCotacaoGR) As Long
'Realiza a gravacao efetiva da cotacao no bd

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim alComando(0 To 1) As Long
Dim sObsDestOrigem1 As String
Dim sObsDestOrigem2 As String
Dim sObsDestOrigem3 As String
Dim sObsDestOrigem4 As String

On Error GoTo Erro_Cotacao_GravaGR

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97253
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 97254

    'Se a condicao de pagamento esta preenchida
    If objCotacao.iCondicaoPagto > 0 Then
        
        'Le e faz lock da condicao
        lErro = CF("Cotacao_CondPagto_LeLock", objCotacao)
        If lErro <> SUCESSO Then gError 97371
        
    End If
    
    'Se o vendedor esta preenchido
    If objCotacao.iCodVendedor > 0 Then
        
        'Le e faz lock do vendedor
        lErro = CF("Cotacao_Vendedor_LeLock", objCotacao)
        If lErro <> SUCESSO Then gError 97372
                   
    End If
    
    'Le e faz o lock do TipoEmbalagem
    lErro = CF("Cotacao_TipoEmbalagem_LeLock", objCotacao)
    If lErro <> SUCESSO Then gError 97373
           
    'Para cada CotacaoContainer, le e faz lock
    lErro = CF("Cotacao_CotacaoContainer_LeLock", objCotacao)
    If lErro <> SUCESSO Then gError 97374
              
    'Para cada CotacaoServico, le e faz lock
    lErro = CF("Cotacao_CotacaoServico_LeLock", objCotacao)
    If lErro <> SUCESSO Then gError 97375
        
    'Le a Cotacao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CotacaoGR WHERE Codigo=? AND FilialEmpresa=? ", 0, lNumIntDoc, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 97272
    
    'Posiciona na Cotacao
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97273
        
    'Se nao Encontrou entao eh Insercao
    If lErro <> AD_SQL_SUCESSO Then
        
        'insere a cotacao
        lErro = CF("Cotacao_Insere", objCotacao)
        If lErro <> SUCESSO Then gError 97294
        
        'insere os relacionamentos
        lErro = Cotacao_Insere_Relacionamentos(objCotacao)
        If lErro <> SUCESSO Then gError 97359
       
    Else
    'Se encontrou entao eh Alteracao
        'Cyntia
        sObsDestOrigem1 = Mid(objCotacao.sObsDestOrigem, 1, 250)
        sObsDestOrigem2 = Mid(objCotacao.sObsDestOrigem, 251, 250)
        sObsDestOrigem3 = Mid(objCotacao.sObsDestOrigem, 501, 250)
        sObsDestOrigem4 = Mid(objCotacao.sObsDestOrigem, 751, 250)
    
        'Altera a Cotacao
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE CotacaoGR SET NumIntDoc=?,FilialEmpresa=?,Codigo=?,Data=?,TipoOperacao=?,Cliente=?,Envio=?,EnvioComplemento=?," _
        & "CodVendedor=?,Indicacao=?,Observacao=?,DataPrevInicio=?,TipoEmbalagem=?,Ajudantes=?,Carga=?,CargaPorConta=?,Descarga=?,DescargaPorconta=?,Ova=?," _
        & "OvaPorConta=?,Desova=?,DesovaPorConta=?,CargaSolta=?,DescCargaSolta=?,CondicaoPagto=?,Situacao=?,Justificativa=?,ObsResultado=?,ValorMercadoria=?,Atendente=?,Responsavel=?,ObsDestOrigem1=?,ObsDestOrigem2=?,ObsDestOrigem3=?,ObsDestOrigem4=?", alComando(0), _
        lNumIntDoc, objCotacao.iFilialEmpresa, objCotacao.lCodigo, objCotacao.dtData, objCotacao.iTipoOperacao, objCotacao.sCliente, objCotacao.iEnvio, objCotacao.sEnvioComplemento, _
        objCotacao.iCodVendedor, objCotacao.sIndicacao, objCotacao.sObservacao, objCotacao.dtDataPrevInicio, objCotacao.iTipoEmbalagem, objCotacao.iAjudantes, objCotacao.iCarga, _
        objCotacao.iCargaPorConta, objCotacao.iDesCarga, objCotacao.iDesCargaPorConta, objCotacao.iOva, objCotacao.iOvaPorConta, objCotacao.iDesova, objCotacao.iDesovaPorConta, objCotacao.iCargaSolta, _
        objCotacao.sDescCargaSolta, objCotacao.iCondicaoPagto, objCotacao.iSituacao, objCotacao.iJustificativa, objCotacao.sObsResultado, objCotacao.dValorMerc, objCotacao.sAtendente, objCotacao.sResponsavel, sObsDestOrigem1, sObsDestOrigem2, sObsDestOrigem3, sObsDestOrigem4)
        
        If lErro <> SUCESSO Then gError 97305
        
        objCotacao.lNumIntDoc = lNumIntDoc
        
        'Deleta os relacionamentos
        lErro = Cotacao_Exclui_Relacionamentos(objCotacao)
        If lErro <> SUCESSO Then gError 97299
        
        'insere os relacionamentos
        lErro = Cotacao_Insere_Relacionamentos(objCotacao)
        If lErro <> SUCESSO Then gError 97300
         
    End If
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 97304

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cotacao_GravaGR = SUCESSO
    
    Exit Function
    
Erro_Cotacao_GravaGR:

    Cotacao_GravaGR = gErr
    
    Select Case gErr
    
        Case 97253
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97254
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                     
        Case 97272, 97273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO1", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
               
        Case 97304
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
               
        Case 97305
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACAO1", gErr, objCotacao.lCodigo, objCotacao.iFilialEmpresa)
                       
        Case 97294, 97299, 97300, 97359, 97371, 97372, 97373, 97374, 97375
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Call Transacao_Rollback

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Cotacao_Exclui_Relacionamentos(objCotacao As ClassCotacaoGR) As Long
'Faz a exclusao nas tabelas em que a cotacao
'em questao se encontra para manter a integridade do bd
'IMPORTANTE: Chamar esta funçao em transacao

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_Cotacao_Exclui_Relacionamentos

    'Excluir os contatos associados a cotacao
    lErro = CF("Cotacao_Exclui_Contato", objCotacao)
    If lErro <> SUCESSO Then gError 97364
    
    'Excluir as cotacoescontainer associados a cotacao se carga solta estiver desmarcada
    lErro = CF("Cotacao_Exclui_CotacaoContainer", objCotacao)
    If lErro <> SUCESSO Then gError 97365
    
    'Excluir as OrigensDestino associadas a cotacao
    lErro = CF("Cotacao_Exclui_CotacaoOrigemDestino", objCotacao)
    If lErro <> SUCESSO Then gError 97367
       
    'Excluir os Servicos associados a cotacao
    lErro = CF("Cotacao_Exclui_CotacaoServico", objCotacao)
    If lErro <> SUCESSO Then gError 97369
        
    Cotacao_Exclui_Relacionamentos = SUCESSO
        
    Exit Function
    
Erro_Cotacao_Exclui_Relacionamentos:

    Cotacao_Exclui_Relacionamentos = gErr
    
    Select Case gErr
        
        Case 97364, 97365, 97367, 97369
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Function

End Function

Private Function Cotacao_Insere_Relacionamentos(objCotacao As ClassCotacaoGR) As Long
'Faz a inclusao nas tabelas em que a cotacao
'em questao se encontra para manter a integridade do bd

Dim lErro As Long

On Error GoTo Erro_Cotacao_Insere_Relacionamentos

    'insere os registros na tabela de Contato
    lErro = CF("Cotacao_Insere_Contato", objCotacao)
    If lErro <> SUCESSO Then gError 97295
       
    'insere os registros na tabela de CotacaoOrigemDestino
    lErro = CF("Cotacao_Insere_CotacaoOrigemDestino", objCotacao)
    If lErro <> SUCESSO Then gError 97296
        
    'insere os registros na tabela de CotacaoContainer
    lErro = CF("Cotacao_Insere_CotacaoContainer", objCotacao)
    If lErro <> SUCESSO Then gError 97297
       
    'insere os registros na tabela de CotacaoServico
    lErro = CF("Cotacao_Insere_CotacaoServico", objCotacao)
    If lErro <> SUCESSO Then gError 97298

    Cotacao_Insere_Relacionamentos = SUCESSO

    Exit Function

Erro_Cotacao_Insere_Relacionamentos:

    Cotacao_Insere_Relacionamentos = gErr
    
    Select Case gErr
        
        Case 97295, 97296, 97297, 97298
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Function SolicitacaoServico_Exclui(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Exclui o Solicitação de Serviço

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNum As Long
Dim lCodigo As Long
Dim lEnderDestino As Long
Dim lEnderOrigem As Long

On Error GoTo Erro_SolicitacaoServico_Exclui

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98204
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98205
    
    'Seleciona a Solicitação de Serviço com o número e filial passados por parâmetro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, EnderecoDestino, EnderecoOrigem FROM SolicitacaoServico WHERE Numero = ? and FilialEmpresa = ?", 0, lNum, lEnderDestino, lEnderOrigem, objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98206

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98207
    
    'Se não encontrou --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 98208

    'Lock da Solicitação de Serviço
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98209
    
    'Verifica se a Solicitação está associada a algum Comprovante de Serviço
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM CompServGR WHERE NumIntDocOrigem = ? ", 0, lCodigo, lNum)
    If lErro <> AD_SQL_SUCESSO Then gError 98210

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98211
    
    'Se estiver associada --> erro
    If lErro = AD_SQL_SUCESSO Then gError 98212
           
    'exclui os serviços relacionados com a solicitação em questão
    lErro = CF("SolServServico_Exclui_Todas", objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98213
    
    'se contém -->Exclui
    lErro = CF("Endereco_Exclui", lEnderDestino)
    If lErro <> SUCESSO Then gError 98216
    
    'se contém -->Exclui
    lErro = CF("Endereco_Exclui", lEnderOrigem)
    If lErro <> SUCESSO Then gError 98217
    
    'Exclui a Solicitação de Serviço
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM SolicitacaoServico", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98214
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98215

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    SolicitacaoServico_Exclui = SUCESSO

    Exit Function

Erro_SolicitacaoServico_Exclui:

    SolicitacaoServico_Exclui = gErr

    Select Case gErr
    
        Case 98204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 98206, 98207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case 98208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_CADASTRADA", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)

        Case 98209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case 98210, 98211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
                             
        Case 98212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_ASSOCIADA_COMPSERV", gErr, lCodigo)
        
        Case 98213, 98216, 98217
        
        Case 98214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)

        Case 98215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function SolicitacaoServico_Grava(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Grava os dados recebidos em objSolicitacaoServico

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objSolServServico As ClassServico
Dim lEnderOrigem As Long
Dim lEnderDestino As Long
Dim lErro_Solicitacao As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long

On Error GoTo Erro_SolicitacaoServico_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98218

    'Abre Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98219
    Next
    
    'Faz o lock de todos registros das tabelas necessárias para a gravação
    lErro = CF("SolicitacaoServico_Lock", objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98263

    'Verifica no Banco de dados se Já existe Solicitação de Serviço com o Número e filial recebidos no obj
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, EnderecoOrigem, EnderecoDestino FROM SolicitacaoServico WHERE Numero=? AND FilialEmpresa=?", 0, lNumIntDoc, lEnderOrigem, lEnderDestino, objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98247

    lErro_Solicitacao = Comando_BuscarPrimeiro(alComando(1))
    If lErro_Solicitacao <> AD_SQL_SUCESSO And lErro_Solicitacao <> AD_SQL_SEM_DADOS Then gError 98248
        
    'Verifica se o endereço de origem já existe, senão...
    If lEnderOrigem = 0 Then
        'Gera um novo código para o endereço
        lErro = CF("Endereco_Automatico", lEnderOrigem)
        If lErro <> SUCESSO Then gError 98249
    End If
    
    'Carrega os dados do endereço em objSolicitacaoServico
    objSolicitacaoServico.objEnderecoOrigem.lCodigo = lEnderOrigem
    objSolicitacaoServico.lEnderecoOrigem = lEnderOrigem
    
    'Grava ou altera o endereço como código passado
    lErro = CF("Endereco_Grava", objSolicitacaoServico.objEnderecoOrigem)
    If lErro <> SUCESSO Then gError 98251
        
    'Verifica se o endereço do destino já existe, senão...
    If lEnderDestino = 0 Then
        'Gera um novo código para o endereço
        lErro = CF("Endereco_Automatico", lEnderDestino)
        If lErro <> SUCESSO Then gError 98252
    End If
    
    'Carrega os dados do endereço em objSolicitacaoServico
    objSolicitacaoServico.objEnderecoDestino.lCodigo = lEnderDestino
    objSolicitacaoServico.lEnderecoDestino = lEnderDestino
    
    'Grava ou altera o endereço como código passado
    lErro = CF("Endereco_Grava", objSolicitacaoServico.objEnderecoDestino)
    If lErro <> SUCESSO Then gError 98254
        
    'Se não Existe a solicitação de Serviço ...
    If lErro_Solicitacao = AD_SQL_SEM_DADOS Then
     
        'Gera um novo NumIntDoc
        lErro = CF("Config_ObterAutomatico_emtrans", "FatConfig", "NUM_PROX_NUMINTDOC_SOLICITACAO", "SolicitacaoServico", "NumIntDoc", lNumIntDoc)
        If lErro <> SUCESSO Then gError 98255
        
        'Joga para o obj
        objSolicitacaoServico.lNumIntDoc = lNumIntDoc
        objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
        
        'cria uma nova Solicitação de Serviço
        lErro = Comando_Executar(alComando(2), "INSERT INTO SolicitacaoServico (NumIntDoc, FilialEmpresa, Numero, CodTabPreco, Cliente, DataEmissao, NumReferencia, DataPedido, TipoOperacao, Despachante, Material, QuantMaterial, UM, ValorMercadoria, TipoEmbalagem, TipoContainer, CodProgNavio, Booking, DataPrevInicio, HoraPrevInicio, DataPrevFim, HoraPrevFim, Observacao, EnderecoOrigem, EnderecoDestino,Porto)" _
        & " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objSolicitacaoServico.lNumIntDoc, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero, objSolicitacaoServico.lCodTabPreco, objSolicitacaoServico.lCliente, objSolicitacaoServico.dtDataEmissao, objSolicitacaoServico.sNumReferencia, objSolicitacaoServico.dtDataPedido, objSolicitacaoServico.iTipoOperacao, objSolicitacaoServico.iDespachante, _
        objSolicitacaoServico.sMaterial, objSolicitacaoServico.dQuantMaterial, objSolicitacaoServico.sUM, objSolicitacaoServico.dValorMercadoria, objSolicitacaoServico.iTipoEmbalagem, objSolicitacaoServico.iTipoContainer, objSolicitacaoServico.lCodProgNavio, objSolicitacaoServico.sBooking, objSolicitacaoServico.dtDataPrevInicio, CDbl(objSolicitacaoServico.dtHoraPrevInicio), objSolicitacaoServico.dtDataPrevFim, CDbl(objSolicitacaoServico.dtHoraPrevFim), _
        objSolicitacaoServico.sObservacao, objSolicitacaoServico.lEnderecoOrigem, objSolicitacaoServico.lEnderecoDestino, objSolicitacaoServico.sPorto)
        If lErro <> AD_SQL_SUCESSO Then gError 98256
                                            
        'Inclui os dados dos serviços relacionados com a solicitação
        lErro = CF("SolServServico_Inclui", objSolicitacaoServico.ColServico, objSolicitacaoServico)
        If lErro <> SUCESSO Then gError 98257
            
    Else
        
        'Carrega o obj com o NumIntDoc recolhido no banco de dados
        objSolicitacaoServico.lNumIntDoc = lNumIntDoc
            
        'Verificar se existem comprovantes de servico relacionados a esta solicitacao
        lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM CompServGR WHERE NumIntDocOrigem=? ", lCodigo, objSolicitacaoServico.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 98193
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98194
        
        If lErro = AD_SQL_SEM_DADOS Then
                    
            'altera uma nova Solicitação de Serviço já existente
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE SolicitacaoServico SET CodTabPreco=?, Cliente=?, DataEmissao=?, NumReferencia=?, DataPedido=?, TipoOperacao=?, Despachante=?, Material=?, QuantMaterial=?, UM=?, ValorMercadoria=?, TipoEmbalagem=?, TipoContainer=?, CodProgNavio=?, Booking=?, DataPrevInicio=?, HoraPrevInicio=?, DataPrevFim=?, HoraPrevFim=?, Observacao=?, EnderecoOrigem=?, EnderecoDestino=?,Porto=?", _
            alComando(1), objSolicitacaoServico.lCodTabPreco, objSolicitacaoServico.lCliente, objSolicitacaoServico.dtDataEmissao, objSolicitacaoServico.sNumReferencia, objSolicitacaoServico.dtDataPedido, objSolicitacaoServico.iTipoOperacao, objSolicitacaoServico.iDespachante, objSolicitacaoServico.sMaterial, objSolicitacaoServico.dQuantMaterial, objSolicitacaoServico.sUM, objSolicitacaoServico.dValorMercadoria, _
            objSolicitacaoServico.iTipoEmbalagem, objSolicitacaoServico.iTipoContainer, objSolicitacaoServico.lCodProgNavio, objSolicitacaoServico.sBooking, objSolicitacaoServico.dtDataPrevInicio, CDbl(objSolicitacaoServico.dtHoraPrevInicio), objSolicitacaoServico.dtDataPrevFim, CDbl(objSolicitacaoServico.dtHoraPrevFim), objSolicitacaoServico.sObservacao, objSolicitacaoServico.lEnderecoOrigem, objSolicitacaoServico.lEnderecoDestino, objSolicitacaoServico.sPorto)
            If lErro <> AD_SQL_SUCESSO Then gError 98258
            
            'Exclui os dados do serviço relacionados com a solicitação
            lErro = CF("SolServServico_Exclui_Todas", objSolicitacaoServico)
            If lErro <> SUCESSO Then gError 98259
                                                
            'Inclui os novos dados do serviço relacionados com a solicitação
            lErro = CF("SolServServico_Inclui", objSolicitacaoServico.ColServico, objSolicitacaoServico)
            If lErro <> SUCESSO Then gError 98260
            
       'Se existem comprovantes relacionados ...
        ElseIf lErro = AD_SQL_SUCESSO Then
            
            lErro = CF("Verifica_Servico", objSolicitacaoServico)
            If lErro <> SUCESSO Then gError 98307
            
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE SolicitacaoServico SET NumReferencia=?, DataPedido=?, DataPrevInicio=?, HoraPrevInicio=?, DataPrevFim=?, HoraPrevFim=?, Observacao=?, EnderecoOrigem=?, EnderecoDestino=?, Material=?, QuantMaterial=?, UM=?, ValorMercadoria=?, TipoEmbalagem=?,Porto=?,CodProgNavio=?, Booking=?", _
            alComando(1), objSolicitacaoServico.sNumReferencia, objSolicitacaoServico.dtDataPedido, objSolicitacaoServico.dtDataPrevInicio, CDbl(objSolicitacaoServico.dtHoraPrevInicio), objSolicitacaoServico.dtDataPrevFim, _
            CDbl(objSolicitacaoServico.dtHoraPrevFim), objSolicitacaoServico.sObservacao, objSolicitacaoServico.lEnderecoOrigem, objSolicitacaoServico.lEnderecoDestino, objSolicitacaoServico.sMaterial, objSolicitacaoServico.dQuantMaterial, objSolicitacaoServico.sUM, objSolicitacaoServico.dValorMercadoria, objSolicitacaoServico.iTipoEmbalagem, objSolicitacaoServico.sPorto, objSolicitacaoServico.lCodProgNavio, objSolicitacaoServico.sBooking)
            If lErro <> AD_SQL_SUCESSO Then gError 98308
            
        End If

    End If
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98261

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
        
    SolicitacaoServico_Grava = SUCESSO

    Exit Function

Erro_SolicitacaoServico_Grava:

    SolicitacaoServico_Grava = gErr

    Select Case gErr

        Case 98218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 98219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98193, 98194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV", gErr, objSolicitacaoServico.lNumIntDoc)
               
        Case 98247, 98248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case 98249, 98251, 98252, 98254, 98255, 98257, 98259, 98260, 98263, 98307
        
        Case 98256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case 98258, 98308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SOLICITACAO", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case 98261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function SolServServico_Exclui_Todas(ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
 'Exclui os Itens da Tabela de SolServServicos
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_SolServServico_Exclui_Todas

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98265
    Next
        
    'Lê e Faz "lock" de SolServServico do numintdoc da solicitação
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocOrigem FROM SolServServico WHERE NumIntDocOrigem = ? ", 0, lNumIntDoc, objSolicitacaoServico.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 98266

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98267
            
    'Enquanto encontrar serviço...
    Do While lErro = AD_SQL_SUCESSO

        'Faz "LockExclusive" de SolServServico
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98268

        'Exclui o SolServServico do banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM SolServServico", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98269

        'Busca próximo Item
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98270

    Loop
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    SolServServico_Exclui_Todas = SUCESSO

    Exit Function

Erro_SolServServico_Exclui_Todas:

    SolServServico_Exclui_Todas = gErr

    Select Case gErr

        Case 98265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98266, 98267, 98270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)

        Case 98268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)

        Case 98269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)
            
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function SolServServico_Exclui(ByVal ColServico As Collection, ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Exclui os Itens da tabela SolServServico no BD
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objSolServServico As ClassServico
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long

On Error GoTo Erro_SolServServico_Exclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98314
    Next
    
    'Para cada item do serviço
    For Each objSolServServico In ColServico
                     
        'Lê e Faz "lock" de SolServServico relacionado ao produto
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocOrigem FROM SolServServico WHERE Produto = ? AND NumIntDocOrigem = ?", 0, lNumIntDoc, objSolServServico.sProduto, objSolicitacaoServico.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 98315
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98316
        
        'Faz "LockExclusive" de SolServServico
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98317

        'Exclui o Item na tabela de SolServServico no BD
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM SolServServico", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98318
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(lComando)
    Next
    
    SolServServico_Exclui = SUCESSO

    Exit Function

Erro_SolServServico_Exclui:

    SolServServico_Exclui = gErr

    Select Case gErr
        
        Case 98314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98315, 98316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)

        Case 98317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)

        Case 98318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(lComando)
    Next
    
    Exit Function

End Function

Function SolServServico_Inclui(ByVal ColServico As Collection, ByVal objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Insere os Itens da tabela SolServServico no BD
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objSolServServico As ClassServico

On Error GoTo Erro_SolServServico_Inclui

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98271
        
    'Para cada item do serviço da Tabela de Solicitação
    For Each objSolServServico In ColServico
             
        'Grava o Item na tabela de SolServServico no BD
        lErro = Comando_Executar(lComando, "INSERT INTO SolServServico (NumIntDocOrigem, Produto, Quantidade) VALUES (?,?,?)", objSolicitacaoServico.lNumIntDoc, objSolServServico.sProduto, objSolServServico.dQuant)
        If lErro <> AD_SQL_SUCESSO Then gError 98272
        
    Next
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    SolServServico_Inclui = SUCESSO

    Exit Function

Erro_SolServServico_Inclui:

    SolServServico_Inclui = gErr

    Select Case gErr
        
        Case 98271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SOLSERVSERVICO", gErr, objSolicitacaoServico.lNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function CompServGR_Exclui(ByVal objCompServ As ClassCompServ) As Long
'Exclui o CompServ

Dim lErro As Long, lTransacao As Long, alComando(0 To 2) As Long, lNumIntDoc As Long, lNumIntNota As Long
Dim iIndice As Integer

On Error GoTo Erro_CompServGR_Exclui

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98521
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98522

    'Le o Comprovante passado como parametro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntNota FROM CompServGR WHERE Codigo=? AND FilialEmpresa=? ", 0, lNumIntDoc, lNumIntNota, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98523
    
    'posiciona no registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98524
    
    'se numintnota estiver preenchido, nao pode excluir
    If lNumIntNota <> 0 Then gError 97841
    
    'Lock Exclusive do CompServ
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98525
    
    objCompServ.lNumIntDoc = lNumIntDoc
    
    'Exclui os itens relacionados com o compserv em questao na
    'tabela intermediaria do relacionamento
    lErro = CompServ_Exclui_CompServItem(objCompServ)
    If lErro <> SUCESSO Then gError 98526
    
    'Exclui o Comprovante
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CompServGR", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 98527

    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 98528

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CompServGR_Exclui = SUCESSO
    
    Exit Function
    
Erro_CompServGR_Exclui:

    CompServGR_Exclui = gErr
    
    Select Case gErr
    
        Case 98521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 98522
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 98523, 98524, 98526
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV3", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        
        Case 98525
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMPSERV", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
                   
        Case 98527
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMPSERV", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
            
        Case 98528
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 97841
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_JA_FOI_FATURADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function CompServ_Exclui_CompServItem(ByVal objCompServ As ClassCompServ) As Long
'Exclui a relacao Itens de Serviço X o Comprovante em questao
'IMPORTANTE: Chamar esta funçao em transacao

Dim lErro As Long, alComando(0 To 1) As Long, lNumIntDoc As Long
Dim iIndice As Integer

On Error GoTo Erro_CompServ_Exclui_CompServItem

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98566
    Next

    'Busca Item associado ao Comprovante
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CompServItemGR WHERE NumIntDocOrigem = ?", 0, lNumIntDoc, objCompServ.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 98567

    'posiciona no registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98568

    Do While lErro = AD_SQL_SUCESSO

        'Exclui Relacao Entre ItemServ e CompServ
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CompServItemGR", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 98569

        'busca a proxima relacao...
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98570

    Loop

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    CompServ_Exclui_CompServItem = SUCESSO

    Exit Function

Erro_CompServ_Exclui_CompServItem:

    CompServ_Exclui_CompServItem = gErr

    Select Case gErr

        Case 98566
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98567, 97568, 98570
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERVITEM", gErr)

        Case 98569
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMPSERVITEM", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CompServGR_Grava(ByVal objCompServ As ClassCompServ) As Long
'Realiza a gravacao efetiva do Comprovante de Servico no bd

Dim lErro As Long, lTransacao As Long, lNumIntDoc As Long, alComando(0 To 4) As Long, lNumIntNota As Long
Dim iIndice As Integer
Dim dQtd As Double, dTotal As Double
Dim objSolServ As New ClassSolicitacaoServico
Dim sProd As String
Dim iFlag As Integer
Dim iSituacao As Integer

On Error GoTo Erro_CompServGR_Grava

    'inicializando strings
    sProd = String(STRING_PRODUTO, 0)
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 98543
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 98544

    'Obtem o total (quantidade)
    lErro = Comando_Executar(alComando(3), "SELECT Quantidade FROM SolServServico WHERE Produto=? AND NumIntDocOrigem=?", dTotal, objCompServ.sProduto, objCompServ.lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 98707
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98708

    'Verificar o quanto falta ainda para chegar na quantidade da solicitacao
    lErro = Comando_Executar(alComando(4), "SELECT SUM(Quantidade) FROM CompServGR WHERE Produto=? AND NumIntDocOrigem=? AND Codigo <> ? AND FilialEmpresa = ?", dQtd, objCompServ.sProduto, objCompServ.lNumIntDocOrigem, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98681
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98682
    
    'Coloca em QTD quanto falta para completar
    dQtd = dTotal - dQtd
            
    'Se qtd for negativa, ja estourou... erro no banco
'    If dQtd < 0 Then gError 98683
    
    'Verifica se estourou...
    If dQtd <= 0 Then gError 98684
    
    'Se quantidade a ser gravada vai estourar....
    If dQtd < objCompServ.dQuantidade Then gError 98685
    
    'Faz o Lock da SolServ
    objSolServ.lNumIntDoc = objCompServ.lNumIntDocOrigem
    
    lErro = CF("SolServ_Lock", objSolServ)
    If lErro <> SUCESSO Then gError 98623
    
    'Faz o lock dos itens de servico
    lErro = CF("ItemServico_Lock2", objCompServ)
    If lErro <> SUCESSO Then gError 98624
    
    'Le o compservgr
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Situacao, NumIntDoc, Produto, NumIntNota FROM CompServGR WHERE Codigo=? AND FilialEmpresa=? ", 0, iSituacao, lNumIntDoc, sProd, lNumIntNota, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 98625
        
    'Posiciona no registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98626
        
    'Se nao Encontrou entao eh Insercao
    If lErro <> AD_SQL_SUCESSO Then
        
        'insere o CompServ
        lErro = CompServ_Insere(objCompServ)
        If lErro <> SUCESSO Then gError 98627
        
        'insere os itens do compserv
        lErro = CompServ_Insere_CompServItens(objCompServ)
        If lErro <> SUCESSO Then gError 98628
       
    Else
    'Se encontrou entao eh Alteracao
        
        'início trecho comentado por Rafael Menezes em 01/10/2002
        'If lNumIntNota > 0 Then gError 98732
        'fim trecho comentado por Rafael Menezes em 01/10/2002
        
        If iSituacao = STATUS_CONCLUIDO Then gError 109112
        
        If objSolServ.lNumIntDoc <> objCompServ.lNumIntDocOrigem Then gError 98733
        
        If sProd <> objCompServ.sProduto Then gError 98734
        
        objCompServ.lNumIntDoc = lNumIntDoc
        
        'Altera o CompServ
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CompServGR SET DataEmissao=?,QuantMaterial=?,UM=?,ValorMercadoria=?, " _
        & "FretePeso=?,Pedagio=?,AdValoren=?,DataDemurrage=?,CodigoContainer=?,ValorContainer=?,Tara=?,Lacre=?,Observacao=?,Situacao=?,Quantidade=?" _
        , alComando(1), objCompServ.dtDataEmissao, objCompServ.dQuantMaterial, objCompServ.sUM, _
        objCompServ.dValorMercadoria, objCompServ.dFretePeso, objCompServ.dPedagio, objCompServ.dAdValoren, objCompServ.dtDataDemurrage, objCompServ.sCodigoContainer, objCompServ.dValorContainer, objCompServ.dTara, objCompServ.sLacre, objCompServ.sObservacao, objCompServ.iSituacao, objCompServ.dQuantidade)
        If lErro <> SUCESSO Then gError 98629
        
        'Deleta os itens relacionados ao compserv
        lErro = CF("CompServ_Exclui_CompServItem", objCompServ)
        If lErro <> SUCESSO Then gError 98630
        
        'insere os novos itens...
        lErro = CompServ_Insere_CompServItens(objCompServ)
        If lErro <> SUCESSO Then gError 98631
         
    End If
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 98632

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CompServGR_Grava = SUCESSO
    
    Exit Function

Erro_CompServGR_Grava:

    CompServGR_Grava = gErr

    Select Case gErr

        Case 98543
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 98544
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 98623, 98624, 98627, 98628, 98630, 98631
        
        Case 98625, 98626, 98681, 98682
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)

        Case 98629
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPSERV3", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        
        Case 98632
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        'Case 98683
        '    Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_EXTRAPOLOU_QTD_SOLSERVSERVICO", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa, Abs(dQtd))
    
        Case 98684
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_TOTALMENTE_COMPROVADO", gErr)
        
        Case 98685
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_COMPSERV_VAI_ESTOURAR", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa, objCompServ.dQuantidade)
        
        Case 98707, 98708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SOLSERVSERVICO1", gErr)
        
        'início trecho comentado por Rafael Menezes em 01/10/2002
        'Case 98732
        '    Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_JA_FOI_FATURADO", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        'fim trecho comentado por Rafael Menezes em 01/10/2002

        
        Case 98733
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLSERV_NAO_PODE_SER_ALTERADA", gErr)
        
        Case 98734
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_ALTERADO", gErr)
            
        Case 109112
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPROVANTE_JA_CONCLUIDO", gErr, objCompServ.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CompServ_Insere(ByVal objCompServ As ClassCompServ) As Long
'Faz a insercao do comprovante de servico, gera o numintdoc do mesmo.

Dim lErro As Long, lComando As Long, lNumIntDoc As Long
Dim sCamposCompServGR As String
        
On Error GoTo Erro_CompServ_Insere
        
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98633
            
    'guarda na string os campos para facilitar no select...
    sCamposCompServGR = "NumIntDoc,NumIntDocOrigem,FilialEmpresa, Codigo, DataEmissao, Produto, " _
    & "QuantMaterial, UM, ValorMercadoria, FretePeso, Pedagio, AdValoren, DataDemurrage, CodigoContainer, ValorContainer, Tara, " _
    & "Lacre, Observacao, Situacao, Quantidade, NumIntNota"
       
    'Pegar o numintdoc
    lErro = CF("Config_ObterNumInt", "FatConfig", "NUMINTDOC_PROX_COMPSERV", lNumIntDoc)
    If lErro <> SUCESSO Then gError 98634
    
    'insere o registro na tabela CompServGR
    lErro = Comando_Executar(lComando, "INSERT INTO CompServGR (" & sCamposCompServGR & ") values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
    lNumIntDoc, objCompServ.lNumIntDocOrigem, objCompServ.iFilialEmpresa, objCompServ.lCodigo, objCompServ.dtDataEmissao, objCompServ.sProduto, objCompServ.dQuantMaterial, _
    objCompServ.sUM, objCompServ.dValorMercadoria, objCompServ.dFretePeso, objCompServ.dPedagio, objCompServ.dAdValoren, objCompServ.dtDataDemurrage, objCompServ.sCodigoContainer, objCompServ.dValorContainer, objCompServ.dTara, objCompServ.sLacre, objCompServ.sObservacao, objCompServ.iSituacao, objCompServ.dQuantidade, objCompServ.lNumIntNota)
    If lErro <> AD_SQL_SUCESSO Then gError 98635

    objCompServ.lNumIntDoc = lNumIntDoc
       
    'fechar comando
    Call Comando_Fechar(lComando)
    
    CompServ_Insere = SUCESSO

    Exit Function

Erro_CompServ_Insere:

    CompServ_Insere = gErr
    
    Select Case gErr
    
        Case 98633
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98634
                    
        Case 98635
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMPSERV", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select
    
    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CompServ_Insere_CompServItens(ByVal objCompServ As ClassCompServ) As Long
'Faz a inclusao dos Itens Relacionados ao Comprovante

Dim lErro As Long, lComando As Long, lNumIntDoc As Long
Dim objCompServItem As ClassCompServItem
Dim sCamposCompServItemGR As String
        
On Error GoTo Erro_CompServ_Insere_CompServItens

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98640
    
    'guarda na string os campos para facilitar no select...
    sCamposCompServItemGR = "NumIntDoc, NumIntDocOrigem, CodItemServico, DataPrev, HoraPrev, DataInicio, HoraInicio, DataFim, HoraFim, " _
    & "DocIntTipo, DocIntNumero, DocIntDataEmissao, DocExtTipo, DocExtNumero, DocExtDataEmissao, DocExtDataRec, DocExtHoraRec, PlacaCaminhao, PlacaCarreta, Motorista, Observacao"
    
    For Each objCompServItem In objCompServ.colCompServItem
    
        'Pegar o numintdoc do Item a ser gravado
        lErro = CF("Config_ObterNumInt", "FatConfig", "NUMINTDOC_PROX_COMPSERVITEM", lNumIntDoc)
        If lErro <> SUCESSO Then gError 98641
        
        lErro = Comando_Executar(lComando, "INSERT INTO CompServItemGR (" & sCamposCompServItemGR & ") values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        lNumIntDoc, objCompServ.lNumIntDoc, objCompServItem.iCodItemServico, objCompServItem.dtDataPrev, CDbl(objCompServItem.dtHoraPrev), objCompServItem.dtDataInicio, CDbl(objCompServItem.dtHoraInicio), _
        objCompServItem.dtDataFim, CDbl(objCompServItem.dtHoraFim), objCompServItem.iDocIntTipo, objCompServItem.sDocIntNumero, objCompServItem.dtDocIntDataEmissao, objCompServItem.iDocExtTipo, _
        objCompServItem.sDocExtNumero, objCompServItem.dtDocExtDataEmissao, objCompServItem.dtDocExtDataRec, CDbl(objCompServItem.dtDocExtHoraRec), objCompServItem.sPlacaCaminhao, objCompServItem.sPlacaCarreta, objCompServItem.sMotorista, objCompServItem.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 98642

    Next

    'fechar comando
    Call Comando_Fechar(lComando)
    
    CompServ_Insere_CompServItens = SUCESSO

    Exit Function

Erro_CompServ_Insere_CompServItens:

    CompServ_Insere_CompServItens = gErr
    
    Select Case gErr

        Case 98640
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98641
                      
        Case 98642
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMPSERVITEM", gErr, objCompServ.lCodigo, objCompServ.iFilialEmpresa, objCompServItem.iCodItemServico)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select

    'fechar comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'mario
Public Function NFiscal_Grava_NFiscal_Cust(ByVal objNFiscal As ClassNFiscal) As Long
'Grava a parte customizada da nota fiscal

Dim lErro As Long
Dim objCompNF As ClassCompServ
Dim lCodigo As Long, lNumIntNota As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_NFiscal_Grava_NFiscal_Cust

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97513
    Next
    
    'Para cada Comprovante da Nota Fiscal
    For Each objCompNF In objNFiscal.colComprovServ
               
        'Atualiza a tabela de comprovante com o numintnota
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, NumIntNota FROM CompServGR WHERE Codigo = ?", 0, lCodigo, lNumIntNota, objCompNF.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 99096
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99097
        
        'Se não encontrou o comprovante --> erro.
        If lErro = AD_SQL_SEM_DADOS Then gError 99098
        
        If lNumIntNota <> 0 Then gError 99176
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CompServGR SET NumIntNota = ?", alComando(1), objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 99099
        
    Next
            
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Grava_NFiscal_Cust = SUCESSO

    Exit Function

Erro_NFiscal_Grava_NFiscal_Cust:
    
    NFiscal_Grava_NFiscal_Cust = gErr

    Select Case gErr
        
        Case 97513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 99096, 99097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
            
        Case 99098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompNF.lCodigo, giFilialEmpresa)
            
        Case 99099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPROVANTESSERVICOS", gErr)
            
        Case 99176
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_JA_FATURADO", gErr, lCodigo)
     
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

'mario
Public Function NFiscal_Lock_Gravacao_Cust(ByVal objNFiscal As ClassNFiscal) As Long
'executa a parte customizadas dos lock dos registros associados a nota fiscal.

Dim lErro As Long
Dim iIndice As Long
Dim objCompNF As New ClassCompServ
Dim lCodigo As Long
Dim lComando As Long

On Error GoTo Erro_NFiscal_Lock_Gravacao_Cust

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97513
        
    'Lock shared nos Comprovantes
    For Each objCompNF In objNFiscal.colComprovServ
        
        'Seleciona cada comprovante usado
        lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM CompServGR WHERE Codigo = ?", lCodigo, objCompNF.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 99100
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99101
        
        'Se não encontrou o comprovante --> erro.
        If lErro = AD_SQL_SEM_DADOS Then gError 99102
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 99103
        
    Next
    
    Call Comando_Fechar(lComando)
    
    NFiscal_Lock_Gravacao_Cust = SUCESSO

    Exit Function

Erro_NFiscal_Lock_Gravacao_Cust:

    NFiscal_Lock_Gravacao_Cust = gErr

    Select Case gErr

        Case 97513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 99100, 99101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
            
        Case 99102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, objCompNF.lCodigo, giFilialEmpresa)
            
        Case 99103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMPSERV", gErr, objCompNF.lCodigo, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function NotaFiscalSaida_Excluir_Usu(ByVal objNFiscal As ClassNFiscal) As Long
    
Dim lErro As Long

On Error GoTo Erro_NotaFiscalSaida_Excluir_Usu

    'Atualiza os comprovantes com NumintNota = 0
    lErro = CF("Atualiza_Comprovante", objNFiscal)
    If lErro <> SUCESSO Then gError 99106
    
    NotaFiscalSaida_Excluir_Usu = SUCESSO
    
    Exit Function
    
Erro_NotaFiscalSaida_Excluir_Usu:

    NotaFiscalSaida_Excluir_Usu = gErr
    
    Select Case gErr

        Case 99106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select

End Function
    
Public Function Atualiza_Comprovante(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lCodigo As Long

On Error GoTo Erro_Atualiza_comprovante
    
    'Abre o comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 99104
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 99105
    
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM CompServGR WHERE NumIntNota = ?", 0, lCodigo, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 99107
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99108
    
'    If lErro = AD_SQL_SEM_DADOS Then gError 99109

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(lComando2, "UPDATE CompServGR SET NumIntNota=?", lComando1, NUMINTNOTA_COMPROVANTE)
        If lErro <> AD_SQL_SUCESSO Then gError 99110

        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 99124

    Loop
    
    'fechar comando
    Call Comando_Fechar(lComando1)
    
    Call Comando_Fechar(lComando2)
    
    Atualiza_Comprovante = SUCESSO

    Exit Function

Erro_Atualiza_comprovante:

    Atualiza_Comprovante = gErr
    
    Select Case gErr

        Case 99104, 99105
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 99107, 99108, 99124
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPSERV1", gErr)
            
        Case 99109
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPSERV_NAO_CADASTRADO", gErr, lCodigo, giFilialEmpresa)
                    
        Case 99110
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMPSERV", gErr, lCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
                
    End Select

    'fechar comando
    Call Comando_Fechar(lComando1)
    
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

    
    

