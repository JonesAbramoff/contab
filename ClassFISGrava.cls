VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Const MODREGFIS_SIMPLES_FAT = 9000 'modelo de reg arq icms p/nota de simples faturamento
Const MODREGFIS_FAT_ANTECIP = 9001 'modelo de reg arq icms p/nota de faturamento antecipado
Const MODREGFIS_DEV_SIMB_CONSIG = 9002 'modelo de reg arq icms p/nota de saída por devolução simbólica de consignação
Const MODREGFIS_DEV_SIMPLES_FAT = 9003 'p/devolucao de simples fat de consign
Const MODREGFIS_FAT_RESSARC_ST = 9004 'p/ressarcimento icms substituicao tributaria
Const MODREGFIS_SIMPLES_FAT_COMPRA_CONTA_E_ORDEM = 9005 'Compra à ordem - simples faturamento
Const MODREGFIS_SIMPLES_FAT_VENDA_CONTA_E_ORDEM = 9006 'Venda à ordem - simples faturamento
Const MODREGFIS_COMPRA_CONTA_E_ORDEM = 9007 'Compra à ordem
Const MODREGFIS_IMPORTACAO_PARCELADA = 9008 '"remessas parceladas de que trata a nota fiscal (entrada) nº ......."
Const MODREGFIS_ANULACAO_SERVICO_TRANSPORTE = 9009

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Arquivos_Exclui_Modelo_EmTrans(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Exclui os arquivos relacionados ao modelo passado como parâmetro
'objIN86Modelo RECEBE(Input) o código do modelo que terá seus arquivos excluídos
'DEVE SER CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iModelo As Integer

On Error GoTo Erro_IN86Arquivos_Exclui_Modelo_EmTrans

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103542
    Next

    'Lê os arquivos do modelo passado como parâmetro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Modelo FROM IN86Arquivos WHERE Modelo = ?", 0, iModelo, objIN86Modelo.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103543

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103544

    'Se não encontrou nenhum arquivo => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103605
    
    'Enquanto houver arquivos
    Do While lErro <> AD_SQL_SEM_DADOS

        'Deleta o arquivo do BD
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM IN86Arquivos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 103545

        'busca o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103546

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    IN86Arquivos_Exclui_Modelo_EmTrans = SUCESSO

    Exit Function

Erro_IN86Arquivos_Exclui_Modelo_EmTrans:

    IN86Arquivos_Exclui_Modelo_EmTrans = gErr

    Select Case gErr
    
        Case 103542
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103605
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86ARQUIVOS_NAO_ENCONTRADO", gErr)

        Case 103543, 103544, 103546
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86ARQUIVOS_LEITURA", gErr)
        
        Case 103545
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86ARQUIVOS_EXCLUSAO", gErr, objIN86Modelo.sModelo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150294)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Arquivos_Grava_EmTrans(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Grava os arquivos passados como parâmetro
'objIN86Modelo RECEBE(Input) os arquivos que serão gravados e o código do modelo relacionado aos arquivos
'DEVE SER CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim objIN86Arquivo As ClassIN86Arquivos
Dim lComando As Long

On Error GoTo Erro_IN86Arquivos_Grava_EmTrans

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103547

    'Para cada arquivo a ser gravado
    For Each objIN86Arquivo In objIN86Modelo.colIN86Arquivos

        'Insere o arquivo no BD
        lErro = Comando_Executar(lComando, "INSERT INTO IN86Arquivos (Modelo,Tipo,Selecionado,Nome, RelatDUMP, RelatAcompanhamento, Layout,FilialEmpresa, NumEtiqueta) VALUES (?,?,?,?,?,?,?,?,?)", _
        objIN86Modelo.iCodigo, objIN86Arquivo.iTipo, objIN86Arquivo.iSelecionado, objIN86Arquivo.sNome, objIN86Arquivo.iDUMP, objIN86Arquivo.iRelatAcompanhamento, objIN86Arquivo.iLayout, objIN86Arquivo.iFilialEmpresa, objIN86Arquivo.iNumEtiqueta)
        If lErro <> AD_SQL_SUCESSO Then gError 103548

    Next

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86Arquivos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_IN86Arquivos_Grava_EmTrans:

    IN86Arquivos_Grava_EmTrans = gErr

    Select Case gErr

        Case 103547
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103548
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86ARQUIVOS_INSERCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150295)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Modelo_Grava_BD_EmTrans(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Grava o modelo para geração do IN86 e as configurações dos arquivos desse modelo
'objIN86Modelo RECEBE(Input) os dados do modelo que será gravado

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_IN86Modelo_Grava_BD_EmTrans

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103549

    'Insere o modelo no BD
    lErro = Comando_Executar(lComando, "INSERT INTO IN86Modelos(Codigo, Modelo, DataInicio, DataFim, MeioEntrega, Etiquetas)VALUES( ? , ? , ? , ? , ? , ? )", _
    objIN86Modelo.iCodigo, objIN86Modelo.sModelo, objIN86Modelo.dtDataInicio, objIN86Modelo.dtDataFim, objIN86Modelo.iMeioEntrega, objIN86Modelo.iEtiquetas)
    If lErro <> AD_SQL_SUCESSO Then gError 103550

    'Grava os arquivos relacionados ao modelo
    lErro = CF("IN86Arquivos_Grava_EmTrans", objIN86Modelo)
    If lErro <> SUCESSO Then gError 103551

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)

    IN86Modelo_Grava_BD_EmTrans = SUCESSO

    Exit Function

Erro_IN86Modelo_Grava_BD_EmTrans:

    IN86Modelo_Grava_BD_EmTrans = gErr

    Select Case gErr

        Case 103549
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103550
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_INSERCAO", gErr)

        Case 103551

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150296)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Modelo_Grava(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Dispara a gravação de um modelo para geração do IN86
'objIN86Modelo RECEBE(Input) os dados do modelo que será gravado

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_IN86Modelo_Grava

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103552
    Next

    'Executa a abertura da transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 103553

    'se o código do modelo está preenchido, ou seja, se é um modelo já existente que está sendo alterado
    If objIN86Modelo.iCodigo <> 0 Then

        'tenta encontrá-lo no Banco de Dados
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM IN86Modelos WHERE Codigo = ?", 0, lCodigo, objIN86Modelo.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 103554

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103555

        'se não encontrou->Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 103556

    'se o código não está preenchido, ou é inserção d eum novo modelo ou alteração de um modelo que não estava carregado na combo
    Else

        'tenta encontrá-lo no Banco de Dados pelo nome
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM IN86Modelos WHERE Modelo = ?", 0, lCodigo, objIN86Modelo.sModelo)
        If lErro <> AD_SQL_SUCESSO Then gError 103557

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103558

    End If
    
    'se encontrou o modelo no BD
    If lErro = AD_SQL_SUCESSO Then

        'Atualiza-o
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE IN86Modelos SET DataInicio=? , DataFim=? , MeioEntrega = ? , Etiquetas = ?", alComando(0), objIN86Modelo.dtDataInicio, objIN86Modelo.dtDataFim, objIN86Modelo.iMeioEntrega, objIN86Modelo.iEtiquetas)
        If lErro <> AD_SQL_SUCESSO Then gError 103559

        'Exclui os arquivos relacionados
        lErro = IN86Arquivos_Exclui_Modelo_EmTrans(objIN86Modelo)
        If lErro <> SUCESSO Then gError 103560

        'Insere os novos arquivos do modelo em questão
        lErro = IN86Arquivos_Grava_EmTrans(objIN86Modelo)
        If lErro <> SUCESSO Then gError 103561

    'se não encontrá-lo
    Else

        'gera o próximo código de modelo
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_PROX_IN86MODELO", lCodigo)
        If lErro <> SUCESSO Then gError 103562
        
        'carrega o objeto com o codigo gerado
        objIN86Modelo.iCodigo = lCodigo

        'grava o novo modelo no BD
        lErro = IN86Modelo_Grava_BD_EmTrans(objIN86Modelo)
        If lErro <> SUCESSO Then gError 103563

    End If

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Efetua o commit da transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 103564

    IN86Modelo_Grava = SUCESSO

    Exit Function

Erro_IN86Modelo_Grava:

    IN86Modelo_Grava = gErr

    Select Case gErr

        Case 103552
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103553
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103554, 103555, 103557, 103558
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_LEITURA", gErr)

        Case 103559
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_ALTERACAO", gErr, objIN86Modelo.sModelo)
            
        Case 103556
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_NAO_ENCONTRADO", gErr)

        Case 103560 To 103563

        Case 103564
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150297)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Efetua rollback da transacao
    Call Transacao_Rollback
    
    Exit Function
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function IN86Modelos_Exclui(ByVal objIN86Modelo As ClassIN86Modelos) As Long
'Dispara a exclusão de um modelo para geração do IN86
'objIN86Modelo RECEBE(Input) o código do modelo que será excluído

Dim lErro As Long
Dim alComando(1) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_IN86Modelos_Exclui

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103574
    Next

    'Executa a abertura da transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 103575

    'seleciona o modelo pelo código
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM IN86Modelos WHERE Codigo=?", 0, iCodigo, objIN86Modelo.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103576

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103577

    'se não achou--> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103578
    
    'efetua lock do modelo
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 127001

    'exclui os arquivos vinculados ao modelo selecionado para exclusão
    lErro = IN86Arquivos_Exclui_Modelo_EmTrans(objIN86Modelo)
    If lErro <> SUCESSO Then gError 103579

    'exclui o modelo do Bd
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM IN86Modelos", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 103580

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Efetua o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 103581

    IN86Modelos_Exclui = SUCESSO

    Exit Function

Erro_IN86Modelos_Exclui:

    IN86Modelos_Exclui = gErr

    Select Case gErr

        Case 103574
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103575
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103576, 103577
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_LEITURA", gErr)
            
        Case 103580
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_EXCLUSAO", gErr, objIN86Modelo.sModelo)

        Case 103578
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MODELOS_NAO_ENCONTRADO", gErr)

        Case 103579, 127001
        
        Case 103581
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150298)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Efetua o rollback da transação
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Private Function NF_Exclui_LivRegESEmitentes(ByVal lNumIntDocRegES As Long, ByVal lNumIntEmitente As Long, alComando() As Long) As Long
'Exclui emitente de LivRegESEmitentes se nao houver outro registro em livreges c/aquele emitente.

Dim lErro As Long, lNumIntDocRegESAux As Long, lNumIntDocLivFechado As Long

On Error GoTo Erro_NF_Exclui_LivRegESEmitentes

    'verifica se emitente está associado a outra nf
    lErro = Comando_Executar(alComando(10), "SELECT NumIntDoc FROM LivRegES where NumIntEmitente = ? AND NumIntDoc <> ?", lNumIntDocRegESAux, lNumIntEmitente, lNumIntDocRegES)
    If lErro <> AD_SQL_SUCESSO Then gError 81554
    
    lErro = Comando_BuscarProximo(alComando(10))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81555
    
    'se nao estiver
    If lErro <> AD_SQL_SUCESSO Then
    
        lErro = Comando_ExecutarPos(alComando(11), "SELECT NumIntDocLivFechado FROM LivRegESEmitentes WHERE Codigo = ?", 0, lNumIntDocLivFechado, lNumIntEmitente)
        If lErro <> AD_SQL_SUCESSO Then gError 81556
        
        lErro = Comando_BuscarProximo(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81557
    
        'se o registro nao está associado a livro fechado
        If lNumIntDocLivFechado = 0 Then
        
            lErro = Comando_ExecutarPos(alComando(12), "DELETE FROM LivRegESEmitentes", alComando(11))
            If lErro <> AD_SQL_SUCESSO Then gError 81558
            
        End If
        
    End If
    
    NF_Exclui_LivRegESEmitentes = SUCESSO
     
    Exit Function
    
Erro_NF_Exclui_LivRegESEmitentes:

    NF_Exclui_LivRegESEmitentes = gErr
     
    Select Case gErr
          
        Case 81554, 81555
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case 81556, 81557
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESEMITENTES", gErr)
        
        Case 81558
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESEMITENTES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150299)
     
    End Select
     
    Exit Function

End Function

Private Function NF_Exclui_LivRegESItemNF(ByVal lNumIntDocRegES As Long, alComando() As Long) As Long
'Exclui registros de livregesitemnf associados a nf. Pode excluir tambem registros de LivRegESCadProd

Dim lErro As Long, lNumIntCadProd As Long, lNumIntDocItemNF As Long, dtDataFinal As Date

On Error GoTo Erro_NF_Exclui_LivRegESItemNF

    'seleciona e exclui regs de livregesitemnf
    lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntCadProd FROM LivRegESItemNF WHERE NumIntDocRegES = ?", 0, lNumIntCadProd, lNumIntDocRegES)
    If lErro <> AD_SQL_SUCESSO Then gError 81559
    
    lErro = Comando_BuscarProximo(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81560
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'verifica se reg de LivRegESCadProd está associado a outra nf
        lErro = Comando_Executar(alComando(7), "SELECT NumIntDoc FROM LivRegESItemNF WHERE NumIntCadProd = ? AND NumIntDocRegES <> ?", lNumIntDocItemNF, lNumIntCadProd, lNumIntDocRegES)
        If lErro <> AD_SQL_SUCESSO Then gError 81561
        
        lErro = Comando_BuscarProximo(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81562
    
        'se nao estiver pode-se exclui-lo
        If lErro <> AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(8), "SELECT DataFinal FROM LivRegESCadProd WHERE NumIntDoc = ?", 0, dtDataFinal, lNumIntCadProd)
            If lErro <> AD_SQL_SUCESSO Then gError 81563
            
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81564
    
            If lErro = AD_SQL_SUCESSO Then
            
                'se o cadastro do produto nao foi alterado
                If dtDataFinal <> DATA_NULA Then
                
                    lErro = Comando_ExecutarPos(alComando(9), "DELETE FROM LivRegESCadProd", alComando(8))
                    If lErro <> AD_SQL_SUCESSO Then gError 81565
                
                End If
            
            End If
            
        End If
        
        lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM LivRegESItemNF", alComando(5))
        If lErro <> AD_SQL_SUCESSO Then gError 81566
        
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81567
    
    Loop
    
    NF_Exclui_LivRegESItemNF = SUCESSO
     
    Exit Function
    
Erro_NF_Exclui_LivRegESItemNF:

    NF_Exclui_LivRegESItemNF = gErr
     
    Select Case gErr
          
        Case 81559, 81560, 81561, 81562, 81567
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESITEM", gErr)
        
        Case 81563, 81654
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESCADPROD", gErr)
        
        Case 81565
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESCADPROD", gErr)
        
        Case 81566
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESITEM", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150300)
     
    End Select
     
    Exit Function

End Function

Function NotaFiscal_Exclui_Fis(ByVal objNFiscal As ClassNFiscal) As Long
'Exclui das Tabelas de Livros Fiscais a Nota Fiscal
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, iIndice As Integer, alComando(1 To 12) As Long
Dim lNumIntDocRegES As Long, lNumIntDocRegESLinha As Long, lNumIntDocLivFechado As Long, lNumIntEmitente As Long

On Error GoTo Erro_NotaFiscal_Exclui_Fis

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81568
    Next
    
    If gobjFIS.dtFisBloqDataLimite <> DATA_NULA Then
    
        If objNFiscal.dtDataEntrada <> DATA_NULA Then
            If objNFiscal.dtDataEntrada < gobjFIS.dtFisBloqDataLimite Then gError 213154
        Else
            If objNFiscal.dtDataEmissao < gobjFIS.dtFisBloqDataLimite Then gError 213154
        End If
        
    End If
    
    'pesquisa registro em livreges
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, NumIntDocLivFechado, NumIntEmitente FROM LivRegES WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", 0, lNumIntDocRegES, lNumIntDocLivFechado, lNumIntEmitente, TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 81569
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81570
    
    'Se encontrou registro no Livro Fiscal
    If lErro <> AD_SQL_SEM_DADOS Then
        
        'bloqueia o registro
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 81571
        
        'verifica se livro está fechado
        If lNumIntDocLivFechado <> 0 Then gError 81572
        
        'seleciona e exclui registros de livregeslinha
        lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM LivRegESLinha WHERE NumIntDocRegES = ?", 0, lNumIntDocRegESLinha, lNumIntDocRegES)
        If lErro <> AD_SQL_SUCESSO Then gError 81573
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81574
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM LivRegESLinha", alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 81575
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81576
    
        Loop
        
        'exclui regs de LivRegESItemNF e livregescadprod
        lErro = NF_Exclui_LivRegESItemNF(lNumIntDocRegES, alComando)
        If lErro <> SUCESSO Then gError 81577
        
        If lNumIntEmitente <> 0 Then
        
            'exclui registro de emitente
            lErro = NF_Exclui_LivRegESEmitentes(lNumIntDocRegES, lNumIntEmitente, alComando)
            If lErro <> SUCESSO Then gError 81578
        
        End If
        
        'exclui registro de livreges
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM LivRegES", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 81579
        
    End If
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NotaFiscal_Exclui_Fis = SUCESSO
     
    Exit Function
    
Erro_NotaFiscal_Exclui_Fis:

    NotaFiscal_Exclui_Fis = gErr
     
    Select Case gErr
          
        Case 213154
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_BLOQUEIA_DATA", gErr, "Movimento Fiscal", gobjFIS.dtFisBloqDataLimite)
        
        Case 81577, 81578
        
        Case 81568
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81569, 81570
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case 81571
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
        
        Case 81572
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGES_FECHADO", gErr)
        
        Case 81573, 81574, 81576
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESLINHA", gErr)
        
        Case 81575
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESLINHA", gErr)
        
        Case 81579
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150301)
     
    End Select
     
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NotaFiscal_Grava_Fis(ByVal objNFiscal As ClassNFiscal) As Long
'Grava nas Tabelas de Livros Fiscais a Nota Fiscal

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objEndereco As New ClassEndereco
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_NotaFiscal_Grava_Fis

    If gobjFIS.dtFisBloqDataLimite <> DATA_NULA Then
    
        If objNFiscal.dtDataEntrada <> DATA_NULA Then
            If objNFiscal.dtDataEntrada < gobjFIS.dtFisBloqDataLimite Then gError 213154
        Else
            If objNFiscal.dtDataEmissao < gobjFIS.dtFisBloqDataLimite Then gError 213154
        End If
        
    End If

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o TipoDocInfo passado
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27263 Then gError 69200
    
    'não encontrado no BD --> Erro
    If lErro = 27263 Then gError 69201
    
    'só vou gerar info para notas que interessem p/livros de ICMS/IPI
    If objTipoDocInfo.iModeloArqICMS <> 0 Then
    
        objFilialEmpresa.iCodFilial = objNFiscal.iFilialEmpresa
        
        'Le os dados da Filial Empresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 69234
                
        If lErro <> SUCESSO Then gError 69235
        
        '???? Pode ser melhorado pois o Cliente já foi lido na gravacao da Nota Fiscal
        If objNFiscal.lCliente > 0 Then
            
            objFilialCliente.lCodCliente = objNFiscal.lCliente
            objFilialCliente.iCodFilial = objNFiscal.iFilialCli
                    
            'Le os Dados da Filial Cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 69226
            
            If lErro = 12567 Then gError 69227
                    
            objCliente.lCodigo = objNFiscal.lCliente
            
            'Le os Dados do Cliente
            lErro = CF("Cliente_ObterDadosBasicos", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 69228
            
            If lErro = 12293 Then gError 69229
            
            'obtenho o codigo do endereco principal
            objEndereco.lCodigo = objFilialCliente.lEndereco
                
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 69497
                
            If lErro = 12309 Then gError 69498
                    
        ElseIf objNFiscal.lFornecedor > 0 Then
        
            objFilialFornecedor.lCodFornecedor = objNFiscal.lFornecedor
            objFilialFornecedor.iCodFilial = objNFiscal.iFilialForn
            
            'Le os Dados da Filial do Fornecedor
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 69263
            
            If lErro = 12929 Then gError 69264
            
            objFornecedor.lCodigo = objNFiscal.lFornecedor
            
            'Le os dados do Fornecedor (???? pode melhorar)
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 69265
            
            If lErro = 12729 Then gError 69266
            
            objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
            'Le o Endereço do Fornecedor
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 69268
            
            If lErro = 12309 Then gError 69269
                
        End If
        
        'Verifica se é uma nota Fiscal de Entrada ou é uma Nota Fiscal de Saída
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
            
            'Se for uma Nota Fiscal de Saída chama
            lErro = NotaFiscal_Saida_Grava_Fis(objNFiscal, objTipoDocInfo, objCliente, objFilialCliente, objFornecedor, objFilialFornecedor, objEndereco)
            If lErro <> SUCESSO Then gError 69202
        
        ElseIf objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
        
            'Se for uma Nota Fiscal de Entrada chama
            lErro = NotaFiscal_Entrada_Grava_Fis(objNFiscal, objTipoDocInfo, objCliente, objFilialCliente, objFornecedor, objFilialFornecedor, objEndereco, objFilialEmpresa)
            If lErro <> SUCESSO Then gError 69203
        
        End If

    End If
    
    NotaFiscal_Grava_Fis = SUCESSO
    
    Exit Function
    
Erro_NotaFiscal_Grava_Fis:

    NotaFiscal_Grava_Fis = gErr
    
    Select Case gErr
        
        Case 213154
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_BLOQUEIA_DATA", gErr, "Movimento Fiscal", gobjFIS.dtFisBloqDataLimite)
        
        Case 69200, 69202, 69203
        
        Case 69201
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case 69226, 69228, 69497, 69263, 69265, 69268, 69234
        
        Case 69227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
        
        Case 69229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case 69235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case 69498, 69269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case 69264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
        
        Case 69266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150302)

    End Select
    
    Exit Function
        
End Function

Function NotaFiscal_Saida_Grava_Fis(ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objFornecedor As ClassFornecedor, ByVal objFilialFornecedor As ClassFilialFornecedor, ByVal objEndereco As ClassEndereco) As Long
'Grava a  Nota Fiscal de Saída no Fis

Dim objLivRegES As New ClassLivRegES
Dim objLivRegESEmitentes As New ClassLivRegESEmitentes
Dim colLivRegESItemNF As New Collection
Dim colLivRegESLinha As New Collection
Dim lErro As Long

On Error GoTo Erro_NotaFiscal_Saida_Grava_Fis
    
    'Gera o Registro de LivRegESEmitentes
'    lErro = NF_Grava_LivRegESEmitentes(objNFiscal, objLivRegES, objTipoDocInfo, objLivRegESEmitentes, objCliente, objFilialCliente, objFornecedor, objFilialFornecedor, objEndereco)
'    If lErro <> SUCESSO Then gError 69204
    
    'Gera o Registro de LivRegES
    lErro = NF_Saida_Grava_LivRegES(objNFiscal, objLivRegES, objTipoDocInfo, objCliente, objFilialCliente, objFornecedor, objFilialFornecedor, objEndereco)
    If lErro <> SUCESSO Then gError 69205
    
    'Gera as linhas para livros fiscais
    lErro = NotaFiscal_GeraLinhas_Fis(objTipoDocInfo.iModeloArqICMS, colLivRegESLinha, colLivRegESItemNF, objNFiscal, objLivRegES)
    If lErro <> SUCESSO Then gError 90560
    
    'Grava as Coleções, que estão ordenadas e completadas com o Complemento
    lErro = NF_Grava_LivRegES_Colecao(colLivRegESLinha, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69210
    
    NotaFiscal_Saida_Grava_Fis = SUCESSO
    
    Exit Function
    
Erro_NotaFiscal_Saida_Grava_Fis:

    NotaFiscal_Saida_Grava_Fis = gErr
    
    Select Case gErr
        
        Case 69204, 69205, 69210, 81628, 90560
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150303)

    End Select
    
    Exit Function
    
End Function

Function NF_Saida_Move_Tela_LivRegESLinha(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection) As Long
'Preenhe a coleção

Dim lErro As Long

On Error GoTo Erro_NF_Saida_Move_Tela_LivRegESLinha

    lErro = NF_Item_Move_Tela_LivRegESLinha(objNFiscal, objItemNF, objLivRegES, colLivRegESLinha)
    If lErro <> SUCESSO Then gError 81663
    
    NF_Saida_Move_Tela_LivRegESLinha = SUCESSO
        
    Exit Function
    
Erro_NF_Saida_Move_Tela_LivRegESLinha:

    NF_Saida_Move_Tela_LivRegESLinha = gErr
    
    Select Case gErr
        
        Case 81663
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150304)

    End Select
    
    Exit Function

End Function

Function NF_Saida_Grava_LivRegES(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objFornecedor As ClassFornecedor, ByVal objFilialFornecedor As ClassFilialFornecedor, ByVal objEndereco As ClassEndereco) As Long
'Preenche o objLivRegES com os dados da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objFilialEmpresa As New AdmFiliais
Dim lComando As Long
Dim sInscricaoAux As String

On Error GoTo Erro_NF_Saida_Grava_LivRegES
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69224
    
    objLivRegES.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objLivRegES.dtData = objNFiscal.dtDataEmissao
    objLivRegES.dtDataEmissao = objNFiscal.dtDataEmissao
    objLivRegES.iDestinatario = objTipoDocInfo.iDestinatario
    objLivRegES.iEmitente = objTipoDocInfo.iEmitente
    
    If right(UCase(objNFiscal.sSerie), 2) = "-E" Then
        objLivRegES.sSerie = left(objNFiscal.sSerie, Len(objNFiscal.sSerie) - 2)
    Else
        objLivRegES.sSerie = objNFiscal.sSerie
    End If
    
    'nfe federal
    If objNFiscal.iNFe <> 0 And objTipoDocInfo.iModeloArqICMS = 1 Then
        objLivRegES.iModelo = 55
    ElseIf objNFiscal.iNFe = MARCADO And objTipoDocInfo.iModeloArqICMS = 8 Then
        objLivRegES.iModelo = 57 'CONHECIMENTO DE TRANSPORTE ELETRÔNICO – CT-e (CÓDIGO 57)
    Else
        objLivRegES.iModelo = objTipoDocInfo.iModeloArqICMS
    End If
    
    objLivRegES.iOrigem = objTipoDocInfo.iOrigem
    objLivRegES.iSituacao = SITUACAO_NORMAL
    objLivRegES.iStatusLivro = STATUS_FIS_ORIGINAL
    objLivRegES.iTipo = TIPO_REGES_SAIDA
    objLivRegES.iFreteRespons = objNFiscal.iFreteRespons
    'inclui este teste em funcao do conhecimentos de transporte
    If Not (objNFiscal.objTributacaoNF Is Nothing) Then
        objLivRegES.dPISValor = objNFiscal.objTributacaoNF.dPISCredito
        objLivRegES.dCOFINSValor = objNFiscal.objTributacaoNF.dCOFINSCredito
    Else
        objLivRegES.dPISValor = 0
        objLivRegES.dCOFINSValor = 0
    End If
    
    'Gera número automático para o Item de apuração ICMS
    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69225
    
    objLivRegES.lNumIntDocOrigem = objNFiscal.lNumIntDoc
    objLivRegES.lNumIntDoc = lNumIntDoc
    objLivRegES.iTipoNumIntDocOrigem = TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL
    objLivRegES.lNumIntNF = objNFiscal.lNumIntDoc
    objLivRegES.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
    
    If objLivRegES.iDestinatario = DESTINATARIO_CLIENTE Then
        
        'Arrumar uma solução viavel
        objLivRegES.sCgc = objFilialCliente.sCgc
        Call Formata_String_Numero(objFilialCliente.sInscricaoEstadual, sInscricaoAux)
        objLivRegES.sInscricaoEstadual = sInscricaoAux
        objLivRegES.iIEIsento = objFilialCliente.iIEIsento
        objLivRegES.sNome = objCliente.sRazaoSocial
        'objLivRegES.sUF = objEndereco.sSiglaEstado
        
        If objEndereco.iCodigoPais <> PAIS_BRASIL Then
            objLivRegES.sUF = "EX"
        Else
            objLivRegES.sUF = objEndereco.sSiglaEstado
        End If
    
    ElseIf objLivRegES.iDestinatario = DESTINATARIO_FORNECEDOR Then
        
        'Arrumar uma solução viavel
        objLivRegES.sCgc = objFilialFornecedor.sCgc
        Call Formata_String_Numero(objFilialFornecedor.sInscricaoEstadual, sInscricaoAux)
        objLivRegES.sInscricaoEstadual = sInscricaoAux
        objLivRegES.iIEIsento = objFilialFornecedor.iIEIsento
        objLivRegES.sNome = objFornecedor.sRazaoSocial
        'objLivRegES.sUF = objEndereco.sSiglaEstado
        
        If objEndereco.iCodigoPais <> PAIS_BRASIL Then
            objLivRegES.sUF = "EX"
        Else
            objLivRegES.sUF = objEndereco.sSiglaEstado
        End If
        
    ElseIf objLivRegES.iDestinatario = DESTINATARIO_EMPRESA Then
        
        objFilialEmpresa.iCodFilial = objNFiscal.iFilialEmpresa
        
        'Le os dados da Filial Empresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 69234
                
        If lErro <> SUCESSO Then gError 69235
        
        'Arrumar uma solução viavel
        objLivRegES.sCgc = objFilialEmpresa.sCgc
        Call Formata_String_Numero(objFilialEmpresa.sInscricaoEstadual, sInscricaoAux)
        objLivRegES.sInscricaoEstadual = sInscricaoAux
        objLivRegES.iIEIsento = IIf(Len(Trim(sInscricaoAux)) = 0, 1, 0)
        objLivRegES.sNome = gsNomeEmpresa
        objLivRegES.sUF = objFilialEmpresa.objEndereco.sSiglaEstado
        
    End If
    
    'Faz a inserção na Tabela de LivResES
    lErro = Comando_Executar(lComando, "INSERT INTO LivRegES (NumIntDoc, TipoNumIntDocOrigem, NumIntDocOrigem, StatusLivro, Tipo, CGC, InscricaoEstadual, Nome, Data, DataEmissao, UF, Serie, NumNotaFiscal, Situacao, Modelo, Emitente, Destinatario, Origem, NumIntEmitente, FilialEmpresa, CIF_FOB, PISValor, COFINSValor, IEIsento) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objLivRegES.lNumIntDoc, objLivRegES.iTipoNumIntDocOrigem, objLivRegES.lNumIntDocOrigem, objLivRegES.iStatusLivro, objLivRegES.iTipo, objLivRegES.sCgc, objLivRegES.sInscricaoEstadual, objLivRegES.sNome, objLivRegES.dtData, objLivRegES.dtDataEmissao, objLivRegES.sUF, objLivRegES.sSerie, objLivRegES.lNumNotaFiscal, objLivRegES.iSituacao, objLivRegES.iModelo, objLivRegES.iEmitente, objLivRegES.iDestinatario, objLivRegES.iOrigem, objLivRegES.lNumIntEmitente, objLivRegES.iFilialEmpresa, objLivRegES.iFreteRespons, objLivRegES.dPISValor, objLivRegES.dCOFINSValor, objLivRegES.iIEIsento)
    If lErro <> AD_SQL_SUCESSO Then gError 69236
                
    Call Comando_Fechar(lComando)
                
    NF_Saida_Grava_LivRegES = SUCESSO
    
    Exit Function
    
Erro_NF_Saida_Grava_LivRegES:

    NF_Saida_Grava_LivRegES = gErr

    Select Case gErr
    
        Case 69224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69225, 69234
        
        Case 69236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150305)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function NF_Grava_LivRegESCadProd(ByVal dtData As Date, ByVal objItemNF As ClassItemNF, objLivRegESCadProd As ClassLivRegESCadProd) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objProduto As New ClassProduto
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim lComando As Long, dtDataInicialAnterior As Date
Dim sCodigoIpiAux As String, lComando1 As Long, lComando2 As Long

On Error GoTo Erro_NF_Grava_LivRegESCadProd

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69237

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 81648

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 81649

    'Lê no BD para ver se este Produto já não está cadastrado na tabela LivResESCadProd
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, DataInicial FROM LivRegESCadProd WHERE Produto = ? ORDER BY DataInicial", 0, lNumIntDoc, dtDataInicialAnterior, objItemNF.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 69238

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69239

    'Se não estiver então preenche o objeto
    If lErro = AD_SQL_SEM_DADOS Then
    
        Set objLivRegESCadProd = New ClassLivRegESCadProd

        objLivRegESCadProd.dAliquotaICMS = 0 '????? definir Jones
        objLivRegESCadProd.dAliquotaIPI = 0  '????? definir Jones
        objLivRegESCadProd.dBaseCalculoICMSSubst = 0 '????? definir Jones
        objLivRegESCadProd.dReducaoBaseCalculoICMS = 0 '????? definir Jones
        objLivRegESCadProd.dtDataFinal = DATA_NULA
        objLivRegESCadProd.dtDataInicial = dtData
        
        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_CADPROD", lNumIntDoc)
        If lErro <> SUCESSO Then gError 69240
        
        objLivRegESCadProd.lNumIntDoc = lNumIntDoc
        
        objLivRegESCadProd.sDescricao = left(objItemNF.sDescricaoItem, 50)
        objLivRegESCadProd.sProduto = objItemNF.sProduto
        objLivRegESCadProd.sSiglaUM = objItemNF.sUnidadeMed
        
        objProduto.sCodigo = objItemNF.sProduto
                
        'Lê o Produto para ler a conta Aplicação
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 69241
                
        If lErro = 28030 Then gError 69242
        
        objTipoTribICMS.iTipo = objItemNF.objTributacaoItemNF.iICMSTipo
            
        objLivRegESCadProd.sCodigoNCM = Mid(objProduto.sIPICodigo, 1, 8)
        
        lErro = CF("TipoTribICMS_Le", objTipoTribICMS)
        If lErro <> SUCESSO And lErro <> 21534 Then gError 69243
        
        If lErro = 21534 Then gError 69244
        
        objLivRegESCadProd.sSituacaoTrib = CStr(objProduto.iOrigemMercadoria) & left(CStr(objTipoTribICMS.iTipoTribCST), 1)
        
        'Faz a inserção na Tabela de LivResESCadProd
        lErro = Comando_Executar(lComando1, "INSERT INTO LivRegESCadProd ( NumIntDoc, Produto, CodigoNCM, Descricao, SiglaUM, SituacaoTrib, AliquotaIPI, AliquotaICMS, ReducaoBaseCalculoICMS, BaseCalculoICMSSubst, DataInicial) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objLivRegESCadProd.lNumIntDoc, objLivRegESCadProd.sProduto, objLivRegESCadProd.sCodigoNCM, objLivRegESCadProd.sDescricao, objLivRegESCadProd.sSiglaUM, objLivRegESCadProd.sSituacaoTrib, objLivRegESCadProd.dAliquotaIPI, objLivRegESCadProd.dAliquotaICMS, objLivRegESCadProd.dReducaoBaseCalculoICMS, objLivRegESCadProd.dBaseCalculoICMSSubst, objLivRegESCadProd.dtDataInicial)
        If lErro <> AD_SQL_SUCESSO Then gError 69245
        
    Else
    
        If dtDataInicialAnterior > dtData Then
        
            'Atualiza a data inicial
            lErro = Comando_ExecutarPos(lComando2, "UPDATE LivRegESCadProd SET DataInicial = ?", lComando, dtData)
            If lErro <> AD_SQL_SUCESSO Then gError 81647
        
        End If
    
    End If
    
    objLivRegESCadProd.lNumIntDoc = lNumIntDoc
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    NF_Grava_LivRegESCadProd = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegESCadProd:

    NF_Grava_LivRegESCadProd = gErr

    Select Case gErr
    
        Case 69237, 81648, 81649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69238, 69239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESCADPROD", gErr)
        
        Case 69240, 69241, 69243
        
        Case 69242
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 69244
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_ICMS_INEXISTENTE", gErr, objTipoTribICMS.iTipo)
        
        Case 69245
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESCADPROD", gErr)
        
        Case 81647
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGESCADPROD", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150306)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function NF_Move_Memoria_LivRegESItemNF(ByVal objItemNF As ClassItemNF, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objLivRegESCadProd As ClassLivRegESCadProd, ByVal colLivRegESItemNF As Collection) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_NF_Move_Memoria_LivRegESItemNF

    Set objLivRegESItemNF = New ClassLivRegESItemNF
    
    With objItemNF.objTributacaoItemNF
        objLivRegESItemNF.dAliquotaICMS = .dICMSAliquota
        objLivRegESItemNF.dQuantidade = objItemNF.dQuantidade
        objLivRegESItemNF.dValorICMSBase = .dICMSBase
        objLivRegESItemNF.dValorICMSSubstBase = .dICMSSubstBase
        objLivRegESItemNF.dValorIPI = .dIPIValor
        objLivRegESItemNF.dValorProduto = Round(objItemNF.dPrecoUnitario * objItemNF.dQuantidade, 2)
        objLivRegESItemNF.iNumItem = objItemNF.iItem
        objLivRegESItemNF.dValorDescontoDespAcess = Round(objItemNF.dValorDesconto + objItemNF.dValorDescGlobal, 2)
        objLivRegESItemNF.iTipoTribICMS = .iICMSTipo
        objLivRegESItemNF.iTipoTribIPI = .iIPITipo
        objLivRegESItemNF.dAliquotaIPI = .dIPIAliquota
        objLivRegESItemNF.dRedBaseICMS = .dICMSPercRedBase
        objLivRegESItemNF.dRedBaseIPI = .dIPIPercRedBase
        objLivRegESItemNF.iTipoTributacao = .iTipoTributacao
        objLivRegESItemNF.dAliquotaSubst = .dICMSSubstAliquota
        objLivRegESItemNF.dValorSubst = .dICMSSubstValor
        objLivRegESItemNF.dValorICMS = .dICMSValor
        objLivRegESItemNF.dValorIPIBase = .dIPIBaseCalculo
    End With
    
    'Gera número automático para o Item de apuração ICMS
    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69246
    
    objLivRegESItemNF.lNumIntDoc = lNumIntDoc
    objLivRegESItemNF.sCFOP = objItemNF.objTributacaoItemNF.sNaturezaOp
    
    objLivRegESItemNF.lNumIntCadProd = objLivRegESCadProd.lNumIntDoc
    objLivRegESItemNF.lNumIntDocRegES = objLivRegES.lNumIntDoc
        
    colLivRegESItemNF.Add objLivRegESItemNF
    
    NF_Move_Memoria_LivRegESItemNF = SUCESSO
    
    Exit Function
    
Erro_NF_Move_Memoria_LivRegESItemNF:

    NF_Move_Memoria_LivRegESItemNF = gErr

    Select Case gErr
        
        Case 69246
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150307)

    End Select
    
    Exit Function

End Function

Function NF_Grava_LivRegESEmitentes(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objLivRegESEmitentes As ClassLivRegESEmitentes, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objFornecedor As ClassFornecedor, ByVal objFilialFornecedor As ClassFilialFornecedor, ByVal objEndereco As ClassEndereco) As Long
'Preenhe o objLivRegESEmitentes com os dados da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim sTabelaConfig As String
Dim sCodigo As String
Dim sTabelaObjeto As String
Dim sCampo As String
Dim lCodigo As Long
Dim lComando As Long
Dim lComando2 As Long
Dim sInscricaoAux As String

On Error GoTo Erro_NF_Grava_LivRegESEmitentes
        
    'Move o Emitente para o objeto
    If objNFiscal.lCliente <> 0 Then
                    
        'Abre o Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 69247
    
        'Abre o Comando
        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then gError 69248
    
        'Le no BD para ver se este emitente já não está cadastrado
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM LivRegESEmitentes WHERE CliForn = ? AND Filial = ? AND TipoEmitente = ? AND NumIntDocLivFechado = 0", lCodigo, objNFiscal.lCliente, objNFiscal.iFilialCli, EMITENTE_CLIENTE)
        If lErro <> AD_SQL_SUCESSO Then gError 69249

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
        'Se não estiver então preenche o objeto
        If lErro = AD_SQL_SEM_DADOS Then
                    
            objLivRegESEmitentes.iFilial = objNFiscal.iFilialCli
            objLivRegESEmitentes.iTipoEmitente = EMITENTE_CLIENTE
            objLivRegESEmitentes.lCliForn = objNFiscal.lCliente
            
            sCodigo = "FISC_NUM_PROX_REG_EMITENTES"
            sTabelaConfig = "FISConfig"
            sTabelaObjeto = "LivRegESEmitentes"
            sCampo = "Codigo"
    
            'Chama Config_ObterAutomatico
            lErro = CF("Config_ObterAutomatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
            If lErro <> SUCESSO Then gError 69255
            
            objLivRegESEmitentes.lCodigo = lCodigo
            objLivRegESEmitentes.sCgc = objFilialCliente.sCgc
            Call Formata_String_Numero(objFilialCliente.sInscricaoEstadual, sInscricaoAux)
            objLivRegESEmitentes.sInscricaoEstadual = sInscricaoAux
            objLivRegESEmitentes.sNome = objCliente.sRazaoSocial
            'objLivRegESEmitentes.sUF = objEndereco.sSiglaEstado
            
            If objEndereco.iCodigoPais <> PAIS_BRASIL Then
                objLivRegESEmitentes.sUF = "EX"
            Else
                objLivRegESEmitentes.sUF = objEndereco.sSiglaEstado
            End If
               
            'colocar aqui o código para gravação do emitente
            lErro = Comando_Executar(lComando2, "INSERT INTO LivRegESEmitentes (Codigo, Nome, CGC, InscricaoEstadual, UF, TipoEmitente, CliForn, Filial) VALUES (?,?,?,?,?,?,?,?)", objLivRegESEmitentes.lCodigo, objLivRegESEmitentes.sNome, objLivRegESEmitentes.sCgc, objLivRegESEmitentes.sInscricaoEstadual, objLivRegESEmitentes.sUF, objLivRegESEmitentes.iTipoEmitente, objLivRegESEmitentes.lCliForn, objLivRegESEmitentes.iFilial)
            If lErro <> AD_SQL_SUCESSO Then gError 69258
               
        End If
        
    ElseIf objNFiscal.lFornecedor <> 0 Then
        
        'Abre o Comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 69259
    
        'Abre o Comando
        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then gError 69260
    
        'Le no BD para ver se este emitente já não está cadastrado
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM LivRegESEmitentes WHERE CliForn = ? AND Filial = ? AND TipoEmitente = ?", lCodigo, objNFiscal.lFornecedor, objNFiscal.iFilialForn, EMITENTE_FORNECEDOR)
        If lErro <> AD_SQL_SUCESSO Then gError 69261

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69262
    
        'Se não estiver então preenche o objeto
        If lErro = AD_SQL_SEM_DADOS Then
        
            objLivRegESEmitentes.iFilial = objNFiscal.iFilialForn
            objLivRegESEmitentes.iTipoEmitente = EMITENTE_FORNECEDOR
            objLivRegESEmitentes.lCliForn = objNFiscal.lFornecedor
            
            sCodigo = "FISC_NUM_PROX_REG_EMITENTES"
            sTabelaConfig = "FISConfig"
            sTabelaObjeto = "LivRegESEmitentes"
            sCampo = "Codigo"
    
            'Chama Config_ObterAutomatico
            lErro = CF("Config_ObterAutomatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
            If lErro <> SUCESSO Then gError 69267
            
            objLivRegESEmitentes.lCodigo = lCodigo
            objLivRegESEmitentes.sCgc = objFilialFornecedor.sCgc
            Call Formata_String_Numero(objFilialFornecedor.sInscricaoEstadual, sInscricaoAux)
            objLivRegESEmitentes.sInscricaoEstadual = sInscricaoAux
            objLivRegESEmitentes.sNome = objFornecedor.sRazaoSocial
            'objLivRegESEmitentes.sUF = objEndereco.sSiglaEstado
            
            If objEndereco.iCodigoPais <> PAIS_BRASIL Then
                objLivRegESEmitentes.sUF = "EX"
            Else
                objLivRegESEmitentes.sUF = objEndereco.sSiglaEstado
            End If
            
            'colocar aqui o código para gravação do emitente
            lErro = Comando_Executar(lComando2, "INSERT INTO LivRegESEmitentes(Codigo, Nome, CGC, InscricaoEstadual, UF, TipoEmitente, CliForn, Filial) VALUES (?,?,?,?,?,?,?,?)", objLivRegESEmitentes.lCodigo, objLivRegESEmitentes.sNome, objLivRegESEmitentes.sCgc, objLivRegESEmitentes.sInscricaoEstadual, objLivRegESEmitentes.sUF, objLivRegESEmitentes.iTipoEmitente, objLivRegESEmitentes.lCliForn, objLivRegESEmitentes.iFilial)
            If lErro <> AD_SQL_SUCESSO Then gError 69270
            
        End If
        
    End If
        
    objLivRegES.lNumIntEmitente = lCodigo
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    NF_Grava_LivRegESEmitentes = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegESEmitentes:

    NF_Grava_LivRegESEmitentes = gErr

    Select Case gErr
        
        Case 69247, 69248, 69259, 69260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69249, 69250, 69261, 69262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESEMITENTES", gErr)
        
        Case 69251, 69253, 69255, 69256, 69263, 69265, 69267, 69268
        
        Case 69252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
        
        Case 69254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case 69257, 69269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)
        
        Case 69258, 69270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESEMITENTES", gErr)
        
        Case 69264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
        
        Case 69266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150308)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function NF_Saida_Move_Tela_Complemento(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long
'Move para as coleções os Itens para o Frete, Seguro, Despesas

Dim lErro As Long

On Error GoTo Erro_NF_Saida_Move_Tela_Complemento

    'Move os dados do Frete
    If objNFiscal.dValorFrete > 0 Then
        
        lErro = NF_Move_Memoria_LivRegESItemNF_Frete(objNFiscal, objLivRegES, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 69272
                
    End If
    
    'Move os dados do Seguro
    If objNFiscal.dValorSeguro > 0 Then
                
        lErro = NF_Move_Memoria_LivRegESItemNF_Seguro(objNFiscal, objLivRegES, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 69274
                
    End If
    
    'Move os dados do Outras
    If objNFiscal.dValorOutrasDespesas > 0 Then
        
        lErro = NF_Move_Memoria_LivRegESItemNF_Outras(objNFiscal, objLivRegES, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 69276
                
    End If
        
    NF_Saida_Move_Tela_Complemento = SUCESSO
    
    Exit Function
    
Erro_NF_Saida_Move_Tela_Complemento:

    NF_Saida_Move_Tela_Complemento = gErr

    Select Case gErr
        
        Case 69271, 69272, 69273, 69274, 69275, 69276
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150309)

    End Select
    
    Exit Function

End Function

Function NF_Move_Memoria_LivRegESItemNF_Frete(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESItemNF As Collection) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_NF_Move_Memoria_LivRegESItemNF_Frete

    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69304
    
    objLivRegESItemNF.lNumIntDoc = lNumIntDoc
    objLivRegESItemNF.lNumIntDocRegES = objLivRegES.lNumIntDoc
    
    objLivRegESItemNF.iNumItem = 991
    objLivRegESItemNF.sCFOP = objNFiscal.objTributacao.sNaturezaOpInterna
    objLivRegESItemNF.dValorDescontoDespAcess = objNFiscal.dValorFrete
    
    colLivRegESItemNF.Add objLivRegESItemNF
    
    NF_Move_Memoria_LivRegESItemNF_Frete = SUCESSO
    
    Exit Function
    
Erro_NF_Move_Memoria_LivRegESItemNF_Frete:

    NF_Move_Memoria_LivRegESItemNF_Frete = gErr

    Select Case gErr
        
        Case 69304
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150313)

    End Select
    
    Exit Function

End Function

Function NF_Move_Memoria_LivRegESItemNF_Seguro(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESItemNF As Collection) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_NF_Move_Memoria_LivRegESItemNF_Seguro

    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69305
    
    objLivRegESItemNF.lNumIntDoc = lNumIntDoc
    objLivRegESItemNF.lNumIntDocRegES = objLivRegES.lNumIntDoc
    
    objLivRegESItemNF.iNumItem = 992
    objLivRegESItemNF.sCFOP = objNFiscal.objTributacao.sNaturezaOpInterna
    objLivRegESItemNF.dValorDescontoDespAcess = objNFiscal.dValorSeguro
    
    colLivRegESItemNF.Add objLivRegESItemNF
    
    NF_Move_Memoria_LivRegESItemNF_Seguro = SUCESSO
    
    Exit Function
    
Erro_NF_Move_Memoria_LivRegESItemNF_Seguro:

    NF_Move_Memoria_LivRegESItemNF_Seguro = gErr

    Select Case gErr
        
        Case 69305
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150314)

    End Select
    
    Exit Function

End Function

Function NF_Move_Memoria_LivRegESItemNF_Outras(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESItemNF As Collection) As Long
'Preenhe o objLivRegES com os dados do item da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_NF_Move_Memoria_LivRegESItemNF_Outras

    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69306
    
    objLivRegESItemNF.lNumIntDoc = lNumIntDoc
    objLivRegESItemNF.lNumIntDocRegES = objLivRegES.lNumIntDoc
    
    objLivRegESItemNF.iNumItem = 999
    objLivRegESItemNF.sCFOP = objNFiscal.objTributacao.sNaturezaOpInterna
    objLivRegESItemNF.dValorDescontoDespAcess = objNFiscal.dValorOutrasDespesas
    
    colLivRegESItemNF.Add objLivRegESItemNF
    
    NF_Move_Memoria_LivRegESItemNF_Outras = SUCESSO
    
    Exit Function
    
Erro_NF_Move_Memoria_LivRegESItemNF_Outras:

    NF_Move_Memoria_LivRegESItemNF_Outras = gErr

    Select Case gErr
        
        Case 69306
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150315)

    End Select
    
    Exit Function

End Function

Function NF_Grava_LivRegES_Colecao(ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long

Dim lErro As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim lComando As Long
Dim lComando2 As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_NF_Grava_LivRegES_Colecao
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69307
    
    'Abre o Comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 69308
    
    For Each objLivRegESLinha In colLivRegESLinha
        
        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGESLINHA", lNumIntDoc)
        If lErro <> SUCESSO Then gError 69223
    
        objLivRegESLinha.lNumIntDoc = lNumIntDoc
        
        'insere as linhas do Livro Fiscal desta Nota Fiscal
        lErro = Comando_Executar(lComando2, "INSERT INTO LivRegESLinha (NumIntDoc, NumIntDocRegES, NaturezaOp, ValorTotal, ValorICMSBase, ValorICMS, ValorICMSIsentoNaoTrib, ValorICMSOutras, ValorICMSSubstBase, ValorICMSSubstRet, AliquotaICMS, ValorDespAcess, ValorIPI, AliquotaIPI, ValorIPIBase, ValorIPIIsentoNaoTrib, ValorIPIOutras, ClassifContabil, ValorContabil, ObservacaoLivFisc, CodigoICMS, CodigoIPI) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objLivRegESLinha.lNumIntDoc, objLivRegESLinha.lNumIntDocRegES, objLivRegESLinha.sNaturezaOp, Arredonda_Moeda(objLivRegESLinha.dValorTotal), Arredonda_Moeda(objLivRegESLinha.dValorICMSBase), Arredonda_Moeda(objLivRegESLinha.dValorICMS), Arredonda_Moeda(objLivRegESLinha.dValorICMSIsentoNaoTrib), Arredonda_Moeda(objLivRegESLinha.dValorICMSOutras), Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstBase), Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstRet), objLivRegESLinha.dAliquotaICMS, Arredonda_Moeda(objLivRegESLinha.dValorDespAcess), Arredonda_Moeda(objLivRegESLinha.dValorIPI), objLivRegESLinha.dAliquotaIPI, Arredonda_Moeda(objLivRegESLinha.dValorIPIBase), _
            Arredonda_Moeda(objLivRegESLinha.dValorIPIIsentoNaoTrib), Arredonda_Moeda(objLivRegESLinha.dValorIPIOutras), objLivRegESLinha.sClassifContabil, Arredonda_Moeda(objLivRegESLinha.dValorContabil), objLivRegESLinha.sObservacaoLivFisc, objLivRegESLinha.iCodigoICMS, objLivRegESLinha.iCodigoIPI)
        If lErro <> AD_SQL_SUCESSO Then gError 69309
    
    Next
    
    For Each objLivRegESItemNF In colLivRegESItemNF
        
        'insere as linhas do Livro Fiscal desta Nota Fiscal
        lErro = Comando_Executar(lComando2, "INSERT INTO LivRegESItemNF (NumIntDoc, NumIntDocRegES, NumItem, CFOP, NumIntCadProd, Quantidade, ValorProduto, ValorDescontoDespAcess, ValorICMSBase, ValorICMSSubstBase, ValorIPI, AliquotaICMS, TipoTribICMS, TipoTribIPI, ValorICMS, ValorIPIBase, AliquotaIPI, RedBaseICMS, RedBaseIPI, TipoTributacao, AliquotaSubst, ValorSubst) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objLivRegESItemNF.lNumIntDoc, objLivRegESItemNF.lNumIntDocRegES, objLivRegESItemNF.iNumItem, objLivRegESItemNF.sCFOP, objLivRegESItemNF.lNumIntCadProd, objLivRegESItemNF.dQuantidade, objLivRegESItemNF.dValorProduto, objLivRegESItemNF.dValorDescontoDespAcess, objLivRegESItemNF.dValorICMSBase, objLivRegESItemNF.dValorICMSSubstBase, objLivRegESItemNF.dValorIPI, objLivRegESItemNF.dAliquotaICMS, objLivRegESItemNF.iTipoTribICMS, objLivRegESItemNF.iTipoTribIPI, objLivRegESItemNF.dValorICMS, objLivRegESItemNF.dValorIPIBase, objLivRegESItemNF.dAliquotaIPI, objLivRegESItemNF.dRedBaseICMS, objLivRegESItemNF.dRedBaseIPI, objLivRegESItemNF.iTipoTributacao, objLivRegESItemNF.dAliquotaSubst, objLivRegESItemNF.dValorSubst)
        If lErro <> AD_SQL_SUCESSO Then gError 69310
    
    Next
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    NF_Grava_LivRegES_Colecao = SUCESSO
    
    Exit Function
    
Erro_NF_Grava_LivRegES_Colecao:

    NF_Grava_LivRegES_Colecao = gErr
    
    Select Case gErr
        
        Case 69223
        
        Case 69307, 69308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESLINHA", gErr)
        
        Case 69310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESITEMNF", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150316)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

'-----------------------  Entrada  ----------------------------------

Function NotaFiscal_Entrada_Grava_Fis(ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objFornecedor As ClassFornecedor, ByVal objFilialFornecedor As ClassFilialFornecedor, ByVal objEndereco As ClassEndereco, ByVal objFilialEmpresa As AdmFiliais) As Long
'Grava a  Nota Fiscal de Saída no Fis

Dim objItemNF As New ClassItemNF
Dim objLivRegES As New ClassLivRegES
Dim objLivRegESEmitentes As New ClassLivRegESEmitentes
Dim colLivRegESItemNF As New Collection
Dim objLivRegESCadProd As New ClassLivRegESCadProd
Dim colLivRegESLinha As New Collection, objNFiscalOrig As New ClassNFiscal
Dim lErro As Long, objLivRegESLinha As ClassLivRegESLinha

On Error GoTo Erro_NotaFiscal_Entrada_Grava_Fis
    
    'Gera o Registro de LivRegESEmitentes
    lErro = NF_Grava_LivRegESEmitentes(objNFiscal, objLivRegES, objTipoDocInfo, objLivRegESEmitentes, objCliente, objFilialCliente, objFornecedor, objFilialFornecedor, objEndereco)
    If lErro <> SUCESSO Then gError 69204
    
    'Gera o Registro de LivRegES
    lErro = NF_Entrada_Grava_LivRegES(objNFiscal, objLivRegES, objTipoDocInfo, objCliente, objFilialCliente, objFornecedor, objFilialFornecedor, objEndereco, objFilialEmpresa)
    If lErro <> SUCESSO Then gError 69205
    
    'Gera as linhas para livros fiscais
    lErro = NotaFiscal_GeraLinhas_Fis(objTipoDocInfo.iModeloArqICMS, colLivRegESLinha, colLivRegESItemNF, objNFiscal, objLivRegES)
    If lErro <> SUCESSO Then gError 90561
    
    'Grava as Coleções, que estão ordenadas e completadas com o Complemento
    lErro = NF_Grava_LivRegES_Colecao(colLivRegESLinha, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69210
    
    NotaFiscal_Entrada_Grava_Fis = SUCESSO
            
    Exit Function
    
Erro_NotaFiscal_Entrada_Grava_Fis:

    NotaFiscal_Entrada_Grava_Fis = gErr
    
    Select Case gErr
        
        Case 69204, 69205, 69210, 81628, 90561
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150317)

    End Select
    
    Exit Function
    
End Function

Function NF_Entrada_Move_Tela_LivRegESLinha(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection) As Long
'Preenhe a coleção

Dim lErro As Long

On Error GoTo Erro_NF_Entrada_Move_Tela_LivRegESLinha

    lErro = NF_Item_Move_Tela_LivRegESLinha(objNFiscal, objItemNF, objLivRegES, colLivRegESLinha)
    If lErro <> SUCESSO Then gError 81664
        
    NF_Entrada_Move_Tela_LivRegESLinha = SUCESSO
        
    Exit Function
    
Erro_NF_Entrada_Move_Tela_LivRegESLinha:

    NF_Entrada_Move_Tela_LivRegESLinha = gErr
    
    Select Case gErr
        
        Case 81664
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150318)

    End Select
    
    Exit Function

End Function

Function NF_Entrada_Grava_LivRegES(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal objCliente As ClassCliente, ByVal objFilialCliente As ClassFilialCliente, ByVal objFornecedor As ClassFornecedor, ByVal objFilialFornecedor As ClassFilialFornecedor, ByVal objEndereco As ClassEndereco, ByVal objFilialEmpresa As AdmFiliais) As Long
'Preenche o objLivRegES com os dados da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim sInscricaoAux As String

On Error GoTo Erro_NF_Entrada_Grava_LivRegES
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69224
    
    objLivRegES.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objLivRegES.dtData = objNFiscal.dtDataEntrada
    objLivRegES.dtDataEmissao = objNFiscal.dtDataEmissao
    objLivRegES.iDestinatario = objTipoDocInfo.iDestinatario
    objLivRegES.iEmitente = objTipoDocInfo.iEmitente
    
    If right(UCase(objNFiscal.sSerie), 2) = "-E" Then
        objLivRegES.sSerie = left(objNFiscal.sSerie, Len(objNFiscal.sSerie) - 2)
    Else
        objLivRegES.sSerie = objNFiscal.sSerie
    End If
    
    'nfe federal
    If objNFiscal.iNFe <> 0 And objTipoDocInfo.iModeloArqICMS = 1 Then
        objLivRegES.iModelo = 55
    ElseIf objNFiscal.iNFe = MARCADO And objTipoDocInfo.iModeloArqICMS = 8 Then
        objLivRegES.iModelo = 57 'CONHECIMENTO DE TRANSPORTE ELETRÔNICO – CT-e (CÓDIGO 57)
    Else
        objLivRegES.iModelo = objTipoDocInfo.iModeloArqICMS
    End If
    
    objLivRegES.iOrigem = objTipoDocInfo.iOrigem
    objLivRegES.iSituacao = SITUACAO_NORMAL
    objLivRegES.iStatusLivro = STATUS_FIS_ORIGINAL
    objLivRegES.iTipo = TIPO_REGES_ENTRADA
    objLivRegES.iFreteRespons = objNFiscal.iFreteRespons
    objLivRegES.dPISValor = objNFiscal.objTributacaoNF.dPISCredito
    objLivRegES.dCOFINSValor = objNFiscal.objTributacaoNF.dCOFINSCredito
    
    'Gera número automático para o Item de apuração ICMS
    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69225
    
    objLivRegES.lNumIntDocOrigem = objNFiscal.lNumIntDoc
    objLivRegES.iTipoNumIntDocOrigem = TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL
    objLivRegES.lNumIntDoc = lNumIntDoc
    objLivRegES.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
        
    If objLivRegES.iOrigem = EMITENTE_CLIENTE Then
        
        'Arrumar uma solução viavel
        objLivRegES.sCgc = objFilialCliente.sCgc
        Call Formata_String_Numero(objFilialCliente.sInscricaoEstadual, sInscricaoAux)
        objLivRegES.sInscricaoEstadual = sInscricaoAux
        objLivRegES.sNome = objCliente.sRazaoSocial
        'objLivRegES.sUF = objEndereco.sSiglaEstado
        
        If objEndereco.iCodigoPais <> PAIS_BRASIL Then
            objLivRegES.sUF = "EX"
        Else
            objLivRegES.sUF = objEndereco.sSiglaEstado
        End If
    
    ElseIf objLivRegES.iOrigem = EMITENTE_FORNECEDOR Then
        
        'Arrumar uma solução viavel
        objLivRegES.sCgc = objFilialFornecedor.sCgc
        Call Formata_String_Numero(objFilialFornecedor.sInscricaoEstadual, sInscricaoAux)
        objLivRegES.sInscricaoEstadual = sInscricaoAux
        objLivRegES.sNome = objFornecedor.sRazaoSocial
        'objLivRegES.sUF = objEndereco.sSiglaEstado
  
        If objEndereco.iCodigoPais <> PAIS_BRASIL Then
            objLivRegES.sUF = "EX"
        Else
            objLivRegES.sUF = objEndereco.sSiglaEstado
        End If
    
    ElseIf objLivRegES.iOrigem = EMITENTE_EMPRESA Then
        
        'Arrumar uma solução viavel
        objLivRegES.sCgc = objFilialEmpresa.sCgc
        Call Formata_String_Numero(objFilialEmpresa.sInscricaoEstadual, sInscricaoAux)
        objLivRegES.sInscricaoEstadual = sInscricaoAux
        objLivRegES.sNome = gsNomeEmpresa
        objLivRegES.sUF = objFilialEmpresa.objEndereco.sSiglaEstado
        
    End If
    
    'Faz a inserção na Tabela de LivResES
    lErro = Comando_Executar(lComando, "INSERT INTO LivRegES (NumIntDoc, TipoNumIntDocOrigem, NumIntDocOrigem, StatusLivro, Tipo, CGC, InscricaoEstadual, Nome, Data, DataEmissao, UF, Serie, NumNotaFiscal, Situacao, Modelo, Emitente, Destinatario, Origem, NumIntEmitente, FilialEmpresa, CIF_FOB, PISValor, COFINSValor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objLivRegES.lNumIntDoc, objLivRegES.iTipoNumIntDocOrigem, objLivRegES.lNumIntDocOrigem, objLivRegES.iStatusLivro, objLivRegES.iTipo, objLivRegES.sCgc, objLivRegES.sInscricaoEstadual, objLivRegES.sNome, objLivRegES.dtData, objLivRegES.dtDataEmissao, objLivRegES.sUF, objLivRegES.sSerie, objLivRegES.lNumNotaFiscal, objLivRegES.iSituacao, objLivRegES.iModelo, objLivRegES.iEmitente, objLivRegES.iDestinatario, objLivRegES.iOrigem, objLivRegES.lNumIntEmitente, objLivRegES.iFilialEmpresa, objLivRegES.iFreteRespons, objLivRegES.dPISValor, objLivRegES.dCOFINSValor)
    If lErro <> AD_SQL_SUCESSO Then gError 78410
                
    Call Comando_Fechar(lComando)
                
    NF_Entrada_Grava_LivRegES = SUCESSO
    
    Exit Function
    
Erro_NF_Entrada_Grava_LivRegES:

    NF_Entrada_Grava_LivRegES = gErr

    Select Case gErr
    
        Case 69224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69225
        
        Case 78410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150319)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function NF_Entrada_Move_Tela_Complemento(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long
'Move para as coleções os Itens para o Frete, Seguro, Despesas

Dim lErro As Long

On Error GoTo Erro_NF_Entrada_Move_Tela_Complemento

    'Move os dados do Frete
    If objNFiscal.dValorFrete > 0 Then
        
        lErro = NF_Move_Memoria_LivRegESItemNF_Frete(objNFiscal, objLivRegES, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 69272
                
    End If
    
    'Move os dados do Seguro
    If objNFiscal.dValorSeguro > 0 Then
        
        lErro = NF_Move_Memoria_LivRegESItemNF_Seguro(objNFiscal, objLivRegES, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 69274
                
    End If
    
    'Move os dados do Outras
    If objNFiscal.dValorOutrasDespesas > 0 Then
        
        lErro = NF_Move_Memoria_LivRegESItemNF_Outras(objNFiscal, objLivRegES, colLivRegESItemNF)
        If lErro <> SUCESSO Then gError 69276
                
    End If
        
    NF_Entrada_Move_Tela_Complemento = SUCESSO
    
    Exit Function
    
Erro_NF_Entrada_Move_Tela_Complemento:

    NF_Entrada_Move_Tela_Complemento = gErr

    Select Case gErr
        
        Case 69271, 69272, 69273, 69274, 69275, 69276
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150320)

    End Select
    
    Exit Function

End Function

Function Repreenche_Linha_LivroFiscal(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal objTela As Object) As Long
'Preenche o Grid de Lancamentos dos Livros de Reg de Entrada e de Saída

Dim lErro As Long
'Dim objItemNF As New ClassItemNF
'Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim iIndice As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Repreenche_Linha_LivroFiscal
        
    lErro = Recalcula_Linhas_LivroFiscal(objNFiscal, objLivRegES)
    If lErro <> SUCESSO Then gError 81672
    
    iIndice = 0
    
    With objTela
        
        Call Grid_Limpa(objTela.objGridLancamento)
        
        For Each objLivRegESLinha In objLivRegES.colLinhas
            
            iIndice = iIndice + 1
            
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_UF_Col) = objLivRegES.sUF
            If objLivRegESLinha.dValorContabil > 0 Then .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_ValorContabil_Col) = Format(objLivRegESLinha.dValorContabil, "Standard")
            
            If objLivRegESLinha.sClassifContabil <> "" Then

                lErro = Mascara_RetornaContaEnxuta(objLivRegESLinha.sClassifContabil, sContaEnxuta)
                If lErro <> SUCESSO Then gError 69509

                .Controls("ContabilLanc").PromptInclude = False
                .Controls("ContabilLanc").Text = sContaEnxuta
                .Controls("ContabilLanc").PromptInclude = True

                .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_Contabil_Col) = .Controls("ContabilLanc").Text

            End If
            
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_Fiscal_Col) = objLivRegESLinha.sNaturezaOp
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_BaseCalculoICMS_Col) = Format(objLivRegESLinha.dValorICMSBase, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_AliquotaICMS_Col) = Format(objLivRegESLinha.dAliquotaICMS, "Percent")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_ImpostoCreditadoICMS_Col) = Format(objLivRegESLinha.dValorICMS, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_IsentaICMS_Col) = Format(objLivRegESLinha.dValorICMSIsentoNaoTrib, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_OutrasICMS_Col) = Format(objLivRegESLinha.dValorICMSOutras, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_BaseCalculoIPI_Col) = Format(objLivRegESLinha.dValorIPIBase, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_AliquotaIPI_Col) = Format(objLivRegESLinha.dAliquotaIPI, "Percent")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_ImpostoCreditadoIPI_Col) = Format(objLivRegESLinha.dValorIPI, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_IsentaIPI_Col) = Format(objLivRegESLinha.dValorIPIIsentoNaoTrib, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_OutrasIPI_Col) = Format(objLivRegESLinha.dValorIPIOutras, "Standard")
            .Controls("GridLancamentos").TextMatrix(iIndice, .iGrid_Observacao_Col) = objLivRegESLinha.sObservacaoLivFisc
        
        Next
    
        .objGridLancamento.iLinhasExistentes = iIndice
        
    End With
    
    Repreenche_Linha_LivroFiscal = SUCESSO
    
    Exit Function
    
Erro_Repreenche_Linha_LivroFiscal:

    Repreenche_Linha_LivroFiscal = gErr
    
    Select Case gErr
        
        Case 81672
        
        Case 69509
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objLivRegESLinha.sClassifContabil)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150324)
        
    End Select
    
    Exit Function

End Function

Function Edicao_RegEntrada_Altera_Item(ByVal objItemNF As ClassItemNF, ByVal objNFiscal As ClassNFiscal, ByVal colLivRegESItemNF As Collection) As Long
'Modifica os valores da Tributação para o Item Passado

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_Edicao_RegEntrada_Altera_Item
        
    For Each objLivRegESItemNF In colLivRegESItemNF
    
        If objItemNF.iItem = objLivRegESItemNF.iNumItem Then
            objLivRegESItemNF.dAliquotaICMS = objItemNF.objTributacaoItemNF.dICMSAliquota
            objLivRegESItemNF.dValorICMSBase = objItemNF.objTributacaoItemNF.dICMSBase
            objLivRegESItemNF.dValorICMSSubstBase = objItemNF.objTributacaoItemNF.dICMSSubstBase
            objLivRegESItemNF.dValorIPI = objItemNF.objTributacaoItemNF.dIPIValor
            objLivRegESItemNF.dValorDescontoDespAcess = Round(objItemNF.dValorDesconto + objItemNF.dValorDescGlobal, 2)
            objLivRegESItemNF.sCFOP = objItemNF.objTributacaoItemNF.sNaturezaOp
            
            objLivRegESItemNF.iTipoTribICMS = objItemNF.objTributacaoItemNF.iICMSTipo
            objLivRegESItemNF.iTipoTribIPI = objItemNF.objTributacaoItemNF.iIPITipo
            objLivRegESItemNF.dAliquotaIPI = objItemNF.objTributacaoItemNF.dIPIAliquota
            objLivRegESItemNF.dRedBaseICMS = objItemNF.objTributacaoItemNF.dICMSPercRedBase
            objLivRegESItemNF.dRedBaseIPI = objItemNF.objTributacaoItemNF.dIPIPercRedBase
            objLivRegESItemNF.iTipoTributacao = objItemNF.objTributacaoItemNF.iTipoTributacao
            objLivRegESItemNF.dAliquotaSubst = objItemNF.objTributacaoItemNF.dICMSSubstAliquota
            objLivRegESItemNF.dValorSubst = objItemNF.objTributacaoItemNF.dICMSSubstValor
            objLivRegESItemNF.dValorICMS = objItemNF.objTributacaoItemNF.dICMSValor
            objLivRegESItemNF.dValorIPIBase = objItemNF.objTributacaoItemNF.dIPIBaseCalculo
            objLivRegESItemNF.dValorICMSCredito = objItemNF.objTributacaoItemNF.dICMSCredito
            
            Exit For
        End If
    Next
    
    Edicao_RegEntrada_Altera_Item = SUCESSO
    
    Exit Function
    
Erro_Edicao_RegEntrada_Altera_Item:

    Edicao_RegEntrada_Altera_Item = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150325)

    End Select
    
    Exit Function

End Function

Function Edicao_RegSaida_Altera_Complemento(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long
'Função usada para a alteração dos itens da Nota para a Tela de Edição de Registro de Saida
'Move para a coleção de Linha os Itens para o Frete, Seguro, Despesas
'No caso do Item NF modifica somente a Tributação

Dim lErro As Long

On Error GoTo Erro_Edicao_RegSaida_Altera_Complemento
        
    'Modifica a tributação do item para Frete
    lErro = Edicao_RegEntrada_Altera_Item_Frete(objNFiscal, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69511
                
    'modifica a tributação do item para o Seguro
    lErro = Edicao_RegEntrada_Altera_Item_Seguro(objNFiscal, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69513
    
    'Modifica a tributação do Item para outras despesas
    lErro = Edicao_RegEntrada_Altera_Item_Outras(objNFiscal, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69515
        
    Edicao_RegSaida_Altera_Complemento = SUCESSO
    
    Exit Function
    
Erro_Edicao_RegSaida_Altera_Complemento:

    Edicao_RegSaida_Altera_Complemento = gErr

    Select Case gErr
        
        Case 69510, 69511, 69512, 69513, 69514, 69515
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150326)

    End Select
    
    Exit Function

End Function

Function Edicao_RegEntrada_Altera_Complemento(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection) As Long
'Função usada para a alteração dos itens da Nota para a Tela de Edição de Registro de Entrada
'Move para a coleção de Linha os Itens para o Frete, Seguro, Despesas
'No caso do Item NF modifica somente a Tributação

Dim lErro As Long

On Error GoTo Erro_Edicao_RegEntrada_Altera_Complemento
    
    'Modifica a tributação do item para Frete
    lErro = Edicao_RegEntrada_Altera_Item_Frete(objNFiscal, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69517
            
    'modifica a tributação do item para o Seguro
    lErro = Edicao_RegEntrada_Altera_Item_Seguro(objNFiscal, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69519
            
    'Modifica a tributação do Item para outras despesas
    lErro = Edicao_RegEntrada_Altera_Item_Outras(objNFiscal, colLivRegESItemNF)
    If lErro <> SUCESSO Then gError 69521
                
    Edicao_RegEntrada_Altera_Complemento = SUCESSO
    
    Exit Function
    
Erro_Edicao_RegEntrada_Altera_Complemento:

    Edicao_RegEntrada_Altera_Complemento = gErr

    Select Case gErr
        
        Case 69516, 69517, 69518, 69519, 69520, 69521
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150327)

    End Select
    
    Exit Function

End Function

Function Edicao_RegEntrada_Altera_Item_Frete(ByVal objNFiscal As ClassNFiscal, ByVal colLivRegESItemNF As Collection) As Long
'Modifica os dados da Tributação para o Frete

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_Edicao_RegEntrada_Altera_Item_Frete

    For Each objLivRegESItemNF In colLivRegESItemNF
    
        If objLivRegESItemNF.iNumItem = 991 Then
            objLivRegESItemNF.dValorDescontoDespAcess = objNFiscal.dValorFrete
            Exit For
        End If
    Next
    
    Edicao_RegEntrada_Altera_Item_Frete = SUCESSO
    
    Exit Function
    
Erro_Edicao_RegEntrada_Altera_Item_Frete:

    Edicao_RegEntrada_Altera_Item_Frete = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150328)

    End Select
    
    Exit Function

End Function

Function Edicao_RegEntrada_Altera_Item_Seguro(ByVal objNFiscal As ClassNFiscal, ByVal colLivRegESItemNF As Collection) As Long
'Modifica os dados da Triutação para o Seguro

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_Edicao_RegEntrada_Altera_Item_Seguro

    For Each objLivRegESItemNF In colLivRegESItemNF
    
        If objLivRegESItemNF.iNumItem = 992 Then
            objLivRegESItemNF.dValorDescontoDespAcess = objNFiscal.dValorSeguro
            Exit For
        End If
    Next
        
    Edicao_RegEntrada_Altera_Item_Seguro = SUCESSO
    
    Exit Function
    
Erro_Edicao_RegEntrada_Altera_Item_Seguro:

    Edicao_RegEntrada_Altera_Item_Seguro = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150329)

    End Select
    
    Exit Function

End Function

Function Edicao_RegEntrada_Altera_Item_Outras(ByVal objNFiscal As ClassNFiscal, ByVal colLivRegESItemNF As Collection) As Long
'Modifica os dados de Tributação para Outras despesas

Dim lErro As Long
Dim lNumIntDoc As Long
Dim objLivRegESItemNF As New ClassLivRegESItemNF

On Error GoTo Erro_Edicao_RegEntrada_Altera_Item_Outras

    For Each objLivRegESItemNF In colLivRegESItemNF
    
        If objLivRegESItemNF.iNumItem = 999 Then
            objLivRegESItemNF.dValorDescontoDespAcess = objNFiscal.dValorOutrasDespesas
            Exit For
        End If
    Next
    
    Edicao_RegEntrada_Altera_Item_Outras = SUCESSO
    
    Exit Function
    
Erro_Edicao_RegEntrada_Altera_Item_Outras:

    Edicao_RegEntrada_Altera_Item_Outras = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150330)

    End Select
    
    Exit Function

End Function

Function RegApuracaoICMS_Exclui(ByVal objApuracao As ClassRegApuracao) As Long
'Exclui um Registro de apuração ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDocLivFechado As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim alComando(0 To 2) As Long
Dim lTransacao As Long

On Error GoTo Erro_RegApuracaoICMS_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70020
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70021

    'Seleciona a de apuração ICMS a partir da DataInidical, Final e FilialEmpresa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegApuracaoICMS WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70022

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70023
        
    'Se não encontrou a apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70266
    
    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocLivFechado = APURACAO_ICMS_ABERTA) Then gError 70024
    
    'Verifica se Data Inicial e a data Final da Apuração ICMS são iguais ao do Livro Fiscal correspontente a Filial Empresa passada
    lErro = Comando_Executar(alComando(1), "SELECT DataInicial, DataFinal FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", dtDataInicial, dtDataFinal, LIVRO_APURACAO_ICMS_CODIGO, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70028
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70029
    
    'Se não encontrou o Livro Fiscal, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70030
    
    'Se a Data Inicial e a Final da Apuração ICMS não são iguais ao do Livro Fiscal correspondente, Erro
    If dtDataInicial <> objApuracao.dtDataInicial Or dtDataFinal <> objApuracao.dtDataFinal Then gError 70031
    
    'Faz "lock" em ApuracaoICMS
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70025

    'Exclui ApuracaoICMS do Banco de dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RegApuracaoICMS", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70026

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70027

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegApuracaoICMS_Exclui = SUCESSO

    Exit Function

Erro_RegApuracaoICMS_Exclui:

    RegApuracaoICMS_Exclui = gErr

    Select Case gErr

        Case 70020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70022, 70023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)

        Case 70024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMS_FECHADO", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 70025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOICMS", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 70026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGAPURACAOICMS", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 70027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 70028, 70029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
            
        Case 70030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr, LIVRO_APURACAO_ICMS_CODIGO, objApuracao.iFilialEmpresa)
            
        Case 70031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_DIFERENTE_LIVROFILIAL", gErr, objApuracao.iFilialEmpresa, dtDataInicial, dtDataFinal)
            
        Case 70266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMS_NAO_CADASTRADA", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150332)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoICMS_Grava(ByVal objApuracao As ClassRegApuracao) As Long
'Grava um Registro de apuração ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim lNumIntDocLivFechado As Long
Dim lNumIntDoc As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim sInscricaoAux As String

On Error GoTo Erro_RegApuracaoICMS_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70032
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70033

    'Verifica se Data Inicial e a data Final do APuração ICMS são iguais ao do Livro Fiscal correspontente a Filial Empresa passada
    lErro = Comando_Executar(alComando(1), "SELECT DataInicial, DataFinal FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", dtDataInicial, dtDataFinal, LIVRO_APURACAO_ICMS_CODIGO, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70048
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70049
    
    'Se não encontrou o Livro Fiscal, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70050
    
    'Se a Data Inicial e a Final da Apuração ICMS não são iguais ao do Livro Fiscal correspondente, Erro
    If dtDataInicial <> objApuracao.dtDataInicial Or dtDataFinal <> objApuracao.dtDataFinal Then gError 70051

    'Verifica se a apuração ICMS está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegApuracaoICMS WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70034

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70035

    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocLivFechado = APURACAO_ICMS_ABERTA) Then gError 70036
    
    'Retira os separadores da Inscricao Estadual
    Call Formata_String_Numero(objApuracao.sInscricaoEstadual, sInscricaoAux)

    objApuracao.sInscricaoEstadual = sInscricaoAux
    
    'Se a Apuração ICMS já estiver cadastrado ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no RegApuracaoICMS
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70052

        'Atualiza os dados da apuração ICMS gravando em Saldo Inicial o Saldo Credor Anterior
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoICMS SET SaldoCredorInicial = ?, DataEntregaGIA = ?, LocalEntregaGIA = ?, Observacoes =?, CGC = ?, InscricaoEstadual = ?, Nome = ?, Municipio = ?, UF = ?, Logradouro = ?, Numero = ?, Complemento = ?, Bairro = ?, CEP = ?, Contato = ?, TelContato = ?", alComando(0), _
        objApuracao.dSaldoCredorInicial, objApuracao.dtDataEntregaGIA, objApuracao.sLocalEntregaGIA, objApuracao.sObservacoes, objApuracao.sCgc, objApuracao.sInscricaoEstadual, objApuracao.sNome, objApuracao.sMunicipio, objApuracao.sUF, objApuracao.sLogradouro, objApuracao.lNumero, objApuracao.sComplemento, objApuracao.sBairro, objApuracao.sCEP, objApuracao.sContato, objApuracao.sTelContato)
        If lErro <> AD_SQL_SUCESSO Then gError 70053

    'Se a Apuração ICMS não estava cadastrada ==> INSERÇÃO
    Else

        'Gera número automático para a de apuração ICMS colocando
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_REGAPURICMS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 70054

        'Guarda o NumIntDoc da apuração ICMS
        objApuracao.lNumIntDoc = lNumIntDoc
        
        'Insere nova apuração ICMS
        lErro = Comando_Executar(alComando(0), "INSERT INTO RegApuracaoICMS (NumIntDoc, FilialEmpresa, DataInicial, DataFinal, SaldoCredorInicial, DataEntregaGIA, LocalEntregaGIA, Observacoes, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objApuracao.lNumIntDoc, objApuracao.iFilialEmpresa, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.dSaldoCredorInicial, objApuracao.dtDataEntregaGIA, objApuracao.sLocalEntregaGIA, objApuracao.sObservacoes, objApuracao.sCgc, objApuracao.sInscricaoEstadual, objApuracao.sNome, objApuracao.sMunicipio, objApuracao.sUF, objApuracao.sLogradouro, objApuracao.lNumero, objApuracao.sComplemento, objApuracao.sBairro, objApuracao.sCEP, objApuracao.sContato, objApuracao.sTelContato)
        If lErro <> AD_SQL_SUCESSO Then gError 70055

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70056

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegApuracaoICMS_Grava = SUCESSO

    Exit Function

Erro_RegApuracaoICMS_Grava:

    RegApuracaoICMS_Grava = gErr

    Select Case gErr

        Case 70032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70034, 70035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)

        Case 70036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMS_FECHADO", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
        
        Case 70048, 70049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
                    
        Case 70050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr, LIVRO_APURACAO_ICMS_CODIGO, objApuracao.iFilialEmpresa)
               
        Case 70051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_DIFERENTE_LIVROFILIAL", gErr, objApuracao.iFilialEmpresa, dtDataInicial, dtDataFinal)
        
        Case 70052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOICMS", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 70053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOICMS", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
        
        Case 70054
        
        Case 70055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGAPURACAOICMS", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 70056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150333)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoIPIItem_Exclui(ByVal objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Exclui um Tipo de Registro de apuração IPI

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDocApuracao As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long

On Error GoTo Erro_RegApuracaoIPIItem_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 79064
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 79065

    'Seleciona o Item de apuração IPI a partir do código, Data e Descrição passados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocApuracao FROM RegApuracaoIPIItem WHERE TipoReg = ? AND Descricao = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocApuracao, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79066

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79067

    'Se não encontrou o Item de apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79068
    
    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocApuracao = APURACAO_IPI_ABERTA) Then gError 79069

    'Faz "lock" em ApuracaoIPIItem
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 79070

    'Exclui ApuracaoIPIItem do Banco de dados
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RegApuracaoIPIItem", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 79071

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 79072

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegApuracaoIPIItem_Exclui = SUCESSO

    Exit Function

Erro_RegApuracaoIPIItem_Exclui:

    RegApuracaoIPIItem_Exclui = gErr

    Select Case gErr

        Case 79064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 79066, 79067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)

        Case 79068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPIITEM_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)
        
        Case 79069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPIITEM_FECHADO", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOIPIITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGAPURACAOIPIITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150334)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoIPIItem_Grava(ByVal objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Grava um Item de Registro de apuração IPI

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lNumIntDocApuracao As Long
Dim lNumIntDoc As Long
Dim iLancamento As Integer

On Error GoTo Erro_RegApuracaoIPIItem_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 79073
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 79074

    'Verifica se este Item permite Lancamentos
    lErro = Comando_Executar(alComando(2), "SELECT Lancamento FROM TiposRegApuracaoIPI WHERE Codigo = ?", iLancamento, objRegApuracaoItem.iTipoReg)
    If lErro <> AD_SQL_SUCESSO Then gError 79075
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79076
    
    'Se não encontrou o tipo de registro de apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79077
        
    'Se não permite lancamentos --> Erro
    If iLancamento = TIPOREGAPURACAO_NAO_PERMITE_LANCAMENTO Then gError 79078
    
    'Verifica se o Item de apuração IPI está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocApuracao FROM RegApuracaoIPIItem WHERE TipoReg = ? AND Descricao = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocApuracao, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79079
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79080

    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocApuracao = APURACAO_IPI_ABERTA) Then gError 79081

    'Se o Item já estiver cadastrado ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no RegApuracaoIPIItem
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 79082

        'Atualiza o valor do item de apuração IPI
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoIPIItem SET Valor = ?", alComando(0), objRegApuracaoItem.dValor)
        If lErro <> AD_SQL_SUCESSO Then gError 79083

    'Se o Item de registro de apuração IPI nãi estava cadastrada ==> INSERÇÃO
    Else

        'Gera número automático para o Item de apuração IPI
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_REGAPURIPIITEM", lNumIntDoc)
        If lErro <> SUCESSO Then gError 79084
        
        'Guarda o NumIntDoc do item de apuração IPI
        objRegApuracaoItem.lNumIntDoc = lNumIntDoc
        
        'Insere novo Item de apuração IPI
        lErro = Comando_Executar(alComando(0), "INSERT INTO RegApuracaoIPIItem (NumIntDoc, TipoReg, Descricao, Data, Valor, FilialEmpresa) VALUES (?,?,?,?,?,?)", objRegApuracaoItem.lNumIntDoc, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, objRegApuracaoItem.dValor, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 79085

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 79086

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegApuracaoIPIItem_Grava = SUCESSO

    Exit Function

Erro_RegApuracaoIPIItem_Grava:

    RegApuracaoIPIItem_Grava = gErr

    Select Case gErr

        Case 79073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 79079, 79080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)

        Case 79081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPIITEM_FECHADO", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOIPIITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOIPIITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGAPURACAOIPIITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 79086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 79084
        
        Case 79075, 79076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOIPI", gErr)
                
        Case 79077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOIPI_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg)
                
        Case 79078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOIPI_NAO_ACEITA_LANCAMENTO", gErr, objRegApuracaoItem.iTipoReg)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150335)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoICMSItem_Exclui(ByVal objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Exclui um Tipo de Registro de apuração ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDocApuracao As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long

On Error GoTo Erro_RegApuracaoICMSItem_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67956
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 67957

    'Seleciona o Item de apuração ICMS a partir do código, Data e Descrição passados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocApuracao FROM RegApuracaoICMSItem WHERE TipoReg = ? AND Descricao = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocApuracao, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67958

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67959

    'Se não encontrou o Item de apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70267
    
    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocApuracao = APURACAO_ICMS_ABERTA) Then gError 67960

    'Faz "lock" em ApuracaoICMSItem
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 67961

    'Exclui ApuracaoICMSItem do Banco de dados
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RegApuracaoICMSItem", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 67962

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 67963

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegApuracaoICMSItem_Exclui = SUCESSO

    Exit Function

Erro_RegApuracaoICMSItem_Exclui:

    RegApuracaoICMSItem_Exclui = gErr

    Select Case gErr

        Case 67956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 67958, 67959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSITEM", gErr)

        Case 67960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMSITEM_FECHADO", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOICMSITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGAPURACAOICMSITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 70267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMSITEM_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150336)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoICMSItem_Grava(ByVal objRegApuracaoItem As ClassRegApuracaoItem) As Long
'Grava um Item de Registro de apuração ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lNumIntDocApuracao As Long
Dim lNumIntDoc As Long
Dim iLancamento As Integer

On Error GoTo Erro_RegApuracaoICMSItem_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67965
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 67966

    'Verifica se este Item permite Lancamentos
    lErro = Comando_Executar(alComando(2), "SELECT Lancamento FROM TiposRegApuracaoICMS WHERE Codigo = ?", iLancamento, objRegApuracaoItem.iTipoReg)
    If lErro <> AD_SQL_SUCESSO Then gError 69998
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69999
    
    'Se não encontrou o tipo de registro de apuração ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 78000
        
    'Se não permite lancamentos --> Erro
    If iLancamento = TIPOREGAPURACAO_NAO_PERMITE_LANCAMENTO Then gError 78001
    
    'Verifica se o Item de apuração ICMS está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocApuracao FROM RegApuracaoICMSItem WHERE TipoReg = ? AND Descricao = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocApuracao, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67967
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67968

    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocApuracao = APURACAO_ICMS_ABERTA) Then gError 67969

    'Se o Item já estiver cadastrado ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no RegApuracaoICMSItem
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 67970

        'Atualiza o valor do item de apuração ICMS
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoICMSItem SET Valor = ?", alComando(0), objRegApuracaoItem.dValor)
        If lErro <> AD_SQL_SUCESSO Then gError 67971

    'Se o Item de registro de apuração ICMS nãi estava cadastrada ==> INSERÇÃO
    Else

        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_REGAPURICMSITEM", lNumIntDoc)
        If lErro <> SUCESSO Then gError 67974
        
        'Guarda o NumIntDoc do item de apuração ICMS
        objRegApuracaoItem.lNumIntDoc = lNumIntDoc
        
        'Insere novo Item de apuração ICMS
        lErro = Comando_Executar(alComando(0), "INSERT INTO RegApuracaoICMSItem (NumIntDoc, TipoReg, Descricao, Data, Valor, FilialEmpresa) VALUES (?,?,?,?,?,?)", objRegApuracaoItem.lNumIntDoc, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData, objRegApuracaoItem.dValor, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 67972

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 67973

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegApuracaoICMSItem_Grava = SUCESSO

    Exit Function

Erro_RegApuracaoICMSItem_Grava:

    RegApuracaoICMSItem_Grava = gErr

    Select Case gErr

        Case 67965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 67967, 67968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSITEM", gErr)

        Case 67969
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOICMSITEM_FECHADO", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOICMSITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67971
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOICMSITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGAPURACAOICMSITEM", gErr, objRegApuracaoItem.iTipoReg, objRegApuracaoItem.sDescricao, objRegApuracaoItem.dtData)

        Case 67973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 67974
        
        Case 69998, 69999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOICMS", gErr)
                
        Case 78000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOICMS_NAO_CADASTRADA", gErr, objRegApuracaoItem.iTipoReg)
                
        Case 78001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOICMS_NAO_ACEITA_LANCAMENTO", gErr, objRegApuracaoItem.iTipoReg)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150337)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoIPI_Exclui(ByVal objApuracao As ClassRegApuracao) As Long
'Exclui um Registro de apuração IPI

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntDocLivFechado As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim alComando(0 To 2) As Long
Dim lTransacao As Long

On Error GoTo Erro_RegApuracaoIPI_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 79120
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 79121

    'Seleciona a de apuração IPI a partir da DataInidical, Final e FilialEmpresa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegApuracaoIPI WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79122

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79123
        
    'Se não encontrou a apuração IPI, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79124
    
    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocLivFechado = APURACAO_IPI_ABERTA) Then gError 79125
    
    'Verifica se Data Inicial e a data Final da Apuração IPI são iguais ao do Livro Fiscal correspontente a Filial Empresa passada
    lErro = Comando_Executar(alComando(1), "SELECT DataInicial, DataFinal FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", dtDataInicial, dtDataFinal, LIVRO_APURACAO_IPI_CODIGO, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79126
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79127
    
    'Se não encontrou o Livro Fiscal, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79128
    
    'Se a Data Inicial e a Final da Apuração IPI não são iguais ao do Livro Fiscal correspondente, Erro
    If dtDataInicial <> objApuracao.dtDataInicial Or dtDataFinal <> objApuracao.dtDataFinal Then gError 79129
    
    'Faz "lock" em ApuracaoIPI
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 79130

    'Exclui ApuracaoIPI do Banco de dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RegApuracaoIPI", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 79131

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 79132

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegApuracaoIPI_Exclui = SUCESSO

    Exit Function

Erro_RegApuracaoIPI_Exclui:

    RegApuracaoIPI_Exclui = gErr

    Select Case gErr

        Case 79120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 79122, 79123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)

        Case 79124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPI_NAO_CADASTRADA", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
        
        Case 79125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPI_FECHADO", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 79126, 79127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
            
        Case 79128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr, LIVRO_APURACAO_IPI_CODIGO, objApuracao.iFilialEmpresa)
            
        Case 79129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_DIFERENTE_LIVROFILIAL", gErr, objApuracao.iFilialEmpresa, dtDataInicial, dtDataFinal)
            
        Case 79130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOIPI", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 79131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGAPURACAOIPI", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 79132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150338)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegApuracaoIPI_Grava(ByVal objApuracao As ClassRegApuracao) As Long
'Grava um Registro de apuração IPI

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lNumIntDocLivFechado As Long
Dim lNumIntDoc As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim sInscricaoAux As String

On Error GoTo Erro_RegApuracaoIPI_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 79133
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 79134

    'Verifica se Data Inicial e a data Final do APuração IPI são iguais ao do Livro Fiscal correspontente a Filial Empresa passada
    lErro = Comando_Executar(alComando(1), "SELECT DataInicial, DataFinal FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", dtDataInicial, dtDataFinal, LIVRO_APURACAO_IPI_CODIGO, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79135
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79136
    
    'Se não encontrou o Livro Fiscal, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79137
    
    'Se a Data Inicial e a Final da Apuração IPI não são iguais ao do Livro Fiscal correspondente, Erro
    If dtDataInicial <> objApuracao.dtDataInicial Or dtDataFinal <> objApuracao.dtDataFinal Then gError 79138

    'Verifica se a apuração IPI está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegApuracaoIPI WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 79139

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79140

    'Se a Apuração já foi fechada, erro
    If Not (lNumIntDocLivFechado = APURACAO_IPI_ABERTA) Then gError 79141

    'Se a Apuração IPI já estiver cadastrado ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no RegApuracaoIPI
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 79142

        'Atualiza os dados da apuração IPI gravando em Saldo Inicial o Saldo Credor Anterior
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoIPI SET SaldoCredorInicial = ?, DataEntregaGIA = ?, LocalEntregaGIA = ?, Observacoes =?, CGC = ?, InscricaoEstadual = ?, Nome = ?, Municipio = ?, UF = ?, Logradouro = ?, Numero = ?, Complemento = ?, Bairro = ?, CEP = ?, Contato = ?, TelContato = ?", alComando(0), _
        objApuracao.dSaldoCredorInicial, objApuracao.dtDataEntregaGIA, objApuracao.sLocalEntregaGIA, objApuracao.sObservacoes, objApuracao.sCgc, objApuracao.sInscricaoEstadual, objApuracao.sNome, objApuracao.sMunicipio, objApuracao.sUF, objApuracao.sLogradouro, objApuracao.lNumero, objApuracao.sComplemento, objApuracao.sBairro, objApuracao.sCEP, objApuracao.sContato, objApuracao.sTelContato)
        If lErro <> AD_SQL_SUCESSO Then gError 79143

    'Se a Apuração IPI não estava cadastrada ==> INSERÇÃO
    Else

        'Gera número automático para a de apuração IPI colocando
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_REGAPURIPI", lNumIntDoc)
        If lErro <> SUCESSO Then gError 79144

        'Guarda o NumIntDoc da apuração IPI
        objApuracao.lNumIntDoc = lNumIntDoc
        
        'Retira os separadores da Inscricao Estadual
        Call Formata_String_Numero(objApuracao.sInscricaoEstadual, sInscricaoAux)
        
        objApuracao.sInscricaoEstadual = sInscricaoAux
        
        'Insere nova apuração IPI
        lErro = Comando_Executar(alComando(2), "INSERT INTO RegApuracaoIPI (NumIntDoc, FilialEmpresa, DataInicial, DataFinal, SaldoCredorInicial, DataEntregaGIA, LocalEntregaGIA, Observacoes, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objApuracao.lNumIntDoc, objApuracao.iFilialEmpresa, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.dSaldoCredorInicial, objApuracao.dtDataEntregaGIA, objApuracao.sLocalEntregaGIA, objApuracao.sObservacoes, objApuracao.sCgc, objApuracao.sInscricaoEstadual, objApuracao.sNome, objApuracao.sMunicipio, objApuracao.sUF, objApuracao.sLogradouro, objApuracao.lNumero, objApuracao.sComplemento, objApuracao.sBairro, objApuracao.sCEP, objApuracao.sContato, objApuracao.sTelContato)
        If lErro <> AD_SQL_SUCESSO Then gError 79145

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 79146

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegApuracaoIPI_Grava = SUCESSO

    Exit Function

Erro_RegApuracaoIPI_Grava:

    RegApuracaoIPI_Grava = gErr

    Select Case gErr

        Case 79133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 79135, 79136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
                    
        Case 79137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr, LIVRO_APURACAO_IPI_CODIGO, objApuracao.iFilialEmpresa)
               
        Case 79138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_APURACAO_DIFERENTE_LIVROFILIAL", gErr, objApuracao.iFilialEmpresa, dtDataInicial, dtDataFinal)
        
        Case 79139, 79140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)

        Case 79141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOIPI_FECHADO", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
        
        Case 79142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOIPI", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 79143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOIPI", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)
        
        Case 79144
        
        Case 79145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGAPURACAOIPI", gErr, objApuracao.dtDataInicial, objApuracao.dtDataFinal, objApuracao.iFilialEmpresa)

        Case 79146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150339)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function GNRICMS_Exclui(ByVal objGNRICMS As ClassGNRICMS) As Long
'Exclui uma Guia de Recolhimento de ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim lNumIntRegApuracaoICMS As Long
Dim lNumIntArqICMS As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long

On Error GoTo Erro_GNRICMS_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70108
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70109

    'Seleciona a Guia de ICMS a partir do código
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntRegApuracaoICMS, NumIntArqICMS FROM GNRICMS WHERE Codigo = ?", 0, lNumIntRegApuracaoICMS, lNumIntArqICMS, objGNRICMS.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 70110

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70111
    
    'Se não encontrou a Guia de ICMS, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70265
    
    'Se a Apuração relacionada a Guia ICMS em questão já foi fechada, erro
    If Not (lNumIntRegApuracaoICMS = APURACAO_ICMS_ABERTA) Then gError 70112

    'Se o arquivo de ICMS relacionado a Guia de ICMS em questão já foi fechado
    If Not (lNumIntArqICMS = ARQUIVO_ICMS_ABERTO) Then gError 70113
    
    'Faz "lock" na Guia de ICMS
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70114

    'Exclui Guia de ICMS do Banco de dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM GNRICMS", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70115

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70116

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    GNRICMS_Exclui = SUCESSO

    Exit Function

Erro_GNRICMS_Exclui:

    GNRICMS_Exclui = gErr

    Select Case gErr

        Case 70108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70110, 70111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GNRICMS", gErr)

        Case 70112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GNRICMS_VINCULADO_APURACAOICMS", gErr, objGNRICMS.lCodigo)
        
        Case 70113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GNRICMS_VINCULADO_ARQICMS", gErr, objGNRICMS.lCodigo)
            
        Case 70114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_GNRICMS", gErr, objGNRICMS.lCodigo)

        Case 70115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GNRICMS", gErr, objGNRICMS.lCodigo)

        Case 70116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 70265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GNRICMS_NAO_CADASTRADA", gErr, objGNRICMS.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150340)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function GNRICMS_Grava(ByVal objGNRICMS As ClassGNRICMS) As Long
'Grava uma Guia de ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim lNumIntRegApuracaoICMS As Long
Dim lNumIntArqICMS As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_GNRICMS_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70117
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70118

    'Verifica se a Guia de ICMS está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntRegApuracaoICMS, NumIntArqICMS FROM GNRICMS WHERE Codigo = ?", 0, lNumIntRegApuracaoICMS, lNumIntArqICMS, objGNRICMS.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 70119

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70120

    'Se a Apuração relacionada a Guia ICMS em questão já foi fechada, erro
    If Not (lNumIntRegApuracaoICMS = APURACAO_ICMS_ABERTA) Then gError 70121

    'Se o arquivo de ICMS relacionado a Guia de ICMS em questão já foi fechado
    If Not (lNumIntArqICMS = ARQUIVO_ICMS_ABERTO) Then gError 70122

    'Se a Guia de ICMS já estiver cadastrada ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no Guia de ICMS
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70123

        'Atualiza os dados da Guia de ICMS
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE GNRICMS SET Tipo = ?, DataPagto = ?, CGCSubstTrib = ?, InscricaoEstadual = ?, UFSubstTrib = ?, UFDestino = ?, Banco = ?, Agencia = ?, Numero = ?, Valor = ?, Vencimento = ?, DataRef = ?, Convenio = ?, FilialEmpresa = ?,CodReceita = ?,CodObrigRecolher = ?", alComando(0), objGNRICMS.iTipo, objGNRICMS.dtDataPagto, objGNRICMS.sCGCSubstTrib, objGNRICMS.sInscricaoEstadual, objGNRICMS.sUFSubstTrib, objGNRICMS.sUFDestino, objGNRICMS.iBanco, objGNRICMS.iAgencia, objGNRICMS.sNumero, objGNRICMS.dValor, objGNRICMS.dtVencimento, objGNRICMS.dtDataRef, objGNRICMS.sConvenio, objGNRICMS.iFilialEmpresa, objGNRICMS.sCodReceita, objGNRICMS.sCodObrigRecolher)
        If lErro <> AD_SQL_SUCESSO Then gError 70124

    'Se a Guia de ICMS não estava cadastrada ==> INSERÇÃO
    Else

        'Gera número automático para a Guia de ICMS colocando
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_GNRICMS", lNumIntDoc)
        If lErro <> SUCESSO Then gError 70125

        'Guarda o NumIntDoc da Guia de ICMS
        objGNRICMS.lNumIntDoc = lNumIntDoc

        'Insere nova Guia de ICMS
        lErro = Comando_Executar(alComando(0), "INSERT INTO GNRICMS (NumIntDoc, Codigo, Tipo, DataPagto, CGCSubstTrib, InscricaoEstadual, UFSubstTrib, UFDestino, Banco, Agencia, Numero, Valor, Vencimento, DataRef, Convenio, FilialEmpresa,CodReceita,CodObrigRecolher) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objGNRICMS.lNumIntDoc, objGNRICMS.lCodigo, objGNRICMS.iTipo, objGNRICMS.dtDataPagto, objGNRICMS.sCGCSubstTrib, objGNRICMS.sInscricaoEstadual, objGNRICMS.sUFSubstTrib, objGNRICMS.sUFDestino, objGNRICMS.iBanco, objGNRICMS.iAgencia, objGNRICMS.sNumero, objGNRICMS.dValor, objGNRICMS.dtVencimento, objGNRICMS.dtDataRef, objGNRICMS.sConvenio, objGNRICMS.iFilialEmpresa, objGNRICMS.sCodReceita, objGNRICMS.sCodObrigRecolher)
        If lErro <> AD_SQL_SUCESSO Then gError 70126

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70127

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    GNRICMS_Grava = SUCESSO

    Exit Function

Erro_GNRICMS_Grava:

    GNRICMS_Grava = gErr

    Select Case gErr

        Case 70117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70119, 70120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GNRICMS", gErr)

        Case 70121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GNRICMS_VINCULADO_APURACAOICMS", gErr, objGNRICMS.lCodigo)
            
        Case 70122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GNRICMS_VINCULADO_ARQICMS", gErr, objGNRICMS.lCodigo)

        Case 70123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_GNRICMS", gErr, objGNRICMS.lCodigo)

        Case 70124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GNRICMS", gErr, objGNRICMS.lCodigo)

        Case 70125

        Case 70126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GNRICMS", gErr, objGNRICMS.lCodigo)

        Case 70127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150341)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function GNRICMS_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próximo Tipo de Registro de apuração ICMS

Dim lErro As Long
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String
Dim sCampo As String

On Error GoTo Erro_GNRICMS_Codigo_Automatico

    sCodigo = "FISC_NUM_PROX_GNRICMS"
    sTabelaConfig = "FISConfig"
    sTabelaObjeto = "GNRICMS"
    sCampo = "Codigo"

    'Chama Config_ObterAutomatico
    lErro = CF("Config_ObterAutomatico", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
    If lErro <> SUCESSO Then gError 70091

    GNRICMS_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_GNRICMS_Codigo_Automatico:

    GNRICMS_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 70091 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150342)

    End Select

    Exit Function

End Function

Function GeracaoArqICMS_Limpar_TabelasAuxiliares() As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lNumIntReg As Long
Dim lNumIntGeracao As Long
Dim iIndice As Integer

On Error GoTo Erro_GeracaoArqICMS_Limpar_TabelasAuxiliares

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69831
    Next

    'Le os dados Auxiliares para Exclusão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntReg FROM GeracaoArqICMS", 0, lNumIntReg)
    If lErro <> AD_SQL_SUCESSO Then gError 69832

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69833
            
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui os dados Auxiliares
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM GeracaoArqICMS", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 69834
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69995
            
    Loop
    
    'Le os dados Auxiliares para Exclusão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntLivRegESCadProd FROM GeracaoArqICMSProd", 0, lNumIntGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 69835

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69836
        
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui os dados Auxiliares
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM GeracaoArqICMSProd", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 69837
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69996
    
    Loop
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    GeracaoArqICMS_Limpar_TabelasAuxiliares = SUCESSO

    Exit Function

Erro_GeracaoArqICMS_Limpar_TabelasAuxiliares:

    GeracaoArqICMS_Limpar_TabelasAuxiliares = gErr

    Select Case gErr
        
        Case 69831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69832, 69833, 69995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERACAOARQICMS", gErr)
        
        Case 69834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GERACAOARQICMS", gErr)
        
        Case 69835, 69836, 69996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GERACAOARQICMSPROD", gErr)
        
        Case 69837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GERACAOARQICMSPROD", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150343)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function GeracaoArqICMS_Grava_FIS(ByVal objInfoArqICMS As ClassInfoArqICMS, ByVal iTipo As Integer, ByVal lComando As Long) As Long
'Grava o Tipo no BD
'Tem que ser chamado  em transacao

Dim lErro As Long

On Error GoTo Erro_GeracaoArqICMS_Grava_FIS

    objInfoArqICMS.lNumIntReg = objInfoArqICMS.lNumIntReg + 1

    'Insere no Bd o Tipo passado
    lErro = Comando_Executar(lComando, "INSERT INTO GeracaoArqICMS (NumIntReg, UFDest, Tipo, ChaveOrd, Conteudo) VALUES (?,?,?,?,?)", objInfoArqICMS.lNumIntReg, objInfoArqICMS.sUFDestino, iTipo, objInfoArqICMS.sChaveOrd, objInfoArqICMS.sBuffer)
    If lErro <> AD_SQL_SUCESSO Then gError 69830

    GeracaoArqICMS_Grava_FIS = SUCESSO

    Exit Function

Erro_GeracaoArqICMS_Grava_FIS:

    GeracaoArqICMS_Grava_FIS = gErr

    Select Case gErr

        Case 69830
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABELA_GERACAOARQICMS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150344)

    End Select

    Exit Function

End Function

Function InfoArqICMS_Grava(ByVal objInfoArqICMS As ClassInfoArqICMS) As Long

Dim lErro As Long
Dim tInfoArqICMS As typeInfoArqICMS
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sInscEstAux As String

On Error GoTo Erro_InfoArqICMS_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69838
    Next
    
    tInfoArqICMS.sNomeArquivo = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    'Verifica se existe um Arquivo de ICMS com as datas sobrepondo
    lErro = Comando_Executar(alComando(1), "SELECT DataInicial, DataFinal, NomeArquivo FROM InfoArqICMS WHERE FilialEmpresa = ? AND Integral = ? AND ((DataInicial >= ? AND DataInicial <= ?) OR (DataFinal >= ? AND DataFinal <= ?) OR (DataInicial <= ? AND DataFinal >= ?))", tInfoArqICMS.dtDataInicial, tInfoArqICMS.dtDataFinal, tInfoArqICMS.sNomeArquivo, objInfoArqICMS.iFilialEmpresa, IIf(objInfoArqICMS.bIntegral, 1, 0), objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 69839

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69840

    If lErro = AD_SQL_SUCESSO Then gError 69841

    'Gera número automático para o Arquivo de ICMS
    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_INFOARQICMS", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69842

    objInfoArqICMS.lNumIntDoc = lNumIntDoc
    
    Call Formata_String_Numero(objInfoArqICMS.sInscricaoEstadual, sInscEstAux)
    
    objInfoArqICMS.sInscricaoEstadual = sInscEstAux

    'Insere Arquivo ICMS
    lErro = Comando_Executar(alComando(2), "INSERT INTO InfoArqICMS (NumIntDoc, FilialEmpresa,DataInicial, DataFinal, Integral, CGC, InscricaoEstadual, Nome, Municipio, UF, Logradouro, Numero, Complemento, Bairro, CEP, Contato, TelContato, NomeArquivo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objInfoArqICMS.lNumIntDoc, objInfoArqICMS.iFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, IIf(objInfoArqICMS.bIntegral, 1, 0), objInfoArqICMS.sCgc, objInfoArqICMS.sInscricaoEstadual, objInfoArqICMS.sNome, objInfoArqICMS.sMunicipio, objInfoArqICMS.sUF, objInfoArqICMS.sLogradouro, objInfoArqICMS.lNumero, objInfoArqICMS.sComplemento, objInfoArqICMS.sBairro, objInfoArqICMS.sCEP, objInfoArqICMS.sContato, objInfoArqICMS.sTelContato, objInfoArqICMS.sNomeArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 69843
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    InfoArqICMS_Grava = SUCESSO

    Exit Function

Erro_InfoArqICMS_Grava:

    InfoArqICMS_Grava = gErr

    Select Case gErr

        Case 69838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69839, 69840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOARQICMS", gErr)

        Case 69841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INFOARQICMS_JA_CADASTRADO", gErr, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, tInfoArqICMS.dtDataInicial, tInfoArqICMS.dtDataFinal)
        
        Case 69842
        
        Case 69843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_INFOARQICMS", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150345)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function ConfiguraFIS_GravarTrans(ByVal objFIS As ClassFIS) As Long
'Efetua a gravação dos dados com a transação já aberta em função _
 anterior

Dim lErro As Long
Dim sConteudo As String
Dim objFISConfig As New ClassFISConfig

On Error GoTo Erro_ConfiguraFIS_GravarTrans

    'Carrega o objConfig com as informações necessárias
    objFISConfig.sConteudo = CStr(objFIS.iCodFiscalServico)
    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_CODIGO_FISCAL_SERVICOS"

    lErro = FISConfig_GravarTrans(objFISConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Carrega o objConfig com as informações necessárias
    objFISConfig.sConteudo = CStr(objFIS.dtFisBloqDataLimite)
    objFISConfig.iFilialEmpresa = EMPRESA_TODA
    objFISConfig.sCodigo = "FISC_BLOQUEIO_DATA_LIMITE"

    lErro = FISConfig_GravarTrans(objFISConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    ConfiguraFIS_GravarTrans = SUCESSO

    Exit Function
    
Erro_ConfiguraFIS_GravarTrans:
    
    ConfiguraFIS_GravarTrans = gErr
    
        Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150347)

    End Select

    Exit Function

End Function
        
Function TipoRegApuracaoICMS_Exclui(ByVal objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Exclui um Tipo de Registro de apuração ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim iPreCadastrado As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TipoRegApuracaoICMS_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67896
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 67897
    
    'Seleciona o Tipo de registro de apuração ICMS a partir do código passado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT PreCadastrado FROM TiposRegApuracaoICMS WHERE Codigo = ?", 0, iPreCadastrado, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 67898

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67899
    
    'Se não encontrou o Tipo de Registro de Apuração ICMS, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70264
    
    'Se for um tipo Pré-cadastrado, erro
    If iPreCadastrado = TIPOREGAPURACAO_PRE_CADASTRADO Then gError 67900
    
    'Faz "lock" em TipoRegApuracaoICMS
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 67901

    'Verifica se existe relaciomento entre o TipoRegApuracaoICMS e RegApuracaoICMSItem
    lErro = Comando_Executar(alComando(1), "SELECT RegApuracaoICMSItem.NumIntDoc FROM RegApuracaoICMSItem WHERE RegApuracaoICMSItem.TipoReg = ?", lNumIntDoc, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 67927
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67928
    
    'Se encontrou relacionamento, Erro
    If lErro = AD_SQL_SUCESSO Then gError 67929
    
    'Exclui o Tipo de Reg. de apuração ICMS do Banco de dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TiposRegApuracaoICMS", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 67902

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 67903

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoRegApuracaoICMS_Exclui = SUCESSO

    Exit Function

Erro_TipoRegApuracaoICMS_Exclui:

    TipoRegApuracaoICMS_Exclui = gErr

    Select Case gErr

        Case 67896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 67898, 67899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOICMS", gErr)

        Case 67900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOICMS_PRE_CADASTRADO", gErr, objTipoRegApuracao.iCodigo)
        
        Case 67901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOREGAPURACAOICMS", gErr, objTipoRegApuracao.iCodigo)
                
        Case 67902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOREGAPURACAOICMS", gErr, objTipoRegApuracao.iCodigo)

        Case 67903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 67927, 67928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSITEM", gErr)
        
        Case 67929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAO_VINCULADO_REGAPURACAOICMSITEM", gErr, objTipoRegApuracao.iCodigo)
                
        Case 70264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOICMS_NAO_CADASTRADA", gErr, objTipoRegApuracao.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150348)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoRegApuracaoICMS_Grava(ByVal objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Grava um tipo de Registro de apuração ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim iPreCadastrada As Integer

On Error GoTo Erro_TipoRegApuracaoICMS_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67904
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 67905

    'Verifica se o Tipo de Registro de apuração ICMS está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT PreCadastrado FROM TiposRegApuracaoICMS WHERE Codigo = ?", 0, iPreCadastrada, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 67906

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67907

    'Se o tipo de Registro de apuração estiver cadastrada ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no TipoRegApuraçãoICMS
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 67909
            
        'Se for um tipo pré-cadastrado, erro
        If iPreCadastrada = TIPOREGAPURACAO_PRE_CADASTRADO Then gError 67908
        
        'Atualiza os dados do Tipo de Registro de apuração ICMS
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposRegApuracaoICMS SET Descricao = ?, Secao = ?", alComando(0), objTipoRegApuracao.sDescricao, objTipoRegApuracao.iSecao)
        If lErro <> AD_SQL_SUCESSO Then gError 67910

    'Se o tipo de registro de apuração ICMS nãi estava cadastrada ==> INSERÇÃO
    Else

        'Insere novo Registro de Tipo de Apuração ICMS
        lErro = Comando_Executar(alComando(0), "INSERT INTO TiposRegApuracaoICMS (Codigo, Descricao, Secao, Lancamento) VALUES (?, ?, ?, ?)", objTipoRegApuracao.iCodigo, objTipoRegApuracao.sDescricao, objTipoRegApuracao.iSecao, TIPOREGAPURACAO_PERMITE_LANCAMENTO)
        If lErro <> AD_SQL_SUCESSO Then gError 67911

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 67912

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    TipoRegApuracaoICMS_Grava = SUCESSO

    Exit Function

Erro_TipoRegApuracaoICMS_Grava:

    TipoRegApuracaoICMS_Grava = gErr

    Select Case gErr

        Case 67904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 67906, 67907
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOICMS", gErr)

        Case 67908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOICMS_PRE_CADASTRADO", gErr, objTipoRegApuracao.iCodigo)
            
        Case 67909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOREGAPURACAOICMS", gErr, objTipoRegApuracao.iCodigo)

        Case 67910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPOREGAPURACAOICMS", gErr, objTipoRegApuracao.iCodigo)

        Case 67911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOREGAPURACAOICMS", gErr, objTipoRegApuracao.iCodigo)

        Case 67912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150349)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoRegApuracaoICMS_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próximo Tipo de Registro de apuração ICMS

Dim lErro As Long
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String
Dim sCampo As String

On Error GoTo Erro_TipoRegApuracaoICMS_Codigo_Automatico

    sCodigo = "FISC_NUM_PROX_TIPOREGAPURICMS"
    sTabelaConfig = "FISConfig"
    sTabelaObjeto = "TiposRegApuracaoICMS"
    sCampo = "Codigo"

    'Chama Config_ObterAutomatico
    lErro = CF("Config_ObterAutomatico", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
    If lErro <> SUCESSO Then gError 67914

    TipoRegApuracaoICMS_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_TipoRegApuracaoICMS_Codigo_Automatico:

    TipoRegApuracaoICMS_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 67914

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150350)

    End Select

    Exit Function

End Function

Function TipoRegApuracaoIPI_Exclui(ByVal objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Exclui um Tipo de Registro de apuração IPI

Dim lErro As Long
Dim iIndice As Integer
Dim iPreCadastrado As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TipoRegApuracaoIPI_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 79025
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 79026

    'Seleciona o Tipo de registro de apuração IPI a partir do código passado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT PreCadastrado FROM TiposRegApuracaoIPI WHERE Codigo = ?", 0, iPreCadastrado, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 79027

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79028

    'Se não encontrou o Tipo de Registro de Apuração IPI, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 79029

    'Se for um tipo Pré-cadastrado, erro
    If iPreCadastrado = TIPOREGAPURACAO_PRE_CADASTRADO Then gError 79030

    'Faz "lock" em TipoRegApuracaoIPI
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 79031

    'Verifica se existe relaciomento entre o TipoRegApuracaoIPI e RegApuracaoIPIItem
    lErro = Comando_Executar(alComando(1), "SELECT RegApuracaoIPIItem.NumIntDoc FROM RegApuracaoIPIItem WHERE RegApuracaoIPIItem.TipoReg = ?", lNumIntDoc, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 79032

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79033

    'Se encontrou relacionamento, Erro
    If lErro = AD_SQL_SUCESSO Then gError 79034

    'Exclui o Tipo de Reg. de apuração IPI do Banco de dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TiposRegApuracaoIPI", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 79035

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 79036

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoRegApuracaoIPI_Exclui = SUCESSO

    Exit Function

Erro_TipoRegApuracaoIPI_Exclui:

    TipoRegApuracaoIPI_Exclui = gErr

    Select Case gErr

        Case 79025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 79027, 79028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOIPI", gErr)

        Case 79029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOIPI_NAO_CADASTRADA", gErr, objTipoRegApuracao.iCodigo)

        Case 79030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOIPI_PRE_CADASTRADO", gErr, objTipoRegApuracao.iCodigo)

        Case 79031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOREGAPURACAOIPI", gErr, objTipoRegApuracao.iCodigo)

        Case 79032, 79033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)

        Case 79034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAO_VINCULADO_REGAPURACAOIPIITEM", gErr, objTipoRegApuracao.iCodigo)

        Case 79035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOREGAPURACAOIPI", gErr, objTipoRegApuracao.iCodigo)

        Case 79036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150351)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoRegApuracaoIPI_Grava(ByVal objTipoRegApuracao As ClassTiposRegApuracao) As Long
'Grava um tipo de Registro de apuração IPI

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim iPreCadastrada As Integer

On Error GoTo Erro_TipoRegApuracaoIPI_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 79037
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 79038

    'Verifica se o Tipo de Registro de apuração IPI está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT PreCadastrado FROM TiposRegApuracaoIPI WHERE Codigo = ?", 0, iPreCadastrada, objTipoRegApuracao.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 79039

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79040

    'Se o tipo de Registro de apuração estiver cadastrada ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Faz "lock" no TipoRegApuraçãoIPI
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 79041

        'Se for um tipo pré-cadastrado, erro
        If iPreCadastrada = TIPOREGAPURACAO_PRE_CADASTRADO Then gError 79042

        'Atualiza os dados do Tipo de Registro de apuração IPI
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposRegApuracaoIPI SET Descricao = ?, Secao = ?", alComando(0), objTipoRegApuracao.sDescricao, objTipoRegApuracao.iSecao)
        If lErro <> AD_SQL_SUCESSO Then gError 79043

    'Se o tipo de registro de apuração IPI nãi estava cadastrada ==> INSERÇÃO
    Else

        'Insere novo Registro de Tipo de Apuração IPI
        lErro = Comando_Executar(alComando(0), "INSERT INTO TiposRegApuracaoIPI (Codigo, Descricao, Secao, Lancamento) VALUES (?, ?, ?, ?)", objTipoRegApuracao.iCodigo, objTipoRegApuracao.sDescricao, objTipoRegApuracao.iSecao, TIPOREGAPURACAO_PERMITE_LANCAMENTO)
        If lErro <> AD_SQL_SUCESSO Then gError 79044

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 79045

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    TipoRegApuracaoIPI_Grava = SUCESSO

    Exit Function

Erro_TipoRegApuracaoIPI_Grava:

    TipoRegApuracaoIPI_Grava = gErr

    Select Case gErr

        Case 79037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 79038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 79039, 79040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOREGAPURACAOIPI", gErr)

        Case 79041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOREGAPURACAOIPI", gErr, objTipoRegApuracao.iCodigo)

        Case 79042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOREGAPURACAOIPI_PRE_CADASTRADO", gErr, objTipoRegApuracao.iCodigo)

        Case 79043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPOREGAPURACAOIPI", gErr, objTipoRegApuracao.iCodigo)

        Case 79044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOREGAPURACAOIPI", gErr, objTipoRegApuracao.iCodigo)

        Case 79045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150352)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoRegApuracaoIPI_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próximo Tipo de Registro de apuração IPI

Dim lErro As Long
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String
Dim sCampo As String

On Error GoTo Erro_TipoRegApuracaoIPI_Codigo_Automatico

    sCodigo = "FISC_NUM_PROX_TIPOREGAPURIPI"
    sTabelaConfig = "FISConfig"
    sTabelaObjeto = "TiposRegApuracaoIPI"
    sCampo = "Codigo"

    'Chama Config_ObterAutomatico
    lErro = CF("Config_ObterAutomatico", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
    If lErro <> SUCESSO Then gError 79046

    TipoRegApuracaoIPI_Codigo_Automatico = SUCESSO

    Exit Function

Erro_TipoRegApuracaoIPI_Codigo_Automatico:

    TipoRegApuracaoIPI_Codigo_Automatico = gErr

    Select Case gErr

        Case 79046

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150353)

    End Select

    Exit Function

End Function

Function LivroFilial_Exclui(ByVal objLivroFilial As ClassLivrosFilial) As Long
'Exclui um Livro Fiscal da Filial Empresa passada em objLivroFilial

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim dtDataFinal As Date
Dim dtData As Date
Dim lNumIntDoc As Long

On Error GoTo Erro_LivroFilial_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70188
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70189

    'Seleciona o Livro Fiscal a partir código do Livro e da Filial Empresa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT DataFinal FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", 0, dtDataFinal, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70190

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70191

    'Se não encontrou o Livro Fiscal, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70192
    
    'Se o Livro for de Apuração de ICMS
    If objLivroFilial.iCodLivro = LIVRO_APURACAO_ICMS_CODIGO Then
    
        'Verifica o Livro Fiscal está vinculado com uma Apuração ICMS
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM RegApuracaoICMS WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? AND NumIntDocLivFechado = ?", lNumIntDoc, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa, APURACAO_ICMS_ABERTA)
        If lErro <> AD_SQL_SUCESSO Then gError 70222
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70223
        
        'Se encontrou vínculo, erro
        If lErro = AD_SQL_SUCESSO Then gError 70224
    
    'Se o Livro for de Apuração de IPI
    ElseIf objLivroFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO Then
        
        'Verifica se o Livro Fiscal está vinculado com uma Apuração IPI
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM RegApuracaoIPI WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? AND NumIntDocLivFechado = ?", lNumIntDoc, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa, APURACAO_ICMS_ABERTA)
        If lErro <> AD_SQL_SUCESSO Then gError 70225
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70226
        
        'Se encontrou vínculo, erro
        If lErro = AD_SQL_SUCESSO Then gError 70227
                    
    'Se o Livro for de Registro de Inventário
    ElseIf objLivroFilial.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO Then
    
        'Verifica se o Livro Fiscal está vinculado a um Registro de Inventário gerado dentro do intervalo de datas
        lErro = Comando_Executar(alComando(3), "SELECT Data FROM RegInventario WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ?", dtData, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 70928
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70929
        
        'Se encontrou vínculo, erro
        If lErro = AD_SQL_SUCESSO Then gError 70930
    
    End If
    
    'Faz "lock" em LivrosFilial
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70194

    'Exclui o Livro Fiscal da Filial Empressa passado
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM LivrosFilial", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70195

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70196

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    LivroFilial_Exclui = SUCESSO

    Exit Function

Erro_LivroFilial_Exclui:

    LivroFilial_Exclui = gErr

    Select Case gErr

        Case 70188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70190, 70191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)

        Case 70192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
        
        Case 70194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 70195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 70196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 70222, 70223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
        
        Case 70224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_VINCULADO_REGAPURACAOICMS", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
        
        Case 70225, 70226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
        
        Case 70227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_VINCULADO_REGAPURACAOIPI", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
                                                                    
        Case 70928, 70929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)
        
        Case 70930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_VINCULADO_REGINVENTARIO", gErr, dtData)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150354)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function LivroFilial_Grava(ByVal objLivroFilial As ClassLivrosFilial) As Long
'Grava um Livro Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim dtDataFinal As Date
Dim dtDataInicial As Date
Dim objLivroFechado As New ClassLivrosFechados
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataInv As Date

On Error GoTo Erro_LivroFilial_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70197
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70198
    
    'Lê dados da último Livro Fiscal Fechado de Código e Filial Empresa passados
    objLivroFechado.iCodLivro = objLivroFilial.iCodLivro
    objLivroFechado.iFilialEmpresa = objLivroFilial.iFilialEmpresa
    lErro = CF("LivrosFechados_Le_UltimaData", objLivroFechado)
    If lErro <> SUCESSO And lErro <> 70231 Then gError 70215
    
    'Se encontrou o Livro Fechado
    If lErro = SUCESSO Then
    
        'Verifica se a Data Inicial de Livros Filial é menor ou igual a Data Final do último Livro Fiscal já Fechado
        If objLivroFilial.dtDataInicial <= objLivroFechado.dtDataFinal Then gError 70201
        
    End If
    
    'Verifica se o Livro Filial está cadastrado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT DataFinal, DataInicial FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", 0, dtDataFinal, dtDataInicial, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70199

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70200
    
    'Se o Livro já estiver cadastrado ==> ALTERAÇÃO
    If lErro = AD_SQL_SUCESSO Then
               
        'Se o Livro for de Apuração de ICMS
        If objLivroFilial.iCodLivro = LIVRO_APURACAO_ICMS_CODIGO Then
        
            'Verifica o Livro Fiscal está vinculado com uma Apuração ICMS
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM RegApuracaoICMS WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? AND NumIntDocLivFechado = ?", lNumIntDoc, dtDataInicial, dtDataFinal, objLivroFilial.iFilialEmpresa, APURACAO_ICMS_ABERTA)
            If lErro <> AD_SQL_SUCESSO Then gError 70216
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70217
            
            'Se encontrou vínculo
            If lErro = AD_SQL_SUCESSO Then
                
                'Verifica se as datas são diferentes
                If objLivroFilial.dtDataInicial <> dtDataInicial Or objLivroFilial.dtDataFinal <> dtDataFinal Then
                    'Avisa que só será possivel Alterar o campo Livro e Folha Inicial pois o livro já está vinculado com uma Apuração
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LIVRO_APURACAO_ICMS_ALTERAR_LIVRO_FOLHA")
                    
                    'Se não deseja contatinuar --> Sai por erro
                    If vbMsgRes = vbNo Then
                        gError 70218
                    ElseIf vbMsgRes = vbYes Then
                        'Se não mantém as datas
                        objLivroFilial.dtDataInicial = dtDataInicial
                        objLivroFilial.dtDataFinal = dtDataFinal
                    End If
                End If
            End If
            
        'Se o Livro for de Apuração de IPI
        ElseIf objLivroFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO Then
            
            'Verifica se o Livro Fiscal está vinculado com uma Apuração IPI
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM RegApuracaoIPI WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ? AND NumIntDocLivFechado = ?", lNumIntDoc, dtDataInicial, dtDataFinal, objLivroFilial.iFilialEmpresa, APURACAO_ICMS_ABERTA)
            If lErro <> AD_SQL_SUCESSO Then gError 70219
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70220
            
            'Se encontrou vínculo
            If lErro = AD_SQL_SUCESSO Then
                'Verifica se as datas são diferentes
                If objLivroFilial.dtDataInicial <> dtDataInicial Or objLivroFilial.dtDataFinal <> dtDataFinal Then
                    'Avisa que só será possivel Alterar o campo Livro e Folha Inicial pois o livro já está vinculado com uma Apuração
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LIVRO_APURACAO_IPI_ALTERAR_LIVRO_FOLHA")
                    
                    'Se não deseja continuar sai por Erro
                    If vbMsgRes = vbNo Then
                        gError 70221
                    ElseIf vbMsgRes = vbYes Then
                        'Se não mantém as datas
                        objLivroFilial.dtDataInicial = dtDataInicial
                        objLivroFilial.dtDataFinal = dtDataFinal
                    End If
                End If
            End If
                        
        'Se o Livro for de Registro de Inventário
        ElseIf objLivroFilial.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO Then
            
            'Verifica se o Livro Fiscal está vinculado com uma Registro de Inventário
            lErro = Comando_Executar(alComando(2), "SELECT Data FROM RegInventario WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? AND NumIntDocLivFechado = ?", dtDataInv, dtDataInicial, dtDataFinal, objLivroFilial.iFilialEmpresa, LIVRO_ABERTO)
            If lErro <> AD_SQL_SUCESSO Then gError 70931
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70932
            
            'Se encontrou vínculo
            If lErro = AD_SQL_SUCESSO Then
                'Verifica se as datas são diferentes
                If objLivroFilial.dtDataInicial <> dtDataInicial Or objLivroFilial.dtDataFinal <> dtDataFinal Then
                    'Avisa que só será possivel Alterar o campo Livro e Folha Inicial pois o livro já está vinculado com uma Apuração
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LIVRO_REG_INVENTARIO_ALTERAR_LIVRO_FOLHA")
                    
                    'Se não deseja continuar sai por Erro
                    If vbMsgRes = vbNo Then
                        gError 70933
                    ElseIf vbMsgRes = vbYes Then
                        'Se não mantém as datas
                        objLivroFilial.dtDataInicial = dtDataInicial
                        objLivroFilial.dtDataFinal = dtDataFinal
                    End If
                End If
            End If
        
        End If
        
        'Faz "lock" em LivrosFilial
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70202

        'Atualiza dados do Livro da Filial
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE LivrosFilial SET Imprime = ?, NumeroProxLivro = ?, NumeroProxFolha = ?, Periodicidade = ?, DataInicial = ?, DataFinal = ?, ImpressoEm = ?", alComando(0), objLivroFilial.iImprime, objLivroFilial.iNumeroProxLivro, objLivroFilial.iNumeroProxFolha, objLivroFilial.iPeriodicidade, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.dtImpressoEm)
        If lErro <> AD_SQL_SUCESSO Then gError 70203

    'Se o Item de registro de apuração ICMS nãi estava cadastrada ==> INSERÇÃO
    Else
    
        'Insere novo Item de apuração ICMS
        lErro = Comando_Executar(alComando(0), "INSERT INTO LivrosFilial (CodLivro, FilialEmpresa, Imprime, NumeroProxLivro, NumeroProxFolha, Periodicidade, DataInicial, DataFinal, ImpressoEm) VALUES (?,?,?,?,?,?,?,?,?)", objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.iImprime, objLivroFilial.iNumeroProxLivro, objLivroFilial.iNumeroProxFolha, objLivroFilial.iPeriodicidade, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.dtImpressoEm)
        If lErro <> AD_SQL_SUCESSO Then gError 70204

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70205

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LivroFilial_Grava = SUCESSO

    Exit Function

Erro_LivroFilial_Grava:

    LivroFilial_Grava = gErr

    Select Case gErr

        Case 70197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70199, 70200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)

        Case 70201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINALLIVRO", gErr, objLivroFilial.dtDataInicial, objLivroFechado.dtDataFinal)

        Case 70202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 70203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 70204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 70205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 70215
        
        Case 70216, 70217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
                
        Case 70218, 70221, 70933
                        
        Case 70219, 70220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
        
        Case 70931, 70932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150355)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function EdicaoRegES_Grava(ByVal objLivRegES As ClassLivRegES) As Long
'Grava as Linhas do Livro e Altera os Itens da Nota Fiscal

Dim lErro As Long
Dim lTransacao As Long, objNFiscal As New ClassNFiscal

On Error GoTo Erro_EdicaoRegES_Grava
        
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 69475
    
    If gobjFIS.iRegravaNF <> 0 And objLivRegES.iTipoNumIntDocOrigem = 0 And objLivRegES.lNumIntDocOrigem <> 0 Then
    
        objNFiscal.lNumIntDoc = objLivRegES.lNumIntDocOrigem
    
        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
            
        'Se não encontrou a NFiscal --> erro
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Lê os Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        lErro = CF("NotaFiscal_Exclui_Fis", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    'Altera o Status e o Estado do Livro Fiscal
    lErro = CF("EdicaoRegES_Grava_AlteraStatus_Trans", objLivRegES)
    If lErro <> SUCESSO Then gError 69509
    
    'Grava as Linhas do Livro
    lErro = CF("EdicaoRegES_Grava_Linha_Trans", objLivRegES)
    If lErro <> SUCESSO Then gError 69507
    
    'Altera os Itens da Nota fiscal
    lErro = CF("EdicaoRegES_Grava_Item_Trans", objLivRegES)
    If lErro <> SUCESSO Then gError 69508
    
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 69488
    
    EdicaoRegES_Grava = SUCESSO
    
    Exit Function
    
Erro_EdicaoRegES_Grava:

    EdicaoRegES_Grava = gErr
    
    Select Case gErr
        
        Case 69475
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                    
        Case 69488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 69507, 69508, 69509, ERRO_SEM_MENSAGEM
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150356)

    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function EdicaoRegES_Grava_Linha_Trans(ByVal objLivRegES As ClassLivRegES) As Long
'Grava as Linhas do Livro

Dim lErro As Long
Dim lNumIntDoc As Long
Dim alComando(3) As Long
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim iIndice As Integer

On Error GoTo Erro_EdicaoRegES_Grava_Linha_Trans
        
    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69476
    Next
    
    'Verica se não tem nenhum item Gravado
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM LivRegESLinha WHERE NumIntDocRegES = ?", 0, lNumIntDoc, objLivRegES.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69477
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69478
        
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 69479
        
    'Se tiver
    Do While lErro = AD_SQL_SUCESSO
    
        'Apaga os Itens da Tabela de LivRegESLinha
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM LivRegESLinha", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 69480
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69490
        
    Loop
     
    'Insere os novos lançamentos
    For Each objLivRegESLinha In objLivRegES.colLinhas
        
        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGESLINHA", lNumIntDoc)
        If lErro <> SUCESSO Then gError 69491
    
        objLivRegESLinha.lNumIntDoc = lNumIntDoc
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO LivRegESLinha(NumIntDoc, NumIntDocRegES, NaturezaOp, ValorTotal, ValorICMSBase, ValorICMS, ValorICMSIsentoNaoTrib, ValorICMSOutras, ValorICMSSubstBase, ValorICMSSubstRet, AliquotaICMS, ValorDespAcess, ValorIPI, AliquotaIPI, ValorIPIBase, ValorIPIIsentoNaoTrib, ValorIPIOutras, ClassifContabil, ValorContabil, ObservacaoLivFisc, CodigoICMS, CodigoIPI) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objLivRegESLinha.lNumIntDoc, objLivRegESLinha.lNumIntDocRegES, objLivRegESLinha.sNaturezaOp, Arredonda_Moeda(objLivRegESLinha.dValorTotal), _
            Arredonda_Moeda(objLivRegESLinha.dValorICMSBase), Arredonda_Moeda(objLivRegESLinha.dValorICMS), Arredonda_Moeda(objLivRegESLinha.dValorICMSIsentoNaoTrib), Arredonda_Moeda(objLivRegESLinha.dValorICMSOutras), Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstBase), Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstRet), objLivRegESLinha.dAliquotaICMS, Arredonda_Moeda(objLivRegESLinha.dValorDespAcess), Arredonda_Moeda(objLivRegESLinha.dValorIPI), objLivRegESLinha.dAliquotaIPI, Arredonda_Moeda(objLivRegESLinha.dValorIPIBase), Arredonda_Moeda(objLivRegESLinha.dValorIPIIsentoNaoTrib) _
        , Arredonda_Moeda(objLivRegESLinha.dValorIPIOutras), objLivRegESLinha.sClassifContabil, Arredonda_Moeda(objLivRegESLinha.dValorContabil), objLivRegESLinha.sObservacaoLivFisc, objLivRegESLinha.iCodigoICMS, objLivRegESLinha.iCodigoIPI)
        If lErro <> AD_SQL_SUCESSO Then gError 69481
    
    Next
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    EdicaoRegES_Grava_Linha_Trans = SUCESSO
    
    Exit Function
    
Erro_EdicaoRegES_Grava_Linha_Trans:

    EdicaoRegES_Grava_Linha_Trans = gErr
    
    Select Case gErr
        
        Case 69476
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69477, 69478, 69490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESLINHA", gErr)
        
        Case 69479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGESLINHA", gErr)
        
        Case 69480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESLINHA", gErr)
         
        Case 69481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LIVREGESLINHA", gErr)
        
        Case 69491
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150357)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function EdicaoRegES_Grava_Item_Trans(ByVal objLivRegES As ClassLivRegES) As Long
'Altera os Itens da Nota Fiscal

Dim lErro As Long
Dim lNumIntDoc As Long, lNumIntDocItem As Long
Dim alComando(6) As Long
Dim iIndice As Integer
Dim iNumItem As Integer
Dim objLivRegESItem As New ClassLivRegESItemNF
Dim iEncontrou As Integer, iTipo As Integer

On Error GoTo Erro_EdicaoRegES_Grava_Item_Trans
        
    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69510
    Next
    
    If objLivRegES.iTipoNumIntDocOrigem = TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL Then
    
        lErro = Comando_Executar(alComando(4), "SELECT T.Tipo FROM NFiscal N, TiposDocInfo T WHERE N.TipoNFiscal = T.Codigo AND N.NumIntDoc = ?", iTipo, objLivRegES.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 69482
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69483
    
    End If
    
    'Seleciona os Itens da Tabela de LivRegESItemNF
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumItem FROM LivRegESItemNF WHERE NumIntDocRegES = ? ORDER BY NumItem", 0, iNumItem, objLivRegES.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69482
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69483
    
    'Faz o Lock exclusivo no Recebimento
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 69484
    
    'Se encontrou
    Do While lErro = AD_SQL_SUCESSO
        
        iEncontrou = 0
            
        'Procura na Coleção
        For Each objLivRegESItem In objLivRegES.ColItensNF
            
            'Se encontrou o Item
            If objLivRegESItem.iNumItem = iNumItem Then
                
                iEncontrou = 1
                
                'Atualiza no BD os seguintes campos: CFOP, ValorDescontoDespAcess, ValorICMSBase, ValorIPI, AliquotaICMS
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE LivRegESItemNF SET CFOP = ?, ValorDescontoDespAcess = ?, ValorICMSBase = ?, ValorIPI = ?, AliquotaICMS = ?, TipoTribICMS = ?, TipoTribIPI = ?, ValorICMS = ?, ValorIPIBase = ?, AliquotaIPI = ?, RedBaseICMS = ?, RedBaseIPI = ?, TipoTributacao = ?, AliquotaSubst = ?, ValorSubst = ?, ValorICMSSubstBase = ?", alComando(1), objLivRegESItem.sCFOP, objLivRegESItem.dValorDescontoDespAcess, objLivRegESItem.dValorICMSBase, objLivRegESItem.dValorIPI, objLivRegESItem.dAliquotaICMS, objLivRegESItem.iTipoTribICMS, objLivRegESItem.iTipoTribIPI, objLivRegESItem.dValorICMS, objLivRegESItem.dValorIPIBase, objLivRegESItem.dAliquotaIPI, objLivRegESItem.dRedBaseICMS, objLivRegESItem.dRedBaseIPI, objLivRegESItem.iTipoTributacao, objLivRegESItem.dAliquotaSubst, objLivRegESItem.dValorSubst, objLivRegESItem.dValorICMSSubstBase)
                If lErro <> AD_SQL_SUCESSO Then gError 69485
                
                'se a nota é externa
                If iTipo = TIPODOCINFO_TIPO_NFEXT And iNumItem < 991 Then
                
                    'alterar cfop e outros dados cuja perspectiva é interna
                    lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDocItem FROM TributacaoDocItem WHERE TipoDoc = ? AND NumIntDoc = ? AND Item = ?", 0, lNumIntDocItem, TIPODOC_TRIB_NF, objLivRegES.lNumIntDocOrigem, iNumItem)
                    If lErro <> AD_SQL_SUCESSO Then gError 69482
                    
                    lErro = Comando_BuscarProximo(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69483
                    
                    'Faz o Lock exclusivo no Recebimento
                    lErro = Comando_LockExclusive(alComando(5))
                    If lErro <> AD_SQL_SUCESSO Then gError 69484

                    lErro = Comando_ExecutarPos(alComando(6), "UPDATE TributacaoDocItem SET NaturezaOp = ?", alComando(5), objLivRegESItem.sCFOP)
                    If lErro <> AD_SQL_SUCESSO Then gError 69485
                
                End If
                
                Exit For
                
            End If
            
        Next
        
        If iEncontrou = 0 Then
            'Se não Encontrou --> Exclui
            lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM LivRegESItemNF", alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 69486
        End If
        
        'Pega o Próximo Item
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69487
        
    Loop
        
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    EdicaoRegES_Grava_Item_Trans = SUCESSO
    
    Exit Function
    
Erro_EdicaoRegES_Grava_Item_Trans:

    EdicaoRegES_Grava_Item_Trans = gErr
    
    Select Case gErr
            
        Case 69482, 69483, 69487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESITEM", gErr)
        
        Case 69484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGESITEM", gErr)
        
        Case 69485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGESITEM", gErr)
        
        Case 69486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESITEM", gErr)
        
        Case 69510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150358)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function EdicaoRegES_Grava_AlteraStatus_Trans(ByVal objLivRegES As ClassLivRegES) As Long
'Altera o Status do Livro Fiscal

Dim lErro As Long
Dim alComando(2) As Long
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim lNumIntLivroFechado As Long

On Error GoTo Erro_EdicaoRegES_Grava_AlteraStatus_Trans
        
    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69511
    Next
    
    'Seleciona os LivRegES
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, NumIntDocLivFechado FROM LivRegES WHERE NumIntDoc = ?", 0, lNumIntDoc, lNumIntLivroFechado, objLivRegES.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69512
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69513
    
    If lNumIntLivroFechado <> 0 Then gError 69661
    
    'Faz o Lock exclusivo no Livro
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 69514
    
    'Se encontrou
    If lErro = AD_SQL_SUCESSO Then
        
        'Atualiza o Status do Livro
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE LivRegES SET StatusLivro = ?, UF = ?, PISValor = ?, COFINSValor = ?, Data = ?, DataEmissao = ?", alComando(1), objLivRegES.iStatusLivro, objLivRegES.sUF, objLivRegES.dPISValor, objLivRegES.dCOFINSValor, objLivRegES.dtData, objLivRegES.dtDataEmissao)
        If lErro <> AD_SQL_SUCESSO Then gError 69515
                    
    End If
            
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    EdicaoRegES_Grava_AlteraStatus_Trans = SUCESSO
    
    Exit Function
    
Erro_EdicaoRegES_Grava_AlteraStatus_Trans:

    EdicaoRegES_Grava_AlteraStatus_Trans = gErr
    
    Select Case gErr
            
        Case 69511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69512, 69513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case 69514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
        
        Case 69515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
        
        Case 69661
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROREGES_JA_FECHADO", gErr, objLivRegES.sSerie, objLivRegES.lNumNotaFiscal)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150359)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function RegInventario_Exclui(ByVal objRegInventario As ClassRegInventario) As Long
'Exclui um Registro de Inventário

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 6) As Long
Dim lTransacao As Long
Dim lNumIntDocLivFechado As Long
Dim iAlmoxarifado As Integer
Dim lNumIntDocReq As Long

On Error GoTo Erro_RegInventario_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70313
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70314

    'Seleciona o Registro de Inventário a partir do Produto, Data e FilialEmpresa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegInventario WHERE Produto = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70315

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70316

    'Se o Livro Fiscal relacionado ao Registro de inventário já foi fechado, erro
    If Not (lNumIntDocLivFechado = LIVRO_ABERTO) Then gError 70317

    'Se o Produto possui Almoxarifado associado
    If objRegInventario.iAlmoxarifado <> 0 Then

        'Lê o Registro de Inventário da tabela RegInventarioAlmox
        lErro = Comando_ExecutarPos(alComando(2), "SELECT Almoxarifado FROM RegInventarioAlmox WHERE Produto = ? AND Data = ? AND FilialEmpresa = ? AND Almoxarifado = ?", 0, iAlmoxarifado, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa, objRegInventario.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 70934

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70935

        'Se não encontrou RegInventarioAlmox, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 70936

        'Faz "Lock" em RegInventarioAlmox
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 70937

        'Exclui RegInventarioAlmox
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM RegInventarioAlmox", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 70938

    End If

    'Verifica se existem outros registros em RegInventarioAlmox para o Produto e a Data em questão
    lErro = Comando_Executar(alComando(4), "SELECT Almoxarifado FROM RegInventarioAlmox WHERE Produto = ? AND Data = ? AND FilialEmpresa = ?", iAlmoxarifado, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70941

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70942

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then

        'Faz "lock" em RegInventario
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70320

        'Exclui Registro de Inventário do Banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RegInventario", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70321

    End If

    'Lê Registro de Inventário em RegInventarioAlmox
    lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDocMovEstReq FROM RegInventarioReq WHERE Produto = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocReq, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70941

    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70942

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Faz "lock" em RegInventarioAlmox
        lErro = Comando_LockExclusive(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 70320

        'Exclui Registro de Inventario de RegInventarioAlmox
        lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM RegInventarioReq", alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 70321
        
        'Busca próximo Registro de inventário
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70942

    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70322

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegInventario_Exclui = SUCESSO

    Exit Function

Erro_RegInventario_Exclui:

    RegInventario_Exclui = gErr

    Select Case gErr

        Case 70313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70315, 70316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)

        Case 70317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_FECHADO", gErr, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGINVENTARIO", gErr, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGINVENTARIO", gErr, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 70934, 70935
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIOALMOX", gErr)

        Case 70936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIOALMOX_NAO_CADASTRADO", gErr, objRegInventario.sProduto, objRegInventario.iAlmoxarifado, objRegInventario.dtData)

        Case 70937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGINVENTARIOALMOX", gErr, objRegInventario.sProduto, objRegInventario.iAlmoxarifado, objRegInventario.dtData)

        Case 70938
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGINVENTARIOALMOX", gErr, objRegInventario.sProduto, objRegInventario.iAlmoxarifado, objRegInventario.dtData)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150360)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegInventario_Atualiza(ByVal objRegInventario As ClassRegInventario) As Long
'Atualiza dados de RegInventário

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim lNumIntDocLivFechado As Long

On Error GoTo Erro_RegInventario_Atualiza

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70323
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70324

    'Verifica se o Registro de inventário está cadastrado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegInventario WHERE Produto = ? AND Data = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70325

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70326

    'Se não encontrou o Registro de Inventário, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 70327

    'Se o Livro Fiscal relacionado ao Registro de inventário já foi fechado, erro
    If Not (lNumIntDocLivFechado = LIVRO_ABERTO) Then gError 70328

    'Faz "lock" no RegInventario
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 70329

    'Atualiza os dados do Registro de Inventário
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegInventario SET Observacoes = ?", alComando(0), objRegInventario.sObservacoes)
    If lErro <> AD_SQL_SUCESSO Then gError 70330

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70331

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegInventario_Atualiza = SUCESSO

    Exit Function

Erro_RegInventario_Atualiza:

    RegInventario_Atualiza = gErr

    Select Case gErr

        Case 70323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70325, 70326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)

        Case 70327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_NAO_CADASTRADO", gErr, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_FECHADO", gErr, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGINVENTARIO", gErr, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGINVENTARIO", gErr, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150361)

    End Select

    Call Transacao_Rollback

    'Fechamento de comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RegInventarioTodos_Exclui(ByVal objRegInventario As ClassRegInventario) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDocLivFechado As Long
Dim iAlmoxarifado As Integer
Dim lNumIntDocReq As Long

On Error GoTo Erro_RegInventarioTodos_Exclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70947
    Next

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 70948

    'Procura por Registro de Inventário com a data passada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDocLivFechado FROM RegInventario WHERE Data = ? AND FilialEmpresa = ?", 0, lNumIntDocLivFechado, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70949

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70950

    'Se o Livro já foi fechado, erro
    If Not (lNumIntDocLivFechado = LIVRO_ABERTO) Then gError 70951

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Faz "Lock" no Registro de Inventario
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70952

        'Exclui Registro de Inventário
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RegInventario", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70953

        'Faz o "UnLock" do Registro de Inventário
        lErro = Comando_Unlock(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 101080
        
        'Busca próximo Registro de Inventário
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70954

    Loop

    'Lê Registro de Inventário em RegInventarioAlmox
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Almoxarifado FROM RegInventarioAlmox WHERE Data = ? AND FilialEmpresa = ?", 0, iAlmoxarifado, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70955

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70956

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Faz "lock" em RegInventarioAlmox
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 70960

        'Exclui Registro de Inventario de RegInventarioAlmox
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM RegInventarioAlmox", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 70957

        'Faz o "UnLock" do Registro de Inventário
        lErro = Comando_Unlock(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 101081
        
        'Busca próximo Registro de inventário
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70958

    Loop
    
    'Lê Registro de Inventário em RegInventarioAlmox
    lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDocMovEstReq FROM RegInventarioReq WHERE Data = ? AND FilialEmpresa = ?", 0, lNumIntDocReq, objRegInventario.dtData, objRegInventario.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 70955

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70956

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Faz "lock" em RegInventarioAlmox
        lErro = Comando_LockExclusive(alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 70960

        'Exclui Registro de Inventario de RegInventarioAlmox
        lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM RegInventarioReq", alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 70957

        'Faz o "UnLock" do Registro de Inventário
        lErro = Comando_Unlock(alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 101081
        
        'Busca próximo Registro de inventário
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70958

    Loop

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 70959

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegInventarioTodos_Exclui = SUCESSO

    Exit Function

Erro_RegInventarioTodos_Exclui:

    RegInventarioTodos_Exclui = gErr

    Select Case gErr

        Case 70947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70948
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 70949, 70950, 70954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)

        Case 70951
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_FECHADO", gErr, objRegInventario.dtData, objRegInventario.iFilialEmpresa)

        Case 70952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGINVENTARIO1", gErr, objRegInventario.dtData)

        Case 70953
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGINVENTARIO1", gErr, objRegInventario.dtData)

        Case 70955, 70956, 70958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIOALMOX", gErr)

        Case 70957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGINVENTARIOALMOX1", gErr, objRegInventario.dtData)

        Case 70959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 70960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGINVENTARIOALMOX1", gErr, objRegInventario.dtData)

        Case 101080
            Call Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_REGINVENTARIO", gErr, objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData)
        
        Case 101081
            Call Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_REGINVENTARIOALMOX", gErr, objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iAlmoxarifado)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150362)

    End Select

    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RegInventario_Geracao_Trans(ByVal dtData As Date) As Long
'Gera Registro de Inventário para todos Produtos
'??? deveria carregar uma colecao de almoxarifados em memoria p/evitar ler varias vezes o mesmo reg. Jones 05/03/2001
'??? filialempresa deveria ser parametro passado

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim sProduto As String
Dim objRegInventario As ClassRegInventario
Dim objProduto As New ClassProduto
Dim iAlmoxPadrao As Integer, dQuantidadeNossa As Double
Dim iIndice As Integer, dValorUnitario As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemMovReq As ClassItemMovEstoque

On Error GoTo Erro_RegInventario_Geracao_Trans

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70245
    Next

    sProduto = String(STRING_PRODUTO, 0)

    '??? dados do produto deveriam ser lidos neste select em vez de Produto_Le abaixo
    'Lê todos os Produtos que tenham Data de Inventário (EstoqueProduto) menor ou igual a dtData e FilialEmpresa = giFilialEmpresa
    lErro = Comando_Executar(alComando(0), "SELECT EstoqueProduto.Produto FROM EstoqueProduto, Almoxarifado, Produtos WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.DataInicial <= ? AND Produtos.Codigo = EstoqueProduto.Produto GROUP BY EstoqueProduto.Produto ORDER BY EstoqueProduto.Produto", sProduto, giFilialEmpresa, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 70246

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70247

    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 69805

    'Enquanto Encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Lê o Produto em questão
        objProduto.sCodigo = sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 70251

        'Se não encontrou o Produto, Erro
        If lErro = 28030 Then gError 70252

        Set objRegInventario = New ClassRegInventario

        'Guarda dados lidos em objRegInventario
        objRegInventario.iFilialEmpresa = giFilialEmpresa
        objRegInventario.sProduto = sProduto
        objRegInventario.dtData = dtData
        objRegInventario.sDescricao = objProduto.sDescricao
        objRegInventario.sModelo = objProduto.sModelo
        objRegInventario.sIPICodigo = objProduto.sIPICodigo
        objRegInventario.sSiglaUMEstoque = objProduto.sSiglaUMEstoque

        objRegInventario.iNatureza = objProduto.iNatureza

        objRegInventario.sObservacoes = ""

        'Se o Produto não é estocável
        If objProduto.iControleEstoque = PRODUTO_SEM_ESTOQUE Then
            'Guarda a conta contábil de aplicação
            objRegInventario.sContaContabil = objProduto.sContaContabil
        'Se não
        Else
            'Guarda a Conta Contábil de  Estoque Produto
            
            'Lê o Almoxarifado Padrão do Produto em questão
            lErro = CF("AlmoxarifadoPadrao_Le", objRegInventario.iFilialEmpresa, sProduto, iAlmoxPadrao)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 70539
    
            'Se encontrou
            If lErro = SUCESSO And iAlmoxPadrao <> 0 Then
    
                'Lê o Almoxarifado
                objAlmoxarifado.iCodigo = iAlmoxPadrao
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 70541
    
                'Se não encontrou o Almoxarifado, erro
                If lErro = 25056 Then gError 70542
    
            End If
            
            objRegInventario.sContaContabil = objAlmoxarifado.sContaContabil
            
            If gobjFIS.iRegInvPegaCtaProd = MARCADO Then
                If Len(Trim(objProduto.sContaContabil)) > 0 Then objRegInventario.sContaContabil = objProduto.sContaContabil
            End If
        End If

        objRegInventario.lNumIntDocLivFechado = 0

        'Calcula as Quantidades e Custos na Data do Inventário
        lErro = CF("RegInventario_Obter_Quantidades", objRegInventario)
        If lErro <> SUCESSO Then gError 69887
        
'        'Calcula as Quantidades e Custos na Data do Inventário
'        lErro = CF("RegInventario_Obter_Custos", objRegInventario)
'        If lErro <> SUCESSO Then gError 69888
        
        objRegInventario.dQtdeDeTercConosco = objRegInventario.dQuantBenef3 + objRegInventario.dQuantConserto3 + objRegInventario.dQuantConsig3 + objRegInventario.dQuantDemo3 + objRegInventario.dQuantOutras3
        objRegInventario.dQtdeNossaEmTerc = objRegInventario.dQuantBenef + objRegInventario.dQuantConserto + objRegInventario.dQuantConsig + objRegInventario.dQuantDemo + objRegInventario.dQuantOutras
        dQuantidadeNossa = objRegInventario.dQuantidadeUMEstoque - objRegInventario.dQtdeDeTercConosco
                
        dValorUnitario = (objRegInventario.dValorEstoque - (objRegInventario.dValorBenef3 + objRegInventario.dValorConserto3 + objRegInventario.dValorConsig3 + objRegInventario.dValorDemo3 + objRegInventario.dValorOutras3))

        'Insere Registro em RegInventario
        lErro = Comando_Executar(alComando(2), "INSERT INTO RegInventario (FilialEmpresa, Produto, Data, NumIntDocLivFechado, Descricao, Modelo, IPICodigo, SiglaUMEstoque, QuantidadeUMEstoque, ValorUnitario, Natureza, QtdeNossaEmTerc, QtdeDeTercConosco, Observacoes, ContaContabil, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantBenef, QuantBenef3, CustoConsig3, CustoConsig, CustoDemo3, CustoDemo, CustoConserto3, CustoConserto, CustoOutras3, " & _
            "CustoOutras, CustoBenef, CustoBenef3, QuantEmProcesso,QuantEmProcessoBenef3,QuantEmProcessoOutros3, ValorEmProcesso,ValorEmProcessoBenef3,ValorEmProcessoOutros3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.lNumIntDocLivFechado, objRegInventario.sDescricao, objRegInventario.sModelo, objRegInventario.sIPICodigo, objRegInventario.sSiglaUMEstoque, _
            dQuantidadeNossa, dValorUnitario, objRegInventario.iNatureza, objRegInventario.dQtdeNossaEmTerc, objRegInventario.dQtdeDeTercConosco, objRegInventario.sObservacoes, objRegInventario.sContaContabil, objRegInventario.dQuantConsig3, objRegInventario.dQuantConsig, objRegInventario.dQuantDemo3, objRegInventario.dQuantDemo, objRegInventario.dQuantConserto3, objRegInventario.dQuantConserto, objRegInventario.dQuantOutras3, objRegInventario.dQuantOutras, objRegInventario.dQuantBenef, objRegInventario.dQuantBenef3, _
            objRegInventario.dValorConsig3, objRegInventario.dValorConsig, objRegInventario.dValorDemo3, objRegInventario.dValorDemo, objRegInventario.dValorConserto3, objRegInventario.dValorConserto, objRegInventario.dValorOutras3, objRegInventario.dValorOutras, objRegInventario.dValorBenef, objRegInventario.dValorBenef3, objRegInventario.dQuantEmProcesso, objRegInventario.dQuantEmProcessoBenef3, objRegInventario.dQuantEmProcessoOutros3, objRegInventario.dValorEmProcesso, objRegInventario.dValorEmProcessoBenef3, objRegInventario.dValorEmProcessoOutros3)
        If lErro <> AD_SQL_SUCESSO Then gError 70253

        For Each objItemMovReq In objRegInventario.colRequisicoes
            'Insere Registro em RegInventarioReq
            lErro = Comando_Executar(alComando(3), "INSERT INTO RegInventarioReq (FilialEmpresa, Produto, Data, CodigoReq, NumIntDocMovEstReq, DataReq, Quantidade, Valor, UM, TipoMov)  VALUES (?,?,?,?,?,?,?,?,?,?)", _
            objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData, objItemMovReq.lCodigo, objItemMovReq.lNumIntDoc, objItemMovReq.dtData, objItemMovReq.dQuantidade, objItemMovReq.dCusto, objItemMovReq.sSiglaUM, objItemMovReq.iTipoMov)
            If lErro <> AD_SQL_SUCESSO Then gError 70253
        Next
        
        'Busca Próximo Produto
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70248

    Loop

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegInventario_Geracao_Trans = SUCESSO

    Exit Function

Erro_RegInventario_Geracao_Trans:

    RegInventario_Geracao_Trans = gErr

    Select Case gErr

        Case 70245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70246, 70247, 70248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO3", gErr)

        Case 70249, 70250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST1", gErr, giFilialEmpresa, sProduto)

        Case 70251, 70539, 70541, 69887, 69888

        Case 70252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 70253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGINVENTARIO", gErr, objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData)

        Case 70524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDDIAEST_NAO_CADASTRADO", gErr, sProduto)

        Case 70542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)

        Case 69805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REG_INVENTARIO_SEM_PRODUTO_DATA", gErr, objRegInventario.dtData)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150363)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RegInventarioAlmox_Geracao_Trans(ByVal dtData As Date)
'Gera Registro de Inventário para todos Produtos agrupados por Almoxarifado

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iAlmoxarifado As Integer
Dim sProduto As String
Dim objProduto As New ClassProduto
Dim iIndice As Integer, sProdutoAnterior As String
Dim objRegInventario As ClassRegInventario

On Error GoTo Erro_RegInventarioAlmox_Geracao_Trans

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70254
    Next

    sProduto = String(STRING_PRODUTO, 0)
    sProdutoAnterior = ""
    
    'Lê os Produtos de Estoque Produto que Tenha DataInicial < dtData e que seja de Almoxarifado de Filial Empresa = giFilialEmpresa
    lErro = Comando_Executar(alComando(0), "SELECT EstoqueProduto.Almoxarifado, EstoqueProduto.Produto FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ? AND DataInicial <= ? ORDER BY EstoqueProduto.Produto", iAlmoxarifado, sProduto, giFilialEmpresa, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 70255

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70256

    'Enquanto Encontrar
    Do While lErro = AD_SQL_SUCESSO

        If sProduto <> sProdutoAnterior Then
        
            'Lê o Produto em questão
            objProduto.sCodigo = sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 70260
    
            'Se não encontrou o Produto, Erro
            If lErro = 28030 Then gError 70261
    
            sProdutoAnterior = sProduto
            
        End If
        
        Set objRegInventario = New ClassRegInventario

        'Guarda dados lidos em objRegInventario
        objRegInventario.iFilialEmpresa = giFilialEmpresa
        objRegInventario.sProduto = sProduto
        objRegInventario.dtData = dtData

        objRegInventario.iNatureza = objProduto.iNatureza
        objRegInventario.iAlmoxarifado = iAlmoxarifado
        
        objRegInventario.lNumIntDocLivFechado = 0

        'Calcula as Quantidades e Custos na Data do Inventário
        lErro = CF("RegInventario_Obter_Quantidades_Almox", objRegInventario)
        If lErro <> SUCESSO Then gError 69889
        
        objRegInventario.dQtdeDeTercConosco = objRegInventario.dQuantConsig3 + objRegInventario.dQuantDemo3 + objRegInventario.dQuantConserto3 + objRegInventario.dQuantOutras3 + objRegInventario.dQuantBenef3
        
        'Insere Registro em RegInventario
        lErro = Comando_Executar(alComando(2), "INSERT INTO RegInventarioAlmox (FilialEmpresa, Produto, Data, Almoxarifado, QuantidadeUMEstoque, QtdeDeTercConosco, QuantConsig3, QuantDemo3, QuantConserto3, QuantOutras3, QuantBenef3) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData, objRegInventario.iAlmoxarifado, objRegInventario.dQuantidadeUMEstoque, objRegInventario.dQtdeDeTercConosco, objRegInventario.dQuantConsig3, objRegInventario.dQuantDemo3, objRegInventario.dQuantConserto3, objRegInventario.dQuantOutras3, objRegInventario.dQuantBenef3)
        If lErro <> AD_SQL_SUCESSO Then gError 70262

        'Busca próximo Almoxarifado
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70257

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RegInventarioAlmox_Geracao_Trans = SUCESSO

    Exit Function

Erro_RegInventarioAlmox_Geracao_Trans:

    RegInventarioAlmox_Geracao_Trans = gErr

    Select Case gErr

        Case 70254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 70255, 70256, 70257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO3", gErr)

        Case 70258, 70259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, iAlmoxarifado, sProduto)

        Case 70260, 69889

        Case 70261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 70262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGINVENTARIO", gErr, objRegInventario.iFilialEmpresa, objRegInventario.sProduto, objRegInventario.dtData)

        Case 70523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDDIAESTALM_NAO_CADASTRADO", gErr, sProduto, iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150364)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function Livro_Fechado_Insere(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Insere o registro no livro Fechado com os valores que veio da Tabela de Livro Filial passado no objLivroFechado

Dim lNumIntDoc As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Livro_Fechado_Insere

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69618
    Next

    'Gera número automático para o Item de apuração ICMS
    lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVRO_FECHADO", lNumIntDoc)
    If lErro <> SUCESSO Then gError 69619

    objLivroFechado.lNumIntDoc = lNumIntDoc

    'Insere novo Item de apuração ICMS
    lErro = Comando_Executar(alComando(0), "INSERT INTO LivrosFechados (NumIntDoc, CodLivro, FilialEmpresa, NumeroLivro, DataInicial, DataFinal, DataImpressao, FolhaInicial) VALUES (?,?,?,?,?,?,?,?)", objLivroFechado.lNumIntDoc, objLivroFechado.iCodLivro, objLivroFechado.iFilialEmpresa, objLivroFechado.iNumeroLivro, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.dtDataImpressao, objLivroFechado.iFolhaInicial)
    If lErro <> AD_SQL_SUCESSO Then gError 69620

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Livro_Fechado_Insere = SUCESSO

    Exit Function

Erro_Livro_Fechado_Insere:

    Livro_Fechado_Insere = gErr

    Select Case gErr

        Case 69618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69619
        
        Case 69620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150365)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Livro_Filial_Atualiza_Trans(ByVal objLivroFilial As ClassLivrosFilial, ByVal objLivroFilialLido As ClassLivrosFilial) As Long
'Atualiza os Valores de DataInicial, DataFinal, NumeroProxLivro, NumeroProxFolha, ImpressoEm
'Depois retorna o Valor que está no BD para poder inserir no Livro Fechado

Dim lErro As Long
Dim tLivroFilial As typeLivrosFilial
Dim iDiferenca As Integer
Dim alComando(0 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Livro_Filial_Atualiza_Trans

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69621
    Next

    'Lê dados do Livro Filial da FilialEmpresa e do Livro passados em objLivroFilial
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodLivro, FilialEmpresa, NumeroProxLivro, NumeroProxFolha, Imprime, Periodicidade, DataInicial, DataFinal, ImpressoEm FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", 0, _
    tLivroFilial.iCodLivro, tLivroFilial.iFilialEmpresa, tLivroFilial.iNumeroProxLivro, tLivroFilial.iNumeroProxFolha, tLivroFilial.iImprime, tLivroFilial.iPeriodicidade, tLivroFilial.dtDataInicial, tLivroFilial.dtDataFinal, tLivroFilial.dtImpressoEm, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69622

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69623

    'Se não encontrou o livro, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 69624

    'Retorna o Valor Lido do BD para Atualizar a Tabela de Livros Fechados
    objLivroFilialLido.iCodLivro = tLivroFilial.iCodLivro
    objLivroFilialLido.iFilialEmpresa = tLivroFilial.iFilialEmpresa
    objLivroFilialLido.iNumeroProxLivro = tLivroFilial.iNumeroProxLivro
    objLivroFilialLido.iNumeroProxFolha = tLivroFilial.iNumeroProxFolha
    objLivroFilialLido.iImprime = tLivroFilial.iImprime
    objLivroFilialLido.iPeriodicidade = tLivroFilial.iPeriodicidade
    objLivroFilialLido.dtDataInicial = tLivroFilial.dtDataInicial
    objLivroFilialLido.dtDataFinal = tLivroFilial.dtDataFinal
    objLivroFilialLido.dtImpressoEm = tLivroFilial.dtImpressoEm

    'Faz "lock" em LivrosFilial
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 69625

    'Calculo da Periodicidade para o periodo Seguinte
    If tLivroFilial.iPeriodicidade = PERIODICIDADE_LIVRE Then
        iDiferenca = objLivroFilial.dtDataInicial - objLivroFilial.dtDataFinal
        objLivroFilial.dtDataInicial = objLivroFilial.dtDataFinal + 1
        objLivroFilial.dtDataFinal = objLivroFilial.dtDataInicial + iDiferenca
    Else
        objLivroFilial.dtDataInicial = objLivroFilial.dtDataFinal + 1
        objLivroFilial.dtDataFinal = Calcula_Periodicidade(tLivroFilial.iPeriodicidade, objLivroFilial.dtDataInicial)
    End If

    objLivroFilial.iImprime = IMPRIME_LIVROFISCAL
    objLivroFilial.dtImpressoEm = DATA_NULA

    'Atualiza dados do Livro da Filial
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE LivrosFilial SET Imprime = ?, NumeroProxLivro = ?, NumeroProxFolha = ?, DataInicial = ?, DataFinal = ?, ImpressoEm = ?", alComando(0), objLivroFilial.iImprime, objLivroFilial.iNumeroProxLivro, objLivroFilial.iNumeroProxFolha, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.dtImpressoEm)
    If lErro <> AD_SQL_SUCESSO Then gError 69626

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Livro_Filial_Atualiza_Trans = SUCESSO

    Exit Function

Erro_Livro_Filial_Atualiza_Trans:

    Livro_Filial_Atualiza_Trans = gErr

    Select Case gErr

        Case 69621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69622, 69623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
        
        Case 69624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFILIAL_NAO_CADASTRADO", gErr)
        
        Case 69625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFILIAL", gErr)
        
        Case 69626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVROSFILIAL", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150366)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegEntrada(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Entrada

Dim lErro As Long
Dim lNumIntDoc As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Fechamento_Livro_RegEntrada

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69627
    Next

    'Lê os Registros de Entrada colocando o NumIntDoc do livro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM LivRegES WHERE Data >= ? AND Data <= ? AND Tipo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, TIPO_REGES_ENTRADA, objLivroFechado.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69628

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69629

    Do While lErro = AD_SQL_SUCESSO

        'Faz "lock" em LivrosFilial
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 69632
    
        'Atualiza o NumIntDocLivFechado dos Registros de Saida
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE LIvRegES SET NumIntDocLivFechado = ?", alComando(0), objLivroFechado.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69630

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69631

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Fechamento_Livro_RegEntrada = SUCESSO
    
    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegEntrada:

    Rotina_Fechamento_Livro_RegEntrada = gErr

    Select Case gErr

        Case 69627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69628, 69629, 69631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case 69630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
                
        Case 69632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150367)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Fechamento_Livro(ByVal objLivrosFilial As ClassLivrosFilial) As Long
'Esta Rotina faz o Fechamento do Livro passado em objLivroFilial

Dim lErro As Long
Dim objLivroFechado As New ClassLivrosFechados
Dim objLivroFilialLido As New ClassLivrosFilial
Dim lTransacao As Long

On Error GoTo Erro_Rotina_Fechamento_Livro

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 69609

    'Atualiza para o próximo livro já ficar aberto
    lErro = CF("Livro_Filial_Atualiza_Trans", objLivrosFilial, objLivroFilialLido)
    If lErro <> SUCESSO Then gError 69610

    'Passa os dados para o Livro Fechado para ser Fechado o Livro
    objLivroFechado.iCodLivro = objLivroFilialLido.iCodLivro
    objLivroFechado.iFilialEmpresa = objLivroFilialLido.iFilialEmpresa
    objLivroFechado.iNumeroLivro = objLivroFilialLido.iNumeroProxLivro
    objLivroFechado.iFolhaInicial = objLivroFilialLido.iNumeroProxFolha
    objLivroFechado.dtDataInicial = objLivroFilialLido.dtDataInicial
    objLivroFechado.dtDataFinal = objLivroFilialLido.dtDataFinal
    objLivroFechado.dtDataImpressao = objLivroFilialLido.dtImpressoEm

    'Insere o registro no livro Fechado com os valores que veio da Tabela de Livro Filial
    lErro = CF("Livro_Fechado_Insere", objLivroFechado)
    If lErro <> SUCESSO Then gError 69611

    'Depende de que Livro está sendo Fechado
    Select Case objLivrosFilial.iCodLivro

        Case LIVRO_REG_ENTRADA_ICMS_IPI_CODIGO
            lErro = CF("Rotina_Fechamento_Livro_RegEntrada", objLivroFechado)
            If lErro <> SUCESSO Then gError 69612

        Case LIVRO_REG_SAIDA_ICMS_IPI_CODIGO
            lErro = CF("Rotina_Fechamento_Livro_RegSaida", objLivroFechado)
            If lErro <> SUCESSO Then gError 69613

        Case LIVRO_REG_INVENTARIO_CODIGO
            lErro = CF("Rotina_Fechamento_Livro_RegInventario", objLivroFechado)
            If lErro <> SUCESSO Then gError 69614

        Case LIVRO_APURACAO_ICMS_CODIGO
            lErro = CF("Rotina_Fechamento_Livro_RegApuracaoICMS", objLivroFechado)
            If lErro <> SUCESSO Then gError 69615

        Case LIVRO_APURACAO_IPI_CODIGO
            lErro = CF("Rotina_Fechamento_Livro_RegApuracaoIPI", objLivroFechado)
            If lErro <> SUCESSO Then gError 69616

    End Select

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 69617

    Rotina_Fechamento_Livro = SUCESSO
    
    Exit Function
    
Erro_Rotina_Fechamento_Livro:

    Rotina_Fechamento_Livro = gErr

    Select Case gErr

        Case 70197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 69610, 69611, 69612, 69613, 69614, 69615, 69616
        
        Case 69617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150368)
    
    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegSaida(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Saida

Dim lErro As Long
Dim lNumIntDoc As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Fechamento_Livro_RegSaida

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69633
    Next

    'Lê os Registros de Saida colocando o NumIntDoc do livro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM LivRegES WHERE Data >= ? AND Data <= ? AND Tipo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, TIPO_REGES_SAIDA, objLivroFechado.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69634

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69635

    Do While lErro = AD_SQL_SUCESSO

        'Faz "lock" em LivrosFilial
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 69636

        'Atualiza o NumIntDocLivFechado dos Registros de Saida
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE LIvRegES SET NumIntDocLivFechado = ?", alComando(0), objLivroFechado.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69637

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69638

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Fechamento_Livro_RegSaida = SUCESSO

    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegSaida:

    Rotina_Fechamento_Livro_RegSaida = gErr

    Select Case gErr

        Case 69633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69634, 69635, 69638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case 69636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
        
        Case 69637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150369)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegInventario(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Inventario

Dim lErro As Long
Dim lNumIntDoc As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dtDataAuxiliar As Date

On Error GoTo Erro_Rotina_Fechamento_Livro_RegInventario

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69639
    Next

    'Lê os Registros de Inventario colocando o NumIntDoc do livro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Data FROM RegInventario WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ?", 0, dtDataAuxiliar, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69640

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69641

    Do While lErro = AD_SQL_SUCESSO

        'Faz "lock" em LivrosFilial
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 69642
        
        'Atualiza o NumIntDocLivFechado dos Registros de Saida
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegInventario SET NumIntDocLivFechado = ?", alComando(0), objLivroFechado.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69643

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69644

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Fechamento_Livro_RegInventario = SUCESSO
    
    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegInventario:

    Rotina_Fechamento_Livro_RegInventario = gErr

    Select Case gErr

        Case 69639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69640, 69641, 69644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 69642
        
        Case 69643
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150370)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegApuracaoICMS(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Apuracao ICMS

Dim lErro As Long
Dim lNumIntDoc As Long, lNumIntDocApuracao As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Fechamento_Livro_RegApuracaoICMS

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69645
    Next

    'Lê os Registros de Apuracao de ICMS colocando o NumIntDoc do livro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RegApuracaoICMS WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", 0, lNumIntDocApuracao, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69646

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69647

    If lErro = AD_SQL_SUCESSO Then

        'Atualiza o NumIntDocLivFechado dos Registros de Apuração de ICMS
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoICMS SET NumIntDocLivFechado = ?", alComando(0), objLivroFechado.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69648

        'Lê os Registros de Apuracao de ICMS Item colocando o NumIntDoc da Apuracao nos Itens
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM RegApuracaoICMSItem WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ?", 0, lNumIntDoc, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 69649

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69650

        Do While lErro = AD_SQL_SUCESSO

            'Atualiza o NumIntDocApuracao dos Registros de Apuração de ICMS Item
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE RegApuracaoICMSItem SET NumIntDocApuracao = ?", alComando(2), lNumIntDocApuracao)
            If lErro <> AD_SQL_SUCESSO Then gError 69651

            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69652

        Loop

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Fechamento_Livro_RegApuracaoICMS = SUCESSO
    
    Exit Function
    
Erro_Rotina_Fechamento_Livro_RegApuracaoICMS:

    Rotina_Fechamento_Livro_RegApuracaoICMS = gErr

    Select Case gErr

        Case 69645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69646, 69647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
        
        Case 69648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOICMS", gErr, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
        
        Case 69649, 69650, 69652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMSITEM", gErr)
                
        Case 69651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOICMSITEM1", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150371)
        
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Fechamento_Livro_RegApuracaoIPI(ByVal objLivroFechado As ClassLivrosFechados) As Long
'Fecha o Livro de Apuracao IPI

Dim lErro As Long
Dim lNumIntDoc As Long, lNumIntDocApuracao As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Fechamento_Livro_RegApuracaoIPI

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69653
    Next

    'Lê os Registros de Apuracao de IPI colocando o NumIntDoc do livro
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RegApuracaoIPI WHERE DataInicial = ? AND DataFinal = ? AND FilialEmpresa = ?", 0, lNumIntDocApuracao, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69654

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69655

    If lErro = AD_SQL_SUCESSO Then

        'Atualiza o NumIntDocLivFechado dos Registros de Apuração de IPI
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoIPI SET NumIntDocLivFechado = ?", alComando(0), objLivroFechado.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69656

        'Lê os Registros de Apuracao de IPI Item colocando o NumIntDoc da Apuracao nos Itens
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM RegApuracaoIPIItem WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ?", 0, lNumIntDoc, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 69657

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69658

        Do While lErro = AD_SQL_SUCESSO

            'Atualiza o NumIntDocApuracao dos Registros de Apuração de IPI Item
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE RegApuracaoIPIItem SET NumIntDocApuracao = ?", alComando(2), lNumIntDocApuracao)
            If lErro <> AD_SQL_SUCESSO Then gError 69659

            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69660

        Loop

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Fechamento_Livro_RegApuracaoIPI = SUCESSO

    Exit Function

Erro_Rotina_Fechamento_Livro_RegApuracaoIPI:

    Rotina_Fechamento_Livro_RegApuracaoIPI = gErr

    Select Case gErr

        Case 69653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69654, 69655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
        
        Case 69656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOIPI", gErr, objLivroFechado.dtDataInicial, objLivroFechado.dtDataFinal, objLivroFechado.iFilialEmpresa)
        
        Case 69657, 69658, 69660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)
                
        Case 69659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOIPIITEM1", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150372)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function LivrosFilial_Atualiza_DataImpresao(ByVal objLivroFilial As ClassLivrosFilial) As Long

Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim iCodLivro As Integer
Dim lTransacao As Long

On Error GoTo Erro_LivrosFilial_Atualiza_DataImpresao

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 70793
    Next
            
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 75391
            
    'Procura por Livro Fiscal da filial empresa em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodLivro FROM LivrosFilial WHERE FilialEmpresa = ? AND CodLivro = ?", 0, iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.iCodLivro)
    If lErro <> AD_SQL_SUCESSO Then gError 70794
     
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 70795
        
    'Se não encontrou o Livro Fiscal da filial, erro
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Procura Livro já fechado com intervalo de datas Inicial e Final maior do que as datas passadas
        lErro = Comando_ExecutarPos(alComando(2), "SELECT CodLivro FROM LivrosFechados WHERE DataInicial <= ? AND DataFinal >= ? AND FilialEmpresa = ? AND CodLivro = ?", 0, iCodLivro, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa, objLivroFilial.iCodLivro)
        If lErro <> AD_SQL_SUCESSO Then gError 69807
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69808
        
        'Se encontrou o Livro Fechado
        If lErro = AD_SQL_SUCESSO Then
        
            'Faz "Lock" em LivrosFilial
            lErro = Comando_LockExclusive(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 69809
        
            'Atualiza data de impressão de LivrosFilial
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE LivrosFechados SET DataImpressao = ?", alComando(2), gdtDataAtual)
            If lErro <> AD_SQL_SUCESSO Then gError 69810
        
        End If
        
    Else
    
        'Faz "Lock" em LivrosFilial
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 70797
    
        'Atualiza data de impressão de LivrosFilial
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE LivrosFilial SET ImpressoEm = ?", alComando(0), gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 70798
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 75392
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    LivrosFilial_Atualiza_DataImpresao = SUCESSO
    
    Exit Function
    
Erro_LivrosFilial_Atualiza_DataImpresao:

    Select Case gErr
    
        Case 70793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 70794, 70795
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)
        
        Case 70797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
        
        Case 70798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
            
        Case 75391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 75392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case 69807, 69808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)
        
        Case 69809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFECHADOS", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
        
        Case 69810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVROSFECHADOS", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150373)
    
    End Select
    
    Call Transacao_Rollback
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function NotaFiscal_Cancela_Fis(ByVal objNFiscal As ClassNFiscal) As Long
'Registra nas tabelas usadas por livros fiscais que a nf foi cancelada

Dim lErro As Long, iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim alComando(0 To 4) As Long, lNumIntDocRegES As Long
Dim lNumIntDocRegESLinha As Long

On Error GoTo Erro_NotaFiscal_Cancela_Fis

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81528
    Next
            
    If gobjFIS.dtFisBloqDataLimite <> DATA_NULA Then
    
        If objNFiscal.dtDataEntrada <> DATA_NULA Then
            If objNFiscal.dtDataEntrada < gobjFIS.dtFisBloqDataLimite Then gError 213154
        Else
            If objNFiscal.dtDataEmissao < gobjFIS.dtFisBloqDataLimite Then gError 213154
        End If
        
    End If
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o TipoDocInfo passado
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27263 Then gError 69200
    
    'não encontrado no BD --> Erro
    If lErro = 27263 Then gError 69201
    
    'só vou atualizar info para notas que interessem p/livros de ICMS/IPI
    If objTipoDocInfo.iModeloArqICMS <> 0 Then
    
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM LivRegES WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", 0, lNumIntDocRegES, TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 81529
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81535
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE LivRegES SET Situacao = ?", alComando(0), 1)
            If lErro <> AD_SQL_SUCESSO Then gError 81531
        
        End If
        
        'deixar apenas uma linha no reges, com valores zerados e com a observacao "CANCELADA"
        
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM LivRegESLinha WHERE NumIntDocRegES = ?", 0, lNumIntDocRegESLinha, lNumIntDocRegES)
        If lErro <> AD_SQL_SUCESSO Then gError 81532
                
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81536
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE LivRegESLinha SET ObservacaoLivFisc = ?, ValorTotal =0, ValorICMSBase = 0, ValorICMS = 0, ValorICMSIsentoNaoTrib = 0, ValorICMSOutras = 0, ValorICMSSubstBase = 0, ValorICMSSubstRet = 0, AliquotaICMS = 0, ValorDespAcess = 0, ValorIPI = 0, AliquotaIPI = 0, ValorIPIBase = 0, ValorIPIIsentoNaoTrib = 0, ValorIPIOutras = 0, ClassifContabil = ?, ValorContabil = 0, CodigoICMS = 3, CodigoIPI = 3", alComando(2), "CANCELADA", "")
            If lErro <> AD_SQL_SUCESSO Then gError 81534
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81537
            
            Do While lErro = AD_SQL_SUCESSO
            
                lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM LivRegESLinha", alComando(2))
                If lErro <> AD_SQL_SUCESSO Then gError 81538
            
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81539
            
            Loop
    
        End If
    
    End If
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NotaFiscal_Cancela_Fis = SUCESSO
    
    Exit Function
    
Erro_NotaFiscal_Cancela_Fis:

    NotaFiscal_Cancela_Fis = gErr
    
    Select Case gErr
        
        Case 213154
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_BLOQUEIA_DATA", gErr, "Movimento Fiscal", gobjFIS.dtFisBloqDataLimite)
        
        Case 81528
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 81529, 81535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
            
        Case 81530
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVREGES_NAO_ENCONTRADO", gErr)
            
        Case 81531
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
            
        Case 81532, 81536, 81537, 81539
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESLINHA", gErr)
            
        Case 81533
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVREGESLINHA_NAO_ENCONTRADO", gErr)
            
        Case 81534
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGESLINHA", gErr)
            
        Case 81538
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVREGESLINHA", gErr)
        
        Case 69200
        
        Case 69201
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150374)

    End Select
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Function NotaFiscal_Altera_Fis(ByVal objNFiscal As ClassNFiscal) As Long
'Trata a alteracao de dados de uma nf em relacao aos livros fiscais
'obs.:atualmente tratando apenas a alteracao das datas de emissao e de entrada

Dim lErro As Long, dtDataEmissaoBD As Date, dtDataBD As Date
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim alComando(1 To 2) As Long, lNumIntDocRegES As Long
Dim lNumIntLivroFechado As Long, iIndice As Integer

On Error GoTo Erro_NotaFiscal_Altera_Fis

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81657
    Next
            
    If gobjFIS.dtFisBloqDataLimite <> DATA_NULA Then
    
        If objNFiscal.dtDataEntrada <> DATA_NULA Then
            If objNFiscal.dtDataEntrada < gobjFIS.dtFisBloqDataLimite Then gError 213154
        Else
            If objNFiscal.dtDataEmissao < gobjFIS.dtFisBloqDataLimite Then gError 213154
        End If
        
    End If
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o TipoDocInfo passado
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27263 Then gError 81654
    
    'não encontrado no BD --> Erro
    If lErro = 27263 Then gError 81655
    
    'só vou atualizar info para notas que interessem p/livros de ICMS/IPI
    If objTipoDocInfo.iModeloArqICMS <> 0 Then
    
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, NumIntDocLivFechado, DataEmissao, Data FROM LivRegES WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", 0, lNumIntDocRegES, lNumIntLivroFechado, dtDataEmissaoBD, dtDataBD, TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 81656
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81662
        
        If lErro <> AD_SQL_SUCESSO Then gError 81658
        
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 81659
        
        'Se for uma Nota Fiscal de Saída
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
            
            'se alterou a data de emissao
            If dtDataEmissaoBD <> objNFiscal.dtDataEmissao Or dtDataBD <> objNFiscal.dtDataEmissao Then
            
                '??? incluir aviso
                If lNumIntLivroFechado <> 0 Then
                End If
                
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE LivRegES SET DataEmissao = ?, Data = ?", alComando(1), objNFiscal.dtDataEmissao, objNFiscal.dtDataEmissao)
                If lErro <> AD_SQL_SUCESSO Then gError 81660
            
            End If
            
        'Se for uma Nota Fiscal de Entrada
        ElseIf objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
            
            'se alterou a data de emissao ou a data de entrada
            If dtDataEmissaoBD <> objNFiscal.dtDataEmissao Or dtDataBD <> objNFiscal.dtDataEntrada Then
            
                '??? incluir aviso
                If lNumIntLivroFechado <> 0 Then
                End If
                
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE LivRegES SET DataEmissao = ?, Data = ?", alComando(1), objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
                If lErro <> AD_SQL_SUCESSO Then gError 81661
            
            End If
        
        End If
    
    End If
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NotaFiscal_Altera_Fis = SUCESSO
    
    Exit Function
    
Erro_NotaFiscal_Altera_Fis:

    NotaFiscal_Altera_Fis = gErr
    
    Select Case gErr
        
        Case 213154
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_BLOQUEIA_DATA", gErr, "Movimento Fiscal", gobjFIS.dtFisBloqDataLimite)
        
        Case 81654
        
        Case 81655
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case 81657
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 81656, 81662
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
            
        Case 81658
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVREGES_NAO_ENCONTRADO", gErr)
            
        Case 81660, 81661
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
            
        Case 81659
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150375)

    End Select
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Function NF_Item_Move_Tela_LivRegESLinha(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objLivRegES As ClassLivRegES, ByVal colLivRegESLinha As Collection) As Long
'Preenhe a coleção

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sContaContabil As String, bProdutoLido As Boolean

On Error GoTo Erro_NF_Item_Move_Tela_LivRegESLinha

    'se a Empresa nao desconsidera a cta contabil nos livros fiscais
    If gobjFIS.iLivFiscSemCtaCtb = 0 Then
    
        'se o tipo de controle do item nao está setado
        If objItemNF.iControleEstoque = 0 Then
            
            objProduto.sCodigo = objItemNF.sProduto
            
            'Lê o Produto para ler a conta Aplicação
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 69219
                
            If lErro = 28030 Then gError 69220
        
            objItemNF.iControleEstoque = objProduto.iControleEstoque
            
            bProdutoLido = True
            
        Else
        
            bProdutoLido = False
            
        End If
        
        'Se o Produto é estocável pega a Conta
        If objItemNF.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then
        
            If bProdutoLido = False Then
            
                objProduto.sCodigo = objItemNF.sProduto
                
                'Lê o Produto para ler a conta Aplicação
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 69219
                    
                If lErro = 28030 Then gError 69220
                
            End If
            
            sContaContabil = objProduto.sContaContabil
            
        Else
            If objItemNF.ColAlocacoes.Count = 1 Then
                
                objEstoqueProduto.sProduto = objItemNF.sProduto
                objEstoqueProduto.iAlmoxarifado = objItemNF.ColAlocacoes(1).iAlmoxarifado
    
                'Lê o EstoqueProduto para ler a conta de Estoque
                lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
                If lErro <> SUCESSO And lErro <> 21306 Then gError 69221
    
                If lErro = 21306 Then gError 69222
    
                sContaContabil = objEstoqueProduto.sContaContabil
            
            Else
                sContaContabil = ""
            End If
            
        End If
        
    Else
    
        sContaContabil = ""
        
    End If
    
    lErro = NF_Move_Tela_LivRegESLinha_Obj(objNFiscal, objItemNF.objTributacaoItemNF, objItemNF.dValorTotal, sContaContabil, colLivRegESLinha, objLivRegES, objItemNF.dValorDesconto, objItemNF.dValorDescGlobal)
    If lErro <> SUCESSO Then gError 81667
    
    NF_Item_Move_Tela_LivRegESLinha = SUCESSO
        
    Exit Function
    
Erro_NF_Item_Move_Tela_LivRegESLinha:

    NF_Item_Move_Tela_LivRegESLinha = gErr
    
    Select Case gErr
        
        Case 69219, 69221, 81667
        
        Case 69220
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 69222
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150376)

    End Select
    
    Exit Function

End Function

Private Function NF_Move_Tela_LivRegESLinha_Obj(ByVal objNFiscal As ClassNFiscal, ByVal objTributacaoObj As Object, ByVal dValorObj As Double, ByVal sContaFormatada As String, ByVal colLivRegESLinha As Collection, ByVal objLivRegES As ClassLivRegES, ByVal dValorDesconto As Double, ByVal dValorDescGlobal As Double) As Long

Dim iEncontrou As Integer
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTribIPI As New ClassTipoTribIPI
Dim lErro As Long, dAliquotaIPI As Double, dAliquotaICMS As Double
Dim objProduto As New ClassProduto, dValorIPIBase As Double, dValorICMSBase As Double, dValorICMSOutras As Double
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lNumIntDoc As Long, dDifBaseObj As Double, dValorDespImp As Double, dValorICMSBaseDespImp As Double, dValorICMSDespImp As Double
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto, dICMSOutrasAux As Double
Dim sContaContabil As String, dICMSValRedBase As Double, dIPIValRedBase As Double
'Dim objFilialEmpresa As New AdmFiliais, dValorContabil As Double
Dim dValorRateioItem As Double, dValorAux As Double, dValorTotalItem As Double, dAcrescimoIPI As Double

On Error GoTo Erro_NF_Move_Tela_LivRegESLinha_Obj

'    objFilialEmpresa.iCodFilial = objNFiscal.iFilialEmpresa
'
'    'Le os dados da Filial Empresa
'    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
'    If lErro <> SUCESSO And lErro <> 27378 Then gError 69234
'
'    If lErro <> SUCESSO Then gError 69235
'
    If objNFiscal.iTipoNFiscal <> DOCINFO_NFIEIMP And objNFiscal.iTipoNFiscal <> DOCINFO_NFIEIMPSE And objNFiscal.iTipoNFiscal <> DOCINFO_NFIEIMPPC Then
        '?????? WAGNER, VEJA SE ESTAS DUAS LINHAS ABAIXO SAO NECESSARIAS P/EVITAR O ERRO QUE DEU NA DETECTOR DE NAO INCLUIR FRETE, SEGURO E OUTRAS NO IPIOutras
        dValorRateioItem = objTributacaoObj.dValorFreteItem + objTributacaoObj.dValorOutrasDespesasItem + objTributacaoObj.dValorSeguroItem
        dValorObj = dValorObj + dValorRateioItem
        '?????? FIM DA OBSERVACAO
    End If
    
    objTipoDeTributacao.iTipo = objTributacaoObj.iTipoTributacao

    lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
    If lErro <> SUCESSO And lErro <> 27259 Then Error 27793

    If lErro = 27259 Then Error 27794
    
    'Lê o Tipo de Tributação para pegar a coluna no Livro Fiscal
    objTipoTribICMS.iTipo = objTributacaoObj.iICMSTipo
    
    lErro = CF("TipoTribICMS_Le", objTipoTribICMS)
    If lErro <> SUCESSO And lErro <> 21534 Then gError 69277
                        
    If lErro = 21534 Then gError 69278
                        
    'Lê o Tipo de Tributação para pegar a coluna no Livro Fiscal
    objTipoTribIPI.iTipo = objTributacaoObj.iIPITipo

    lErro = CF("TipoTribIPI_Le", objTipoTribIPI)
    If lErro <> SUCESSO And lErro <> 21620 Then gError 69279

    If lErro = 21620 Then gError 69280
        
    If objTributacaoObj.dICMSCredito <> 0 Then
        dICMSValRedBase = Round(objTributacaoObj.dICMSBase * objTributacaoObj.dICMSPercRedBase, 4)
        dValorICMSBase = Round(objTributacaoObj.dICMSBase - dICMSValRedBase, 4)
        dAliquotaICMS = objTributacaoObj.dICMSAliquota
    End If
    
    If objTributacaoObj.dIPICredito <> 0 Then
        dIPIValRedBase = Round(objTributacaoObj.dIPIBaseCalculo * objTributacaoObj.dIPIPercRedBase, 4)
        dValorIPIBase = Round(objTributacaoObj.dIPIBaseCalculo - dIPIValRedBase, 4)
        dAliquotaIPI = objTributacaoObj.dIPIAliquota
    End If

'    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP And objTributacaoObj.iTipo = ITEM_TIPO_NORMAL Then
'        dValorDespImp = objTributacaoObj.dDespImpICMSBase
'        dValorICMSDespImp = objTributacaoObj.dDespImpICMSValor
'        If objTributacaoObj.dDespImpICMSCredito <> 0 Then
'            dValorICMSBaseDespImp = objTributacaoObj.dDespImpICMSBase
'            dValorICMSBase = Round(dValorICMSBase + dValorICMSBaseDespImp, 4)
'        End If
'    End If
    
    'Trata o Preenchimento do objLivRegESLinha
    For Each objLivRegESLinha In colLivRegESLinha
    
        'verifica se já existe na colecao item com as mesmas Natureza, Aliquotas (IPI e ICMS), conta contabil e colunas no livro (IPI e ICMS)
        If objLivRegESLinha.sNaturezaOp = objTributacaoObj.sNaturezaOp And objLivRegESLinha.dAliquotaICMS = dAliquotaICMS And objLivRegESLinha.dAliquotaIPI = dAliquotaIPI And objLivRegESLinha.sClassifContabil = sContaFormatada Then
            
            iEncontrou = 1
                    
            Exit For
        
        End If
    
    Next
    
    'Se não encontrou nenhuma Linha com a mesma Natureza e Alíquota
    If iEncontrou = 0 Then
        
        'insere uma Linha
        Set objLivRegESLinha = New ClassLivRegESLinha
        
        objLivRegESLinha.lNumIntDocRegES = objLivRegES.lNumIntDoc
        
        objLivRegESLinha.dAliquotaICMS = dAliquotaICMS
        objLivRegESLinha.sNaturezaOp = objTributacaoObj.sNaturezaOp
        objLivRegESLinha.dAliquotaIPI = dAliquotaIPI
        objLivRegESLinha.iCodigoICMS = objTipoTribICMS.iColunaNoLivroSaida
        objLivRegESLinha.iCodigoIPI = objTipoTribIPI.iColunaNoLivro
        objLivRegESLinha.sClassifContabil = sContaFormatada
        
        colLivRegESLinha.Add objLivRegESLinha
        
    End If
    
    dAcrescimoIPI = 0
            
    objLivRegESLinha.dValorICMSBase = Round(objLivRegESLinha.dValorICMSBase + dValorICMSBase, 4)
    objLivRegESLinha.dValorIPIBase = Round(objLivRegESLinha.dValorIPIBase + dValorIPIBase, 4)
    
    objLivRegESLinha.dValorICMSSubstBase = Round(objLivRegESLinha.dValorICMSSubstBase + objTributacaoObj.dICMSSubstBase, 4)
    objLivRegESLinha.dValorICMSSubstRet = Round(objLivRegESLinha.dValorICMSSubstRet + objTributacaoObj.dICMSSubstValor, 4)
        
    dValorTotalItem = 0
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
        
        dValorTotalItem = Arredonda_Moeda(dValorObj + objTributacaoObj.dIPIValor + objTributacaoObj.dICMSSubstValor + objTributacaoObj.dValorFreteItem + objTributacaoObj.dValorOutrasDespesasItem + objTributacaoObj.dValorSeguroItem)
       
        If objTributacaoObj.dICMSBase <> 0 And objTributacaoObj.dICMSBase > dValorObj And gobjCRFAT.iNFImportacaoTribFlag03 = DESMARCADO Then
            dValorTotalItem = objTributacaoObj.dICMSBase
        Else
            If gobjCRFAT.iNFImportacaoTribFlag04 = DESMARCADO Then dValorTotalItem = dValorTotalItem + objTributacaoObj.dICMSValor
            If gobjCRFAT.iNFImportacaoTribFlag08 = MARCADO Then dValorTotalItem = dValorTotalItem - objTributacaoObj.dValorOutrasDespesasItem
            If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dValorTotalItem = dValorTotalItem + objTributacaoObj.dIIValor
        
            If gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO Then dValorTotalItem = dValorTotalItem + objTributacaoObj.dPISValor + objTributacaoObj.dCOFINSValor
        
        End If
        
        If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dAcrescimoIPI = dAcrescimoIPI + objTributacaoObj.dIIValor
        
        objLivRegESLinha.dValorContabil = Round(objLivRegESLinha.dValorContabil + dValorTotalItem, 4)
        objLivRegESLinha.dValorTotal = Round(objLivRegESLinha.dValorTotal + dValorTotalItem, 4)
        If objTributacaoObj.iTipo = ITEM_TIPO_NORMAL Then objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + objTributacaoObj.dValorOutrasDespesasItem + objTributacaoObj.dICMSValor + IIf(gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO, objTributacaoObj.dPISValor + objTributacaoObj.dCOFINSValor, 0), 4)
        If dValorICMSBase = 0 Then dValorAux = objTributacaoObj.dValorOutrasDespesasItem
        
    Else
    
    '    dValorContabil = Round(dValorObj + objTributacaoObj.dIPIValor, 4)
        objLivRegESLinha.dValorContabil = Round(objLivRegESLinha.dValorContabil + dValorObj + objTributacaoObj.dIPIValor - dValorDescGlobal + objTributacaoObj.dICMSSubstValor, 4)
        objLivRegESLinha.dValorTotal = Round(objLivRegESLinha.dValorTotal + dValorObj + objTributacaoObj.dIPIValor - dValorDescGlobal + objTributacaoObj.dICMSSubstValor, 4)
        
        dValorAux = 0

    End If
    
    '    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
    '        objLivRegESLinha.dValorContabil = Round(objLivRegESLinha.dValorContabil + objTributacaoObj.dValorFreteItem + objTributacaoObj.dValorOutrasDespesasItem + objTributacaoObj.dValorSeguroItem + objTributacaoObj.dICMSValor, 4)
    '        objLivRegESLinha.dValorTotal = Round(objLivRegESLinha.dValorTotal + objTributacaoObj.dValorFreteItem + objTributacaoObj.dValorOutrasDespesasItem + objTributacaoObj.dValorSeguroItem + objTributacaoObj.dICMSValor, 4)
    '        If objTributacaoObj.iTipo = ITEM_TIPO_NORMAL Then objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + objTributacaoObj.dValorOutrasDespesasItem + objTributacaoObj.dICMSValor, 4)
    '        If dValorICMSBase = 0 Then dValorAux = objTributacaoObj.dValorOutrasDespesasItem
    '    End If
    
'    If objFilialEmpresa.iContribuinteIPI <> 0 Then
        
        If objTributacaoObj.dICMSCredito <> 0 Then
            objLivRegESLinha.dValorICMS = Round(objLivRegESLinha.dValorICMS + objTributacaoObj.dICMSValor + dValorICMSDespImp, 4)
            
            If right(objTributacaoObj.sCST, 2) = "51" Then
                objLivRegESLinha.dValorICMS51 = Round(objLivRegESLinha.dValorICMS51 + objTributacaoObj.dICMSValor + dValorICMSDespImp, 4)
            End If
            
            objLivRegESLinha.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha.dValorICMSIsentoNaoTrib + dICMSValRedBase, 4)
            If objTributacaoObj.dICMSBase < (dValorObj - dValorDescGlobal) Then
                objLivRegESLinha.dValorICMSOutras = Round(objLivRegESLinha.dValorICMSOutras + dValorObj - objTributacaoObj.dICMSBase - dValorDescGlobal, 4)
            End If
    
        ElseIf objTipoTribICMS.iColunaNoLivroSaida = TIPO_ISENTO_NAO_TRIBUTADO Then
            objLivRegESLinha.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha.dValorICMSIsentoNaoTrib + IIf(objTributacaoObj.dICMSBase <= (dValorObj - dValorDescGlobal), dValorObj - dValorDescGlobal, objTributacaoObj.dICMSBase), 4)
        Else
            dICMSOutrasAux = IIf(objTributacaoObj.dICMSBase <= (dValorObj - dValorDescGlobal + dValorAux), dValorObj - dValorDescGlobal + dValorAux, objTributacaoObj.dICMSBase)
            If objTipoDeTributacao.iICMSBaseComIPI And dICMSOutrasAux > objTributacaoObj.dIPIValor Then
                dICMSOutrasAux = dICMSOutrasAux - objTributacaoObj.dIPIValor
            End If
            objLivRegESLinha.dValorICMSOutras = Round(objLivRegESLinha.dValorICMSOutras + dICMSOutrasAux, 4)
        End If
        
'    Else
'
'        If objTributacaoObj.dICMSCredito <> 0 Then
'            objLivRegESLinha.dValorICMS = Round(objLivRegESLinha.dValorICMS + objTributacaoObj.dICMSValor, 4)
'            If objTributacaoObj.dICMSBase < dValorObj Then
'                dValorICMSOutras = dValorObj - objTributacaoObj.dICMSBase
'                objLivRegESLinha.dValorICMSOutras = Round(objLivRegESLinha.dValorICMSOutras + dValorObj - objTributacaoObj.dICMSBase, 4)
'            End If
'            If dValorContabil > (dValorICMSBase + dValorICMSOutras) Then
'                objLivRegESLinha.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha.dValorICMSIsentoNaoTrib + (dValorContabil - (dValorICMSBase + dValorICMSOutras)), 4)
'            End If
'        ElseIf objTipoTribICMS.iColunaNoLivroSaida = TIPO_ISENTO_NAO_TRIBUTADO Then
'            objLivRegESLinha.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha.dValorICMSIsentoNaoTrib + dValorObj + objTributacaoObj.dIPIValor, 4)
'        Else
'            objLivRegESLinha.dValorICMSOutras = Round(objLivRegESLinha.dValorICMSOutras + dValorObj + objTributacaoObj.dIPIValor, 4)
'        End If
'
'    End If
    
    If gobjCRFAT.iIPIAbateDesconto <> 0 Then
    
        If objTributacaoObj.dIPICredito <> 0 Then
            objLivRegESLinha.dValorIPI = Round(objLivRegESLinha.dValorIPI + objTributacaoObj.dIPIValor, 4)
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dIPIValRedBase, 4)
            If objTributacaoObj.dIPIBaseCalculo < (dValorObj - dValorDescGlobal) Then
                objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj + dAcrescimoIPI - objTributacaoObj.dIPIBaseCalculo - dValorDescGlobal, 4)
            End If
        ElseIf objTipoTribIPI.iColunaNoLivro = TIPO_ISENTO_NAO_TRIBUTADO Then
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dValorObj + dAcrescimoIPI - dValorDescGlobal, 4)
        Else
            objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj + dAcrescimoIPI + IIf(gobjFIS.iIPIOutrasIncluiIPI <> 0, objTributacaoObj.dIPIValor, 0) - dValorDescGlobal, 4)
        End If

    Else
    
        If objTributacaoObj.dIPICredito <> 0 Then
            objLivRegESLinha.dValorIPI = Round(objLivRegESLinha.dValorIPI + objTributacaoObj.dIPIValor, 4)
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dIPIValRedBase, 4)
            If objTributacaoObj.dIPIBaseCalculo < (dValorObj) Then
                objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj + dAcrescimoIPI - objTributacaoObj.dIPIBaseCalculo, 4)
            End If
        ElseIf objTipoTribIPI.iColunaNoLivro = TIPO_ISENTO_NAO_TRIBUTADO Then
            objLivRegESLinha.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha.dValorIPIIsentoNaoTrib + dValorObj + dAcrescimoIPI, 4)
        Else
            objLivRegESLinha.dValorIPIOutras = Round(objLivRegESLinha.dValorIPIOutras + dValorObj + dAcrescimoIPI + IIf(gobjFIS.iIPIOutrasIncluiIPI <> 0, objTributacaoObj.dIPIValor, 0) - dValorDescGlobal, 4) 'alterado em 13/06/07 para abater o desconto
        End If
    
    End If
    
    NF_Move_Tela_LivRegESLinha_Obj = SUCESSO
     
    Exit Function
    
Erro_NF_Move_Tela_LivRegESLinha_Obj:

    NF_Move_Tela_LivRegESLinha_Obj = gErr
     
    Select Case gErr
          
        Case 69277, 69279, 27793, 69234
        
'        Case 69235
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case 69278, 27794
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_ICMS_INEXISTENTE", gErr, objTipoTribICMS.iTipo)
        
        Case 69280
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_IPI_INEXISTENTE", gErr, objTipoTribIPI.iTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150378)
     
    End Select
     
    Exit Function
    
End Function

Public Function LivRegESLinha_RefazerNFs(ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal bSoNovas As Boolean) As Long
'recalcula os registros de LivRegESLinha para um grupo de nfs
'iFilialEmpresa NAO PODE ser EMPRESA_TODA

Dim lErro As Long, iIndice As Integer
Dim tLivRegES As typeLivRegES, lNumIntNF As Long
Dim alComando(1 To 10) As Long
Dim lTransacao As Long, objItemNF As ClassItemNF, objLivRegESCadProd As New ClassLivRegESCadProd
Dim objNFiscal As ClassNFiscal, objLivRegES As ClassLivRegES

On Error GoTo Erro_LivRegESLinha_RefazerNFs

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81669
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81670

    If Not bSoNovas Then
    
        With tLivRegES
            lErro = Comando_Executar(alComando(1), "SELECT NumIntDocOrigem FROM LivRegES WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND Situacao = 0 ORDER BY Data", _
                .lNumIntDocOrigem, iFilialEmpresa, dtDataInicial, dtDataFinal)
            .iFilialEmpresa = iFilialEmpresa
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 81675
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81676
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objNFiscal = New ClassNFiscal
            Set objLivRegES = New ClassLivRegES
            
            objNFiscal.lNumIntDoc = tLivRegES.lNumIntDocOrigem
            objNFiscal.iFilialEmpresa = tLivRegES.iFilialEmpresa
            
            'Tenta ler a nota Fiscal passada por parametro
            lErro = CF("NFiscal_Le", objNFiscal)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 69402
                
            'Se não encontrou a NFiscal --> erro
            If lErro <> SUCESSO Then gError 69403
            
            objLivRegES.lNumIntNF = objNFiscal.lNumIntDoc
            
            'Le o Livro Fiscal de Acordo com a Nota Fiscal
            lErro = CF("LivRegES_Le_NotaFiscal", objLivRegES)
            If lErro <> SUCESSO And lErro <> 69464 Then gError 69406
            
            'Se não encontrar --> Erro
            If lErro = 69464 Then gError 69505
            
            'Lê os Ítens da Nota Fiscal
            lErro = CF("NFiscalItens_Le", objNFiscal)
            If lErro <> SUCESSO Then gError 69395
                    
            'Lê a parte de Tributação
            lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
            If lErro <> SUCESSO Then gError 69396
        
            'Lê a parte de tributacao dos Ítens da Nota Fiscal
            lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
            If lErro <> SUCESSO Then gError 69397
            
            'Pega a tributacao do Livro e não da Nota
            Call Preenche_Tributacao_ItemRegES(objLivRegES, objNFiscal)
        
            lErro = Recalcula_Linhas_LivroFiscal(objNFiscal, objLivRegES)
            If lErro <> SUCESSO Then gError 81673
            
            'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
            For Each objItemNF In objNFiscal.ColItensNF
                
                'Move os dados do Item NF para a coleção colLivRegESCadProd
                lErro = NF_Grava_LivRegESCadProd(objLivRegES.dtData, objItemNF, objLivRegESCadProd)
                If lErro <> SUCESSO Then gError 69207
                
            Next
            
            'Grava as Linhas do Livro
            lErro = CF("EdicaoRegES_Grava_Linha_Trans", objLivRegES)
            If lErro <> SUCESSO Then gError 69507
                
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81677
        
        Loop
    
    End If
    
    'tratar nfs que tiveram tiposdocinfo alterados p/passar a aparecer em livros fiscais
    lErro = Comando_Executar(alComando(2), "SELECT NFiscal.NumIntDoc FROM NFiscal, TiposDocInfo WHERE NFiscal.Filialempresa = ? AND Nfiscal.DataEmissao >= ? AND Nfiscal.DataEmissao <= ? AND TiposDocInfo.ModeloArqIcms <> 0 AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND NOT EXISTS (SELECT NumIntDoc FROM LivREGES WHERE NumIntDocORigem = NFiscal.NumIntDoc)", lNumIntNF, iFilialEmpresa, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 81729
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81730
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNFiscal = New ClassNFiscal
        objNFiscal.lNumIntDoc = lNumIntNF
        
        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 69402
            
        'Se não encontrou a NFiscal --> erro
        If lErro <> SUCESSO Then gError 69403
            
        'Lê os Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 69395
                
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO Then gError 69396
    
        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError 69397
            
        lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
        If lErro <> SUCESSO Then gError 81731
            
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81732
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81671

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    LivRegESLinha_RefazerNFs = SUCESSO
     
    Exit Function
    
Erro_LivRegESLinha_RefazerNFs:

    LivRegESLinha_RefazerNFs = gErr
     
    Select Case gErr
          
        Case 81729 To 81732
        
        Case 81673, 81674
        
        Case 81669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 81675, 81676, 81677
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150379)
     
    End Select
     
    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Recalcula_Linhas_LivroFiscal(ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES) As Long
'Faz calculo automatico das linhas dos Livros de Reg de Entrada e de Saída correspondentes à nota fiscal passada.

Dim lErro As Long
Dim objItemNF As New ClassItemNF, objNFiscalOrig As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objLivRegESLinha As New ClassLivRegESLinha

On Error GoTo Erro_Recalcula_Linhas_LivroFiscal
        
    Set objLivRegES.colLinhas = New Collection
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o TipoDocInfo passado
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27263 Then gError 69501
    
    'não encontrado no BD
    If lErro = 27263 Then gError 69502
    
    'só vou gerar info para notas que interessem p/livros de ICMS/IPI
    If objTipoDocInfo.iModeloArqICMS <> 0 Then
    
        If objNFiscal.iStatus = STATUS_CANCELADO Then
        
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "CANCELADA", "", objLivRegES.colLinhas)
        
        Else
        
            'para nao gerar detalhamento p/notas de simples faturamento
            Select Case objTipoDocInfo.iModeloArqICMS
            
                Case 0
                    'não faz nada
                
                Case MODREGFIS_DEV_SIMB_CONSIG
                
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Compra em Consignação - NF nº " + CStr(objNFiscal.lNumNotaFiscal) + " de " + CStr(objNFiscal.dtDataEmissao), "", objLivRegES.colLinhas)
                            
                Case MODREGFIS_FAT_ANTECIP 'faturamento antecipado
                    'vai incluir uma linha com a identificacao da nota e uma observacao
                
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "faturamento antecipado", "", objLivRegES.colLinhas)
                    
                Case MODREGFIS_SIMPLES_FAT 'simples faturamento
                    'vai incluir uma linha com a identificacao da nota e uma observacao
                
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "venda em consignação", "venda em consig. nf ", objLivRegES.colLinhas)
                                    
                Case MODREGFIS_DEV_SIMPLES_FAT 'devolucao de simples faturamento
                    'vai incluir uma linha com a identificacao da nota e uma observacao
                
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "devolução de venda em consignação", "devolução de venda em consig. nf ", objLivRegES.colLinhas)
        
                Case MODREGFIS_FAT_RESSARC_ST 'ressarcimento tributario
                    'vai incluir uma linha com a identificacao da nota e uma observacao
                
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Ressarcimento de ICMS substituição tributária", "Ressarc. ICMS subst. trib. ", objLivRegES.colLinhas)
                
                Case MODREGFIS_SIMPLES_FAT_COMPRA_CONTA_E_ORDEM '= 9005 'Compra à ordem - simples faturamento
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Compra à ordem - simples faturamento", "Compra à ordem - simples faturamento ", objLivRegES.colLinhas)

                Case MODREGFIS_SIMPLES_FAT_VENDA_CONTA_E_ORDEM '= 9006 'Venda à ordem - simples faturamento
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Venda à ordem - simples faturamento", "Venda à ordem - simples faturamento ", objLivRegES.colLinhas)

                Case MODREGFIS_COMPRA_CONTA_E_ORDEM '= 9007 'Compra à ordem
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Compra à ordem", "Compra à ordem ", objLivRegES.colLinhas)
                
                Case MODREGFIS_IMPORTACAO_PARCELADA
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Remessas parceladas", "Remessas parceladas de que trata a nf ", objLivRegES.colLinhas)
                
                Case MODREGFIS_ANULACAO_SERVICO_TRANSPORTE
                    Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Anulação de Serviço de Transporte", "", objLivRegES.colLinhas)
                
                Case Else
                
                    'se existe desconto global à nf
                    If objNFiscal.dValorDesconto <> 0 Then
                    
                        lErro = NF_DescGlobal_Rateia(objNFiscal)
                        If lErro <> SUCESSO Then gError 69208
                    
                    End If
            
                    'Verifica se é uma nota Fiscal de Entrada ou se é uma Nota Fiscal de Saída
                    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
                        
                        'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
                        For Each objItemNF In objNFiscal.ColItensNF
                        
                            'Grava os dados da Nota para a coleção objLivRegES.colLinhas
                            lErro = NF_Saida_Move_Tela_LivRegESLinha(objNFiscal, objItemNF, objLivRegES, objLivRegES.colLinhas)
                            If lErro <> SUCESSO Then gError 69503
                                
                            'Move os dados do Item NF para a coleção colLivRegESItemNF
                            lErro = Edicao_RegEntrada_Altera_Item(objItemNF, objNFiscal, objLivRegES.ColItensNF)
                            If lErro <> SUCESSO Then gError 69504
                                
                        Next
                        
                        'Move os dados de Frete, Seguro, Outras
                        lErro = Edicao_RegSaida_Altera_Complemento(objNFiscal, objLivRegES, objLivRegES.colLinhas, objLivRegES.ColItensNF)
                        If lErro <> SUCESSO Then gError 69505
                        
                    ElseIf objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Then
                    
                        For Each objItemNF In objNFiscal.ColItensNF
                        
                            'Grava os dados da Nota para a coleção objLivRegES.colLinhas
                            lErro = NF_Entrada_Move_Tela_LivRegESLinha(objNFiscal, objItemNF, objLivRegES, objLivRegES.colLinhas)
                            If lErro <> SUCESSO Then gError 69506
                                
                            'Move os dados do Item NF para a coleção colLivRegESItemNF
                            lErro = Edicao_RegEntrada_Altera_Item(objItemNF, objNFiscal, objLivRegES.ColItensNF)
                            If lErro <> SUCESSO Then gError 69507
                        
                        Next
                        
                        'Move os dados de Frete, Seguro, Outras
                        lErro = Edicao_RegEntrada_Altera_Complemento(objNFiscal, objLivRegES, objLivRegES.colLinhas, objLivRegES.ColItensNF)
                        If lErro <> SUCESSO Then gError 69508
                    
                    End If
                
                    'Ajustas os valores para bater igual a tributação
                    lErro = NF_LivRegESLinha_Ajusta(objNFiscal, objLivRegES.colLinhas)
                    If lErro <> SUCESSO Then gError 69507
            
            End Select
    
        End If
        
    End If
    
    Recalcula_Linhas_LivroFiscal = SUCESSO
    
    Exit Function
    
Erro_Recalcula_Linhas_LivroFiscal:

    Recalcula_Linhas_LivroFiscal = gErr
    
    Select Case gErr
        
        Case 69501, 69503, 69504, 69505, 69506, 69507, 69508
        
        Case 69502
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150380)
        
    End Select
    
    Exit Function

End Function

Private Sub Preenche_Tributacao_ItemRegES(ByVal objLivRegES As ClassLivRegES, ByVal objNFiscal As ClassNFiscal)
'??? provisoria: apenas p/apoio ao recalculo do livro

Dim objLivRegESItemNF As New ClassLivRegESItemNF
Dim objItemNF As New ClassItemNF
    
    For Each objLivRegESItemNF In objLivRegES.ColItensNF
    
        For Each objItemNF In objNFiscal.ColItensNF
            
            'Caso seja um item
            If objLivRegESItemNF.iNumItem = objItemNF.iItem Then
        
                objItemNF.objTributacaoItemNF.dICMSAliquota = objLivRegESItemNF.dAliquotaICMS
                objItemNF.objTributacaoItemNF.dICMSBase = objLivRegESItemNF.dValorICMSBase
                objItemNF.objTributacaoItemNF.dICMSSubstBase = objLivRegESItemNF.dValorICMSSubstBase
                objItemNF.objTributacaoItemNF.dIPIValor = objLivRegESItemNF.dValorIPI
                objItemNF.objTributacaoItemNF.iICMSTipo = objLivRegESItemNF.iTipoTribICMS
                objItemNF.objTributacaoItemNF.iIPITipo = objLivRegESItemNF.iTipoTribIPI
                objItemNF.objTributacaoItemNF.dIPIAliquota = objLivRegESItemNF.dAliquotaIPI
                objItemNF.objTributacaoItemNF.dICMSPercRedBase = objLivRegESItemNF.dRedBaseICMS
                objItemNF.objTributacaoItemNF.dIPIPercRedBase = objLivRegESItemNF.dRedBaseIPI
                objItemNF.objTributacaoItemNF.iTipoTributacao = objLivRegESItemNF.iTipoTributacao
                objItemNF.objTributacaoItemNF.dICMSSubstAliquota = objLivRegESItemNF.dAliquotaSubst
                objItemNF.objTributacaoItemNF.dICMSSubstValor = objLivRegESItemNF.dValorSubst
                objItemNF.objTributacaoItemNF.dICMSValor = objLivRegESItemNF.dValorICMS
                objItemNF.objTributacaoItemNF.dIPIBaseCalculo = objLivRegESItemNF.dValorIPIBase
                objItemNF.objTributacaoItemNF.sNaturezaOp = objLivRegESItemNF.sCFOP
                Exit For
                        
            End If
        
        Next
        
    Next
        
End Sub

Function NotaFiscal_GeraLinhas_Fis(ByVal iModeloArqICMS As Integer, ByVal colLivRegESLinha As Collection, ByVal colLivRegESItemNF As Collection, ByVal objNFiscal As ClassNFiscal, ByVal objLivRegES As ClassLivRegES) As Long
'Gera as linhas para livros fiscais de acordo com o valor passado iModeloArqICMS
'Em geral, iModeloArqICMS será 1, o que indica que as linhas serão geradas no Case Else
'Caso o valor seja diferente de 1, significa que será gerada uma linha "especial"

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objLivRegESCadProd As New ClassLivRegESCadProd
Dim objLivRegESLinha As ClassLivRegESLinha
Dim objNFiscalOrig As New ClassNFiscal

On Error GoTo Erro_NotaFiscal_GeraLinhas_Fis

    'Gera as linhas de livros fiscais conforme o valor de iModeloArqICMS
    Select Case iModeloArqICMS
    
        Case MODREGFIS_DEV_SIMB_CONSIG
        
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Compra em Consignação - NF nº " + CStr(objNFiscal.lNumNotaFiscal) + " de " + CStr(objNFiscal.dtDataEmissao), "", colLivRegESLinha)
                    
        'Se for Faturamento Antecipado
        Case MODREGFIS_FAT_ANTECIP
        
            'vai incluir uma linha com a identificacao da nota e uma observacao
        
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "faturamento antecipado", "", colLivRegESLinha)
        
        Case MODREGFIS_SIMPLES_FAT 'simples faturamento
            'vai incluir uma linha com a identificacao da nota e uma observacao
        
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "venda em consignação", "venda em consig. nf ", colLivRegESLinha)
        
        Case MODREGFIS_DEV_SIMPLES_FAT 'devolucao de simples faturamento
            'vai incluir uma linha com a identificacao da nota e uma observacao
        
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "devolução de venda em consignação", "devolução de venda em consig. nf ", colLivRegESLinha)
        
        Case MODREGFIS_FAT_RESSARC_ST 'ressarcimento tributario
        
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Ressarcimento de ICMS substituição tributária", "Ressarc. ICMS subst. trib. ", colLivRegESLinha)

        Case MODREGFIS_SIMPLES_FAT_COMPRA_CONTA_E_ORDEM '= 9005 'Compra à ordem - simples faturamento
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Compra à ordem - simples faturamento", "Compra à ordem - simples faturamento ", colLivRegESLinha)

        Case MODREGFIS_SIMPLES_FAT_VENDA_CONTA_E_ORDEM '= 9006 'Venda à ordem - simples faturamento
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Venda à ordem - simples faturamento", "Venda à ordem - simples faturamento ", colLivRegESLinha)

        Case MODREGFIS_COMPRA_CONTA_E_ORDEM '= 9007 'Compra à ordem
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Compra à ordem", "Compra à ordem ", colLivRegESLinha)
        
        Case MODREGFIS_IMPORTACAO_PARCELADA
            Call LivRegESLinha_Observacao(objLivRegES, objLivRegESLinha, objNFiscal, "Remessas parceladas", "Remessas parceladas de que trata a nf ", colLivRegESLinha)
        
        Case Else
        
            'se existe desconto global à nf
            If objNFiscal.dValorDesconto <> 0 Then
            
                lErro = NF_DescGlobal_Rateia(objNFiscal)
                If lErro <> SUCESSO Then gError 69208
            
            End If
            
            'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
            For Each objItemNF In objNFiscal.ColItensNF
                
                'Grava os dados da Nota para a coleção colLivRegESLinha
                lErro = NF_Saida_Move_Tela_LivRegESLinha(objNFiscal, objItemNF, objLivRegES, colLivRegESLinha)
                If lErro <> SUCESSO Then gError 69206
                
                'Move os dados do Item NF para a coleção colLivRegESCadProd
                lErro = NF_Grava_LivRegESCadProd(objLivRegES.dtData, objItemNF, objLivRegESCadProd)
                If lErro <> SUCESSO Then gError 69207
                
                'Move os dados do Item NF para a coleção colLivRegESItemNF
                lErro = NF_Move_Memoria_LivRegESItemNF(objItemNF, objNFiscal, objLivRegES, objLivRegESCadProd, colLivRegESItemNF)
                If lErro <> SUCESSO Then gError 69208
                        
            Next
            
            'gError 99999
            
            'Move os dados de Frete, Seguro, Outras
            lErro = NF_Saida_Move_Tela_Complemento(objNFiscal, objLivRegES, colLivRegESLinha, colLivRegESItemNF)
            If lErro <> SUCESSO Then gError 69209
        
            'Ajustas os valores para bater igual a tributação
            lErro = NF_LivRegESLinha_Ajusta(objNFiscal, colLivRegESLinha)
            If lErro <> SUCESSO Then gError 69206
        
    End Select

    If iModeloArqICMS >= MODREGFIS_SIMPLES_FAT Then
    
        'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
        For Each objItemNF In objNFiscal.ColItensNF
            
            'Move os dados do Item NF para a coleção colLivRegESCadProd
            lErro = NF_Grava_LivRegESCadProd(objLivRegES.dtData, objItemNF, objLivRegESCadProd)
            If lErro <> SUCESSO Then gError 69207
            
        Next
    
    End If
    
    NotaFiscal_GeraLinhas_Fis = SUCESSO
    
    Exit Function

Erro_NotaFiscal_GeraLinhas_Fis:

    NotaFiscal_GeraLinhas_Fis = gErr
    
    Select Case gErr
    
        Case 69206 To 69209, 81628
        
        Case 81629
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscalOrig.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150381)
    
    End Select
    
    Exit Function
        
End Function

Private Function LivRegESLinha_Observacao(ByVal objLivRegES As ClassLivRegES, objLivRegESLinha As ClassLivRegESLinha, ByVal objNFiscal As ClassNFiscal, ByVal sTextoSemNFOrig As String, ByVal sTextoComNFOrig As String, ByVal colLivRegESLinha As Collection) As Long
'vai incluir uma linha com a identificacao da nota e uma observacao que pode ou nao incluir ref a nf original

Dim lErro As Long, objNFiscalOrig As New ClassNFiscal

On Error GoTo Erro_LivRegESLinha_Observacao

    Set objLivRegESLinha = New ClassLivRegESLinha

    objLivRegESLinha.lNumIntDocRegES = objLivRegES.lNumIntDoc
    
    objLivRegESLinha.sNaturezaOp = ""
    objLivRegESLinha.iCodigoICMS = 0
    objLivRegESLinha.iCodigoIPI = 0
    objLivRegESLinha.sClassifContabil = ""
    
    objLivRegESLinha.dValorContabil = 0
    objLivRegESLinha.dValorTotal = 0
    
    If objNFiscal.lNumIntNotaOriginal = 0 Or sTextoComNFOrig = "" Then
    
        'objLivRegESLinha.sObservacaoLivFisc = "venda em consignação"
        objLivRegESLinha.sObservacaoLivFisc = sTextoSemNFOrig
                
    Else
    
        objNFiscalOrig.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        lErro = CF("NFiscal_Le", objNFiscalOrig)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 81628
        If lErro <> SUCESSO Then gError 81629
        
        'objLivRegESLinha.sObservacaoLivFisc = "venda em consig. nf " & CStr(objNFiscalOrig.lNumNotaFiscal) & " de " & CStr(objNFiscalOrig.dtDataEmissao)
        objLivRegESLinha.sObservacaoLivFisc = sTextoComNFOrig & CStr(objNFiscalOrig.lNumNotaFiscal) & " de " & CStr(objNFiscalOrig.dtDataEmissao)
    
    End If
    
    colLivRegESLinha.Add objLivRegESLinha
    
    LivRegESLinha_Observacao = SUCESSO
     
    Exit Function
    
Erro_LivRegESLinha_Observacao:

    LivRegESLinha_Observacao = gErr
     
    Select Case gErr
          
        Case 81628, 81629
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150382)
     
    End Select
     
    Exit Function

End Function

Function GeracaoArqICMSProd_Grava(ByVal objInfoArqICMS As ClassInfoArqICMS, ByVal objGeracaoArqICMSProd As ClassGeracaoArqICMSProd, ByVal lComando As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim lNumInt As Long

On Error GoTo Erro_GeracaoArqICMSProd_Grava

    lErro = Comando_Executar(lComando, "SELECT NumIntLivRegESCadProd FROM GeracaoArqICMSProd WHERE UFDest = ? AND NumIntLivRegESCadProd = ?", lNumInt, objInfoArqICMS.sUFDestino, objGeracaoArqICMSProd.lNumIntLivRegESCadProd)
    If lErro <> AD_SQL_SUCESSO Then Error 58832

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58833

    If lErro = AD_SQL_SEM_DADOS Then

        'Insere no Bd o Tipo passado (75) para nao haver repeticoes
        lErro = Comando_Executar(lComando2, "INSERT INTO GeracaoArqICMSProd (UFDest, NumIntLivRegESCadProd) VALUES (?,?)", objInfoArqICMS.sUFDestino, objGeracaoArqICMSProd.lNumIntLivRegESCadProd)
        If lErro <> AD_SQL_SUCESSO Then Error 58834

    End If

    GeracaoArqICMSProd_Grava = SUCESSO

    Exit Function

Erro_GeracaoArqICMSProd_Grava:

    GeracaoArqICMSProd_Grava = Err

    Select Case Err
        
        Case 58830, 58831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 58832, 58833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMSPROD", Err)

        Case 58834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABELA_GERACAOARQICMSPROD", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150383)

    End Select

    Exit Function

End Function

Public Function DIPJ_Preenche(ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFicha As Integer, ByVal iAnoRegulamento As Integer, lNumIntRel As Long) As Long
'Aciona rotinas que lêem Informações, calculam valores e gravam nas Tabelas sobre Entradas e Saidas de IPI
'num determinado período

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Dipj_Preenche

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 121085
        
    'obtem valor de NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 121087
    
    Select Case iAnoRegulamento
    
        Case 2007, 2008, 2009, 2010, 2011, 2012, 2013
    
            Select Case iFicha
            
                Case 21
                    lErro = DIPJ2007_GeraFicha_2122(iFilialEmpresa, 21, lNumIntRel, dtDataInicial, dtDataFinal, iAnoRegulamento - 1)
                    If lErro <> SUCESSO Then gError 121086
            
                Case 22
                    lErro = DIPJ2007_GeraFicha_2122(iFilialEmpresa, 22, lNumIntRel, dtDataInicial, dtDataFinal, iAnoRegulamento - 1)
                    If lErro <> SUCESSO Then gError 121086
            
                Case 23
                    lErro = DIPJ2007_GeraFicha_2325(iFilialEmpresa, 23, lNumIntRel, dtDataInicial, dtDataFinal, iAnoRegulamento - 1)
                    If lErro <> SUCESSO Then gError 121086
            
                Case 24
                    lErro = DIPJ2007_GeraFicha_2426(iFilialEmpresa, 24, lNumIntRel, dtDataInicial, dtDataFinal, iAnoRegulamento - 1)
                    If lErro <> SUCESSO Then gError 121086
            
                Case 25
                    lErro = DIPJ2007_GeraFicha_2325(iFilialEmpresa, 25, lNumIntRel, dtDataInicial, dtDataFinal, iAnoRegulamento - 1)
                    If lErro <> SUCESSO Then gError 121086
            
                Case 26
                    lErro = DIPJ2007_GeraFicha_2426(iFilialEmpresa, 26, lNumIntRel, dtDataInicial, dtDataFinal, iAnoRegulamento - 1)
                    If lErro <> SUCESSO Then gError 121086
            
                Case Else
                    gError 121099
                
'                Case 25
'                    'Preenche tabela DIPJ_25_27 com registros para preencher as fichas 25 e 27
'                    lErro = DIPJ_GeraFichas_25_27(iFilialEmpresa, lNumIntRel, dtDataInicial, dtDataFinal)
'                    If lErro <> SUCESSO Then gError 121092
'
'                Case 26
'                    'Preenche tabela DIPJ_26_28 com registros para preencher as fichas 26 e 28
'                    lErro = DIPJ_GeraFichas_26_28(iFilialEmpresa, lNumIntRel, dtDataInicial, dtDataFinal)
'                    If lErro <> SUCESSO Then gError 121093
'
'                Case 30
'                    'Preenche tabela DIPJ_23_24 com registros para preencher as fichas 30 e 31
'                    lErro = DIPJ_GeraFichas_30_31(iFilialEmpresa, lNumIntRel, dtDataInicial, dtDataFinal)
'                    If lErro <> SUCESSO Then gError 121093
'
'                Case 32
'                    'Preenche tabela DIPJ_25_27 com registros para preencher as fichas 32 e 34
'                    lErro = DIPJ_GeraFichas_32_34(iFilialEmpresa, lNumIntRel, dtDataInicial, dtDataFinal)
'                    If lErro <> SUCESSO Then gError 121093
'
'                Case 33
'                    'Preenche tabela DIPJ_26_28 com registros para preencher as fichas 33 e 35
'                    lErro = DIPJ_GeraFichas_33_35(iFilialEmpresa, lNumIntRel, dtDataInicial, dtDataFinal)
'                    If lErro <> SUCESSO Then gError 121093
            
            End Select
    
        Case Else
            gError 121099
    
    End Select
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 121100
        
    Exit Function
    
    DIPJ_Preenche = SUCESSO

    Exit Function

Erro_Dipj_Preenche:

    DIPJ_Preenche = gErr

    Select Case gErr

        Case 121086, 121087
        
        Case 121099
            Call Rotina_Erro(vbOKOnly, "ERRO_DIPJ_PARAMETRO_INVALIDO", gErr)
        
        Case 121085
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 121100
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150384)

    End Select
       
    Call Transacao_Rollback

    Exit Function

End Function

Function DIPJ_GeraFichas_23_24(ByVal iFilialEmpresa As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Preenche tabela DIPJ_23_24 com registros para preencher as fichas 23 e 24

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ_GeraFichas_23_24
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 23 e 24 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_23_24 (NumIntRel, Ficha, IPIBase, IPIOutros, IPIValor) (SELECT ?, GrupoDIPJ2324, SUM(ValorIPIBase), SUM(ValorIPIIsentoNaoTrib+ValorIPIOutras), SUM(ValorIPI) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 group by grupodipj2324 )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ_GeraFichas_23_24 = SUCESSO

    Exit Function

Erro_DIPJ_GeraFichas_23_24:

    DIPJ_GeraFichas_23_24 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150385)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function DIPJ_GeraFichas_25_27(ByVal iFilialEmpresa As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Preenche tabela DIPJ_25_27 com registros para preencher as fichas 25 e 27

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ_GeraFichas_25_27
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 25 e 27 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 25, CGC, SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj2527 = 25 group by cgc )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 25, '00000000000000', SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj2527 = 26 )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 27, CGC, SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj2527 = 27 group by cgc )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 27, '00000000000000', SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj2527 = 28 )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ_GeraFichas_25_27 = SUCESSO

    Exit Function

Erro_DIPJ_GeraFichas_25_27:

    DIPJ_GeraFichas_25_27 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150386)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function DIPJ_GeraFichas_26_28(ByVal iFilialEmpresa As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Preenche tabela DIPJ_26_28 com registros para preencher as fichas 26 e 28

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ_GeraFichas_26_28
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 26 e 28 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_26_28 (NumIntRel, Ficha, IPICodigo, Nome, Valor) (SELECT ?, 26 ficha, IPICodigo, Produtos.Descricao, SUM(ValorProduto) Valor FROM LivRegES, LivRegESItemNF, LivRegESCadProd, Produtos, NaturezaOP WHERE " & _
        " LivRegES.NumIntDoc = LivRegESItemNF.NumIntDocRegES AND LivRegESItemNF.NumIntCadProd = LivRegESCadProd.NumIntDoc AND LivRegESItemNF.CFOP = NaturezaOP.Codigo AND Grupodipj2628 = 26 AND LivRegESCadProd.Produto = Produtos.Codigo AND LivRegES.Data BETWEEN ? AND ? AND LivRegES.FilialEmpresa = ? and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 GROUP BY IPICodigo, Produtos.Descricao)", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_26_28 (NumIntRel, Ficha, IPICodigo, Nome, Valor) (SELECT ?, 28 ficha, IPICodigo, Produtos.Descricao, SUM(ValorProduto) Valor FROM LivRegES, LivRegESItemNF, LivRegESCadProd, Produtos, NaturezaOP WHERE " & _
        " LivRegES.NumIntDoc = LivRegESItemNF.NumIntDocRegES AND LivRegESItemNF.NumIntCadProd = LivRegESCadProd.NumIntDoc AND LivRegESItemNF.CFOP = NaturezaOP.Codigo AND Grupodipj2628 = 28 AND LivRegESCadProd.Produto = Produtos.Codigo AND LivRegES.Data BETWEEN ? AND ? AND LivRegES.FilialEmpresa = ? and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 GROUP BY IPICodigo, Produtos.Descricao)", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ_GeraFichas_26_28 = SUCESSO

    Exit Function

Erro_DIPJ_GeraFichas_26_28:

    DIPJ_GeraFichas_26_28 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150387)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Public Function InventarioTercProd_Grava_Em_Trans(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'grava os registro passados por parametro na tabela InventarioTercProd
'APENAS EM TRANSAÇÃO ABERTA

Dim lErro As Long
Dim lComando As Long
Dim objInventarioTercProd As ClassInventarioTercProd

On Error GoTo Erro_InventarioTercProd_Grava_Em_Trans

    'abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 119716
        
    'exclui os registros da tabela InvnetarioTercProd p/ inserir depois
    lErro = CF("InventarioTercProd_Exclui_Em_Trans", objInventarioTerc)
    If lErro <> SUCESSO And lErro <> 119724 Then gError 119718

    'Para cada inventariotercprod carregado na collection
    For Each objInventarioTercProd In objInventarioTerc.colInventarioTercProd
        
        'insere os novos registros na tabela InventarioTercProd
        lErro = Comando_Executar(lComando, "INSERT INTO InventarioTercProd(FilialEmpresa, TipoTerc, CodTerc, FilialTerc, Data, CodEscaninho, Produto, QuantTotal) VALUES (?,?,?,?,?,?,?,?)", _
        objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho, objInventarioTercProd.sProduto, objInventarioTercProd.dQuantTotal)
        If lErro <> AD_SQL_SUCESSO Then gError 119719
               
    Next
     
    'fecha o comando
    Call Comando_Fechar(lComando)
        
    InventarioTercProd_Grava_Em_Trans = SUCESSO

    Exit Function

Erro_InventarioTercProd_Grava_Em_Trans:

    InventarioTercProd_Grava_Em_Trans = gErr

    Select Case gErr
        
        Case 119718
        
        Case 119716
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 119719
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INVENTARIOTERCPROD", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150388)
                
    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function InventarioTercProd_Exclui_Em_Trans(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'exclui os registros da tabela InventarioTercProd através das chaves passadas como parametro
'funciona apenas EM TRANSNÇÂO ABERTA

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iFilialEmpresa As Integer
Dim iIndice As Integer

On Error GoTo Erro_InventarioTercProd_Exclui_Em_Trans
    
    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 119721
    Next

    'Verifica se existe o registro em inventariotercprod
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM InventarioTercProd WHERE FilialEmpresa = ? AND TipoTerc = ? AND CodTerc = ? AND FilialTerc = ? AND Data = ? AND CodEscaninho = ?", 0, iFilialEmpresa, objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho)
    If lErro <> AD_SQL_SUCESSO Then gError 119722

    'busca o 1º registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119723

    'sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119724

    'enquanto houver reggistros
    Do While lErro = SUCESSO
    
        'Faz o lock exclusivo do registro
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 119730
    
        'Exclui o registro encontrado
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InventarioTercProd", alComando(0))
        If lErro <> SUCESSO Then gError 119731
    
        'busca o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119732
    
    Loop
    
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InventarioTercProd_Exclui_Em_Trans = SUCESSO

    Exit Function

Erro_InventarioTercProd_Exclui_Em_Trans:

    InventarioTercProd_Exclui_Em_Trans = gErr

    Select Case gErr
    
        Case 119721
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)
    
        Case 119722, 119723, 119732
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOTERCPROD", gErr)
    
        Case 119724 'sem dados
    
        Case 119730
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_LOCK_EXCLUSIVE", gErr)
    
        Case 119731
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUIR_INVENTARIOTERCPROD", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150389)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function InventarioTerc_Exclui(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'exclui os registros das tabelas inventarioterc e inventariotercprod (relação de produtos com terceiros)

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_InventarioTerc_Exclui

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 119749

    'exclui o registro de 3º na tabela inventarioterc
    lErro = CF("InventarioTerc_Exclui_Em_Trans", objInventarioTerc)
    If lErro <> SUCESSO And lErro <> 119763 Then gError 119774

    'erro sem dados
    If lErro = 19763 Then gError 119778

    'exclui o(s) registro(s) da tabela InventariTercProd
    lErro = CF("InventarioTercProd_Exclui_Em_Trans", objInventarioTerc)
    If lErro <> SUCESSO And lErro <> 119724 Then gError 119750

    'erro sem dados
    If lErro = 119724 Then gError 119751

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 119782
    
    InventarioTerc_Exclui = SUCESSO
    
    Exit Function

Erro_InventarioTerc_Exclui:

    InventarioTerc_Exclui = gErr

    Select Case gErr
    
        Case 119749
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 119750, 119774
        
        Case 19778
            Call Rotina_Erro(vbOKOnly, "ERRO_INVENTARIOTERC_NAO_ENCONTRADO", gErr)
        
        Case 119751
            Call Rotina_Erro(vbOKOnly, "ERRO_INVENTARIOTERCPROD_NAO_ENCONTRADO", gErr)
        
        Case 119782
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150390)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Function InventarioTerc_Exclui_Em_Trans(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'exclui os registros da tabela InventarioTerc através das chaves passadas como parametro
'APENNAS EM TRANSACAO ABERTA

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iFilialEmpresa As Integer
Dim iIndice As Integer

On Error GoTo Erro_InventarioTerc_Exclui_Em_Trans
    
    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 119760
    Next

    'Verifica se existe o registro em inventarioterc
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM InventarioTerc WHERE FilialEmpresa = ? AND TipoTerc = ? AND CodTerc = ? AND FilialTerc = ? AND Data = ? AND CodEscaninho = ?", 0, iFilialEmpresa, objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho)
    If lErro <> AD_SQL_SUCESSO Then gError 119761

    'busca o 1º registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119762

    'sem dados
    If lErro = AD_SQL_SEM_DADOS Then gError 119763

    'faz um lock exclusivo do registro
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 119773

    'Exclui o registro encontrado
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InventarioTerc", alComando(0))
    If lErro <> SUCESSO Then gError 119764
        
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InventarioTerc_Exclui_Em_Trans = SUCESSO

    Exit Function

Erro_InventarioTerc_Exclui_Em_Trans:

    InventarioTerc_Exclui_Em_Trans = gErr

    Select Case gErr
    
        Case 119760
            Call Rotina_Erro(vbOKOnly, "ERRO_COMANDO_ABRIR", gErr)
    
        Case 119762, 119761
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOTERC", gErr)
    
        Case 119763
    
        Case 119773
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXCLUSIVE", gErr)
            
        Case 119764
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUIR_INVENTARIOTERC", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150391)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function InventarioTerc_Grava(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'grava os registros de 3ºs em inventarioterc e inventariotercprod

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_InventarioTerc_Grava

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 119765

    'locka os registros das tabelas relacionadas
    lErro = CF("InventarioTercProd_Lock_ObjetosVinculados", objInventarioTerc)
    If lErro <> SUCESSO Then gError 119717

    '??? acertar dic e colocar CF
    'inclui o registro de 3º na tabela inventarioterc
    lErro = InventarioTerc_Grava_Em_Trans(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119779

    '??? acertar dic e colocar CF
    'inclui o(s) registro(s) da tabela InventariTercProd
    lErro = InventarioTercProd_Grava_Em_Trans(objInventarioTerc)
    If lErro <> SUCESSO Then gError 119780

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 119781
    
    InventarioTerc_Grava = SUCESSO
    
    Exit Function

Erro_InventarioTerc_Grava:

    InventarioTerc_Grava = gErr

    Select Case gErr
    
        Case 119765
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 119780, 119779, 119717
        
        Case 119781
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150392)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Function InventarioTerc_Grava_Em_Trans(ByVal objInventarioTerc As ClassInventarioTerc) As Long
'grava os registro passados por parametro na tabela InventarioTerc
'APENAS EM TRANSAÇÃO ABERTA

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iFilial As Integer
Dim iIndice As Integer

On Error GoTo Erro_InventarioTerc_Grava_Em_Trans

    'abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 119766
    Next
        
    'procura o registro em InventarioTerc a partir das chaves passadas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM InventarioTerc WHERE FilialEmpresa = ? AND TipoTerc = ? AND CodTerc = ? AND FilialTerc = ? AND Data = ? AND CodEscaninho = ?", 0, iFilial, objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho)
    If lErro <> AD_SQL_SUCESSO Then gError 119767

    'busca o 1º registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119768
    
    'se encontrou registros
    If lErro = AD_SQL_SUCESSO Then
        
        'faz um lock no registro encontrado p/ evitar sua exclusão
        'os registros serão utiliaados em outra rotina
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 119748
    
    Else
        
        'insere os novos registros na tabela InventarioTerc
        lErro = Comando_Executar(alComando(1), "INSERT INTO InventarioTerc(FilialEmpresa, TipoTerc, CodTerc, FilialTerc, Data, CodEscaninho) VALUES (?,?,?,?,?,?)", _
        objInventarioTerc.iFilialEmpresa, objInventarioTerc.iTipoTerc, objInventarioTerc.lCodTerc, objInventarioTerc.iFilialTerc, objInventarioTerc.dtData, objInventarioTerc.iCodEscaninho)
        If lErro <> AD_SQL_SUCESSO Then gError 119769
 
    End If
 
    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    InventarioTerc_Grava_Em_Trans = SUCESSO

    Exit Function

Erro_InventarioTerc_Grava_Em_Trans:

    InventarioTerc_Grava_Em_Trans = gErr

    Select Case gErr
        
        Case 119766
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 119767, 119768
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOTERC", gErr)
                    
        Case 119748
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_SHARED_INVENTARIOTERC", gErr)
                    
        Case 119769
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INVENTARIOTERC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150393)
                
    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Sub RegEmitentes_Completa030703()

Dim lErro As Long, lTransacao As Long, iIndice As Integer
Dim alComando(0 To 7) As Long, lCodigo As Long
Dim lNumIntDoc As Long, lCliente As Long, iFilialCli As Integer, lFornecedor As Long, iFilialForn As Integer, sCgc As String, sInscricaoEstadual As String, sNome As String, sUF As String
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String, sCampo As String

On Error GoTo Erro_RegEmitentes_Completa030703

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103542
    Next
    
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 103553
    
    sCgc = String(255, 0)
    sInscricaoEstadual = String(255, 0)
    sNome = String(255, 0)
    sUF = String(255, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT Livreges.NumIntDoc, NFiscal.Fornecedor, NFiscal.FilialForn, Livreges.cgc, Livreges.inscricaoestadual, Livreges.nome,Livreges.uf FROM Livreges, NFiscal, TiposDocInfo WHERE NumIntEmitente = 0 AND TipoNumIntDocOrigem =0 AND Livreges.NumIntDocOrigem = NFiscal.NumIntDoc AND NFiscal.TipoNFiscal = TiposDOCInfo.codigo AND TiposDOCInfo.Tipo IN (1,3) AND NFiscal.Fornecedor <> 0", _
        lNumIntDoc, lFornecedor, iFilialForn, sCgc, sInscricaoEstadual, sNome, sUF)
    If lErro <> AD_SQL_SUCESSO Then gError 69250
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Le no BD para ver se este emitente já não está cadastrado
        lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM LivRegESEmitentes WHERE CliForn = ? AND Filial = ? AND TipoEmitente = ? AND NumIntDocLivFechado = 0", lCodigo, lFornecedor, iFilialForn, EMITENTE_FORNECEDOR)
        If lErro <> AD_SQL_SUCESSO Then gError 69249

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
        'Se não estiver então preenche o objeto
        If lErro = AD_SQL_SEM_DADOS Then
                    
            sCodigo = "FISC_NUM_PROX_REG_EMITENTES"
            sTabelaConfig = "FISConfig"
            sTabelaObjeto = "LivRegESEmitentes"
            sCampo = "Codigo"
    
            'Chama Config_ObterAutomatico
            lErro = CF("Config_ObterAutomatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
            If lErro <> SUCESSO Then gError 69255
            
            'colocar aqui o código para gravação do emitente
            lErro = Comando_Executar(alComando(3), "INSERT INTO LivRegESEmitentes (Codigo, Nome, CGC, InscricaoEstadual, UF, TipoEmitente, CliForn, Filial) VALUES (?,?,?,?,?,?,?,?)", lCodigo, sNome, sCgc, sInscricaoEstadual, sUF, EMITENTE_FORNECEDOR, lFornecedor, iFilialForn)
            If lErro <> AD_SQL_SUCESSO Then gError 69258
               
        End If
        
        lErro = Comando_Executar(alComando(4), "UPDATE LivRegES SET NumIntEmitente = ? WHERE NumIntDoc = ?", lCodigo, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69258
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
    Loop
    
    sCgc = String(255, 0)
    sInscricaoEstadual = String(255, 0)
    sNome = String(255, 0)
    sUF = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Livreges.NumIntDoc, NFiscal.Cliente, NFiscal.FilialCli, Livreges.cgc, Livreges.inscricaoestadual, Livreges.nome,Livreges.uf FROM Livreges, NFiscal, TiposDocInfo WHERE NumIntEmitente = 0 AND TipoNumIntDocOrigem =0 AND Livreges.NumIntDocOrigem = NFiscal.NumIntDoc AND NFiscal.TipoNFiscal = TiposDOCInfo.codigo AND TiposDOCInfo.Tipo IN (1,3) AND NFiscal.Cliente <> 0", _
        lNumIntDoc, lCliente, iFilialCli, sCgc, sInscricaoEstadual, sNome, sUF)

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Le no BD para ver se este emitente já não está cadastrado
        lErro = Comando_Executar(alComando(5), "SELECT Codigo FROM LivRegESEmitentes WHERE CliForn = ? AND Filial = ? AND TipoEmitente = ? AND NumIntDocLivFechado = 0", lCodigo, lCliente, iFilialCli, EMITENTE_CLIENTE)
        If lErro <> AD_SQL_SUCESSO Then gError 69249

        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
        'Se não estiver então preenche o objeto
        If lErro = AD_SQL_SEM_DADOS Then
                    
            sCodigo = "FISC_NUM_PROX_REG_EMITENTES"
            sTabelaConfig = "FISConfig"
            sTabelaObjeto = "LivRegESEmitentes"
            sCampo = "Codigo"
    
            'Chama Config_ObterAutomatico
            lErro = CF("Config_ObterAutomatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
            If lErro <> SUCESSO Then gError 69255
            
            'colocar aqui o código para gravação do emitente
            lErro = Comando_Executar(alComando(6), "INSERT INTO LivRegESEmitentes (Codigo, Nome, CGC, InscricaoEstadual, UF, TipoEmitente, CliForn, Filial) VALUES (?,?,?,?,?,?,?,?)", lCodigo, sNome, sCgc, sInscricaoEstadual, sUF, EMITENTE_CLIENTE, lCliente, iFilialCli)
            If lErro <> AD_SQL_SUCESSO Then gError 69258
               
        End If
        
        lErro = Comando_Executar(alComando(7), "UPDATE LivRegES SET NumIntEmitente = ? WHERE NumIntDoc = ?", lCodigo, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 69258
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69250
    
    Loop
    
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 103564
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MsgBox ("SUCESSO")
    
    Exit Sub
    
Erro_RegEmitentes_Completa030703:

    MsgBox ("ERRO")
    
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150394)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Sub

End Sub

Public Function GuiaICMS_Grava(ByVal objGuiasICMS As ClassGuiasICMS) As Long

Dim lErro As Long, lTransacao As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim sNumero As String
Dim dtData As Date
Dim iFilialEmpresa As Integer
Dim sTabela As String

On Error GoTo Erro_GuiaICMS_Grava

    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 125237
    Next

    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 124073
    
    If objGuiasICMS.iICMSST = MARCADO Then
        sTabela = "GuiasICMSST"
    Else
        sTabela = "GuiasICMS"
    End If
    
    sNumero = String(STRING_GUIASICMS_NUMERO, 0)
    
    If sTabela = "GuiasICMSST" Then
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Numero, Data, FilialEmpresa FROM " & sTabela & " WHERE Numero = ? AND Data = ? AND FilialEmpresa = ? AND UF = ?", 0, sNumero, dtData, iFilialEmpresa, objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa, objGuiasICMS.sUF)
    Else
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Numero, Data, FilialEmpresa FROM " & sTabela & " WHERE Numero = ? AND Data = ? AND FilialEmpresa = ?", 0, sNumero, dtData, iFilialEmpresa, objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 125238
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125239
    
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 125240
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE " & sTabela & " SET DataEntrega = ?, LocalEntrega = ?, Valor = ?, OrgaoArrecadador = ?, ApuracaoDe = ?, ApuracaoAte = ?,CodReceita = ?,CodObrigRecolher = ?,Vencimento = ?", alComando(1), objGuiasICMS.dtDataEntrega, objGuiasICMS.sLocalEntrega, objGuiasICMS.dValor, objGuiasICMS.sOrgaoArrecadador, objGuiasICMS.dtApuracaoDe, objGuiasICMS.dtApuracaoAte, objGuiasICMS.sCodReceita, objGuiasICMS.sCodObrigRecolher, objGuiasICMS.dtVencimento)
        If lErro <> AD_SQL_SUCESSO Then gError 125241
        
    Else
    
        If sTabela = "GuiasICMSST" Then
            lErro = Comando_Executar(alComando(3), "INSERT INTO " & sTabela & " (Numero, Data, FilialEmpresa, DataEntrega, Valor, LocalEntrega, OrgaoArrecadador, ApuracaoDe, ApuracaoAte,CodReceita,CodObrigRecolher,Vencimento, UF) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtDataEntrega, objGuiasICMS.dValor, objGuiasICMS.sLocalEntrega, objGuiasICMS.sOrgaoArrecadador, objGuiasICMS.dtApuracaoDe, objGuiasICMS.dtApuracaoAte, objGuiasICMS.sCodReceita, objGuiasICMS.sCodObrigRecolher, objGuiasICMS.dtVencimento, objGuiasICMS.sUF)
        Else
            lErro = Comando_Executar(alComando(3), "INSERT INTO " & sTabela & " (Numero, Data, FilialEmpresa, DataEntrega, Valor, LocalEntrega, OrgaoArrecadador, ApuracaoDe, ApuracaoAte,CodReceita,CodObrigRecolher,Vencimento) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtDataEntrega, objGuiasICMS.dValor, objGuiasICMS.sLocalEntrega, objGuiasICMS.sOrgaoArrecadador, objGuiasICMS.dtApuracaoDe, objGuiasICMS.dtApuracaoAte, objGuiasICMS.sCodReceita, objGuiasICMS.sCodObrigRecolher, objGuiasICMS.dtVencimento)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 125242
        
    End If
    
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 124074
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    GuiaICMS_Grava = SUCESSO
    
    Exit Function
    
Erro_GuiaICMS_Grava:

    GuiaICMS_Grava = gErr
    
    Select Case gErr
    
        Case 124073
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 124074
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
        
        Case 125237
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 125238, 125239
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIASICMS", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
        
        Case 125240
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_GUIASICMS", gErr)
        
        Case 125241
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GUIASICMS", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
            
        Case 125242
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GUIASICMS", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150395)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Public Function GuiaICMS_Exclui(ByVal objGuiasICMS As ClassGuiasICMS) As Long

Dim lErro As Long, lTransacao As Long
Dim alComando(1 To 3)  As Long
Dim iIndice As Integer
Dim sNumero As String
Dim dtData As Date
Dim iFilialEmpresa As Integer
Dim sTabela As String

On Error GoTo Erro_GuiaICMS_Exclui

    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 125243
    Next
    
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 124075
    
    If objGuiasICMS.iICMSST = MARCADO Then
        sTabela = "GuiasICMSST"
    Else
        sTabela = "GuiasICMS"
    End If
    
    sNumero = String(STRING_GUIASICMS_NUMERO, 0)

    If sTabela = "GuiasICMSST" Then
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Numero, Data, FilialEmpresa FROM " & sTabela & " WHERE Numero = ? AND Data = ? AND FilialEmpresa = ? AND UF = ?", 0, sNumero, dtData, iFilialEmpresa, objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa, objGuiasICMS.sUF)
    Else
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Numero, Data, FilialEmpresa FROM " & sTabela & " WHERE Numero = ? AND Data = ? AND FilialEmpresa = ?", 0, sNumero, dtData, iFilialEmpresa, objGuiasICMS.sNumero, objGuiasICMS.dtData, objGuiasICMS.iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 125244
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125245
    
    If lErro = AD_SQL_SEM_DADOS Then gError 125246
    
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 125247
    
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM " & sTabela, alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 125248
    
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 124076
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    GuiaICMS_Exclui = SUCESSO
    
    Exit Function
    
Erro_GuiaICMS_Exclui:

    GuiaICMS_Exclui = gErr
    
    Select Case gErr
    
        Case 124075
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 124076
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
        
        Case 125243
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 125244, 125245
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GUIASICMS", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
        
        Case 125246 'registro não encontrado
        
        Case 125247
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_GUIASICMS", gErr)
            
        Case 125248
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GUIASICMS", gErr, objGuiasICMS.iFilialEmpresa, objGuiasICMS.dtData, objGuiasICMS.sNumero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150396)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Function NF_DescGlobal_Rateia(ByVal objNFiscal As ClassNFiscal) As Long
'faz o rateio do desconto global à nf

Dim lErro As Long, dValorBrutoItens As Double, dValorDescGlobal As Double, dFator As Double
Dim objItemNF As New ClassItemNF, objLivRegESItemNF As ClassLivRegESItemNF, dDescAux As Double
Dim objLivRegESLinha As ClassLivRegESLinha, dSaldoDesc As Double, iIndice As Integer
Dim dValorDescAcum As Double, dDif As Double

On Error GoTo Erro_NF_DescGlobal_Rateia

    dValorDescGlobal = objNFiscal.dValorDesconto
    
    For Each objItemNF In objNFiscal.ColItensNF
    
        dValorBrutoItens = dValorBrutoItens + objItemNF.dValorTotal
    
    Next
        
    If dValorBrutoItens <> 0 Then
    
        For Each objItemNF In objNFiscal.ColItensNF
        
            dDescAux = Int(100 * objItemNF.dValorTotal * dValorDescGlobal / dValorBrutoItens) / 100
                
            objItemNF.dValorDescGlobal = dDescAux
            dValorDescAcum = dValorDescAcum + dDescAux
                    
'''            If objItemNF.objTributacaoItemNF.dICMSBase <> 0 Then
'''
'''                objItemNF.objTributacaoItemNF.dICMSBase = objItemNF.objTributacaoItemNF.dICMSBase - dDescAux
'''                objItemNF.objTributacaoItemNF.dICMSValor = CDbl(Format(objItemNF.objTributacaoItemNF.dICMSBase * objItemNF.objTributacaoItemNF.dICMSAliquota, "0.0000"))
'''
'''            End If
        
        Next
    
        dDif = dValorDescGlobal - dValorDescAcum
        
        Do While dDif >= DELTA_VALORMONETARIO
        
            For Each objItemNF In objNFiscal.ColItensNF
            
                objItemNF.dValorDescGlobal = objItemNF.dValorDescGlobal + 0.01
                
'''                If objItemNF.objTributacaoItemNF.dICMSBase <> 0 Then
'''
'''                    objItemNF.objTributacaoItemNF.dICMSBase = objItemNF.objTributacaoItemNF.dICMSBase - 0.01
'''                    objItemNF.objTributacaoItemNF.dICMSValor = CDbl(Format(objItemNF.objTributacaoItemNF.dICMSBase * objItemNF.objTributacaoItemNF.dICMSAliquota, "0.0000"))
'''
'''                End If
                
                dDif = dDif - 0.01
                If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
                
            Next
            
        Loop
        
    End If
    
    NF_DescGlobal_Rateia = SUCESSO
        
    Exit Function
    
Erro_NF_DescGlobal_Rateia:

    NF_DescGlobal_Rateia = gErr
    
    Select Case gErr
        
        Case 81501
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150397)

    End Select
    
    Exit Function

End Function

Function DIPJ_GeraFichas_30_31(ByVal iFilialEmpresa As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Preenche tabela DIPJ_23_24 com registros para preencher as fichas 30 e 31

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ_GeraFichas_30_31
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 30 e 31 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_23_24 (NumIntRel, Ficha, IPIBase, IPIOutros, IPIValor) (SELECT ?, GrupoDIPJ3031, SUM(ValorIPIBase), SUM(ValorIPIIsentoNaoTrib+ValorIPIOutras), SUM(ValorIPI) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and GrupoDIPJ3031 <> 0 group by GrupoDIPJ3031 )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ_GeraFichas_30_31 = SUCESSO

    Exit Function

Erro_DIPJ_GeraFichas_30_31:

    DIPJ_GeraFichas_30_31 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150398)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function DIPJ_GeraFichas_32_34(ByVal iFilialEmpresa As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Preenche tabela DIPJ_25_27 com registros para preencher as fichas 32 e 34

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ_GeraFichas_32_34
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 32 e 34 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 32, CGC, SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj3234 = 32 group by cgc )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 32, '00000000000000', SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj3234 = 33 )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, Valor) (SELECT ?, 34, CGC, SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, Naturezaop WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and naturezaop.codigo = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and grupodipj3234 = 34 group by cgc )", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ_GeraFichas_32_34 = SUCESSO

    Exit Function

Erro_DIPJ_GeraFichas_32_34:

    DIPJ_GeraFichas_32_34 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150399)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function DIPJ_GeraFichas_33_35(ByVal iFilialEmpresa As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Preenche tabela DIPJ_26_28 com registros para preencher as fichas 33 e 35

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ_GeraFichas_33_35
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 33 e 35 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_26_28 (NumIntRel, Ficha, IPICodigo, Nome, Valor) (SELECT ?, 33 ficha, IPICodigo, Produtos.Descricao, SUM(ValorProduto) Valor FROM LivRegES, LivRegESItemNF, LivRegESCadProd, Produtos, NaturezaOP WHERE " & _
        " LivRegES.NumIntDoc = LivRegESItemNF.NumIntDocRegES AND LivRegESItemNF.NumIntCadProd = LivRegESCadProd.NumIntDoc AND LivRegESItemNF.CFOP = NaturezaOP.Codigo AND grupodipj3335 = 33 AND LivRegESCadProd.Produto = Produtos.Codigo AND LivRegES.Data BETWEEN ? AND ? AND LivRegES.FilialEmpresa = ? and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 GROUP BY IPICodigo, Produtos.Descricao)", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_26_28 (NumIntRel, Ficha, IPICodigo, Nome, Valor) (SELECT ?, 35 ficha, IPICodigo, Produtos.Descricao, SUM(ValorProduto) Valor FROM LivRegES, LivRegESItemNF, LivRegESCadProd, Produtos, NaturezaOP WHERE " & _
        " LivRegES.NumIntDoc = LivRegESItemNF.NumIntDocRegES AND LivRegESItemNF.NumIntCadProd = LivRegESCadProd.NumIntDoc AND LivRegESItemNF.CFOP = NaturezaOP.Codigo AND grupodipj3335 = 35 AND LivRegESCadProd.Produto = Produtos.Codigo AND LivRegES.Data BETWEEN ? AND ? AND LivRegES.FilialEmpresa = ? and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 GROUP BY IPICodigo, Produtos.Descricao)", lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ_GeraFichas_33_35 = SUCESSO

    Exit Function

Erro_DIPJ_GeraFichas_33_35:

    DIPJ_GeraFichas_33_35 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150400)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function


'############################################################################
'Inserido por Wagner
Function GeracaoArqICMSProdCupom_Grava(ByVal sProduto As String, ByVal sUFDest As String, objItemNF As ClassItemNF, ByVal dtData As Date, ByVal lComando As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim iAux As Integer
Dim lNumIntDoc As Long
Dim objLivRegESCadProd As New ClassLivRegESCadProd

On Error GoTo Erro_GeracaoArqICMSProdCupom_Grava

    lErro = Comando_Executar(lComando, "SELECT 1 FROM GeracaoArqICMSProd AS G, LivRegESCadProd AS L WHERE L.NumIntDoc = G.NumIntLivRegESCadProd AND L.Produto = ? AND G.UFDest = ?", iAux, sProduto, sUFDest)
    If lErro <> AD_SQL_SUCESSO Then gError 138390

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138391

    If lErro = AD_SQL_SEM_DADOS Then

        lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM LivRegESCadProd AS L WHERE L.Produto = ? ORDER BY L.DataInicial DESC", lNumIntDoc, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 138392

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138393

        If lErro = AD_SQL_SEM_DADOS Then
                
            lErro = NF_Grava_LivRegESCadProd(dtData, objItemNF, objLivRegESCadProd)
            If lErro <> SUCESSO Then gError 138397
        
            lNumIntDoc = objLivRegESCadProd.lNumIntDoc
        
        End If

        'Insere no Bd o Tipo passado (75) para nao haver repeticoes
        lErro = Comando_Executar(lComando2, "INSERT INTO GeracaoArqICMSProd (UFDest, NumIntLivRegESCadProd) VALUES (?,?)", sUFDest, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 138394

    End If

    GeracaoArqICMSProdCupom_Grava = SUCESSO

    Exit Function

Erro_GeracaoArqICMSProdCupom_Grava:

    GeracaoArqICMSProdCupom_Grava = gErr

    Select Case gErr
        
        Case 138390 To 138393
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMSPROD", gErr)

        Case 138394
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABELA_GERACAOARQICMSPROD", gErr)

        Case 138397

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150401)

    End Select

    Exit Function

End Function

'##############################################################
'INSERIDO POR WAGNER PARA REABRIR LIVROS FISCAIS
Public Function LivroFilial_Reabre_EmTrans(ByVal objLivroFilial As ClassLivrosFilial) As Long
'Grava um Livro Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim iAux As Integer

On Error GoTo Erro_LivroFilial_Reabre_EmTrans

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131820
    Next
       
    'Verifica se o Livro Filial está cadastrado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodLivro FROM LivrosFilial WHERE CodLivro = ? AND FilialEmpresa = ?", 0, iAux, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 131822

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131823
    
    If lErro = AD_SQL_SUCESSO Then
                     
        'Faz "lock" em LivrosFilial
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 131824

        'Atualiza dados do Livro da Filial
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE LivrosFilial SET DataInicial = ?, DataFinal = ?", alComando(0), objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 131825

        lErro = LivroFilial_Exclui_Fechados(objLivroFilial)
        If lErro <> SUCESSO Then gError 131857
                
        'Se o Livro for de Apuração de ICMS
        If objLivroFilial.iCodLivro = LIVRO_APURACAO_ICMS_CODIGO Then
        
            lErro = LivroFilial_ReabreICMS(objLivroFilial)
            If lErro <> SUCESSO Then gError 131850

        'Se o Livro for de Apuração de IPI
        ElseIf objLivroFilial.iCodLivro = LIVRO_APURACAO_IPI_CODIGO Then
            
            lErro = LivroFilial_ReabreIPI(objLivroFilial)
            If lErro <> SUCESSO Then gError 131852
                        
        'Se o Livro for de Registro de Inventário
        ElseIf objLivroFilial.iCodLivro = LIVRO_REG_INVENTARIO_CODIGO Then
            
            lErro = LivroFilial_ReabreINV(objLivroFilial)
            If lErro <> SUCESSO Then gError 131853
                    
        End If
    
    End If

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LivroFilial_Reabre_EmTrans = SUCESSO

    Exit Function

Erro_LivroFilial_Reabre_EmTrans:

    LivroFilial_Reabre_EmTrans = gErr

    Select Case gErr

        Case 131820
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131822, 131823
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFILIAL", gErr)

        Case 131824
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 131825
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVROSFILIAL", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa)

        Case 131850, 131852, 131853, 131857
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150402)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function LivroFilial_ReabreICMS(ByVal objLivroFilial As ClassLivrosFilial) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_LivroFilial_ReabreICMS

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131849
    Next

    'Verifica o Livro Fiscal está vinculado com uma Apuração ICMS
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM RegApuracaoICMS WHERE DataInicial >= ? AND DataFinal <= ? AND FilialEmpresa = ? ", 0, lNumIntDoc, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 131830
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131831
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 131844
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoICMS SET NumIntDocLivFechado = ?", alComando(2), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 131832
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131833
    
    Loop

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LivroFilial_ReabreICMS = SUCESSO

    Exit Function

Erro_LivroFilial_ReabreICMS:

    LivroFilial_ReabreICMS = gErr

    Select Case gErr
            
        Case 131830, 131831, 131833
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOICMS", gErr)
                
        Case 131832
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOICMS", gErr)
                    
        Case 131844
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOICMS", gErr, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa)
                    
        Case 131849
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150403)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function LivroFilial_ReabreIPI(ByVal objLivroFilial As ClassLivrosFilial) As Long

Dim lErro As Long

Dim lNumIntDoc As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_LivroFilial_ReabreIPI

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131851
    Next

    'Verifica se o Livro Fiscal está vinculado com uma Apuração IPI
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM RegApuracaoIPI WHERE DataInicial >= ? AND DataFinal <= ? AND FilialEmpresa = ? ", 0, lNumIntDoc, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 131834
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131835
    
    Do While lErro = AD_SQL_SUCESSO
        
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 131845
       
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoIPI SET NumIntDocLivFechado = ?", alComando(2), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 131836
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131837
    
    Loop
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LivroFilial_ReabreIPI = SUCESSO

    Exit Function

Erro_LivroFilial_ReabreIPI:

    LivroFilial_ReabreIPI = gErr

    Select Case gErr
            
        Case 131834, 131835, 131837
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPI", gErr)
        
        Case 131836
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGAPURACAOIPI", gErr)
        
        Case 131845
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGAPURACAOIPI", gErr, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa)
        
        Case 131851
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150404)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function LivroFilial_ReabreINV(ByVal objLivroFilial As ClassLivrosFilial) As Long

Dim lErro As Long
Dim dtData As Date
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_LivroFilial_ReabreINV

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131854
    Next

    'Verifica se o Livro Fiscal está vinculado com uma Registro de Inventário
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Data FROM RegInventario WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? ", 0, dtData, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal, objLivroFilial.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 131838
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131839
    
    'Se encontrou vínculo
    Do While lErro = AD_SQL_SUCESSO
        
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 131846
        
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegInventario SET NumIntDocLivFechado = ?", alComando(2), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 131840
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131841
    
    Loop
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LivroFilial_ReabreINV = SUCESSO

    Exit Function

Erro_LivroFilial_ReabreINV:

    LivroFilial_ReabreINV = gErr

    Select Case gErr
            
        Case 131838, 131839, 131841
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)
            
        Case 131840
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REGINVENTARIO", gErr)
        
        Case 131846
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGINVENTARIO1", gErr, dtData)
        
        Case 131854
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150405)

    End Select


    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

    
End Function


Public Function LivroFilial_Reabre(ByVal objLivroFilial As ClassLivrosFilial) As Long
'Reabre um Livro fiscal

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_LivroFilial_Reabre

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 131821
    
    lErro = CF("LivroFilial_Reabre_EmTrans", objLivroFilial)
    If lErro <> SUCESSO Then gError 131855


    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 131842
    
    LivroFilial_Reabre = SUCESSO

    Exit Function

Erro_LivroFilial_Reabre:

    LivroFilial_Reabre = gErr

    Select Case gErr
    
        Case 131821
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 131842
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 131855
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150406)

    End Select

    Call Transacao_Rollback

    Exit Function
    
End Function

Private Function LivroFilial_Exclui_Fechados(ByVal objLivroFilial As ClassLivrosFilial) As Long

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim lNumIntDoc As Long
Dim lNumIntDocLivRegES As Long
Dim iIndice As Integer

On Error GoTo Erro_LivroFilial_Exclui_Fechados

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 131856
    Next

    'Verifica Livros Fechados
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM LivrosFechados WHERE CodLivro = ? AND FilialEmpresa = ? AND DataInicial >= ? AND DataFinal <= ?", 0, lNumIntDoc, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 131826

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131827
    
    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 131843

        lErro = Comando_ExecutarPos(alComando(1), "DELETE LivrosFechados", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 131828
        
        '#######################################
        'Inserido por Wagner 06/02/2006
        'Desassocia o Livro com os registros de LivRegES
        lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc FROM LivRegES WHERE NumIntDocLivFechado = ? ", 0, lNumIntDocLivRegES, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141651
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141652
        
        Do While lErro = AD_SQL_SUCESSO
    
            lErro = Comando_LockExclusive(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 141653
    
            lErro = Comando_ExecutarPos(alComando(4), "UPDATE LivRegES SET NumIntDocLivFechado = ? ", alComando(3), 0)
            If lErro <> AD_SQL_SUCESSO Then gError 141654
        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141655
    
        Loop
        '#######################################
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131829

    Loop
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    LivroFilial_Exclui_Fechados = SUCESSO

    Exit Function

Erro_LivroFilial_Exclui_Fechados:

    LivroFilial_Exclui_Fechados = gErr

    Select Case gErr

        Case 131826, 131827, 131829
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVROSFECHADOS", gErr)
        
        Case 131828
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LIVROSFECHADOS", gErr)
                       
        Case 131843
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVROSFECHADOS", gErr, objLivroFilial.iCodLivro, objLivroFilial.iFilialEmpresa, objLivroFilial.dtDataInicial, objLivroFilial.dtDataFinal)
        
        Case 131856
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 141651, 141652, 141655
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
            
        Case 141653
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
        
        Case 141654
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150407)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function
    
End Function

'########################################################################
'Inserido por Wagner 25/04/2006
Public Function RelRCPEMod3_Prepara(lNumIntRel As Long, ByVal objRelRCPESeleciona As ClassRelRCPESeleciona) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RelRCPEMod3_Prepara
  
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 177362
    
    lErro = RelRCPEMod3_Prepara_EmTrans(lNumIntRel, objRelRCPESeleciona)
    If lErro <> SUCESSO Then gError 177391
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 177363
  
    RelRCPEMod3_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelRCPEMod3_Prepara:
  
    RelRCPEMod3_Prepara = gErr
 
    Select Case gErr
 
        Case 177362
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 177363
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 177391
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177364)
 
    End Select

    Call Transacao_Rollback
 
    Exit Function
 
End Function

Public Function RelRCPEMod3_Prepara_EmTrans(lNumIntRel As Long, ByVal objRelRCPESeleciona As ClassRelRCPESeleciona) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objProduto As ClassProduto
Dim iSeqPag As Integer
Dim ColProduto As New Collection
Dim colRelRCPECab As New Collection

On Error GoTo Erro_RelRCPEMod3_Prepara_EmTrans
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_REL_RCPEMOD3", lNumIntRel)
    If lErro <> SUCESSO Then gError 177365
    
    'Pega os produtos filtrados por ProdutoDe e ProdutoAte
    lErro = CF("Produto_Le_Faixa", objRelRCPESeleciona.sProdutoDe, objRelRCPESeleciona.sProdutoAte, ColProduto)
    If lErro <> SUCESSO Then gError 177366
    
    For Each objProduto In ColProduto
            
        'Se possui controle de estoque prossegue
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
    
            'Obtém os principais dados do relatório em colRelRCPECab
            lErro = RelRCPEMod3_Prepara_EmTrans1(lNumIntRel, objRelRCPESeleciona, objProduto, colRelRCPECab)
            If lErro <> SUCESSO Then gError 177367
        
        End If
    
    Next
    
    'Obtém os dados calculados e grava na tabela temporária
    lErro = RelRCPEMod3_Prepara_EmTrans2(lNumIntRel, objRelRCPESeleciona.iFilialEmpresa, colRelRCPECab)
    If lErro <> SUCESSO Then gError 177368
    
    RelRCPEMod3_Prepara_EmTrans = SUCESSO
 
    Exit Function
 
Erro_RelRCPEMod3_Prepara_EmTrans:
 
    RelRCPEMod3_Prepara_EmTrans = gErr
 
    Select Case gErr
 
        Case 177365 To 177368
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177369)
 
    End Select
 
    Exit Function
 
End Function

Public Function RelRCPEMod3_Prepara_EmTrans1(ByVal lNumIntRel As Long, ByVal objRelRCPESeleciona As ClassRelRCPESeleciona, ByVal objProduto As ClassProduto, ByVal colRelRCPECab As Collection) As Long
'Aqui chegam os produtos selecionados que tem controle de estoque, falta filtrar:
'Os que tem almoxarifado para filial empresa passada
'E retirar os que não tem movimentação e nem estoque inicial
 
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iQuebra As Integer
Dim sEntradaOuSaida As String
Dim lCodigoMovEst As Long
Dim lNumIntDocMovEst As Long
Dim objRelRCPECab As New ClassRelRCPECab
Dim objRelRCPEDet As ClassRelRCPEDet
Dim objRelRCPEDetAux As ClassRelRCPEDet
Dim dTotalE As Double
Dim dTotalS As Double
Dim iQuebraAnt As Integer
Dim dFator As Double
Dim iQuebraAux As Integer
Dim bTemMov As Boolean

On Error GoTo Erro_RelRCPEMod3_Prepara_EmTrans1

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 177370
    Next
    
    Set objRelRCPECab.objEstoqueProduto = New ClassEstoqueProduto
    Set objRelRCPECab.objProdutoFilial = New ClassProdutoFilial
    
    Set objRelRCPECab.objProduto = objProduto
    objRelRCPECab.lFolha = objRelRCPESeleciona.lFolha
    objRelRCPECab.dtDataAte = objRelRCPESeleciona.dtDataAte
    objRelRCPECab.dtDataDe = objRelRCPESeleciona.dtDataDe
    
    objRelRCPECab.objEstoqueProduto.sProduto = objRelRCPECab.objProduto.sCodigo
    
    'Lê as quantidades do produtos nos almoxs da filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objRelRCPECab.objEstoqueProduto, objRelRCPESeleciona.iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 52979 Then gError 177371
    
    objRelRCPECab.objProdutoFilial.sProduto = objRelRCPECab.objProduto.sCodigo
    objRelRCPECab.objProdutoFilial.iFilialEmpresa = objRelRCPESeleciona.iFilialEmpresa

    'Le ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objRelRCPECab.objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 177372
    
    If lErro = SUCESSO Then
    
        sEntradaOuSaida = String(STRING_TIPOMOV_EST_ENTRADAOUSAIDA, 0)
        
        'Le os movimentos de estoque para o produto e período informado
        lErro = Comando_Executar(alComando(1), "SELECT Quebra, EntradaOuSaida, Codigo, NumIntDoc FROM MovsRegMod3 WHERE Data BETWEEN ? AND ? AND Produto = ? ORDER BY Quebra, Data, EntradaOuSaida, NumIntDoc", iQuebra, sEntradaOuSaida, lCodigoMovEst, lNumIntDocMovEst, objRelRCPESeleciona.dtDataDe, objRelRCPESeleciona.dtDataAte, objRelRCPECab.objProduto.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 177373
        
        'Tenta ler o primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177374
        
        iQuebraAnt = iQuebra
                
        Do While lErro = AD_SQL_SUCESSO
                
            Set objRelRCPEDet = New ClassRelRCPEDet
            Set objRelRCPEDet.objItemMovEst = New ClassItemMovEstoque
            Set objRelRCPEDet.objTipoMovEst = New ClassTipoMovEst
                        
            objRelRCPEDet.objItemMovEst.lNumIntDoc = lNumIntDocMovEst
            
            'Le o movimento de estoque associado
            lErro = CF("MovimentoEstoque_Le_NumIntDoc", alComando(2), objRelRCPEDet.objItemMovEst)
            If lErro <> SUCESSO And lErro <> 83748 Then gError 177375
        
            objRelRCPEDet.iAno = Year(objRelRCPEDet.objItemMovEst.dtData)
            objRelRCPEDet.iDia = Day(objRelRCPEDet.objItemMovEst.dtData)
            objRelRCPEDet.iMes = Month(objRelRCPEDet.objItemMovEst.dtData)
            objRelRCPEDet.iQuebra = iQuebra
            objRelRCPEDet.sEntradaOuSaida = sEntradaOuSaida
            
            'Obtém o campo código
            lErro = RelRCPEMod3_Prepara_EmTrans3(objRelRCPECab, objRelRCPEDet)
            If lErro <> SUCESSO Then gError 177376
        
            objRelRCPEDet.objTipoMovEst.iCodigo = objRelRCPEDet.objItemMovEst.iTipoMov
            
            'ler os dados referentes ao tipo de movimento
            lErro = CF("TipoMovEstoque_Le", objRelRCPEDet.objTipoMovEst)
            If lErro <> SUCESSO Then gError 177377
            
            'Se a origem for um item de Nota fiscal
            If objRelRCPEDet.objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Then
            
                Set objRelRCPEDet.objItemNF = New ClassItemNF
                Set objRelRCPEDet.objNFiscal = New ClassNFiscal
                
                objRelRCPEDet.objItemNF.lNumIntDoc = objRelRCPEDet.objItemMovEst.lNumIntDocOrigem
                
                'Le o item de nota associado ao movimento de estoque
                lErro = CF("ItemNFiscal_Le", objRelRCPEDet.objItemNF)
                If lErro <> SUCESSO And lErro <> 35225 Then gError 177378
                
                objRelRCPEDet.objNFiscal.lNumIntDoc = objRelRCPEDet.objItemNF.lNumIntNF
                
                'Le a nota que contém o item
                lErro = CF("NFiscal_Le", objRelRCPEDet.objNFiscal)
                If lErro <> SUCESSO And lErro <> 31442 Then gError 177379
                
            End If
            
            'Fator de conv. de UMs
            lErro = CF("UM_Conversao", objRelRCPECab.objProduto.iClasseUM, objRelRCPEDet.objItemMovEst.sSiglaUM, objRelRCPECab.objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 177380
            
            objRelRCPEDet.objItemMovEst.dQuantidadeEst = objRelRCPEDet.objItemMovEst.dQuantidade * dFator
            
            If sEntradaOuSaida = "E" Then
                dTotalE = dTotalE + objRelRCPEDet.objItemMovEst.dQuantidadeEst
            Else
                dTotalS = dTotalS + objRelRCPEDet.objItemMovEst.dQuantidadeEst
            End If
        
            objRelRCPECab.ColDetalhe.Add objRelRCPEDet
       
            'Le o próximo registro
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177381
    
            'Se mudou a quebra ou foi o último
            If iQuebraAnt <> iQuebra Or lErro = AD_SQL_SEM_DADOS Then
                
                For Each objRelRCPEDetAux In objRelRCPECab.ColDetalhe
                    If objRelRCPEDetAux.iQuebra = iQuebraAnt Then
                        objRelRCPEDetAux.dTotalEntrada = dTotalE
                        objRelRCPEDetAux.dTotalSaida = dTotalS
                    End If
                Next
                dTotalS = 0
                dTotalE = 0
                iQuebraAnt = iQuebra
            End If
            
        Loop
        
        'Para cada quebra verifica se teve movimentação, senão insere
        For iQuebraAux = 1 To 3
        
            bTemMov = False
        
            For Each objRelRCPEDetAux In objRelRCPECab.ColDetalhe
                If objRelRCPEDetAux.iQuebra = iQuebraAux Then
                    bTemMov = True
                    Exit For
                End If
            Next
        
            If Not bTemMov Then
            
                Set objRelRCPEDet = New ClassRelRCPEDet
                               
                Set objRelRCPEDet.objItemMovEst = New ClassItemMovEstoque
                Set objRelRCPEDet.objItemNF = New ClassItemNF
                Set objRelRCPEDet.objNFiscal = New ClassNFiscal
                Set objRelRCPEDet.objTipoMovEst = New ClassTipoMovEst
                        
                objRelRCPEDet.iQuebra = iQuebraAux
                objRelRCPEDet.objItemMovEst.dtData = DATA_NULA
                        
                objRelRCPECab.ColDetalhe.Add objRelRCPEDet
                
            End If
        
        Next
        
        colRelRCPECab.Add objRelRCPECab
        
    End If
   
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelRCPEMod3_Prepara_EmTrans1 = SUCESSO
    
    Exit Function
 
Erro_RelRCPEMod3_Prepara_EmTrans1:
 
    RelRCPEMod3_Prepara_EmTrans1 = gErr
 
    Select Case gErr
        
        Case 177370
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 177371, 177372, 177375 To 177380
        
        Case 177373, 177374, 177381
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177382)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Public Function RelRCPEMod3_Prepara_EmTrans2(ByVal lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal colRelRCPECab As Collection) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim objRelRCPECab As ClassRelRCPECab
Dim objRelRCPEDet As ClassRelRCPEDet
Dim iSeqCab As Integer
Dim iSeqDet As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim objRegInventarioFil As ClassRegInventario
Dim dQuantEstInicial As Double
Dim sSerie As String
Dim lNumNotaFiscal As Long
Dim dtDataEmissaoNF As Date
Dim sNaturezaNF As String
Dim objItemNF As ClassItemNF
Dim lNumIntDocItemNF As Long
Dim lNumIntDocNF As Long
Dim dIPIBaseCalculo As Double
Dim dIPICredito As Double
Dim dSaldoAtual As Double
Dim iQuebraAnt As Integer
Dim dMovimentacao As Double
Dim bTemMov As Boolean

On Error GoTo Erro_RelRCPEMod3_Prepara_EmTrans2

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 177383
    Next
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    'Le os dados da Filial Empresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 177384
    
    iSeqCab = 0

    For Each objRelRCPECab In colRelRCPECab
    
        iSeqCab = iSeqCab + 1
        iSeqDet = 0
        
        Set objRegInventarioFil = New ClassRegInventario
    
        With objRegInventarioFil
            .iFilialEmpresa = iFilialEmpresa
            .sProduto = objRelRCPECab.objProduto.sCodigo
            .iAlmoxarifado = 0
            .dtData = DateAdd("d", -1, objRelRCPECab.dtDataDe)
        End With
        
        'Pega o estoque do dia anterior
        lErro = CF("RegInventario_Obter_Quantidades", objRegInventarioFil)
        If lErro <> SUCESSO Then gError 177385
        
        objRegInventarioFil.dQtdeDeTercConosco = objRegInventarioFil.dQuantBenef3 + objRegInventarioFil.dQuantConserto3 + objRegInventarioFil.dQuantConsig3 + objRegInventarioFil.dQuantDemo3 + objRegInventarioFil.dQuantOutras3
        objRegInventarioFil.dQtdeNossaEmTerc = objRegInventarioFil.dQuantBenef + objRegInventarioFil.dQuantConserto + objRegInventarioFil.dQuantConsig + objRegInventarioFil.dQuantDemo + objRegInventarioFil.dQuantOutras
        
        'Verifica se o produto teve movimentação
        bTemMov = False
        For Each objRelRCPEDet In objRelRCPECab.ColDetalhe
            If objRelRCPEDet.objItemMovEst.dtData <> DATA_NULA Then
                    bTemMov = True
                Exit For
            End If
        Next
        
        'Se tem alguma movimentação ou saldo inicial(Nosso ou Nosso em Terceiros ou de Terceiros)
        If bTemMov Or (objRegInventarioFil.dQtdeNossaEmTerc <> 0) Or (objRegInventarioFil.dQtdeDeTercConosco <> 0) Or (objRegInventarioFil.dQuantidadeUMEstoque - objRegInventarioFil.dQtdeDeTercConosco <> 0) Then
        
            'Insere o cabeçalho
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelRCPEMod3Cab (NumIntRel, SeqCab, Empresa, InscricaoEstadual, CNPJ, Folha, DataDe, DataAte, Produto, DescricaoProduto, UMEstoque, ClassificacaoFiscal, FilialEmpresa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, iSeqCab, gsNomeEmpresa, objFilialEmpresa.sInscricaoEstadual, objFilialEmpresa.sCgc, objRelRCPECab.lFolha, objRelRCPECab.dtDataDe, objRelRCPECab.dtDataAte, objRelRCPECab.objProduto.sCodigo, objRelRCPECab.objProduto.sDescricao, objRelRCPECab.objProduto.sSiglaUMEstoque, objRelRCPECab.objProduto.sIPICodigo, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 177386
    
            dMovimentacao = 0
    
            'Para cada item do detalhe
            For Each objRelRCPEDet In objRelRCPECab.ColDetalhe
            
                iSeqDet = iSeqDet + 1
            
                Select Case objRelRCPEDet.iQuebra
                    Case 1 'Nosso
                        dQuantEstInicial = objRegInventarioFil.dQuantidadeUMEstoque - objRegInventarioFil.dQtdeDeTercConosco
                    Case 2 'Nosso em Terceiros
                        dQuantEstInicial = objRegInventarioFil.dQtdeNossaEmTerc
                    Case 3 'De terceiros
                        dQuantEstInicial = objRegInventarioFil.dQtdeDeTercConosco
                    Case Else
                        dQuantEstInicial = 0
                End Select
                
                If iQuebraAnt <> objRelRCPEDet.iQuebra Then
                    iQuebraAnt = objRelRCPEDet.iQuebra
                    dMovimentacao = 0
                End If
                
                If objRelRCPEDet.sEntradaOuSaida = "E" Then
                    dMovimentacao = dMovimentacao + objRelRCPEDet.objItemMovEst.dQuantidadeEst
                Else
                    dMovimentacao = dMovimentacao - objRelRCPEDet.objItemMovEst.dQuantidadeEst
                End If
                
                dSaldoAtual = dQuantEstInicial + dMovimentacao
                
                sNaturezaNF = ""
                dIPIBaseCalculo = 0
                dIPICredito = 0
                            
                If objRelRCPEDet.objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Then
                    sSerie = objRelRCPEDet.objNFiscal.sSerie
                    lNumNotaFiscal = objRelRCPEDet.objNFiscal.lNumNotaFiscal
                    dtDataEmissaoNF = objRelRCPEDet.objNFiscal.dtDataEmissao
                    lNumIntDocNF = objRelRCPEDet.objNFiscal.lNumIntDoc
                    lNumIntDocItemNF = objRelRCPEDet.objItemNF.lNumIntDoc
                    
                    'Lê os Ítens da Nota Fiscal
                    lErro = CF("NFiscalItens_Le", objRelRCPEDet.objNFiscal)
                    If lErro <> SUCESSO Then gError 177392
            
                    'Le a parte Tributária
                    lErro = CF("NFiscalItens_Le_Trib", objRelRCPEDet.objNFiscal)
                    If lErro <> SUCESSO And lErro <> 22877 Then gError 177387
                    
                    'Pega a natureza de operação
                    For Each objItemNF In objRelRCPEDet.objNFiscal.ColItensNF
                        If objItemNF.iItem = objRelRCPEDet.objItemNF.iItem Then
                            sNaturezaNF = objItemNF.objTributacaoItemNF.sNaturezaOp
                            dIPIBaseCalculo = objItemNF.objTributacaoItemNF.dIPIBaseCalculo
                            dIPICredito = objItemNF.objTributacaoItemNF.dIPICredito
                            
                            If dIPIBaseCalculo < DELTA_VALORMONETARIO Then dIPIBaseCalculo = objItemNF.dValorTotal
                            
                            Exit For
                        End If
                    Next
                    
                Else
                    sSerie = ""
                    lNumNotaFiscal = 0
                    dtDataEmissaoNF = DATA_NULA
                    lNumIntDocNF = 0
                    lNumIntDocItemNF = 0
                End If
                
                'Se não tem estoque e não tem movimentação não insere a quebra
                If dQuantEstInicial <> 0 Or objRelRCPEDet.objItemMovEst.dtData <> DATA_NULA Then
                
                    lErro = Comando_Executar(alComando(2), "INSERT INTO RelRCPEMod3Det (NumIntRel, SeqCab, SeqDet,Quebra,TipoNumIntDocOrigem, Sigla, CodigoOP, Serie, NumeroNF, CodMovEstoque, DataMov, DataEmissaoNF, NaturezaNF, EntradaOuSaida, Codigo, Quantidade, IPIBase, ValorIPI, SaldoAtualData, EstoqueInicial,TotalEntrada, TotalSaida, NumIntDocItemNF, NumIntDocNF, NumIntDocMovEst ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, iSeqCab, iSeqDet, objRelRCPEDet.iQuebra, objRelRCPEDet.objItemMovEst.iTipoNumIntDocOrigem, objRelRCPEDet.objTipoMovEst.sSigla, objRelRCPEDet.objItemMovEst.sOPCodigo, sSerie, lNumNotaFiscal, objRelRCPEDet.objItemMovEst.lCodigo, objRelRCPEDet.objItemMovEst.dtData, dtDataEmissaoNF, sNaturezaNF, objRelRCPEDet.sEntradaOuSaida, objRelRCPEDet.iCodigo, _
                    objRelRCPEDet.objItemMovEst.dQuantidadeEst, dIPIBaseCalculo, dIPICredito, dSaldoAtual, dQuantEstInicial, objRelRCPEDet.dTotalEntrada, objRelRCPEDet.dTotalSaida, lNumIntDocItemNF, lNumIntDocNF, objRelRCPEDet.objItemMovEst.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 177388
                    
                End If
        
            Next
            
        End If
    
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelRCPEMod3_Prepara_EmTrans2 = SUCESSO
 
    Exit Function
 
Erro_RelRCPEMod3_Prepara_EmTrans2:
 
    RelRCPEMod3_Prepara_EmTrans2 = gErr
 
    Select Case gErr
    
        Case 177383
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 177384, 177385, 177387
        
        Case 177386
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELRCPEMOD3CAB", gErr)
 
        Case 177388
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELRCPEMOD3DET", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177389)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
 
End Function

Function RelRCPEMod3_Prepara_EmTrans3(ByVal objRelRCPECab As ClassRelRCPECab, ByVal objRelRCPEDet As ClassRelRCPEDet) As Long
'retorna em iCodigo se o movimento deve ser considerado como:
'1: no próprio estabelecimento
'2: em outro estabelecimento
'3: diversos

'iQuebra: 1: nosso conosco, 2: nosso em 3os, 3: de 3os conosco

Dim lErro As Long

On Error GoTo Erro_RelRCPEMod3_Prepara_EmTrans3

    objRelRCPEDet.iCodigo = 0
    
    Select Case objRelRCPEDet.iQuebra
    
        Case 1 'nosso conosco
    
            Select Case objRelRCPEDet.sEntradaOuSaida
            
                'entradas
                Case "E"
                
                    'se producao entrada => coluna 1
                    If objRelRCPEDet.objItemMovEst.iTipoMov = MOV_EST_PRODUCAO Then objRelRCPEDet.iCodigo = 1
                    
                    'se entrada de material beneficiado => coluna 2
                    If objRelRCPEDet.objItemMovEst.iTipoMov = 139 Then objRelRCPEDet.iCodigo = 2 '??? depois incluir novo tipo
                    
                    If objRelRCPEDet.iCodigo = 0 Then objRelRCPEDet.iCodigo = 3
                
                'saidas
                Case Else
                                    
                    Select Case objRelRCPECab.objProduto.iNatureza
                    
                        'se produto acabado
                        Case NATUREZA_PROD_PRODUTO_ACABADO
                    
                            'se produzido na propria filial
                            If objRelRCPECab.objProdutoFilial.iProdNaFilial <> 0 Then
                                objRelRCPEDet.iCodigo = 1
                            Else
                                objRelRCPEDet.iCodigo = 2
                            End If
                            
                        Case NATUREZA_PROD_MATERIA_PRIMA, NATUREZA_PROD_PRODUTO_INTERMEDIARIO, NATUREZA_PROD_EMBALAGENS
                        
                            'se requisicao para producao => 1
                            If objRelRCPEDet.objItemMovEst.iTipoMov = MOV_EST_REQ_PRODUCAO Then objRelRCPEDet.iCodigo = 1
                            
                            'se remessa de material para beneficiamento => 2
                            If objRelRCPEDet.objItemMovEst.iTipoMov = MOV_EST_MAT_NOSSO_PARA_BENEF_SAIDA Then objRelRCPEDet.iCodigo = 2
                            
                            If objRelRCPEDet.iCodigo = 0 Then objRelRCPEDet.iCodigo = 3
        
                        Case Else
                        
                            objRelRCPEDet.iCodigo = 3
        
                    End Select
            
            End Select
        
        Case 2 'nosso em 3os
            
            objRelRCPEDet.iCodigo = 3
            
        Case 3 'de 3os conosco
        
            objRelRCPEDet.iCodigo = 3
    
    End Select
    
    RelRCPEMod3_Prepara_EmTrans3 = SUCESSO
     
    Exit Function
    
Erro_RelRCPEMod3_Prepara_EmTrans3:

    RelRCPEMod3_Prepara_EmTrans3 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177390)
     
    End Select
     
    Exit Function

End Function
'########################################################################

Function RelDeclanIPM_Prepara(ByVal bImprimeRel As Boolean, ByVal sNomeArquivo As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicio As Date, ByVal dtDataFim As Date, lNumIntRel As Long) As Long

Dim lErro As Long, iSeq As Integer, sInscricaoEstadualAux As String
Dim lTransacao As Long, iIndice As Integer
Dim alComando(0 To 1) As Long, sRegistro As String
Dim sInscricaoEstadual As String, sCNPJ As String, sRazaoSocial As String, sNaturezaOp As String
Dim lNumNotaFiscal As Long, dtDataEmissao As Date, dValorTotal As Double
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_RelDeclanIPM_Prepara

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103552
    Next

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 103508
        
    'obtem valor de NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 103509
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    'Le os dados da Filial Empresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 103510
    If lErro <> SUCESSO Then gError 103511
    
    sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    sCNPJ = String(STRING_CGC, 0)
    sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT FiliaisClientes.InscricaoEstadual, FiliaisClientes.CGC, Clientes.RazaoSocial, NFiscal.NumNotaFiscal, NFiscal.DataEmissao, NFiscal.ValorTotal, NFiscal.NaturezaOp FROM NFiscal, Clientes, FiliaisClientes, Enderecos WHERE NFiscal.Status <> ? AND NFiscal.FilialEmpresa = ? AND NFiscal.Cliente = Clientes.Codigo AND Clientes.Codigo = FiliaisClientes.CodCliente AND NFiscal.FilialCli = FiliaisClientes.CodFilial AND FiliaisClientes.Endereco = Enderecos.Codigo AND Enderecos.SiglaEstado = ? AND FiliaisClientes.InscricaoEstadual <> '' AND DataEmissao BETWEEN ? AND ? AND NFiscal.NaturezaOp BETWEEN ? AND ? ORDER BY DataEmissao, NumNotafiscal", _
        sInscricaoEstadual, sCNPJ, sRazaoSocial, lNumNotaFiscal, dtDataEmissao, dValorTotal, sNaturezaOp, _
        STATUS_CANCELADO, iFilialEmpresa, objFilialEmpresa.objEndereco.sSiglaEstado, dtDataInicio, dtDataFim, "5101", "5125")
    If lErro <> AD_SQL_SUCESSO Then gError 103512
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103513

    'Se não encontrou nenhum registro => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 103514
    
    If sNomeArquivo = "" Then sNomeArquivo = "c:\sge\programa\declan" & CStr(lNumIntRel) & ".txt"
    
    Open sNomeArquivo For Output As #1
    
    'Enquanto houver registros
    Do While lErro <> AD_SQL_SEM_DADOS

        iSeq = iSeq + 1
        
        sInscricaoEstadualAux = IIf(Len(Replace(sInscricaoEstadual, ".", "")) > 8, "11111111", Replace(sInscricaoEstadual, ".", ""))
        
        'insere o registro na tabela temporaria
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelDeclanIPM (NumIntRel,Seq,InscEstadualCliente,CNPJ,RazaoSocialCli,NumNotaFiscal,DataNotaFiscal,ValorNotaFiscal,CFOP) VALUES (?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, iSeq, sInscricaoEstadualAux, sCNPJ, sRazaoSocial, lNumNotaFiscal, dtDataEmissao, dValorTotal, sNaturezaOp)
        If lErro <> AD_SQL_SUCESSO Then gError 103545

        sRegistro = "3" & FormataCpoTexto(sInscricaoEstadualAux, 8) & FormataCpoTexto(sCNPJ, 14) & FormataCpoTexto(sRazaoSocial, 64) & FormataCpoTexto(CStr(lNumNotaFiscal), 9) & Format(dtDataEmissao, "dd/mm/yyyy") & FormataCpoTexto(Arredonda_Moeda(dValorTotal * 100), 12) & sNaturezaOp
        
        Print #1, sRegistro

        'busca o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103546

    Loop
    
    Close #1
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 103515
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelDeclanIPM_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelDeclanIPM_Prepara:

    RelDeclanIPM_Prepara = gErr
     
    Select Case gErr
          
        Case 103515
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 103509, 103510
        
        Case 103508
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 103552
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 103511
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, iFilialEmpresa)
        
        Case 103512, 103513, 103514, 103546
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDECLAN", gErr)
        
        Case 103545
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_RELDECLAN", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130507)
     
    End Select
     
    'Efetua rollback da transacao
    Call Transacao_Rollback
    
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'#####################################################################
'Inserido por Wagner 28/08/2006
Public Function LivRegES_Atualiza_Emitentes(ByVal objLivRegESAtualiza As ClassLivRegESAtualiza) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(1 To 5) As Long

On Error GoTo Erro_LivRegES_Atualiza_Emitentes

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182012
    Next

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182013
    
    'Atualiza as informações de Clientes e Fornecedores nas Tabelas
    'de LivRegES e LivRegESEmitentes
    lErro = LivRegES_Atualiza_Emitentes_Trans(objLivRegESAtualiza, alComando)
    If lErro <> SUCESSO Then gError 182014
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182015
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    LivRegES_Atualiza_Emitentes = SUCESSO
     
    Exit Function
    
Erro_LivRegES_Atualiza_Emitentes:

    LivRegES_Atualiza_Emitentes = gErr
     
    Select Case gErr
          
        Case 182012
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 182013
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182014
        
        Case 182015
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182016)
     
    End Select
         
    'Efetua rollback da transacao
    Call Transacao_Rollback
    
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function LivRegES_Atualiza_Emitentes_Trans(ByVal objLivRegESAtualiza As ClassLivRegESAtualiza, alComando() As Long) As Long
'Atualiza as informações de Clientes e Fornecedores nas Tabelas
'de LivRegES e LivRegESEmitentes

Dim lErro As Long
Dim sSQL As String
Dim lNumIntDoc As Long
Dim lNumIntEmitente As Long
Dim colFilialCliente As New Collection
Dim colFilialClienteEnd As New Collection
Dim colFilialFornecedor As New Collection
Dim colFilialFornecedorEnd As New Collection
Dim lCliente As Long
Dim lFornecedor As Long
Dim iFilialCli As Integer
Dim iFilialForn As Integer
Dim sCgc As String
Dim sIE As String
Dim sUF As String
Dim iAux As Integer
Dim colEmitentes As New Collection
Dim bAchou As Boolean
Dim vlNumIntEmitente As Variant
Dim iIEIsento As Integer

On Error GoTo Erro_LivRegES_Atualiza_Emitentes_Trans

    'Prepara o SQL
    lErro = LivRegES_Atualiza_PreparaSQL(objLivRegESAtualiza, sSQL)
    If lErro <> SUCESSO Then gError 182041

    'Passa os parâmetros
    lErro = LivRegES_Atualiza_PreparaInt(alComando(1), sSQL, TIPO_NUMINTDOC_ORIGEM_NOTAFISCAL, objLivRegESAtualiza.dtData, objLivRegESAtualiza.lClienteDe, objLivRegESAtualiza.lClienteAte, objLivRegESAtualiza.lFornecedorDe, objLivRegESAtualiza.lFornecedorAte, lNumIntDoc, lCliente, lFornecedor, iFilialCli, iFilialForn)
    If lErro <> SUCESSO Then gError 182042
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182043

    'Enquanto existirem dados de acordo com o filtro passado
    Do While lErro <> AD_SQL_SEM_DADOS
   
        'Le posicionado em LivRegES
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntEmitente FROM LivRegES WHERE NumIntDoc = ?", 0, lNumIntEmitente, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 182044
  
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182045
    
        'Sem Dados = > Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 182046
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> SUCESSO Then gError 182047
        
        'Com s informações do código do cliente\fornecedor e filial busca os dados que serão
        'atualizados
        lErro = LivRegES_Atualiza_Obtem_CliForn(lCliente, lFornecedor, iFilialCli, iFilialForn, colFilialFornecedor, colFilialFornecedorEnd, colFilialCliente, colFilialClienteEnd, sUF, sCgc, sIE, iIEIsento)
        If lErro <> SUCESSO Then gError 182048
        
        'Atualiza a tabela LivRegES
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE LivRegES SET CGC = ?, UF = ? , InscricaoEstadual = ?, IEIsento = ?", alComando(2), sCgc, sUF, sIE, iIEIsento)
        If lErro <> AD_SQL_SUCESSO Then gError 182049
        
        'Se tiver associada a um emitente
        If lNumIntEmitente <> 0 Then
        
            'Busca na coleção de emitentes já atualizados
            bAchou = False
            For Each vlNumIntEmitente In colEmitentes
                If vlNumIntEmitente = lNumIntEmitente Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            'Se encontrou não precisa atualizar de novo,
            'Se não encontrou ....
            If Not bAchou Then
            
                'Le posicionado em LivRegESEmitentes
                lErro = Comando_ExecutarPos(alComando(4), "SELECT 1 FROM LivRegESEmitentes WHERE Codigo = ?", 0, iAux, lNumIntEmitente)
                If lErro <> AD_SQL_SUCESSO Then gError 182050
        
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182051
          
                'Sem Dados = > Erro
                If lErro = AD_SQL_SEM_DADOS Then gError 182052
          
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(4))
                If lErro <> SUCESSO Then gError 182053
              
                'Atualiza a tabela LivRegES
                lErro = Comando_ExecutarPos(alComando(5), "UPDATE LivRegESEmitentes SET CGC = ?, UF = ? , InscricaoEstadual = ?", alComando(4), sCgc, sUF, sIE)
                If lErro <> AD_SQL_SUCESSO Then gError 182054
                
                'Adiciona o Código do Emitente na tabela para evitar alterações desnecessárias
                colEmitentes.Add lNumIntEmitente
                
            End If
            
        End If
    
        'Busca Primeiro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182055
    
    Loop
    
    LivRegES_Atualiza_Emitentes_Trans = SUCESSO
     
    Exit Function
    
Erro_LivRegES_Atualiza_Emitentes_Trans:

    LivRegES_Atualiza_Emitentes_Trans = gErr
     
    Select Case gErr
    
        Case 182041, 182042, 182048
        
        Case 182043, 182044, 182045, 182055
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
            
        Case 182046
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVREGES_NAO_ENCONTRADO", gErr)
        
        Case 182047
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGES", gErr)
            
        Case 182049
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
        
        Case 182050, 182051
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESEMITENTES", gErr)
            
        Case 182052
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LIVREGESEMITENTES", gErr)
        
        Case 182053
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGESEMITENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182017)
     
    End Select
     
    Exit Function

End Function

Private Function LivRegES_Atualiza_PreparaSQL(ByVal objLivRegESAtualiza As ClassLivRegESAtualiza, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String

On Error GoTo Erro_LivRegES_Atualiza_PreparaSQL

    sSQL = "SELECT L.NumIntDoc, NF.Cliente, NF.Fornecedor, NF.FilialCli, NF.FilialForn " & _
            "FROM LivRegES AS L, NFiscal AS NF "
    sWhere = "WHERE L.TipoNumIntDocOrigem = ? AND " & _
            "L.NumIntDocOrigem = NF.NumIntDoc AND L.DataEmissao >= ? "
            
    If objLivRegESAtualiza.iIncluiLivRegFechados = DESMARCADO Then
        sWhere = sWhere & "AND L.NumIntDocLivFechado = 0 "
    End If
    
    If objLivRegESAtualiza.iAtualizaCliente = MARCADO Then
        sWhere = sWhere & "AND NF.Cliente <> 0 "
    Else
        sWhere = sWhere & "AND NF.Fornecedor <> 0 "
    End If
    
    If objLivRegESAtualiza.lClienteDe <> 0 Then
        sWhere = sWhere & "AND NF.Cliente >= ? "
    End If
    
    If objLivRegESAtualiza.lClienteAte <> 0 Then
        sWhere = sWhere & "AND NF.Cliente <= ? "
    End If
    
    If objLivRegESAtualiza.lFornecedorDe <> 0 Then
        sWhere = sWhere & "AND NF.Fornecedor >= ? "
    End If
    
    If objLivRegESAtualiza.lFornecedorAte <> 0 Then
        sWhere = sWhere & "AND NF.Fornecedor <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    LivRegES_Atualiza_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_LivRegES_Atualiza_PreparaSQL:

    LivRegES_Atualiza_PreparaSQL = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182018)
     
    End Select
     
    Exit Function

End Function

Private Function LivRegES_Atualiza_PreparaInt(ByVal lComando As Long, ByVal sSQL As String, ByVal viTipoNumIntOrigem As Variant, ByVal vdtData As Variant, ByVal vlClienteDe As Variant, ByVal vlClienteAte As Variant, ByVal vlFornecedorDe As Variant, ByVal vlFornecedorAte As Variant, vlNumIntDoc As Variant, vlCliente As Variant, vlFornecedor As Variant, viFilialCli As Variant, viFilialForn As Variant) As Long

Dim lErro As Long

On Error GoTo Erro_LivRegES_Atualiza_PreparaInt

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182020
                       
    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182021

    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182022

    lErro = Comando_BindVarInt(lComando, vlFornecedor)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182023

    lErro = Comando_BindVarInt(lComando, viFilialCli)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182024

    lErro = Comando_BindVarInt(lComando, viFilialForn)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182025

    lErro = Comando_BindVarInt(lComando, viTipoNumIntOrigem)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182026

    lErro = Comando_BindVarInt(lComando, vdtData)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182027
    
    If vlClienteDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 182028
    End If

    If vlClienteAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 182029
    End If

    If vlFornecedorDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlFornecedorDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 182030
    End If

    If vlFornecedorAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlFornecedorAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 182031
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 182032
    
    LivRegES_Atualiza_PreparaInt = SUCESSO
     
    Exit Function
    
Erro_LivRegES_Atualiza_PreparaInt:

    LivRegES_Atualiza_PreparaInt = gErr
     
    Select Case gErr

        Case 182020 To 182032
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182019)
     
    End Select
     
    Exit Function

End Function

Private Function LivRegES_Atualiza_Obtem_CliForn(ByVal lCliente As Long, ByVal lFornecedor As Long, ByVal iFilialCli As Integer, ByVal iFilialForn As Integer, ByVal colFilialFornecedor As Collection, ByVal colFilialFornecedorEnd As Collection, ByVal colFilialCliente As Collection, ByVal colFilialClienteEnd As Collection, sUF As String, sCgc As String, sIE As String, iIEIsento As Integer) As Long
'Le a UF, CGC e Inscrição estatual ligados a Filial Cliente\Fornecedor ligados a Nota Fiscal.

Dim lErro As Long
Dim objFilialCliente As ClassFilialCliente
Dim objFilialFornecedor As ClassFilialFornecedor
Dim objEndereco As ClassEndereco
Dim bAchou As Boolean
Dim iIndice As Integer
Dim sIEAux As String

On Error GoTo Erro_LivRegES_Atualiza_Obtem_CliForn
    
    'Se for uma Nota Ligada a Um Cliente
    If lCliente <> 0 Then
    
        'Verifica se a filial Cliente já foi lida
        bAchou = False
        iIndice = 0
        For Each objFilialCliente In colFilialCliente
            iIndice = iIndice + 1
            If objFilialCliente.lCodCliente = lCliente And objFilialCliente.iCodFilial = iFilialCli Then
                bAchou = True
                Set objEndereco = colFilialClienteEnd.Item(iIndice)
                Exit For
            End If
        Next
        
        'Se não leu ainda, lê e guarda
        If Not bAchou Then
        
            Set objFilialCliente = New ClassFilialCliente
            Set objEndereco = New ClassEndereco
            
            objFilialCliente.lCodCliente = lCliente
            objFilialCliente.iCodFilial = iFilialCli
        
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 182033
            
            lErro = CF("FilialCliente_Lock", lCliente, iFilialCli)
            If lErro <> SUCESSO And lErro <> 43050 Then gError 182034
            
            If lErro <> SUCESSO Then gError 182035
            
            objEndereco.lCodigo = objFilialCliente.lEndereco
            
            'Le o Endereco de Principal
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 182036
            
            colFilialCliente.Add objFilialCliente
            colFilialClienteEnd.Add objEndereco
                            
        End If
        
        sCgc = objFilialCliente.sCgc
        sIE = objFilialCliente.sInscricaoEstadual
        iIEIsento = objFilialCliente.iIEIsento
    
    Else
    
        'Verifica se a filial Fornecedor já foi lido
        bAchou = False
        iIndice = 0
        For Each objFilialFornecedor In colFilialFornecedor
            iIndice = iIndice + 1
            If objFilialFornecedor.lCodFornecedor = lFornecedor And objFilialFornecedor.iCodFilial = iFilialForn Then
                bAchou = True
                Set objEndereco = colFilialFornecedorEnd.Item(iIndice)
                Exit For
            End If
        Next
    
        'Se não leu ainda, lê e guarda
        If Not bAchou Then
        
            Set objFilialFornecedor = New ClassFilialFornecedor
            Set objEndereco = New ClassEndereco
        
            objFilialFornecedor.lCodFornecedor = lFornecedor
            objFilialFornecedor.iCodFilial = iFilialForn
            
            lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 12929 Then gError 182037
            
            lErro = CF("FilialFornecedor_Lock1", objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18900 Then gError 182038
            
            If lErro <> SUCESSO Then gError 182039
            
            objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
            'Le o Endereco de Principal
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 182040
            
            colFilialFornecedor.Add objFilialFornecedor
            colFilialFornecedorEnd.Add objEndereco
                            
        End If
        
        sCgc = objFilialFornecedor.sCgc
        sIE = objFilialFornecedor.sInscricaoEstadual
        iIEIsento = objFilialFornecedor.iIEIsento
    
    End If
    
    sIEAux = sIE
    sIE = ""
    Call Formata_String_Numero(sIEAux, sIE)
    'sIE = Replace(Replace(sIE, ".", ""), "-", "")
    
    'Se for de outro pais a sigla do estado é "EX"
    If objEndereco.iCodigoPais <> PAIS_BRASIL Then
        sUF = "EX"
    Else
        sUF = objEndereco.sSiglaEstado
    End If
    
    LivRegES_Atualiza_Obtem_CliForn = SUCESSO
     
    Exit Function
    
Erro_LivRegES_Atualiza_Obtem_CliForn:

    LivRegES_Atualiza_Obtem_CliForn = gErr
     
    Select Case gErr
    
        Case 182033, 182034, 182036, 182037, 182038, 182040
        
        Case 182035
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
        
        Case 182039
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182017)
     
    End Select
     
    Exit Function

End Function
'#####################################################################

Function Sintegra_SQLRegES(iStatusLivro As Integer, iSituacao As Integer) As Long
    Sintegra_SQLRegES = SUCESSO
End Function

Private Function DIPJ2007_GeraFicha_2426(ByVal iFilialEmpresa As Integer, ByVal iFicha As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iAnoBase As Integer) As Long
'Preenche tabela DIPJ2007_2426 com registros para preencher a ficha 24 ou 26

Dim lErro As Long, iIndice As Integer, objProduto As ClassProduto, objProd As ClassProduto
Dim alComando(0 To 5) As Long, lNumIntDocRegES As Long, bAchou As Boolean
Dim tProd As typeProduto, dValorProduto As Double, dValorRatear As Double, dValorBaseRateio As Double
Dim colProdutosNF As Collection, colProdutosComRateio As New Collection

On Error GoTo Erro_DIPJ2007_GeraFicha_2426
    
    'abre o comando
    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 121106
    Next
        
    'Insere os dados para as fichas 23 e 24 da DIPJ
    lErro = Comando_Executar(alComando(0), "INSERT INTO DIPJ_26_28 (NumIntRel, Ficha, IPICodigo, Nome, Valor) (SELECT ?, ?, IPICodigo, Produtos.Descricao, SUM(ValorProduto) Valor FROM LivRegES, LivRegESItemNF, LivRegESCadProd, Produtos, DIPJFichasIPINatOP WHERE " & _
        " LivRegES.NumIntDoc = LivRegESItemNF.NumIntDocRegES AND LivRegESItemNF.NumIntCadProd = LivRegESCadProd.NumIntDoc AND LivRegESItemNF.CFOP = DIPJFichasIPINatOP.NaturezaOP AND DIPJFichasIPINatOP.Ficha = ? AND DIPJFichasIPINatOP.AnoBase = ? AND LivRegESCadProd.Produto = Produtos.Codigo AND LivRegES.Data BETWEEN ? AND ? AND LivRegES.FilialEmpresa = ? and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 GROUP BY IPICodigo, Produtos.Descricao )", _
        lNumIntRel, iFicha, iFicha, iAnoBase, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'pesquisar notas que tenham frete, seguro e outras despesas
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDocRegES, SUM(ValorDescontoDespAcess) FROM LivRegESItemNF, LivRegES L1 WHERE LivRegESItemNF.NumIntDocRegES = L1.NumIntDoc AND NumItem IN (991, 992, 999) AND L1.Data BETWEEN ? AND ? AND L1.FilialEmpresa = ? and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 AND EXISTS " & _
        "(SELECT L2.NumIntDoc FROM LivRegES L2, LivRegESItemNF, DIPJFichasIPINatOP WHERE L2.NumIntDoc = LivRegESItemNF.NumIntDocRegES AND LivRegESItemNF.CFOP = DIPJFichasIPINatOP.NaturezaOP AND DIPJFichasIPINatOP.Ficha = ? AND DIPJFichasIPINatOP.AnoBase = ? AND NumItem < 900 AND L1.NumIntDoc = L2.NumIntDoc) GROUP BY NumIntDocRegES", _
        lNumIntDocRegES, dValorRatear, dtDataInicial, dtDataFinal, iFilialEmpresa, iFicha, iAnoBase)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 121107
    
    Do While lErro = AD_SQL_SUCESSO
    
        dValorBaseRateio = 0
        Set colProdutosNF = New Collection
        
        'pesquisar itens da nota
        tProd.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        tProd.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        lErro = Comando_Executar(alComando(2), "SELECT IPICodigo, Produtos.Descricao, ValorProduto FROM LivRegESItemNF, LivRegESCadProd, Produtos, DIPJFichasIPINatOP WHERE LivRegESItemNF.NumIntCadProd = LivRegESCadProd.NumIntDoc AND LivRegESCadProd.Produto = Produtos.Codigo AND NumItem < 900 AND LivRegESItemNF.NumIntDocRegES = ? AND LivRegESItemNF.CFOP = DIPJFichasIPINatOP.NaturezaOP AND DIPJFichasIPINatOP.Ficha = ? AND DIPJFichasIPINatOP.AnoBase = ?", _
            tProd.sIPICodigo, tProd.sDescricao, dValorProduto, lNumIntDocRegES, iFicha, iAnoBase)
        If lErro <> AD_SQL_SUCESSO Then gError 121107
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 121107
        
        Do While lErro = AD_SQL_SUCESSO
        
            dValorBaseRateio = dValorBaseRateio + dValorProduto
            
            Set objProduto = New ClassProduto
            
            With objProduto
                .sDescricao = tProd.sDescricao
                .sIPICodigo = tProd.sIPICodigo
                .dPesoBruto = dValorProduto
            End With
    
            colProdutosNF.Add objProduto
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 121107
    
        Loop
        
        'incorporar os itens ao produtos com rateio
        For Each objProd In colProdutosNF
                
            bAchou = False
            For Each objProduto In colProdutosComRateio
            
                If objProd.sDescricao = objProduto.sDescricao And objProd.sIPICodigo = objProduto.sIPICodigo Then
                    bAchou = True
                    If dValorBaseRateio <> 0 Then objProduto.dPesoLiq = objProduto.dPesoLiq + Arredonda_Moeda((objProd.dPesoBruto / dValorBaseRateio) * dValorRatear)
                    Exit For
                End If
            
            Next
            
            If bAchou = False Then
            
                If dValorBaseRateio <> 0 Then
                    objProd.dPesoLiq = Arredonda_Moeda((objProd.dPesoBruto / dValorBaseRateio) * dValorRatear)
                Else
                    objProd.dPesoLiq = 0
                End If
                colProdutosComRateio.Add objProd
                
            End If
        
        Next
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 121107
    
    Loop
    
    For Each objProd In colProdutosComRateio
    
        lErro = Comando_Executar(alComando(3), "UPDATE DIPJ_26_28 SET Valor = Valor + ? WHERE NumIntRel = ? AND Ficha = ? AND IPICodigo = ? AND Nome = ?", _
            objProd.dPesoLiq, lNumIntRel, iFicha, objProd.sIPICodigo, objProd.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    DIPJ2007_GeraFicha_2426 = SUCESSO

    Exit Function

Erro_DIPJ2007_GeraFicha_2426:

    DIPJ2007_GeraFicha_2426 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150385)

    End Select

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Function DIPJ2007_GeraFicha_2122(ByVal iFilialEmpresa As Integer, ByVal iFicha As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iAnoBase As Integer) As Long
'Preenche tabela DIPJ_23_24 com registros para preencher as fichas 30 e 31

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ2007_GeraFicha_2122
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 30 e 31 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_23_24 (NumIntRel, Ficha, IPIBase, IPIOutros, IPIValor) (SELECT ?, Ficha, SUM(ValorIPIBase), SUM(ValorIPIIsentoNaoTrib+ValorIPIOutras), SUM(ValorIPI) FROM LivRegES, LivRegESLinha, DIPJFichasIPINatOP WHERE LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and DIPJFichasIPINatOP.NaturezaOp = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 and DIPJFichasIPINatOP.Anobase = ? AND DIPJFichasIPINatOP.Ficha BETWEEN ? AND ? group by DIPJFichasIPINatOP.Ficha )", _
        lNumIntRel, dtDataInicial, dtDataFinal, iFilialEmpresa, iAnoBase, iFicha * 100, (iFicha * 100) + 99)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ2007_GeraFicha_2122 = SUCESSO

    Exit Function

Erro_DIPJ2007_GeraFicha_2122:

    DIPJ2007_GeraFicha_2122 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150398)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Function DIPJ2007_GeraFicha_2325(ByVal iFilialEmpresa As Integer, ByVal iFicha As Integer, ByVal lNumIntRel As Long, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iAnoBase As Integer) As Long
'Preenche tabela DIPJ_25_27 com registros para preencher as fichas 23 e 25

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_DIPJ2007_GeraFicha_2325
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 121106
        
    'Insere os dados para as fichas 25 e 27 da DIPJ
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, NumIntEmitente, Valor) (SELECT ?, ?, CGC, 0, SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, DIPJFichasIPINatOP WHERE DIPJFichasIPINatOP.AnoBase = ? AND DIPJFichasIPINatOP.Ficha = ? AND LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and DIPJFichasIPINatOP.NaturezaOp = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 AND UF <> 'EX' AND UF <> '' group by cgc )", lNumIntRel, iFicha, iAnoBase, iFicha, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    lErro = Comando_Executar(lComando, "INSERT INTO DIPJ_25_27 (NumIntRel, Ficha, CNPJ, NumIntEmitente, Valor) (SELECT ?, ?, '00000000000000', NumIntEmitente, SUM(ValorIPIBase+ValorIPIIsentoNaoTrib+ValorIPIOutras) FROM LivRegES, LivRegESLinha, DIPJFichasIPINatOP WHERE DIPJFichasIPINatOP.AnoBase = ? AND DIPJFichasIPINatOP.Ficha = ? AND LivRegES.Numintdoc = LivRegESLinha.Numintdocreges AND Livreges.data between ? and ? and filialempresa = ? and DIPJFichasIPINatOP.NaturezaOp = livregeslinha.naturezaop and StatusLivro <> 2 and StatusLivro <> 3 and situacao = 0 AND (UF = 'EX' OR UF = '') group by NumIntEmitente )", lNumIntRel, iFicha, iAnoBase, iFicha, dtDataInicial, dtDataFinal, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 121107
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    DIPJ2007_GeraFicha_2325 = SUCESSO

    Exit Function

Erro_DIPJ2007_GeraFicha_2325:

    DIPJ2007_GeraFicha_2325 = gErr

    Select Case gErr

        Case 121106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 121107
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IPIENTRADACREDITO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150386)

    End Select

    'fecha o comando
    Call Comando_Fechar(lComando)
        
    Exit Function

End Function

Private Function NF_LivRegESLinha_Ajusta(ByVal objNFiscal As ClassNFiscal, ByVal colLivRegESLinha As Collection) As Long

Dim lErro As Long
Dim dValorTotal As Double
Dim dValorICMS As Double, dValorICMS51
Dim dValorIPI As Double, dValorICMSSubst As Double
Dim objLivRegESLinha As New ClassLivRegESLinha
Dim dValorDif As Double

On Error GoTo Erro_NF_LivRegESLinha_Ajusta

    dValorTotal = 0
    dValorICMS = 0
    dValorIPI = 0
    dValorICMSSubst = 0

    For Each objLivRegESLinha In colLivRegESLinha
        objLivRegESLinha.dValorTotal = Arredonda_Moeda(objLivRegESLinha.dValorTotal)
        objLivRegESLinha.dValorICMS = Arredonda_Moeda(objLivRegESLinha.dValorICMS)
        objLivRegESLinha.dValorIPI = Arredonda_Moeda(objLivRegESLinha.dValorIPI)
        objLivRegESLinha.dValorICMSSubstRet = Arredonda_Moeda(objLivRegESLinha.dValorICMSSubstRet)
        dValorTotal = dValorTotal + objLivRegESLinha.dValorTotal
        dValorICMS = dValorICMS + objLivRegESLinha.dValorICMS
        dValorICMS51 = dValorICMS51 + objLivRegESLinha.dValorICMS51
        dValorIPI = dValorIPI + objLivRegESLinha.dValorIPI
        dValorICMSSubst = dValorICMSSubst + objLivRegESLinha.dValorICMSSubstRet
    Next
    
    dValorDif = objNFiscal.dValorTotal - dValorTotal
    If Abs(dValorDif) > DELTA_VALORMONETARIO Then
        For Each objLivRegESLinha In colLivRegESLinha
            If objLivRegESLinha.dValorTotal <> 0 Then
                objLivRegESLinha.dValorTotal = Arredonda_Moeda(objLivRegESLinha.dValorTotal + dValorDif)
                objLivRegESLinha.dValorContabil = objLivRegESLinha.dValorTotal
                Exit For
            End If
        Next
    End If
        
    dValorDif = objNFiscal.objTributacaoNF.dIPICredito - dValorIPI
    If Abs(dValorDif) > DELTA_VALORMONETARIO Then
        For Each objLivRegESLinha In colLivRegESLinha
            If objLivRegESLinha.dValorIPI <> 0 Then
                objLivRegESLinha.dValorIPI = objLivRegESLinha.dValorIPI + dValorDif
                Exit For
            End If
        Next
    End If

    dValorDif = objNFiscal.objTributacaoNF.dICMSCredito - (dValorICMS - dValorICMS51)
    If Abs(dValorDif) > DELTA_VALORMONETARIO Then
        For Each objLivRegESLinha In colLivRegESLinha
            If objLivRegESLinha.dValorICMS <> 0 Then
                objLivRegESLinha.dValorICMS = objLivRegESLinha.dValorICMS + dValorDif
                Exit For
            End If
        Next
    End If

    dValorDif = objNFiscal.objTributacaoNF.dICMSSubstValor - dValorICMSSubst
    If Abs(dValorDif) > DELTA_VALORMONETARIO Then
        For Each objLivRegESLinha In colLivRegESLinha
            If objLivRegESLinha.dValorICMSSubstRet <> 0 Then
                objLivRegESLinha.dValorICMSSubstRet = objLivRegESLinha.dValorICMSSubstRet + dValorDif
                Exit For
            End If
        Next
    End If

    For Each objLivRegESLinha In colLivRegESLinha
        If objLivRegESLinha.dValorICMSSubstRet <> 0 And objLivRegESLinha.sObservacaoLivFisc = "" Then
            
            objLivRegESLinha.sObservacaoLivFisc = "ICMS ST BC " & Format(objLivRegESLinha.dValorICMSSubstBase, "Standard") & " ICMS ST " & Format(objLivRegESLinha.dValorICMSSubstRet, "standard")
            
        End If
    Next
    
    NF_LivRegESLinha_Ajusta = SUCESSO

    Exit Function

Erro_NF_LivRegESLinha_Ajusta:

    NF_LivRegESLinha_Ajusta = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150386)

    End Select
        
    Exit Function

End Function

Function RelOpLancEstICMSST_Prepara(lNumIntRel As Long, ByVal sEstadoDe As String, ByVal sEstadoAte As String, ByVal dtDataDe As Date, ByVal dtDataAte As Date) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim lNumIntDoc As Long
Dim sUF As String
Dim dtData As Date
Dim sSerie As String
Dim lNumNotaFiscal As Long
Dim sNome As String
Dim dtDataEmissao As Date
Dim dValorICMSSubstBase As Double
Dim dValorICMSSubstRet As Double

On Error GoTo Erro_RelOpLancEstICMSST_Prepara

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 198697
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 198698
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_REL_APURICMSSUBST", lNumIntRel)
    If lErro <> SUCESSO Then gError 198699
    
    sUF = String(STRING_ESTADO_SIGLA, 0)
    sSerie = String(STRING_SERIE, 0)
    sNome = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)

    'Verifica se o Tipo de Registro de apuração IPI está cadastrada
    lErro = Comando_Executar(alComando(0), "SELECT R.NumIntDoc, R.UF,R.Data,R.Serie,R.NumNotaFiscal,R.Nome, R.DataEmissao, SUM(L.ValorICMSSubstBase) AS ValorICMSSubstBase, SUM(ValorICMSSubstRet) AS  ValorICMSSubstRet FROM LivRegES AS R, LivRegESLinha AS L, NFiscal AS N WHERE L.NumIntDocRegES = R.NumIntDoc AND N.NumIntDoc = R.NumIntDocOrigem AND R.UF <> 'EX' AND L.ValorICMSSubstBase <> 0 AND N.Status <> 7 AND R.StatusLivro IN (0,1) AND R.Tipo = 1 AND R.Situacao = 0 AND (?=? OR R.Data >=?) AND (?=? OR R.Data <=?) AND (?=? OR R.UF >=?) AND (?=? OR R.UF <=?) GROUP BY R.NumIntDoc, R.UF,R.Data,R.Serie,R.NumNotaFiscal,R.Nome, R.DataEmissao", lNumIntDoc, sUF, dtData, sSerie, lNumNotaFiscal, sNome, dtDataEmissao, dValorICMSSubstBase, dValorICMSSubstRet, DATA_NULA, dtDataDe, dtDataDe, DATA_NULA, dtDataAte, dtDataAte, "", sEstadoDe, sEstadoDe, "", sEstadoAte, sEstadoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 198700

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198701

    Do While lErro = AD_SQL_SUCESSO

        'Insere novo Registro de Tipo de Apuração IPI
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelApurICMSSubst (NumIntRel,NumIntDocLivRegES,UF,Cred_Deb,Data,Serie,NumNotaFiscal,Nome,DataEmissao,ValorICMSSubstBase,ValorICMSSubstRet) VALUES (?,?,?,?,?,?,?,?,?,?,?)", lNumIntRel, lNumIntDoc, sUF, 0, dtData, sSerie, lNumNotaFiscal, left(sNome, 40), dtDataEmissao, dValorICMSSubstBase, dValorICMSSubstRet)
        If lErro <> AD_SQL_SUCESSO Then gError 198702

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198703

    Loop

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 198704

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    RelOpLancEstICMSST_Prepara = SUCESSO

    Exit Function

Erro_RelOpLancEstICMSST_Prepara:

    RelOpLancEstICMSST_Prepara = gErr

    Select Case gErr

        Case 198697
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198698
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 198700, 198701, 198703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)

        Case 198702
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGAPURICMSSUBST", gErr)

        Case 198704
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198705)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Importacao_Acerto_260808() As Long

Dim lErro As Long
Dim iIndice As Integer, sSQL As String
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lNumIntDoc As Long, lNumIntDocRegES As Long, lNumIntDocNF As Long
Dim dValorOutras As Double, sNaturezaOp As String, iTipoTrib As Integer

On Error GoTo Erro_Importacao_Acerto_260808

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 11111
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 11112
    
    sSQL = "SELECT l1.numintdoc, nfiscal.numintdoc, nfiscal.valoroutrasdespesas, compl.NaturezaOP, compl.TipoTributacao FROM tributacaocomplnf compl, nfiscal, livreges l1 WHERE compl.numintdoc = nfiscal.numintdoc AND compl.tipo = 4 AND l1.numintdocorigem = nfiscal.numintdoc AND tiponfiscal = 119 AND valoroutrasdespesas <> 0 AND NOT EXISTS (SELECT * FROM livreges, livregesitemnf where livregesitemnf.numitem = 999 and livreges.numintdoc = livregesitemnf.numintdocreges and nfiscal.numintdoc = livreges.numintdocorigem) order by l1.numintdoc"
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
    lErro = Comando_Executar(alComando(1), sSQL, lNumIntDocRegES, lNumIntDocNF, dValorOutras, sNaturezaOp, iTipoTrib)
    If lErro <> AD_SQL_SUCESSO Then gError 11113
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 11114
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Gera número automático para o Item de apuração ICMS
        lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
        If lErro <> SUCESSO Then gError 11115
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO LivRegESItemNF (NumIntDoc, NumIntDocRegES,NumItem,CFOP,NumIntCadProd,Quantidade,ValorProduto,ValorDescontoDespAcess,ValorICMSBase,ValorICMSSubstBase,ValorIPI,AliquotaICMS,TipoTribICMS,TipoTribIPI,ValorICMS,ValorIPIBase,AliquotaIPI,RedBaseICMS,RedBaseIPI,TipoTributacao,AliquotaSubst,ValorSubst) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntDoc, lNumIntDocRegES, 999, sNaturezaOp, 0, 0, 0, dValorOutras, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, iTipoTrib, 0, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 11119
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 11116
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 11117

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Importacao_Acerto_260808 = SUCESSO
    
    Exit Function
    
Erro_Importacao_Acerto_260808:

    Importacao_Acerto_260808 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 11118)

    End Select
    
    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RelRegApurIPI_Prepara(ByVal iFilialEmpresa As Integer, ByVal dtDataInicio As Date, ByVal dtDataFim As Date, ByVal iFolha As Integer, lNumIntRel As Long) As Long

Dim lErro As Long, iSeq As Integer
Dim lTransacao As Long, iIndice As Integer
Dim alComando(0 To 7) As Long
Dim iCodigo As Integer
Dim sDescricao As String
Dim iSecao As Integer
Dim iGrupoAnt As Integer
Dim iLancamento As Integer
Dim dValor As Double
Dim objRegItem As ClassRegApuracaoItem
Dim colRegItem As Collection
Dim sDescricaoItem As String
Dim dValorItem As Double
Dim iGrupo As Integer
Dim sDescGrupo As String
Dim bIncluiObs As Boolean
Dim sCgc  As String, sIE As String
Dim dValores(1 To 30) As Double
Dim objApuracao As ClassRegApuracao
Dim dValorAux As Double

On Error GoTo Erro_RelRegApurIPI_Prepara

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202742
    Next

    'Abertura transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 202743
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_REL_REGAPURIPI", lNumIntRel)
    If lErro <> SUCESSO Then gError 202744
    
    sDescricao = String(255, 0)
     
    lErro = Comando_Executar(alComando(0), "SELECT Codigo, Descricao, Secao, Lancamento FROM TiposRegApuracaoIPI ORDER BY Codigo", iCodigo, sDescricao, iSecao, iLancamento)
    If lErro <> AD_SQL_SUCESSO Then gError 202745

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202746
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set colRegItem = New Collection
        dValor = 0
           
        If iLancamento = MARCADO Then
    
            sDescricaoItem = String(255, 0)
    
            lErro = Comando_Executar(alComando(1), "SELECT Descricao, Valor FROM RegApuracaoIPIItem WHERE FilialEmpresa = ? AND Data BETWEEN ? AND ? AND TipoReg = ? ORDER BY Data", sDescricaoItem, dValorItem, iFilialEmpresa, dtDataInicio, dtDataFim, iCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 202747
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202748
    
            Do While lErro <> AD_SQL_SEM_DADOS
    
                Set objRegItem = New ClassRegApuracaoItem
    
                objRegItem.dValor = dValorItem
                objRegItem.sDescricao = sDescricaoItem
                
                dValor = dValor + dValorItem
                
                colRegItem.Add objRegItem
    
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202749
    
            Loop
        
        Else
            Select Case iCodigo
                Case 1
                    '==========> PRECISA AJUSTAR FUNÇÃO
                    lErro = CF("LivRegES_Le_Entradas", dValorAux, dValor, dtDataInicio, dtDataFim, iFilialEmpresa, alComando(5), LIVREGES_ENTRADA_NACIONAL)
                    If lErro <> SUCESSO Then gError 202750
                Case 2
                    '==========> PRECISA AJUSTAR FUNÇÃO
                    lErro = CF("LivRegES_Le_Entradas", dValorAux, dValor, dtDataInicio, dtDataFim, iFilialEmpresa, alComando(5), LIVREGES_ENTRADA_EXTERNA)
                    If lErro <> SUCESSO Then gError 202751
                Case 3
                    '==========> PRECISA AJUSTAR FUNÇÃO
                    lErro = CF("LivRegES_Le_Saidas", dValorAux, dValor, dtDataInicio, dtDataFim, iFilialEmpresa, alComando(7), LIVREGES_SAIDA_EXTERNA)
                    If lErro <> SUCESSO Then gError 202752
                Case 6
                    dValor = dValores(1) + dValores(2) + dValores(3) + dValores(4) + dValores(5)
                Case 7
                    Set objApuracao = New ClassRegApuracao

                    objApuracao.dtDataInicial = dtDataInicio
                    objApuracao.dtDataFinal = dtDataFim
                    objApuracao.iFilialEmpresa = iFilialEmpresa
                    
                    'Le o Saldo Credor do último Período
                    lErro = CF("ApuracaoIPI_Le_SaldoCredor", objApuracao, alComando(6))
                    If lErro <> SUCESSO And lErro <> 78049 Then gError 202753

                    dValor = objApuracao.dSaldoCredorInicial

                Case 8
                    dValor = dValores(7) + dValores(6)
                Case 9
                    '==========> PRECISA AJUSTAR FUNÇÃO
                    lErro = CF("LivRegES_Le_Saidas", dValorAux, dValor, dtDataInicio, dtDataFim, iFilialEmpresa, alComando(7), LIVREGES_SAIDA_NACIONAL)
                    If lErro <> SUCESSO Then gError 202754
                Case 11
                    dValor = 0 '==========> VER COMO PEGAR O VALOR
                Case 13
                    dValor = dValores(9) + dValores(10) + dValores(11) + dValores(12)
                Case 14
                    dValor = dValores(13)
                Case 15
                    dValor = dValores(8)
                Case 16
                    dValor = dValores(14) - dValores(15)
                Case 17
                    dValor = dValores(15) - dValores(14)
            End Select
            
        End If
        
        iIndice = 1
        dValores(iCodigo) = dValor
        
        bIncluiObs = False
        Select Case iSecao
            Case 0
                iGrupo = 3
                sDescGrupo = "DEMONSTRATIVO DE DÉBITOS"
                If iGrupoAnt <> iGrupo Then bIncluiObs = True
            Case 1
                iGrupo = 1
                sDescGrupo = "DEMONSTRATIVO DE CRÉDITOS"
            Case 2
                iGrupo = 4
                sDescGrupo = "APURAÇÃO DO SALDO"
        End Select
        iGrupoAnt = iGrupo
        
        If bIncluiObs Then
        
            iSeq = iSeq + 1
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelRegApurIPIDet (NumIntRel,Seq,Grupo,DescGrupo,Descricao,ValorSoma,ValorLinha) VALUES (?,?,?,?,?,?,?) ", lNumIntRel, iSeq, 2, "OBSERVAÇÃO", "", -1, -1)
            If lErro <> AD_SQL_SUCESSO Then gError 202755
        
        End If
        
        iSeq = iSeq + 1
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelRegApurIPIDet (NumIntRel,Seq,Grupo,DescGrupo,Descricao,ValorSoma,ValorLinha) VALUES (?,?,?,?,?,?,?) ", lNumIntRel, iSeq, iGrupo, sDescGrupo, Format(iCodigo, "000") & SEPARADOR & sDescricao, dValor, -1)
        If lErro <> AD_SQL_SUCESSO Then gError 202756
        
        For Each objRegItem In colRegItem
    
            iSeq = iSeq + 1
            lErro = Comando_Executar(alComando(2), "INSERT INTO RelRegApurIPIDet (NumIntRel,Seq,Grupo,DescGrupo,Descricao,ValorSoma,ValorLinha) VALUES (?,?,?,?,?,?,?)", lNumIntRel, iSeq, iGrupo, sDescGrupo, "          " & sDescricaoItem, -1, dValorItem)
            If lErro <> AD_SQL_SUCESSO Then gError 202757
    
        Next
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202758
    
    Loop
    
    sCgc = String(STRING_CGC, 0)
    sIE = String(STRING_INSCR_EST, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT CGC, InscricaoEstadual FROM FiliaisEmpresa WHERE FilialEmpresa = ?", sCgc, sIE, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 202759

    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202760
    
    lErro = Comando_Executar(alComando(4), "INSERT INTO RelRegApurIPI (NumIntRel,RazaoEmp,IE, CGC,Folha, DataDe, DataAte) VALUES (?,?,?,?,?,?,?)", lNumIntRel, gsNomeEmpresa, sIE, sCgc, iFolha, dtDataInicio, dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 202761
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 202762
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelRegApurIPI_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelRegApurIPI_Prepara:

    RelRegApurIPI_Prepara = gErr
     
    Select Case gErr

        Case 202742
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 202743
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                   
        Case 202744, 202750 To 202754

        Case 202745, 202746, 202758
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSREGAPURACAOIPI", gErr)

        Case 202747 To 202749
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOIPIITEM", gErr)

        Case 202755, 202756, 202757
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELREGAPURIPIDET", gErr)

        Case 202759, 202760
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", gErr)

        Case 202761
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELREGAPURIPI", gErr)

        Case 202762
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202763)
     
    End Select
     
    'Efetua rollback da transacao
    Call Transacao_Rollback
    
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function LivRegESLinha_RefazerItensNFs(ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'recalcula os registros de LivRegESItensNF para um grupo de nfs
'iFilialEmpresa NAO PODE ser EMPRESA_TODA

Dim lErro As Long, iIndice As Integer, lNumIntDoc As Long, lNumIntDocRegES As Long
Dim tLivRegES As typeLivRegES, lNumIntNF As Long
Dim alComando(1 To 10) As Long
Dim lTransacao As Long
Dim objNFiscal As ClassNFiscal, objLivRegES As ClassLivRegES
Dim objItemNF As ClassItemNF
Dim objLivRegESItemNF As ClassLivRegESItemNF, colLivRegESItemNF As Collection

On Error GoTo Erro_LivRegESLinha_RefazerItensNFs

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81669
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81670

    'tratar nfs que tiveram tiposdocinfo alterados p/passar a aparecer em livros fiscais
    lErro = Comando_Executar(alComando(2), "SELECT NFiscal.NumIntDoc, LivRegES.NumIntDoc FROM NFiscal, TiposDocInfo, LivRegES WHERE LivRegES.NumIntDocOrigem = NFiscal.NumIntDoc AND NFiscal.Filialempresa = ? AND Nfiscal.DataEmissao >= ? AND Nfiscal.DataEmissao <= ? AND TiposDocInfo.ModeloArqIcms =1 AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND NOT EXISTS (SELECT L.NumIntDoc FROM LivREGES L, LivRegESItemNF I WHERE L.NumIntDoc = I.NumIntDocRegES AND NumIntDocORigem = NFiscal.NumIntDoc)", lNumIntNF, lNumIntDocRegES, iFilialEmpresa, dtDataInicial, dtDataFinal)
    If lErro <> SUCESSO Then gError 81729
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81730
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNFiscal = New ClassNFiscal
        objNFiscal.lNumIntDoc = lNumIntNF
        
        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 69402
            
        'Se não encontrou a NFiscal --> erro
        If lErro <> SUCESSO Then gError 69403
            
        'Lê os Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 69395
                
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO Then gError 69396
    
        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError 69397
            
        Set colLivRegESItemNF = New Collection
        
        'Varre os Itens para gerar LivRegLinha, LivRegESNF, LivRegESCadProd
        For Each objItemNF In objNFiscal.ColItensNF
            
            Set objLivRegESItemNF = New ClassLivRegESItemNF
            
            With objItemNF.objTributacaoItemNF
                objLivRegESItemNF.dAliquotaICMS = .dICMSAliquota
                objLivRegESItemNF.dQuantidade = objItemNF.dQuantidade
                objLivRegESItemNF.dValorICMSBase = .dICMSBase
                objLivRegESItemNF.dValorICMSSubstBase = .dICMSSubstBase
                objLivRegESItemNF.dValorIPI = .dIPIValor
                objLivRegESItemNF.dValorProduto = Round(objItemNF.dPrecoUnitario * objItemNF.dQuantidade, 2)
                objLivRegESItemNF.iNumItem = objItemNF.iItem
                objLivRegESItemNF.dValorDescontoDespAcess = Round(objItemNF.dValorDesconto + objItemNF.dValorDescGlobal, 2)
                objLivRegESItemNF.iTipoTribICMS = .iICMSTipo
                objLivRegESItemNF.iTipoTribIPI = .iIPITipo
                objLivRegESItemNF.dAliquotaIPI = .dIPIAliquota
                objLivRegESItemNF.dRedBaseICMS = .dICMSPercRedBase
                objLivRegESItemNF.dRedBaseIPI = .dIPIPercRedBase
                objLivRegESItemNF.iTipoTributacao = .iTipoTributacao
                objLivRegESItemNF.dAliquotaSubst = .dICMSSubstAliquota
                objLivRegESItemNF.dValorSubst = .dICMSSubstValor
                objLivRegESItemNF.dValorICMS = .dICMSValor
                objLivRegESItemNF.dValorIPIBase = .dIPIBaseCalculo
            End With
            
            'Gera número automático para o Item de apuração ICMS
            lErro = CF("Config_ObterNumInt", "FISConfig", "FISC_NUM_INT_LIVREGES_ITEMNF", lNumIntDoc)
            If lErro <> SUCESSO Then gError 69246
            
            objLivRegESItemNF.lNumIntDoc = lNumIntDoc
            objLivRegESItemNF.sCFOP = objItemNF.objTributacaoItemNF.sNaturezaOp
            
            objLivRegESItemNF.lNumIntCadProd = 0
            objLivRegESItemNF.lNumIntDocRegES = lNumIntDocRegES
                
            colLivRegESItemNF.Add objLivRegESItemNF
                    
        Next
            
        For Each objLivRegESItemNF In colLivRegESItemNF
            
            'insere as linhas do Livro Fiscal desta Nota Fiscal
            lErro = Comando_Executar(alComando(3), "INSERT INTO LivRegESItemNF (NumIntDoc, NumIntDocRegES, NumItem, CFOP, NumIntCadProd, Quantidade, ValorProduto, ValorDescontoDespAcess, ValorICMSBase, ValorICMSSubstBase, ValorIPI, AliquotaICMS, TipoTribICMS, TipoTribIPI, ValorICMS, ValorIPIBase, AliquotaIPI, RedBaseICMS, RedBaseIPI, TipoTributacao, AliquotaSubst, ValorSubst) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objLivRegESItemNF.lNumIntDoc, objLivRegESItemNF.lNumIntDocRegES, objLivRegESItemNF.iNumItem, objLivRegESItemNF.sCFOP, objLivRegESItemNF.lNumIntCadProd, objLivRegESItemNF.dQuantidade, objLivRegESItemNF.dValorProduto, objLivRegESItemNF.dValorDescontoDespAcess, objLivRegESItemNF.dValorICMSBase, objLivRegESItemNF.dValorICMSSubstBase, objLivRegESItemNF.dValorIPI, objLivRegESItemNF.dAliquotaICMS, objLivRegESItemNF.iTipoTribICMS, objLivRegESItemNF.iTipoTribIPI, objLivRegESItemNF.dValorICMS, objLivRegESItemNF.dValorIPIBase, objLivRegESItemNF.dAliquotaIPI, objLivRegESItemNF.dRedBaseICMS, objLivRegESItemNF.dRedBaseIPI, objLivRegESItemNF.iTipoTributacao, objLivRegESItemNF.dAliquotaSubst, objLivRegESItemNF.dValorSubst)
            If lErro <> AD_SQL_SUCESSO Then gError 69310
        
        Next
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81732
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81671

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MsgBox ("sucesso")
    
    LivRegESLinha_RefazerItensNFs = SUCESSO
     
    Exit Function
    
Erro_LivRegESLinha_RefazerItensNFs:

    LivRegESLinha_RefazerItensNFs = gErr
     
    Select Case gErr
          
        Case 81729 To 81732
        
        Case 81673, 81674
        
        Case 81669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 81675, 81676, 81677
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150379)
     
    End Select
     
    Call Transacao_Rollback

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function RegApuracaoPISCofins_Grava(ByVal objApurPISCofins As ClassRegApuracaoPISCofins) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RegApuracaoPISCofins_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213004

    lErro = CF("RegApuracaoPISCofins_Grava_EmTrans", objApurPISCofins)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213005

    RegApuracaoPISCofins_Grava = SUCESSO

    Exit Function

Erro_RegApuracaoPISCofins_Grava:

    RegApuracaoPISCofins_Grava = gErr

    Select Case gErr

        Case 213004
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213005
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213006)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RegApuracaoPISCofins_Grava_EmTrans(ByVal objApurPISCofins As ClassRegApuracaoPISCofins) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RegApuracaoPISCofins_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213007
    Next

    'Le a tabelaRegApuracaoPISCofins
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Tipo= ?  AND Ano= ?  AND Mes= ?  AND OrigCred= ?  AND CNPJCedCred= ?  AND CodCred= ? ", _
                0, iAux, objApurPISCofins.iFilialEmpresa, objApurPISCofins.iTipo, objApurPISCofins.iAno, objApurPISCofins.iMes, objApurPISCofins.iOrigCred, objApurPISCofins.sCNPJCedCred, objApurPISCofins.sCodCred)
    If lErro <> AD_SQL_SUCESSO Then gError 213008

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213009

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 213010

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RegApuracaoPISCofins SET VlCredApu= ? , VlCredExtApu= ? , VlTotCredApu= ? , VlCredDescPAAnt= ? , " & _
                "VlCredPerPAAnt= ? , VlCredDCompPAAnt= ? , SdCredDispEFD= ? , VlCredDescEFD= ? , VlCredPerEFD= ? , " & _
                "VlCredDCompEFD= ? , VlCredTrans= ? , VlCredOut= ? , SdCredFim= ? ", alComando(0), _
                objApurPISCofins.dVlCredApu, objApurPISCofins.dVlCredExtApu, objApurPISCofins.dVlTotCredApu, objApurPISCofins.dVlCredDescPAAnt, _
                objApurPISCofins.dVlCredPerPAAnt, objApurPISCofins.dVlCredDCompPAAnt, objApurPISCofins.dSdCredDispEFD, objApurPISCofins.dVlCredDescEFD, objApurPISCofins.dVlCredPerEFD, _
                objApurPISCofins.dVlCredDCompEFD, objApurPISCofins.dVlCredTrans, objApurPISCofins.dVlCredOut, objApurPISCofins.dSdCredFim)
        If lErro <> AD_SQL_SUCESSO Then gError 213011

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_INT_PROX_REGAPURACAOPISCOFINS", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objApurPISCofins.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO RegApuracaoPISCofins( NumIntDoc, FilialEmpresa, Tipo, Ano, Mes, " & _
                "OrigCred, CNPJCedCred, CodCred, VlCredApu, VlCredExtApu, " & _
                "VlTotCredApu, VlCredDescPAAnt, VlCredPerPAAnt, VlCredDCompPAAnt, SdCredDispEFD, " & _
                "VlCredDescEFD, VlCredPerEFD, VlCredDCompEFD, VlCredTrans, VlCredOut, " & _
                "SdCredFim) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objApurPISCofins.lNumIntDoc, objApurPISCofins.iFilialEmpresa, objApurPISCofins.iTipo, objApurPISCofins.iAno, objApurPISCofins.iMes, _
                objApurPISCofins.iOrigCred, objApurPISCofins.sCNPJCedCred, objApurPISCofins.sCodCred, objApurPISCofins.dVlCredApu, objApurPISCofins.dVlCredExtApu, _
                objApurPISCofins.dVlTotCredApu, objApurPISCofins.dVlCredDescPAAnt, objApurPISCofins.dVlCredPerPAAnt, objApurPISCofins.dVlCredDCompPAAnt, objApurPISCofins.dSdCredDispEFD, _
                objApurPISCofins.dVlCredDescEFD, objApurPISCofins.dVlCredPerEFD, objApurPISCofins.dVlCredDCompEFD, objApurPISCofins.dVlCredTrans, objApurPISCofins.dVlCredOut, _
                objApurPISCofins.dSdCredFim)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213012

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegApuracaoPISCofins_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RegApuracaoPISCofins_Grava_EmTrans:

    RegApuracaoPISCofins_Grava_EmTrans = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 213007
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213008, 213009
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOPISCOFINS", gErr)

        Case 213010
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_REGAPURACAOPISCOFINS", gErr)

        Case 2130011
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_REGAPURACAOPISCOFINS", gErr)

        Case 213012
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REGAPURACAOPISCOFINS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213013)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function RegApuracaoPISCofins_Exclui(ByVal objApurPISCofins As ClassRegApuracaoPISCofins) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RegApuracaoPISCofins_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213014

    lErro = CF("RegApuracaoPISCofins_Exclui_EmTrans", objApurPISCofins)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213015

    RegApuracaoPISCofins_Exclui = SUCESSO

    Exit Function

Erro_RegApuracaoPISCofins_Exclui:

    RegApuracaoPISCofins_Exclui = gErr

    Select Case gErr

        Case 213014
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213015
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213016)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RegApuracaoPISCofins_Exclui_EmTrans(ByVal objApurPISCofins As ClassRegApuracaoPISCofins) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_RegApuracaoPISCofins_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213017
    Next

    'Le a tabelaRegApuracaoPISCofins
    lErro = Comando_ExecutarPos(alComando(0), "SELECT FilialEmpresa FROM RegApuracaoPISCofins WHERE FilialEmpresa = ? AND Tipo= ?  AND Ano= ?  AND Mes= ?  AND OrigCred= ?  AND CNPJCedCred= ?  AND CodCred= ? ", _
                0, iAux, objApurPISCofins.iFilialEmpresa, objApurPISCofins.iTipo, objApurPISCofins.iAno, objApurPISCofins.iMes, objApurPISCofins.iOrigCred, objApurPISCofins.sCNPJCedCred, objApurPISCofins.sCodCred)
    If lErro <> AD_SQL_SUCESSO Then gError 213018

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213019

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 213020

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 213021

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RegApuracaoPISCofins", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 213022

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegApuracaoPISCofins_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_RegApuracaoPISCofins_Exclui_EmTrans:

    RegApuracaoPISCofins_Exclui_EmTrans = gErr

    Select Case gErr

        Case 213017
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213018, 213019
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGAPURACAOPISCOFINS", gErr)

        Case 213020
            Call Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOPISCOFINS_NAO_CADASTRADO", gErr, objApurPISCofins.iTipo, objApurPISCofins.iAno, objApurPISCofins.iMes, objApurPISCofins.iOrigCred, objApurPISCofins.sCNPJCedCred, objApurPISCofins.sCodCred)

        Case 213021
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_REGAPURACAOPISCOFINS", gErr)

        Case 213022
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REGAPURACAOPISCOFINS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213023)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Gerar_SINTEGRA_PAFECF(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long

Dim lErro As Long, sConteudo As String
Dim objFilial As New AdmFiliais, sBuffer As String
Dim sArqCaixaConfig As String
Dim sArquivo1 As String
Dim iFilialEmpresa As Integer
Dim iCodEmpresa As Integer
Dim iCodCaixa As Integer
Dim sRetorno As String
Dim lTamanho As Long
Dim sDirMVTEF As String
Dim sTel As String
Dim sTel1 As String
Dim objInfoArqICMS As New ClassInfoArqICMS
Dim iPos As Integer
Dim objInfoArqICMSAux As New ClassInfoArqICMS
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long


On Error GoTo Erro_Gerar_SINTEGRA_PAFECF

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69978
    Next

    sArqCaixaConfig = "c:\windows\caixaconfig.ini"

    sArquivo1 = Dir(sArqCaixaConfig)

    If Len(sArquivo1) = 0 Then

        sArqCaixaConfig = "\\tsclient\c\windows\caixaconfig.ini"

        sArquivo1 = Dir(sArqCaixaConfig)

        'o arquivo caixaconfig.ini nao foi encontrado
        If Len(sArquivo1) = 0 Then gError 210110
    
    End If
    
    lTamanho = 150
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "FilialEmpresa", CONSTANTE_ERRO, sRetorno, lTamanho, sArqCaixaConfig)
    
    If sRetorno = CStr(CONSTANTE_ERRO) Then gError 210111
    
    If sRetorno = String(lTamanho, 0) Then gError 210112
    
    iFilialEmpresa = StrParaInt(sRetorno)

    If iFilialEmpresa <> giFilialEmpresa Then gError 210116
    
    'pega o código da empresa
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "CodEmpresa", CONSTANTE_ERRO, sRetorno, lTamanho, sArqCaixaConfig)
    
    If sRetorno = CStr(CONSTANTE_ERRO) Then gError 210113
    
    If sRetorno = String(lTamanho, 0) Then gError 210114
    
    iCodEmpresa = StrParaInt(sRetorno)
    
    
    If glEmpresa <> iCodEmpresa Then gError 210115
    
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "CodCaixa", CONSTANTE_ERRO, sRetorno, lTamanho, sArqCaixaConfig)
    
    If sRetorno = CStr(CONSTANTE_ERRO) Then gError 210107
    
    If sRetorno = String(lTamanho, 0) Then gError 210108
    
    iCodCaixa = StrParaInt(sRetorno)

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirMVTEF", CONSTANTE_ERRO, sRetorno, lTamanho, sArqCaixaConfig)
    If sRetorno <> String(lTamanho, 0) Then sDirMVTEF = StringZ(sRetorno)
    
    If right(sDirMVTEF, 1) <> "\" Then sDirMVTEF = sDirMVTEF & "\"

    
    objFilial.iCodFilial = giFilialEmpresa
    
    'Lê a Filial Empresa
    lErro = CF("FilialEmpresa_Le", objFilial)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 210122

    'Se não encontrou a Filial da Empresa --> Erro
    If lErro <> SUCESSO Then gError 210123
    

    objInfoArqICMS.iFilialEmpresa = giFilialEmpresa
    objInfoArqICMS.dtDataFinal = dtDataFinal
    objInfoArqICMS.dtDataInicial = dtDataInicial
    
    
    objInfoArqICMS.lNumero = objFilial.objEndereco.lNumero
    objInfoArqICMS.sComplemento = objFilial.objEndereco.sComplemento
    objInfoArqICMS.sContato = objFilial.objEndereco.sContato
    If Len(objFilial.objEndereco.sLogradouro) > 0 Then
        objInfoArqICMS.sLogradouro = objFilial.objEndereco.sLogradouro
    Else
        objInfoArqICMS.sLogradouro = objFilial.objEndereco.sEndereco
    End If
    
    objInfoArqICMS.sNomeArquivo = sDirMVTEF & "SINTEGRA.TXT"
    If Len(objFilial.objEndereco.sTelNumero1) > 0 Then
    
        sTel1 = objFilial.objEndereco.sTelNumero1
    
    
    Else
    
        sTel1 = objFilial.objEndereco.sTelefone1
        
    End If
    
    iPos = InStr(1, sTel1, "-")
    If iPos <> 0 Then
        sTel = left(Trim(sTel1), iPos - 1) & right(Trim(sTel1), Len(Trim(sTel1)) - iPos)
    Else
        sTel = Trim(sTel1)
    End If
    
    objInfoArqICMS.sTelContato = sTel
    objInfoArqICMS.sNome = gsNomeEmpresa
    objInfoArqICMS.dtReg74DataInicial = DATA_NULA
    objInfoArqICMS.dtReg74DataFinal = DATA_NULA
    objInfoArqICMS.iGeraSaidas = 1
    objInfoArqICMS.iGeraEntradas = 1
    objInfoArqICMS.bIntegral = True
    objInfoArqICMS.sFinalidade = "1"  'normal
    objInfoArqICMS.iGeraReg54 = 1
    objInfoArqICMS.iGeraReg51 = 1
    
    'Verifica se existe um Arquivo de ICMS com as datas sobrepondo
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM InfoArqICMS WHERE FilialEmpresa = ? AND Integral = ? AND ((DataInicial >= ? AND DataInicial <= ?) OR (DataFinal >= ? AND DataFinal <= ?) OR (DataInicial <= ? AND DataFinal >= ?))", lNumIntDoc, objInfoArqICMS.iFilialEmpresa, IIf(objInfoArqICMS.bIntegral, 1, 0), objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 210961

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210962

    Do While lErro = AD_SQL_SUCESSO

        objInfoArqICMSAux.lNumIntDoc = lNumIntDoc

        'Exclui a apuração de ICMS
        lErro = InfoArqICMS_Exclui_Trans(objInfoArqICMSAux)
        If lErro <> SUCESSO Then gError 210963

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210964

    Loop

    lErro = CF("GeracaoArqICMS_Limpar_TabelasAuxiliares")
    If lErro <> SUCESSO Then gError 210987

    lErro = Gerar_Arquivo_ICMS_FIS_Trans(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 210117
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Gerar_SINTEGRA_PAFECF = SUCESSO

    Exit Function

Erro_Gerar_SINTEGRA_PAFECF:

    Gerar_SINTEGRA_PAFECF = gErr

    Select Case gErr

        Case 210110
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_NAO_ENCONTRADO", gErr)
        
        Case 210111, 210112
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_FILIALEMPRESA", gErr)

        Case 210113, 210114
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_CODEMPRESA", gErr)

        Case 210115
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_EMPRESA_DIFERENTE", gErr)
        
        Case 210116
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_FILIALEMPRESA_DIFERENTE", gErr)
        
        Case 210117, 210122, 210123, 210963, 210987
        
        Case 210961, 210962, 210964
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOARQICMS", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210118)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Gerar_SPEDFISCAL_PAFECF(ByVal dtDataInicio As Date, ByVal dtDataFim As Date) As Long

Dim lErro As Long, sConteudo As String
Dim objFilial As New AdmFiliais, sBuffer As String
Dim lTransacao As Long, lComando As Long
Dim sArqCaixaConfig As String
Dim sArquivo1 As String
Dim iFilialEmpresa As Integer
Dim iCodEmpresa As Integer
Dim iCodCaixa As Integer
Dim sRetorno As String
Dim lTamanho As Long
Dim sDirMVTEF As String
Dim sDiretorio As String
Dim sNomeArqParam As String
Dim objTela As Object
Dim iIncluiRegInv As Integer
Dim iMotivoRegInv As Integer
Dim dtDataInv As Date



On Error GoTo Erro_Gerar_SPEDFISCAL_PAFECF


    sArqCaixaConfig = "c:\windows\caixaconfig.ini"

    sArquivo1 = Dir(sArqCaixaConfig)

    If Len(sArquivo1) = 0 Then

        sArqCaixaConfig = "\\tsclient\c\windows\caixaconfig.ini"

        sArquivo1 = Dir(sArqCaixaConfig)

        'o arquivo caixaconfig.ini nao foi encontrado
        If Len(sArquivo1) = 0 Then gError 210127
    
    End If
    
    lTamanho = 150
    
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_EMPRESA, "FilialEmpresa", CONSTANTE_ERRO, sRetorno, lTamanho, sArqCaixaConfig)
    
    If sRetorno = CStr(CONSTANTE_ERRO) Then gError 210128
    
    If sRetorno = String(lTamanho, 0) Then gError 210129
    
    iFilialEmpresa = StrParaInt(sRetorno)

    If iFilialEmpresa <> giFilialEmpresa Then gError 210130
    

    lTamanho = 255
    sRetorno = String(lTamanho, 0)
    
    Call GetPrivateProfileString(APLICACAO_DADOS, "DirMVTEF", CONSTANTE_ERRO, sRetorno, lTamanho, sArqCaixaConfig)
    If sRetorno <> String(lTamanho, 0) Then sDirMVTEF = StringZ(sRetorno)
    
    If right(sDirMVTEF, 1) <> "\" Then sDirMVTEF = sDirMVTEF & "\"

    sDiretorio = sDirMVTEF & "SPEDFISCAL_PAFECF.TMP"
    
    lErro = CF("Gera_Sped_Fiscal_Trans", sDiretorio, iFilialEmpresa, dtDataInicio, dtDataFim, objTela, iIncluiRegInv, iMotivoRegInv, dtDataInv)
    If lErro <> SUCESSO Then gError 210132
    
    Name sDiretorio As sDirMVTEF & "SPEDFISCAL_PAFECF.TXT"
    
    Gerar_SPEDFISCAL_PAFECF = SUCESSO

    Exit Function

Erro_Gerar_SPEDFISCAL_PAFECF:

    Gerar_SPEDFISCAL_PAFECF = gErr

    Select Case gErr

        Case 210127
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_NAO_ENCONTRADO", gErr)
        
        Case 210128, 210129
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_FILIALEMPRESA", gErr)

        Case 210130
            Call Rotina_Erro(vbOKOnly, "ERRO_CAIXACONFIG_FILIALEMPRESA_DIFERENTE", gErr)
        
        Case 210131, 210132
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210133)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Gerar_Arquivo_ICMS_FIS(objInfoArqICMS As ClassInfoArqICMS) As Long

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Gerar_Arquivo_ICMS_FIS

    'abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 69792

    lErro = Gerar_Arquivo_ICMS_FIS_Trans(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 210134
    
    'Insere no BD todas as Alterações
    Call Transacao_Commit

    Gerar_Arquivo_ICMS_FIS = SUCESSO

    Exit Function

Erro_Gerar_Arquivo_ICMS_FIS:

    Gerar_Arquivo_ICMS_FIS = gErr

    Select Case gErr

        Case 69792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 210134
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160734)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Gerar_Arquivo_ICMS_FIS_Trans(objInfoArqICMS As ClassInfoArqICMS) As Long

Dim lErro As Long, sConteudo As String
Dim objFilial As New AdmFiliais, sBuffer As String
Dim lComando As Long

On Error GoTo Erro_Gerar_Arquivo_ICMS_FIS_Trans

    'Limpa as tabelas que são usadas para ordenação
    lErro = CF("GeracaoArqICMS_Limpar_TabelasAuxiliares")
    If lErro <> SUCESSO Then gError 69797
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69792
    
    'Quarda os dados da Filial Empresa
    objFilial.iCodFilial = giFilialEmpresa

    lErro = CF("FilialEmpresa_Le", objFilial)
    If lErro <> SUCESSO Then gError 69793

    objInfoArqICMS.sBairro = left(Trim(objFilial.objEndereco.sBairro), 15)
    objInfoArqICMS.sCEP = objFilial.objEndereco.sCEP
    objInfoArqICMS.sCgc = objFilial.sCgc
    objInfoArqICMS.sInscricaoEstadual = objFilial.sInscricaoEstadual
    objInfoArqICMS.sMunicipio = objFilial.objEndereco.sCidade
    objInfoArqICMS.sUF = objFilial.objEndereco.sSiglaEstado

    'Grava o Arquivo no BD
    lErro = CF("InfoArqICMS_Grava", objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69794

    sConteudo = String(255, 0)
    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM FisConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, "FISC_SINTEGRA_GERA_REG_51", giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 69792

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69792

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then
        objInfoArqICMS.iGeraReg51 = IIf(objFilial.iContribuinteIPI = 0, 0, 1)
    Else
        objInfoArqICMS.iGeraReg51 = StrParaInt(sConteudo)
    End If
    
    'Percorrer os dados dos Registros de Entrada e Saída
    lErro = Processar_RegistrosES(objInfoArqICMS, objFilial)
    If lErro <> SUCESSO Then gError 69795

    lErro = Processar_SubstTributaria(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69795
    
    lErro = Processar_ConhecFrete_Transp(objFilial, objInfoArqICMS)
    If lErro <> SUCESSO Then gError 81762
    
    lErro = Processar_RegsInventario(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 130028
    
    'Insere os dados no Arquivo de ICMS
    If objInfoArqICMS.bIntegral Then
        
        If gcolModulo.Ativo(MODULO_LOJA) = MODULO_ATIVO Then
        
            '****incluído por Cyntia******
            lErro = Processar_CuponsFiscais(objInfoArqICMS)
            If lErro <> SUCESSO Then gError 112081
            '****até aqui*****************
        
        End If
        
        lErro = InserirRegsArqICMS1(objFilial, objInfoArqICMS)
    Else
        lErro = InserirRegsArqICMS(objFilial, objInfoArqICMS)
    End If
    If lErro <> SUCESSO Then gError 69796
        
    'Limpa as tabelas que são usadas para ordenação
    lErro = CF("GeracaoArqICMS_Limpar_TabelasAuxiliares")
    If lErro <> SUCESSO Then gError 69797
    
    Call Comando_Fechar(lComando)
    
    Gerar_Arquivo_ICMS_FIS_Trans = SUCESSO

    Exit Function

Erro_Gerar_Arquivo_ICMS_FIS_Trans:

    Gerar_Arquivo_ICMS_FIS_Trans = gErr

    Select Case gErr

        Case 69792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 69793, 69794, 69795, 69796, 69797, 81762, 112081, 130028
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160734)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Processar_RegistrosES(objInfoArqICMS As ClassInfoArqICMS, objFilial As AdmFiliais) As Long
'Processa os Registros de Entrada e Saidas

Dim lErro As Long
Dim tLivRegES As typeLivRegES
Dim tLivRegESItemNF As typeLivRegESItemNF
Dim tLivRegESCadProd As typeLivRegESCadProd
Dim tLivRegESLinha As typeLivRegESLinha
Dim objGeracaoArqICMSProd As New ClassGeracaoArqICMSProd
Dim iIndice As Integer, bAchou As Boolean
Dim alComando(1 To 18) As Long, objTipoTributacao As New ClassTipoDeTributacaoMovto
Dim objLivRegESLinha50 As ClassLivRegESLinha, colLivRegESLinha50 As Collection
Dim objLivRegESLinha51 As ClassLivRegESLinha, colLivRegESLinha51 As Collection
Dim bEntra As Boolean, sSQL1 As String, sSQL2 As String
Dim iStatusLivro As Integer, iSituacao As Integer
Dim iItem As Integer, sProduto As String, dQuantidade As Double, dPrecoTotal As Double, sNaturezaOp As String, dValorTotal As Double, sCST As String

On Error GoTo Erro_Processar_RegistrosES

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69798
    Next

    'Inicializa os Buffers
    tLivRegES.sSerie = String(STRING_SERIE, 0)
    tLivRegES.sCgc = String(STRING_CGC, 0)
    tLivRegES.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tLivRegES.sUF = String(STRING_ESTADO, 0)
    tLivRegES.sNome = String(STRING_NOME_REGES, 0)

    sSQL1 = "SELECT NumIntDoc, FilialEmpresa, NumIntDocLivFechado, TipoNumIntDocOrigem, NumIntDocOrigem, StatusLivro, Tipo, CGC, InscricaoEstadual, IEIsento, Nome, DataEmissao, Data, UF, Serie, NumNotaFiscal, Situacao, Modelo, Emitente, Destinatario, Origem, NumIntEmitente, NumIntRegApuracaoICMS, NumIntRegApuracaoIPI, CIF_FOB FROM LivRegES WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ?"
    If objInfoArqICMS.iGeraEntradas = 0 Then sSQL1 = sSQL1 & " AND Tipo <> 0"
    If objInfoArqICMS.iGeraSaidas = 0 Then sSQL1 = sSQL1 & " AND Tipo <> 1"
    sSQL2 = sSQL1 & " AND UF <> ?"
    
    'Le os Registros de Entrada e de Saída
    If objInfoArqICMS.bIntegral Then
        lErro = Comando_ExecutarPos(alComando(1), sSQL1, 0, _
            tLivRegES.lNumIntDoc, tLivRegES.iFilialEmpresa, tLivRegES.lNumIntLivroFechado, tLivRegES.iTipoNumIntDocOrigem, tLivRegES.lNumIntDocOrigem, iStatusLivro, tLivRegES.iTipo, tLivRegES.sCgc, tLivRegES.sInscricaoEstadual, tLivRegES.iIEIsento, tLivRegES.sNome, tLivRegES.dtDataEmissao, tLivRegES.dtData, tLivRegES.sUF, tLivRegES.sSerie, tLivRegES.lNumNotaFiscal, iSituacao, tLivRegES.iModelo, tLivRegES.iEmitente, tLivRegES.iDestinatario, tLivRegES.iOrigem, tLivRegES.lNumIntEmitente, tLivRegES.lNumIntRegApuracaoICMS, tLivRegES.lNumIntRegApuracaoIPI, tLivRegES.iCIF_FOB, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    Else
        'filtra p/só pegar as saidas p/UFs diferentes da UF da filial da empresa
        lErro = Comando_ExecutarPos(alComando(1), sSQL2, 0, _
            tLivRegES.lNumIntDoc, tLivRegES.iFilialEmpresa, tLivRegES.lNumIntLivroFechado, tLivRegES.iTipoNumIntDocOrigem, tLivRegES.lNumIntDocOrigem, iStatusLivro, tLivRegES.iTipo, tLivRegES.sCgc, tLivRegES.sInscricaoEstadual, tLivRegES.iIEIsento, tLivRegES.sNome, tLivRegES.dtDataEmissao, tLivRegES.dtData, tLivRegES.sUF, tLivRegES.sSerie, tLivRegES.lNumNotaFiscal, iSituacao, tLivRegES.iModelo, tLivRegES.iEmitente, tLivRegES.iDestinatario, tLivRegES.iOrigem, tLivRegES.lNumIntEmitente, tLivRegES.lNumIntRegApuracaoICMS, tLivRegES.lNumIntRegApuracaoIPI, tLivRegES.iCIF_FOB, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, objInfoArqICMS.sUF)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 69799

    'Busca o Primeiro Registro de Entrada ou Saída
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69800

    'Loop por Registros de Entrada ou Saída
    Do While lErro = SUCESSO
        
        lErro = CF("Sintegra_SQLRegES", iStatusLivro, iSituacao)
        If lErro <> SUCESSO Then gError 184010
        
        If iStatusLivro <> STATUS_FIS_ALTERADA_EXCLUIDA And iStatusLivro <> STATUS_FIS_ORIGINAL_EXCLUIDA Then
        
            tLivRegES.iStatusLivro = iStatusLivro
            tLivRegES.iSituacao = iSituacao
            
            If tLivRegES.lNumNotaFiscal > 999999 Then tLivRegES.lNumNotaFiscal = tLivRegES.lNumNotaFiscal Mod 1000000
            
            If objInfoArqICMS.bIntegral Then
                objInfoArqICMS.sUFDestino = objInfoArqICMS.sUF
            Else
                objInfoArqICMS.sUFDestino = tLivRegES.sUF
            End If
    
            'se for p/gerar reg tipo 50 das notas "normais"
            Select Case tLivRegES.iModelo
            
                Case 1, 3, 6, 7, 22, 55
            
                    If tLivRegES.iSituacao = 0 Then
            
                        If (tLivRegES.iModelo = 1 Or tLivRegES.iModelo = 55) And objInfoArqICMS.iGeraReg54 = 1 Then
                        
                            tLivRegESItemNF.sCFOP = String(STRING_NATUREZAOP_CODIGO, 0)
                    
                            'Lê os itens da Nota que estão no Reg. de Entrada e Saída
                            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, NumIntDocRegES, NumItem, CFOP, NumIntCadProd, Quantidade, ValorProduto, ValorDescontoDespAcess, ValorICMSBase, ValorICMSSubstBase, ValorIPI, AliquotaICMS, TipoTributacao, TipoTribICMS FROM LivRegESItemNF WHERE NumIntDocRegES = ? ORDER BY NumItem ASC", _
                                    tLivRegESItemNF.lNumIntDoc, tLivRegESItemNF.lNumIntDocRegES, tLivRegESItemNF.iNumItem, tLivRegESItemNF.sCFOP, tLivRegESItemNF.lNumIntCadProd, tLivRegESItemNF.dQuantidade, tLivRegESItemNF.dValorProduto, tLivRegESItemNF.dValorDescontoDespAcess, tLivRegESItemNF.dValorICMSBase, tLivRegESItemNF.dValorICMSSubstBase, tLivRegESItemNF.dValorIPI, tLivRegESItemNF.dAliquotaICMS, tLivRegESItemNF.iTipoTributacao, tLivRegESItemNF.iTipoTribICMS, tLivRegES.lNumIntDoc)
                            If lErro <> AD_SQL_SUCESSO Then gError 69801
                    
                            lErro = Comando_BuscarPrimeiro(alComando(2))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69802
                    
                            'Loop por Item de Nota Fiscal
                            Do While lErro = SUCESSO
                            
                                objTipoTributacao.iTipo = tLivRegESItemNF.iTipoTributacao
                                
                                lErro = CF("TipoTributacao_Le", objTipoTributacao)
                                If lErro = SUCESSO Then
                                
                                    'o cfop 1201 foi incluido apenas pq o default de devolucao de vendas ainda nao é alterado automaticamente qdo tem ST
                                    If objTipoTributacao.iICMSCredita = 0 Or objFilial.iSuperSimples <> 0 Or (tLivRegES.iTipo = TIPO_REGES_ENTRADA And tLivRegESItemNF.dValorICMSSubstBase <> 0 And _
                                        InStr("1201,1410,1411,1660,1661,1662,2410,2411,2660,2661,2662", tLivRegESItemNF.sCFOP) = 0) Then
                                        
                                        With tLivRegESItemNF
                                            .dValorICMSBase = 0
                                            .dRedBaseICMS = 0
                                            .dAliquotaICMS = 0
                                            .dValorICMS = 0
                                            .dValorICMSSubstBase = 0
                                            .dAliquotaSubst = 0
                                            .dValorSubst = 0
                                        End With
                                    End If
                                    
                                    If objTipoTributacao.iIPICredita = 0 Then
                                        tLivRegESItemNF.dValorIPI = 0
                                    End If
                                
                                End If
                                
                                If tLivRegESItemNF.iNumItem < 990 Then
                                                        
                                    'Inicializa os Buffers
                                    tLivRegESCadProd.sSiglaUM = String(STRING_ARQICMS_UM, 0)
                                    tLivRegESCadProd.sDescricao = String(STRING_ARQICMS_DESCRICAO_PRODUTO, 0)
                                    tLivRegESCadProd.sProduto = String(STRING_PRODUTO, 0)
                                    tLivRegESCadProd.sCodigoNCM = String(STRING_CODIGONCM, 0)
                                    tLivRegESCadProd.sSituacaoTrib = String(STRING_SITUACAOTRIB, 0)
                        
                                    'Lê o cadastro Produto nos Livros Fiscais
                                    lErro = Comando_Executar(alComando(3), "SELECT L2.NumIntDoc, L2.Produto, L2.DataInicial, L2.DataFinal, L2.CodigoNCM, LEFT(L2.Descricao, 53) Descricao, L2.SiglaUM, L2.SituacaoTrib, L2.AliquotaIPI, L2.AliquotaICMS, L2.ReducaoBaseCalculoICMS, L2.BaseCalculoICMSSubst FROM LivRegESCadProd L1, LivRegESCadProd L2 WHERE L1.Produto = L2.Produto AND L1.NumIntDoc = ? AND L2.DataInicial <= ? ORDER BY L2.DataInicial DESC", _
                                            tLivRegESCadProd.lNumIntDoc, tLivRegESCadProd.sProduto, tLivRegESCadProd.dtDataInicial, tLivRegESCadProd.dtDataFinal, tLivRegESCadProd.sCodigoNCM, tLivRegESCadProd.sDescricao, tLivRegESCadProd.sSiglaUM, tLivRegESCadProd.sSituacaoTrib, tLivRegESCadProd.dAliquotaIPI, tLivRegESCadProd.dAliquotaICMS, tLivRegESCadProd.dReducaoBaseCalculoICMS, tLivRegESCadProd.dBaseCalculoICMSSubst, tLivRegESItemNF.lNumIntCadProd, objInfoArqICMS.dtDataFinal)
                                    If lErro <> AD_SQL_SUCESSO Then gError 69803
                        
                                    lErro = Comando_BuscarPrimeiro(alComando(3))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69804
                        
                                    'Se não encontrar --> Erro
                                    If lErro = AD_SQL_SEM_DADOS Then gError 69805
                        
                                    Do While lErro <> AD_SQL_SEM_DADOS
                                    
                                        'inicializa a variavel que vai dizer se o o registro serah gravado ou nao
                                        bEntra = True
                                    
                                        'inicio da barreira
                                        'se a data nao for data_nula
                                        If tLivRegESCadProd.dtDataFinal <> DATA_NULA Then
                                            
                                            'se data final do registro for menor a data inicial do periodo
                                            If tLivRegESCadProd.dtDataFinal < objInfoArqICMS.dtDataInicial Then
                                            
                                                'se o registro nao possui data final como nula e
                                                'se a data final dele eh menor que a data do periodo
                                                'significa que ele nao deve ser gravado... logo
                                                'ele nao deve entrar...
                                                bEntra = False
                                        
                                            End If
                                        
                                        End If
                                        
                                        'se for pra gravar...
                                        If bEntra = True Then
                                    
                                            'e Grava o prototipo do Tipo 75 no BD para evitar repeticoes
                                            objGeracaoArqICMSProd.lNumIntLivRegESCadProd = tLivRegESCadProd.lNumIntDoc
                                
                                            'Só grava o tipo75 se ele ainda não estiver no BD
                                            lErro = CF("GeracaoArqICMSProd_Grava", objInfoArqICMS, objGeracaoArqICMSProd, alComando(6), alComando(7))
                                            If lErro <> SUCESSO Then gError 69808
                                
                                        End If
                                        
                                        If tLivRegESCadProd.dtDataInicial < objInfoArqICMS.dtDataInicial Then Exit Do
                                            
                                        lErro = Comando_BuscarProximo(alComando(3))
                                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1
                                        
                                    Loop
                                
                                Else
                                    tLivRegESCadProd.sProduto = ""
                                End If
                                
                                Call MontarReg54(objInfoArqICMS, tLivRegES, tLivRegESItemNF, tLivRegESCadProd, alComando(13))
                    
                                'Conta o registro 54
                                objInfoArqICMS.lTotalReg54 = objInfoArqICMS.lTotalReg54 + 1
                    
                                'e Grava no BD o Tipo 54
                                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 54, alComando(5))
                                If lErro <> SUCESSO Then gError 69807
                    
                                'Busca o Proximo Item no BD
                                lErro = Comando_BuscarProximo(alComando(2))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69809
                    
                            Loop
                        
                        End If
                    
                    End If
                    
                    Set colLivRegESLinha50 = New Collection
                    Set colLivRegESLinha51 = New Collection
                    
                    tLivRegESLinha.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
                    tLivRegESLinha.sObservacaoLivFisc = String(STRING_OBSERVACAO, 0)
                    tLivRegESLinha.sClassifContabil = String(STRING_CONTA, 0)
            
                    'Lê as Linhas dos Livro para gerar o Tipo 50 e 51
                    lErro = Comando_Executar(alComando(8), "SELECT NumIntDoc, NumIntDocRegES, NaturezaOp, ValorTotal, ValorICMSBase, ValorICMS, ValorICMSIsentoNaoTrib, ValorICMSOutras, ValorICMSSubstBase, ValorICMSSubstRet, AliquotaICMS, ValorDespAcess, ValorIPI, AliquotaIPI, ValorIPIBase, ValorIPIIsentoNaoTrib, ValorIPIOutras, ClassifContabil, ValorContabil, ObservacaoLivFisc FROM LivRegESLinha WHERE NumIntDocRegES = ? ORDER BY NumIntDoc ASC", tLivRegESLinha.lNumIntDoc, tLivRegESLinha.lNumIntDocRegES, tLivRegESLinha.sNaturezaOp, tLivRegESLinha.dValorTotal, tLivRegESLinha.dValorICMSBase, tLivRegESLinha.dValorICMS, tLivRegESLinha.dValorICMSIsentoNaoTrib, tLivRegESLinha.dValorICMSOutras, tLivRegESLinha.dValorICMSSubstBase, tLivRegESLinha.dValorICMSSubstRet, tLivRegESLinha.dAliquotaICMS, tLivRegESLinha.dValorDespAcess, tLivRegESLinha.dValorIPI, tLivRegESLinha.dAliquotaIPI, tLivRegESLinha.dValorIPIBase, tLivRegESLinha.dValorIPIIsentoNaoTrib, _
                    tLivRegESLinha.dValorIPIOutras, tLivRegESLinha.sClassifContabil, tLivRegESLinha.dValorContabil, tLivRegESLinha.sObservacaoLivFisc, tLivRegES.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 69810
            
                    lErro = Comando_BuscarPrimeiro(alComando(8))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69811
            
                    'Loop por Linha inserindo no BD
                    Do While lErro = SUCESSO
            
                        bAchou = False
                        For Each objLivRegESLinha50 In colLivRegESLinha50
                    
                            If objLivRegESLinha50.sNaturezaOp = tLivRegESLinha.sNaturezaOp And objLivRegESLinha50.dAliquotaICMS = tLivRegESLinha.dAliquotaICMS Then
                            
                                objLivRegESLinha50.dValorTotal = Round(objLivRegESLinha50.dValorTotal + tLivRegESLinha.dValorTotal, 2)
                                objLivRegESLinha50.dValorICMSBase = Round(objLivRegESLinha50.dValorICMSBase + tLivRegESLinha.dValorICMSBase, 2)
                                objLivRegESLinha50.dValorICMS = Round(objLivRegESLinha50.dValorICMS + tLivRegESLinha.dValorICMS, 2)
                                objLivRegESLinha50.dValorICMSIsentoNaoTrib = Round(objLivRegESLinha50.dValorICMSIsentoNaoTrib + tLivRegESLinha.dValorICMSIsentoNaoTrib, 2)
                                objLivRegESLinha50.dValorICMSOutras = Round(objLivRegESLinha50.dValorICMSOutras + tLivRegESLinha.dValorICMSOutras, 2)
                                
                                bAchou = True
                                Exit For
                                
                            End If
                    
                        Next
                        
                        If bAchou = False Then
                        
                            Set objLivRegESLinha50 = New ClassLivRegESLinha
                        
                            With objLivRegESLinha50
                            
                                .sNaturezaOp = tLivRegESLinha.sNaturezaOp
                                .dAliquotaICMS = tLivRegESLinha.dAliquotaICMS
                                .dValorTotal = tLivRegESLinha.dValorTotal
                                .dValorICMSBase = tLivRegESLinha.dValorICMSBase
                                .dValorICMS = tLivRegESLinha.dValorICMS
                                .dValorICMSIsentoNaoTrib = tLivRegESLinha.dValorICMSIsentoNaoTrib
                                .dValorICMSOutras = tLivRegESLinha.dValorICMSOutras
                                                            
                            End With
                            
                            Call colLivRegESLinha50.Add(objLivRegESLinha50)
                            
                        End If
                        
                        If (tLivRegES.iModelo = 1 Or tLivRegES.iModelo = 55) And objInfoArqICMS.iGeraReg51 <> 0 Then
                        
                            bAchou = False
                            For Each objLivRegESLinha51 In colLivRegESLinha51
                        
                                If objLivRegESLinha51.sNaturezaOp = tLivRegESLinha.sNaturezaOp Then
                                
                                    objLivRegESLinha51.dValorTotal = Round(objLivRegESLinha51.dValorTotal + tLivRegESLinha.dValorTotal, 2)
                                    objLivRegESLinha51.dValorIPI = Round(objLivRegESLinha51.dValorIPI + tLivRegESLinha.dValorIPI, 2)
                                    objLivRegESLinha51.dValorIPIIsentoNaoTrib = Round(objLivRegESLinha51.dValorIPIIsentoNaoTrib + tLivRegESLinha.dValorIPIIsentoNaoTrib, 2)
                                    objLivRegESLinha51.dValorIPIOutras = Round(objLivRegESLinha51.dValorIPIOutras + tLivRegESLinha.dValorIPIOutras, 2)
                                    
                                    bAchou = True
                                    Exit For
                                    
                                End If
                        
                            Next
                            
                            If bAchou = False Then
                            
                                Set objLivRegESLinha51 = New ClassLivRegESLinha
                            
                                With objLivRegESLinha51
                                
                                    .sNaturezaOp = tLivRegESLinha.sNaturezaOp
                                    .dValorTotal = tLivRegESLinha.dValorTotal
                                    .dValorIPI = tLivRegESLinha.dValorIPI
                                    .dValorIPIIsentoNaoTrib = tLivRegESLinha.dValorIPIIsentoNaoTrib
                                    .dValorIPIOutras = tLivRegESLinha.dValorIPIOutras
                                                                
                                End With
                                
                                Call colLivRegESLinha51.Add(objLivRegESLinha51)
                                
                            End If
                        
                        End If
                        
                        lErro = Comando_BuscarProximo(alComando(8))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69814
            
                    Loop
        
                    For Each objLivRegESLinha50 In colLivRegESLinha50
                    
                        'Monta o Buffer do Tipo 50 de Acordo com as Linhas do Livro Fiscal
                        lErro = MontarReg50(objInfoArqICMS, tLivRegES, objLivRegESLinha50)
                        If lErro <> SUCESSO Then gError 69812
            
                        'Conta o Registro 50 para gerar o Tipo 90
                        objInfoArqICMS.lTotalReg50 = objInfoArqICMS.lTotalReg50 + 1
            
                        'e Grava no BD o Tipo 50
                        lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 50, alComando(9))
                        If lErro <> SUCESSO Then gError 69812
                        
                    Next
                               
                    If objInfoArqICMS.iGeraReg51 <> 0 Then
                    
                        For Each objLivRegESLinha51 In colLivRegESLinha51
                            
                            'Monta o Buffer do Tipo 51 de Acordo com as Linhas do Livro Fiscal
                            lErro = MontarReg51(objInfoArqICMS, tLivRegES, objLivRegESLinha51)
                            If lErro <> SUCESSO Then gError 69813
                
                            'Conta o Registro 51 para gerar o Tipo 90
                            objInfoArqICMS.lTotalReg51 = objInfoArqICMS.lTotalReg51 + 1
                
                            'e Grava no BD o Tipo 50
                            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 51, alComando(10))
                            If lErro <> SUCESSO Then gError 69813
                            
                        Next
            
                    End If
                    
                Case 8, 57 'conecimento de frete normal ou eletronico
                
                    tLivRegESLinha.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
                    tLivRegESLinha.sObservacaoLivFisc = String(STRING_OBSERVACAO, 0)
                    tLivRegESLinha.sClassifContabil = String(STRING_CONTA, 0)
            
                    'Lê as Linhas dos Livro para gerar o Tipo 70
                    lErro = Comando_Executar(alComando(14), "SELECT LivRegESLinha.NumIntDoc, LivRegESLinha.NumIntDocRegES, LivRegESLinha.NaturezaOp, LivRegESLinha.ValorTotal, LivRegESLinha.ValorICMSBase, LivRegESLinha.ValorICMS, LivRegESLinha.ValorICMSIsentoNaoTrib, LivRegESLinha.ValorICMSOutras, LivRegESLinha.ValorICMSSubstBase, LivRegESLinha.ValorICMSSubstRet, LivRegESLinha.AliquotaICMS, LivRegESLinha.ValorDespAcess, LivRegESLinha.ValorIPI, LivRegESLinha.AliquotaIPI, LivRegESLinha.ValorIPIBase, LivRegESLinha.ValorIPIIsentoNaoTrib, LivRegESLinha.ValorIPIOutras, LivRegESLinha.ClassifContabil, LivRegESLinha.ValorContabil, LivRegESLinha.ObservacaoLivFisc " & _
                        "FROM LivRegESLinha, LivRegEs, NFiscal WHERE LivRegESLinha.NumIntDocRegES = LivRegEs.NumIntDoc AND LivRegEs.NumIntDocOrigem = NFiscal.NumIntDoc AND LivRegESLinha.NumIntDocRegES = ? ORDER BY LivRegESLinha.NumIntDoc ASC", tLivRegESLinha.lNumIntDoc, tLivRegESLinha.lNumIntDocRegES, tLivRegESLinha.sNaturezaOp, tLivRegESLinha.dValorTotal, tLivRegESLinha.dValorICMSBase, tLivRegESLinha.dValorICMS, tLivRegESLinha.dValorICMSIsentoNaoTrib, tLivRegESLinha.dValorICMSOutras, tLivRegESLinha.dValorICMSSubstBase, tLivRegESLinha.dValorICMSSubstRet, tLivRegESLinha.dAliquotaICMS, tLivRegESLinha.dValorDespAcess, tLivRegESLinha.dValorIPI, tLivRegESLinha.dAliquotaIPI, tLivRegESLinha.dValorIPIBase, tLivRegESLinha.dValorIPIIsentoNaoTrib, _
                        tLivRegESLinha.dValorIPIOutras, tLivRegESLinha.sClassifContabil, tLivRegESLinha.dValorContabil, tLivRegESLinha.sObservacaoLivFisc, tLivRegES.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 69810
            
                    lErro = Comando_BuscarPrimeiro(alComando(14))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69811
                        
                    'Monta o Buffer do Tipo 70
                    Call MontarReg70(objInfoArqICMS, tLivRegES, tLivRegESLinha)
        
                    'Conta o Registro 70 para gerar o Tipo 90
                    objInfoArqICMS.lTotalReg70 = objInfoArqICMS.lTotalReg70 + 1
        
                    'e Grava no BD o Tipo 70
                    lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 70, alComando(9))
                    If lErro <> SUCESSO Then gError 81646
        
                Case 9000 To 9999
            
                    tLivRegES.iModelo = 55
            
                    If tLivRegES.iSituacao = 0 Then
            
                        sProduto = String(STRING_PRODUTO, 0)
                        sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
                        sCST = String(STRING_MAXIMO, 0)
                        'Lê os itens da Nota diretamente na tabela
                        lErro = Comando_Executar(alComando(15), "SELECT TDI.Item, TDI.Produto, TDI.Quantidade, TDI.PrecoTotal, TDI.NaturezaOp, TDI.CST, NF.ValorTotal FROM TributacaoDocItem TDI, NFiscal NF WHERE TDI.TipoDoc = 0 AND TDI.NumIntDoc = NF.NumIntDoc AND TDI.NumIntDoc = ? ORDER BY TDI.Item", _
                                iItem, sProduto, dQuantidade, dPrecoTotal, sNaturezaOp, sCST, dValorTotal, tLivRegES.lNumIntDocOrigem)
                        If lErro <> AD_SQL_SUCESSO Then gError 69801
                
                        lErro = Comando_BuscarPrimeiro(alComando(15))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69802
                    
                        'Loop por Item de Nota Fiscal
                        Do While lErro = SUCESSO
                        
                            If objInfoArqICMS.iGeraReg54 = 1 Then
                        
                                'Inicializa os Buffers
                                tLivRegESCadProd.sSiglaUM = String(STRING_ARQICMS_UM, 0)
                                tLivRegESCadProd.sDescricao = String(STRING_ARQICMS_DESCRICAO_PRODUTO, 0)
                                tLivRegESCadProd.sProduto = String(STRING_PRODUTO, 0)
                                tLivRegESCadProd.sCodigoNCM = String(STRING_CODIGONCM, 0)
                                tLivRegESCadProd.sSituacaoTrib = String(STRING_SITUACAOTRIB, 0)
                    
                                'Lê o cadastro Produto nos Livros Fiscais
                                lErro = Comando_Executar(alComando(3), "SELECT L2.NumIntDoc, L2.Produto, L2.DataInicial, L2.DataFinal, L2.CodigoNCM, LEFT(L2.Descricao, 53) Descricao, L2.SiglaUM, L2.SituacaoTrib, L2.AliquotaIPI, L2.AliquotaICMS, L2.ReducaoBaseCalculoICMS, L2.BaseCalculoICMSSubst FROM LivRegESCadProd L2 WHERE L2.Produto = ? AND L2.DataInicial <= ? ORDER BY L2.DataInicial DESC", _
                                        tLivRegESCadProd.lNumIntDoc, tLivRegESCadProd.sProduto, tLivRegESCadProd.dtDataInicial, tLivRegESCadProd.dtDataFinal, tLivRegESCadProd.sCodigoNCM, tLivRegESCadProd.sDescricao, tLivRegESCadProd.sSiglaUM, tLivRegESCadProd.sSituacaoTrib, tLivRegESCadProd.dAliquotaIPI, tLivRegESCadProd.dAliquotaICMS, tLivRegESCadProd.dReducaoBaseCalculoICMS, tLivRegESCadProd.dBaseCalculoICMSSubst, sProduto, objInfoArqICMS.dtDataFinal)
                                If lErro <> AD_SQL_SUCESSO Then gError 69803
                    
                                lErro = Comando_BuscarPrimeiro(alComando(3))
                                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69804
                    
                                'Se não encontrar --> Erro
                                If lErro = AD_SQL_SEM_DADOS Then gError 69805
                    
                                Do While lErro <> AD_SQL_SEM_DADOS
                                
                                    'inicializa a variavel que vai dizer se o o registro serah gravado ou nao
                                    bEntra = True
                                
                                    'inicio da barreira
                                    'se a data nao for data_nula
                                    If tLivRegESCadProd.dtDataFinal <> DATA_NULA Then
                                        
                                        'se data final do registro for menor a data inicial do periodo
                                        If tLivRegESCadProd.dtDataFinal < objInfoArqICMS.dtDataInicial Then
                                        
                                            'se o registro nao possui data final como nula e
                                            'se a data final dele eh menor que a data do periodo
                                            'significa que ele nao deve ser gravado... logo
                                            'ele nao deve entrar...
                                            bEntra = False
                                    
                                        End If
                                    
                                    End If
                                    
                                    'se for pra gravar...
                                    If bEntra = True Then
                                
                                        'e Grava o prototipo do Tipo 75 no BD para evitar repeticoes
                                        objGeracaoArqICMSProd.lNumIntLivRegESCadProd = tLivRegESCadProd.lNumIntDoc
                            
                                        'Só grava o tipo75 se ele ainda não estiver no BD
                                        lErro = CF("GeracaoArqICMSProd_Grava", objInfoArqICMS, objGeracaoArqICMSProd, alComando(6), alComando(7))
                                        If lErro <> SUCESSO Then gError 69808
                            
                                    End If
                                    
                                    If tLivRegESCadProd.dtDataInicial < objInfoArqICMS.dtDataInicial Then Exit Do
                                        
                                    lErro = Comando_BuscarProximo(alComando(3))
                                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1
                                    
                                Loop
                                
                                objInfoArqICMS.sBuffer = "54" & FormataCpoNum(tLivRegES.sCgc, 14) & FormataCpoNum(tLivRegES.iModelo, 2) & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(sNaturezaOp, 4) & FormataCpoNum(sCST, 3) & FormataCpoNum(iItem, 3) & FormataCpoTexto(sProduto, 14) & FormataCpoQuantidade(dQuantidade, 11) & FormataCpoValor(dPrecoTotal, 12) & FormataCpoValor(0, 12) & FormataCpoValor(0, 12) & FormataCpoValor(0, 12) & FormataCpoValor(0, 12) & FormataCpoValor(0, 4)
                                objInfoArqICMS.sChaveOrd = FormataCpoNum(tLivRegES.sCgc, 14) & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(iItem, 3)
                                
                                'Conta o registro 54
                                objInfoArqICMS.lTotalReg54 = objInfoArqICMS.lTotalReg54 + 1
                    
                                'e Grava no BD o Tipo 54
                                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 54, alComando(5))
                                If lErro <> SUCESSO Then gError 69807
                    
                            End If
                            
                            'Busca o Proximo Item no BD
                            lErro = Comando_BuscarProximo(alComando(15))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69809
                
                        Loop
                                            
                    Else
                    
                        sNaturezaOp = ""
                        dValorTotal = 0
                        
                    End If
                    
                    tLivRegESLinha.sObservacaoLivFisc = String(STRING_OBSERVACAO, 0)
                
                    'Lê as Linhas dos Livro para gerar o Tipo 50 e 51
                    lErro = Comando_Executar(alComando(8), "SELECT ObservacaoLivFisc FROM LivRegESLinha WHERE NumIntDocRegES = ? ORDER BY NumIntDoc", tLivRegESLinha.sObservacaoLivFisc, tLivRegES.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 69810
                
                    lErro = Comando_BuscarPrimeiro(alComando(8))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69811
                
                    If lErro = SUCESSO Then
                
                         Set objLivRegESLinha50 = New ClassLivRegESLinha
                     
                         objLivRegESLinha50.sNaturezaOp = sNaturezaOp
                         objLivRegESLinha50.dValorTotal = dValorTotal
                         
                         'Monta o Buffer do Tipo 50 de Acordo com as Linhas do Livro Fiscal
                         lErro = MontarReg50(objInfoArqICMS, tLivRegES, objLivRegESLinha50)
                         If lErro <> SUCESSO Then gError 69812
                
                         'Conta o Registro 50 para gerar o Tipo 90
                         objInfoArqICMS.lTotalReg50 = objInfoArqICMS.lTotalReg50 + 1
                
                         'e Grava no BD o Tipo 50
                         lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 50, alComando(9))
                         If lErro <> SUCESSO Then gError 69812
                         
                         If objInfoArqICMS.iGeraReg51 <> 0 Then
                         
                             Set objLivRegESLinha51 = New ClassLivRegESLinha
                         
                             objLivRegESLinha51.sNaturezaOp = sNaturezaOp
                             objLivRegESLinha51.dValorTotal = dValorTotal
                                                             
                             'Monta o Buffer do Tipo 51 de Acordo com as Linhas do Livro Fiscal
                             lErro = MontarReg51(objInfoArqICMS, tLivRegES, objLivRegESLinha51)
                             If lErro <> SUCESSO Then gError 69813
                 
                             'Conta o Registro 51 para gerar o Tipo 90
                             objInfoArqICMS.lTotalReg51 = objInfoArqICMS.lTotalReg51 + 1
                 
                             'e Grava no BD o Tipo 50
                             lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 51, alComando(10))
                             If lErro <> SUCESSO Then gError 69813
                         
                         End If
                
                    End If
                
                Case Else
                
            End Select
        
        End If
        
        'Pega o Reg de Entrada ou Saida
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69816

    Loop
    
    'Tratar NFe Inutilizadas
    lErro = Trata_NFe_Inutilizadas(objInfoArqICMS, alComando(17), alComando(9))
    If lErro <> SUCESSO Then gError 69812

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Processar_RegistrosES = SUCESSO

    Exit Function

Erro_Processar_RegistrosES:

    Processar_RegistrosES = gErr

    Select Case gErr
    
        Case 69798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
               
        Case 69799, 69800, 69816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
                
        Case 69801, 69802, 69809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESITEM", gErr)
                
        Case 69803, 69804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGCADPROD", gErr)
        
        Case 69805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVREGCADPROD_NAO_ENCONTRADO", gErr)
        
        Case 69807, 69808, 81646, 184010
                
        Case 69810, 69811, 69814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGESLINHA", gErr)
        
        Case 69812, 69813
        
        Case 69815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LIVREGES", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160735)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function MontarReg54(objInfoArqICMS As ClassInfoArqICMS, tLivRegES As typeLivRegES, tLivRegESItemNF As typeLivRegESItemNF, tLivRegESCadProd As typeLivRegESCadProd, ByVal lComando As Long) As Long
'preenche sBuffer com registro do tipo 54

Dim sSituacaoTrib As String
Dim lErro As Long, objTipoTribICMS As New ClassTipoTribICMS

On Error GoTo Erro_MontarReg54

'??? Jones: o codigo do produto deve estar mascarado

    If Trim(tLivRegES.sSerie) = "0" Then tLivRegES.sSerie = "" '??? dava erro no validador 2005
    
    If tLivRegES.lNumNotaFiscal > 999999 Then Exit Function 'Teste

    'alteracao convenio 69/02
    If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
    
        objInfoArqICMS.sBuffer = "54" & FormataCpoNum(tLivRegES.sCgc, 14) & FormataCpoNum(tLivRegES.iModelo, 2) & FormataCpoTexto(tLivRegES.sSerie, 3) & "  " & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(tLivRegESItemNF.sCFOP, 3) & FormataCpoNum(tLivRegESItemNF.iNumItem, 3) & FormataCpoTexto(LTrim(tLivRegESCadProd.sProduto), 14) & FormataCpoQuantidade(tLivRegESItemNF.dQuantidade, 13) & FormataCpoValor(tLivRegESItemNF.dValorProduto, 12) & FormataCpoValor(tLivRegESItemNF.dValorDescontoDespAcess, 12) & FormataCpoValor(tLivRegESItemNF.dValorICMSBase, 12) & FormataCpoValor(tLivRegESItemNF.dValorICMSSubstBase, 12) & FormataCpoValor(tLivRegESItemNF.dValorIPI, 12) & FormataCpoValor(tLivRegESItemNF.dAliquotaICMS * 100, 4)
        
    Else
    
        If tLivRegESItemNF.iNumItem <= 990 Then
        
            'Le a situação tributaria do Produto
            lErro = CF("Produto_Le_SituacaoTrib", sSituacaoTrib, tLivRegESCadProd.sProduto, tLivRegESItemNF.iTipoTribICMS, lComando)
            If lErro <> SUCESSO Then Error 58572
            
            objInfoArqICMS.sBuffer = "54" & FormataCpoNum(tLivRegES.sCgc, 14) & FormataCpoNum(tLivRegES.iModelo, 2) & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(tLivRegESItemNF.sCFOP, 4) & FormataCpoTexto(sSituacaoTrib, 3) & FormataCpoNum(tLivRegESItemNF.iNumItem, 3) & FormataCpoTexto(LTrim(tLivRegESCadProd.sProduto), 14) & FormataCpoQuantidade(tLivRegESItemNF.dQuantidade, 11) & FormataCpoValor(tLivRegESItemNF.dValorProduto, 12) & FormataCpoValor(tLivRegESItemNF.dValorDescontoDespAcess, 12) & FormataCpoValor(tLivRegESItemNF.dValorICMSBase, 12) & FormataCpoValor(tLivRegESItemNF.dValorICMSSubstBase, 12) & FormataCpoValor(tLivRegESItemNF.dValorIPI, 12) & FormataCpoValor(tLivRegESItemNF.dAliquotaICMS * 100, 4)
    
        Else
        
            sSituacaoTrib = "   "
        
            objInfoArqICMS.sBuffer = "54" & FormataCpoNum(tLivRegES.sCgc, 14) & FormataCpoNum(tLivRegES.iModelo, 2) & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(tLivRegESItemNF.sCFOP, 4) & FormataCpoTexto(sSituacaoTrib, 3) & FormataCpoNum(tLivRegESItemNF.iNumItem, 3) & FormataCpoTexto(LTrim(tLivRegESCadProd.sProduto), 14) & FormataCpoQuantidade(tLivRegESItemNF.dQuantidade, 11) & FormataCpoValor(tLivRegESItemNF.dValorProduto, 12) & FormataCpoValor(tLivRegESItemNF.dValorDescontoDespAcess, 12) & FormataCpoValor(tLivRegESItemNF.dValorICMSBase, 12) & FormataCpoValor(tLivRegESItemNF.dValorICMSSubstBase, 12) & FormataCpoValor(tLivRegESItemNF.dValorIPI, 12) & FormataCpoValor(tLivRegESItemNF.dAliquotaICMS * 100, 4)
    
        End If
        
    End If

    objInfoArqICMS.sChaveOrd = FormataCpoNum(tLivRegES.sCgc, 14) & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(tLivRegESItemNF.iNumItem, 3)
    
    MontarReg54 = SUCESSO

    Exit Function
    
Erro_MontarReg54:

    MontarReg54 = gErr
     
    Select Case gErr
          
        Case 58572, 124160
        
        Case 124161
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_ICMS_INEXISTENTE", gErr, objTipoTribICMS.iTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160736)
     
    End Select
     
    Exit Function

End Function

Private Function MontarReg75(sBuffer As String, objInfoArqICMS As ClassInfoArqICMS, tTipo75 As typeTipo75) As Long
'preenche sBuffer com registro do tipo 75

'??? Jones: o codigo do produto deve estar mascarado
    
    sBuffer = "75" & FormataCpoData(IIf(tTipo75.dtDataInicial < objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataInicial, tTipo75.dtDataInicial)) & FormataCpoData(IIf(tTipo75.dtDataFinal <> DATA_NULA And tTipo75.dtDataFinal <= objInfoArqICMS.dtDataFinal, tTipo75.dtDataFinal, objInfoArqICMS.dtDataFinal)) & FormataCpoTexto(LTrim(tTipo75.sCodigoProduto), 14) & FormataCpoTexto(tTipo75.sCodigoNCM, 8) & FormataCpoTexto(tTipo75.sDescricaoProduto, 53) & FormataCpoTexto(tTipo75.sUnidadeMedida, 6) & tTipo75.sSituacaoTributaria & "0" & FormataCpoNum(tTipo75.dAliquotaIPI, 4) & FormataCpoNum(tTipo75.dAliquotaICMS, 4) & FormataCpoNum(tTipo75.dReducaoBaseICMS, 4) & FormataCpoValor(tTipo75.dBaseICMSSubst, 12)
    
    MontarReg75 = SUCESSO
    
End Function

'Private Function MontarReg75(objInfoArqICMS As ClassInfoArqICMS, tLivRegES As typeLivRegES, tLivRegESItemNF As typeLivRegESItemNF, tLivRegESCadProd As typeLivRegESCadProd) As Long
''preenche sBuffer com registro do tipo 75
'
'    objInfoArqICMS.sBuffer = "75" & FormataCpoData(objInfoArqICMS.dtDataDe) & FormataCpoData(objInfoArqICMS.dtDataAte) & FormataCpoTexto(tLivRegESCadProd.sProduto, 14) & FormataCpoTexto(tLivRegESCadProd.sCodigoNCM, 8) & FormataCpoTexto(tLivRegESCadProd.sDescricao, 53) & FormataCpoTexto(tLivRegESCadProd.sSiglaUM, 6) & tLivRegESCadProd.sSituacaoTrib & "0" & FormataCpoNum(tLivRegESCadProd.dAliquotaIPI * 100, 4) & FormataCpoNum(tLivRegESCadProd.dAliquotaICMS, 4) & FormataCpoNum(tLivRegESCadProd.dReducaoBaseCalculoICMS, 4) & FormataCpoValor(tLivRegESCadProd.dBaseCalculoICMSSubst, 12)
'
'    MontarReg75 = SUCESSO
'
'End Function

Private Function LivRegES_Situacao_Sintegra(objInfoArqICMS As ClassInfoArqICMS, tLivRegES As typeLivRegES, sSituacao As String) As Long

Dim lComando As Long, iSituacao As Integer
Dim lErro As Long, lNumIntNF As Long, itpAmb As Integer

On Error GoTo Erro_LivRegES_Situacao_Sintegra

    iSituacao = tLivRegES.iSituacao

    If iSituacao = 1 And tLivRegES.iTipoNumIntDocOrigem = 0 Then
    
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 201544
    
        lErro = Comando_Executar(lComando, "SELECT tpAmb FROM NFeFedProtNFe WHERE NumIntNF = ? AND cStat IN ('110', '301','302') AND tpAmb = 0", itpAmb, tLivRegES.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 201545
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201546
    
        'nota fiscal denegada
        If lErro = AD_SQL_SUCESSO Then iSituacao = 4
    
        Call Comando_Fechar(lComando)
    
    End If
    
    Select Case iSituacao
    
        Case 0 'Documento Fiscal Normal
            sSituacao = "N"
        
        Case 1 'Documento Fiscal Cancelado
            sSituacao = "S"
            
        Case 2 'Lançamento Extemporâneo de Documento Fiscal Normal
            sSituacao = "E"
    
        Case 3 'Lançamento Extemporâneo de Documento Fiscal Cancelado
            sSituacao = "X"
        
        Case 4 'Documento com USO DENEGADO - exclusivamente para uso dos emitentes de Nota Fiscal Eletrônica - Modelo 55
            sSituacao = "2"
        
        Case 5 'Documento com USO inutilizado - exclusivamente para uso dos emitentes de Nota Fiscal Eletrônica - Modelo 55
            sSituacao = "4"
        
        Case Else 'Documento Fiscal Normal
            gError 201543
            
    End Select
    
    LivRegES_Situacao_Sintegra = SUCESSO

    Exit Function
    
Erro_LivRegES_Situacao_Sintegra:

    LivRegES_Situacao_Sintegra = gErr

    Select Case gErr

        Case 201544
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201543
            Call Rotina_Erro(vbOKOnly, "ERRO_LIVREGES_SITUACAO_INVALIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 20541)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function MontarReg50(objInfoArqICMS As ClassInfoArqICMS, tLivRegES As typeLivRegES, objLivRegESLinha As ClassLivRegESLinha) As Long
'preenche sBuffer com registro do tipo 50

Dim sEmitente As String, sSituacao As String
Dim lErro As Long

On Error GoTo Erro_MontarReg50

    If Trim(tLivRegES.sSerie) = "0" Then tLivRegES.sSerie = "" '??? dava erro no validador 2005
    
    If tLivRegES.lNumNotaFiscal > 999999 Then Exit Function 'Teste

    objInfoArqICMS.sBuffer = "50" & FormataCpoNum(tLivRegES.sCgc, 14) & InscEst_FormatadaArqICMS(tLivRegES.sInscricaoEstadual, tLivRegES.iIEIsento) & FormataCpoData(tLivRegES.dtData) & FormataCpoTexto(tLivRegES.sUF, 2) & FormataCpoNum(tLivRegES.iModelo, 2) & FormataCpoTexto(tLivRegES.sSerie, 3)
    
    'alteracao convenio 69/02
    If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
    
        objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoTexto("", 2) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(objLivRegESLinha.sNaturezaOp, 3)
        
    Else
    
        sEmitente = IIf(tLivRegES.iEmitente = EMITENTE_EMPRESA, "P", "T")
        
        objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(objLivRegESLinha.sNaturezaOp, 4) & FormataCpoTexto(sEmitente, 1)
    
    End If

    lErro = LivRegES_Situacao_Sintegra(objInfoArqICMS, tLivRegES, sSituacao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoValor(objLivRegESLinha.dValorTotal, 13) & FormataCpoValor(objLivRegESLinha.dValorICMSBase, 13) & FormataCpoValor(objLivRegESLinha.dValorICMS, 13) & FormataCpoValor(objLivRegESLinha.dValorICMSIsentoNaoTrib, 13) & FormataCpoValor(objLivRegESLinha.dValorICMSOutras, 13) & FormataCpoValor(objLivRegESLinha.dAliquotaICMS * 100, 4) & FormataCpoTexto(sSituacao, 1)
    
    objInfoArqICMS.sChaveOrd = FormataCpoData(tLivRegES.dtData) & FormataCpoNum(tLivRegES.iModelo, 2) & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6)

    MontarReg50 = SUCESSO

    Exit Function
    
Erro_MontarReg50:

    MontarReg50 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201542)

    End Select
    
    Exit Function

End Function

Private Function MontarReg70(objInfoArqICMS As ClassInfoArqICMS, tLivRegES As typeLivRegES, tLivRegESLinha As typeLivRegESLinha) As Long
'preenche sBuffer com registro do tipo 70
Dim sSerieSubSerie As String

    Select Case Len(tLivRegES.sSerie)
    
        Case 0
            sSerieSubSerie = "   "
            
        Case 1
            If tLivRegES.iModelo = 8 Then
            
                If IsNumeric(tLivRegES.sSerie) Then
                    sSerieSubSerie = "U" & tLivRegES.sSerie & " "
                Else
                    sSerieSubSerie = "   "
                End If
            
            Else
            
                sSerieSubSerie = tLivRegES.sSerie & "  "
            
            End If
            
        Case 2
            sSerieSubSerie = tLivRegES.sSerie & " "
            
        Case Else
            sSerieSubSerie = left(tLivRegES.sSerie, 3)
        
    End Select
    
    objInfoArqICMS.sBuffer = "70" & FormataCpoNum(tLivRegES.sCgc, 14) & InscEst_FormatadaArqICMS(tLivRegES.sInscricaoEstadual, tLivRegES.iIEIsento) & FormataCpoData(tLivRegES.dtData) & FormataCpoTexto(tLivRegES.sUF, 2) & FormataCpoNum(tLivRegES.iModelo, 2) & sSerieSubSerie & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6)
    
    'alteracao convenio 69/02
    If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
         objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(tLivRegESLinha.sNaturezaOp, 3) & FormataCpoValor(tLivRegESLinha.dValorTotal, 14)
    Else
         objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(tLivRegESLinha.sNaturezaOp, 4) & FormataCpoValor(tLivRegESLinha.dValorTotal, 13)
    End If

    objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoValor(tLivRegESLinha.dValorICMSBase, 14) & FormataCpoValor(tLivRegESLinha.dValorICMS, 14) & FormataCpoValor(tLivRegESLinha.dValorICMSIsentoNaoTrib, 14) & FormataCpoValor(tLivRegESLinha.dValorICMSOutras, 14) & FormataCpoNum(tLivRegES.iCIF_FOB, 1) & FormataCpoTexto(IIf(tLivRegES.iSituacao = 0, "N", "S"), 1)

    objInfoArqICMS.sChaveOrd = FormataCpoData(tLivRegES.dtData)

    MontarReg70 = SUCESSO

End Function

Private Function MontarReg51(objInfoArqICMS As ClassInfoArqICMS, tLivRegES As typeLivRegES, objLivRegESLinha As ClassLivRegESLinha) As Long
'preenche sBuffer com registro do tipo 51

Dim sSituacao As String
Dim lErro As Long

On Error GoTo Erro_MontarReg51

    If Trim(tLivRegES.sSerie) = "0" Then tLivRegES.sSerie = "" '??? dava erro no validador 2005
    
    If tLivRegES.lNumNotaFiscal > 999999 Then Exit Function 'Teste

    objInfoArqICMS.sBuffer = "51" & FormataCpoNum(tLivRegES.sCgc, 14) & InscEst_FormatadaArqICMS(tLivRegES.sInscricaoEstadual, tLivRegES.iIEIsento) & FormataCpoData(tLivRegES.dtData) & FormataCpoTexto(tLivRegES.sUF, 2)
    
    'alteracao convenio 69/02
    If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
        objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoTexto(tLivRegES.sSerie, 2) & FormataCpoTexto("", 2) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(objLivRegESLinha.sNaturezaOp, 3)
    Else
        objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoTexto(tLivRegES.sSerie, 3) & FormataCpoNum(tLivRegES.lNumNotaFiscal, 6) & FormataCpoNum(objLivRegESLinha.sNaturezaOp, 4)
    End If

    lErro = LivRegES_Situacao_Sintegra(objInfoArqICMS, tLivRegES, sSituacao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoValor(objLivRegESLinha.dValorTotal, 13) & FormataCpoValor(objLivRegESLinha.dValorIPI, 13) & FormataCpoValor(objLivRegESLinha.dValorIPIIsentoNaoTrib, 13) & FormataCpoValor(objLivRegESLinha.dValorIPIOutras, 13) & FormataCpoTexto("", 20) & FormataCpoTexto(sSituacao, 1)

    objInfoArqICMS.sChaveOrd = FormataCpoData(tLivRegES.dtData)

    MontarReg51 = SUCESSO

    Exit Function
    
Erro_MontarReg51:

    MontarReg51 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201542)

    End Select
    
    Exit Function

End Function

Private Function MontarReg10(sBuffer As String, objFilial As AdmFiliais, objInfoArqICMS As ClassInfoArqICMS) As Long
'preenche sBuffer com registro do tipo 10

Dim lErro As Long

On Error GoTo Erro_MontarReg10

    lErro = Tamanho_Valida("FilialCGC", objFilial.sCgc, 14)
    If lErro <> SUCESSO Then gError 200239
    
    lErro = Tamanho_Valida("FilialIE", InscEst_FormatadaArqICMS(objFilial.sInscricaoEstadual), 14)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialNome", objFilial.sNome, 35)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialCidade", objFilial.objEndereco.sCidade, 30)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialUF", objFilial.objEndereco.sSiglaEstado, 2)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialFax", Replace(objFilial.objEndereco.sFax, " ", ""), 10)
    If lErro <> SUCESSO Then gError 200239

    sBuffer = "10" & FormataCpoNum(objFilial.sCgc, 14) & InscEst_FormatadaArqICMS(objInfoArqICMS.sInscricaoEstadual) & FormataCpoTexto(objInfoArqICMS.sNome, 35) & FormataCpoTexto(objFilial.objEndereco.sCidade, 30) & FormataCpoTexto(objFilial.objEndereco.sSiglaEstado, 2) & FormataCpoNum(Replace(objFilial.objEndereco.sFax, " ", ""), 10) & FormataCpoData(objInfoArqICMS.dtDataInicial) & FormataCpoData(objInfoArqICMS.dtDataFinal)
    'alteracao convenio 69/02
    If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
        sBuffer = sBuffer & "1"
    ElseIf objInfoArqICMS.dtDataFinal < #1/1/2004# Then
        sBuffer = sBuffer & "2"
    Else
        sBuffer = sBuffer & "3"
    End If
    sBuffer = sBuffer & IIf(objInfoArqICMS.bIntegral, "3", "2") & CStr(Codigo_Extrai(objInfoArqICMS.sFinalidade))

    MontarReg10 = SUCESSO

    Exit Function

Erro_MontarReg10:

    Select Case gErr
    
        Case 200239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160737)

    End Select

    Exit Function

End Function

Private Function MontarReg11(sBuffer As String, objFilial As AdmFiliais, objInfoArqICMS As ClassInfoArqICMS) As Long
'preenche sBuffer com registro do tipo 11

Dim lErro As Long

On Error GoTo Erro_MontarReg11

    lErro = Tamanho_Valida("FilialEndLogradouro", objInfoArqICMS.sLogradouro, 34)
    If lErro <> SUCESSO Then gError 200239
    
    lErro = Tamanho_Valida("FilialEndNumero", objInfoArqICMS.lNumero, 5)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialEndComplemento", objInfoArqICMS.sComplemento, 22)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialEndBairro", objFilial.objEndereco.sBairro, 15)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialEndCEP", objFilial.objEndereco.sCEP, 8)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialEndContato", objInfoArqICMS.sContato, 28)
    If lErro <> SUCESSO Then gError 200239

    lErro = Tamanho_Valida("FilialEndTelContato", objInfoArqICMS.sTelContato, 12)
    If lErro <> SUCESSO Then gError 200239

    sBuffer = "11" & FormataCpoTexto(objInfoArqICMS.sLogradouro, 34) & FormataCpoNum(objInfoArqICMS.lNumero, 5) & FormataCpoTexto(objInfoArqICMS.sComplemento, 22) & FormataCpoTexto(objFilial.objEndereco.sBairro, 15) & FormataCpoNum(objFilial.objEndereco.sCEP, 8) & FormataCpoTexto(objInfoArqICMS.sContato, 28) & FormataCpoNum(objInfoArqICMS.sTelContato, 12)

    MontarReg11 = SUCESSO

    Exit Function

Erro_MontarReg11:

    Select Case gErr
    
        Case 200239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160738)

    End Select

    Exit Function

End Function

Private Function MontarReg90(sBuffer As String, objInfoArqICMS As ClassInfoArqICMS) As Long
'preenche sBuffer com registro do tipo 90

Dim mvarlNumRegTotal As Long

    sBuffer = "90" & FormataCpoNum(objInfoArqICMS.sCgc, 14) & InscEst_FormatadaArqICMS(objInfoArqICMS.sInscricaoEstadual)

    If objInfoArqICMS.lTotalReg50 > 0 Then
        sBuffer = sBuffer & "50" & FormataCpoNum(objInfoArqICMS.lTotalReg50, 8)
    End If

    If objInfoArqICMS.lTotalReg51 > 0 Then
        sBuffer = sBuffer & "51" & FormataCpoNum(objInfoArqICMS.lTotalReg51, 8)
    End If

    If objInfoArqICMS.lTotalReg53 > 0 Then
        sBuffer = sBuffer & "53" & FormataCpoNum(objInfoArqICMS.lTotalReg53, 8)
    End If
    
    If objInfoArqICMS.lTotalReg54 > 0 Then
        sBuffer = sBuffer & "54" & FormataCpoNum(objInfoArqICMS.lTotalReg54, 8)
    End If
    
    If objInfoArqICMS.lTotalReg60 > 0 Then
        sBuffer = sBuffer & "60" & FormataCpoNum(objInfoArqICMS.lTotalReg60, 8)
    End If

    If objInfoArqICMS.lTotalReg61 > 0 Then
        sBuffer = sBuffer & "61" & FormataCpoNum(objInfoArqICMS.lTotalReg61, 8)
    End If

    If objInfoArqICMS.lTotalReg70 > 0 Then
        sBuffer = sBuffer & "70" & FormataCpoNum(objInfoArqICMS.lTotalReg70, 8)
    End If

    If objInfoArqICMS.lTotalReg71 > 0 Then
        sBuffer = sBuffer & "71" & FormataCpoNum(objInfoArqICMS.lTotalReg71, 8)
    End If

    If objInfoArqICMS.lTotalReg74 > 0 Then
        sBuffer = sBuffer & "74" & FormataCpoNum(objInfoArqICMS.lTotalReg74, 8)
    End If

    If objInfoArqICMS.lTotalReg75 > 0 Then
        sBuffer = sBuffer & "75" & FormataCpoNum(objInfoArqICMS.lTotalReg75, 8)
    End If

    mvarlNumRegTotal = objInfoArqICMS.lTotalReg50 + objInfoArqICMS.lTotalReg51 + objInfoArqICMS.lTotalReg53 + objInfoArqICMS.lTotalReg54 + objInfoArqICMS.lTotalReg60 + objInfoArqICMS.lTotalReg61 + objInfoArqICMS.lTotalReg70 + objInfoArqICMS.lTotalReg71 + objInfoArqICMS.lTotalReg74 + objInfoArqICMS.lTotalReg75 + 3

    sBuffer = sBuffer & "99" & FormataCpoNum(mvarlNumRegTotal, 8)

    sBuffer = FormataCpoTexto(sBuffer, 125)
    sBuffer = sBuffer & "1"
    
    MontarReg90 = SUCESSO

End Function

'Não Subir
Private Function InserirRegsArqICMS1(objFilial As AdmFiliais, objInfoArqICMS As ClassInfoArqICMS) As Long
'gera o arquivo para a objInfoArqICMS.sUFDestino

Dim sBuffer As String
Dim iIndice As Integer
Dim lErro As Long, sNomeArq As String

On Error GoTo Erro_InserirRegsArqICMS1

    If objInfoArqICMS.bIntegral Then
        sNomeArq = objInfoArqICMS.sNomeArquivo
    Else
        sNomeArq = objInfoArqICMS.sNomeArquivo & objInfoArqICMS.sUFDestino
    End If
    
    Open sNomeArq For Output As #1

    lErro = MontarReg10(sBuffer, objFilial, objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69817

    'incluir reg tipo 10
    Print #1, sBuffer

    lErro = MontarReg11(sBuffer, objFilial, objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69818

    'incluir reg tipo 11
    Print #1, sBuffer

    'percorrer a tabela GeracaoArqICMS onde ficaram armazenados os regs montados acima e criar o arquivo
    lErro = GeracaoArqICMS_Inclui_registros(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69819

    'Monta e insere os Registro Tipo 75
    lErro = GeracaoArqICMSProd_Inclui_Tipos75(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69820

    'Inclui o Reg 90
    lErro = MontarReg90(sBuffer, objInfoArqICMS)
    If lErro <> SUCESSO Then gError 69821

    Print #1, sBuffer

    Close

    InserirRegsArqICMS1 = SUCESSO

    Exit Function

Erro_InserirRegsArqICMS1:

    InserirRegsArqICMS1 = gErr

    Select Case gErr

        Case 69817, 69818, 69819, 69820, 69821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160739)

    End Select

    Close

    Exit Function

End Function

'Não Subir
Function GeracaoArqICMS_Inclui_registros(objInfoArqICMS As ClassInfoArqICMS) As Long
'Le todos os Tipos armazenados no BD no, em Ordem

Dim lErro As Long
Dim sBuffer As String
Dim lComando As Long, iTipo As Integer

On Error GoTo Erro_GeracaoArqICMS_Inclui_registros

    With objInfoArqICMS
        .lTotalReg60 = 0
        .lTotalReg50 = 0
        .lTotalReg51 = 0
        .lTotalReg53 = 0
        .lTotalReg54 = 0
        .lTotalReg70 = 0
        .lTotalReg71 = 0
    End With
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69822

    sBuffer = String(127, 0)

    'Seleciona todos os Tipos
    lErro = Comando_Executar(lComando, "SELECT Tipo, Conteudo FROM GeracaoArqICMS WHERE UFDest = ? ORDER BY Tipo, ChaveOrd ASC", iTipo, sBuffer, objInfoArqICMS.sUFDestino)
    If lErro <> SUCESSO Then gError 69823

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69824

    Do While lErro = AD_SQL_SUCESSO

        Select Case iTipo
        
            Case 50
                objInfoArqICMS.lTotalReg50 = objInfoArqICMS.lTotalReg50 + 1
                
            Case 51
                objInfoArqICMS.lTotalReg51 = objInfoArqICMS.lTotalReg51 + 1
                
            Case 53
                objInfoArqICMS.lTotalReg53 = objInfoArqICMS.lTotalReg53 + 1
                
            Case 54
                objInfoArqICMS.lTotalReg54 = objInfoArqICMS.lTotalReg54 + 1
                
            Case 60
                objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
                
            Case 70
                objInfoArqICMS.lTotalReg70 = objInfoArqICMS.lTotalReg70 + 1
                
            Case 71
                objInfoArqICMS.lTotalReg71 = objInfoArqICMS.lTotalReg71 + 1
                
        End Select
        
        Print #1, sBuffer

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69825

    Loop

    Call Comando_Fechar(lComando)

    GeracaoArqICMS_Inclui_registros = SUCESSO

    Exit Function

Erro_GeracaoArqICMS_Inclui_registros:

    GeracaoArqICMS_Inclui_registros = gErr

    Select Case gErr

        Case 69822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69823, 69824, 69825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160740)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Não Subir
Function GeracaoArqICMSProd_Inclui_Tipos75(objInfoArqICMS As ClassInfoArqICMS) As Long
'Le todos os Tipos 75 armazenados no BD no, em Ordem e inclui no Arquivo

Dim lErro As Long
Dim sBuffer As String
Dim lComando As Long
Dim tTipo75 As typeTipo75
Dim iOrigemMercadoria As Integer
Dim iTipoTribCST As Integer

On Error GoTo Erro_GeracaoArqICMSProd_Inclui_Tipos75

    objInfoArqICMS.lTotalReg75 = 0
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69826

    tTipo75.sCodigoProduto = String(STRING_PRODUTO, 0)
    tTipo75.sDescricaoProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tTipo75.sUnidadeMedida = String(STRING_UM_SIGLA, 0)
    tTipo75.sCodigoNCM = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tTipo75.sSituacaoTributaria = String(STRING_SITUACAOTRIB, 0)
    
    'Seleciona todos os Tipos
    lErro = Comando_Executar(lComando, "SELECT LivRegESCadProd.Produto, LivRegESCadProd.DataInicial, LivRegESCadProd.DataFinal, LivRegESCadProd.Descricao, LivRegESCadProd.SiglaUM, LivRegESCadProd.SituacaoTrib, LivRegESCadProd.CodigoNCM, AliquotaIPI, AliquotaICMS, ReducaoBaseCalculoICMS, BaseCalculoICMSSubst FROM LivRegESCadProd, GeracaoArqICMSProd WHERE GeracaoArqICMSProd.NumIntLivRegESCadProd = LivRegESCadProd.NumIntDoc AND GeracaoArqICMSProd.UFDest = ? ORDER BY LivRegESCadProd.Produto", tTipo75.sCodigoProduto, tTipo75.dtDataInicial, tTipo75.dtDataFinal, tTipo75.sDescricaoProduto, tTipo75.sUnidadeMedida, tTipo75.sSituacaoTributaria, tTipo75.sCodigoNCM, tTipo75.dAliquotaIPI, tTipo75.dAliquotaICMS, tTipo75.dReducaoBaseICMS, tTipo75.dBaseICMSSubst, objInfoArqICMS.sUFDestino)
    If lErro <> SUCESSO Then gError 69827

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69828

    Do While lErro = AD_SQL_SUCESSO

        'Conta o Tipo 75 para a Geracao do Tipo 90
        objInfoArqICMS.lTotalReg75 = objInfoArqICMS.lTotalReg75 + 1

        Call MontarReg75(sBuffer, objInfoArqICMS, tTipo75)

        Print #1, sBuffer

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69829

    Loop

    Call Comando_Fechar(lComando)

    GeracaoArqICMSProd_Inclui_Tipos75 = SUCESSO

    Exit Function

Erro_GeracaoArqICMSProd_Inclui_Tipos75:

    GeracaoArqICMSProd_Inclui_Tipos75 = gErr

    Select Case gErr

        Case 69826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69827, 69828, 69829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMSPROD", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160741)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


'??? corrigir a versao em globaisadm2.vbp
Function FormataCpoQuantidade(vData As Variant, iTam As Integer) As String
'formata campo Quantidade alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim sValor As String
Dim iTamanhoValor As Integer
Dim dValor As Double
Dim sDecimais As String
Dim sInteiro As String
Dim sNumero As String
Dim sEspacos As String
Dim iData As Integer

    sValor = Format(vData, "##########0.000")
    sDecimais = Mid(sValor, Len(sValor) - 2, 3)
    sInteiro = Mid(sValor, 1, Len(sValor) - 4)
    sNumero = sInteiro & sDecimais

    iData = iTam - Len(sNumero)
    sEspacos = String(iData, "0")

    FormataCpoQuantidade = sEspacos & sInteiro & sDecimais

End Function

Function Processar_ConhecFrete_Transp(objFilial As AdmFiliais, objInfoArqICMS As ClassInfoArqICMS) As Long
'gerar registros 70 e 71 p/transportadoras

Dim iIndice As Integer, alComando(0 To 1) As Long, dOutrasTrib As Double, dBaseTrib As Double
Dim lErro As Long, dtDataEmissao As Date, sNaturezaOp, iCIF_FOB As Integer, lNumNFcli As Long
Dim sSerieSubSerie As String, sUF As String, dValorICMSIsentoNaoTrib As String, iIncluiPedagio As Integer
Dim lNumNotaFiscal As Long, iStatus As Integer, sSerie As String, sInscricaoEstadual As String, sCgc As String, sCGCRemetente As String, sInscEstadualRemetente As String, sUFRemetente As String, sCGCDestinatario As String, sInscEstadualDestinatario As String, sUFDestinatario As String, dValorTotal As Double, dBaseCalculo As Double, dValorICMS As Double, dPedagio As Double, sNotasFiscais As String, dValorMercadoria As Double

On Error GoTo Erro_Processar_ConhecFrete_Transp

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81763
    Next

    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    sUF = String(STRING_ESTADO_SIGLA, 0)
    sSerie = String(STRING_SERIE, 0)
    sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    sCgc = String(STRING_CGC, 0)
    sCGCRemetente = String(STRING_CGC, 0)
    sInscEstadualRemetente = String(STRING_INSCR_EST, 0)
    sUFRemetente = String(STRING_ESTADO_SIGLA, 0)
    sCGCDestinatario = String(STRING_CGC, 0)
    sInscEstadualDestinatario = String(STRING_INSCR_EST, 0)
    sUFDestinatario = String(STRING_ESTADO_SIGLA, 0)
    sNotasFiscais = String(STRING_CONHECIMENTOFRETE_NOTAS, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT NFiscal.NaturezaOP, Enderecos.SiglaEstado, DataEmissao, NumNotaFiscal, Status, Serie, InscricaoEstadual, CGC, CGCRemetente, InscEstadualRemetente, UFRemetente, CGCDestinatario, InscEstadualDestinatario, UFDestinatario, ValorTotal, BaseCalculo, ValorICMS, Pedagio, NotasFiscais, ConhecimentoFrete.ValorMercadoria, IncluiPedagio FROM Enderecos, FiliaisClientes, Clientes, NFiscal, ConhecimentoFrete WHERE FiliaisClientes.Endereco = Enderecos.Codigo and NFiscal.NumIntDoc = ConhecimentoFrete.NumIntNFiscal AND DataEmissao >= ? AND DataEmissao <= ? AND Clientes.Codigo = FiliaisClientes.CodCliente AND NFiscal.Cliente = FiliaisClientes.CodCliente AND NFiscal.FilialCli = FiliaisClientes.CodFilial", _
        sNaturezaOp, sUF, dtDataEmissao, lNumNotaFiscal, iStatus, sSerie, sInscricaoEstadual, sCgc, sCGCRemetente, sInscEstadualRemetente, sUFRemetente, sCGCDestinatario, sInscEstadualDestinatario, sUFDestinatario, dValorTotal, dBaseCalculo, dValorICMS, dPedagio, sNotasFiscais, dValorMercadoria, iIncluiPedagio, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 81764

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81765

    Do While lErro = AD_SQL_SUCESSO

        If objInfoArqICMS.bIntegral Then
            objInfoArqICMS.sUFDestino = objInfoArqICMS.sUF
        Else
            objInfoArqICMS.sUFDestino = sUFDestinatario
        End If
        
        Select Case Len(sSerie)
        
            Case 0
                sSerieSubSerie = "   "
                
            Case 1
                If IsNumeric(sSerie) Then
                    sSerieSubSerie = "U" & sSerie & " "
                Else
                    sSerieSubSerie = "   "
                End If
                
            Case 2
                sSerieSubSerie = sSerie & " "
                
            Case Else
                sSerieSubSerie = left(sSerie, 3)
            
        End Select
        
        If dValorICMS <> 0 Then
        
            dBaseTrib = dBaseCalculo
            
            'se o pedagio nao está dentro dos outros valores (e portanto da base de calculo)
            If iIncluiPedagio = 0 Then
                                
                dValorICMSIsentoNaoTrib = Round(dValorTotal - dBaseCalculo - dPedagio, 2)
                dOutrasTrib = dPedagio
            Else
                dValorICMSIsentoNaoTrib = Round(dValorTotal - dBaseCalculo, 2)
                dOutrasTrib = 0
            End If
            
        Else
        
            dBaseTrib = 0
            
            If iIncluiPedagio = 0 Then
                dValorICMSIsentoNaoTrib = Round(dValorTotal - dPedagio, 2)
                dOutrasTrib = dPedagio
            Else
                dValorICMSIsentoNaoTrib = Round(dValorTotal, 2)
                dOutrasTrib = 0
            End If
        
        End If
        
        If Len(Trim(sCgc)) <> 0 Then
        
            If Trim(sCGCDestinatario) <> Trim(sCGCRemetente) Then
            
                If Trim(sCgc) = Trim(sCGCDestinatario) Then
                    iCIF_FOB = FRETE_DESTINATARIO
                Else
                    iCIF_FOB = FRETE_EMITENTE
                End If
            
            Else
            
                If Trim(sInscricaoEstadual) = Trim(sInscEstadualDestinatario) Then
                    iCIF_FOB = FRETE_DESTINATARIO
                Else
                    iCIF_FOB = FRETE_EMITENTE
                End If
                
            End If
            
        Else
        
            iCIF_FOB = FRETE_EMITENTE
            
        End If
        
        objInfoArqICMS.sBuffer = "70" & FormataCpoNum(sCgc, 14) & InscEst_FormatadaArqICMS(sInscricaoEstadual) & FormataCpoData(dtDataEmissao) & FormataCpoTexto(sUF, 2) & FormataCpoNum(8, 2) & sSerieSubSerie & FormataCpoNum(lNumNotaFiscal, 6)
        
        'alteracao convenio 69/02
        If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(sNaturezaOp, 3) & FormataCpoValor(dValorTotal, 14)
        Else
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(sNaturezaOp, 4) & FormataCpoValor(dValorTotal, 13)
        End If
                    
        objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoValor(dBaseTrib, 14) & FormataCpoValor(dValorICMS, 14) & FormataCpoValor(dValorICMSIsentoNaoTrib, 14) & FormataCpoValor(dOutrasTrib, 14) & FormataCpoNum(iCIF_FOB, 1) & FormataCpoTexto(IIf(iStatus <> 7, "N", "S"), 1)
    
        objInfoArqICMS.sChaveOrd = FormataCpoData(dtDataEmissao) & FormataCpoNum(lNumNotaFiscal, 6)
        
        'Conta o Registro 70 para gerar o Tipo 90
        objInfoArqICMS.lTotalReg70 = objInfoArqICMS.lTotalReg70 + 1

        'e Grava no BD o Tipo 70
        lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 70, alComando(1))
        If lErro <> SUCESSO Then gError 81768
        
        'se o conhec de frete nao foi cancelado
        If iStatus <> 7 Then
        
            '??? em tese poderia gerar n registros 71 p/cada 70 mas isto ainda nao foi implementado. Jones 09/11/2001
            
            'se nao tenho dados minimos, entao vou pular a nf
            If NFsObtemNum(sNotasFiscais) <> "" And dValorMercadoria <> 0 Then
            
                objInfoArqICMS.sBuffer = "71" & FormataCpoNum(sCgc, 14) & InscEst_FormatadaArqICMS(sInscricaoEstadual) & FormataCpoData(dtDataEmissao) & FormataCpoTexto(sUF, 2) & FormataCpoNum(8, 2) & sSerieSubSerie & FormataCpoNum(lNumNotaFiscal, 6)
                
                If iCIF_FOB = FRETE_EMITENTE Then
                     objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoTexto(IIf(sUFDestinatario = "", "EX", sUFDestinatario), 2) & FormataCpoNum(sCGCDestinatario, 14) & FormataCpoTexto(IIf(Trim(sInscEstadualDestinatario) = "", "ISENTO", sInscEstadualDestinatario), 14)
                Else
                     objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoTexto(IIf(sUFRemetente = "", "EX", sUFRemetente), 2) & FormataCpoNum(sCGCRemetente, 14) & FormataCpoTexto(IIf(Trim(sInscEstadualRemetente) = "", "ISENTO", sInscEstadualRemetente), 14)
                End If
                
                'alteracao convenio 69/02
                If objInfoArqICMS.dtDataFinal < #1/1/2003# Then
                    objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoData(dtDataEmissao) & FormataCpoTexto("01    ", 5) & FormataCpoNum(NFsObtemNum(sNotasFiscais), 6) & FormataCpoValor(dValorMercadoria, 14) & FormataCpoTexto("", 12)
                Else
                    objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoData(dtDataEmissao) & FormataCpoTexto("01   ", 5) & FormataCpoNum(NFsObtemNum(sNotasFiscais), 6) & FormataCpoValor(dValorMercadoria, 14) & FormataCpoTexto("", 12)
                End If
                
                objInfoArqICMS.sChaveOrd = FormataCpoData(dtDataEmissao) & FormataCpoNum(lNumNotaFiscal, 6)
                
                'Conta o Registro 71 para gerar o Tipo 90
                objInfoArqICMS.lTotalReg71 = objInfoArqICMS.lTotalReg71 + 1
        
                'e Grava no BD o Tipo 71
                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 71, alComando(1))
                If lErro <> SUCESSO Then gError 81766
        
            End If
            
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81767
    
    Loop
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Processar_ConhecFrete_Transp = SUCESSO
     
    Exit Function
    
Erro_Processar_ConhecFrete_Transp:

    Processar_ConhecFrete_Transp = gErr
     
    Select Case gErr
          
        Case 81766, 81768
        
        Case 81764, 81765, 81767
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CF_ARQICMS", gErr)
        
        Case 81763
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160742)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function NFsObtemNum(sNotasFiscais As String) As String

Dim sResult As String, iTam As Integer, iIndice As Integer, sChar As String

    For iIndice = 1 To Len(sNotasFiscais)
    
        sChar = Mid(sNotasFiscais, iIndice, 1)
        
        If InStr("0123456789", sChar) = 0 Then Exit For
    
        sResult = sResult & sChar
        
    Next
    
    If Len(sResult) > 6 Then sResult = right(sResult, 6)
    
    NFsObtemNum = sResult
    
End Function

'Não Subir
Private Function InserirRegsArqICMS(objFilial As AdmFiliais, objInfoArqICMS As ClassInfoArqICMS) As Long
'gera os arquivos para todas as UFs destino (interestaduais)

Dim lErro As Long, sUFDestino As String, lComando As Long

On Error GoTo Erro_InserirRegsArqICMS

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81874
    
    sUFDestino = String(STRING_ESTADO_SIGLA, 0)
    
    'foi feito link c/tabela de estados p/garantir que sejam estados validos
    lErro = Comando_Executar(lComando, "SELECT DISTINCT UFDest FROM GeracaoArqICMS, Estados WHERE GeracaoArqICMS.UFDest = Estados.Sigla", sUFDestino)
    If lErro <> AD_SQL_SUCESSO Then gError 81875
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81876
    
    Do While lErro = AD_SQL_SUCESSO
    
        objInfoArqICMS.sUFDestino = sUFDestino
        
        lErro = InserirRegsArqICMS1(objFilial, objInfoArqICMS)
        If lErro <> SUCESSO Then gError 81877
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81878
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    InserirRegsArqICMS = SUCESSO

    Exit Function

Erro_InserirRegsArqICMS:

    InserirRegsArqICMS = gErr

    Select Case gErr

        Case 81874
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81875, 81876, 81878
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_GERACAOARQICMS", gErr)
        
        Case 81877

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160743)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'****feito por Cyntia******

Function Processar_CuponsFiscais(objInfoArqICMS As ClassInfoArqICMS) As Long
'gerar registros 60 - subtipos M e A
'Função modificada por Wagner

Dim iIndice As Integer
Dim alComando(0 To 7) As Long
Dim tLojaArqFisICMS As typeLojaArqFisICMS
Dim iCodECF As Integer
Dim dtData As Date
Dim iFilial As Integer
Dim lErro As Long
Dim dCancICMS As Double
Dim dDescICMS As Double
Dim dTotalizador As Double
Dim iISS As Integer
Dim dAliquota As Double
Dim dVendaBruta As Double

On Error GoTo Erro_Processar_CuponsFiscais

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 112082
    Next
    
    With tLojaArqFisICMS
    
    .sNumSerieECF = String(50, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT CodECF, FilialEmpresa, Data, NumSerieECF, COOIni, COOFim, CRZ, GrandeTotal, CRO, VendaBruta FROM LojaArqFisMestre WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? ", .iCodECF, .iFilialEmpresa, .dtData, .sNumSerieECF, .lCOOIni, .lCOOFim, .lCRZ, .dGrandeTotal, .lCRO, .dVendaBruta, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 112083

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112084
    
    'geração dos mestres
    Do While lErro = AD_SQL_SUCESSO

        dVendaBruta = 0
        
        .sSituacaoTrib = String(4, 0)
        
        lErro = Comando_Executar(alComando(2), "SELECT SituacaoTrib, Totalizador FROM LojaArqFisAnalitico WHERE CodECF = ? AND Data = ? AND FilialEmpresa = ? AND (SituacaoTrib = 'F' OR SituacaoTrib = 'N' OR SituacaoTrib = 'I' OR SituacaoTrib = 'S') ", .sSituacaoTrib, .dTotalizador, .iCodECF, .dtData, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 112086
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112087
        
        'geração dos analíticos
        Do While lErro = AD_SQL_SUCESSO
    
            If .sSituacaoTrib = "S" Then .sSituacaoTrib = "ISS"
    
            objInfoArqICMS.sBuffer = "60" & "A" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & FormataCpoTexto(.sSituacaoTrib, 4) & FormataCpoValor(.dTotalizador, 12) & FormataCpoTexto("", 79)
            dVendaBruta = dVendaBruta + .dTotalizador
            objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 2
            
            'Conta o Registro 60
            objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
    
            'e Grava no BD o Tipo 60
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(3))
            If lErro <> SUCESSO Then gError 112088
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112089
        
        Loop
            
        lErro = Comando_Executar(alComando(4), "SELECT CancelamentoICMS, DescontoICMS FROM LojaReducao WHERE CodECF = ? AND DataMovimento = ? AND FilialEmpresa = ?", dCancICMS, dDescICMS, .iCodECF, .dtData, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 210955
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210956
        
        'geração dos analíticos
        Do While lErro = AD_SQL_SUCESSO
    
            objInfoArqICMS.sBuffer = "60" & "A" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & "CANC" & FormataCpoValor(dCancICMS, 12) & FormataCpoTexto("", 79)
            dVendaBruta = dVendaBruta + dCancICMS
        
            objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 2
            
            'Conta o Registro 60
            objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
    
            'e Grava no BD o Tipo 60
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(5))
            If lErro <> SUCESSO Then gError 210957
            
            objInfoArqICMS.sBuffer = "60" & "A" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & "DESC" & FormataCpoValor(dDescICMS, 12) & FormataCpoTexto("", 79)
            dVendaBruta = dVendaBruta + dDescICMS
        
            objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 2
            
            'Conta o Registro 60
            objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
    
            'e Grava no BD o Tipo 60
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(5))
            If lErro <> SUCESSO Then gError 210957
            
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210958
        
        Loop
        
                        
        lErro = Comando_Executar(alComando(6), "SELECT Aliquota, Totalizador FROM LojaReducaoAliquota WHERE CodECF = ? AND Data = ? AND FilialEmpresa = ? AND ISS=0", dAliquota, dTotalizador, .iCodECF, .dtData, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 210955
    
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210956
        
        'geração dos analíticos
        Do While lErro = AD_SQL_SUCESSO
    
            objInfoArqICMS.sBuffer = "60" & "A" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & Format(dAliquota * 100, "0000") & FormataCpoValor(dTotalizador, 12) & FormataCpoTexto("", 79)
            dVendaBruta = dVendaBruta + dTotalizador
        
            objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 2
            
            'Conta o Registro 60
            objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
    
            'e Grava no BD o Tipo 60
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(7))
            If lErro <> SUCESSO Then gError 210957
            
            
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210958
        
        Loop
    
        objInfoArqICMS.sBuffer = "60" & "M" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & FormataCpoNum(.iCodECF, 3) & "2D" & FormataCpoNum(.lCOOIni, 6) & FormataCpoNum(.lCOOFim, 6) & FormataCpoNum(.lCRZ, 6) & FormataCpoNum(.lCRO, 3) & FormataCpoValor(dVendaBruta, 16) & FormataCpoValor(.dGrandeTotal, 16) & FormataCpoTexto("", 37)
    
        objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 1
        
        'Conta o Registro 60
        objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1

        'e Grava no BD o Tipo 60
        lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(1))
        If lErro <> SUCESSO Then gError 112085
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112090
    
    Loop
    
    End With
   
    'gerar o D, I  e R para cada produto caso a unidade federada solicite
    lErro = Gera_60Subtipos_I(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 112141
    
    If objInfoArqICMS.sUF = "XX" Then
            lErro = Gera_60Subtipos_D(objInfoArqICMS)
            If lErro <> SUCESSO Then gError 112141
    End If
    
    If objInfoArqICMS.sUF = "XX" Then
            lErro = Gera_60Subtipos_R(objInfoArqICMS)
            If lErro <> SUCESSO Then gError 112141
    End If
    
    lErro = Gera_61(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 112141
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Processar_CuponsFiscais = SUCESSO
     
    Exit Function
    
Erro_Processar_CuponsFiscais:

    Processar_CuponsFiscais = gErr
     
    Select Case gErr
        
        Case 112082
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 112083, 112084, 112090
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAARQFISMESTRE", gErr)
        
        Case 112085, 112088, 112141, 210952, 210953
        
        Case 112086, 112087, 112089
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAARQFISANALITICO", gErr)
        
        Case 210950, 210951, 210954
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAREDUCAO", gErr)
        
        Case 210955, 210956, 210958
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOJAREDUCAOALIQUOTA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160744)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

'#################################################################
'Inserido por Wagner

Function Gera_60Subtipos_I(objInfoArqICMS As ClassInfoArqICMS) As Long
'gerar registros 60 - subtipo I

Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim tLojaArqFisICMS As typeLojaArqFisICMS
Dim iItem As Integer
Dim dPrecoUnitario As Double
Dim lErro As Long
Dim sProduto As String
Dim dValorDesconto As Double
Dim bAchou As Boolean
Dim objItemNF As ClassItemNF
Dim sUM As String
Dim objProduto As ClassProduto
Dim iTipo As Integer
Dim objTribItemNF As ClassTributacaoDocItem
Dim sAliquota As String
Dim iStatus As Integer

On Error GoTo Erro_Gera_60Subtipos_I

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 112144
    Next
    
    With tLojaArqFisICMS
    
    .sNumSerieECF = String(50, 0)
    sProduto = String(20, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    .sSituacaoTrib = String(STRING_ITEMCFISCAL_SITUACAOTRIB, 0)
    
    'gera o subtipo I
    lErro = Comando_Executar(alComando(0), "SELECT C.Numero, C.DataEmissao, L.NumSerieECF, I.Produto, I.Quantidade, I.Item, I.PrecoUnitario, I.ValorDesconto," & _
 "I.AliquotaICMS, I.UnidadeMed , T.Tipo, I.SituacaoTrib, I.Status FROM CupomFiscal AS C INNER JOIN  ItensCupomFiscal AS I ON " & _
    "I.NumIntCupom = C.NumIntDoc AND I.FilialEmpresa = C.FilialEmpresa INNER JOIN " & _
 "LojaArqFisMestre  AS L ON L.CodECF = C.ECF  LEFT OUTER JOIN TiposTribICMS AS T ON T.SituacaoTribECF = Left(I.SituacaoTrib,1) " & _
 " WHERE C.NFeChaveAcesso = '' AND C.DataReducao >= ? AND C.DataReducao <= ? AND C.FilialEmpresa = L.FilialEmpresa  AND L.Data = C.DataEmissao AND " & _
  " C.FilialEmpresa= ? AND C.Tipo = 1 ORDER BY C.NumSerieECF, C.DataEmissao", _
    .lCOOIni, .dtData, .sNumSerieECF, sProduto, .dTotalizador, iItem, dPrecoUnitario, dValorDesconto, .dAliquota, sUM, iTipo, _
    .sSituacaoTrib, iStatus, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 112145
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
    
    Do While lErro = AD_SQL_SUCESSO
    
        If iStatus = STATUS_CANCELADO Then
            sAliquota = "CANC"
            .dAliquota = 0
        ElseIf left(.sSituacaoTrib, 1) = "T" Then
            sAliquota = right(.sSituacaoTrib, 4)
        ElseIf left(.sSituacaoTrib, 1) = "S" Then
            sAliquota = "ISS"
            .dAliquota = 0
        Else
            sAliquota = left(.sSituacaoTrib, 1)
            .dAliquota = 0
        End If
    
                    
        objInfoArqICMS.sBuffer = "60" & "I" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & "2D" & FormataCpoNum(.lCOOIni, 6) & FormataCpoNum(iItem, 3) & FormataCpoTexto(sProduto, 14) & FormataCpoNum(Fix(.dTotalizador * 1000), 13) & FormataCpoNum(Fix(dPrecoUnitario * 1000), 13) & FormataCpoValor((dPrecoUnitario - dValorDesconto), 12) & FormataCpoTexto(sAliquota, 4) & FormataCpoValor((dPrecoUnitario - dValorDesconto) * .dAliquota, 12) & FormataCpoTexto("", 16)
    
        objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 4
        
        'Conta o Registro 60
        objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1

        'e Grava no BD o Tipo 60
        lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(1))
        If lErro <> SUCESSO Then gError 112142
        
        Set objItemNF = New ClassItemNF
        Set objProduto = New ClassProduto
        Set objTribItemNF = New ClassTributacaoDocItem
        
        objProduto.sCodigo = sProduto
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 138396
        
        objItemNF.sProduto = sProduto
        objItemNF.sUnidadeMed = sUM
        objItemNF.sDescricaoItem = objProduto.sDescricao
        objTribItemNF.iICMSTipo = iTipo
        
        Set objItemNF.objTributacaoItemNF = objTribItemNF
        
        lErro = CF("GeracaoArqICMSProdCupom_Grava", sProduto, objInfoArqICMS.sUFDestino, objItemNF, .dtData, alComando(1), alComando(2))
        If lErro <> SUCESSO Then gError 138395
                    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112143
                    
    Loop
    
    End With
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Gera_60Subtipos_I = SUCESSO
     
    Exit Function
    
Erro_Gera_60Subtipos_I:

    Gera_60Subtipos_I = gErr
     
    Select Case gErr
        
        Case 112144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 112143, 112145, 112146
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CF_ARQICMS", gErr)
        
        Case 112142, 138395, 138396
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160745)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Gera_60Subtipos_R(objInfoArqICMS As ClassInfoArqICMS) As Long
'gerar registros 60 - subtipo R

Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim tLojaArqFisICMS As typeLojaArqFisICMS
Dim iItem As Integer
Dim dPrecoUnitario As Double
Dim lErro As Long
Dim sProduto As String
Dim dValorDesconto As Double
Dim colProdutoD As New Collection
Dim colProdutoR As New Collection
Dim objItensCF As New ClassItemNF
Dim sNumSerie As String
Dim dtData As Date
Dim bAchou As Boolean
Dim sUM As String
Dim objProduto As ClassProduto
Dim iTipo As Integer
Dim objTribItemNF As ClassTributacaoDocItem
Dim bAcabou As Boolean
Dim sAliquota As String

On Error GoTo Erro_Gera_60Subtipos_R

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 112144
    Next
    
    With tLojaArqFisICMS
    
    .sNumSerieECF = String(50, 0)
    sProduto = String(20, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    .sSituacaoTrib = String(4, 0)
    
    'gera o subtipo R
    lErro = Comando_Executar(alComando(0), "SELECT C.Numero, C.DataEmissao, L.NumSerieECF, I.Produto, I.Quantidade, I.Item, I.PrecoUnitario, I.ValorDesconto, I.AliquotaICMS, I.UnidadeMed , T.Tipo, I.SituacaoTrib FROM CupomFiscal AS C, ItensCupomFiscal AS I, LojaArqFisMestre  AS L, TiposTribICMS AS T WHERE C.NFeChaveAcesso = '' AND L.CodECF = C.ECF AND C.FilialEmpresa=? AND I.NumIntCupom = C.NumIntDoc AND I.FilialEmpresa = C.FilialEmpresa AND C.DataEmissao >= ? AND C.DataEmissao <= ? AND C.FilialEmpresa = L.FilialEmpresa  AND L.Data = C.DataEmissao AND T.SituacaoTribECF = I.SituacaoTrib AND C.Tipo = 1 ORDER BY NumSerieECF, DataEmissao", .lCOOIni, .dtData, .sNumSerieECF, sProduto, .dTotalizador, iItem, dPrecoUnitario, dValorDesconto, .dAliquota, sUM, iTipo, .sSituacaoTrib, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 112145
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
    
    dtData = .dtData
    
    bAcabou = False
    
    If lErro = AD_SQL_SEM_DADOS Then bAcabou = True
    
    Do While Not bAcabou
                                                                        
        If lErro = AD_SQL_SEM_DADOS Then bAcabou = True
                                        
        '***gera quando chegar o momento os registros do subtipo R
        'toda vez q alterar o mes  --> gravo registros novos
        If (Month(dtData) <> Month(.dtData)) Or bAcabou Then
            
            'Para cada produto...
            For Each objItensCF In colProdutoR
            
                If objItensCF.dPercDesc = 0 Then
                    sAliquota = FormataCpoTexto(objItensCF.sDescricaoItem, 4)
                Else
                    sAliquota = FormataCpoValor(objItensCF.dPercDesc * 100, 4)
                End If
           
                objInfoArqICMS.sBuffer = "60" & "R" & Format(dtData, "mmyyyy") & FormataCpoTexto(objItensCF.sProduto, 14) & FormataCpoNum(Fix(objItensCF.dQuantidade * 1000), 13) & FormataCpoNum(Fix(objItensCF.dPrecoUnitario * 100), 16) & FormataCpoNum(Fix((objItensCF.dPrecoUnitario - objItensCF.dValorDesconto) * 100), 16) & sAliquota & FormataCpoTexto("", 54)
                objInfoArqICMS.sChaveOrd = 5 & Month(dtData) & Year(dtData) & FormataCpoNum(objItensCF.sProduto, 14)
                
                'Conta o Registro 60
                objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
        
                'e Grava no BD o Tipo 60
                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(2))
                If lErro <> SUCESSO Then gError 112147
            
                Set objProduto = New ClassProduto
        
                objProduto.sCodigo = objItensCF.sProduto
        
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 138396
        
                objItensCF.sDescricaoItem = objProduto.sDescricao
        
                lErro = CF("GeracaoArqICMSProdCupom_Grava", objItensCF.sProduto, objInfoArqICMS.sUFDestino, objItensCF, .dtData, alComando(1), alComando(2))
                If lErro <> SUCESSO Then gError 138395
            
            Next
            
            dtData = .dtData
            
            Set colProdutoR = New Collection
            
        End If
                    
        For Each objItensCF In colProdutoR
             If objItensCF.sProduto = sProduto Then
                objItensCF.dPrecoUnitario = objItensCF.dPrecoUnitario + (dPrecoUnitario * .dTotalizador)
                objItensCF.dQuantidade = objItensCF.dQuantidade + .dTotalizador
                objItensCF.dValorDesconto = objItensCF.dValorDesconto + (dValorDesconto)
                bAchou = True
                Exit For
             End If
        Next
        If Not (bAchou) Then
            Set objItensCF = New ClassItemNF
            Set objTribItemNF = New ClassTributacaoDocItem
            
            objItensCF.sProduto = sProduto
            objItensCF.dPrecoUnitario = (dPrecoUnitario * .dTotalizador)
            objItensCF.dQuantidade = .dTotalizador
            objItensCF.dValorDesconto = dValorDesconto
            objItensCF.dPercDesc = .dAliquota
            objItensCF.sUnidadeMed = sUM
            objItensCF.sDescricaoItem = .sSituacaoTrib
            
            objTribItemNF.iICMSTipo = iTipo
            
            Set objItensCF.objTributacaoItemNF = objTribItemNF
            
            colProdutoR.Add objItensCF
        End If
                
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112143
    
    Loop
    
    End With
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Gera_60Subtipos_R = SUCESSO
     
    Exit Function
    
Erro_Gera_60Subtipos_R:

    Gera_60Subtipos_R = gErr
     
    Select Case gErr
        
        Case 112144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 112145, 112146, 112143
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CF_ARQICMS", gErr)
        
        Case 112142, 112147, 112148, 138395
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160746)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Gera_60Subtipos_D(objInfoArqICMS As ClassInfoArqICMS) As Long
'gerar registros 60 - subtipo D

Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim tLojaArqFisICMS As typeLojaArqFisICMS
Dim iItem As Integer
Dim dPrecoUnitario As Double
Dim lErro As Long
Dim sProduto As String
Dim dValorDesconto As Double
Dim colProdutoD As New Collection
Dim objItensCF As New ClassItemNF
Dim sNumSerie As String
Dim dtData As Date
Dim bAchou As Boolean
Dim sUM As String
Dim objProduto As ClassProduto
Dim iTipo As Integer
Dim objTribItemNF As ClassTributacaoDocItem
Dim bAcabou As Boolean
Dim sAliquota As String

On Error GoTo Erro_Gera_60Subtipos_D

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 112144
    Next
    
    With tLojaArqFisICMS
    
    .sNumSerieECF = String(50, 0)
    sProduto = String(20, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    .sSituacaoTrib = String(4, 0)
    
    'gera o subtipo D
    lErro = Comando_Executar(alComando(0), "SELECT C.Numero, C.DataEmissao, L.NumSerieECF, I.Produto, I.Quantidade, I.Item, I.PrecoUnitario, I.ValorDesconto, I.AliquotaICMS, I.UnidadeMed , T.Tipo, I.SituacaoTrib FROM CupomFiscal AS C, ItensCupomFiscal AS I, LojaArqFisMestre  AS L, TiposTribICMS AS T WHERE C.NFeChaveAcesso = '' AND L.CodECF = C.ECF AND C.FilialEmpresa=? AND I.NumIntCupom = C.NumIntDoc AND I.FilialEmpresa = C.FilialEmpresa AND C.DataEmissao >= ? AND C.DataEmissao <= ? AND C.FilialEmpresa = L.FilialEmpresa  AND L.Data = C.DataEmissao AND T.SituacaoTribECF = I.SituacaoTrib AND C.Tipo = 1 ORDER BY NumSerieECF, DataEmissao", .lCOOIni, .dtData, .sNumSerieECF, sProduto, .dTotalizador, iItem, dPrecoUnitario, dValorDesconto, .dAliquota, sUM, iTipo, .sSituacaoTrib, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 112145
       
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
    
    dtData = .dtData
    sNumSerie = .sNumSerieECF
    
    bAcabou = False
    
    If lErro = AD_SQL_SEM_DADOS Then bAcabou = True
    
    Do While Not bAcabou
                           
        If lErro = AD_SQL_SEM_DADOS Then bAcabou = True
                                                     
        '***gera quando chegar o momento os registros do subtipo D
        'toda vez q alterar o dia ou meu ECf --> gravo registros novos
        If (dtData <> .dtData Or sNumSerie <> .sNumSerieECF) Or bAcabou Then
            
            'Para cada produto...
            For Each objItensCF In colProdutoD
                        
                If objItensCF.dPercDesc = 0 Then
                    sAliquota = FormataCpoTexto(objItensCF.sDescricaoItem, 4)
                Else
                    sAliquota = FormataCpoValor(objItensCF.dPercDesc * 100, 4)
                End If
                        
                objInfoArqICMS.sBuffer = "60" & "D" & FormataCpoData(dtData) & FormataCpoTexto(sNumSerie, 20) & FormataCpoTexto(objItensCF.sProduto, 14) & FormataCpoNum(Fix(objItensCF.dQuantidade * 1000), 13) & FormataCpoNum(Fix(objItensCF.dPrecoUnitario * 100), 16) & FormataCpoNum(Fix((objItensCF.dPrecoUnitario - objItensCF.dValorDesconto) * 100), 16) & sAliquota & FormataCpoValor(objItensCF.dPrecoUnitario * objItensCF.dPercDesc, 13) & FormataCpoTexto("", 19)
                objInfoArqICMS.sChaveOrd = FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & 3
                
                'Conta o Registro 60
                objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
        
                'e Grava no BD o Tipo 60
                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(2))
                If lErro <> SUCESSO Then gError 112148
            
                Set objProduto = New ClassProduto
        
                objProduto.sCodigo = objItensCF.sProduto
        
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 138396
        
                objItensCF.sDescricaoItem = objProduto.sDescricao
        
                lErro = CF("GeracaoArqICMSProdCupom_Grava", objItensCF.sProduto, objInfoArqICMS.sUFDestino, objItensCF, .dtData, alComando(1), alComando(2))
                If lErro <> SUCESSO Then gError 138395
            
            Next
            
            dtData = .dtData
            sNumSerie = .sNumSerieECF
            
            Set colProdutoD = New Collection
        End If
                    
        For Each objItensCF In colProdutoD
            If objItensCF.sProduto = sProduto Then
                objItensCF.dPrecoUnitario = objItensCF.dPrecoUnitario + (dPrecoUnitario * .dTotalizador)
                objItensCF.dQuantidade = objItensCF.dQuantidade + .dTotalizador
                objItensCF.dValorDesconto = objItensCF.dValorDesconto + (dValorDesconto)
                               
                bAchou = True
                Exit For
             End If
        Next
        If Not (bAchou) Then
            Set objItensCF = New ClassItemNF
            Set objTribItemNF = New ClassTributacaoDocItem
            
            objItensCF.sProduto = sProduto
            objItensCF.dQuantidade = .dTotalizador
            objItensCF.dPrecoUnitario = dPrecoUnitario * .dTotalizador
            objItensCF.dValorDesconto = dValorDesconto
            objItensCF.dPercDesc = .dAliquota
            objItensCF.sUnidadeMed = sUM
            objItensCF.sDescricaoItem = .sSituacaoTrib
            
            objTribItemNF.iICMSTipo = iTipo
            
            Set objItensCF.objTributacaoItemNF = objTribItemNF
            
            colProdutoD.Add objItensCF
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112143
    
    Loop
    
    End With
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Gera_60Subtipos_D = SUCESSO
     
    Exit Function
    
Erro_Gera_60Subtipos_D:

    Gera_60Subtipos_D = gErr
     
    Select Case gErr
        
        Case 112144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 112145, 112146, 112143
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CF_ARQICMS", gErr)
        
        Case 112142, 112147, 112148, 138395
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160747)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

'Function Gera_60Subtipos(objInfoArqICMS As ClassInfoArqICMS) As Long
''gerar registros 60 - subtipo D, I e R
'
'Dim iIndice As Integer
'Dim alComando(0 To 2) As Long
'Dim tLojaArqFisICMS As typeLojaArqFisICMS
'Dim iItem As Integer
'Dim dPrecoUnitario As Double
'Dim lErro As Long
'Dim sProduto As String
'Dim dValorDesconto As Double
'Dim colProdutoD As New Collection
'Dim colProdutoR As New Collection
'Dim objItensCF As New ClassItemNF
'Dim sNumSerie As String
'Dim dtData As Date
'Dim bAchou As Boolean
'
'On Error GoTo Erro_Gera_60Subtipos
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 112144
'    Next
'
'    With tLojaArqFisICMS
'
'    .sNumSerieECF = String(10, 0)
'    sProduto = String(20, 0)
'
'    'gera o subtipo I
'    lErro = Comando_Executar(alComando(0), "SELECT Numero, DataEmissao, NumSerieECF, Produto, Quantidade, Item, PrecoUnitario, ValorDesconto, AliquotaICMS FROM CupomFiscal, ItensCupomFiscal, LojaArqFisMestre WHERE LojaArqFisMestre.CodECF = CupomFiscal.ECF AND CupomFiscal.FilialEmpresa=? AND ItensCupomFiscal.NumIntCupom = CupomFiscal.NumIntDoc AND ItensCupomFiscal.FilialEmpresa = CupomFiscal.FilialEmpresa AND  DataEmissao >= ? AND DataEmissao <= ? AND CupomFiscal.FilialEmpresa = LojaArqFisMestre.FilialEmpresa  AND LojaArqFisMestre.Data = DataEmissao ORDER BY NumSerieECF, DataEmissao", .lContadorInicio, .dtData, .sNumSerieECF, sProduto, .dTotalizador, iItem, dValorDesconto, .dAliquota, giFilialEmpresa, CDate(DataInicial.Text), CDate(DataFinal.Text))
'    If lErro <> AD_SQL_SUCESSO Then gError 112145
'
'    dtData = .dtData
'    sNumSerie = .sNumSerieECF
'
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        objInfoArqICMS.sBuffer = "60" & "I" & FormataCpoData(.dtData) & FormataCpoTexto(.sNumSerieECF, 20) & "2D" & FormataCpoNum(.lContadorInicio, 6) & FormataCpoNum(iItem, 3) & FormataCpoTexto(sProduto, 14) & FormataCpoNum(.dTotalizador * 1000, 13) & FormataCpoNum(dPrecoUnitario * 1000, 13) & FormataCpoNum((dPrecoUnitario - dValorDesconto) * 100, 12) & FormataCpoNum(.dAliquota * 100, 4) & FormataCpoNum((dPrecoUnitario - dValorDesconto) * .dAliquota, 12) & FormataCpoTexto("", 16)
'
'        objInfoArqICMS.sChaveOrd = .dtData & FormataCpoNum(.sNumSerieECF, 20) & 4
'
'        'Conta o Registro 60
'        objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
'
'        'e Grava no BD o Tipo 60
'        lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(1))
'        If lErro <> SUCESSO Then gError 112142
'
'        '***gera quando chegar o momento os registros do subtipo R
'        'toda vez q alterar o mes  --> gravo registros novos
'        If Month(dtData) <> Month(.dtData) Then
'
'            'Para cada produto...
'            For Each objItensCF In colProdutoR
'
'                objInfoArqICMS.sBuffer = "60" & "R" & Format(dtData, "mmaaaa") & FormataCpoTexto(objItensCF.sProduto, 14) & FormataCpoNum(objItensCF.dQuantidade * 1000, 13) & FormataCpoNum(objItensCF.dValorTotal * 100, 16) * FormataCpoNum(objItensCF.dValorDesconto * 100, 16) & FormataCpoNum(objItensCF.dPercDesc * 100, 4) & FormataCpoTexto("", 54)
'
'                objInfoArqICMS.sChaveOrd = 5 & Month(dtData) & Year(dtData) & FormataCpoNum(objItensCF.sProduto, 14)
'
'                'Conta o Registro 60
'                objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
'
'                'e Grava no BD o Tipo 60
'                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(2))
'                If lErro <> SUCESSO Then gError 112147
'            Next
'
'            Set colProdutoR = New Collection
'        End If
'
'        For Each objItensCF In colProdutoR
'             If objItensCF.sProduto = sProduto Then
'                objItensCF.dValorTotal = objItensCF.dValorTotal + dPrecoUnitario
'                objItensCF.dQuantidade = objItensCF.dQuantidade + .dTotalizador
'                objItensCF.dValorDesconto = objItensCF.dValorDesconto + (dPrecoUnitario - dValorDesconto)
'                bAchou = True
'                Exit For
'             End If
'        Next
'        If Not (bAchou) Then
'            Set objItensCF = New ClassItemNF
'            objItensCF.sProduto = sProduto
'            objItensCF.dValorTotal = dPrecoUnitario
'            objItensCF.dQuantidade = .dTotalizador
'            objItensCF.dValorDesconto = dPrecoUnitario - dValorDesconto
'            objItensCF.dPercDesc = .dAliquota
'
'            colProdutoR.Add objItensCF
'        End If
'
'        '***gera quando chegar o momento os registros do subtipo D
'        'toda vez q alterar o dia ou meu ECf --> gravo registros novos
'        If dtData <> .dtData Or sNumSerie <> .sNumSerieECF Then
'
'            'Para cada produto...
'            For Each objItensCF In colProdutoD
'
'                objInfoArqICMS.sBuffer = "60" & "D" & FormataCpoData(dtData) & FormataCpoNum(sNumSerie, 20) & FormataCpoTexto(objItensCF.sProduto, 14) & FormataCpoNum(objItensCF.dQuantidade * 1000, 13) & FormataCpoNum(objItensCF.dValorTotal * 100, 16) * FormataCpoNum(objItensCF.dValorDesconto * 100, 16) & FormataCpoNum(objItensCF.dPercDesc * 100, 4) & FormataCpoNum(objItensCF.dValorDesconto * objItensCF.dPercDesc, 13) & FormataCpoTexto("", 19)
'
'                objInfoArqICMS.sChaveOrd = .dtData & FormataCpoNum(.sNumSerieECF, 20) & 3
'
'                'Conta o Registro 60
'                objInfoArqICMS.lTotalReg60 = objInfoArqICMS.lTotalReg60 + 1
'
'                'e Grava no BD o Tipo 60
'                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 60, alComando(2))
'                If lErro <> SUCESSO Then gError 112148
'            Next
'            dtData = .dtData
'            sNumSerie = .sNumSerieECF
'
'            Set colProdutoD = New Collection
'        End If
'
'        For Each objItensCF In colProdutoD
'            If objItensCF.sProduto = sProduto Then
'                objItensCF.dValorTotal = objItensCF.dValorTotal + dPrecoUnitario
'                objItensCF.dQuantidade = objItensCF.dQuantidade + .dTotalizador
'                objItensCF.dValorDesconto = objItensCF.dValorDesconto + (dPrecoUnitario - dValorDesconto)
'                bAchou = True
'                Exit For
'             End If
'        Next
'        If Not (bAchou) Then
'            Set objItensCF = New ClassItemNF
'            objItensCF.sProduto = sProduto
'            objItensCF.dValorTotal = dPrecoUnitario
'            objItensCF.dQuantidade = .dTotalizador
'            objItensCF.dValorDesconto = dPrecoUnitario - dValorDesconto
'            objItensCF.dPercDesc = .dAliquota
'
'            colProdutoD.Add objItensCF
'        End If
'
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112143
'
'    Loop
'
'    End With
'
'    'Fecha os  comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Comando_Fechar (alComando(iIndice))
'    Next
'
'    Gera_60Subtipos = SUCESSO
'
'    Exit Function
'
'Erro_Gera_60Subtipos:
'
'    Gera_60Subtipos = gErr
'
'    Select Case gErr
'
'        Case 112144
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 112145, 112146, 112143
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CF_ARQICMS", gErr)
'
'        Case 112142, 112147, 112148
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160748)
'
'    End Select
'
'    'Fecha os  comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Comando_Fechar (alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'###################################################################

Private Function Processar_RegsInventario(objInfoArqICMS As ClassInfoArqICMS) As Long
'gera os dados para os registros do tipo 74

Dim lErro As Long

On Error GoTo Erro_Processar_RegsInventario

    If objInfoArqICMS.dtReg74DataInicial <> DATA_NULA Then
        
        lErro = Processar_RegsInventario1(objInfoArqICMS, objInfoArqICMS.dtReg74DataInicial)
        If lErro <> SUCESSO Then gError 130025
        
    End If
    
    If objInfoArqICMS.dtReg74DataFinal <> DATA_NULA And objInfoArqICMS.dtReg74DataInicial <> objInfoArqICMS.dtReg74DataFinal Then
        
        lErro = Processar_RegsInventario1(objInfoArqICMS, objInfoArqICMS.dtReg74DataFinal)
        If lErro <> SUCESSO Then gError 130026
        
    End If
    
    Processar_RegsInventario = SUCESSO
     
    Exit Function
    
Erro_Processar_RegsInventario:

    Processar_RegsInventario = gErr
     
    Select Case gErr
          
        Case 130025, 130026
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184429)
     
    End Select
     
    Exit Function

End Function

Private Function Processar_RegsInventario1(objInfoArqICMS As ClassInfoArqICMS, ByVal dtDataInventario As Date) As Long
'gera os dados para os registros do tipo 74

Dim lErro As Long, iIndice As Integer, alComando(0 To 7) As Long
Dim sProduto As String, dQuantidade As Double, dValor As Double, dValorUnitario As Double

On Error GoTo Erro_Processar_RegsInventario1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130027
    Next

    'inicializar strings
    sProduto = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT Produto, QuantidadeUMEstoque, ValorUnitario FROM RegInventario, Produtos, NaturezasProduto WHERE RegInventario.Produto = Produtos.Codigo AND Produtos.Natureza = NaturezasProduto.Codigo AND InventarioICMS <> 0 AND FilialEmpresa = ? AND Data = ? ORDER BY Produto", sProduto, dQuantidade, dValorUnitario, objInfoArqICMS.iFilialEmpresa, dtDataInventario)
    If lErro <> AD_SQL_SUCESSO Then gError 130029

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130030

    Do While lErro = AD_SQL_SUCESSO

        If Abs(dQuantidade) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Processar_RegsInventario2(objInfoArqICMS, dtDataInventario, sProduto, alComando())
            '??? If lErro <> SUCESSO Then gError 130038
            If lErro = SUCESSO Then
            
                dValor = Arredonda_Moeda(dValorUnitario, 2)
                objInfoArqICMS.sBuffer = "74" & FormataCpoData(dtDataInventario) & FormataCpoTexto(sProduto, 14) & FormataCpoQuantidade(dQuantidade, 13) & FormataCpoValor(dValor, 13) & "1" & FormataCpoNum(0, 14) & FormataCpoTexto("", 14 + 2 + 45)
            
                objInfoArqICMS.sChaveOrd = FormataCpoData(dtDataInventario) & FormataCpoTexto(sProduto, 14)
                
                'Conta o Registro 74 para gerar o Tipo 90
                objInfoArqICMS.lTotalReg74 = objInfoArqICMS.lTotalReg74 + 1
        
                'e Grava no BD o Tipo 74
                lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 74, alComando(1))
                If lErro <> SUCESSO Then gError 130032
        
            End If
            
        End If
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130031
    
    Loop
    
    '??? falta pesquisar e incluir regs referentes a saldos em/de 3os
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Processar_RegsInventario1 = SUCESSO
     
    Exit Function
    
Erro_Processar_RegsInventario1:

    Processar_RegsInventario1 = gErr
     
    Select Case gErr
          
        Case 130027
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130029 To 130031
            Call Rotina_Erro(vbOKOnly, "ERRO_LEIURA_INVENTARIO_ARQICMS", gErr)
        
        Case 130032, 130038
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184428)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Processar_RegsInventario2(objInfoArqICMS As ClassInfoArqICMS, ByVal dtDataInventario As Date, ByVal sProduto As String, alComando() As Long) As Long

Dim lErro As Long, lNumIntDoc As Long, dtDataInicial As Date, dtDataFinal As Date, bEntra As Boolean
Dim objGeracaoArqICMSProd As New ClassGeracaoArqICMSProd

On Error GoTo Erro_Processar_RegsInventario2

    'Lê o cadastro Produto nos Livros Fiscais
    lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc, DataInicial, DataFinal FROM LivRegESCadProd WHERE Produto = ? AND DataInicial <= ? ORDER BY DataInicial DESC", _
            lNumIntDoc, dtDataInicial, dtDataFinal, sProduto, dtDataInventario)
    If lErro <> AD_SQL_SUCESSO Then gError 130033

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130034

    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 130035

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'inicializa a variavel que vai dizer se o o registro serah gravado ou nao
        bEntra = True
    
        'inicio da barreira
        'se a data nao for data_nula
        If dtDataFinal <> DATA_NULA Then
            
            'se data final do registro for menor a data inicial do periodo
            If dtDataFinal < dtDataInventario Then
            
                'se o registro nao possui data final como nula e
                'se a data final dele eh menor que a data do periodo
                'significa que ele nao deve ser gravado... logo
                'ele nao deve entrar...
                bEntra = False
        
            End If
        
        End If
        
        'se for pra gravar...
        If bEntra = True Then
    
            'e Grava o prototipo do Tipo 75 no BD para evitar repeticoes
            objGeracaoArqICMSProd.lNumIntLivRegESCadProd = lNumIntDoc

            'Só grava o tipo75 se ele ainda não estiver no BD
            lErro = CF("GeracaoArqICMSProd_Grava", objInfoArqICMS, objGeracaoArqICMSProd, alComando(6), alComando(7))
            If lErro <> SUCESSO Then gError 130036

        End If
        
        If dtDataInicial < dtDataInventario Then Exit Do
            
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130037
        
    Loop

    Processar_RegsInventario2 = SUCESSO
     
    Exit Function
    
Erro_Processar_RegsInventario2:

    Processar_RegsInventario2 = gErr
     
    Select Case gErr
          
        Case 130033, 130034, 130037
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_ARQICMS", gErr)
        
        Case 130035
'???            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_INFOPROD_ARQICMS", gErr)
        
        Case 130036
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184427)
     
    End Select
     
    Exit Function

End Function

Private Function Processar_SubstTributaria(objInfoArqICMS As ClassInfoArqICMS) As Long

Dim lErro As Long, lComando As Long, lComando2 As Long, iEmitente As Integer
Dim dFreteSegOutras As Double, iFilialEmpresa As Integer, lNumIntDocRegES As Long, iTipo As Integer, dtDataEmissao As Date, dtData As Date, sUF As String
Dim sSerie As String, iModelo As Integer, sCgc As String, sInscricaoEstadual As String, lNumNotaFiscal As Long, sNaturezaOp As String
Dim dICMSSubstBase As Double, dICMSSubstValor As Double, iSituacao As Integer, iIEIsento As Integer

On Error GoTo Erro_Processar_SubstTributaria

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 184707

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 184708

    sUF = String(STRING_ESTADO, 0)
    sSerie = String(STRING_SERIE, 0)
    sCgc = String(STRING_CGC, 0)
    sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    
    If objInfoArqICMS.bIntegral Then
        lErro = Comando_Executar(lComando, "SELECT Emitente, FreteSegOutras, Situacao, FilialEmpresa,NumIntDocRegES,Tipo,DataEmissao,Data,UF,Serie,modelo,CGC,InscricaoEstadual,IEIsento,NumNotaFiscal,NaturezaOP,ICMSSubstBase,ICMSSubstValor FROM GeracaoSintegraICMSSubst WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? ORDER BY Data, Tipo, NumNotaFiscal", _
            iEmitente, dFreteSegOutras, iSituacao, iFilialEmpresa, lNumIntDocRegES, iTipo, dtDataEmissao, dtData, sUF, sSerie, iModelo, sCgc, sInscricaoEstadual, iIEIsento, lNumNotaFiscal, sNaturezaOp, dICMSSubstBase, dICMSSubstValor, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    Else
        lErro = Comando_Executar(lComando, "SELECT Emitente, FreteSegOutras, Situacao, FilialEmpresa,NumIntDocRegES,Tipo,DataEmissao,Data,UF,Serie,modelo,CGC,InscricaoEstadual,IEIsento,NumNotaFiscal,NaturezaOP,ICMSSubstBase,ICMSSubstValor FROM GeracaoSintegraICMSSubst WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND UF <> ? ORDER BY Data, Tipo, NumNotaFiscal", _
            iEmitente, dFreteSegOutras, iSituacao, iFilialEmpresa, lNumIntDocRegES, iTipo, dtDataEmissao, dtData, sUF, sSerie, iModelo, sCgc, sInscricaoEstadual, iIEIsento, lNumNotaFiscal, sNaturezaOp, dICMSSubstBase, dICMSSubstValor, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, objInfoArqICMS.sUF)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 184709
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184710
    
    Do While lErro = AD_SQL_SUCESSO
    
        If Trim(sSerie) = "0" Then sSerie = "" '??? dava erro no validador 2005
        
        If lNumNotaFiscal <= 999999 And iModelo < 100 Then

            If objInfoArqICMS.bIntegral Then
                objInfoArqICMS.sUFDestino = objInfoArqICMS.sUF
            Else
                objInfoArqICMS.sUFDestino = sUF
            End If
    
            objInfoArqICMS.sBuffer = "53" & FormataCpoNum(sCgc, 14) & InscEst_FormatadaArqICMS(sInscricaoEstadual, iIEIsento) & FormataCpoData(dtData) & FormataCpoTexto(sUF, 2) & FormataCpoNum(iModelo, 2)
            
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoTexto(sSerie, 3) & FormataCpoNum(lNumNotaFiscal, 6) & FormataCpoNum(sNaturezaOp, 4) & IIf(iEmitente = EMITENTE_EMPRESA, "P", "T")
        
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoValor(dICMSSubstBase, 13) & FormataCpoValor(dICMSSubstValor, 13) & FormataCpoValor(dFreteSegOutras, 13) & FormataCpoTexto(IIf(iSituacao = 0, "N", "S"), 1) & " " & FormataCpoTexto("", 29)
        
            objInfoArqICMS.sChaveOrd = FormataCpoData(dtData)
    
            'Conta o Registro 53 para gerar o Tipo 90
            objInfoArqICMS.lTotalReg53 = objInfoArqICMS.lTotalReg53 + 1
    
            'e Grava no BD o Tipo 53
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 53, lComando2)
            If lErro <> SUCESSO Then gError 184711
    
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184712
    
    Loop
    
    'Fecha os  comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Processar_SubstTributaria = SUCESSO
    
    Exit Function
    
Erro_Processar_SubstTributaria:

    Processar_SubstTributaria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184706)

    End Select
    
    'Fecha os  comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Private Function Tamanho_Valida(ByVal sCampo As String, ByVal sValor As String, ByVal iTamanho As Integer) As Long

On Error GoTo Erro_Tamanho_Valida

    If Len(sValor) > iTamanho Then gError 200238

    Tamanho_Valida = SUCESSO
    
    Exit Function
    
Erro_Tamanho_Valida:

    Tamanho_Valida = gErr

    Select Case gErr
    
        Case 200238
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPO_COM_TAM_MAIOR_PERM", gErr, sCampo, sValor, Len(sValor), iTamanho)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200239)

    End Select
    
    Exit Function
    
End Function

Private Function InscEst_FormatadaArqICMS(sIE As String, Optional ByVal iIEIsento As Integer = -1) As String
    Dim sIESemPontos As String
    Call Formata_String_Numero(sIE, sIESemPontos)
    If iIEIsento <> -1 Then
        InscEst_FormatadaArqICMS = FormataCpoTexto(IIf(iIEIsento = MARCADO, "ISENTO", sIESemPontos), 14)
    Else
        InscEst_FormatadaArqICMS = FormataCpoTexto(IIf(Len(Trim(sIE)) = 0, "ISENTO", sIESemPontos), 14)
    End If
End Function

Private Function Trata_NFe_Inutilizadas(ByVal objInfoArqICMS As ClassInfoArqICMS, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long

Dim lErro As Long
Dim sEmitente As String
Dim lNumNFInut As Long, lNumNFDeInut As Long, lNumNFAteInut As Long
Dim sSerieNFInut As String, dtDataNFInut As Date, iFilialEmpNFInut As Integer, sCGCNFInut As String
Dim sNumDe As String, sNumAte As String

On Error GoTo Erro_Trata_Nfe_Inutilizadas

    sEmitente = "P"
    sSerieNFInut = String(STRING_SERIE, 0)
    sNumDe = String(STRING_MAXIMO, 0)
    sNumAte = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando1, "SELECT N.Data, N.Serie, N.nNFIni, N.nNFFim  FROM NFeFedRetInutNFe AS N WHERE N.Data BETWEEN ? AND ? AND N.tpAmb = ? AND N.cStat = ? AND N.FilialEmpresa = ? ORDER BY Data, nNFIni ", dtDataNFInut, sSerieNFInut, sNumDe, sNumAte, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal, 0, "102", objInfoArqICMS.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 211742
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211743
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lNumNFDeInut = StrParaLong(sNumDe)
        lNumNFAteInut = StrParaLong(sNumAte)

        For lNumNFInut = lNumNFDeInut To lNumNFAteInut
    
            objInfoArqICMS.sBuffer = "50" & FormataCpoNum("", 14) & InscEst_FormatadaArqICMS("", 0) & FormataCpoData(dtDataNFInut) & FormataCpoTexto("", 2) & FormataCpoNum(55, 2) & FormataCpoTexto(sSerieNFInut, 3)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(lNumNFInut, 6) & FormataCpoNum("", 4) & FormataCpoTexto(sEmitente, 1)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoValor(0, 13) & FormataCpoValor(0, 13) & FormataCpoValor(0, 13) & FormataCpoValor(0, 13) & FormataCpoValor(0, 13) & FormataCpoValor(0, 4) & FormataCpoTexto("4", 1)
            objInfoArqICMS.sChaveOrd = FormataCpoData(dtDataNFInut) & FormataCpoNum(55, 2) & FormataCpoTexto(sSerieNFInut, 3) & FormataCpoNum(lNumNFInut, 6)
    
            'Conta o Registro 50 para gerar o Tipo 90
            objInfoArqICMS.lTotalReg50 = objInfoArqICMS.lTotalReg50 + 1

            'e Grava no BD o Tipo 50
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 50, lComando2)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Next
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211744

    Loop

    Trata_NFe_Inutilizadas = SUCESSO
    
    Exit Function
    
Erro_Trata_Nfe_Inutilizadas:

    Trata_NFe_Inutilizadas = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 211742 To 211744
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184706)

    End Select
    
    Exit Function

End Function

Function FISConfig_GravarTrans(ByVal objFISConfig As ClassFISConfig) As Long
'Efetua a gravação dos dados com a transação já aberta em função _
 anterior

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sConteudo As String

On Error GoTo Erro_FISConfig_GravarTrans

    'Abre os  comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 80154
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 80167
    
    'Inicializa String
    sConteudo = String(STRING_CONTEUDO, 0)

    'Seleciona na tabela de FisConfig o codigo passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM FisConfig WHERE Codigo = ? AND FilialEmpresa =? ", 0, sConteudo, objFISConfig.sCodigo, objFISConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 80155

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80156

    'Nao encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then gError 80157

    'É uma alteracao => atualiza o campo Conteudo da tabela FisConfig
    If lErro = AD_SQL_SUCESSO Then
                    
        'Atualiza a tabela FisConfig
        lErro = Comando_ExecutarPos(lComando1, "UPDATE FisConfig SET Conteudo = ? ", lComando, objFISConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 80158

    End If
        
    'Fecha os  comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    FISConfig_GravarTrans = SUCESSO

    Exit Function
    
Erro_FISConfig_GravarTrans:
    
    FISConfig_GravarTrans = gErr
    
        Select Case gErr
        
        Case 80154, 80167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 80155, 80156
        
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FISCONFIG", gErr, objFISConfig.sCodigo)

        Case 80157, 80158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FISCONFIG", gErr, objFISConfig.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150347)

    End Select

    'Fecha os  comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function InfoArqICMS_Exclui(ByVal objInfoArqICMS As ClassInfoArqICMS) As Long
'Exclui o Arquivo de ICMS

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_InfoArqICMS_Exclui

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 69979

    lErro = InfoArqICMS_Exclui_Trans(objInfoArqICMS)
    If lErro <> SUCESSO Then gError 210959
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 69985

    InfoArqICMS_Exclui = SUCESSO

    Exit Function

Erro_InfoArqICMS_Exclui:

    InfoArqICMS_Exclui = gErr

    Select Case gErr

        Case 69979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 69985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 210959
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150346)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function InfoArqICMS_Exclui_Trans(ByVal objInfoArqICMS As ClassInfoArqICMS) As Long
'Exclui o Arquivo de ICMS

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 4) As Long
Dim lNumIntDoc As Long
Dim lNumIntDocProduto As Long

On Error GoTo Erro_InfoArqICMS_Exclui_Trans

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69978
    Next

    'Seleciona O Arquivo de ICMS apartir do NumIntDoc
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM InfoArqICMS WHERE NumIntDoc = ?", 0, lNumIntDoc, objInfoArqICMS.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 69980

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69981
        
    'Se não encontrou o Arquivo --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 69982
    
    'Faz "lock" em ApuracaoICMS
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 69983

    'Exclui Arquivo de ICMS do Banco de dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM InfoArqICMS", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 69984
    
    'Confirma a transação
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InfoArqICMS_Exclui_Trans = SUCESSO

    Exit Function

Erro_InfoArqICMS_Exclui_Trans:

    InfoArqICMS_Exclui_Trans = gErr

    Select Case gErr

        Case 69978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 69980, 69981
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOARQICMS", gErr)
        
        Case 69980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INFOARQICMS_NAO_CADASTRADO", gErr, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
        
        Case 69983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_INFOARQICMS", gErr)

        Case 69984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INFOARQICMS", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210960)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Gera_61(objInfoArqICMS As ClassInfoArqICMS) As Long
'gerar registros 60 - subtipo I

Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim tLojaArqFisICMS As typeLojaArqFisICMS
Dim iItem As Integer
Dim dPrecoUnitario As Double
Dim lErro As Long
Dim sProduto As String
Dim dValorDesconto As Double
Dim bAchou As Boolean
Dim objItemNF As ClassItemNF
Dim sUM As String
Dim objProduto As ClassProduto
Dim iTipo As Integer
Dim objTribItemNF As ClassTributacaoDocItem
Dim sAliquota As String
Dim iStatus As Integer
Dim lNumNotaFiscalMin As Long, lNumNotaFiscalMax As Long, sSerie As String, dtDataEmissao As Date, dTotal As Double

On Error GoTo Erro_Gera_61

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 112144
    Next
    
    sSerie = String(STRING_SERIE, 0)
    lErro = Comando_Executar(alComando(3), "SELECT MIN(NumNotaFiscal), MAX(NumNotaFiscal), Serie, DataEmissaoNF FROM CupomFiscal WHERE NFeChaveAcesso = '' AND Tipo = 1 AND FilialEmpresa = ? AND DataEmissaoNF >= ? AND DataEmissaoNF <=? AND NumNotaFiscal <> 0 GROUP BY Serie, DataEmissaoNF", _
        lNumNotaFiscalMin, lNumNotaFiscalMax, sSerie, dtDataEmissao, giFilialEmpresa, objInfoArqICMS.dtDataInicial, objInfoArqICMS.dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 112145
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
        
    Do While lErro <> AD_SQL_SEM_DADOS
        
        lErro = Comando_Executar(alComando(0), "SELECT SUM(I.Quantidade* I.PrecoUnitario-I.ValorDesconto) FROM CupomFiscal AS C INNER JOIN  ItensCupomFiscal AS I ON " & _
            "I.NumIntCupom = C.NumIntDoc AND I.FilialEmpresa = C.FilialEmpresa " & _
            " WHERE I.Status <> 7 AND C.Status <> 7 AND C.NumNotaFiscal <> 0 AND C.DataEmissaoNF = ? AND C.Serie = ? AND " & _
            " C.FilialEmpresa= ? AND C.Tipo = 1", _
            dTotal, dtDataEmissao, sSerie, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 112145
        
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
        
        Do While lErro = AD_SQL_SUCESSO
        
            objInfoArqICMS.sBuffer = "61" & FormataCpoTexto("", 28) & FormataCpoData(dtDataEmissao) & "02" & FormataCpoTexto("D  " & sSerie, 5) & FormataCpoNum(lNumNotaFiscalMin, 6) & FormataCpoNum(lNumNotaFiscalMax, 6) & FormataCpoNum(Arredonda_Moeda(dTotal) * 100, 13)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(Arredonda_Moeda(Arredonda_Moeda(dTotal)) * 100, 13)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(0, 12)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(0, 13)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(0, 13)
            objInfoArqICMS.sBuffer = objInfoArqICMS.sBuffer & FormataCpoNum(0, 4) & " "
            
            objInfoArqICMS.sChaveOrd = "61" & FormataCpoData(dtDataEmissao)
            
            'Conta o Registro 60
            objInfoArqICMS.lTotalReg61 = objInfoArqICMS.lTotalReg61 + 1
    
            'e Grava no BD o Tipo 60
            lErro = CF("GeracaoArqICMS_Grava_FIS", objInfoArqICMS, 61, alComando(1))
            If lErro <> SUCESSO Then gError 112142
            
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112143
                        
        Loop
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112146
    
    Loop
    
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Gera_61 = SUCESSO
     
    Exit Function
    
Erro_Gera_61:

    Gera_61 = gErr
     
    Select Case gErr
        
        Case 112144
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 112143, 112145, 112146
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CF_ARQICMS", gErr)
        
        Case 112142, 138395, 138396
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160745)
     
    End Select
     
    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Exit Function

End Function

Function DFC_Gera(ByVal iFilialEmpresa As Integer, ByVal sDiretorio As String, ByVal iAno As Integer) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objFilialEmpresa As New AdmFiliais
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim sLinha As String, dValor As Double
Dim sIE As String, sCRCContador As String
Dim lCount As Long, sCount As String
Dim sRegistro As String

On Error GoTo Erro_DFC_Gera

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213839
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213840
    Next
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
        
    Call Formata_String_Numero(objFilialEmpresa.sInscricaoEstadual, sIE)
    sIE = FormataCpoNum(sIE, 10)
    
    'Call Formata_String_Numero(objFilialEmpresa.sCRCContador, sCRCContador)
    sCRCContador = Replace(Replace(Replace(Replace(Mid(objFilialEmpresa.sCRCContador, 2), ".", ""), "-", ""), "/", ""), "\", "")
    sCRCContador = FormataCpoTexto(sCRCContador, 15)
    
    sLinha = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(alComando(0), "SELECT COUNT(*) FROM DFC_Linha WHERE FilialEmpresa = ? AND Ano = ?", _
        lCount, iFilialEmpresa, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 213841
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213842
    
    sCount = FormataCpoNum(lCount, 3)
    
    sRegistro = "1"
    sRegistro = sRegistro & "21"
    sRegistro = sRegistro & sIE
    sRegistro = sRegistro & objFilialEmpresa.sCgc
    sRegistro = sRegistro & CStr(iAno) & "00"
    sRegistro = sRegistro & "F"
    sRegistro = sRegistro & sCRCContador
    sRegistro = sRegistro & "8"
    sRegistro = sRegistro & String(59, " ")
    sRegistro = sRegistro & sCount
    
    Open sDiretorio For Output As #1
    
    Print #1, sRegistro
        
    sLinha = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT Linha, Valor FROM DFC_Linha WHERE FilialEmpresa = ? AND Ano = ? ORDER BY Linha", _
        sLinha, dValor, iFilialEmpresa, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 213843
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213844
    
    Do While lErro = AD_SQL_SUCESSO
    
        sRegistro = "2"
        sRegistro = sRegistro & sLinha
        sRegistro = sRegistro & FormataCpoNum(Fix(dValor), 15)
    
        Print #1, sRegistro
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213845
    Loop
    
    Close #1

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213846
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_SPEDFISCAL_SUCESSO", sDiretorio)
    
    DFC_Gera = SUCESSO
    
    Exit Function
    
Erro_DFC_Gera:

    DFC_Gera = gErr
    
    Select Case gErr
    
        Case 213839
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213840
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 213841 To 213845
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LIVREGES", gErr)
            
        Case 213846
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213847)
     
    End Select
    
    Close #1
     
    Call Transacao_Rollback

End Function


