VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATFormata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Function UM_Conversao(ByVal iClasseUM As Integer, ByVal sUnidade1 As String, ByVal sUnidade2 As String, dFator As Double) As Long
'Calcula o fator de conversão entre 2 Unidades de Medida
'De sUnidade1 para sUnidade2

'??? se for a mesma unidade retornar 1,
'??? se uma das unidades for a base basta uma leitura

Dim lErro As Long
Dim objUnidade1 As New ClassUnidadeDeMedida
Dim objUnidade2 As New ClassUnidadeDeMedida

On Error GoTo Erro_UM_Conversao

    'Se forem iguais retorna 1
    If sUnidade1 = sUnidade2 Then
    
        dFator = 1
         
    Else 'Se forem diferentes.....
        
        'Preenche chave de objUnidade1
        objUnidade1.iClasse = iClasseUM
        objUnidade1.sSigla = sUnidade1
    
        'Lê objUnidade1
        lErro = CF("UM_Le", objUnidade1)
        If lErro <> SUCESSO And lErro <> 23775 Then Error 23770
        
        If lErro = 23775 Then Error 55992
        
        'Se Editavel de objUnidade1 for UM_EDITAVEL, atualiza objUnidade1 no BD para Editavel=UM_NAO_EDITAVEL
        If objUnidade1.iEditavel = UM_EDITAVEL Then
                        
            'Chama UM_Grava_NaoEditavel(objUnidade1)
            lErro = UM_Grava_NaoEditavel(objUnidade1)
            If lErro <> SUCESSO Then Error 33558
            
        End If
        
        'Preenche chave de objUnidade2
        objUnidade2.iClasse = iClasseUM
        objUnidade2.sSigla = sUnidade2
    
        'Lê objUnidade2
        lErro = CF("UM_Le", objUnidade2)
        If lErro <> SUCESSO And lErro <> 23775 Then Error 23771
    
        If lErro = 23775 Then Error 55993
    
        'Se Editavel de objUnidade2 for UM_EDITAVEL, atualiza objUnidade2 no BD para Editavel=UM_NAO_EDITAVEL
        If objUnidade2.iEditavel = UM_EDITAVEL Then
        
            'Chama UM_Grava_NaoEditavel(objUnidade2)
            lErro = UM_Grava_NaoEditavel(objUnidade2)
            If lErro <> SUCESSO Then Error 33559
            
        End If
        
        'Calcula dFator
        dFator = (objUnidade1.dQuantidade) / (objUnidade2.dQuantidade)

    End If
    
    UM_Conversao = SUCESSO

    Exit Function

Erro_UM_Conversao:

    UM_Conversao = Err

    Select Case Err

        Case 23770, 23771, 33558, 33559

        Case 55992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidade1.iClasse, objUnidade1.sSigla)

        Case 55993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidade2.iClasse, objUnidade2.sSigla)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151671)

    End Select

    Exit Function

End Function

Function UM_Grava_NaoEditavel(ByVal objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Atualiza o campo Editável da tabela UnidadesDeMedida
'ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iEditavel As Integer
Dim lTransacao As Long
Dim lConexao As Long

On Error GoTo Erro_UM_Grava_NaoEditavel

    lErro = Abre_Conexao(lConexao)
    If lErro <> SUCESSO Then Error 33560

    'Abertura de transação
    lTransacao = Transacao_AbrirExt(lConexao)
    If lTransacao = 0 Then Error 33560

    'Inicializa comandos
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 33560

    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 33561

    'Pesquisa no BD a Unidade de Medida em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Editavel FROM UnidadesdeMedida WHERE Classe=? AND Sigla=?", 0, iEditavel, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 33563

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33564

    If lErro = AD_SQL_SEM_DADOS Then Error 33565
    
    If iEditavel = UM_EDITAVEL Then
             
        'Alterar o campo editável p/ não editável
        lErro = Comando_ExecutarPos(lComando1, "UPDATE UnidadesDeMedida SET Editavel = ?", lComando, UM_NAO_EDITAVEL)
        If lErro <> AD_SQL_SUCESSO Then Error 33566
                
    End If
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacao)
    If lErro <> AD_SQL_SUCESSO Then Error 33560

    Call Conexao_FecharExt(lConexao)

    UM_Grava_NaoEditavel = SUCESSO

    Exit Function

Erro_UM_Grava_NaoEditavel:

    UM_Grava_NaoEditavel = Err

    Select Case Err

        Case 33560, 33561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 33563, 33564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 33565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 33566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_UNIDADESDEMEDIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151672)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Call Transacao_RollbackExt(lTransacao)
    
    Call Conexao_FecharExt(lConexao)
    
    Exit Function
        
End Function



Sub Converte_NaturezaProduto(ByVal iNatureza As Integer, sNatureza As String)
'converte a natureza do produto do valor inteiro para o seu correspondente textual (materia prima/produto intermadiario, ...)

    Select Case iNatureza

        Case NATUREZA_PROD_MATERIA_PRIMA
            sNatureza = "Matéria Prima"
        
        Case NATUREZA_PROD_PRODUTO_INTERMEDIARIO
            sNatureza = "Produto Intermediário"
            
        Case NATUREZA_PROD_EMBALAGENS
            sNatureza = "Embalagens"
            
        Case NATUREZA_PROD_PRODUTO_ACABADO
            sNatureza = "Produto Acabado"
            
        Case NATUREZA_PROD_PRODUTO_REVENDA
            sNatureza = "Produto p/Revenda"
            
        Case NATUREZA_PROD_PRODUTO_MANUTENCAO
            sNatureza = "Produto p/Manutenção"
            
        Case NATUREZA_PROD_OUTROS
            sNatureza = "Outros"

        Case NATUREZA_PROD_SERVICO
            sNatureza = "Outros"
     
        Case NATUREZA_PROD_SERVICO
            sNatureza = "Serviços"
        
        Case NATUREZA_PROD_PRODUTO_EM_PROCESSO
            sNatureza = "Produtos em Processo"
        
        Case NATUREZA_PROD_SUBPRODUTO
            sNatureza = "Subproduto"
        
        Case NATUREZA_PROD_MATERIAL_DE_USO_E_CONSUMO
            sNatureza = "Mat. Uso e Consumo"
        
        Case NATUREZA_PROD_ATIVO_IMOBILIZADO
            sNatureza = "Ativo Imobilizado"
        
        Case NATUREZA_PROD_OUTROS_INSUMOS
            sNatureza = "Outros Insumos"

    End Select
    
End Sub

'Está Funcao ja Esta em classMatFormata
Function UMEstoque_Conversao(ByVal sProduto As String, ByVal sUM As String, ByVal dQuantidade As Double, dQuantidadeConvertida As Double) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_UMEstoque_Conversao

    'Formata o produto
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 64197

    objProduto.sCodigo = sProdutoFormatado
                        
    'Lê o produto para pegar a UM_Estoque
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 64191
                    
    If lErro = 28030 Then Error 64192
    
    'Lê o fator de conversão da UM_Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 64193
                    
    'Converte a UM da quantidade
    dQuantidadeConvertida = dQuantidade * dFator

    UMEstoque_Conversao = SUCESSO
        
    Exit Function
    
Erro_UMEstoque_Conversao:
    
    UMEstoque_Conversao = Err
    
    Select Case Err
            
        Case 64191, 64193, 64197 'Tratados nas rotinas chamadas
        
        Case 64192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151673)

    End Select

    Exit Function
        
End Function

Function MovEstoqueLista_TrataParametros(ByVal objTela As Object, colSelecao As Collection)

    If colSelecao.Item(1) = MOV_EST_REQ_PRODUCAO Then
        objTela.Caption = "Lista de Requisições para Produção"
    End If

    If colSelecao.Item(1) = MOV_EST_PRODUCAO Then
        objTela.Caption = "Lista de Entradas de Material Produzido no Estoque"
    End If

    If colSelecao.Item(1) = MOV_EST_CONSUMO Then
        objTela.Caption = "Lista de Requisições para Consumo"
    End If

    If colSelecao Is Nothing Then Set colSelecao = New Collection

    colSelecao.Add giFilialEmpresa

End Function

Public Function UM_Conversao_Rel(ByVal iClasseUM As Integer, ByVal sUnidade1 As String, ByVal sUnidade2 As String, dFator As Double) As Long
'Calcula o fator de conversão entre 2 Unidades de Medida
'De sUnidade1 para sUnidade2
'ATENÇÃO => Esta função não exibe msg de Erro !!!

'??? se for a mesma unidade retornar 1,
'??? se uma das unidades for a base basta uma leitura

Dim lErro As Long
Dim objUnidade1 As New ClassUnidadeDeMedida
Dim objUnidade2 As New ClassUnidadeDeMedida

On Error GoTo Erro_UM_Conversao_Rel

    'Se forem iguais retorna 1
    If sUnidade1 = sUnidade2 Then
    
        dFator = 1
         
    Else 'Se forem diferentes.....
        
        'Preenche chave de objUnidade1
        objUnidade1.iClasse = iClasseUM
        objUnidade1.sSigla = sUnidade1
    
        'Lê objUnidade1
        lErro = CF("UM_Le", objUnidade1)
        If lErro <> SUCESSO And lErro <> 22537 Then Error 23770
        
        If lErro = 22537 Then Error 55992
        
        'Se Editavel de objUnidade1 for UM_EDITAVEL, atualiza objUnidade1 no BD para Editavel=UM_NAO_EDITAVEL
        If objUnidade1.iEditavel = UM_EDITAVEL Then
                        
            'Chama UM_Grava_NaoEditavel(objUnidade1)
            lErro = UM_Grava_NaoEditavel(objUnidade1)
            If lErro <> SUCESSO Then Error 33558
            
        End If
        
        'Preenche chave de objUnidade2
        objUnidade2.iClasse = iClasseUM
        objUnidade2.sSigla = sUnidade2
    
        'Lê objUnidade2
        lErro = CF("UM_Le", objUnidade2)
        If lErro <> SUCESSO And lErro <> 22537 Then Error 23771
    
        If lErro = 22537 Then Error 55993
    
        'Se Editavel de objUnidade2 for UM_EDITAVEL, atualiza objUnidade2 no BD para Editavel=UM_NAO_EDITAVEL
        If objUnidade2.iEditavel = UM_EDITAVEL Then
        
            'Chama UM_Grava_NaoEditavel(objUnidade2)
            lErro = UM_Grava_NaoEditavel(objUnidade2)
            If lErro <> SUCESSO Then Error 33559
            
        End If
        
        'Calcula dFator
        dFator = (objUnidade1.dQuantidade) / (objUnidade2.dQuantidade)

    End If
    
    UM_Conversao_Rel = SUCESSO

    Exit Function

Erro_UM_Conversao_Rel:

    UM_Conversao_Rel = Err

    Select Case Err

        Case 23770, 23771, 33558, 33559, 55992, 55993

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151674)

    End Select

End Function

Private Function Abre_Conexao(lConexao As Long) As Long

Dim lErro As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim sStringConexao As String

On Error GoTo Erro_Abre_Conexao

    lErro = CF("Empresa_Le_StringConexao", sStringConexao)
    If lErro <> SUCESSO Then gError 182723

    sDSN = sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexao = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexao = 0 Then gError 182724
    
    Abre_Conexao = SUCESSO
    
    Exit Function

Erro_Abre_Conexao:

    Abre_Conexao = gErr

    Select Case gErr
    
        Case 182723, 182724, 182725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182726)

    End Select

    Exit Function

End Function
