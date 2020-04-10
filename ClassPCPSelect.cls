VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPCPSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Public Function Maquinas_Le(ByVal objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tMaquinas As typeMaquinas

On Error GoTo Erro_Maquinas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134137

    'Alocação de espaço no buffer
    tMaquinas.sNomeReduzido = String(STRING_MAQUINA_NOMEREDUZIDO, 0)
    tMaquinas.sDescricao = String(STRING_MAQUINA_DESCRICAO, 0)
    tMaquinas.sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaMaquinas
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, NumIntDoc, NomeReduzido, Descricao, " & _
                "TempoMovimentacao, TempoPreparacao, TempoDescarga, Recurso, CustoHora, Produto, Peso, Largura, Comprimento, Espessura FROM Maquinas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                tMaquinas.iCodigo, tMaquinas.iFilialEmpresa, tMaquinas.lNumIntDoc, tMaquinas.sNomeReduzido, _
                tMaquinas.sDescricao, tMaquinas.dTempoMovimentacao, tMaquinas.dTempoPreparacao, tMaquinas.dTempoDescarga, _
                tMaquinas.iRecurso, tMaquinas.dCustoHora, tMaquinas.sProduto, tMaquinas.dPeso, tMaquinas.dLargura, tMaquinas.dComprimento, tMaquinas.dEspessura, objMaquinas.iCodigo, objMaquinas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 134138

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134139

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 103090

    objMaquinas.iCodigo = tMaquinas.iCodigo
    objMaquinas.iFilialEmpresa = tMaquinas.iFilialEmpresa
    objMaquinas.lNumIntDoc = tMaquinas.lNumIntDoc
    objMaquinas.sNomeReduzido = tMaquinas.sNomeReduzido
    objMaquinas.sDescricao = tMaquinas.sDescricao
    objMaquinas.dTempoMovimentacao = tMaquinas.dTempoMovimentacao
    objMaquinas.dTempoPreparacao = tMaquinas.dTempoPreparacao
    objMaquinas.dTempoDescarga = tMaquinas.dTempoDescarga
    objMaquinas.iRecurso = tMaquinas.iRecurso
    objMaquinas.dCustoHora = tMaquinas.dCustoHora
    objMaquinas.sProduto = tMaquinas.sProduto

    objMaquinas.dPeso = tMaquinas.dPeso
    objMaquinas.dLargura = tMaquinas.dLargura
    objMaquinas.dComprimento = tMaquinas.dComprimento
    objMaquinas.dEspessura = tMaquinas.dEspessura

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Maquinas_Le = SUCESSO

    Exit Function

Erro_Maquinas_Le:

    Maquinas_Le = gErr

    Select Case gErr

        Case 134137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134138, 134139
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

        Case 103090 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152405)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Maquinas_Le_NomeReduzido(ByVal objMaquina As ClassMaquinas) As Long
'Faz a leitura de uma máquina na tabela de maquinas através do NomeReduzido

Dim lErro As Long
Dim lComando As Long
Dim tMaquina As typeMaquinas

On Error GoTo Erro_Maquinas_Le_NomeReduzido

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103099

    tMaquina.sDescricao = String(STRING_MAQUINA_DESCRICAO, 0)
    tMaquina.sNomeReduzido = String(STRING_MAQUINA_NOMEREDUZIDO, 0)
    tMaquina.sProduto = String(STRING_PRODUTO, 0)

    'Tenta selecionar uma máquina através do código e filialempresa
    lErro = Comando_Executar(lComando, "SELECT NomeReduzido, Codigo, FilialEmpresa, NumIntDoc, Descricao, TempoMovimentacao, TempoPreparacao, TempoDescarga, Recurso, CustoHora, Produto, Peso, Largura, Comprimento, Espessura FROM Maquinas WHERE NomeReduzido = ?", _
    tMaquina.sNomeReduzido, tMaquina.iCodigo, tMaquina.iFilialEmpresa, tMaquina.lNumIntDoc, tMaquina.sDescricao, tMaquina.dTempoMovimentacao, tMaquina.dTempoPreparacao, tMaquina.dTempoDescarga, _
    tMaquina.iRecurso, tMaquina.dCustoHora, tMaquina.sProduto, tMaquina.dPeso, tMaquina.dLargura, tMaquina.dComprimento, tMaquina.dEspessura, objMaquina.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 103101

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103102

    'se a máquina não foi encontrada
    If lErro = AD_SQL_SEM_DADOS Then gError 103100

    objMaquina.iCodigo = tMaquina.iCodigo
    objMaquina.iFilialEmpresa = tMaquina.iFilialEmpresa
    objMaquina.lNumIntDoc = tMaquina.lNumIntDoc
    objMaquina.sDescricao = tMaquina.sDescricao
    objMaquina.sNomeReduzido = tMaquina.sNomeReduzido
    objMaquina.dTempoMovimentacao = tMaquina.dTempoMovimentacao
    objMaquina.dTempoPreparacao = tMaquina.dTempoPreparacao
    objMaquina.dTempoDescarga = tMaquina.dTempoDescarga
    objMaquina.iRecurso = tMaquina.iRecurso
    objMaquina.dCustoHora = tMaquina.dCustoHora
    objMaquina.sProduto = tMaquina.sProduto
    
    objMaquina.dPeso = tMaquina.dPeso
    objMaquina.dLargura = tMaquina.dLargura
    objMaquina.dComprimento = tMaquina.dComprimento
    objMaquina.dEspessura = tMaquina.dEspessura
    
    Call Comando_Fechar(lComando)

    Maquinas_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Maquinas_Le_NomeReduzido:

    Maquinas_Le_NomeReduzido = gErr

    Select Case gErr

            Case 103100 'tratado na rotina chamadora

            Case 103101, 103102
                Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

            Case 103099
                Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152406)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TP_Maquina_Le(ByVal objMaquinaControle As Object, ByVal objMaquina As ClassMaquinas, Optional iCria As Integer = 1) As Long
'Lê o Maquina com Código ou NomeRed em objMaquinaControle.Text
'Devolve em objMaquina. Coloca NomeReduzido em .Text

Dim sMaquina As String
Dim Maquina As Object
Dim eTipoMaquina As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TP_Maquina_Le

    Set Maquina = objMaquinaControle
    sMaquina = Trim(Maquina.Text)
    eTipoMaquina = Tipo_Maquina(sMaquina)

    Select Case eTipoMaquina

    Case TIPO_STRING

        objMaquina.sNomeReduzido = sMaquina
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 103085
        If lErro <> SUCESSO Then gError 103086

        Maquina.Text = objMaquina.sNomeReduzido

    Case TIPO_CODIGO

        objMaquina.iCodigo = StrParaInt(sMaquina)
        objMaquina.iFilialEmpresa = giFilialEmpresa
        lErro = CF("Maquinas_Le", objMaquina)
        If lErro <> SUCESSO And lErro <> 103090 Then gError 103094
        If lErro <> SUCESSO Then gError 103095

        Maquina.Text = objMaquina.sNomeReduzido

    Case TIPO_OVERFLOW

        gError 103096

    Case TIPO_DECIMAL

        gError 103097

    Case TIPO_NAO_POSITIVO

        gError 103098

    End Select

    TP_Maquina_Le = SUCESSO

    Exit Function

Erro_TP_Maquina_Le:

    TP_Maquina_Le = gErr

    Select Case gErr

        Case 103085, 103094 'Tratados nas rotinas chamadas

        Case 103086  'Maquina com NomeReduzido não cadastrado

            If iCria = 1 Then

                'Envia aviso que Maquina não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_MAQUINA2", objMaquina.sNomeReduzido)

                If vbMsgRes = vbYes Then
                    'Chama tela de Maquinas
                    lErro = Chama_Tela("Maquinas", objMaquina)
                End If
            End If

        Case 103095  'Maquina com Código não cadastrado

            If iCria = 1 Then

                'Envia aviso que Maquina não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_MAQUINA1", objMaquina.iCodigo)

                If vbMsgRes = vbYes Then
                    'Chama tela de Maquinas
                    lErro = Chama_Tela("Maquinas", objMaquina)
                End If
            End If

        Case 103096
            Call Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", gErr, sMaquina)

        Case 103097
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sMaquina)

        Case 103098
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sMaquina)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152407)

    End Select

End Function

Private Function Tipo_Maquina(ByVal sMaquina As String) As enumTipo

    If Len(Trim(sMaquina)) = 0 Then
        Tipo_Maquina = TIPO_VAZIO
    ElseIf Not IsNumeric(sMaquina) Then
        Tipo_Maquina = TIPO_STRING
    ElseIf Int(CDbl(sMaquina)) <> CDbl(sMaquina) Then
        Tipo_Maquina = TIPO_DECIMAL
    ElseIf CDbl(sMaquina) <= 0 Then
        Tipo_Maquina = TIPO_NAO_POSITIVO
    ElseIf CDbl(sMaquina) > NUM_MAX_MAQUINAS Then
        Tipo_Maquina = TIPO_OVERFLOW
    Else
        Tipo_Maquina = TIPO_CODIGO
    End If

End Function

Public Function Maquinas_Le_NumIntDoc(ByVal objMaquina As ClassMaquinas) As Long
'Faz a leitura de uma máquina na tabela de maquinas através do NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim tMaquina As typeMaquinas

On Error GoTo Erro_Maquinas_Le_NumIntDoc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106350

    tMaquina.sDescricao = String(STRING_MAQUINA_DESCRICAO, 0)
    tMaquina.sNomeReduzido = String(STRING_MAQUINA_NOMEREDUZIDO, 0)
    tMaquina.sProduto = String(STRING_PRODUTO, 0)

    'Tenta selecionar uma máquina através do código e filialempresa
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Descricao, NomeReduzido, TempoMovimentacao, TempoPreparacao, TempoDescarga, Recurso, CustoHora, Produto, Peso, Largura, Comprimento, Espessura FROM Maquinas WHERE NumIntDoc = ?", _
    tMaquina.iCodigo, tMaquina.iFilialEmpresa, tMaquina.sDescricao, tMaquina.sNomeReduzido, tMaquina.dTempoMovimentacao, tMaquina.dTempoPreparacao, tMaquina.dTempoDescarga, _
    tMaquina.iRecurso, tMaquina.dCustoHora, tMaquina.sProduto, tMaquina.dPeso, tMaquina.dLargura, tMaquina.dComprimento, tMaquina.dEspessura, objMaquina.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 106351

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106352

    'se a máquina não foi encontrada
    If lErro = AD_SQL_SEM_DADOS Then gError 106353

    objMaquina.iCodigo = tMaquina.iCodigo
    objMaquina.iFilialEmpresa = tMaquina.iFilialEmpresa
    objMaquina.sDescricao = tMaquina.sDescricao
    objMaquina.sNomeReduzido = tMaquina.sNomeReduzido
    objMaquina.dTempoMovimentacao = tMaquina.dTempoMovimentacao
    objMaquina.dTempoPreparacao = tMaquina.dTempoPreparacao
    objMaquina.dTempoDescarga = tMaquina.dTempoDescarga
    objMaquina.iRecurso = tMaquina.iRecurso
    objMaquina.dCustoHora = tMaquina.dCustoHora
    objMaquina.sProduto = tMaquina.sProduto
    
    objMaquina.dPeso = tMaquina.dPeso
    objMaquina.dLargura = tMaquina.dLargura
    objMaquina.dComprimento = tMaquina.dComprimento
    objMaquina.dEspessura = tMaquina.dEspessura
    
    Call Comando_Fechar(lComando)

    Maquinas_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_Maquinas_Le_NumIntDoc:

    Maquinas_Le_NumIntDoc = gErr

    Select Case gErr

            Case 106353 'tratado na rotina chamadora

            Case 106351, 106352
                Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

            Case 106350
                Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152408)

    End Select

    Call Comando_Fechar(lComando)

End Function

Public Function TP_Maquina_Le2(ByVal objMaquinaControle As Object, ByVal objMaquina As ClassMaquinas, Optional iCria As Integer = 1) As Long
'Lê o Maquina com Código, NomeRed ou Codigo & SEPARADOR & NoemReduzido em objMaquinaControle.Text
'Devolve em objMaquina. Coloca Código - NomeReduzido em .Text

Dim sMaquina As String
Dim Maquina As Object
Dim eTipoMaquina As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TP_Maquina_Le2

    Set Maquina = objMaquinaControle

    'Tenta extrair o código da máquina
    iCodigo = Codigo_Extrai(Trim(Maquina.Text))

    If iCodigo = 0 Then

            objMaquina.sNomeReduzido = Trim(Maquina.Text)
            lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
            If lErro <> SUCESSO And lErro <> 103100 Then gError 106450
            If lErro <> SUCESSO Then gError 106451

    Else

            objMaquina.iCodigo = iCodigo
            objMaquina.iFilialEmpresa = giFilialEmpresa
            lErro = CF("Maquinas_Le", objMaquina)
            If lErro <> SUCESSO And lErro <> 103090 Then gError 106452
            If lErro <> SUCESSO Then gError 106453

    End If

    Maquina.Text = objMaquina.iCodigo & SEPARADOR & objMaquina.sNomeReduzido

    TP_Maquina_Le2 = SUCESSO

    Exit Function

Erro_TP_Maquina_Le2:

    TP_Maquina_Le2 = gErr

    Select Case gErr

        Case 106450, 106452 'Tratados nas rotinas chamadas

        Case 106451  'Maquina com NomeReduzido não cadastrado

        Case 106453  'Maquina com Código não cadastrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152409)

    End Select

End Function

Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)

    Nome_Extrai = sString

End Function

Public Function Verifica_Produto_Agua(ByVal sProduto, bEhAgua As Boolean) As Long
'Verifica se o Produto passado é água

'*****************************************************************************
'Este Resultado abaixo foi a base para comparação do produto com o produto AGUA
'Favor não remover o trecho comentado abaixo

            'Codigo               Tipo   Descricao                                                NomeReduzido
            '-------------------- ------ -------------------------------------------------- --------------------
            'H000010             12     AGUA                                                       AGUA
            'H000012             12     AGUA DESMINERALIZADA                        AGUA DESMINERAL
            
'*****************************************************************************
    
    Select Case sProduto
    
        Case "H000010", "H000012"
            bEhAgua = True
            
        Case Else
            bEhAgua = False
            
    End Select

End Function

Public Function Obter_AdicaoMP_ProdOP(ByVal sProduto As String, ByVal sCodOP As String, colItensOP As Collection) As Long
'Monta uma colecao com os insumos correspodentes ao Produto/OP passados.

Dim lErro As Long
Dim iIndice As Integer
Dim sVersao As String
Dim lComando As Long
Dim dFatorConv As Double
Dim dQuantidadeOP As Double, sUMDest As String
Dim objProduto As New ClassProduto, sSiglaUM As String
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos
Dim objProdutoKitProdutosTEMP As New ClassProdutoKitProdutos

On Error GoTo Erro_Obter_AdicaoMP_ProdOP
    
    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108600
    
    'Inicialização da string
    sVersao = String(STRING_KIT_VERSAO, 0)
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    
    'Prepara a Execucao do Comando SQL
    lErro = Comando_Executar(lComando, "SELECT Versao, Quantidade, SiglaUM FROM ItensOrdemProducaoBaixadas WHERE Codigo = ? AND Produto = ? AND FilialEmpresa = ?", sVersao, dQuantidadeOP, sSiglaUM, sCodOP, sProduto, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 108631
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108602
    
    'Se não Encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108632
    
    'Le o Produto
    objProduto.sCodigo = sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 108604
    
    'Se nao encontrou => Erro
    If lErro = 28030 Then gError 108605
    
    sUMDest = "KG"
    
    'Transforma produto em kg
    'Faz a Conversao de unidades de medida
    lErro = CF("UM_Conversao_Rel", objProduto.iClasseUM, sSiglaUM, sUMDest, dFatorConv)
    If lErro <> SUCESSO Then gError 108609
        
    objProdutoKitProdutos.sProdutoRaiz = sProduto
    objProdutoKitProdutos.sVersao = sVersao
    objProdutoKitProdutos.dQuantidade = dQuantidadeOP * dFatorConv
    objProdutoKitProdutos.iClasseUM = objProduto.iClasseUM
    objProdutoKitProdutos.sUnidadeMed = sUMDest
    
    'Obtem uma Colecao com os insumos do kit já convertidos para sUMDest proporcionais à qtde à ser produzida
    'nao inclui embalagem
    lErro = OrdemProducao_Le_Col_Empenho3(objProdutoKitProdutos, colItensOP, sUMDest)
    If lErro <> SUCESSO And lErro <> 106393 Then gError 108606
    
    'Se nao Encontrou => Erro
    If lErro = 106393 Then gError 108633
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Obter_AdicaoMP_ProdOP = SUCESSO
    
    Exit Function

Erro_Obter_AdicaoMP_ProdOP:

    Obter_AdicaoMP_ProdOP = gErr

    Select Case gErr
    
        Case 108600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108631, 108602, 108604, 108606, 108633, 108608, 108609
        
        Case 108632
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", gErr)
            
        Case 108605
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152410)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Public Function Obter_AdicaoMP_MatReq(ByVal sProduto As String, ByVal sCodOP As String, colItensMovEst As Collection) As Long
'Monta uma colecao com os itens de Material Requisitado

Dim lErro As Long, iTipoMov As Integer, dFator As Double
Dim iIndice As Integer, iNatureza As Integer, iClasseUM As Integer
Dim lComando As Long
Dim bAchou As Boolean
Dim dFatorConv As Double
Dim dQuantidadeOP As Double
Dim objProduto As New ClassProduto
Dim tProdutoKitProdutos As typeProdutoKitProdutos
Dim objProdutoKitProdutos As ClassProdutoKitProdutos
Dim objProdutoKitProdutosTEMP As New ClassProdutoKitProdutos

On Error GoTo Erro_Obter_AdicaoMP_MatReq
    
    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108600
    
    'Instancia obj
    Set objProdutoKitProdutos = New ClassProdutoKitProdutos
            
    'Inicializar as strings
    tProdutoKitProdutos.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKitProdutos.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    
    'Prepara a Execucao do Comando SQL
    lErro = Comando_Executar(lComando, "SELECT Natureza, ClasseUM, TipoMov, Produto, Quantidade, SiglaUM FROM MatReqProdOP WHERE ProdutoRaiz=? AND FilialEmpresa=? AND Codigo=?", _
                                                iNatureza, iClasseUM, iTipoMov, tProdutoKitProdutos.sProduto, tProdutoKitProdutos.dQuantidade, tProdutoKitProdutos.sSiglaUMEstoque, sProduto, giFilialEmpresa, sCodOP)
    If lErro <> AD_SQL_SUCESSO Then gError 108601
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108602
    
    'Se não Encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108603
        
    'Enquanto tiver dados => Preenche a colecao
    While lErro = AD_SQL_SUCESSO
        
        If iNatureza <> NATUREZA_PROD_EMBALAGENS Then
        
            Select Case iTipoMov
            
                Case MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3
                    dFator = 1
    
                Case MOV_EST_DEV_MATERIAL_PRODUCAO
                    dFator = -1
                
                Case Else
                    dFator = 0
                
            End Select
                
            If dFator <> 0 Then
            
                objProdutoKitProdutos.sProdutoRaiz = sProduto
                objProdutoKitProdutos.sProduto = tProdutoKitProdutos.sProduto
                objProdutoKitProdutos.sSiglaUMEstoque = tProdutoKitProdutos.sSiglaUMEstoque
                objProdutoKitProdutos.dQuantidade = tProdutoKitProdutos.dQuantidade * dFator
                
                bAchou = False
                
                'Percorre na colecao para nao haver duplicidade
                For Each objProdutoKitProdutosTEMP In colItensMovEst
                
                    'Verifica se já existe na colecao
                    If objProdutoKitProdutosTEMP.sProduto = objProdutoKitProdutos.sProduto Then
                    
                        If tProdutoKitProdutos.sSiglaUMEstoque <> "KG" Then
                        
                            'Transforma produto em kg
                            'Faz a Conversao de unidades de medida
                            lErro = CF("UM_Conversao_Trans", iClasseUM, tProdutoKitProdutos.sSiglaUMEstoque, "KG", dFatorConv)
                            If lErro <> SUCESSO Then gError 108609
                            
                            objProdutoKitProdutos.dQuantidade = objProdutoKitProdutos.dQuantidade * dFatorConv
                            
                        End If
                        
                        'Adiciona à quantidade já existente
                        objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade + (objProdutoKitProdutos.dQuantidade)
                        
                        bAchou = True
                        
                        Exit For
                            
                    End If
                
                Next
                    
                'Se nao encontrou na colecao => Adiciona
                If bAchou = False Then
                    
                    colItensMovEst.Add objProdutoKitProdutos
                    
                    'Instancia outro obj
                    Set objProdutoKitProdutos = New ClassProdutoKitProdutos
                
                End If
            
            End If
        
        End If
        
        'Busca o Próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108612
    
    Wend
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    Obter_AdicaoMP_MatReq = SUCESSO
    
    Exit Function

Erro_Obter_AdicaoMP_MatReq:

    Obter_AdicaoMP_MatReq = gErr

    Select Case gErr
    
        Case 108600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108601, 108602, 108612
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 108609
        
        Case 108603, 108604
            
        Case 108605
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152411)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function OrdemProducao_Le_Col_Empenho3(ByVal objProdutoKitProdutos As ClassProdutoKitProdutos, ByVal colItensOP As Collection, ByVal sUMDest As String) As Long
'Le a partir do Produto/Versao todos os componentes do Kit e os retorna na colecao (não há duplicidade na colecao)
'Obtem uma Colecao com os insumos do kit já convertidos para sUMDest proporcionais à qtde à ser produzida
'nao inclui embalagem

Dim lErro As Long
Dim dFator As Double, dFatorRaiz As Double
Dim lComando As Long
Dim bAchou As Boolean
Dim dQuantidade As Double
Dim dQuantidadePerda As Double
Dim objProduto As New ClassProduto
Dim objProdutoKit As New ClassProdutoKit
Dim tProdutoKitProdutos As typeProdutoKitProdutos
Dim objProdutoKitProdutosTEMP As New ClassProdutoKitProdutos
Dim objProdutoKitProdutosTEMP2 As New ClassProdutoKitProdutos

On Error GoTo Erro_OrdemProducao_Le_Col_Empenho3

    'Aertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106390
    
    'Faz a leitura de dados do produto raiz no kit
    objProdutoKit.sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
    objProdutoKit.sVersao = objProdutoKitProdutos.sVersao
    
    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKit)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 108585
    
    'Se nao encontrou => Erro
    If lErro = 34875 Then gError 108586
    
    'descobre o fator de conversao da UM raiz p/UM dest
    lErro = CF("UM_Conversao_Rel", objProdutoKitProdutos.iClasseUM, objProdutoKit.sUnidadeMed, sUMDest, dFatorRaiz)
    If lErro <> SUCESSO Then gError 55248
    
    'Inicializa as strings do Type
    With tProdutoKitProdutos
    
        .sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)
        .sProduto = String(STRING_PRODUTO, 0)
    
        'para obter apenas os nós do nivel 1
        lErro = Comando_Executar(lComando, "SELECT ProdutoKit.Produto, Quantidade, ProdutoKit.UnidadeMed, Composicao, ControleEstoque, SiglaUMEstoque, ClasseUM FROM ProdutoKit, Produtos WHERE ProdutoRaiz = ? AND Versao = ? AND ProdutoKit.Produto = Produtos.Codigo AND ProdutoKit.Nivel = 1", .sProduto, .dQuantidade, .sUnidadeMed, .iComposicao, .iControleEstoque, .sSiglaUMEstoque, .iClasseUM, objProdutoKitProdutos.sProdutoRaiz, objProdutoKitProdutos.sVersao)
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 106391

    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106392
    
    'Se nao Encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 106393
    
    'Enquanto tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS
        
        bAchou = False
        
        'Transfere os Dados para o obj
        With objProdutoKitProdutosTEMP
        
            .dQuantidade = tProdutoKitProdutos.dQuantidade
            .iClasseUM = tProdutoKitProdutos.iClasseUM
            .iComposicao = tProdutoKitProdutos.iComposicao
            .iControleEstoque = tProdutoKitProdutos.iControleEstoque
            .sProduto = tProdutoKitProdutos.sProduto
            .sSiglaUMEstoque = tProdutoKitProdutos.sSiglaUMEstoque
            .sUnidadeMed = tProdutoKitProdutos.sUnidadeMed
            .sProdutoRaiz = objProdutoKitProdutos.sProdutoRaiz
            .sVersao = objProdutoKitProdutos.sVersao
        
        End With
        
        'Faz a Conversao de unidades de medida
        lErro = CF("UM_Conversao_Rel", objProdutoKitProdutosTEMP.iClasseUM, objProdutoKitProdutosTEMP.sUnidadeMed, sUMDest, dFator)
        If lErro <> SUCESSO Then gError 106394
        
        objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade * dFator
            
        If objProdutoKitProdutosTEMP.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
        
            objProdutoKitProdutosTEMP.dQuantidade = objProdutoKitProdutosTEMP.dQuantidade * (objProdutoKitProdutos.dQuantidade / (objProdutoKit.dQuantidade * dFatorRaiz))
            
        End If
                
        'Percorre na colecao para nao haver duplicidade
        For Each objProdutoKitProdutosTEMP2 In colItensOP
        
            'Verifica se já existe na colecao
            If objProdutoKitProdutosTEMP2.sProduto = objProdutoKitProdutosTEMP.sProduto Then
            
                'Adiciona à quantidade já existente
                objProdutoKitProdutosTEMP2.dQuantidade = objProdutoKitProdutosTEMP2.dQuantidade + objProdutoKitProdutosTEMP.dQuantidade
                bAchou = True
                Exit For
                    
            End If
        
        Next
            
        'Se nao encontrou na colecao => Adiciona
        If bAchou = False Then
            
            'Adiciona na colecao
            colItensOP.Add objProdutoKitProdutosTEMP
            
            'Instancia outro obj
            Set objProdutoKitProdutosTEMP = New ClassProdutoKitProdutos
            
        End If
            
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106395
    
    Loop
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

    OrdemProducao_Le_Col_Empenho3 = SUCESSO
    
    Exit Function
    
Erro_OrdemProducao_Le_Col_Empenho3:

    OrdemProducao_Le_Col_Empenho3 = gErr
    
    Select Case gErr
    
        Case 106394
            dFator = 0
            Resume Next
            
        Case 106390
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 106391, 106392, 106395
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT_PRODUTOS", gErr)
            '??? Erro ao tentar ler as tabelas de Produtos e ProdutoKit.
            
        Case 106393 'Sem Dados
        
        Case 106396, 106397, 106421, 108585
        
        Case 108586
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRODUTO_RAIZ", gErr, objProdutoKit.sProdutoRaiz)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152412)
    
    End Select
    
    'Fechamento do Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ProdutoKit_Le_PrimeiroNivel_Versao(objKit As ClassKit) As Long
'Lê os componentes do Kit da versão recebida, se a versão não for informada Lê a versão padrão.
'Os componentes serão retornados em objKit.colComponentes

Dim lErro As Long
Dim lComando As Long
Dim tProdutoKit As typeProdutoKit
Dim objProdutoKit As ClassProdutoKit
Dim sComando_SQL As String, sProdutoDescricao As String
Dim iClasseUM As Integer

On Error GoTo Erro_ProdutoKit_Le_PrimeiroNivel_Versao

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103197

    'Se nenhuma versão foi informada, pega a versão padrão do Kit
    If Len(Trim(objKit.sVersao)) = 0 Then

        'tenta ler a versão padrão do Kit
        lErro = CF("Kit_Le_Padrao", objKit)
        If lErro <> SUCESSO And lErro <> 106304 Then gError 103202

        'Se não existe versão padrão p/ o Kit -> Erro.
        If lErro <> SUCESSO Then gError 103203

    End If

    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
    sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    '###########################################
    'Inserido por Wagner
    tProdutoKit.sVersaoKitComp = String(STRING_KIT_VERSAO, 0)
    '###########################################

    '###########################################
    'Alterado por Wagner
    sComando_SQL = "SELECT ProdutoKit.PercentualPerda, ProdutoKit.ProdutoRaiz, ProdutoKit.Versao, ProdutoKit.Nivel, ProdutoKit.Seq, ProdutoKit.Produto, Produtos.Descricao, Produtos.ClasseUM, ProdutoKit.SeqPai, ProdutoKit.Quantidade, ProdutoKit.UnidadeMed, ProdutoKit.Composicao, ProdutoKit.VersaoKitComp FROM Produtos, ProdutoKit WHERE Produtos.Codigo = ProdutoKit.Produto AND ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Versao = ? AND (ProdutoKit.Nivel = ? OR ProdutoKit.Nivel = ?) ORDER BY ProdutoKit.Nivel"

    lErro = Comando_Executar(lComando, sComando_SQL, tProdutoKit.dPercentualPerda, tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, sProdutoDescricao, iClasseUM, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, tProdutoKit.sVersaoKitComp, objKit.sProdutoRaiz, objKit.sVersao, KIT_NIVEL_RAIZ, KIT_NIVEL_RAIZ + 1)
    If lErro <> AD_SQL_SUCESSO Then gError 103198
    '###########################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103199

'    If lErro = AD_SQL_SEM_DADOS Then gError 103200

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objProdutoKit = New ClassProdutoKit

        objProdutoKit.dPercentualPerda = tProdutoKit.dPercentualPerda
        objProdutoKit.dQuantidade = tProdutoKit.dQuantidade
        objProdutoKit.iComposicao = tProdutoKit.iComposicao
        objProdutoKit.iNivel = tProdutoKit.iNivel
        objProdutoKit.iSeq = tProdutoKit.iSeq
        objProdutoKit.iSeqPai = tProdutoKit.iSeqPai
        objProdutoKit.sProduto = tProdutoKit.sProduto
        objProdutoKit.sProdutoDesc = sProdutoDescricao
        objProdutoKit.iClasseUM = iClasseUM
        objProdutoKit.sUnidadeMed = tProdutoKit.sUnidadeMed
        objProdutoKit.sProdutoRaiz = tProdutoKit.sProdutoRaiz
        objProdutoKit.sVersao = tProdutoKit.sVersao
        
        '###########################################
        'Inserido por Wagner
        objProdutoKit.sVersaoKitComp = tProdutoKit.sVersaoKitComp
        '###########################################

        objKit.colComponentes.Add objProdutoKit

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103201

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ProdutoKit_Le_PrimeiroNivel_Versao = SUCESSO

    Exit Function

Erro_ProdutoKit_Le_PrimeiroNivel_Versao:

    ProdutoKit_Le_PrimeiroNivel_Versao = gErr

    Select Case gErr

        Case 103197
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103198, 103199, 103201
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 103200 'Tratado na rotina chamadora

        Case 103202

        Case 103203
'            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz, objKit.sVersao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152413)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function EstoqueProduto_ObterQTD_Filial(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal iFilialEmpresa As Integer) As Long
'Obtém a quantidade atual total (proprio, reservado ou nao, e de 3os na empresa) de um produto em uma filial empresa
'e a data de cadastramento de estoque do produto na filial.dtDataInicial
'os dados sao retornados em objEstoqueProduto.dQuantDispNossa e objEstoqueProduto.

Dim lErro As Long
Dim lComando As Long
Dim sProduto As String
Dim dQuantidade As Double
Dim sComando_SQL As String
Dim dtDataInicial As Date

On Error GoTo Erro_EstoqueProduto_ObterQTD_Filial

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103274

    sComando_SQL = "SELECT MIN(DataInicial) AS DataInicial, SUM(QuantDispNossa+QuantReservada+QuantConsig3+QuantBenef3+QuantOutras3+QuantConserto3+QuantDemo3) FROM EstoqueProduto, Almoxarifado WHERE Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo"

    lErro = Comando_Executar(lComando, sComando_SQL, dtDataInicial, dQuantidade, iFilialEmpresa, objEstoqueProduto.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 103275

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103276

    If lErro = AD_SQL_SEM_DADOS Then gError 103277

    objEstoqueProduto.dQuantDispNossa = dQuantidade
    objEstoqueProduto.dtDataInicial = dtDataInicial

    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueProduto_ObterQTD_Filial = SUCESSO

    Exit Function

Erro_EstoqueProduto_ObterQTD_Filial:

    EstoqueProduto_ObterQTD_Filial = gErr

    Select Case gErr

        Case 103274
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103275, 103276
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr)

        Case 103277 'Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152414)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_ProducaoLocal(ByVal sProduto As String, ByVal iFilialEmpresa As Integer) As Long
'retorna SUCESSO se o produto estiver cadastrado como sendo de producao na filial e 103277 caso nao seja

Dim lErro As Long, lComando As Long, iProdNaFilial As Integer, iCompras As Integer

On Error GoTo Erro_Produto_ProducaoLocal

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103274

    'obtem dados p/verificar se é produzido mas nao na iFilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT ProdNaFilial, Compras FROM Produtos, ProdutosFilial WHERE Produtos.Codigo = ProdutosFilial.Produto AND FilialEmpresa = ? AND Produto = ?", _
        iProdNaFilial, iCompras, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106925
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103276
    If lErro <> AD_SQL_SUCESSO Then gError 103277
    
    'se for produzido mas nao na iFilialEmpresa, deve ser descartado
    If iCompras = PRODUTO_PRODUZIVEL And iProdNaFilial = 0 Then gError 103277
    
    Call Comando_Fechar(lComando)
    
    Produto_ProducaoLocal = SUCESSO
     
    Exit Function
    
Erro_Produto_ProducaoLocal:

    Produto_ProducaoLocal = gErr
     
    Select Case gErr
          
        Case 103274
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103276, 106925
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL_PROD", gErr)
        
        Case 103277 'Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152415)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Maquinas_Le_Itens(ByVal objMaquinas As ClassMaquinas) As Long

Dim lErro As Long
Dim lComando As Long
Dim objMaquinasInsumos As ClassMaquinasInsumos
Dim objMaquinaOperadores As ClassMaquinaOperadores
Dim tMaquinasInsumos As typeMaquinasInsumos
Dim tMaquinaOperadores As typeMaquinaOperadores

On Error GoTo Erro_Maquinas_Le_Itens

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134140

    'Lê MaquinasInsumos
    tMaquinasInsumos.sProduto = String(STRING_PRODUTO, 0)
    tMaquinasInsumos.sUMProduto = String(STRING_UM_SIGLA, 0)
    tMaquinasInsumos.sUMTempo = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDocMaq, Produto, Quantidade, UMProduto, UMTempo FROM MaquinasInsumos WHERE NumIntDocMaq = ?", tMaquinasInsumos.lNumIntDocMaq, tMaquinasInsumos.sProduto, tMaquinasInsumos.dQuantidade, tMaquinasInsumos.sUMProduto, tMaquinasInsumos.sUMTempo, objMaquinas.lNumIntDoc)
    If lErro <> SUCESSO Then gError 134141
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134142

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objMaquinasInsumos = New ClassMaquinasInsumos

        objMaquinasInsumos.lNumIntDocMaq = tMaquinasInsumos.lNumIntDocMaq
        objMaquinasInsumos.sProduto = tMaquinasInsumos.sProduto
        objMaquinasInsumos.dQuantidade = tMaquinasInsumos.dQuantidade
        objMaquinasInsumos.sUMProduto = tMaquinasInsumos.sUMProduto
        objMaquinasInsumos.sUMTempo = tMaquinasInsumos.sUMTempo

        objMaquinas.colProdutos.Add objMaquinasInsumos

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134143

    Loop
    
    'Lê MaquinaOperadores
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocMaq, TipoMaoDeObra, Quantidade, PercentualUso FROM MaquinaOperadores WHERE NumIntDocMaq = ?", tMaquinaOperadores.lNumIntDoc, tMaquinaOperadores.lNumIntDocMaq, tMaquinaOperadores.iTipoMaoDeObra, tMaquinaOperadores.iQuantidade, tMaquinaOperadores.dPercentualUso, objMaquinas.lNumIntDoc)
    If lErro <> SUCESSO Then gError 134663
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134664

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objMaquinaOperadores = New ClassMaquinaOperadores

        objMaquinaOperadores.lNumIntDoc = tMaquinaOperadores.lNumIntDoc
        objMaquinaOperadores.lNumIntDocMaq = tMaquinaOperadores.lNumIntDocMaq
        objMaquinaOperadores.iTipoMaoDeObra = tMaquinaOperadores.iTipoMaoDeObra
        objMaquinaOperadores.iQuantidade = tMaquinaOperadores.iQuantidade
        objMaquinaOperadores.dPercentualUso = tMaquinaOperadores.dPercentualUso

        objMaquinas.colTipoOperadores.Add objMaquinaOperadores

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134665

    Loop
        
    Call Comando_Fechar(lComando)
    
    Maquinas_Le_Itens = SUCESSO

    Exit Function


Erro_Maquinas_Le_Itens:

   Maquinas_Le_Itens = gErr

   Select Case gErr
    
        Case 134140
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 134141, 134142, 134143, 134663, 134664, 134665
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152416)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ApontamentoProducao_ObterDadosGrid(ByVal objApontamentoSeleciona As ClassApontamentoSeleciona) As Long

Dim lErro As Long
Dim lComando As Long
Dim sTexto_SQL As String
Dim lNumIntDoc As Long
Dim lNumIntDocPOPai As Long
Dim lNumIntDocPMP As Long
Dim lNumIntDocOper As Long
Dim iNivel As Integer
Dim iSeq As Integer
Dim sCodOPOrigem As String
Dim sProduto As String
Dim sVersao As String
Dim dQuantidade As Double
Dim sUM As String
Dim lNumIntDocCT As Long
Dim dtDataInicio As Date
Dim dHoraInicio As Double
Dim dtDataFim As Date
Dim dHoraFim As Double
Dim iTemApontamento As Integer
Dim objPO As ClassPlanoOperacional
Dim objApontamento As ClassApontamentoProducao

On Error GoTo Erro_ApontamentoProducao_ObterDadosGrid

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137718
    
    Call ApontamentoProducao_PreparaSQL(objApontamentoSeleciona, sTexto_SQL)
    
    'Inicializa buffers
    sCodOPOrigem = String(STRING_OPCODIGO, 0)
    sProduto = String(STRING_PRODUTO, 0)
    sVersao = String(STRING_KIT_VERSAO, 0)
    sUM = String(STRING_UM_SIGLA, 0)

    lErro = ApontamentoProducao_PreparaInt(lComando, lNumIntDoc, lNumIntDocPOPai, lNumIntDocPMP, _
                        lNumIntDocOper, iNivel, iSeq, sCodOPOrigem, sProduto, sVersao, dQuantidade, sUM, _
                        lNumIntDocCT, dtDataInicio, dHoraInicio, dtDataFim, dHoraFim, _
                        objApontamentoSeleciona.dtDataOPFinal, objApontamentoSeleciona.dtDataOPInicial, _
                        objApontamentoSeleciona.lCTFinal, objApontamentoSeleciona.lCTInicial, _
                        objApontamentoSeleciona.sOPFinal, objApontamentoSeleciona.sOPInicial, _
                        objApontamentoSeleciona.sProdutoFinal, objApontamentoSeleciona.sProdutoInicial, _
                        sTexto_SQL)
    If lErro <> SUCESSO Then gError 137719

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137720

    If lErro = AD_SQL_SEM_DADOS Then gError 137773

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objApontamento = New ClassApontamentoProducao

        objApontamento.lNumIntDocPO = lNumIntDoc
        
        'Le ApontamentoProducao gravado anteriormente
        lErro = ApontamentoProducao_Le(objApontamento)
        If lErro <> SUCESSO And lErro <> 137717 Then gError 137721
        
        'Se encontrou... Tem Apontamentos
        iTemApontamento = IIf(lErro = SUCESSO, MARCADO, DESMARCADO)

        'Se ainda não está Concluido ou exibe todos...
        If objApontamento.iConcluido = DESMARCADO Or objApontamentoSeleciona.iTodos = True Then
        
            Set objPO = New ClassPlanoOperacional
        
            objPO.lNumIntDoc = lNumIntDoc
            objPO.lNumIntDocPOPai = lNumIntDocPOPai
            objPO.lNumIntDocPMP = lNumIntDocPMP
            objPO.lNumIntDocOper = lNumIntDocOper
            objPO.iNivel = iNivel
            objPO.iSeq = iSeq
            objPO.sCodOPOrigem = sCodOPOrigem
            objPO.sProduto = sProduto
            objPO.sVersao = sVersao
            objPO.dQuantidade = dQuantidade
            objPO.sUM = sUM
            objPO.lNumIntDocCT = lNumIntDocCT
            objPO.dtDataInicio = dtDataInicio
            objPO.dtDataFim = dtDataFim
            objPO.iTemApontamento = iTemApontamento

            Set objPO.objApontamento = objApontamento
        
            objApontamentoSeleciona.colPlanoOperacional.Add objPO
            
        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137722

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ApontamentoProducao_ObterDadosGrid = SUCESSO
    
    Exit Function
    
Erro_ApontamentoProducao_ObterDadosGrid:

    ApontamentoProducao_ObterDadosGrid = gErr
    
    Select Case gErr
    
        Case 137718
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 137720, 137722
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAQUINAS", gErr)

        Case 137719, 137721
        
        Case 137773
            'Sem dados
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152417)
        
    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Sub ApontamentoProducao_PreparaSQL(ByVal objApontamentoSeleciona As ClassApontamentoSeleciona, sTexto_SQL As String)
'Monta o SELECT que traz os dados necessários para Obter Dados do Grid
 
Dim sSelect As String
Dim sWhere As String
Dim sWhereAux As String
Dim sFrom As String
Dim sOrderBy As String
Dim sSubSelect As String
 
On Error GoTo Erro_ApontamentoProducao_PreparaSQL
        
    sSelect = "SELECT NumIntDoc, NumIntDocPOPai, NumIntDocPMP, PO.NumIntDocOper, " & _
                        "Nivel, Seq, CodOPOrigem, Produto, Versao, Quantidade, " & _
                        "UM, NumIntDocCT, DataInicio, HoraInicio, DataFim, HoraFim "
 
    sFrom = "FROM PlanoOperacional AS PO " & _
            "INNER JOIN OrdensDeProducao AS OP ON OP.Codigo = PO.CodOPOrigem "
    
    sSubSelect = "(SELECT PMPItens.NumIntDoc " & _
                    "FROM PlanoMestreProducaoItens AS PMPItens, " & _
                    "PlanoMestreProducaoUltimo AS PMP " & _
                    "WHERE PMP.Codgeracao = PMPItens.CodGeracao AND " & _
                    "PMPItens.NumIntDoc = PO.NumIntDocPMP) "
     
    sWhere = "Where Exists " & sSubSelect
    
    sOrderBy = "ORDER BY Nivel DESC, Seq DESC, NumIntDocPOPai DESC"
      
    If objApontamentoSeleciona.dtDataOPFinal <> DATA_NULA Then
        sWhereAux = sWhereAux & "AND OP.DataEmissao <= ? "
    End If
    If objApontamentoSeleciona.dtDataOPInicial <> DATA_NULA Then
        sWhereAux = sWhereAux & "AND OP.DataEmissao >= ? "
    End If
    
    If objApontamentoSeleciona.lCTFinal <> 0 Then
        sWhereAux = sWhereAux & "AND PO.NumIntDocCT <= ? "
    End If
    If objApontamentoSeleciona.lCTInicial <> 0 Then
        sWhereAux = sWhereAux & "AND PO.NumIntDocCT >= ? "
    End If
    
    If objApontamentoSeleciona.sOPFinal <> "" Then
        sWhereAux = sWhereAux & "AND OP.Codigo <= ? "
    End If
    If objApontamentoSeleciona.sOPInicial <> "" Then
        sWhereAux = sWhereAux & "AND OP.Codigo >= ? "
    End If
    
    If objApontamentoSeleciona.sProdutoFinal <> "" Then
        sWhereAux = sWhereAux & "AND PO.Produto <= ? "
    End If
    If objApontamentoSeleciona.sProdutoInicial <> "" Then
        sWhereAux = sWhereAux & "AND PO.Produto >= ? "
    End If
    
    sWhere = sWhere & sWhereAux
    
    sTexto_SQL = sSelect & sFrom & sWhere & sOrderBy
 
    Exit Sub
 
Erro_ApontamentoProducao_PreparaSQL:
 
    Select Case gErr
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152418)
 
    End Select
 
    Exit Sub
 
End Sub

Private Function ApontamentoProducao_PreparaInt(ByVal lComando As Long, vlNumIntDoc As Variant, vlNumIntDocPOPai As Variant, vlNumIntDocPMP As Variant, vlNumIntDocOper As Variant, viNivel As Variant, viSeq As Variant, vsCodOPOrigem As Variant, vsProduto As Variant, vsVersao As Variant, vdQuantidade As Variant, vsUM As Variant, vlNumIntDocCT As Variant, vdtDataInicio As Variant, vdHoraInicio As Variant, vdtDataFim As Variant, vdHoraFim As Variant, ByVal vdtDataFimOP As Variant, ByVal vdtDataIniOP As Variant, ByVal vlCTFim As Variant, ByVal vlCTIni As Variant, ByVal vsCodOPFim As Variant, ByVal vsCodOPIni As Variant, ByVal vsProdutoFim As Variant, ByVal vsProdutoIni As Variant, ByVal sTexto_SQL As String) As Long
'Passa as variáveis de retorno e filtro do SELECT que monta os dados do Grid
 
Dim lErro As Long
 
On Error GoTo Erro_ApontamentoProducao_PreparaInt
 
    lErro = Comando_PrepararInt(lComando, sTexto_SQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137798
                       
    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137799
    
    lErro = Comando_BindVarInt(lComando, vlNumIntDocPOPai)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137800

    lErro = Comando_BindVarInt(lComando, vlNumIntDocPMP)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137801

    lErro = Comando_BindVarInt(lComando, vlNumIntDocOper)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137802
    
    lErro = Comando_BindVarInt(lComando, viNivel)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137803

    lErro = Comando_BindVarInt(lComando, viSeq)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137804

    lErro = Comando_BindVarInt(lComando, vsCodOPOrigem)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137805

    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137806
   
    lErro = Comando_BindVarInt(lComando, vsVersao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137807

    lErro = Comando_BindVarInt(lComando, vdQuantidade)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137808

    lErro = Comando_BindVarInt(lComando, vsUM)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137809

    lErro = Comando_BindVarInt(lComando, vlNumIntDocCT)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137810

    lErro = Comando_BindVarInt(lComando, vdtDataInicio)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137811

    lErro = Comando_BindVarInt(lComando, vdHoraInicio)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137812

    lErro = Comando_BindVarInt(lComando, vdtDataFim)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137813

    lErro = Comando_BindVarInt(lComando, vdHoraFim)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137814
   
    If vdtDataFimOP <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataFimOP)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137815
    End If
    If vdtDataIniOP <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataIniOP)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137816
    End If

    If vlCTFim <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCTFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137817
    End If
    If vlCTIni <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlCTIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137818
    End If

    If vsCodOPFim <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCodOPFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137819
    End If
    If vsCodOPIni <> "" Then
        lErro = Comando_BindVarInt(lComando, vsCodOPIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137820
    End If

    If vsProdutoFim <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdutoFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137821
    End If
    If vsProdutoIni <> "" Then
        lErro = Comando_BindVarInt(lComando, vsProdutoIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 137822
    End If
     
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 137823
 
   ApontamentoProducao_PreparaInt = SUCESSO
 
    Exit Function
 
Erro_ApontamentoProducao_PreparaInt:
 
    ApontamentoProducao_PreparaInt = gErr
 
    Select Case gErr
 
        Case 137798 To 137823
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sTexto_SQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152419)
 
    End Select
 
    Exit Function
 
End Function
 

Public Function CentrodeTrabalho_Le(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCentrodeTrabalho As typeCentrodeTrabalho
Dim iDiaDaSemana As Integer

On Error GoTo Erro_CentrodeTrabalho_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134446

    'Alocação de espaço no buffer
    tCentrodeTrabalho.sNomeReduzido = String(STRING_CENTRODETRABALHO_NOMERED, 0)
    tCentrodeTrabalho.sDescricao = String(STRING_CENTRODETRABALHO_DESCRICAO, 0)

    'Verifica se filtra por FilialEmpresa
    If objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA Then
        
        'Le a tabelaCentrodeTrabalho sem filtrar pela FilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, CargaMin, CargaMax, Turnos, HorasTurno, DomingoUtil, SegundaUtil, TercaUtil, QuartaUtil, QuintaUtil, SextaUtil, SabadoUtil FROM CentrodeTrabalho WHERE Codigo= ? ", _
                    tCentrodeTrabalho.lNumIntDoc, tCentrodeTrabalho.lCodigo, tCentrodeTrabalho.iFilialEmpresa, tCentrodeTrabalho.sNomeReduzido, tCentrodeTrabalho.sDescricao, _
                    tCentrodeTrabalho.dCargaMin, tCentrodeTrabalho.dCargaMax, tCentrodeTrabalho.iTurnos, tCentrodeTrabalho.dHorasTurno, tCentrodeTrabalho.iDiaisUteis(DOMINGO), tCentrodeTrabalho.iDiaisUteis(SEGUNDA), _
                    tCentrodeTrabalho.iDiaisUteis(TERCA), tCentrodeTrabalho.iDiaisUteis(QUARTA), tCentrodeTrabalho.iDiaisUteis(QUINTA), tCentrodeTrabalho.iDiaisUteis(SEXTA), tCentrodeTrabalho.iDiaisUteis(SABADO), _
                    objCentrodeTrabalho.lCodigo)
    
    Else
        
        'Le a tabelaCentrodeTrabalho filtrando pela FilialEmpresa
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, CargaMin, CargaMax, Turnos, HorasTurno, DomingoUtil, SegundaUtil, TercaUtil, QuartaUtil, QuintaUtil, SextaUtil, SabadoUtil FROM CentrodeTrabalho WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                    tCentrodeTrabalho.lNumIntDoc, tCentrodeTrabalho.lCodigo, tCentrodeTrabalho.iFilialEmpresa, tCentrodeTrabalho.sNomeReduzido, tCentrodeTrabalho.sDescricao, _
                    tCentrodeTrabalho.dCargaMin, tCentrodeTrabalho.dCargaMax, tCentrodeTrabalho.iTurnos, tCentrodeTrabalho.dHorasTurno, tCentrodeTrabalho.iDiaisUteis(DOMINGO), tCentrodeTrabalho.iDiaisUteis(SEGUNDA), _
                    tCentrodeTrabalho.iDiaisUteis(TERCA), tCentrodeTrabalho.iDiaisUteis(QUARTA), tCentrodeTrabalho.iDiaisUteis(QUINTA), tCentrodeTrabalho.iDiaisUteis(SEXTA), tCentrodeTrabalho.iDiaisUteis(SABADO), _
                    objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 134447

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134448

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134449

    objCentrodeTrabalho.lNumIntDoc = tCentrodeTrabalho.lNumIntDoc
    objCentrodeTrabalho.lCodigo = tCentrodeTrabalho.lCodigo
    objCentrodeTrabalho.iFilialEmpresa = tCentrodeTrabalho.iFilialEmpresa
    objCentrodeTrabalho.sNomeReduzido = tCentrodeTrabalho.sNomeReduzido
    objCentrodeTrabalho.sDescricao = tCentrodeTrabalho.sDescricao
    objCentrodeTrabalho.dCargaMin = tCentrodeTrabalho.dCargaMin
    objCentrodeTrabalho.dCargaMax = tCentrodeTrabalho.dCargaMax
    objCentrodeTrabalho.iTurnos = tCentrodeTrabalho.iTurnos
    objCentrodeTrabalho.dHorasTurno = tCentrodeTrabalho.dHorasTurno
    
    For iDiaDaSemana = DOMINGO To SABADO
    
        objCentrodeTrabalho.iDiaisUteis(iDiaDaSemana) = tCentrodeTrabalho.iDiaisUteis(iDiaDaSemana)
    
    Next
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    CentrodeTrabalho_Le = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le:

    CentrodeTrabalho_Le = gErr

    Select Case gErr

        Case 134446
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134447, 134448
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)

        Case 134449 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152420)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CentrodeTrabalho_Le_NomeReduzido(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCentrodeTrabalho As typeCentrodeTrabalho
Dim iDiaDaSemana As Integer

On Error GoTo Erro_CentrodeTrabalho_Le_NomeReduzido

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134938

    'Alocação de espaço no buffer
    tCentrodeTrabalho.sNomeReduzido = String(STRING_CENTRODETRABALHO_NOMERED, 0)
    tCentrodeTrabalho.sDescricao = String(STRING_CENTRODETRABALHO_DESCRICAO, 0)

    'Le a tabelaCentrodeTrabalho
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, CargaMin, CargaMax, Turnos, HorasTurno, DomingoUtil, SegundaUtil, TercaUtil, QuartaUtil, QuintaUtil, SextaUtil, SabadoUtil FROM CentrodeTrabalho WHERE NomeReduzido= ? ", _
                tCentrodeTrabalho.lNumIntDoc, tCentrodeTrabalho.lCodigo, tCentrodeTrabalho.iFilialEmpresa, tCentrodeTrabalho.sNomeReduzido, tCentrodeTrabalho.sDescricao, _
                tCentrodeTrabalho.dCargaMin, tCentrodeTrabalho.dCargaMax, tCentrodeTrabalho.iTurnos, tCentrodeTrabalho.dHorasTurno, tCentrodeTrabalho.iDiaisUteis(DOMINGO), tCentrodeTrabalho.iDiaisUteis(SEGUNDA), _
                tCentrodeTrabalho.iDiaisUteis(TERCA), tCentrodeTrabalho.iDiaisUteis(QUARTA), tCentrodeTrabalho.iDiaisUteis(QUINTA), tCentrodeTrabalho.iDiaisUteis(SEXTA), tCentrodeTrabalho.iDiaisUteis(SABADO), _
                objCentrodeTrabalho.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 134939

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134940

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134941

    objCentrodeTrabalho.lNumIntDoc = tCentrodeTrabalho.lNumIntDoc
    objCentrodeTrabalho.lCodigo = tCentrodeTrabalho.lCodigo
    objCentrodeTrabalho.iFilialEmpresa = tCentrodeTrabalho.iFilialEmpresa
    objCentrodeTrabalho.sNomeReduzido = tCentrodeTrabalho.sNomeReduzido
    objCentrodeTrabalho.sDescricao = tCentrodeTrabalho.sDescricao
    objCentrodeTrabalho.dCargaMin = tCentrodeTrabalho.dCargaMin
    objCentrodeTrabalho.dCargaMax = tCentrodeTrabalho.dCargaMax
    objCentrodeTrabalho.iTurnos = tCentrodeTrabalho.iTurnos
    objCentrodeTrabalho.dHorasTurno = tCentrodeTrabalho.dHorasTurno

    For iDiaDaSemana = DOMINGO To SABADO
    
        objCentrodeTrabalho.iDiaisUteis(iDiaDaSemana) = tCentrodeTrabalho.iDiaisUteis(iDiaDaSemana)
    
    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CentrodeTrabalho_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_NomeReduzido:

    CentrodeTrabalho_Le_NomeReduzido = gErr

    Select Case gErr

        Case 134938
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134939, 134940
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)

        Case 134941 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152421)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function Competencias_Le(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCompetencias As typeCompetencias

On Error GoTo Erro_Competencias_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134329

    'Alocação de espaço no buffer
    tCompetencias.sNomeReduzido = String(STRING_COMPETENCIA_NOMERED, 0)
    tCompetencias.sDescricao = String(STRING_COMPETENCIA_DESCRICAO, 0)

    'Le a tabelaCompetencias
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NomeReduzido, Descricao, NumIntDocCT, Padrao FROM Competencias WHERE Codigo= ? ", _
                tCompetencias.lNumIntDoc, tCompetencias.lCodigo, tCompetencias.sNomeReduzido, tCompetencias.sDescricao, tCompetencias.lNumIntDocCT, tCompetencias.iPadrao, objCompetencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 134330

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134331

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134332

    objCompetencias.lNumIntDoc = tCompetencias.lNumIntDoc
    objCompetencias.lCodigo = tCompetencias.lCodigo
    objCompetencias.sNomeReduzido = tCompetencias.sNomeReduzido
    objCompetencias.sDescricao = tCompetencias.sDescricao
    objCompetencias.lNumIntDocCT = tCompetencias.lNumIntDocCT
    objCompetencias.iPadrao = tCompetencias.iPadrao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Competencias_Le = SUCESSO

    Exit Function

Erro_Competencias_Le:

    Competencias_Le = gErr

    Select Case gErr

        Case 134329
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134330, 134331
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 134332 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152422)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Competencias_Le_Padrao(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCompetencias As typeCompetencias

On Error GoTo Erro_Competencias_Le_Padrao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137841

    'Alocação de espaço no buffer
    tCompetencias.sNomeReduzido = String(STRING_COMPETENCIA_NOMERED, 0)
    tCompetencias.sDescricao = String(STRING_COMPETENCIA_DESCRICAO, 0)

    'Le a tabelaCompetencias
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NomeReduzido, Descricao, NumIntDocCT, Padrao FROM Competencias WHERE Padrao= ? ", _
                tCompetencias.lNumIntDoc, tCompetencias.lCodigo, tCompetencias.sNomeReduzido, tCompetencias.sDescricao, tCompetencias.lNumIntDocCT, tCompetencias.iPadrao, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 137842

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137843

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 137844

    objCompetencias.lNumIntDoc = tCompetencias.lNumIntDoc
    objCompetencias.lCodigo = tCompetencias.lCodigo
    objCompetencias.sNomeReduzido = tCompetencias.sNomeReduzido
    objCompetencias.sDescricao = tCompetencias.sDescricao
    objCompetencias.lNumIntDocCT = tCompetencias.lNumIntDocCT
    objCompetencias.iPadrao = tCompetencias.iPadrao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Competencias_Le_Padrao = SUCESSO

    Exit Function

Erro_Competencias_Le_Padrao:

    Competencias_Le_Padrao = gErr

    Select Case gErr

        Case 137841
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137842, 137843
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 137844 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152423)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function Competencias_Le_NomeReduzido(ByVal objCompetencias As ClassCompetencias) As Long
'Lê a tabela de Competências pelo NomeReduzido - Jorge Specian - 26/04/2005
Dim lErro As Long
Dim lComando As Long
Dim tCompetencias As typeCompetencias

On Error GoTo Erro_Competencias_Le_NomeReduzido

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134934

    'Alocação de espaço no buffer
    tCompetencias.sNomeReduzido = String(STRING_COMPETENCIA_NOMERED, 0)
    tCompetencias.sDescricao = String(STRING_COMPETENCIA_DESCRICAO, 0)

    'Le a tabelaCompetencias
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NomeReduzido, Descricao, NumIntDocCT, Padrao FROM Competencias WHERE NomeReduzido= ? ", _
                tCompetencias.lNumIntDoc, tCompetencias.lCodigo, tCompetencias.sNomeReduzido, tCompetencias.sDescricao, tCompetencias.lNumIntDocCT, tCompetencias.iPadrao, objCompetencias.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 134935

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134936

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134937

    objCompetencias.lNumIntDoc = tCompetencias.lNumIntDoc
    objCompetencias.lCodigo = tCompetencias.lCodigo
    objCompetencias.sNomeReduzido = tCompetencias.sNomeReduzido
    objCompetencias.sDescricao = tCompetencias.sDescricao
    objCompetencias.lNumIntDocCT = tCompetencias.lNumIntDocCT
    objCompetencias.iPadrao = tCompetencias.iPadrao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Competencias_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Competencias_Le_NomeReduzido:

    Competencias_Le_NomeReduzido = gErr

    Select Case gErr

        Case 134934
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134935, 134936
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 134937 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152424)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function Competencias_Le_CTPadrao(ByVal objCompetencias As ClassCompetencias, objCT As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDocCT As Long

On Error GoTo Erro_Competencias_Le_CTPadrao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134905

    'Le a tabelaCompetencias
    lErro = Comando_Executar(lComando, "SELECT NumIntDocCT FROM Competencias WHERE Codigo= ? ", _
                lNumIntDocCT, objCompetencias.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 134906

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134907

    'Encontrou a Competencia e CTPadrao está preenchido
    If lErro = SUCESSO And lNumIntDocCT <> 0 Then
        
        objCT.lNumIntDoc = lNumIntDocCT
        
        'Le Centro de Trabalho a partir o NumIntDoc
        lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 134908
        
        'Não encontrou o CT cadastrado em Competencias
        If lErro = 134590 Then gError 134909
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Competencias_Le_CTPadrao = SUCESSO

    Exit Function

Erro_Competencias_Le_CTPadrao:

    Competencias_Le_CTPadrao = gErr

    Select Case gErr

        Case 134905
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134906, 134907
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 134908
            'erro tratado na rotina chamada

        Case 134909 'Sem dados -> Tratado na rotina chamadora
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152425)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Competencias_VerificaCTPadrao(ByVal objCT As ClassCentrodeTrabalho) As Long
'Verifica se o CTPadrao das Competencias estão vazios e dá o aviso que este será o Padrão
'por Jorge Specian - 25/04/2005
Dim lErro As Long
Dim objCTCompetencias As ClassCTCompetencias
Dim objCompetencias As ClassCompetencias
Dim sCompetencias As String
Dim iIndice As Integer

On Error GoTo Erro_Competencias_VerificaCTPadrao

    'Verifica se o CT Padrão da Competencias está vazio
    For Each objCTCompetencias In objCT.colCompetencias
    
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objCTCompetencias.lNumIntDocCompet
        
        'Lê a Competencia pelo NumIntDoc
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 134925
        
        'Se o CT Padrão está vazio vai montando a msg
        If objCompetencias.lNumIntDocCT = 0 Then
        
            If Len(sCompetencias) > 0 Then
                sCompetencias = sCompetencias & ", "
            End If
            sCompetencias = sCompetencias & Trim(CStr(objCompetencias.lCodigo)) & SEPARADOR & Trim(CStr(objCompetencias.sNomeReduzido))
        
        End If
    
    Next
    
    'Troca a vírgula final por "e"
    For iIndice = Len(sCompetencias) To 1 Step -1
    
        If Mid(sCompetencias, iIndice, 1) = "," Then
        
            sCompetencias = Mid(sCompetencias, 1, iIndice - 1) & " e" & right(sCompetencias, Len(sCompetencias) - iIndice)
            Exit For
            
        End If
    
    Next iIndice
    
    'Se tem msg -> exibe o AVISO
    If Len(sCompetencias) > 0 Then
    
        Call Rotina_Aviso(vbOKOnly, "AVISO_CENTROTRABALHO_SERA_PADRAO", sCompetencias)

    End If

    Competencias_VerificaCTPadrao = SUCESSO

    Exit Function

Erro_Competencias_VerificaCTPadrao:

    Competencias_VerificaCTPadrao = gErr

    Select Case gErr

        Case 134925
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152426)

    End Select

    Exit Function

End Function

Public Function Competencias_Le_NumIntDoc(ByVal objCompetencias As ClassCompetencias) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCompetencias As typeCompetencias

On Error GoTo Erro_Competencias_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134333

    'Alocação de espaço no buffer
    tCompetencias.sNomeReduzido = String(STRING_COMPETENCIA_NOMERED, 0)
    tCompetencias.sDescricao = String(STRING_COMPETENCIA_DESCRICAO, 0)

    'Le a tabelaCompetencias
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NomeReduzido, Descricao, NumIntDocCT, Padrao FROM Competencias WHERE NumIntDoc= ? ", _
                tCompetencias.lNumIntDoc, tCompetencias.lCodigo, tCompetencias.sNomeReduzido, tCompetencias.sDescricao, tCompetencias.lNumIntDocCT, tCompetencias.iPadrao, objCompetencias.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134334

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134335

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134336

    objCompetencias.lNumIntDoc = tCompetencias.lNumIntDoc
    objCompetencias.lCodigo = tCompetencias.lCodigo
    objCompetencias.sNomeReduzido = tCompetencias.sNomeReduzido
    objCompetencias.sDescricao = tCompetencias.sDescricao
    objCompetencias.lNumIntDocCT = tCompetencias.lNumIntDocCT
    objCompetencias.iPadrao = tCompetencias.iPadrao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Competencias_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_Competencias_Le_NumIntDoc:

    Competencias_Le_NumIntDoc = gErr

    Select Case gErr

        Case 134333
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134334, 134335
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 134336 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152427)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TaxaDeProducao_Le(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTaxaDeProducao As typeTaxaDeProducao

On Error GoTo Erro_TaxaDeProducao_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134540

    'Alocação de espaço no buffer
    tTaxaDeProducao.sProduto = String(STRING_PRODUTO, 0)
    tTaxaDeProducao.sUMProduto = String(STRING_UM_SIGLA, 0)
    tTaxaDeProducao.sUMTempo = String(STRING_UM_SIGLA, 0)

    'Le a tabelaTaxaDeProducao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, " & _
                "UMTempo, Ativo, Data, DataDesativacao FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocMaq= ? AND NumIntDocCompet= ? AND Ativo = ? ", _
                tTaxaDeProducao.lNumIntDoc, tTaxaDeProducao.dLoteMax, tTaxaDeProducao.dLoteMin, tTaxaDeProducao.dLotePadrao, tTaxaDeProducao.dTempoPreparacao, tTaxaDeProducao.dTempoMovimentacao, tTaxaDeProducao.dTempoDescarga, _
                tTaxaDeProducao.iTipo, tTaxaDeProducao.dQuantidade, tTaxaDeProducao.sUMProduto, tTaxaDeProducao.dTempoOperacao, tTaxaDeProducao.sUMTempo, _
                tTaxaDeProducao.iAtivo, tTaxaDeProducao.dtData, tTaxaDeProducao.dtDataDesativacao, _
                objTaxaDeProducao.sProduto, objTaxaDeProducao.lNumIntDocMaq, objTaxaDeProducao.lNumIntDocCompet, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134541

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134542

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134543

    objTaxaDeProducao.lNumIntDoc = tTaxaDeProducao.lNumIntDoc
    objTaxaDeProducao.dLoteMin = tTaxaDeProducao.dLoteMin
    objTaxaDeProducao.dLoteMax = tTaxaDeProducao.dLoteMax
    objTaxaDeProducao.dLotePadrao = tTaxaDeProducao.dLotePadrao
    objTaxaDeProducao.dTempoPreparacao = tTaxaDeProducao.dTempoPreparacao
    objTaxaDeProducao.dTempoMovimentacao = tTaxaDeProducao.dTempoMovimentacao
    objTaxaDeProducao.dTempoDescarga = tTaxaDeProducao.dTempoDescarga
    objTaxaDeProducao.iTipo = tTaxaDeProducao.iTipo
    objTaxaDeProducao.dQuantidade = tTaxaDeProducao.dQuantidade
    objTaxaDeProducao.sUMProduto = tTaxaDeProducao.sUMProduto
    objTaxaDeProducao.dTempoOperacao = tTaxaDeProducao.dTempoOperacao
    objTaxaDeProducao.sUMTempo = tTaxaDeProducao.sUMTempo
    objTaxaDeProducao.iAtivo = tTaxaDeProducao.iAtivo
    objTaxaDeProducao.dtData = tTaxaDeProducao.dtData
    If tTaxaDeProducao.dtDataDesativacao <> DATA_NULA Then
        objTaxaDeProducao.dtDataDesativacao = tTaxaDeProducao.dtDataDesativacao
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TaxaDeProducao_Le = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Le:

    TaxaDeProducao_Le = gErr

    Select Case gErr

        Case 134540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134541, 134542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134543 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152428)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function RoteirosDeFabricacao_Le(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long

On Error GoTo Erro_RoteirosDeFabricacao_Le

    'Lê o RoteirosDeFabricacao que está sendo Passado
    lErro = CF("RoteirosDeFabricacao_Le_Roteiro", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134622 Then gError 134616
    
    If lErro = 134622 Then gError 134617 'não encontrou dados
    
    'Lê as Operacoes do RoteirosDeFabricacao que está sendo Passado
    lErro = CF("RoteirosDeFabricacao_Le_Operacoes", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134626 Then gError 134618

    RoteirosDeFabricacao_Le = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Le:

    RoteirosDeFabricacao_Le = gErr

    Select Case gErr

        Case 134617 'Sem dados -> Tratado na rotina chamadora
        
        Case 134616, 164618
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152429)

    End Select

    Exit Function

End Function


Public Function RoteirosDeFabricacao_Le_Roteiro(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteirosDeFabricacao As typeRoteirosDeFabricacao

On Error GoTo Erro_RoteirosDeFabricacao_Le_Roteiro

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134619

    'Alocação de espaço no buffer
    tRoteirosDeFabricacao.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tRoteirosDeFabricacao.sVersao = String(STRING_KIT_VERSAO, 0)
    tRoteirosDeFabricacao.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tRoteirosDeFabricacao.sUM = String(STRING_UM_SIGLA, 0)
    tRoteirosDeFabricacao.sAutor = String(STRING_AUTOR_ROT_FAB, 0)

    'Le a tabelaRoteirosDeFabricacao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ProdutoRaiz, Versao, Descricao, DataCriacao, DataUltModificacao, " & _
                "Quantidade, UM, Autor, NumMaxMaqPorOper FROM RoteirosDeFabricacao WHERE ProdutoRaiz= ?  AND Versao= ? ", _
                tRoteirosDeFabricacao.lNumIntDoc, tRoteirosDeFabricacao.sProdutoRaiz, tRoteirosDeFabricacao.sVersao, tRoteirosDeFabricacao.sDescricao, _
                tRoteirosDeFabricacao.dtDataCriacao, tRoteirosDeFabricacao.dtDataUltModificacao, tRoteirosDeFabricacao.dQuantidade, tRoteirosDeFabricacao.sUM, tRoteirosDeFabricacao.sAutor, tRoteirosDeFabricacao.iNumMaxMaqPorOper, _
                objRoteirosDeFabricacao.sProdutoRaiz, objRoteirosDeFabricacao.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 134620

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134621

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134622

    objRoteirosDeFabricacao.lNumIntDoc = tRoteirosDeFabricacao.lNumIntDoc
    objRoteirosDeFabricacao.sProdutoRaiz = tRoteirosDeFabricacao.sProdutoRaiz
    objRoteirosDeFabricacao.sVersao = tRoteirosDeFabricacao.sVersao
    objRoteirosDeFabricacao.sDescricao = tRoteirosDeFabricacao.sDescricao
    objRoteirosDeFabricacao.dtDataCriacao = tRoteirosDeFabricacao.dtDataCriacao
    objRoteirosDeFabricacao.dtDataUltModificacao = tRoteirosDeFabricacao.dtDataUltModificacao
    objRoteirosDeFabricacao.dQuantidade = tRoteirosDeFabricacao.dQuantidade
    objRoteirosDeFabricacao.sUM = tRoteirosDeFabricacao.sUM
    objRoteirosDeFabricacao.sAutor = tRoteirosDeFabricacao.sAutor
    objRoteirosDeFabricacao.iNumMaxMaqPorOper = tRoteirosDeFabricacao.iNumMaxMaqPorOper

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteirosDeFabricacao_Le_Roteiro = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Le_Roteiro:

    RoteirosDeFabricacao_Le_Roteiro = gErr

    Select Case gErr

        Case 134619
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134620, 134621
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)

        Case 134622 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152430)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RoteirosDeFabricacao_Le_Operacoes(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim objOperacoes As ClassOperacoes
Dim tOperacoes As typeOperacoes

On Error GoTo Erro_RoteirosDeFabricacao_Le_Operacoes

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134623

    'Alocação de espaço no buffer
    tOperacoes.sObservacao = String(STRING_OBSERVACAO, 0)
    
    sComandoSQL = "SELECT NumIntDoc, NumIntDocRotFabr, Seq, NumIntDocCompet, NumIntDocCT, "
    sComandoSQL = sComandoSQL & "Observacao, IgnoraTaxaProducao, SeqPai, SeqArvore, Nivel, "
    sComandoSQL = sComandoSQL & "PosicaoArvore, NumMaxMaqPorOper, NumRepeticoes From Operacoes WHERE NumIntDocRotFabr = ? "
    sComandoSQL = sComandoSQL & "Order by Seq"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tOperacoes.lNumIntDoc, _
                            tOperacoes.lNumIntDocRotFabr, tOperacoes.iSeq, _
                            tOperacoes.lNumIntDocCompet, tOperacoes.lNumIntDocCT, _
                            tOperacoes.sObservacao, tOperacoes.iIgnoraTaxaProducao, _
                            tOperacoes.iSeqPai, tOperacoes.iSeqArvore, tOperacoes.iNivel, _
                            tOperacoes.iPosicaoArvore, tOperacoes.iNumMaxMaqPorOper, tOperacoes.iNumRepeticoes, objRoteirosDeFabricacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134624

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134625

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134626

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOperacoes = New ClassOperacoes
        
        objOperacoes.lNumIntDoc = tOperacoes.lNumIntDoc
        objOperacoes.lNumIntDocRotFabr = tOperacoes.lNumIntDocRotFabr
        objOperacoes.iSeq = tOperacoes.iSeq
        objOperacoes.lNumIntDocCompet = tOperacoes.lNumIntDocCompet
        objOperacoes.lNumIntDocCT = tOperacoes.lNumIntDocCT
        objOperacoes.sObservacao = tOperacoes.sObservacao
        objOperacoes.iIgnoraTaxaProducao = tOperacoes.iIgnoraTaxaProducao
        objOperacoes.iSeqPai = tOperacoes.iSeqPai
        objOperacoes.iSeqArvore = tOperacoes.iSeqArvore
        objOperacoes.iNivel = tOperacoes.iNivel
        objOperacoes.iPosicaoArvore = tOperacoes.iPosicaoArvore
        objOperacoes.iNumMaxMaqPorOper = tOperacoes.iNumMaxMaqPorOper
        objOperacoes.iNumRepeticoes = tOperacoes.iNumRepeticoes
        
        'Lê OperacoesTempo associada a Operação que está sendo lida
        lErro = CF("RoteirosDeFabricacao_Le_OperacoesTempo", objOperacoes)
        If lErro <> SUCESSO And lErro <> 134633 Then gError 134627
        
        'Lê as OperacaoInsumos associadas a Operação que está sendo lida
        lErro = CF("RoteirosDeFabricacao_Le_OperacaoInsumos", objOperacoes)
        If lErro <> SUCESSO And lErro <> 134637 Then gError 134628
        
        objRoteirosDeFabricacao.colOperacoes.Add objOperacoes
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134629
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteirosDeFabricacao_Le_Operacoes = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Le_Operacoes:

    RoteirosDeFabricacao_Le_Operacoes = gErr

    Select Case gErr

        Case 134623
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134624, 134625, 134629
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOES", gErr)

        Case 134626 'Sem dados -> Tratado na rotina chamadora
        
        Case 134627, 134628
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152431)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RoteirosDeFabricacao_Le_OperacoesTempo(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOperacoesTempo As ClassOperacoesTempo
Dim tOperacoesTempo As typeOperacoesTempo
Dim sComandoSQL As String

On Error GoTo Erro_RoteirosDeFabricacao_Le_OperacoesTempo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134630

    'Alocação de espaço no buffer
    tOperacoesTempo.sUMTempo = String(STRING_UM_SIGLA, 0)

    sComandoSQL = "SELECT NumIntDocOperacao, LoteMax, LoteMin, LotePadrao, "
    sComandoSQL = sComandoSQL & "TempoPreparacao, TempoMovimentacao, TempoDescarga, "
    sComandoSQL = sComandoSQL & "TempoOperacao, UMTempo, Tipo, NumIntDocMaq "
    sComandoSQL = sComandoSQL & "From OperacoesTempo WHERE NumIntDocOperacao = ?"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tOperacoesTempo.lNumIntDocOperacao, _
                            tOperacoesTempo.dLoteMax, tOperacoesTempo.dLoteMin, _
                            tOperacoesTempo.dLotePadrao, tOperacoesTempo.dTempoPreparacao, _
                            tOperacoesTempo.dTempoMovimentacao, tOperacoesTempo.dTempoDescarga, _
                            tOperacoesTempo.dTempoOperacao, tOperacoesTempo.sUMTempo, _
                            tOperacoesTempo.iTipo, tOperacoesTempo.lNumIntDocMaq, objOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134631

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134632

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134633

    Set objOperacoesTempo = New ClassOperacoesTempo
    
    objOperacoesTempo.lNumIntDocOperacao = tOperacoesTempo.lNumIntDocOperacao
    objOperacoesTempo.dLoteMax = tOperacoesTempo.dLoteMax
    objOperacoesTempo.dLoteMin = tOperacoesTempo.dLoteMin
    objOperacoesTempo.dLotePadrao = tOperacoesTempo.dLotePadrao
    objOperacoesTempo.dTempoPreparacao = tOperacoesTempo.dTempoPreparacao
    objOperacoesTempo.dTempoMovimentacao = tOperacoesTempo.dTempoMovimentacao
    objOperacoesTempo.dTempoDescarga = tOperacoesTempo.dTempoDescarga
    objOperacoesTempo.dTempoOperacao = tOperacoesTempo.dTempoOperacao
    objOperacoesTempo.sUMTempo = tOperacoesTempo.sUMTempo
    objOperacoesTempo.iTipo = tOperacoesTempo.iTipo
    objOperacoesTempo.lNumIntDocMaq = tOperacoesTempo.lNumIntDocMaq
    
    Set objOperacoes.objOperacoesTempo = objOperacoesTempo

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteirosDeFabricacao_Le_OperacoesTempo = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Le_OperacoesTempo:

    RoteirosDeFabricacao_Le_OperacoesTempo = gErr

    Select Case gErr

        Case 134630
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134631, 134632
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOESTEMPO", gErr)

        Case 134633 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152432)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RoteirosDeFabricacao_Le_OperacaoInsumos(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim objOperacaoInsumos As ClassOperacaoInsumos
Dim tOperacaoInsumos As typeOperacaoInsumos

On Error GoTo Erro_RoteirosDeFabricacao_Le_OperacaoInsumos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134634

    'Alocação de espaço no buffer
    tOperacaoInsumos.sProduto = String(STRING_PRODUTO, 0)
    tOperacaoInsumos.sUMProduto = String(STRING_UM_SIGLA, 0)
    tOperacaoInsumos.sVersaoKitComp = String(STRING_KIT_VERSAO, 0)
    
    sComandoSQL = "SELECT NumIntDocOper, Produto, Quantidade, UMProduto, VersaoKitComp, "
    sComandoSQL = sComandoSQL & "Composicao, PercentualPerda, CustoStandard "
    sComandoSQL = sComandoSQL & "From OperacaoInsumos WHERE NumIntDocOper = ?"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tOperacaoInsumos.lNumIntDocOper, _
                            tOperacaoInsumos.sProduto, tOperacaoInsumos.dQuantidade, _
                            tOperacaoInsumos.sUMProduto, tOperacaoInsumos.sVersaoKitComp, _
                            tOperacaoInsumos.iComposicao, tOperacaoInsumos.dPercentualPerda, _
                            tOperacaoInsumos.dCustoStandard, objOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134635

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134636

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134637

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOperacaoInsumos = New ClassOperacaoInsumos
        
        objOperacaoInsumos.lNumIntDocOper = tOperacaoInsumos.lNumIntDocOper
        objOperacaoInsumos.sProduto = tOperacaoInsumos.sProduto
        objOperacaoInsumos.dQuantidade = tOperacaoInsumos.dQuantidade
        objOperacaoInsumos.sUMProduto = tOperacaoInsumos.sUMProduto
        objOperacaoInsumos.sVersaoKitComp = tOperacaoInsumos.sVersaoKitComp
        objOperacaoInsumos.iComposicao = tOperacaoInsumos.iComposicao
        objOperacaoInsumos.dPercentualPerda = tOperacaoInsumos.dPercentualPerda
        objOperacaoInsumos.dCustoStandard = tOperacaoInsumos.dCustoStandard
                            
        objOperacoes.colOperacaoInsumos.Add objOperacaoInsumos
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134638
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteirosDeFabricacao_Le_OperacaoInsumos = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Le_OperacaoInsumos:

    RoteirosDeFabricacao_Le_OperacaoInsumos = gErr

    Select Case gErr

        Case 134634
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134635, 134636, 134638
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACAOINSUMOS", gErr)

        Case 134637 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152433)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function CentrodeTrabalho_Le_CTMaquinas(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTMaquinas As ClassCTMaquinas
Dim tCTMaquinas As typeCTMaquinas

On Error GoTo Erro_CentrodeTrabalho_Le_CTMaquinas

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134455

    lErro = Comando_Executar(lComando, "SELECT NumIntDocMaq, NumIntDocCT, Quantidade FROM CTMaquinas WHERE NumIntDocCT = ? Order by NumIntDocMaq", tCTMaquinas.lNumIntDocMaq, tCTMaquinas.lNumIntDocCT, tCTMaquinas.iQuantidade, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> SUCESSO Then gError 134456
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134457

    'Maquina não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134458

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCTMaquinas = New ClassCTMaquinas

        objCTMaquinas.lNumIntDocMaq = tCTMaquinas.lNumIntDocMaq
        objCTMaquinas.lNumIntDocCT = tCTMaquinas.lNumIntDocCT
        objCTMaquinas.iQuantidade = tCTMaquinas.iQuantidade
        
        'lê os turnos da maquina
        lErro = CentrodeTrabalho_Le_CTMaquinasTurnos(objCTMaquinas)
        If lErro <> SUCESSO And lErro <> 137228 Then gError 137224

        objCentrodeTrabalho.colMaquinas.Add objCTMaquinas

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134459

    Loop
    
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_Le_CTMaquinas = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_CTMaquinas:

   CentrodeTrabalho_Le_CTMaquinas = gErr

   Select Case gErr
    
        Case 134455
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 134456, 134457, 134459
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINAS", gErr)
            
        Case 134458
            'Erro tratado na rotina chamadora
            
        Case 137224
            'Erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152434)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Function CentrodeTrabalho_Le_CTCompetencias(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTCompetencias As ClassCTCompetencias
Dim tCTCompetencias As typeCTCompetencias

On Error GoTo Erro_CentrodeTrabalho_Le_CTCompetencias

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134450
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDocCT, NumIntDocCompet FROM CTCompetencias WHERE NumIntDocCT = ? Order by NumIntDocCompet", tCTCompetencias.lNumIntDocCT, tCTCompetencias.lNumIntDocCompet, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> SUCESSO Then gError 134451
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134452

    'CTCompetencia não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134453

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCTCompetencias = New ClassCTCompetencias

        objCTCompetencias.lNumIntDocCT = tCTCompetencias.lNumIntDocCT
        objCTCompetencias.lNumIntDocCompet = tCTCompetencias.lNumIntDocCompet

        objCentrodeTrabalho.colCompetencias.Add objCTCompetencias

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134454

    Loop
    
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_Le_CTCompetencias = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_CTCompetencias:

   CentrodeTrabalho_Le_CTCompetencias = gErr

   Select Case gErr
    
        Case 134450
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 134451, 134452, 134454
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTCOMPETENCIAS", gErr)
            
        Case 134453
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152435)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function TiposDeMaodeObra_Le(ByVal objTiposDeMaodeObra As ClassTiposDeMaodeObra) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTiposDeMaodeObra As typeTiposDeMaodeObra

On Error GoTo Erro_TiposDeMaodeObra_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137595

    'Alocação de espaço no buffer
    tTiposDeMaodeObra.sDescricao = String(STRING_TIPO_MO_DESCRICAO, 0)
    tTiposDeMaodeObra.sObservacao = String(STRING_TIPO_MO_OBSERVACAO, 0)
    tTiposDeMaodeObra.sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaTiposDeMaodeObra
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Observacao, CustoHora, Produto FROM TiposDeMaodeObra WHERE Codigo= ? ", _
                tTiposDeMaodeObra.iCodigo, tTiposDeMaodeObra.sDescricao, tTiposDeMaodeObra.sObservacao, tTiposDeMaodeObra.dCustoHora, tTiposDeMaodeObra.sProduto, objTiposDeMaodeObra.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 137596

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137597

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 137598

    objTiposDeMaodeObra.iCodigo = tTiposDeMaodeObra.iCodigo
    objTiposDeMaodeObra.sDescricao = tTiposDeMaodeObra.sDescricao
    objTiposDeMaodeObra.sObservacao = tTiposDeMaodeObra.sObservacao
    objTiposDeMaodeObra.dCustoHora = tTiposDeMaodeObra.dCustoHora
    objTiposDeMaodeObra.sProduto = tTiposDeMaodeObra.sProduto

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TiposDeMaodeObra_Le = SUCESSO

    Exit Function

Erro_TiposDeMaodeObra_Le:

    TiposDeMaodeObra_Le = gErr

    Select Case gErr

        Case 137595
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137596, 137597
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEMAODEOBRA", gErr)

        Case 137598 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152436)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CentrodeTrabalho_Le_NumIntDoc(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long
'Lê o CentroDeTrabalho pelo NumIntDoc passado no objCentrodeTrabalho

Dim lErro As Long
Dim lComando As Long
Dim tCentrodeTrabalho As typeCentrodeTrabalho
Dim iDiaDaSemana As Integer

On Error GoTo Erro_CentrodeTrabalho_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137065

    'Alocação de espaço no buffer
    tCentrodeTrabalho.sNomeReduzido = String(STRING_CENTRODETRABALHO_NOMERED, 0)
    tCentrodeTrabalho.sDescricao = String(STRING_CENTRODETRABALHO_DESCRICAO, 0)

    'Le a tabelaCentrodeTrabalho
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, NomeReduzido, Descricao, CargaMin, CargaMax, Turnos, HorasTurno, DomingoUtil, SegundaUtil, TercaUtil, QuartaUtil, QuintaUtil, SextaUtil, SabadoUtil FROM CentrodeTrabalho WHERE NumIntDoc= ? ", _
                tCentrodeTrabalho.lNumIntDoc, tCentrodeTrabalho.lCodigo, tCentrodeTrabalho.iFilialEmpresa, tCentrodeTrabalho.sNomeReduzido, tCentrodeTrabalho.sDescricao, _
                tCentrodeTrabalho.dCargaMin, tCentrodeTrabalho.dCargaMax, tCentrodeTrabalho.iTurnos, tCentrodeTrabalho.dHorasTurno, tCentrodeTrabalho.iDiaisUteis(DOMINGO), tCentrodeTrabalho.iDiaisUteis(SEGUNDA), _
                tCentrodeTrabalho.iDiaisUteis(TERCA), tCentrodeTrabalho.iDiaisUteis(QUARTA), tCentrodeTrabalho.iDiaisUteis(QUINTA), tCentrodeTrabalho.iDiaisUteis(SEXTA), tCentrodeTrabalho.iDiaisUteis(SABADO), _
                objCentrodeTrabalho.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137066

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137067

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134590

    objCentrodeTrabalho.lNumIntDoc = tCentrodeTrabalho.lNumIntDoc
    objCentrodeTrabalho.lCodigo = tCentrodeTrabalho.lCodigo
    objCentrodeTrabalho.iFilialEmpresa = tCentrodeTrabalho.iFilialEmpresa
    objCentrodeTrabalho.sNomeReduzido = tCentrodeTrabalho.sNomeReduzido
    objCentrodeTrabalho.sDescricao = tCentrodeTrabalho.sDescricao
    objCentrodeTrabalho.dCargaMin = tCentrodeTrabalho.dCargaMin
    objCentrodeTrabalho.dCargaMax = tCentrodeTrabalho.dCargaMax
    objCentrodeTrabalho.iTurnos = tCentrodeTrabalho.iTurnos
    objCentrodeTrabalho.dHorasTurno = tCentrodeTrabalho.dHorasTurno

    For iDiaDaSemana = DOMINGO To SABADO
    
        objCentrodeTrabalho.iDiaisUteis(iDiaDaSemana) = tCentrodeTrabalho.iDiaisUteis(iDiaDaSemana)
    
    Next

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CentrodeTrabalho_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_NumIntDoc:

    CentrodeTrabalho_Le_NumIntDoc = gErr

    Select Case gErr

        Case 137065
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137066, 137067
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CENTRODETRABALHO", gErr)

        Case 134590 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152437)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TaxaDeProducao_Seleciona(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sNomeFunc As String

On Error GoTo Erro_TaxaDeProducao_Seleciona

    'Seleciona a TaxaDeProducao com os dados exatos (competencia, máquina e produto)
    lErro = CF("TaxaDeProducao_Seleciona_Exata", objTaxaDeProducao)
    If lErro <> SUCESSO And lErro <> 134592 Then gError 134593
    
    'não encontrei
    If lErro <> SUCESSO Then
    
        'Seleciona a TaxaDeProducao conforme prioridades ...
        For iIndice = 1 To NUM_PRIORIDADES_SELECAO
           
           Select Case iIndice
                
                Case gobjEST.iPrioridadeMaquina
                   sNomeFunc = "TaxaDeproducao_Seleciona_" & STRING_PRIORIDADE_MAQUINAS
                
                Case gobjEST.iPrioridadeProduto
                   sNomeFunc = "TaxaDeproducao_Seleciona_" & STRING_PRIORIDADE_PRODUTOS

                Case gobjEST.iPrioridadeProdutoAnc
                   sNomeFunc = "TaxaDeproducao_Seleciona_" & STRING_PRIORIDADE_PRODUTOS_ANC

           End Select
           
           lErro = CF(sNomeFunc, objTaxaDeProducao)
           If lErro <> SUCESSO And lErro <> 134599 And _
              lErro <> 134604 And lErro <> 134607 Then gError 134595

           If lErro = SUCESSO Then Exit For ' Já encontrou
           
        Next
    
        'não encontrei nada !!!
        If lErro <> SUCESSO Then
    
            'Seleciona a TaxaDeProducao somente pela Competencia
            lErro = CF("TaxaDeProducao_Seleciona_Competencia", objTaxaDeProducao)
            If lErro <> SUCESSO And lErro <> 134674 Then gError 134675
            
            'não encontrei nada MESMO!!!
            If lErro <> SUCESSO Then gError 134600
    
        End If
        
    End If

    TaxaDeProducao_Seleciona = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Seleciona:

    TaxaDeProducao_Seleciona = gErr

    Select Case gErr

        Case 134592, 134595, 134675
            'erro tratado na rotina chamada

        Case 134600 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152438)

    End Select

    Exit Function

End Function

Public Function TaxaDeProducao_Seleciona_Exata(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_TaxaDeProducao_Seleciona_Exata

    'Le a tabelaTaxaDeProducao
    lErro = CF("TaxaDeProducao_Le", objTaxaDeProducao)
    If lErro <> SUCESSO And lErro <> 134543 Then gError 134591
    
    'não encontrei
    If lErro = 134543 Then gError 134592

    TaxaDeProducao_Seleciona_Exata = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Seleciona_Exata:

    TaxaDeProducao_Seleciona_Exata = gErr

    Select Case gErr

        Case 134591
            'erro tratado na rotina chamada

        Case 134592 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152439)

    End Select

    Exit Function

End Function


Public Function TaxaDeProducao_Seleciona_Maquinas(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTaxaDeProducao As typeTaxaDeProducao
Dim sProdutoAncestral As String

On Error GoTo Erro_TaxaDeProducao_Seleciona_Maquinas

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134596

    'Alocação de espaço no buffer
    tTaxaDeProducao.sProduto = String(STRING_PRODUTO, 0)
    tTaxaDeProducao.sUMProduto = String(STRING_UM_SIGLA, 0)
    tTaxaDeProducao.sUMTempo = String(STRING_UM_SIGLA, 0)
    sProdutoAncestral = String(STRING_PRODUTO, 0)
    
    If Len(objTaxaDeProducao.sProduto) <> 0 Then
    
        'verifica qual produto é pai deste
        lErro = Mascara_RetornaProdutoPai(objTaxaDeProducao.sProduto, sProdutoAncestral)
        If lErro <> SUCESSO Then gError 134676
        
    Else
    
        sProdutoAncestral = ""
    
    End If
        
    'Le a tabelaTaxaDeProducao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, " & _
                "UMTempo, Ativo, Data, DataDesativacao FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocMaq= ? AND NumIntDocCompet= ? AND Ativo = ? ", _
                tTaxaDeProducao.lNumIntDoc, tTaxaDeProducao.dLoteMax, tTaxaDeProducao.dLoteMin, tTaxaDeProducao.dLotePadrao, tTaxaDeProducao.dTempoPreparacao, tTaxaDeProducao.dTempoMovimentacao, tTaxaDeProducao.dTempoDescarga, _
                tTaxaDeProducao.iTipo, tTaxaDeProducao.dQuantidade, tTaxaDeProducao.sUMProduto, tTaxaDeProducao.dTempoOperacao, tTaxaDeProducao.sUMTempo, _
                tTaxaDeProducao.iAtivo, tTaxaDeProducao.dtData, tTaxaDeProducao.dtDataDesativacao, _
                sProdutoAncestral, objTaxaDeProducao.lNumIntDocMaq, objTaxaDeProducao.lNumIntDocCompet, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134597

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134598
    
    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then
    
        'não tem mais ancestral, então não vai achar mais
        If Len(sProdutoAncestral) = 0 Then gError 134599

        objTaxaDeProducao.sProduto = sProdutoAncestral
        
        'chamo recursivamente com o novo produto ...
        lErro = TaxaDeProducao_Seleciona_Maquinas(objTaxaDeProducao)
        If lErro <> SUCESSO And lErro <> 134599 Then gError 134677
        
        'não tem mais ancestrais e não achou dados
        If lErro = 134599 Then gError lErro   'gero erro com o mesmo nº devido a recursividade
    
    Else

        objTaxaDeProducao.lNumIntDoc = tTaxaDeProducao.lNumIntDoc
        objTaxaDeProducao.dLoteMin = tTaxaDeProducao.dLoteMin
        objTaxaDeProducao.dLoteMax = tTaxaDeProducao.dLoteMax
        objTaxaDeProducao.dLotePadrao = tTaxaDeProducao.dLotePadrao
        objTaxaDeProducao.dTempoPreparacao = tTaxaDeProducao.dTempoPreparacao
        objTaxaDeProducao.dTempoMovimentacao = tTaxaDeProducao.dTempoMovimentacao
        objTaxaDeProducao.dTempoDescarga = tTaxaDeProducao.dTempoDescarga
        objTaxaDeProducao.iTipo = tTaxaDeProducao.iTipo
        objTaxaDeProducao.dQuantidade = tTaxaDeProducao.dQuantidade
        objTaxaDeProducao.sUMProduto = tTaxaDeProducao.sUMProduto
        objTaxaDeProducao.dTempoOperacao = tTaxaDeProducao.dTempoOperacao
        objTaxaDeProducao.sUMTempo = tTaxaDeProducao.sUMTempo
        objTaxaDeProducao.iAtivo = tTaxaDeProducao.iAtivo
        objTaxaDeProducao.dtData = tTaxaDeProducao.dtData
        If tTaxaDeProducao.dtDataDesativacao <> DATA_NULA Then
            objTaxaDeProducao.dtDataDesativacao = tTaxaDeProducao.dtDataDesativacao
        End If
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TaxaDeProducao_Seleciona_Maquinas = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Seleciona_Maquinas:

    TaxaDeProducao_Seleciona_Maquinas = gErr

    Select Case gErr

        Case 134596
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134597, 134598
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134599 'Sem dados -> Tratado na rotina chamadora
        
        Case 134676, 134677
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152440)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function TaxaDeProducao_Seleciona_Produtos(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTaxaDeProducao As typeTaxaDeProducao

On Error GoTo Erro_TaxaDeProducao_Seleciona_Produtos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134601

    'Alocação de espaço no buffer
    tTaxaDeProducao.sProduto = String(STRING_PRODUTO, 0)
    tTaxaDeProducao.sUMProduto = String(STRING_UM_SIGLA, 0)
    tTaxaDeProducao.sUMTempo = String(STRING_UM_SIGLA, 0)

    'Le a tabelaTaxaDeProducao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, " & _
                "UMTempo, Ativo, Data, DataDesativacao FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocCompet= ? AND NumIntDocMaq = ? AND Ativo = ? ", _
                tTaxaDeProducao.lNumIntDoc, tTaxaDeProducao.dLoteMax, tTaxaDeProducao.dLoteMin, tTaxaDeProducao.dLotePadrao, tTaxaDeProducao.dTempoPreparacao, tTaxaDeProducao.dTempoMovimentacao, tTaxaDeProducao.dTempoDescarga, _
                tTaxaDeProducao.iTipo, tTaxaDeProducao.dQuantidade, tTaxaDeProducao.sUMProduto, tTaxaDeProducao.dTempoOperacao, tTaxaDeProducao.sUMTempo, _
                tTaxaDeProducao.iAtivo, tTaxaDeProducao.dtData, tTaxaDeProducao.dtDataDesativacao, _
                objTaxaDeProducao.sProduto, objTaxaDeProducao.lNumIntDocCompet, 0, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134602

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134603

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134604

    objTaxaDeProducao.lNumIntDoc = tTaxaDeProducao.lNumIntDoc
    objTaxaDeProducao.dLoteMin = tTaxaDeProducao.dLoteMin
    objTaxaDeProducao.dLoteMax = tTaxaDeProducao.dLoteMax
    objTaxaDeProducao.dLotePadrao = tTaxaDeProducao.dLotePadrao
    objTaxaDeProducao.dTempoPreparacao = tTaxaDeProducao.dTempoPreparacao
    objTaxaDeProducao.dTempoMovimentacao = tTaxaDeProducao.dTempoMovimentacao
    objTaxaDeProducao.dTempoDescarga = tTaxaDeProducao.dTempoDescarga
    objTaxaDeProducao.iTipo = tTaxaDeProducao.iTipo
    objTaxaDeProducao.dQuantidade = tTaxaDeProducao.dQuantidade
    objTaxaDeProducao.sUMProduto = tTaxaDeProducao.sUMProduto
    objTaxaDeProducao.dTempoOperacao = tTaxaDeProducao.dTempoOperacao
    objTaxaDeProducao.sUMTempo = tTaxaDeProducao.sUMTempo
    objTaxaDeProducao.iAtivo = tTaxaDeProducao.iAtivo
    objTaxaDeProducao.dtData = tTaxaDeProducao.dtData
    If tTaxaDeProducao.dtDataDesativacao <> DATA_NULA Then
        objTaxaDeProducao.dtDataDesativacao = tTaxaDeProducao.dtDataDesativacao
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TaxaDeProducao_Seleciona_Produtos = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Seleciona_Produtos:

    TaxaDeProducao_Seleciona_Produtos = gErr

    Select Case gErr

        Case 134601
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134602, 134603
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134604 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152441)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function TaxaDeProducao_Seleciona_Produtos_Anc(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTaxaDeProducao As typeTaxaDeProducao
Dim sProdutoAncestral As String

On Error GoTo Erro_TaxaDeProducao_Seleciona_Produtos_Anc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134605

    'Alocação de espaço no buffer
    tTaxaDeProducao.sProduto = String(STRING_PRODUTO, 0)
    tTaxaDeProducao.sUMProduto = String(STRING_UM_SIGLA, 0)
    tTaxaDeProducao.sUMTempo = String(STRING_UM_SIGLA, 0)
    
    If Len(objTaxaDeProducao.sProduto) <> 0 Then
        
        sProdutoAncestral = String(STRING_PRODUTO, 0)
        
        'verifica qual produto é pai deste
        lErro = Mascara_RetornaProdutoPai(objTaxaDeProducao.sProduto, sProdutoAncestral)
        If lErro <> SUCESSO Then gError 134606
        
    End If
    
    'não tem mais ancestral, então não vai achar mais
    If Len(sProdutoAncestral) = 0 Then gError 134607

    'Le a tabelaTaxaDeProducao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, " & _
                "UMTempo, Ativo, Data, DataDesativacao FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocCompet= ? AND NumIntDocMaq = ? AND Ativo = ? ", _
                tTaxaDeProducao.lNumIntDoc, tTaxaDeProducao.dLoteMax, tTaxaDeProducao.dLoteMin, tTaxaDeProducao.dLotePadrao, tTaxaDeProducao.dTempoPreparacao, tTaxaDeProducao.dTempoMovimentacao, tTaxaDeProducao.dTempoDescarga, _
                tTaxaDeProducao.iTipo, tTaxaDeProducao.dQuantidade, tTaxaDeProducao.sUMProduto, tTaxaDeProducao.dTempoOperacao, tTaxaDeProducao.sUMTempo, _
                tTaxaDeProducao.iAtivo, tTaxaDeProducao.dtData, tTaxaDeProducao.dtDataDesativacao, _
                sProdutoAncestral, objTaxaDeProducao.lNumIntDocCompet, 0, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134608

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134609

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then
    
        objTaxaDeProducao.sProduto = sProdutoAncestral
        
        'chamo recursivamente com o novo produto ...
        lErro = TaxaDeProducao_Seleciona_Produtos_Anc(objTaxaDeProducao)
        If lErro <> SUCESSO And lErro <> 134607 Then gError 134610
        
        'não tem mais ancestrais e não achou dados
        If lErro = 134607 Then gError lErro   'gero erro com o mesmo nº devido a recursividade
    
    Else
    
        objTaxaDeProducao.lNumIntDoc = tTaxaDeProducao.lNumIntDoc
        objTaxaDeProducao.dLoteMin = tTaxaDeProducao.dLoteMin
        objTaxaDeProducao.dLoteMax = tTaxaDeProducao.dLoteMax
        objTaxaDeProducao.dLotePadrao = tTaxaDeProducao.dLotePadrao
        objTaxaDeProducao.dTempoPreparacao = tTaxaDeProducao.dTempoPreparacao
        objTaxaDeProducao.dTempoMovimentacao = tTaxaDeProducao.dTempoMovimentacao
        objTaxaDeProducao.dTempoDescarga = tTaxaDeProducao.dTempoDescarga
        objTaxaDeProducao.iTipo = tTaxaDeProducao.iTipo
        objTaxaDeProducao.dQuantidade = tTaxaDeProducao.dQuantidade
        objTaxaDeProducao.sUMProduto = tTaxaDeProducao.sUMProduto
        objTaxaDeProducao.dTempoOperacao = tTaxaDeProducao.dTempoOperacao
        objTaxaDeProducao.sUMTempo = tTaxaDeProducao.sUMTempo
        objTaxaDeProducao.iAtivo = tTaxaDeProducao.iAtivo
        objTaxaDeProducao.dtData = tTaxaDeProducao.dtData
        If tTaxaDeProducao.dtDataDesativacao <> DATA_NULA Then
            objTaxaDeProducao.dtDataDesativacao = tTaxaDeProducao.dtDataDesativacao
        End If
    
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TaxaDeProducao_Seleciona_Produtos_Anc = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Seleciona_Produtos_Anc:

    TaxaDeProducao_Seleciona_Produtos_Anc = gErr

    Select Case gErr

        Case 134605
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134608, 134609
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134607 'Sem dados -> Tratado na rotina chamadora
        
        Case 134606, 134610
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152442)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function TaxaDeProducao_Seleciona_Competencia(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTaxaDeProducao As typeTaxaDeProducao

On Error GoTo Erro_TaxaDeProducao_Seleciona_Competencia

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134671

    'Alocação de espaço no buffer
    tTaxaDeProducao.sProduto = String(STRING_PRODUTO, 0)
    tTaxaDeProducao.sUMProduto = String(STRING_UM_SIGLA, 0)
    tTaxaDeProducao.sUMTempo = String(STRING_UM_SIGLA, 0)

    'Le a tabelaTaxaDeProducao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, " & _
                "UMTempo, Ativo, Data, DataDesativacao FROM TaxaDeProducao WHERE Produto= ? AND NumIntDocMaq= ? AND NumIntDocCompet= ? AND Ativo = ? ", _
                tTaxaDeProducao.lNumIntDoc, tTaxaDeProducao.dLoteMax, tTaxaDeProducao.dLoteMin, tTaxaDeProducao.dLotePadrao, tTaxaDeProducao.dTempoPreparacao, tTaxaDeProducao.dTempoMovimentacao, tTaxaDeProducao.dTempoDescarga, _
                tTaxaDeProducao.iTipo, tTaxaDeProducao.dQuantidade, tTaxaDeProducao.sUMProduto, tTaxaDeProducao.dTempoOperacao, tTaxaDeProducao.sUMTempo, _
                tTaxaDeProducao.iAtivo, tTaxaDeProducao.dtData, tTaxaDeProducao.dtDataDesativacao, _
                "", 0, objTaxaDeProducao.lNumIntDocCompet, TAXA_ATIVA)
    If lErro <> AD_SQL_SUCESSO Then gError 134672

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134673

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 134674

    objTaxaDeProducao.lNumIntDoc = tTaxaDeProducao.lNumIntDoc
    objTaxaDeProducao.dLoteMin = tTaxaDeProducao.dLoteMin
    objTaxaDeProducao.dLoteMax = tTaxaDeProducao.dLoteMax
    objTaxaDeProducao.dLotePadrao = tTaxaDeProducao.dLotePadrao
    objTaxaDeProducao.dTempoPreparacao = tTaxaDeProducao.dTempoPreparacao
    objTaxaDeProducao.dTempoMovimentacao = tTaxaDeProducao.dTempoMovimentacao
    objTaxaDeProducao.dTempoDescarga = tTaxaDeProducao.dTempoDescarga
    objTaxaDeProducao.iTipo = tTaxaDeProducao.iTipo
    objTaxaDeProducao.dQuantidade = tTaxaDeProducao.dQuantidade
    objTaxaDeProducao.sUMProduto = tTaxaDeProducao.sUMProduto
    objTaxaDeProducao.dTempoOperacao = tTaxaDeProducao.dTempoOperacao
    objTaxaDeProducao.sUMTempo = tTaxaDeProducao.sUMTempo
    objTaxaDeProducao.iAtivo = tTaxaDeProducao.iAtivo
    objTaxaDeProducao.dtData = tTaxaDeProducao.dtData
    If tTaxaDeProducao.dtDataDesativacao <> DATA_NULA Then
        objTaxaDeProducao.dtDataDesativacao = tTaxaDeProducao.dtDataDesativacao
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TaxaDeProducao_Seleciona_Competencia = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Seleciona_Competencia:

    TaxaDeProducao_Seleciona_Competencia = gErr

    Select Case gErr

        Case 134671
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134672, 134673
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 134674 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152443)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Public Function Competencias_Le_CTs(ByVal objCompetencias As ClassCompetencias, colCTsCompetencia As Collection) As Long
'Lê a Competência passada retornando numa coleção os Centros de Trabalhos associados
'Jorge Specian (13/04/2005)

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim lNumIntDocCT
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Competencias_Le_CTs

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134640

    sComandoSQL = "SELECT NumIntDocCT From CTCompetencias WHERE NumIntDocCompet = ?"
    
    'Le a tabelaCompetencias
    lErro = Comando_Executar(lComando, sComandoSQL, lNumIntDocCT, objCompetencias.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134641

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134642

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134643

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.lNumIntDoc = lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 134644

        colCTsCompetencia.Add objCentrodeTrabalho
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 134645
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Competencias_Le_CTs = SUCESSO

    Exit Function

Erro_Competencias_Le_CTs:

    Competencias_Le_CTs = gErr

    Select Case gErr

        Case 134640
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134641, 134642, 134645
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPETENCIAS", gErr)

        Case 134643 'Sem dados -> Tratado na rotina chamadora
        
        Case 134644
            'erros tratados na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152444)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Operacao_Le_Pai(ByVal lNumIntDocOper As Long, lNumIntDocPai As Long) As Long
'Lê a Tabela de Operações encontrando o NumIntDoc do Pai a partir do NumIntDoc passado
'por Jorge Specian (14/04/2005)

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lNumIntDocRotFabr As Long
Dim iSeqPai As Integer
Dim lNumIntDocOperPai As Integer
Dim objOperacoes As ClassOperacoes
Dim tOperacoes As typeOperacoes

On Error GoTo Erro_Operacao_Le_Pai

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 134646
    Next

    'Le a tabela Operacoes primeira vez para localizar dados passados
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDocRotFabr, SeqPai From Operacoes WHERE NumIntDoc = ?", _
                            lNumIntDocRotFabr, iSeqPai, lNumIntDocOper)
    If lErro <> AD_SQL_SUCESSO Then gError 134647

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134648

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134649

    'Le a tabela Operacoes segunda vez para localizar dados do pai
    lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc From Operacoes WHERE NumIntDocRotFabr = ? And Seq = ?", _
                            lNumIntDocOperPai, lNumIntDocRotFabr, iSeqPai)
    If lErro <> AD_SQL_SUCESSO Then gError 134650

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134651

    'Operacoes pai encontrada
    If lErro = AD_SQL_SUCESSO Then
    
        lNumIntDocPai = lNumIntDocOperPai
    Else
    
        lNumIntDocPai = 0
    
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Operacao_Le_Pai = SUCESSO

    Exit Function

Erro_Operacao_Le_Pai:

    Operacao_Le_Pai = gErr

    Select Case gErr

        Case 134646
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134647, 134648, 134650, 134651
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOES", gErr)

        Case 134649 'Sem dados -> Tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152445)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Tipo_Competencia(ByVal sCompetencia As String) As enumTipo

    If Len(Trim(sCompetencia)) = 0 Then
        Tipo_Competencia = TIPO_VAZIO
    ElseIf Not IsNumeric(sCompetencia) Then
        Tipo_Competencia = TIPO_STRING
    ElseIf Int(CDbl(sCompetencia)) <> CDbl(sCompetencia) Then
        Tipo_Competencia = TIPO_DECIMAL
    ElseIf CDbl(sCompetencia) <= 0 Then
        Tipo_Competencia = TIPO_NAO_POSITIVO
    ElseIf CDbl(sCompetencia) > NUM_MAX_COMPETENCIAS Then
        Tipo_Competencia = TIPO_OVERFLOW
    Else
        Tipo_Competencia = TIPO_CODIGO
    End If

End Function

Public Function TP_Competencia_Le(ByVal objCompetenciaControle As Object, ByVal objCompetencia As ClassCompetencias, Optional ByVal iCria As Integer = 1) As Long
'Lê o Competencia com Código ou NomeRed em objCompetenciaControle.Text
'Devolve em objCompetencia. Coloca NomeReduzido em .Text

Dim sCompetencia As String
Dim Competencia As Object
Dim eTipoCompetencia As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TP_Competencia_Le

    Set Competencia = objCompetenciaControle
    sCompetencia = Trim(Competencia.Text)
    eTipoCompetencia = Tipo_Competencia(sCompetencia)

    Select Case eTipoCompetencia

    Case TIPO_STRING

        objCompetencia.sNomeReduzido = sCompetencia
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencia)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 134942
        If lErro <> SUCESSO Then gError 134943

        Competencia.Text = objCompetencia.sNomeReduzido

    Case TIPO_CODIGO

        objCompetencia.lCodigo = StrParaInt(sCompetencia)
        lErro = CF("Competencias_Le", objCompetencia)
        If lErro <> SUCESSO And lErro <> 134332 Then gError 134944
        If lErro <> SUCESSO Then gError 134945

        Competencia.Text = objCompetencia.sNomeReduzido

    Case TIPO_OVERFLOW

        gError 134946

    Case TIPO_DECIMAL

        gError 134947

    Case TIPO_NAO_POSITIVO

        gError 134948

    End Select

    TP_Competencia_Le = SUCESSO

    Exit Function

Erro_TP_Competencia_Le:

    TP_Competencia_Le = gErr

    Select Case gErr

        Case 134942, 134944 'Tratados nas rotinas chamadas

        Case 134943  'Competencia com NomeReduzido não cadastrado

            If iCria = 1 Then

                'Envia aviso que Competencia não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COMPETENCIA2", objCompetencia.sNomeReduzido)

                If vbMsgRes = vbYes Then
                    'Chama tela de Competencias
                    lErro = Chama_Tela("Competencias", objCompetencia)
                End If
            End If

        Case 134945  'Competencia com Código não cadastrado

            If iCria = 1 Then

                'Envia aviso que Competencia não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_COMPETENCIA1", objCompetencia.lCodigo)

                If vbMsgRes = vbYes Then
                    'Chama tela de Competencias
                    lErro = Chama_Tela("Competencias", objCompetencia)
                End If
            End If

        Case 134946
            Call Rotina_Erro(vbOKOnly, "ERRO_LONG_OVERFLOW", gErr, sCompetencia)

        Case 134947
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sCompetencia)

        Case 134948
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sCompetencia)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152446)

    End Select

End Function

Private Function Tipo_CentrodeTrabalho(ByVal sCentrodeTrabalho As String) As enumTipo

    If Len(Trim(sCentrodeTrabalho)) = 0 Then
        Tipo_CentrodeTrabalho = TIPO_VAZIO
    ElseIf Not IsNumeric(sCentrodeTrabalho) Then
        Tipo_CentrodeTrabalho = TIPO_STRING
    ElseIf Int(CDbl(sCentrodeTrabalho)) <> CDbl(sCentrodeTrabalho) Then
        Tipo_CentrodeTrabalho = TIPO_DECIMAL
    ElseIf CDbl(sCentrodeTrabalho) <= 0 Then
        Tipo_CentrodeTrabalho = TIPO_NAO_POSITIVO
    ElseIf CDbl(sCentrodeTrabalho) > NUM_MAX_CENTROSDETRABALHOS Then
        Tipo_CentrodeTrabalho = TIPO_OVERFLOW
    Else
        Tipo_CentrodeTrabalho = TIPO_CODIGO
    End If

End Function

Public Function TP_CentrodeTrabalho_Le(ByVal objCentrodeTrabalhoControle As Object, ByVal objCentrodeTrabalho As ClassCentrodeTrabalho, Optional ByVal iCria As Integer = 1) As Long
'Lê o CentrodeTrabalho com Código ou NomeRed em objCentrodeTrabalhoControle.Text
'Devolve em objCentrodeTrabalho. Coloca NomeReduzido em .Text

Dim sCentrodeTrabalho As String
Dim CentrodeTrabalho As Object
Dim eTipoCentrodeTrabalho As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TP_CentrodeTrabalho_Le

    Set CentrodeTrabalho = objCentrodeTrabalhoControle
    sCentrodeTrabalho = Trim(CentrodeTrabalho.Text)
    eTipoCentrodeTrabalho = Tipo_CentrodeTrabalho(sCentrodeTrabalho)

    Select Case eTipoCentrodeTrabalho

    Case TIPO_STRING

        objCentrodeTrabalho.sNomeReduzido = sCentrodeTrabalho
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 134949
        If lErro <> SUCESSO Then gError 134950

        CentrodeTrabalho.Text = objCentrodeTrabalho.sNomeReduzido

    Case TIPO_CODIGO

        objCentrodeTrabalho.lCodigo = StrParaLong(sCentrodeTrabalho)
        lErro = CF("CentrodeTrabalho_Le", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134449 Then gError 134951
        If lErro <> SUCESSO Then gError 134952

        CentrodeTrabalho.Text = objCentrodeTrabalho.sNomeReduzido

    Case TIPO_OVERFLOW

        gError 134953

    Case TIPO_DECIMAL

        gError 134954

    Case TIPO_NAO_POSITIVO

        gError 134955

    End Select

    TP_CentrodeTrabalho_Le = SUCESSO

    Exit Function

Erro_TP_CentrodeTrabalho_Le:

    TP_CentrodeTrabalho_Le = gErr

    Select Case gErr

        Case 134949, 134951 'Tratados nas rotinas chamadas

        Case 134950  'CentrodeTrabalho com NomeReduzido não cadastrado
        
            If objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA Then

                Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO3", gErr, objCentrodeTrabalho.sNomeReduzido)

            Else
            
                If iCria = 1 Then
    
                    'Envia aviso que CentrodeTrabalho não está cadastrado e pergunta se deseja criar
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CENTRODETRABALHO2", objCentrodeTrabalho.sNomeReduzido)
    
                    If vbMsgRes = vbYes Then
                        'Chama tela de CentrodeTrabalhos
                        lErro = Chama_Tela("CentrodeTrabalho", objCentrodeTrabalho)
                    End If
                End If
            
            End If

        Case 134952  'CentrodeTrabalho com Código não cadastrado
        
            If objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA Then

                Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_CADASTRADO2", gErr, objCentrodeTrabalho.lCodigo)
            
            Else
            
                If iCria = 1 Then
    
                    'Envia aviso que CentrodeTrabalho não está cadastrado e pergunta se deseja criar
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CENTRODETRABALHO1", objCentrodeTrabalho.lCodigo, objCentrodeTrabalho.iFilialEmpresa)
    
                    If vbMsgRes = vbYes Then
                        'Chama tela de CentrodeTrabalhos
                        lErro = Chama_Tela("CentrodeTrabalho", objCentrodeTrabalho)
                    End If
                End If
            
            End If

        Case 134953
            Call Rotina_Erro(vbOKOnly, "ERRO_LONG_OVERFLOW", gErr, sCentrodeTrabalho)

        Case 134954
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sCentrodeTrabalho)

        Case 134955
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sCentrodeTrabalho)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152447)

    End Select

End Function

Public Function Operacoes_Le_NumIntDoc(ByVal objOperacoes As ClassOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tOperacoes As typeOperacoes

On Error GoTo Erro_Operacoes_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134969

    'Alocação de espaço no buffer
    tOperacoes.sObservacao = String(STRING_OBSERVACAO, 0)
    
    sComandoSQL = "SELECT NumIntDoc, NumIntDocRotFabr, Seq, NumIntDocCompet, NumIntDocCT, "
    sComandoSQL = sComandoSQL & "Observacao, IgnoraTaxaProducao, SeqPai, SeqArvore, Nivel, "
    sComandoSQL = sComandoSQL & "PosicaoArvore, NumMaxMaqPorOper, NumRepeticoes From Operacoes WHERE NumIntDoc = ? "

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tOperacoes.lNumIntDoc, _
                            tOperacoes.lNumIntDocRotFabr, tOperacoes.iSeq, _
                            tOperacoes.lNumIntDocCompet, tOperacoes.lNumIntDocCT, _
                            tOperacoes.sObservacao, tOperacoes.iIgnoraTaxaProducao, _
                            tOperacoes.iSeqPai, tOperacoes.iSeqArvore, tOperacoes.iNivel, _
                            tOperacoes.iPosicaoArvore, tOperacoes.iNumMaxMaqPorOper, tOperacoes.iNumRepeticoes, objOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 134970

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134971

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 134972

    objOperacoes.lNumIntDoc = tOperacoes.lNumIntDoc
    objOperacoes.lNumIntDocRotFabr = tOperacoes.lNumIntDocRotFabr
    objOperacoes.iSeq = tOperacoes.iSeq
    objOperacoes.lNumIntDocCompet = tOperacoes.lNumIntDocCompet
    objOperacoes.lNumIntDocCT = tOperacoes.lNumIntDocCT
    objOperacoes.sObservacao = tOperacoes.sObservacao
    objOperacoes.iIgnoraTaxaProducao = tOperacoes.iIgnoraTaxaProducao
    objOperacoes.iSeqPai = tOperacoes.iSeqPai
    objOperacoes.iSeqArvore = tOperacoes.iSeqArvore
    objOperacoes.iNivel = tOperacoes.iNivel
    objOperacoes.iPosicaoArvore = tOperacoes.iPosicaoArvore
    objOperacoes.iNumMaxMaqPorOper = tOperacoes.iNumMaxMaqPorOper
    objOperacoes.iNumRepeticoes = tOperacoes.iNumRepeticoes
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Operacoes_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_Operacoes_Le_NumIntDoc:

    Operacoes_Le_NumIntDoc = gErr

    Select Case gErr

        Case 134969
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134970, 134971
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPERACOES", gErr)

        Case 134972 'Sem dados -> Tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152448)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrdemDeProducao_Le_Operacoes(ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComandoAux As Long
Dim sComandoSQL As String
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim tOrdemProducaoOperacoes As typeOrdemProducaoOperacoes
Dim objOPOperacoes As New ClassOrdemProducaoOperacoes
Dim objOPInsumos As New ClassOrdemProducaoInsumos

On Error GoTo Erro_OrdemDeProducao_Le_Operacoes

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137013
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 137013
    
    lComandoAux = lComando

    'Alocação de espaço no buffer
    tOrdemProducaoOperacoes.sProduto = String(STRING_PRODUTO, 0)
    tOrdemProducaoOperacoes.sVersao = String(STRING_KIT_VERSAO, 0)
    tOrdemProducaoOperacoes.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    
    sComandoSQL = "SELECT NumIntDoc, NumIntDocItemOP, Seq, NumIntDocCompet, NumIntDocCT, "
    sComandoSQL = sComandoSQL & "Observacao, SeqPai, SeqArvore, Nivel, "
    sComandoSQL = sComandoSQL & "SeqRoteiro, SeqRoteiroPai, NivelRoteiro, Produto, Versao, "
    sComandoSQL = sComandoSQL & "IgnoraTaxaProducao, ConsideraCarga, Origem, NumIntDocOperOrigem, NumMaxMaqPorOper, NumRepeticoes "
    sComandoSQL = sComandoSQL & "FROM OrdemProducaoOperacoes WHERE NumIntDocItemOP = ? Order by Seq"

    With tOrdemProducaoOperacoes
        
        'Le a tabela Operacoes
        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .lNumIntDocItemOP, _
                                .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                                .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, _
                                .iSeqRoteiroPai, .iNivelRoteiro, .sProduto, .sVersao, _
                                .iIgnoraTaxaProducao, .iConsideraCarga, .iOrigem, _
                                .lNumIntDocOperOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes, objItemOP.lNumIntDoc)
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 137014

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137015
    
    If lErro <> AD_SQL_SUCESSO Then
    
        lComandoAux = lComando1
    
        'Alocação de espaço no buffer
        tOrdemProducaoOperacoes.sProduto = String(STRING_PRODUTO, 0)
        tOrdemProducaoOperacoes.sVersao = String(STRING_KIT_VERSAO, 0)
        tOrdemProducaoOperacoes.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
        
        sComandoSQL = "SELECT NumIntDoc, NumIntDocItemOP, Seq, NumIntDocCompet, NumIntDocCT, "
        sComandoSQL = sComandoSQL & "Observacao, SeqPai, SeqArvore, Nivel, "
        sComandoSQL = sComandoSQL & "SeqRoteiro, SeqRoteiroPai, NivelRoteiro, Produto, Versao, "
        sComandoSQL = sComandoSQL & "IgnoraTaxaProducao, ConsideraCarga, Origem, NumIntDocOperOrigem, NumMaxMaqPorOper, NumRepeticoes "
        sComandoSQL = sComandoSQL & "FROM OrdemProducaoOperacoes WHERE NumIntDoc IN ( "
        sComandoSQL = sComandoSQL & "SELECT PO.NumIntDocOper FROM PlanoOperacional AS PO, PlanoMestreProducaoItens AS PMPI "
        sComandoSQL = sComandoSQL & "WHERE PO.NumIntDocPMP  = PMPI.NumIntDoc AND PMPI.FilialEmpresa = ? AND PO.CodOPOrigem = ? AND "
        sComandoSQL = sComandoSQL & "PO.Produto = ? ) Order by Seq"
    
        With tOrdemProducaoOperacoes
            
            'Le a tabela Operacoes
            lErro = Comando_Executar(lComando1, sComandoSQL, .lNumIntDoc, .lNumIntDocItemOP, _
                                    .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                                    .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, _
                                    .iSeqRoteiroPai, .iNivelRoteiro, .sProduto, .sVersao, _
                                    .iIgnoraTaxaProducao, .iConsideraCarga, .iOrigem, _
                                    .lNumIntDocOperOrigem, .iNumMaxMaqPorOper, .iNumRepeticoes, _
                                    objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.sProduto)
            
        End With
    
        If lErro <> AD_SQL_SUCESSO Then gError 137014
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137015
    
    End If
        

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 137016

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
        
        objOrdemProducaoOperacoes.lNumIntDoc = tOrdemProducaoOperacoes.lNumIntDoc
        objOrdemProducaoOperacoes.lNumIntDocItemOP = tOrdemProducaoOperacoes.lNumIntDocItemOP
        objOrdemProducaoOperacoes.iSeq = tOrdemProducaoOperacoes.iSeq
        objOrdemProducaoOperacoes.lNumIntDocCompet = tOrdemProducaoOperacoes.lNumIntDocCompet
        objOrdemProducaoOperacoes.lNumIntDocCT = tOrdemProducaoOperacoes.lNumIntDocCT
        objOrdemProducaoOperacoes.sObservacao = tOrdemProducaoOperacoes.sObservacao
        objOrdemProducaoOperacoes.iSeqPai = tOrdemProducaoOperacoes.iSeqPai
        objOrdemProducaoOperacoes.iSeqArvore = tOrdemProducaoOperacoes.iSeqArvore
        objOrdemProducaoOperacoes.iNivel = tOrdemProducaoOperacoes.iNivel
        objOrdemProducaoOperacoes.iSeqRoteiro = tOrdemProducaoOperacoes.iSeqRoteiro
        objOrdemProducaoOperacoes.iSeqRoteiroPai = tOrdemProducaoOperacoes.iSeqRoteiroPai
        objOrdemProducaoOperacoes.iNivelRoteiro = tOrdemProducaoOperacoes.iNivelRoteiro
        objOrdemProducaoOperacoes.sProduto = tOrdemProducaoOperacoes.sProduto
        objOrdemProducaoOperacoes.sVersao = tOrdemProducaoOperacoes.sVersao
        objOrdemProducaoOperacoes.iIgnoraTaxaProducao = tOrdemProducaoOperacoes.iIgnoraTaxaProducao
        objOrdemProducaoOperacoes.iConsideraCarga = tOrdemProducaoOperacoes.iConsideraCarga
        objOrdemProducaoOperacoes.iOrigem = tOrdemProducaoOperacoes.iOrigem
        objOrdemProducaoOperacoes.lNumIntDocOperOrigem = tOrdemProducaoOperacoes.lNumIntDocOperOrigem
        objOrdemProducaoOperacoes.iNumMaxMaqPorOper = tOrdemProducaoOperacoes.iNumMaxMaqPorOper
        objOrdemProducaoOperacoes.iNumRepeticoes = tOrdemProducaoOperacoes.iNumRepeticoes
       
        'encontra a quantidade/UM do produto a ser produzido
        If tOrdemProducaoOperacoes.iSeqPai = 0 Then
        
            'se não tem pai é a quantidade/UM do próprio item da OP
            objOrdemProducaoOperacoes.dQuantidade = objItemOP.dQuantidade
            objOrdemProducaoOperacoes.sUMedida = objItemOP.sSiglaUM
        
        Else
        
            'para cada operação anterior
            For Each objOPOperacoes In objItemOP.colOrdemProducaoOperacoes
            
                'se é o pai desta operação
                If objOPOperacoes.iSeq = objOrdemProducaoOperacoes.iSeqPai Then
                    
                    'para cada insumo daquela operação pai
                    For Each objOPInsumos In objOPOperacoes.colOPInsumos
                    
                        'se é o insumo desta operação
                        If objOPInsumos.sProduto = objOrdemProducaoOperacoes.sProduto Then
                        
                            'a quantidade/UM é a deste insumo.
                            objOrdemProducaoOperacoes.dQuantidade = objOPInsumos.dQuantidade
                            objOrdemProducaoOperacoes.sUMedida = objOPInsumos.sUMProduto
                            
                            'encerra a busca por insumo
                            Exit For
                            
                        End If
                    
                    Next
                    
                    'e encerra a busca em operações anteriores
                    Exit For
                
                End If
            
            Next
        
        End If
       
        'Lê as OperacaoInsumos associadas a Operação que está sendo lida
        lErro = OrdemDeProducao_Le_OperacoesInsumos(objOrdemProducaoOperacoes)
        If lErro <> SUCESSO And lErro <> 137022 Then gError 137017
        
        '##############################
        'Inserido por Wagner 07/11/2005
        lErro = ItemOPOperacoesMaquinas_Le(objOrdemProducaoOperacoes)
        If lErro <> SUCESSO Then gError 140774
        '##############################
        
        If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
        
            'Lê as OperacaoTempo associadas a Operação que está sendo lida
            lErro = OrdemDeProducao_Le_OperacoesTempo(objOrdemProducaoOperacoes)
            If lErro <> SUCESSO And lErro <> 137060 Then gError 137043
        
        End If
        
        objItemOP.colOrdemProducaoOperacoes.Add objOrdemProducaoOperacoes
        
        'lErro = Comando_BuscarProximo(lComando)
        lErro = Comando_BuscarProximo(lComandoAux)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137018
        
    Loop

    lComandoAux = 0

    'Fecha Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    OrdemDeProducao_Le_Operacoes = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Le_Operacoes:
    
    OrdemDeProducao_Le_Operacoes = gErr
    
    Select Case gErr
    
        Case 137013
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137014, 137015, 137018
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OP_OPERACOES", gErr)

        Case 137016 'Sem dados -> Tratado na rotina chamadora
        
        Case 137017, 137043, 140774
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152449)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Public Function OrdemDeProducao_Le_OperacoesInsumos(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim tOrdemProducaoInsumos As typeOrdemProducaoInsumos

On Error GoTo Erro_OrdemDeProducao_Le_OperacoesInsumos

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137019

    'Alocação de espaço no buffer
    tOrdemProducaoInsumos.sProduto = String(STRING_PRODUTO, 0)
    tOrdemProducaoInsumos.sUMProduto = String(STRING_UM_SIGLA, 0)
    tOrdemProducaoInsumos.sVersaoKitComp = String(STRING_KIT_VERSAO, 0)
    
    sComandoSQL = "SELECT NumIntDocOper, Produto, Quantidade, UMProduto, VersaoKitComp, "
    sComandoSQL = sComandoSQL & "Composicao, PercentualPerda, CustoStandard "
    sComandoSQL = sComandoSQL & "From OrdemProducaoInsumos WHERE NumIntDocOper = ?"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tOrdemProducaoInsumos.lNumIntDocOper, _
                            tOrdemProducaoInsumos.sProduto, tOrdemProducaoInsumos.dQuantidade, _
                            tOrdemProducaoInsumos.sUMProduto, tOrdemProducaoInsumos.sVersaoKitComp, _
                            tOrdemProducaoInsumos.iComposicao, tOrdemProducaoInsumos.dPercentualPerda, _
                            tOrdemProducaoInsumos.dCustoStandard, objOrdemProducaoOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137020

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137021

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 137022

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos
        
        objOrdemProducaoInsumos.lNumIntDocOper = tOrdemProducaoInsumos.lNumIntDocOper
        objOrdemProducaoInsumos.sProduto = tOrdemProducaoInsumos.sProduto
        objOrdemProducaoInsumos.dQuantidade = tOrdemProducaoInsumos.dQuantidade
        objOrdemProducaoInsumos.sUMProduto = tOrdemProducaoInsumos.sUMProduto
        objOrdemProducaoInsumos.sVersaoKitComp = tOrdemProducaoInsumos.sVersaoKitComp
        objOrdemProducaoInsumos.iComposicao = tOrdemProducaoInsumos.iComposicao
        objOrdemProducaoInsumos.dPercentualPerda = tOrdemProducaoInsumos.dPercentualPerda
        objOrdemProducaoInsumos.dCustoStandard = tOrdemProducaoInsumos.dCustoStandard
                            
        objOrdemProducaoOperacoes.colOPInsumos.Add objOrdemProducaoInsumos
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137023
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    OrdemDeProducao_Le_OperacoesInsumos = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Le_OperacoesInsumos:

    OrdemDeProducao_Le_OperacoesInsumos = gErr

    Select Case gErr

        Case 137019
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137020, 137021, 137023
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOINSUMOS", gErr)

        Case 137022 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152450)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrdemDeProducao_Le_Roteiro(ByVal objItemOP As ClassItemOP, ByVal objOPOperacao As ClassOrdemProducaoOperacoes, ByVal objOPInsumos As ClassOrdemProducaoInsumos, ByVal colOPOperacao As Collection) As Long
Dim lErro As Long
Dim iSeqRoteiro As Integer
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objOrdemProdOper As ClassOrdemProducaoOperacoes
Dim objOrdemProdIns As ClassOrdemProducaoInsumos
Dim objProduto As ClassProduto
Dim dFatorUMProduto As Double
Dim dFatorUMOriginal As Double
Dim dFatorConvQtde As Double
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim dQuantidade As Double
Dim dTempoCorrigido As Double

On Error GoTo Erro_OrdemDeProducao_Le_Roteiro

    'Inicializa Fator de Conversão da Quantidade
    dFatorConvQtde = 1

    For Each objOrdemProdOper In objItemOP.colOrdemProducaoOperacoes
    
        If objOrdemProdOper.lNumIntDoc = objOPInsumos.lNumIntDocOper Then
            Exit For
        End If
    
    Next

    'encontra a Quantidade e UM Originais
    For Each objOrdemProdIns In objOrdemProdOper.colOPInsumos
    
        If objOrdemProdIns.sProduto = objOPInsumos.sProduto And _
            objOrdemProdIns.sVersaoKitComp = objOPInsumos.sVersaoKitComp Then
            Exit For
        End If
    
    Next
    
    If Not objOrdemProdIns Is Nothing Then
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objOPInsumos.sProduto
        
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137044
    
        'Descobre o fator de conversao da UM do obj passado p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOPInsumos.sUMProduto, objProduto.sSiglaUMEstoque, dFatorUMProduto)
        If lErro <> SUCESSO Then gError 137045
            
        'Descobre o fator de conversao da UM Original p/UM de estoque do produto
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOrdemProdIns.sUMProduto, objProduto.sSiglaUMEstoque, dFatorUMOriginal)
        If lErro <> SUCESSO Then gError 137046
    
        dFatorConvQtde = (objOPInsumos.dQuantidade * dFatorUMProduto) / (objOrdemProdIns.dQuantidade * dFatorUMOriginal)
    
    End If
    
    For Each objOrdemProdOper In objItemOP.colOrdemProducaoOperacoes
    
        If objOrdemProdOper.iNivelRoteiro = objOPOperacao.iNivelRoteiro And _
            objOrdemProdOper.sProduto = objOPOperacao.sProduto And _
            Trim(objOrdemProdOper.sVersao) = objOPOperacao.sVersao _
            And objOrdemProdOper.lNumIntDocItemOP = objOPOperacao.lNumIntDocItemOP And _
            objOrdemProdOper.iSeqPai = objOPOperacao.iSeqPai Then
            
            iSeqRoteiro = objOrdemProdOper.iSeqRoteiro
            
        End If
        
        If objOrdemProdOper.iNivelRoteiro = objOPOperacao.iNivelRoteiro And _
            objOrdemProdOper.sProduto = objOPOperacao.sProduto And _
            Trim(objOrdemProdOper.sVersao) = objOPOperacao.sVersao _
            And objOrdemProdOper.lNumIntDocItemOP = objOPOperacao.lNumIntDocItemOP And _
            objOrdemProdOper.iSeqRoteiro = iSeqRoteiro Then
            
            Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
            
            lErro = CF("OPOperacoes_Cria_Copia", objOrdemProdOper, objOrdemProducaoOperacoes)
            If lErro <> SUCESSO Then gError 137042
                        
            'Converte a quantidade de acordo com o ObjOPInsumos passado
            For Each objOrdemProducaoInsumos In objOrdemProducaoOperacoes.colOPInsumos
            
                If objOrdemProducaoInsumos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
            
                    dQuantidade = objOrdemProducaoInsumos.dQuantidade * dFatorConvQtde
                    objOrdemProducaoInsumos.dQuantidade = dQuantidade
                    
                End If
                
            Next
                        
            'Se Ignora a Taxa de Produção ...
            If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
            
                'Se for do Tipo Variável
                If objOrdemProducaoOperacoes.objOperacoesTempo.iTipo <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
                
                    'Corrige conforme a quantidade passada
                    dTempoCorrigido = objOrdemProducaoOperacoes.objOperacoesTempo.dTempoOperacao * dFatorConvQtde
                    objOrdemProducaoOperacoes.objOperacoesTempo.dTempoOperacao = dTempoCorrigido
                                                        
                End If
            
            End If
            
            'Adiciona a Operação na coleção de retorno
            colOPOperacao.Add objOrdemProducaoOperacoes

        End If
    
    Next
    
    OrdemDeProducao_Le_Roteiro = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Le_Roteiro:

    OrdemDeProducao_Le_Roteiro = gErr

    Select Case gErr
        
        Case 137042, 137044, 137045, 137046
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152451)

    End Select

    Exit Function

End Function

Function OrdemDeProducao_Le_Oper_NumIntDoc(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tOrdemProducaoOperacoes As typeOrdemProducaoOperacoes

On Error GoTo Erro_OrdemDeProducao_Le_Oper_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137036

    'Alocação de espaço no buffer
    tOrdemProducaoOperacoes.sProduto = String(STRING_PRODUTO, 0)
    tOrdemProducaoOperacoes.sVersao = String(STRING_KIT_VERSAO, 0)
    tOrdemProducaoOperacoes.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    
    sComandoSQL = "SELECT NumIntDoc, NumIntDocItemOP, Seq, NumIntDocCompet, NumIntDocCT, "
    sComandoSQL = sComandoSQL & "Observacao, SeqPai, SeqArvore, Nivel, "
    sComandoSQL = sComandoSQL & "SeqRoteiro, SeqRoteiroPai, NivelRoteiro, Produto, Versao, "
    sComandoSQL = sComandoSQL & "IgnoraTaxaProducao, ConsideraCarga, Origem, NumIntDocOperOrigem "
    sComandoSQL = sComandoSQL & "FROM OrdemProducaoOperacoes WHERE NumIntDoc = ?"

    With tOrdemProducaoOperacoes
    
        'Le a tabela Operacoes
        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .lNumIntDocItemOP, _
                                .iSeq, .lNumIntDocCompet, .lNumIntDocCT, .sObservacao, _
                                .iSeqPai, .iSeqArvore, .iNivel, .iSeqRoteiro, _
                                .iSeqRoteiroPai, .iNivelRoteiro, .sProduto, .sVersao, _
                                .iIgnoraTaxaProducao, .iConsideraCarga, .iOrigem, _
                                .lNumIntDocOperOrigem, objOrdemProducaoOperacoes.lNumIntDoc)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 137037

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137038

    'Operacoes não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 137039
    
    objOrdemProducaoOperacoes.lNumIntDoc = tOrdemProducaoOperacoes.lNumIntDoc
    objOrdemProducaoOperacoes.lNumIntDocItemOP = tOrdemProducaoOperacoes.lNumIntDocItemOP
    objOrdemProducaoOperacoes.iSeq = tOrdemProducaoOperacoes.iSeq
    objOrdemProducaoOperacoes.lNumIntDocCompet = tOrdemProducaoOperacoes.lNumIntDocCompet
    objOrdemProducaoOperacoes.lNumIntDocCT = tOrdemProducaoOperacoes.lNumIntDocCT
    objOrdemProducaoOperacoes.sObservacao = tOrdemProducaoOperacoes.sObservacao
    objOrdemProducaoOperacoes.iSeqPai = tOrdemProducaoOperacoes.iSeqPai
    objOrdemProducaoOperacoes.iSeqArvore = tOrdemProducaoOperacoes.iSeqArvore
    objOrdemProducaoOperacoes.iNivel = tOrdemProducaoOperacoes.iNivel
    objOrdemProducaoOperacoes.iSeqRoteiro = tOrdemProducaoOperacoes.iSeqRoteiro
    objOrdemProducaoOperacoes.iSeqRoteiroPai = tOrdemProducaoOperacoes.iSeqRoteiroPai
    objOrdemProducaoOperacoes.iNivelRoteiro = tOrdemProducaoOperacoes.iNivelRoteiro
    objOrdemProducaoOperacoes.sProduto = tOrdemProducaoOperacoes.sProduto
    objOrdemProducaoOperacoes.sVersao = tOrdemProducaoOperacoes.sVersao
    objOrdemProducaoOperacoes.iIgnoraTaxaProducao = tOrdemProducaoOperacoes.iIgnoraTaxaProducao
    objOrdemProducaoOperacoes.iConsideraCarga = tOrdemProducaoOperacoes.iConsideraCarga
    objOrdemProducaoOperacoes.iOrigem = tOrdemProducaoOperacoes.iOrigem
    objOrdemProducaoOperacoes.lNumIntDocOperOrigem = tOrdemProducaoOperacoes.lNumIntDocOperOrigem
    
    'Lê as OperacaoInsumos associadas a Operação que está sendo lida
    lErro = OrdemDeProducao_Le_OperacoesInsumos(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO And lErro <> 137022 Then gError 137040
    
    '##############################
    'Inserido por Wagner 07/11/2005
    lErro = ItemOPOperacoesMaquinas_Le(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO Then gError 140855
    '##############################
        
    If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
    
        'Lê as OperacaoTempo associadas a Operação que está sendo lida
        lErro = OrdemDeProducao_Le_OperacoesTempo(objOrdemProducaoOperacoes)
        If lErro <> SUCESSO And lErro <> 137060 Then gError 137071
    
    End If
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    OrdemDeProducao_Le_Oper_NumIntDoc = SUCESSO
    
    Exit Function

Erro_OrdemDeProducao_Le_Oper_NumIntDoc:
    
    OrdemDeProducao_Le_Oper_NumIntDoc = gErr
    
    Select Case gErr
    
        Case 137036
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137037, 137038, 137041
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OP_OPERACOES", gErr)

        Case 137039 'Sem dados -> Tratado na rotina chamadora
        
        Case 137040, 137071, 140855
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152452)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function OrdemDeProducao_Le_OperacoesTempo(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim objOperacoesTempo As ClassOperacoesTempo
Dim tOperacoesTempo As typeOperacoesTempo
Dim sComandoSQL As String

On Error GoTo Erro_OrdemDeProducao_Le_OperacoesTempo

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137057

    'Alocação de espaço no buffer
    tOperacoesTempo.sUMTempo = String(STRING_UM_SIGLA, 0)

    sComandoSQL = "SELECT NumIntDocOperacao, LoteMax, LoteMin, LotePadrao, "
    sComandoSQL = sComandoSQL & "TempoPreparacao, TempoMovimentacao, TempoDescarga, "
    sComandoSQL = sComandoSQL & "TempoOperacao, UMTempo, Tipo, NumIntDocMaq "
    sComandoSQL = sComandoSQL & "From OrdemProducaoTempo WHERE NumIntDocOperacao = ?"

    'Le a tabela Operacoes
    lErro = Comando_Executar(lComando, sComandoSQL, tOperacoesTempo.lNumIntDocOperacao, _
                            tOperacoesTempo.dLoteMax, tOperacoesTempo.dLoteMin, _
                            tOperacoesTempo.dLotePadrao, tOperacoesTempo.dTempoPreparacao, _
                            tOperacoesTempo.dTempoMovimentacao, tOperacoesTempo.dTempoDescarga, _
                            tOperacoesTempo.dTempoOperacao, tOperacoesTempo.sUMTempo, _
                            tOperacoesTempo.iTipo, tOperacoesTempo.lNumIntDocMaq, objOrdemProducaoOperacoes.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137058

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137059

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 137060

    Set objOperacoesTempo = New ClassOperacoesTempo
    
    objOperacoesTempo.lNumIntDocOperacao = tOperacoesTempo.lNumIntDocOperacao
    objOperacoesTempo.dLoteMax = tOperacoesTempo.dLoteMax
    objOperacoesTempo.dLoteMin = tOperacoesTempo.dLoteMin
    objOperacoesTempo.dLotePadrao = tOperacoesTempo.dLotePadrao
    objOperacoesTempo.dTempoPreparacao = tOperacoesTempo.dTempoPreparacao
    objOperacoesTempo.dTempoMovimentacao = tOperacoesTempo.dTempoMovimentacao
    objOperacoesTempo.dTempoDescarga = tOperacoesTempo.dTempoDescarga
    objOperacoesTempo.dTempoOperacao = tOperacoesTempo.dTempoOperacao
    objOperacoesTempo.sUMTempo = tOperacoesTempo.sUMTempo
    objOperacoesTempo.iTipo = tOperacoesTempo.iTipo
    objOperacoesTempo.lNumIntDocMaq = tOperacoesTempo.lNumIntDocMaq
    
    Set objOrdemProducaoOperacoes.objOperacoesTempo = objOperacoesTempo

    'Fecha Comando
    Call Comando_Fechar(lComando)

    OrdemDeProducao_Le_OperacoesTempo = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Le_OperacoesTempo:

    OrdemDeProducao_Le_OperacoesTempo = gErr

    Select Case gErr

        Case 137057
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137058, 137059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOTEMPO", gErr)

        Case 137060 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152453)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_CustoProduto(ByVal objProduto As ClassProduto, dCustoProduto As Double) As Long

Dim lErro As Long

On Error GoTo Erro_Produto_Le_CustoProduto

    dCustoProduto = 0.25

    Produto_Le_CustoProduto = SUCESSO
    
    Exit Function
    
Erro_Produto_Le_CustoProduto:

    Produto_Le_CustoProduto = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152454)

    End Select

    Exit Function
    
End Function

'########################################################################
'Inserido por Wagner

Public Function PlanoMestreProducao_Le(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPMP As typePMP

On Error GoTo Erro_PlanoMestreProducao_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136295

    'Alocação de espaço no buffer
    tPMP.sVersao = String(STRING_PMP_VERSAO, 0)

    'Le a tabelaPlanoMestreProducao
    lErro = Comando_Executar(lComando, "SELECT CodGeracao, DataGeracao, Versao FROM PlanoMestreProducaoUltimo ", _
                tPMP.lCodGeracao, tPMP.dtDataGeracao, tPMP.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 136296

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136297

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136298

    objPMP.lCodGeracao = tPMP.lCodGeracao
    objPMP.dtDataGeracao = tPMP.dtDataGeracao
    objPMP.sVersao = tPMP.sVersao

    lErro = PlanoMestreProducaoItens_Le(objPMP)
    If lErro <> SUCESSO And lErro <> 136303 And lErro <> 136304 Then gError 136299

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoMestreProducao_Le = SUCESSO

    Exit Function

Erro_PlanoMestreProducao_Le:

    PlanoMestreProducao_Le = gErr

    Select Case gErr

        Case 136295
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136296, 136297
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAO", gErr)

        Case 136298 'Sem dados -> Tratado na rotina chamadora
        
        Case 136299 'Tratado na Função Chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152455)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PlanoMestreProducao_Le2(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPMP As typePMP

On Error GoTo Erro_PlanoMestreProducao_Le2

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136470

    'Alocação de espaço no buffer
    tPMP.sVersao = String(STRING_PMP_VERSAO, 0)

    'Le a tabelaPlanoMestreProducao
    lErro = Comando_Executar(lComando, "SELECT CodGeracao, DataGeracao, Versao FROM PlanoMestreProducao WHERE Versao = ? ", _
                tPMP.lCodGeracao, tPMP.dtDataGeracao, tPMP.sVersao, objPMP.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 136471

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136472

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136473

    objPMP.lCodGeracao = tPMP.lCodGeracao
    objPMP.dtDataGeracao = tPMP.dtDataGeracao
    objPMP.sVersao = tPMP.sVersao

    lErro = PlanoMestreProducaoItens_Le(objPMP)
    If lErro <> SUCESSO And lErro <> 136303 And lErro <> 136304 Then gError 136474

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoMestreProducao_Le2 = SUCESSO

    Exit Function

Erro_PlanoMestreProducao_Le2:

    PlanoMestreProducao_Le2 = gErr

    Select Case gErr

        Case 136470
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136471, 136472
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAO", gErr)

        Case 136473 'Sem dados -> Tratado na rotina chamadora
        
        Case 136474 'Tratado na Função Chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152456)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PMPCombo_Le_Versoes(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPMP As typePMP

On Error GoTo Erro_PMPCombo_Le_Versoes

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136463

    objCombo.Clear
    
    'Alocação de espaço no buffer
    tPMP.sVersao = String(STRING_PMP_VERSAO, 0)

    'Le a tabelaPlanoMestreProducao
    lErro = Comando_Executar(lComando, "SELECT Versao, Codgeracao FROM PlanoMestreProducao WHERE Versao <> ? ORDER BY Versao ", tPMP.sVersao, tPMP.lCodGeracao, "")
    If lErro <> AD_SQL_SUCESSO Then gError 136464

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136465

    'Com Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        objCombo.AddItem tPMP.sVersao
        objCombo.ItemData(objCombo.NewIndex) = tPMP.lCodGeracao

        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136466

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PMPCombo_Le_Versoes = SUCESSO

    Exit Function

Erro_PMPCombo_Le_Versoes:

    PMPCombo_Le_Versoes = gErr

    Select Case gErr

        Case 13463
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136464 To 136466
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152457)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PlanoMestreProducaoItens_Le(ByVal objPMP As ClassPMP) As Long

Dim lErro As Long
Dim lComando As Long
Dim objPMPItens As ClassPMPItens
Dim tPMPItens As typePMPItens
Dim objItemOP As ClassItemOP
Dim objOP As ClassOrdemDeProducao

On Error GoTo Erro_PlanoMestreProducaoItens_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136300

    'Alocação de espaço no buffer
    tPMPItens.sProduto = String(STRING_PRODUTO, 0)
    tPMPItens.sVersao = String(STRING_KIT_VERSAO, 0)
    tPMPItens.sUM = String(STRING_UM_SIGLA, 0)
    tPMPItens.sCodOPOrigem = String(STRING_OPCODIGO, 0)

    'Le a tabelaPlanoMestreProducaoItens com OP Ativa
    lErro = Comando_Executar(lComando, "SELECT PMPI.NumIntDoc, PMPI.CodGeracao, PMPI.Produto, PMPI.Versao, PMPI.Quantidade, PMPI.UM, " & _
                "PMPI.DataNecessidade, PMPI.CodOPOrigem, PMPI.Cliente, PMPI.FilialCliente, PMPI.FilialEmpresa, PMPI.Prioridade FROM PlanoMestreProducaoItens AS PMPI, OrdensDeProducao AS OP WHERE OP.Codigo =PMPI.CodOPOrigem AND OP.FilialEmpresa = PMPI.FilialEmpresa AND CodGeracao = ? ", _
                tPMPItens.lNumIntDoc, tPMPItens.lCodGeracao, tPMPItens.sProduto, tPMPItens.sVersao, _
                tPMPItens.dQuantidade, tPMPItens.sUM, tPMPItens.dtDataNecessidade, tPMPItens.sCodOPOrigem, _
                tPMPItens.lCliente, tPMPItens.iFilialCli, tPMPItens.iFilialEmpresa, tPMPItens.iPrioridade, _
                objPMP.lCodGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 136301

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136302

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136303

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPMPItens = New ClassPMPItens
        Set objItemOP = New ClassItemOP
        Set objOP = New ClassOrdemDeProducao
    
        objPMPItens.lNumIntDoc = tPMPItens.lNumIntDoc
        objPMPItens.lCodGeracao = tPMPItens.lCodGeracao
        objPMPItens.sProduto = tPMPItens.sProduto
        objPMPItens.sVersao = tPMPItens.sVersao
        objPMPItens.dQuantidade = tPMPItens.dQuantidade
        objPMPItens.sUM = tPMPItens.sUM
        objPMPItens.dtDataNecessidade = tPMPItens.dtDataNecessidade
        objPMPItens.sCodOPOrigem = tPMPItens.sCodOPOrigem
        objPMPItens.lCliente = tPMPItens.lCliente
        objPMPItens.iFilialCli = tPMPItens.iFilialCli
        objPMPItens.iFilialEmpresa = tPMPItens.iFilialEmpresa
        objPMPItens.iPrioridade = tPMPItens.iPrioridade
        
        objOP.sCodigo = objPMPItens.sCodOPOrigem
        objOP.iFilialEmpresa = objPMPItens.iFilialEmpresa
        
        'tenta ler a OP desejada
        lErro = CF("OrdemProducao_Le", objOP)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 136498
    
        'Le a ordem de produção associada ao PMP
        'Le as demais informações associadas a Ordem de Produção
        lErro = CF("OrdemDeProducao_Le_ComItens", objOP)
        If lErro <> SUCESSO And lErro <> 21960 Then gError 136499
        
        If lErro <> SUCESSO Then
            lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOP)
            If lErro <> SUCESSO And lErro <> 82797 Then gError 138971
        End If
        
        'Encontra o item da OP relacionada ao item do Plano Mestre
        For Each objItemOP In objOP.colItens
            If objItemOP.sProduto = objPMPItens.sProduto Then Exit For
        Next
        
        Set objPMPItens.objItemOP = objItemOP
        Set objPMPItens.objOP = objOP
        
        'Le as etapas
        lErro = PlanoOperacional_Le(objPMPItens)
        If lErro <> SUCESSO And lErro <> 136309 And lErro <> 136310 Then gError 136304
        
        'Le a necessidade de compras
        lErro = PlanoOperacional_Le_RC(objPMPItens)
        If lErro <> SUCESSO Then gError 136501
       
        'Preenche a necessidade de produção
        lErro = PlanoOperacional_Le_OP(objPMPItens)
        If lErro <> SUCESSO Then gError 136513
       
        objPMP.colItens.Add objPMPItens

        'Busca Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136305

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoMestreProducaoItens_Le = SUCESSO

    Exit Function

Erro_PlanoMestreProducaoItens_Le:

    PlanoMestreProducaoItens_Le = gErr

    Select Case gErr

        Case 136300
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136301, 136302, 136305
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAOITENS", gErr)

        Case 136303 'Sem dados -> Tratado na rotina chamadora
        
        Case 136304, 136498, 136499, 136501, 136513 'Tratado ma função chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152458)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PlanoOperacional_Le_RC(ByVal objPMPItem As ClassPMPItens) As Long

Dim lErro As Long
Dim lComando As Long
Dim sProduto As String
Dim sUM As String
Dim dQuantidade As Double
Dim objPO As ClassPlanoOperacional
Dim objPOAux As ClassPlanoOperacional
Dim objReqCompras As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim iAlmoxarifadoPadrao As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_PlanoOperacional_Le_RC

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136502

    For Each objPO In objPMPItem.colPO

        'Alocação de espaço no buffer
        sProduto = String(STRING_PRODUTO, 0)
        sUM = String(STRING_UM_SIGLA * 2, 0)

        'Le a PlanoOperacionalReqCompras
        lErro = Comando_Executar(lComando, "SELECT Produto, Quantidade, UM FROM PlanoOperacionalReqCompras WHERE NumIntDocPO = ?", sProduto, dQuantidade, sUM, objPO.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 136503
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136504
    
        Do While lErro <> AD_SQL_SEM_DADOS

            Set objReqCompras = New ClassRequisicaoCompras
            Set objItemRC = New ClassItemReqCompras
            Set objProduto = New ClassProduto
                                    
            objProduto.sCodigo = sProduto
                                    
            'Le o produto para obter a ClasseUM
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 136505
            
            If lErro <> SUCESSO Then gError 136506
            
            'Cria um ItemRC novo com o produto
            lErro = CF("AlmoxarifadoPadrao_Le", objPO.iFilialEmpresa, sProduto, iAlmoxarifadoPadrao)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 136507
                            
            objReqCompras.dtDataLimite = objPO.dtDataInicio
            objReqCompras.dtData = gdtDataAtual
            objReqCompras.dtDataBaixa = DATA_NULA
            objReqCompras.dtDataEnvio = DATA_NULA
            objReqCompras.iFilialCompra = objPO.iFilialEmpresa
            objReqCompras.iFilialDestino = objPO.iFilialEmpresa
            objReqCompras.iFilialEmpresa = objPO.iFilialEmpresa
            objReqCompras.iTipoDestino = TIPO_DESTINO_EMPRESA
            objReqCompras.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO
            objReqCompras.sOPCodigo = objPO.sCodOPOrigem
    
            objItemRC.iItem = 1
            objItemRC.dQuantidade = dQuantidade
            objItemRC.iAlmoxarifado = iAlmoxarifadoPadrao
            objItemRC.sDescProduto = objProduto.sDescricao
            objItemRC.sProduto = sProduto
            objItemRC.sUM = objProduto.sSiglaUMCompra

            objReqCompras.colItens.Add objItemRC
        
            objPO.colRCFilhas.Add objReqCompras

            'Busca Próximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136508

        Loop

    Next

    PlanoOperacional_Le_RC = SUCESSO

    Exit Function

Erro_PlanoOperacional_Le_RC:

    PlanoOperacional_Le_RC = gErr

    Select Case gErr

        Case 136502
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136503, 136504, 136508
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_POREQCOMPRAS", gErr)

        Case 136505, 136507

        Case 136506
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152459)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PlanoOperacional_Le_OP(ByVal objPMPItem As ClassPMPItens) As Long

Dim lErro As Long
Dim objPO As ClassPlanoOperacional
Dim objPOFilha As ClassPlanoOperacional
Dim objOP As ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim iAlmoxarifadoPadrao As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_PlanoOperacional_Le_OP

    For Each objPO In objPMPItem.colPO

        For Each objPOFilha In objPMPItem.colPO
    
            'Se a etapa é filha , então deve virar uma SubOP
            If objPOFilha.lNumIntDocPOPai = objPO.lNumIntDoc And objPOFilha.sCodOPOrigem <> objPO.sCodOPOrigem Then
                
                Set objOP = New ClassOrdemDeProducao
                Set objItemOP = New ClassItemOP
                Set objProduto = New ClassProduto
                                        
                objProduto.sCodigo = objPOFilha.sProduto
                                        
                'Le o produto para obter a ClasseUM
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 136509
                
                If lErro <> SUCESSO Then gError 136510
                
                objOP.sCodigo = objPOFilha.sCodOPOrigem
                objOP.iFilialEmpresa = objPOFilha.iFilialEmpresa
                objOP.sOPGeradora = objPO.sCodOPOrigem
                objOP.dtDataEmissao = gdtDataAtual
                
                lErro = CF("AlmoxarifadoPadrao_Le", objPO.iFilialEmpresa, objPOFilha.sProduto, iAlmoxarifadoPadrao)
                If lErro <> SUCESSO And lErro <> 23796 Then gError 136511
                If iAlmoxarifadoPadrao = 0 Then gError 136512
                
                objItemOP.iItem = 1
                objItemOP.dQuantidade = objPOFilha.dQuantidade
                objItemOP.dtDataFimProd = objPOFilha.dtDataFim
                objItemOP.iFilialEmpresa = objPOFilha.iFilialEmpresa
                objItemOP.iAlmoxarifado = iAlmoxarifadoPadrao
                objItemOP.iDestinacao = ITEMOP_DESTINACAO_ESTOQUE
                objItemOP.sProduto = objPOFilha.sProduto
                objItemOP.sSiglaUM = objPOFilha.sUM
                objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
                objItemOP.sVersao = objPOFilha.sVersao
                objItemOP.sCodigo = objOP.sCodigo
                
                objOP.colItens.Add objItemOP
            
                objOP.iNumItens = objOP.colItens.Count
                
                objPO.colOPFilhas.Add objOP
            
            End If
    
        Next

    Next

    PlanoOperacional_Le_OP = SUCESSO

    Exit Function

Erro_PlanoOperacional_Le_OP:

    PlanoOperacional_Le_OP = gErr

    Select Case gErr
    
        Case 136509, 136511

        Case 136510
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
               
        Case 136512
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152460)

    End Select

    Exit Function

End Function

Public Function PlanoOperacional_Le(ByVal objPMPItens As ClassPMPItens) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPO As typePlanoOperacional
Dim objPO As ClassPlanoOperacional

On Error GoTo Erro_PlanoOperacional_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136306

    'Alocação de espaço no buffer
    tPO.sCodOPOrigem = String(STRING_OPCODIGO, 0)
    tPO.sProduto = String(STRING_PRODUTO, 0)
    tPO.sVersao = String(STRING_KIT_VERSAO, 0)
    tPO.sUM = String(STRING_UM_SIGLA, 0)

    'Le a tabelaPlanoOperacional
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPOPai, NumIntDocPMP, NumIntDocOper, Nivel, Seq, " & _
                "CodOPOrigem, Produto, Versao, Quantidade, UM, NumIntDocCT, DataInicio, " & _
                "DataFim FROM PlanoOperacional WHERE NumIntDocPMP = ? ", _
                tPO.lNumIntDoc, tPO.lNumIntDocPOPai, tPO.lNumIntDocPMP, tPO.lNumIntDocOper, _
                tPO.iNivel, tPO.iSeq, tPO.sCodOPOrigem, tPO.sProduto, tPO.sVersao, _
                tPO.dQuantidade, tPO.sUM, tPO.lNumIntDocCT, tPO.dtDataInicio, _
                tPO.dtDataFim, _
                objPMPItens.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 136307

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136308

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136309
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objPO = New ClassPlanoOperacional

        objPO.lNumIntDoc = tPO.lNumIntDoc
        objPO.lNumIntDocPOPai = tPO.lNumIntDocPOPai
        objPO.lNumIntDocPMP = tPO.lNumIntDocPMP
        objPO.lNumIntDocOper = tPO.lNumIntDocOper
        objPO.iNivel = tPO.iNivel
        objPO.iSeq = tPO.iSeq
        objPO.sCodOPOrigem = tPO.sCodOPOrigem
        objPO.sProduto = tPO.sProduto
        objPO.sVersao = tPO.sVersao
        objPO.dQuantidade = tPO.dQuantidade
        objPO.sUM = tPO.sUM
        objPO.lNumIntDocCT = tPO.lNumIntDocCT
        objPO.dtDataInicio = tPO.dtDataInicio
        objPO.dtDataFim = tPO.dtDataFim
        objPO.iFilialEmpresa = objPMPItens.iFilialEmpresa
        objPO.iAlterado = REGISTRO_ALTERADO 'Para pode validar sempre que lê (o dado pode estar desatualizado)
               
        Set objPO.objOP = New ClassOrdemDeProducao
               
        objPO.objOP.sCodigo = objPO.sCodOPOrigem
        objPO.objOP.iFilialEmpresa = objPO.iFilialEmpresa
               
        'tenta ler a OP desejada
        lErro = CF("OrdemProducao_Le", objPO.objOP)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 136498
    
        'Le a ordem de produção associada ao PMP
        'Le as demais informações associadas a Ordem de Produção
        lErro = CF("OrdemDeProducao_Le_ComItens", objPO.objOP)
        If lErro <> SUCESSO And lErro <> 21960 Then gError 136499
        
        If lErro <> SUCESSO Then
            lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objPO.objOP)
            If lErro <> SUCESSO And lErro <> 82797 Then gError 138971
        End If
        
        lErro = PlanoOperacionalMaquinas_Le(objPO)
        If lErro <> SUCESSO And lErro <> 136315 Then gError 136310
        
        objPMPItens.colPO.Add objPO

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136311

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoOperacional_Le = SUCESSO

    Exit Function

Erro_PlanoOperacional_Le:

    PlanoOperacional_Le = gErr

    Select Case gErr

        Case 136306
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136307, 136308, 136311
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOOPERACIONAL", gErr)

        Case 136309 'Sem dados -> Tratado na rotina chamadora
        
        Case 136310, 136498, 136499, 138971 'Tratado na função chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152461)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function PlanoOperacionalMaquinas_Le(ByVal objPO As ClassPlanoOperacional) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPOMaquinas As typePOMaquinas
Dim objPOMaquinas As ClassPOMaquinas
Dim objTaxa As ClassTaxaDeProducao
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_PlanoOperacionalMaquinas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136312

    'Le a tabelaPlanoOperacionalMaquinas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocPO, NumIntDocMaq, Quantidade, Data, HorasMaquinas, " & _
                "NumIntDocTxProd FROM PlanoOperacionalMaquinas WHERE NumIntDocPO = ?", _
                tPOMaquinas.lNumIntDoc, tPOMaquinas.lNumIntDocPO, tPOMaquinas.lNumIntDocMaq, tPOMaquinas.iQuantidade, _
                tPOMaquinas.dtData, tPOMaquinas.dHorasMaquina, tPOMaquinas.lNumIntDocTxProd, _
                objPO.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 136313

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136314

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136315

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPOMaquinas = New ClassPOMaquinas
        Set objTaxa = New ClassTaxaDeProducao
        Set objMaquina = New ClassMaquinas

        objPOMaquinas.lNumIntDoc = tPOMaquinas.lNumIntDoc
        objPOMaquinas.lNumIntDocPO = tPOMaquinas.lNumIntDocPO
        objPOMaquinas.lNumIntDocMaq = tPOMaquinas.lNumIntDocMaq
        objPOMaquinas.iQuantidade = tPOMaquinas.iQuantidade
        objPOMaquinas.dtData = tPOMaquinas.dtData
        objPOMaquinas.dHorasMaquina = tPOMaquinas.dHorasMaquina
        objPOMaquinas.lNumIntDocTxProd = tPOMaquinas.lNumIntDocTxProd
        
        objTaxa.sProduto = objPO.sProduto
        objTaxa.lNumIntDocMaq = objPOMaquinas.lNumIntDocMaq
         
        objTaxa.lNumIntDoc = objPOMaquinas.lNumIntDocTxProd
        
        lErro = TaxaDeProducao_Le_NumIntDoc(objTaxa)
        If lErro <> SUCESSO And lErro <> 137151 Then gError 136500
        
        Set objPOMaquinas.objTaxaProducao = objTaxa
        
        objMaquina.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
        
        'Le a Máquina
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 138519
        
        lErro = CF("Maquinas_Le_Itens", objMaquina)
        If lErro <> SUCESSO Then gError 138520
        
        Set objPOMaquinas.objMaquina = objMaquina
        
        objPO.colAlocacaoMaquinas.Add objPOMaquinas

        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136316

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoOperacionalMaquinas_Le = SUCESSO

    Exit Function

Erro_PlanoOperacionalMaquinas_Le:

    PlanoOperacionalMaquinas_Le = gErr

    Select Case gErr

        Case 136312
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136313, 136314, 136316
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOOPERACIONALMAQUINAS", gErr)

        Case 136315 'Sem dados -> Tratado na rotina chamadora

        Case 136500, 138519, 138520

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152462)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'########################################################################

Public Function PlanoMestreProducao_Le_ItensOP(ByVal objPMP As ClassPMP, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tPMP As typePMP
Dim objPMPItens As ClassPMPItens
Dim tPMPItens As typePMPItens

On Error GoTo Erro_PlanoMestreProducao_Le_ItensOP

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137137
    
    'Le a tabelaPlanoMestreProducao
    lErro = Comando_Executar(lComando, "SELECT CodGeracao, DataGeracao FROM PlanoMestreProducaoUltimo ", _
                tPMP.lCodGeracao, tPMP.dtDataGeracao)
    If lErro <> AD_SQL_SUCESSO Then gError 137138
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137139

    'Se tiver Dados
    If lErro <> AD_SQL_SEM_DADOS Then
    
        objPMP.lCodGeracao = tPMP.lCodGeracao
        objPMP.dtDataGeracao = tPMP.dtDataGeracao
        
        'Alocação de espaço no buffer
        tPMPItens.sProduto = String(STRING_PRODUTO, 0)
        tPMPItens.sVersao = String(STRING_KIT_VERSAO, 0)
        tPMPItens.sUM = String(STRING_UM_SIGLA, 0)
        tPMPItens.sCodOPOrigem = String(STRING_OPCODIGO, 0)
    
        'Le a tabelaPlanoMestreProducaoItens
        lErro = Comando_Executar(lComando, "SELECT PMPItens.NumIntDoc, PMPItens.CodGeracao, PMPItens.Produto, " & _
                    "PMPItens.Versao, PMPItens.Quantidade, PMPItens.UM, PMPItens.DataNecessidade, PMPItens.CodOPOrigem, " & _
                    "PMPItens.Cliente, PMPItens.FilialCliente, PMPItens.FilialEmpresa, PMPItens.Prioridade " & _
                    "FROM PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMPItens.NumIntDoc = PO.NumIntDocPMP " & _
                    "AND PMPItens.CodGeracao = ? AND PMPItens.FilialEmpresa = ? AND PO.CodOPOrigem = ? GROUP BY PMPItens.NumIntDoc, PMPItens.CodGeracao, PMPItens.Produto, " & _
                    "PMPItens.Versao, PMPItens.Quantidade, PMPItens.UM, PMPItens.DataNecessidade, PMPItens.CodOPOrigem, " & _
                    "PMPItens.Cliente, PMPItens.FilialCliente, PMPItens.FilialEmpresa, PMPItens.Prioridade ", _
                    tPMPItens.lNumIntDoc, tPMPItens.lCodGeracao, tPMPItens.sProduto, tPMPItens.sVersao, _
                    tPMPItens.dQuantidade, tPMPItens.sUM, tPMPItens.dtDataNecessidade, tPMPItens.sCodOPOrigem, _
                    tPMPItens.lCliente, tPMPItens.iFilialCli, tPMPItens.iFilialEmpresa, tPMPItens.iPrioridade, _
                    objPMP.lCodGeracao, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
                
        If lErro <> AD_SQL_SUCESSO Then gError 137140
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137141
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objPMPItens = New ClassPMPItens
        
            objPMPItens.lNumIntDoc = tPMPItens.lNumIntDoc
            objPMPItens.lCodGeracao = tPMPItens.lCodGeracao
            objPMPItens.sProduto = tPMPItens.sProduto
            objPMPItens.sVersao = tPMPItens.sVersao
            objPMPItens.dQuantidade = tPMPItens.dQuantidade
            objPMPItens.sUM = tPMPItens.sUM
            objPMPItens.dtDataNecessidade = tPMPItens.dtDataNecessidade
            objPMPItens.sCodOPOrigem = tPMPItens.sCodOPOrigem
            objPMPItens.lCliente = tPMPItens.lCliente
            objPMPItens.iFilialCli = tPMPItens.iFilialCli
            objPMPItens.iFilialEmpresa = tPMPItens.iFilialEmpresa
            objPMPItens.iPrioridade = tPMPItens.iPrioridade
            
            lErro = PlanoOperacional_Le(objPMPItens)
            If lErro <> SUCESSO And lErro <> 136309 And lErro <> 136310 Then gError 137142
            
            objPMP.colItens.Add objPMPItens
    
            'Busca Proximo
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137143
    
        Loop
        
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PlanoMestreProducao_Le_ItensOP = SUCESSO

    Exit Function

Erro_PlanoMestreProducao_Le_ItensOP:

    PlanoMestreProducao_Le_ItensOP = gErr

    Select Case gErr

        Case 137137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137138, 137139
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAO", gErr)

        Case 137140, 137141, 137143
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOMESTREPRODUCAOITENS", gErr)

        Case 137142 'Tratado na função chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152463)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function TaxaDeProducao_Le_NumIntDoc(ByVal objTaxaDeProducao As ClassTaxaDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTaxaDeProducao As typeTaxaDeProducao

On Error GoTo Erro_TaxaDeProducao_Le_NumIntDoc

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137148

    'Alocação de espaço no buffer
    tTaxaDeProducao.sProduto = String(STRING_PRODUTO, 0)
    tTaxaDeProducao.sUMProduto = String(STRING_UM_SIGLA, 0)
    tTaxaDeProducao.sUMTempo = String(STRING_UM_SIGLA, 0)

    'Le a tabelaTaxaDeProducao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, LoteMax, LoteMin, LotePadrao, TempoPreparacao, TempoMovimentacao, TempoDescarga, Tipo, Quantidade, UMProduto, TempoOperacao, " & _
                "UMTempo, Ativo, Data, DataDesativacao FROM TaxaDeProducao WHERE NumIntDoc= ? ", _
                tTaxaDeProducao.lNumIntDoc, tTaxaDeProducao.dLoteMax, tTaxaDeProducao.dLoteMin, tTaxaDeProducao.dLotePadrao, tTaxaDeProducao.dTempoPreparacao, tTaxaDeProducao.dTempoMovimentacao, tTaxaDeProducao.dTempoDescarga, _
                tTaxaDeProducao.iTipo, tTaxaDeProducao.dQuantidade, tTaxaDeProducao.sUMProduto, tTaxaDeProducao.dTempoOperacao, tTaxaDeProducao.sUMTempo, tTaxaDeProducao.iAtivo, tTaxaDeProducao.dtData, tTaxaDeProducao.dtDataDesativacao, _
                objTaxaDeProducao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 137149

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137150

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 137151

    objTaxaDeProducao.lNumIntDoc = tTaxaDeProducao.lNumIntDoc
    objTaxaDeProducao.dLoteMin = tTaxaDeProducao.dLoteMin
    objTaxaDeProducao.dLoteMax = tTaxaDeProducao.dLoteMax
    objTaxaDeProducao.dLotePadrao = tTaxaDeProducao.dLotePadrao
    objTaxaDeProducao.dTempoPreparacao = tTaxaDeProducao.dTempoPreparacao
    objTaxaDeProducao.dTempoMovimentacao = tTaxaDeProducao.dTempoMovimentacao
    objTaxaDeProducao.dTempoDescarga = tTaxaDeProducao.dTempoDescarga
    objTaxaDeProducao.iTipo = tTaxaDeProducao.iTipo
    objTaxaDeProducao.dQuantidade = tTaxaDeProducao.dQuantidade
    objTaxaDeProducao.sUMProduto = tTaxaDeProducao.sUMProduto
    objTaxaDeProducao.dTempoOperacao = tTaxaDeProducao.dTempoOperacao
    objTaxaDeProducao.sUMTempo = tTaxaDeProducao.sUMTempo
    objTaxaDeProducao.iAtivo = tTaxaDeProducao.iAtivo
    objTaxaDeProducao.dtData = tTaxaDeProducao.dtData
    If tTaxaDeProducao.dtDataDesativacao <> DATA_NULA Then
        objTaxaDeProducao.dtDataDesativacao = tTaxaDeProducao.dtDataDesativacao
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TaxaDeProducao_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_TaxaDeProducao_Le_NumIntDoc:

    TaxaDeProducao_Le_NumIntDoc = gErr

    Select Case gErr

        Case 137148
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137149, 137150
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TAXADEPRODUCAO", gErr)

        Case 137151 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152464)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function CTMaquinaProgDisponibilidade_Le(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCTMaqProgDisp As typeCTMaqProgDisp

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136561

    'Alocação de espaço no buffer
    tCTMaqProgDisp.sObservacao = String(STRING_OBSERVACAO, 0)

    'Le a tabelaCTMaquinaProgDisponibilidade
    lErro = Comando_Executar(lComando, "SELECT C.NumIntDoc, C.NumIntDocCT, C.NumIntDocMaq, C.Data, C.Quantidade, C.Observacao FROM CTMaquinaProgDisponibilidade AS C WHERE C.Data = ? AND C.NumIntDocCT= ? AND C.NumIntDocMaq= ? ", _
                tCTMaqProgDisp.lNumIntDoc, tCTMaqProgDisp.lNumIntDocCT, tCTMaqProgDisp.lNumIntDocMaq, tCTMaqProgDisp.dtData, tCTMaqProgDisp.iQuantidade, _
                tCTMaqProgDisp.sObservacao, _
                objCTMaqProgDisp.dtData, objCTMaqProgDisp.lNumIntDocCT, objCTMaqProgDisp.lNumIntDocMaq)
    If lErro <> AD_SQL_SUCESSO Then gError 136562

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136563

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136564

    objCTMaqProgDisp.lNumIntDoc = tCTMaqProgDisp.lNumIntDoc
    objCTMaqProgDisp.lNumIntDocCT = tCTMaqProgDisp.lNumIntDocCT
    objCTMaqProgDisp.lNumIntDocMaq = tCTMaqProgDisp.lNumIntDocMaq
    objCTMaqProgDisp.dtData = tCTMaqProgDisp.dtData
    objCTMaqProgDisp.iQuantidade = tCTMaqProgDisp.iQuantidade
    objCTMaqProgDisp.sObservacao = tCTMaqProgDisp.sObservacao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CTMaquinaProgDisponibilidade_Le = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Le:

    CTMaquinaProgDisponibilidade_Le = gErr

    Select Case gErr

        Case 136561
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136562, 136563
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQPROGDISP", gErr)

        Case 136564 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152465)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CTMaquinaProgDisponibilidade_Le_Periodo(ByVal objCTMaqProgDisp As ClassCTMaqProgDisp, colCTMaqProgDisp As Collection) As Long

Dim lErro As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iPeriodoEmDias As Long
Dim iIndice As Integer
Dim objCTMaqProgDispUnitario As ClassCTMaqProgDisp

On Error GoTo Erro_CTMaquinaProgDisponibilidade_Le_Periodo

    dtDataInicial = objCTMaqProgDisp.dtDataDe
    dtDataFinal = objCTMaqProgDisp.dtDataAte
    
    'Se a data inicial não está vazia ...
    If Len(dtDataInicial) > 0 Then
        
        'Se a data final está vazia ... assume a data inicial como final
        If Len(dtDataFinal) = 0 Then
    
            dtDataFinal = dtDataInicial
            
        End If
                        
        'verifica qual o período
        iPeriodoEmDias = CInt(DateDiff("d", dtDataInicial, dtDataFinal))
        
        'para cada dia do período passado
        For iIndice = 0 To iPeriodoEmDias
        
            Set objCTMaqProgDispUnitario = New ClassCTMaqProgDisp
            
            With objCTMaqProgDispUnitario
            
                .lNumIntDocCT = objCTMaqProgDisp.lNumIntDocCT
                .lNumIntDocMaq = objCTMaqProgDisp.lNumIntDocMaq
                
                'se é a primeira data... não implementa nada
                If iIndice = 0 Then
                    .dtData = dtDataInicial
                Else
                    .dtData = DateAdd("d", iIndice, dtDataInicial)
                End If
            
            End With
            
            'Lê CTMaquinaProgDisponibilidade Individualmente
            lErro = CF("CTMaquinaProgDisponibilidade_Le", objCTMaqProgDispUnitario)
            If lErro <> SUCESSO And lErro <> 136564 Then gError 137261
            
            'se conseguiu ler...
            If lErro = SUCESSO Then
            
                'alimenta a coleção
                colCTMaqProgDisp.Add objCTMaqProgDispUnitario
            
            End If
            
        Next
        
    End If

    CTMaquinaProgDisponibilidade_Le_Periodo = SUCESSO

    Exit Function

Erro_CTMaquinaProgDisponibilidade_Le_Periodo:

    CTMaquinaProgDisponibilidade_Le_Periodo = gErr

    Select Case gErr

        Case 137261
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152466)

    End Select

    Exit Function

End Function

Public Function CTMaquinaProgTurno_Le_Periodo(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno, colCTMaqProgTurno As Collection) As Long

Dim lErro As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iPeriodoEmDias As Long
Dim iIndice As Integer
Dim objCTMaqProgTurnoUnitario As ClassCTMaqProgTurno

On Error GoTo Erro_CTMaquinaProgTurno_Le_Periodo

    dtDataInicial = objCTMaqProgTurno.dtDataDe
    dtDataFinal = objCTMaqProgTurno.dtDataAte
    
    'Se a data inicial não está vazia ...
    If Len(dtDataInicial) > 0 Then
        
        'Se a data final está vazia ... assume a data inicial como final
        If Len(dtDataFinal) = 0 Then
    
            dtDataFinal = dtDataInicial
            
        End If
                        
        'verifica qual o período
        iPeriodoEmDias = CInt(DateDiff("d", dtDataInicial, dtDataFinal))
        
        'para cada dia do período passado
        For iIndice = 0 To iPeriodoEmDias
        
            Set objCTMaqProgTurnoUnitario = New ClassCTMaqProgTurno
            
            With objCTMaqProgTurnoUnitario
            
                .lNumIntDocCT = objCTMaqProgTurno.lNumIntDocCT
                .lNumIntDocMaq = objCTMaqProgTurno.lNumIntDocMaq
                
                'se é a primeira data... não implementa nada
                If iIndice = 0 Then
                    .dtData = dtDataInicial
                Else
                    .dtData = DateAdd("d", iIndice, dtDataInicial)
                End If
            
            End With
            
            'Lê CTMaquinaProgTurno Individualmente
            lErro = CF("CTMaquinaTurno_Le", objCTMaqProgTurnoUnitario)
            If lErro <> SUCESSO And lErro <> 136704 Then gError 137262
            
            'se conseguiu ler...
            If lErro = SUCESSO Then
            
                'alimenta a coleção
                colCTMaqProgTurno.Add objCTMaqProgTurnoUnitario
            
            End If
            
        Next
    
    End If
    
    CTMaquinaProgTurno_Le_Periodo = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Le_Periodo:

    CTMaquinaProgTurno_Le_Periodo = gErr

    Select Case gErr

        Case 137262
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152467)

    End Select

    Exit Function

End Function



'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function CTMaquinaProgTurno_Le(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCTMaqProgTurno As typeCTMaqProgTurno

On Error GoTo Erro_CTMaquinaProgTurno_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 136701

    'Alocação de espaço no buffer
    tCTMaqProgTurno.sObservacao = String(STRING_OBSERVACAO, 0)

    'Le a tabelaCTMaquinaProgTurno
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocCT, NumIntDocMaq, Data, Observacao FROM CTMaquinaProgTurno WHERE NumIntDocCT= ?  AND NumIntDocMaq= ?  AND Data= ? ", _
                tCTMaqProgTurno.lNumIntDoc, tCTMaqProgTurno.lNumIntDocCT, tCTMaqProgTurno.lNumIntDocMaq, tCTMaqProgTurno.dtData, _
                tCTMaqProgTurno.sObservacao, _
                objCTMaqProgTurno.lNumIntDocCT, objCTMaqProgTurno.lNumIntDocMaq, objCTMaqProgTurno.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 136702

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136703

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 136704

    objCTMaqProgTurno.lNumIntDoc = tCTMaqProgTurno.lNumIntDoc
    objCTMaqProgTurno.lNumIntDocCT = tCTMaqProgTurno.lNumIntDocCT
    objCTMaqProgTurno.lNumIntDocMaq = tCTMaqProgTurno.lNumIntDocMaq
    objCTMaqProgTurno.dtData = tCTMaqProgTurno.dtData
    objCTMaqProgTurno.sObservacao = tCTMaqProgTurno.sObservacao
    
    lErro = CTMaquinaProgTurno_Le_Itens(objCTMaqProgTurno)
    If lErro <> SUCESSO And lErro <> 137233 Then gError 137235

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CTMaquinaProgTurno_Le = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Le:

    CTMaquinaProgTurno_Le = gErr

    Select Case gErr

        Case 136701
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 136702, 136703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINAPROGTURNO", gErr)

        Case 136704 'Sem dados -> Tratado na rotina chamadora
        
        Case 137235
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152468)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CTMaquinaProgTurno_Le_Itens(ByVal objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTMaqProgTurnoItens As ClassCTMaqProgTurnoItens
Dim iTurno As Integer
Dim dHoras As Double

On Error GoTo Erro_CTMaquinaProgTurno_Le_Itens

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137230

    lErro = Comando_Executar(lComando, "SELECT Turno, Horas FROM CTMaquinaProgTurnoItens WHERE NumIntDocCTMaqPTurno = ? Order by Turno", iTurno, dHoras, objCTMaqProgTurno.lNumIntDoc)
    If lErro <> SUCESSO Then gError 137231
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137232

    'Turno não cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 137233

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCTMaqProgTurnoItens = New ClassCTMaqProgTurnoItens

        objCTMaqProgTurnoItens.lNumIntDocCTMaqPTurno = objCTMaqProgTurno.lNumIntDoc
        objCTMaqProgTurnoItens.iTurno = iTurno
        objCTMaqProgTurnoItens.dHoras = dHoras

        objCTMaqProgTurno.colTurnos.Add objCTMaqProgTurnoItens
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137234

    Loop
    
    Call Comando_Fechar(lComando)
    
    CTMaquinaProgTurno_Le_Itens = SUCESSO

    Exit Function

Erro_CTMaquinaProgTurno_Le_Itens:

   CTMaquinaProgTurno_Le_Itens = gErr

   Select Case gErr
    
        Case 137230
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 137231, 137232, 137234
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQPROGTURNOITENS", gErr)
            
        Case 137233
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152469)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


'##########################

Function CentrodeTrabalho_Le_CTTurnos(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objTurno As ClassTurno
Dim iTurno As Integer
Dim dQtdHoras(1 To 7) As Double
Dim iDiaDaSemana As Integer

On Error GoTo Erro_CentrodeTrabalho_Le_CTTurnos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137204

    lErro = Comando_Executar(lComando, "SELECT Turno, HorasDomingo, HorasSegunda, HorasTerca, HorasQuarta, HorasQuinta, HorasSexta, HorasSabado FROM CTTurnos WHERE NumIntDocCT = ? Order by Turno", iTurno, dQtdHoras(DOMINGO), dQtdHoras(SEGUNDA), dQtdHoras(TERCA), dQtdHoras(QUARTA), dQtdHoras(QUINTA), dQtdHoras(SEXTA), dQtdHoras(SABADO), objCentrodeTrabalho.lNumIntDoc)
    If lErro <> SUCESSO Then gError 137205
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137206

    'Turno não cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 137207

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTurno = New ClassTurno

        objTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
        objTurno.iTurno = iTurno
        
        For iDiaDaSemana = DOMINGO To SABADO
        
            objTurno.dQtdHoras(iDiaDaSemana) = dQtdHoras(iDiaDaSemana)
        
        Next

        objCentrodeTrabalho.colTurnos.Add objTurno
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137208

    Loop
    
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_Le_CTTurnos = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_CTTurnos:

   CentrodeTrabalho_Le_CTTurnos = gErr

   Select Case gErr
    
        Case 137204
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 137205, 137206, 137208
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTTURNOS", gErr)
            
        Case 137207
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152470)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CentrodeTrabalho_Le_CTMaquinasTurnos(ByVal objCTMaquinas As ClassCTMaquinas) As Long

Dim lErro As Long
Dim lComando As Long
Dim objTurno As ClassTurno
Dim iTurno As Integer
Dim dQtdHoras(1 To 7) As Double
Dim iDiaDaSemana As Integer

On Error GoTo Erro_CentrodeTrabalho_Le_CTMaquinasTurnos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137225

    lErro = Comando_Executar(lComando, "SELECT Turno, HorasDomingo, HorasSegunda, HorasTerca, HorasQuarta, HorasQuinta, HorasSexta, HorasSabado FROM CTMaquinasTurnos WHERE NumIntDocMaq = ? And NumIntDocCT = ? Order by Turno", iTurno, dQtdHoras(DOMINGO), dQtdHoras(SEGUNDA), dQtdHoras(TERCA), dQtdHoras(QUARTA), dQtdHoras(QUINTA), dQtdHoras(SEXTA), dQtdHoras(SABADO), objCTMaquinas.lNumIntDocMaq, objCTMaquinas.lNumIntDocCT)
    If lErro <> SUCESSO Then gError 137226
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137227

    'Turno não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 137228

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTurno = New ClassTurno

        objTurno.lNumIntDocMaq = objCTMaquinas.lNumIntDocMaq
        objTurno.lNumIntDocCT = objCTMaquinas.lNumIntDocCT
        objTurno.iTurno = iTurno
        
        For iDiaDaSemana = DOMINGO To SABADO
        
            objTurno.dQtdHoras(iDiaDaSemana) = dQtdHoras(iDiaDaSemana)
        
        Next

        objCTMaquinas.colTurnos.Add objTurno
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137229

    Loop
    
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_Le_CTMaquinasTurnos = SUCESSO

    Exit Function


Erro_CentrodeTrabalho_Le_CTMaquinasTurnos:

   CentrodeTrabalho_Le_CTMaquinasTurnos = gErr

   Select Case gErr
    
        Case 137225
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 137226, 137227, 137229
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINASTURNOS", gErr)
            
        Case 137228
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152471)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Public Function CentrodeTrabalho_Le_Completo(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCentrodeTrabalho As typeCentrodeTrabalho

On Error GoTo Erro_CentrodeTrabalho_Le_Completo

    If objCentrodeTrabalho.lNumIntDoc <> 0 Then
    
        'Lê o CentrodeTrabalho pelo Nome Reduzido
        lErro = CentrodeTrabalho_Le_NumIntDoc(objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 137209
    
    ElseIf Len(objCentrodeTrabalho.sNomeReduzido) <> 0 Then
    
        'Lê o CentrodeTrabalho pelo Nome Reduzido
        lErro = CentrodeTrabalho_Le_NomeReduzido(objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137210
    
    Else
    
        'Lê o CentrodeTrabalho pelo Codigo e Empresa/Filial
        lErro = CentrodeTrabalho_Le(objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134449 Then gError 137211
    
    End If

    'Sem dados
    If lErro <> SUCESSO Then gError 137212
    
    'Lê o CTTurnos
    lErro = CentrodeTrabalho_Le_CTTurnos(objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 137207 Then gError 137213
    
    'Lê o CTCompetencias
    lErro = CentrodeTrabalho_Le_CTCompetencias(objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 134453 Then gError 137214
    
    'Lê o CTMaquinas
    lErro = CentrodeTrabalho_Le_CTMaquinas(objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 134458 Then gError 137215

    'Lê o CTOperadores
    lErro = CentrodeTrabalho_Le_CTOperadores(objCentrodeTrabalho)
    If lErro <> SUCESSO And lErro <> 139100 Then gError 139096

    CentrodeTrabalho_Le_Completo = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_Completo:

    CentrodeTrabalho_Le_Completo = gErr

    Select Case gErr
    
        Case 137209 To 137211, 137213 To 137215, 139096
            'erros tratados nas rotinas chamadas
            
        Case 137212
            'erro tratado na rotina chamadora -> sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152472)

    End Select

    Exit Function

End Function

Public Function ApontamentoProducao_Le(ByVal objApontamentoProducao As ClassApontamentoProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim tApontamentoProducao As typeApontamentoProducao

On Error GoTo Erro_ApontamentoProducao_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 137714

    'Alocação de espaço no buffer
    tApontamentoProducao.sObservacao = String(STRING_OBSERVACAO, 0)

    'Le a tabela ApontamentoProducao
    lErro = Comando_Executar(lComando, "SELECT PercConcluido, Concluido, Quantidade, Data, Observacao FROM ApontamentoProducao WHERE NumIntDocPO= ? ", _
                tApontamentoProducao.dPercConcluido, tApontamentoProducao.iConcluido, tApontamentoProducao.dQuantidade, tApontamentoProducao.dtData, _
                tApontamentoProducao.sObservacao, objApontamentoProducao.lNumIntDocPO)
    If lErro <> AD_SQL_SUCESSO Then gError 137715

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137716

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 137717

    objApontamentoProducao.dPercConcluido = tApontamentoProducao.dPercConcluido
    objApontamentoProducao.iConcluido = tApontamentoProducao.iConcluido
    objApontamentoProducao.dQuantidade = tApontamentoProducao.dQuantidade
    objApontamentoProducao.dtData = tApontamentoProducao.dtData
    objApontamentoProducao.sObservacao = tApontamentoProducao.sObservacao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ApontamentoProducao_Le = SUCESSO

    Exit Function

Erro_ApontamentoProducao_Le:

    ApontamentoProducao_Le = gErr

    Select Case gErr

        Case 137714
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 137715, 137716
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APONTAMENTOPRODUCAO", gErr)

        Case 137717 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152473)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'#############################################################
'Inserido por Wagner
Public Function EstoqueDisponivelProjetado_Le(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal objProdutoFilial As ClassProdutoFilial, ByVal dtData As Date, dEstoqueProjetado As Double, Optional objPMP As ClassPMP) As Long

Dim lErro As Long
Dim colSimulacaoEST As New Collection
Dim objSimulacaoEST As ClassSimulacaoESTItens
Dim colCampos As New Collection
Dim colSaida As New Collection
Dim dtDataAnterior As Date
Dim dMaiorSaldoNegativo As Double
Dim dSaldo As Double

On Error GoTo Erro_EstoqueDisponivelProjetado_Le

    dEstoqueProjetado = 0

    'Le a posição do estoque na data passada
    lErro = EstoqueProjetado_Le(objEstoqueProduto, dtData, colSimulacaoEST, dEstoqueProjetado)
    If lErro <> SUCESSO Then gError 138079
    
    If Not (objPMP Is Nothing) Then
        'Le a posição do estoque na data passada
        lErro = EstoqueProjetadoPMP_Le(objEstoqueProduto, dtData, colSimulacaoEST, dEstoqueProjetado, objPMP)
        If lErro <> SUCESSO Then gError 138080
    End If
    
    'Procura por consumos/venda que vão utilizar o estoque na data
    'colSimulacaoEST retorna com uma coleção de entradas e saídas do estoque posterior a data passada
    
    'Ordena pela prioridade
    colCampos.Add "dtData"
   
    lErro = Ordena_Colecao(colSimulacaoEST, colSaida, colCampos)
    If lErro <> SUCESSO Then gError 138081

    If colSaida.Count > 0 Then
        dtDataAnterior = colSaida.Item(1).dtData
    End If

    dMaiorSaldoNegativo = 0
    dSaldo = 0
    
    For Each objSimulacaoEST In colSaida
    
        'Quando mudar de data zera o saldo e verifica o défice do produto
        If dtDataAnterior <> objSimulacaoEST.dtData Then
        
            'Se o saldo for negativo guarda o valor
            If dSaldo < dMaiorSaldoNegativo Then dMaiorSaldoNegativo = dSaldo
            
            dtDataAnterior = objSimulacaoEST.dtData
            
        End If
        
        'Acumula as quantidades no mesmo dia
        dSaldo = dSaldo + objSimulacaoEST.dQuantidade
    
    Next

    'Se o saldo for negativo guarda o valor
    If dSaldo < dMaiorSaldoNegativo Then dMaiorSaldoNegativo = dSaldo

    objProdutoFilial.sProduto = objEstoqueProduto.sProduto
    objProdutoFilial.iFilialEmpresa = objEstoqueProduto.iFilialEmpresa

    'Le ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 136816

    'Se o saldo posterior a data passada ficou negativo em algum momento é sinal que
    'O Estoque Projetado para data tem a quantidade empenhada
    dEstoqueProjetado = dEstoqueProjetado + dMaiorSaldoNegativo
    
    EstoqueDisponivelProjetado_Le = SUCESSO

    Exit Function

Erro_EstoqueDisponivelProjetado_Le:

    EstoqueDisponivelProjetado_Le = gErr

    Select Case gErr
    
        Case 138079 To 138081, 136816

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152474)

    End Select

    Exit Function

End Function


Public Function EstoqueProjetado_Le(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal dtData As Date, ByVal colSimulacaoEST As Collection, dEstoqueProjetado As Double) As Long

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim dConsumo As Double
Dim dVenda As Double
Dim dProducao As Double
Dim dCompra As Double
Dim dQuantidade As Double
Dim sUM As String
Dim objSimulacaoEST As ClassSimulacaoESTItens
Dim dtDataAux As Date
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double
Dim lNumIntPV As Long
Dim dQuantAux1 As Double
Dim dQuantAux2 As Double

On Error GoTo Erro_EstoqueProjetado_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 138082
    Next
    
    objProduto.sCodigo = objEstoqueProduto.sProduto
    
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138083
    
    'Lê as quantidades do produtos nos almoxs da filial
    lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, objEstoqueProduto.iFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 52979 Then gError 138084
            
    'CONSUMIDO
    lErro = Comando_Executar(alComando(1), "SELECT (E.Quantidade - E.QuantidadeRequisitada), IOP.DataInicioProd " & _
                                        "FROM Empenho AS E, ItensOrdemProducao AS IOP " & _
                                        "WHERE E.NumIntDocItemOP = IOP.NumIntDoc AND IOP.DataInicioProd >= ? AND E.Produto = ? AND E.FilialEmpresa = ? AND " & _
                                        "EXISTS (SELECT PMPItens.NumIntDoc FROM PlanoMestreProducaoUltimo AS PMP, PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMP.CodGeracao = PMPItens.CodGeracao AND PO.NumIntDocPMP = PMPItens.NumIntDoc AND IOP.Produto = PO.Produto AND IOP.Codigo = PO.CodOPOrigem)", _
                                        dQuantidade, dtDataAux, gdtDataAtual, objProduto.sCodigo, objEstoqueProduto.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 138085
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138086
    
    Do While lErro <> AD_SQL_SEM_DADOS
 
        If dtDataAux <= dtData Then
            'Totaliza o consumo projetado até a data
            dConsumo = dConsumo + dQuantidade
        Else
            'Se adata for posterior guarda para verificar empenhos
            Set objSimulacaoEST = New ClassSimulacaoESTItens
        
            objSimulacaoEST.dQuantidade = -1 * dQuantidade
            objSimulacaoEST.dtData = dtDataAux
            objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_COMSUMO
            
            colSimulacaoEST.Add objSimulacaoEST
        
        End If
         
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138087
    
    Loop
    'FIM CONSUMO
        
    'VENDA
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT IPV.Quantidade - IPV.QuantReservada  - IPV.QuantCancelada - IPV.QuantFaturada, IPV.UnidadeMed, IPV.DataEntrega " & _
                                        "FROM ItensPedidoDeVenda AS IPV " & _
                                        "WHERE IPV.DataEntrega >= ? AND IPV.Produto = ? AND IPV.FilialEmpresa = ? AND NOT EXISTS (SELECT IPVE.NumIntPV FROM ItensPVEntrega AS IPVE WHERE IPVE.NumIntItemPV = IPV.NumIntDoc) ", _
                                        dQuantidade, sUM, dtDataAux, gdtDataAtual, objProduto.sCodigo, objEstoqueProduto.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 138088
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138089
   
    Do While lErro <> AD_SQL_SEM_DADOS
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 138090
        
        If dtDataAux <= dtData Then
            'Totaliza o consumo projetado até a data
            dVenda = dVenda + dQuantidade * dFatorConversao
'        Else
'            'Se adata for posterior guarda para verificar empenhos
'            Set objSimulacaoEST = New ClassSimulacaoESTItens
'
'            objSimulacaoEST.dQuantidade = -1 * dQuantidade * dFatorConversao
'            objSimulacaoEST.dtData = dtDataAux
'            objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVVENDA
'
'            colSimulacaoEST.Add objSimulacaoEST
'
        End If
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138091
   
    Loop
    'FIM VENDA
    
    'VENDA COM VÁRIAS DATAS DE ENTREGA
    sUM = String(STRING_UM_SIGLA, 0)
    
    'Le o item de entrega do pedido de venda
    lErro = Comando_Executar(alComando(6), "SELECT IPV.NumIntDoc, IPVE.Quantidade ,IPV.Quantidade - IPV.QuantReservada - IPV.QuantCancelada - IPV.QuantFaturada, IPV.UnidadeMed, IPVE.DataEntrega  " & _
                                        "FROM ItensPedidoDeVenda AS IPV, ItensPVEntrega AS IPVE " & _
                                        "WHERE IPVE.NumIntItemPV = IPV.NumIntDoc AND IPVE.DataEntrega >= ? AND IPV.Produto = ? AND IPV.FilialEmpresa = ? ", _
                                        lNumIntPV, dQuantidade, dQuantAux1, sUM, dtDataAux, gdtDataAtual, objProduto.sCodigo, objEstoqueProduto.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 138088
    
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138089
   
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Le o que vai ter que ser entrega após o item lido acima
        lErro = Comando_Executar(alComando(7), "SELECT SUM(IPVE.Quantidade) " & _
                                            "FROM  ItensPVEntrega AS IPVE " & _
                                            "WHERE IPVE.NumIntItemPV = ? AND IPVE.DataEntrega > ? ", _
                                            dQuantAux2, lNumIntPV, dtDataAux)
        If lErro <> SUCESSO Then gError 138088
        
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138089
        
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
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 138090
        
        If dtDataAux <= dtData Then
            'Totaliza o consumo projetado até a data
            dVenda = dVenda + dQuantidade * dFatorConversao
        End If
        
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138091
   
    Loop
    'FIM VENDA COM VÁRIAS DATAS DE ENTREGA

   
    'PRODUÇÃO
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(4), "SELECT IOP.Quantidade, IOP.SiglaUM, IOP.DataFimProd " & _
                                        "FROM ItensOrdemProducao AS IOP " & _
                                        "WHERE IOP.DataFimProd >= ? AND IOP.Produto = ? AND IOP.FilialEmpresa = ? AND " & _
                                        "EXISTS (SELECT PMPItens.NumIntDoc FROM PlanoMestreProducaoUltimo AS PMP, PlanoMestreProducaoItens AS PMPItens, PlanoOperacional AS PO WHERE PMP.CodGeracao = PMPItens.CodGeracao AND PO.NumIntDocPMP = PMPItens.NumIntDoc AND IOP.Produto = PO.Produto AND IOP.Codigo = PO.CodOPOrigem)", _
                                        dQuantidade, sUM, dtDataAux, gdtDataAtual, objProduto.sCodigo, objEstoqueProduto.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 138092
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138093
   
    Do While lErro <> AD_SQL_SEM_DADOS
           
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 138094
        
        If dtDataAux <= dtData Then
            'Totaliza o consumo projetado até a data
            dProducao = dProducao + dQuantidade * dFatorConversao
        Else
            'Se adata for posterior guarda para verificar empenhos
            Set objSimulacaoEST = New ClassSimulacaoESTItens
        
            objSimulacaoEST.dQuantidade = dQuantidade * dFatorConversao
            objSimulacaoEST.dtData = dtDataAux
            objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAO
            
            colSimulacaoEST.Add objSimulacaoEST
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138095
   
    Loop
    'FIM PRODUÇÃO

    'COMPRA
    sUM = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(4), "SELECT IRC.Quantidade - IRC.QuantCancelada AS Quantidade, IRC.UM, R.DataLimite " & _
                                            "FROM RequisicaoCompra AS R, ItensReqCompra AS IRC " & _
                                            "WHERE R.NumIntDoc = IRC.ReqCompra AND IRC.StatusBaixa = ? AND IRC.Produto = ? AND R.DataLimite >= ? AND R.FilialEmpresa = ? ", _
                                            dQuantidade, sUM, dtDataAux, 0, objProduto.sCodigo, gdtDataAtual, objEstoqueProduto.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 138096
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138097
   
    Do While lErro <> AD_SQL_SEM_DADOS
           
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 138098
        
        If dtDataAux <= dtData Then
            'Totaliza o consumo projetado até a data
            dCompra = dCompra + dQuantidade * dFatorConversao
        Else
            'Se adata for posterior guarda para verificar empenhos
            Set objSimulacaoEST = New ClassSimulacaoESTItens
        
            objSimulacaoEST.dQuantidade = dQuantidade * dFatorConversao
            objSimulacaoEST.dtData = dtDataAux
            objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVCOMPRA
            
            colSimulacaoEST.Add objSimulacaoEST
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138099
   
    Loop
    'FIM COMPRA

    dEstoqueProjetado = objEstoqueProduto.dQuantDisponivel + dProducao + dCompra - dConsumo - dVenda '- objEstoqueProduto.dQuantReservada

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    EstoqueProjetado_Le = SUCESSO

    Exit Function

Erro_EstoqueProjetado_Le:

    EstoqueProjetado_Le = gErr

    Select Case gErr

        Case 138082
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 138083, 138084, 138090, 138094, 138098
        
        Case 138085, 138086, 138087
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO2", gErr)
    
        Case 138088, 138089, 138091
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
            
        Case 138092, 138093, 138095
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO3", gErr)
            
        Case 138096, 138097, 138099
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152475)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Public Function EstoqueProjetadoPMP_Le(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal dtData As Date, ByVal colSimulacaoEST As Collection, dEstoqueProjetado As Double, ByVal objPMP As ClassPMP) As Long
'
'Dim lErro As Long
'Dim objPMPItens As ClassPMPItens
'Dim objPO As ClassPlanoOperacional
'Dim objItemOP As ClassItemOP
'Dim objRC As ClassRequisicaoCompras
'Dim objItemRC As ClassItemReqCompras
'Dim objSimulacaoEST As ClassSimulacaoESTItens
'Dim objOPInsumos As ClassOrdemProducaoInsumos
'Dim objOPOperacao As ClassOrdemProducaoOperacoes
'Dim objOPInsumoAux As ClassOrdemProducaoInsumos
'Dim objOPOperacaoAux As ClassOrdemProducaoOperacoes
'Dim objOPOperacaoAux2 As ClassOrdemProducaoOperacoes
'Dim objOPOperacaoAux3 As ClassOrdemProducaoOperacoes
'Dim dFatorConversao As Double
'Dim dFatorMultiplicacao As Double
'Dim dConsumo As Double
'Dim dVenda As Double
'Dim dProducao As Double
'Dim dCompra As Double
'Dim objProduto As New ClassProduto
'Dim sCodigoOPAnt As String
'Dim sProdutoAnt As String, colPO As Collection
'Dim bPegarPO As Boolean, sCodOPOrigem As String, sProduto As String
'Dim iNivel As Integer, iSeq As Integer, iFilialEmpresa As Integer
'
'On Error GoTo Erro_EstoqueProjetadoPMP_Le
'
'    objProduto.sCodigo = objEstoqueProduto.sProduto
'
'    'Lê o produto
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 138100
'
'    'Acerta as datas das OPs e RCs de acordo com as alterações do usuário
'    lErro = CF("PMP_Acerta_OP_RC", objPMP)
'    If lErro <> SUCESSO Then gError 138101
'
'    'Para cada item do Plano Mestre
'    For Each objPMPItens In objPMP.colItens
'
'        If objPMPItens.iFilialEmpresa = objEstoqueProduto.iFilialEmpresa Then
'
'            Set objItemOP = objPMPItens.objItemOP
'
'            sCodigoOPAnt = ""
'            sProdutoAnt = ""
'
'            'Pega o colOP de acordo com o que já foi calculado e o que está sendo
'            'Tem que ver o cenário completo para calcular direito
'            Set colPO = New Collection
'            sCodOPOrigem = ""
'            sProduto = ""
'            iNivel = 0
'            iSeq = 0
'            iFilialEmpresa = 0
'            bPegarPO = False
'            'Pega os POs que foram calculadas nessa passada da função
'            For Each objPO In objPMPItens.colPO
'                sCodOPOrigem = objPO.sCodOPOrigem
'                sProduto = objPO.sProduto
'                iNivel = objPO.iNivel
'                iSeq = objPO.iSeq
'                iFilialEmpresa = objPO.iFilialEmpresa
'                colPO.Add objPO
'            Next
'            'Se já houve um calculo anterior pega os POs que ainda serão calculados
'            'a partir do que já foi
'            If colPO.Count = 0 Then bPegarPO = True
'            For Each objPO In objPMPItens.colPOAux
'                If bPegarPO Then
'                    colPO.Add objPO
'                End If
'                If sProduto = objPO.sProduto And iNivel = objPO.iNivel And iSeq = objPO.iSeq And iSeq = objPO.iSeq And iFilialEmpresa = objPO.iFilialEmpresa Then
'                    bPegarPO = True
'                End If
'            Next
'
'            'Para cada Etapa
'            For Each objPO In colPO
'
'                If objPO.sProduto = objEstoqueProduto.sProduto Then
'
'                    If sCodigoOPAnt <> objPO.sCodOPOrigem And sProdutoAnt <> objPO.sProduto Then
'
'                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
'                        If lErro <> SUCESSO Then gError 138102
'
'                        If objPO.dtDataFim <= dtData Then
'                            'Totaliza o consumo projetado até a data
'                            dProducao = dProducao + objPO.dQuantidade * dFatorConversao
'
'                        Else
'                            'Se adata for posterior guarda para verificar empenhos
'                            Set objSimulacaoEST = New ClassSimulacaoESTItens
'
'                            objSimulacaoEST.dQuantidade = objPO.dQuantidade * dFatorConversao
'                            objSimulacaoEST.dtData = objPO.dtDataFim
'                            objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAO
'
'                            colSimulacaoEST.Add objSimulacaoEST
'
'                        End If
'
'                        sCodigoOPAnt = objPO.sCodOPOrigem
'                        sProdutoAnt = objPO.sProduto
'
'
'                    End If
'
'                End If
'
'                For Each objOPOperacao In objItemOP.colOrdemProducaoOperacoes
'
'                    If objOPOperacao.lNumIntDoc = objPO.lNumIntDocOper Then
'
'                        For Each objOPInsumos In objOPOperacao.colOPInsumos
'
'                            If objOPInsumos.sProduto = objEstoqueProduto.sProduto Then
'
'                                If objOPOperacao.iNivelRoteiro <> KIT_NIVEL_RAIZ Then
'
'                                   'Tenta achar a Operação Pai
'                                    For Each objOPOperacaoAux In objItemOP.colOrdemProducaoOperacoes
'
'                                        'Se é a Operação Pai
'                                        If objOPOperacaoAux.iSeq = objOPOperacao.iSeqPai Then
'
'                                            'Verifica se é do mesmo roteiro
'                                            If objOPOperacaoAux.iNivelRoteiro = objOPOperacao.iNivelRoteiro And objOPOperacaoAux.iSeqRoteiro = objOPOperacao.iSeqRoteiro Then
'
'                                                'Acha a primeira Operação do Roteiro
'                                                For Each objOPOperacaoAux2 In objItemOP.colOrdemProducaoOperacoes
'
'                                                     If objOPOperacao.iNivelRoteiro = objOPOperacaoAux2.iNivelRoteiro And objOPOperacao.iSeqRoteiro = objOPOperacaoAux2.iSeqRoteiro Then
'                                                        Exit For
'                                                     End If
'
'                                                Next
'
'                                                'Acha a Operação Pai da Primeira Operação do roteiro
'                                                For Each objOPOperacaoAux3 In objItemOP.colOrdemProducaoOperacoes
'
'                                                     If objOPOperacaoAux3.iSeq = objOPOperacaoAux2.iSeqPai Then
'                                                        Exit For
'                                                     End If
'
'                                                Next
'
'                                            Else
'
'                                                Set objOPOperacaoAux3 = objOPOperacaoAux
'
'                                            End If
'
'                                            'Tenta achar o Insumo que originou a OP pai
'                                            For Each objOPInsumoAux In objOPOperacaoAux3.colOPInsumos
'
'                                                If objOPInsumoAux.sProduto = objOPOperacao.sProduto Then
'
'                                                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objOPInsumoAux.sUMProduto, dFatorConversao)
'                                                    If lErro <> SUCESSO Then gError 138103
'
'                                                    dFatorMultiplicacao = (objPO.dQuantidade * dFatorConversao) / objOPInsumoAux.dQuantidade
'
'                                                    Exit For
'
'                                                End If
'
'                                            Next
'
'                                        End If
'
'                                    Next
'
'                                Else
'
'                                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objItemOP.sSiglaUM, dFatorConversao)
'                                    If lErro <> SUCESSO Then gError 138103
'
'                                    dFatorMultiplicacao = (objPO.dQuantidade * dFatorConversao) / objItemOP.dQuantidade
'
'                                End If
'
'                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOPInsumos.sUMProduto, objProduto.sSiglaUMEstoque, dFatorConversao)
'                                If lErro <> SUCESSO Then gError 138103
'
'                                If objPO.dtDataInicio <= dtData Then
'                                    'Totaliza o consumo projetado até a data
'                                    dConsumo = dConsumo + objOPInsumos.dQuantidade * dFatorMultiplicacao * dFatorConversao
'                                Else
'                                    'Se adata for posterior guarda para verificar empenhos
'                                    Set objSimulacaoEST = New ClassSimulacaoESTItens
'
'                                    objSimulacaoEST.dQuantidade = (-1 * objOPInsumos.dQuantidade * dFatorMultiplicacao * dFatorConversao) / (1 - objOPInsumos.dPercentualPerda)
'                                    objSimulacaoEST.dtData = objPO.dtDataInicio
'                                    objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_COMSUMO
'
'                                    colSimulacaoEST.Add objSimulacaoEST
'
'                                End If
'                            End If
'                        Next
'                    End If
'                Next
'
'                For Each objRC In objPO.colRCFilhas
'
'                    For Each objItemRC In objRC.colItens
'
'                        If objItemRC.sProduto = objEstoqueProduto.sProduto Then
'
'                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
'                            If lErro <> SUCESSO Then gError 138104
'
'                            If objRC.dtDataLimite <= dtData Then
'                                'Totaliza o consumo projetado até a data
'                                dCompra = dCompra + objItemRC.dQuantidade * dFatorConversao
'                            Else
'                                'Se adata for posterior guarda para verificar empenhos
'                                Set objSimulacaoEST = New ClassSimulacaoESTItens
'
'                                objSimulacaoEST.dQuantidade = objItemRC.dQuantidade * dFatorConversao
'                                objSimulacaoEST.dtData = objRC.dtDataLimite
'                                objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVCOMPRA
'
'                                colSimulacaoEST.Add objSimulacaoEST
'
'                            End If
'                        End If
'                    Next
'                Next
'            Next
'        End If
'    Next
'
'    dEstoqueProjetado = dEstoqueProjetado + dProducao + dCompra - dConsumo - dVenda
'
'    EstoqueProjetadoPMP_Le = SUCESSO
'
'    Exit Function
'
'Erro_EstoqueProjetadoPMP_Le:
'
'    EstoqueProjetadoPMP_Le = gErr
'
'    Select Case gErr
'
'        Case 138100 To 138104
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152476)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function EstoqueProjetadoPMP_Le(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal dtData As Date, ByVal colSimulacaoEST As Collection, dEstoqueProjetado As Double, ByVal objPMP As ClassPMP) As Long
'Alterado para não só considerar o cenário que está sendo calculado no momento como
'também o cenário completo calculado anteriormente
Dim lErro As Long
Dim objPMPItens As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objItemOP As ClassItemOP, objItemOPAux As ClassItemOP
Dim objRC As ClassRequisicaoCompras
Dim objItemRC As ClassItemReqCompras
Dim objSimulacaoEST As ClassSimulacaoESTItens
Dim objOPInsumos As ClassOrdemProducaoInsumos
Dim objOPOperacao As ClassOrdemProducaoOperacoes
Dim objOPInsumoAux As ClassOrdemProducaoInsumos
Dim objOPOperacaoAux As ClassOrdemProducaoOperacoes
Dim objOPOperacaoAux2 As ClassOrdemProducaoOperacoes
Dim objOPOperacaoAux3 As ClassOrdemProducaoOperacoes
Dim dFatorConversao As Double
Dim dFatorMultiplicacao As Double
Dim dConsumo As Double
Dim dVenda As Double
Dim dProducao As Double
Dim dCompra As Double
Dim objProduto As New ClassProduto
Dim sCodigoOPAnt As String
Dim sProdutoAnt As String, colPO As Collection
Dim bPegarPO As Boolean, sCodOPOrigem As String, sProduto As String
Dim iNivel As Integer, iSeq As Integer, iFilialEmpresa As Integer
Dim objOP As ClassOrdemDeProducao
 
On Error GoTo Erro_EstoqueProjetadoPMP_Le
 
    objProduto.sCodigo = objEstoqueProduto.sProduto
 
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138100
 
    'Acerta as datas das OPs e RCs de acordo com as alterações do usuário
    lErro = CF("PMP_Acerta_OP_RC", objPMP)
    If lErro <> SUCESSO Then gError 138101
        
    'Para cada item do Plano Mestre
    For Each objPMPItens In objPMP.colItens
    
        If objPMPItens.iFilialEmpresa = objEstoqueProduto.iFilialEmpresa Then
    
            Set objItemOP = objPMPItens.objItemOP
                        
            sCodigoOPAnt = ""
            sProdutoAnt = ""
            
            'Pega o colOP de acordo com o que já foi calculado e o que está sendo
            'Tem que ver o cenário completo para calcular direito
            Set colPO = New Collection
            sCodOPOrigem = ""
            sProduto = ""
            iNivel = 0
            iSeq = 0
            iFilialEmpresa = 0
            bPegarPO = False
            'Pega os POs que foram calculadas nessa passada da função
            For Each objPO In objPMPItens.colPO
                sCodOPOrigem = objPO.sCodOPOrigem
                sProduto = objPO.sProduto
                iNivel = objPO.iNivel
                iSeq = objPO.iSeq
                iFilialEmpresa = objPO.iFilialEmpresa
                colPO.Add objPO
            Next
            'Se já houve um calculo anterior pega os POs que ainda serão calculados
            'a partir do que já foi
            If colPO.Count = 0 Then bPegarPO = True
            For Each objPO In objPMPItens.colPOAux
                If bPegarPO Then
                    colPO.Add objPO
                End If
                If sProduto = objPO.sProduto And iNivel = objPO.iNivel And iSeq = objPO.iSeq And iSeq = objPO.iSeq And iFilialEmpresa = objPO.iFilialEmpresa Then
                    bPegarPO = True
                End If
            Next
    
            'Se a OP é planejada considera como produção planejada além das subops
            If objPMPItens.objOP.iStatusOP = ITEMOP_SITUACAO_PLANEJADA Then
            
                If objPMPItens.objItemOP.sProduto = objProduto.sCodigo Then

                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPMPItens.objItemOP.sSiglaUM, objProduto.sSiglaUMEstoque, dFatorConversao)
                    If lErro <> SUCESSO Then gError 138102
    
                    If objPMPItens.objItemOP.dtDataFimProd <= dtData Then
                        'Totaliza o consumo projetado até a data
                        dProducao = dProducao + objPMPItens.objItemOP.dQuantidade * dFatorConversao
    
                    Else
                        'Se adata for posterior guarda para verificar empenhos
                        Set objSimulacaoEST = New ClassSimulacaoESTItens
    
                        objSimulacaoEST.dQuantidade = objPMPItens.objItemOP.dQuantidade * dFatorConversao
                        objSimulacaoEST.dtData = objPMPItens.objItemOP.dtDataFimProd
                        objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAO
    
                        colSimulacaoEST.Add objSimulacaoEST
    
                    End If
                
                End If

            End If
            
            'Para cada Etapa
            For Each objPO In colPO
                
                For Each objOPOperacao In objItemOP.colOrdemProducaoOperacoes
                
                    If objOPOperacao.lNumIntDoc = objPO.lNumIntDocOper Then
                                            
                        For Each objOPInsumos In objOPOperacao.colOPInsumos
                        
                            If objOPInsumos.sProduto = objEstoqueProduto.sProduto Then
                            
                                If objOPOperacao.iNivelRoteiro <> KIT_NIVEL_RAIZ Then
                            
                                   'Tenta achar a Operação Pai
                                    For Each objOPOperacaoAux In objItemOP.colOrdemProducaoOperacoes
                                    
                                        'Se é a Operação Pai
                                        If objOPOperacaoAux.iSeq = objOPOperacao.iSeqPai Then
                                        
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
                                        
                                        End If
                                    
                                    Next
                                    
                                Else
                                
                                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objPO.sUM, objItemOP.sSiglaUM, dFatorConversao)
                                    If lErro <> SUCESSO Then gError 138103
                                
                                    dFatorMultiplicacao = (objPO.dQuantidade * dFatorConversao) / objItemOP.dQuantidade
                                
                                End If
                            
                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objOPInsumos.sUMProduto, objProduto.sSiglaUMEstoque, dFatorConversao)
                                If lErro <> SUCESSO Then gError 138103
                           
                                If objPO.dtDataInicio <= dtData Then
                                    'Totaliza o consumo projetado até a data
                                    dConsumo = dConsumo + objOPInsumos.dQuantidade * dFatorMultiplicacao * dFatorConversao
                                Else
                                    'Se adata for posterior guarda para verificar empenhos
                                    Set objSimulacaoEST = New ClassSimulacaoESTItens
                                
                                    objSimulacaoEST.dQuantidade = (-1 * objOPInsumos.dQuantidade * dFatorMultiplicacao * dFatorConversao) / (1 - objOPInsumos.dPercentualPerda)
                                    objSimulacaoEST.dtData = objPO.dtDataInicio
                                    objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_COMSUMO
                                    
                                    colSimulacaoEST.Add objSimulacaoEST
                                
                                End If
                            End If
                        Next
                    End If
                Next
                    
                For Each objRC In objPO.colRCFilhas
                
                    For Each objItemRC In objRC.colItens
                
                        If objItemRC.sProduto = objEstoqueProduto.sProduto Then
                        
                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemRC.sUM, objProduto.sSiglaUMEstoque, dFatorConversao)
                            If lErro <> SUCESSO Then gError 138104
                    
                            If objRC.dtDataLimite <= dtData Then
                                'Totaliza o consumo projetado até a data
                                dCompra = dCompra + objItemRC.dQuantidade * dFatorConversao
                            Else
                                'Se adata for posterior guarda para verificar empenhos
                                Set objSimulacaoEST = New ClassSimulacaoESTItens
                            
                                objSimulacaoEST.dQuantidade = objItemRC.dQuantidade * dFatorConversao
                                objSimulacaoEST.dtData = objRC.dtDataLimite
                                objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PREVCOMPRA
                                
                                colSimulacaoEST.Add objSimulacaoEST
                            
                            End If
                        End If
                    Next
                Next
                
                
                For Each objOP In objPO.colOPFilhas
                
                    For Each objItemOPAux In objOP.colItens
                
                        If objItemOPAux.sProduto = objEstoqueProduto.sProduto Then
                        
                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemOPAux.sSiglaUM, objProduto.sSiglaUMEstoque, dFatorConversao)
                            If lErro <> SUCESSO Then gError 138104
                    
                            If objItemOPAux.dtDataFimProd <= dtData Then
                                'Totaliza o consumo projetado até a data
                                dProducao = dProducao + objItemOPAux.dQuantidade * dFatorConversao
                            Else
                                'Se adata for posterior guarda para verificar empenhos
                                Set objSimulacaoEST = New ClassSimulacaoESTItens
                            
                                objSimulacaoEST.dQuantidade = objItemOPAux.dQuantidade * dFatorConversao
                                objSimulacaoEST.dtData = objItemOPAux.dtDataFimProd
                                objSimulacaoEST.iTipo = SIMULACAO_ESTOQUE_TIPO_PRODUCAO
                                
                                colSimulacaoEST.Add objSimulacaoEST
                            
                            End If
                        End If
                    Next
                Next
                
            Next
        End If
    Next
    
    dEstoqueProjetado = dEstoqueProjetado + dProducao + dCompra - dConsumo - dVenda

'    'gsMRPTexto = "Indice;Seq;Produto;Data;QtdInsumo;CMR;CSR;VR;PR;CMP;CSP;VP;PP;QtdProjetada;EstoqueSeguranca;QtdAtual"
'    gsMRPTexto = gsMRPTexto & ";" & Formata_Estoque(dCompra) & ";" & Formata_Estoque(dConsumo) & ";" & Formata_Estoque(dVenda) & ";" & Formata_Estoque(dProducao)

    EstoqueProjetadoPMP_Le = SUCESSO
    
    Exit Function
 
Erro_EstoqueProjetadoPMP_Le:
 
    EstoqueProjetadoPMP_Le = gErr
 
    Select Case gErr
    
        Case 138100 To 138104
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152476)
 
    End Select
 
    Exit Function
 
End Function

Public Function Produto_Le_Faixa(ByVal vsProdIni As Variant, ByVal vsProdFim As Variant, ByVal colProduto As Collection) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vsProduto As Variant
Dim alComando(1 To 1) As Long
Dim objProduto As ClassProduto
 
On Error GoTo Erro_Produto_Le_Faixa
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 136705
    Next
    
    sSQL = "SELECT Codigo FROM Produtos "
 
    If vsProdIni <> "" Then
        sWhere = sWhere & "WHERE Codigo >= ? "
    End If
    If vsProdFim <> "" Then
        If Len(Trim(sWhere)) <> 0 Then
            sWhere = sWhere & "AND "
        Else
            sWhere = "WHERE "
        End If
        sWhere = sWhere & "Codigo <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY Codigo "
    
    vsProduto = String(STRING_PRODUTO, 0)
 
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136706
 
    lErro = Comando_BindVarInt(alComando(1), vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 136707
 
    If vsProdIni <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsProdIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136708
    End If
    
    If vsProdFim <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsProdFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 136709
    End If
 
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 136710
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136711
        
    'Para cada Produto
    Do While lErro = SUCESSO
    
        Set objProduto = New ClassProduto
 
        objProduto.sCodigo = vsProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 136712
     
        colProduto.Add objProduto
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 136713
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Produto_Le_Faixa = SUCESSO
 
    Exit Function
 
Erro_Produto_Le_Faixa:
 
    Produto_Le_Faixa = gErr
 
    Select Case gErr
 
        Case 136705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 136706 To 136711, 136713
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case 136712
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152477)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

'#############################################################
'
Public Function OPOperacoes_Cria_Copia(ByVal objOrdemProdOper As ClassOrdemProducaoOperacoes, objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long
'Faz a cópia de objOrdemProdOper para objOrdemProducaoOperacoes

Dim lErro As Long
Dim objOrdemProdIns As ClassOrdemProducaoInsumos
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim objOperacoesTempo As ClassOperacoesTempo

On Error GoTo Erro_OPOperacoes_Cria_Copia

    objOrdemProducaoOperacoes.lNumIntDoc = objOrdemProdOper.lNumIntDoc
    objOrdemProducaoOperacoes.lNumIntDocItemOP = objOrdemProdOper.lNumIntDocItemOP
    objOrdemProducaoOperacoes.iSeq = objOrdemProdOper.iSeq
    objOrdemProducaoOperacoes.lNumIntDocCompet = objOrdemProdOper.lNumIntDocCompet
    objOrdemProducaoOperacoes.lNumIntDocCT = objOrdemProdOper.lNumIntDocCT
    objOrdemProducaoOperacoes.sObservacao = objOrdemProdOper.sObservacao
    objOrdemProducaoOperacoes.iSeqPai = objOrdemProdOper.iSeqPai
    objOrdemProducaoOperacoes.iSeqArvore = objOrdemProdOper.iSeqArvore
    objOrdemProducaoOperacoes.iNivel = objOrdemProdOper.iNivel
    objOrdemProducaoOperacoes.iSeqRoteiro = objOrdemProdOper.iSeqRoteiro
    objOrdemProducaoOperacoes.iSeqRoteiroPai = objOrdemProdOper.iSeqRoteiroPai
    objOrdemProducaoOperacoes.iNivelRoteiro = objOrdemProdOper.iNivelRoteiro
    objOrdemProducaoOperacoes.sProduto = objOrdemProdOper.sProduto
    objOrdemProducaoOperacoes.sVersao = objOrdemProdOper.sVersao
    objOrdemProducaoOperacoes.iIgnoraTaxaProducao = objOrdemProdOper.iIgnoraTaxaProducao
    objOrdemProducaoOperacoes.iConsideraCarga = objOrdemProdOper.iConsideraCarga
    objOrdemProducaoOperacoes.iOrigem = objOrdemProdOper.iOrigem
    objOrdemProducaoOperacoes.lNumIntDocOperOrigem = objOrdemProdOper.lNumIntDocOperOrigem
    objOrdemProducaoOperacoes.iNumMaxMaqPorOper = objOrdemProdOper.iNumMaxMaqPorOper
    objOrdemProducaoOperacoes.iNumRepeticoes = objOrdemProdOper.iNumRepeticoes
    
    For Each objOrdemProdIns In objOrdemProdOper.colOPInsumos
    
        Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos
        
        lErro = OPInsumos_Cria_Copia(objOrdemProdIns, objOrdemProducaoInsumos)
        If lErro <> SUCESSO Then gError 137030

        objOrdemProducaoOperacoes.colOPInsumos.Add objOrdemProducaoInsumos
    
    Next
    
    If objOrdemProdOper.iIgnoraTaxaProducao = MARCADO Then
    
        Set objOperacoesTempo = New ClassOperacoesTempo
        
        lErro = OPTempo_Cria_Copia(objOrdemProdOper.objOperacoesTempo, objOperacoesTempo)
        If lErro <> SUCESSO Then gError 137031

        Set objOrdemProducaoOperacoes.objOperacoesTempo = objOperacoesTempo
    
    End If
    
    OPOperacoes_Cria_Copia = SUCESSO

    Exit Function

Erro_OPOperacoes_Cria_Copia:

    OPOperacoes_Cria_Copia = gErr

    Select Case gErr
    
        Case 137030, 137031
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152478)

    End Select

    Exit Function

End Function

Function OPInsumos_Cria_Copia(ByVal objOrdemProdIns As ClassOrdemProducaoInsumos, objOrdemProducaoInsumos As ClassOrdemProducaoInsumos) As Long
'Faz a cópia de objOrdemProdIns para objOrdemProducaoInsumos

Dim lErro As Long

On Error GoTo Erro_OPInsumos_Cria_Copia

    objOrdemProducaoInsumos.lNumIntDocOper = objOrdemProdIns.lNumIntDocOper
    objOrdemProducaoInsumos.sProduto = objOrdemProdIns.sProduto
    objOrdemProducaoInsumos.sVersaoKitComp = objOrdemProdIns.sVersaoKitComp
    objOrdemProducaoInsumos.sUMProduto = objOrdemProdIns.sUMProduto
    objOrdemProducaoInsumos.dQuantidade = objOrdemProdIns.dQuantidade
    objOrdemProducaoInsumos.dPercentualPerda = objOrdemProdIns.dPercentualPerda
    objOrdemProducaoInsumos.iComposicao = objOrdemProdIns.iComposicao
    objOrdemProducaoInsumos.dCustoStandard = objOrdemProdIns.dCustoStandard
    
    OPInsumos_Cria_Copia = SUCESSO

    Exit Function

Erro_OPInsumos_Cria_Copia:

    OPInsumos_Cria_Copia = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152479)

    End Select

    Exit Function

End Function

Function OPTempo_Cria_Copia(ByVal objOrdemProdTempo As ClassOperacoesTempo, objOperacoesTempo As ClassOperacoesTempo) As Long
'Faz a cópia de objOrdemProdTempo para objOperacoesTempo

Dim lErro As Long

On Error GoTo Erro_OPTempo_Cria_Copia

    objOperacoesTempo.lNumIntDocOperacao = objOrdemProdTempo.lNumIntDocOperacao
    objOperacoesTempo.iTipo = objOrdemProdTempo.iTipo
    objOperacoesTempo.dTempoOperacao = objOrdemProdTempo.dTempoOperacao
    objOperacoesTempo.sUMTempo = objOrdemProdTempo.sUMTempo
    objOperacoesTempo.dTempoPreparacao = objOrdemProdTempo.dTempoPreparacao
    objOperacoesTempo.dTempoMovimentacao = objOrdemProdTempo.dTempoMovimentacao
    objOperacoesTempo.dTempoDescarga = objOrdemProdTempo.dTempoDescarga
    objOperacoesTempo.dLoteMin = objOrdemProdTempo.dLoteMin
    objOperacoesTempo.dLoteMax = objOrdemProdTempo.dLoteMax
    objOperacoesTempo.dLotePadrao = objOrdemProdTempo.dLotePadrao
    objOperacoesTempo.lNumIntDocMaq = objOrdemProdTempo.lNumIntDocMaq
    
    OPTempo_Cria_Copia = SUCESSO

    Exit Function

Erro_OPTempo_Cria_Copia:

    OPTempo_Cria_Copia = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152480)

    End Select

    Exit Function

End Function

'Public Function CusteioRoteiro_Le(ByVal objCusteioRoteiro As ClassCusteioRoteiro) As Long
''Esta função lê o CusteioRoteiro pelo Código ou pelo NumIntDoc se estiver preenchido, sendo
''o NumIntDoc preferêncial para leitura, senão tiver ele lê pelo código
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tCusteioRoteiro As typeCusteioRoteiro
'
'On Error GoTo Erro_CusteioRoteiro_Le
'
'    'Executa a abertura do Comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 137939
'
'    'Alocação de espaço no buffer
'    tCusteioRoteiro.sNomeReduzido = String(STRING_CUSTEIO_NOMEREDUZIDO, 0)
'    tCusteioRoteiro.sDescricao = String(STRING_CUSTEIO_DESCRICAO, 0)
'    tCusteioRoteiro.sProduto = String(STRING_PRODUTO, 0)
'    tCusteioRoteiro.sVersao = String(STRING_KIT_VERSAO, 0)
'    tCusteioRoteiro.sUMedida = String(STRING_UM_SIGLA, 0)
'    tCusteioRoteiro.sObservacao = String(STRING_OBSERVACAO, 0)
'
'    If objCusteioRoteiro.lNumIntDoc <> 0 Then
'
'        'Le a tabela CusteioRoteiroFabricacao
'        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NomeReduzido, Descricao, Produto, Versao, UMedida, Quantidade, DataCusteio, DataValidade, CustoTotalInsumosKit, CustoTotalInsumosMaq, CustoTotalMaoDeObra, PrecoTotalRoteiro, Observacao FROM CusteioRoteiroFabricacao WHERE NumIntDoc= ? ", _
'                    tCusteioRoteiro.lNumIntDoc, tCusteioRoteiro.lCodigo, tCusteioRoteiro.sNomeReduzido, tCusteioRoteiro.sDescricao, tCusteioRoteiro.sProduto, tCusteioRoteiro.sVersao, tCusteioRoteiro.sUMedida, tCusteioRoteiro.dQuantidade, tCusteioRoteiro.dtDataCusteio, tCusteioRoteiro.dtDataValidade, tCusteioRoteiro.dCustoTotalInsumosKit, tCusteioRoteiro.dCustoTotalInsumosMaq, tCusteioRoteiro.dCustoTotalMaoDeObra, tCusteioRoteiro.dPrecoTotalRoteiro, tCusteioRoteiro.sObservacao, objCusteioRoteiro.lNumIntDoc)
'
'    Else
'
'        'Le a tabela CusteioRoteiroFabricacao
'        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, NomeReduzido, Descricao, Produto, Versao, UMedida, Quantidade, DataCusteio, DataValidade, CustoTotalInsumosKit, CustoTotalInsumosMaq, CustoTotalMaoDeObra, PrecoTotalRoteiro, Observacao FROM CusteioRoteiroFabricacao WHERE Codigo= ? ", _
'                    tCusteioRoteiro.lNumIntDoc, tCusteioRoteiro.lCodigo, tCusteioRoteiro.sNomeReduzido, tCusteioRoteiro.sDescricao, tCusteioRoteiro.sProduto, tCusteioRoteiro.sVersao, tCusteioRoteiro.sUMedida, tCusteioRoteiro.dQuantidade, tCusteioRoteiro.dtDataCusteio, tCusteioRoteiro.dtDataValidade, tCusteioRoteiro.dCustoTotalInsumosKit, tCusteioRoteiro.dCustoTotalInsumosMaq, tCusteioRoteiro.dCustoTotalMaoDeObra, tCusteioRoteiro.dPrecoTotalRoteiro, tCusteioRoteiro.sObservacao, objCusteioRoteiro.lCodigo)
'
'    End If
'
'    If lErro <> AD_SQL_SUCESSO Then gError 137940
'
'    'Busca Primeiro
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137941
'
'    'se não tiver dados ...
'    If lErro = AD_SQL_SEM_DADOS Then gError 137942
'
'    objCusteioRoteiro.lNumIntDoc = tCusteioRoteiro.lNumIntDoc
'    objCusteioRoteiro.lCodigo = tCusteioRoteiro.lCodigo
'    objCusteioRoteiro.sNomeReduzido = tCusteioRoteiro.sNomeReduzido
'    objCusteioRoteiro.sDescricao = tCusteioRoteiro.sDescricao
'    objCusteioRoteiro.sProduto = tCusteioRoteiro.sProduto
'    objCusteioRoteiro.sVersao = tCusteioRoteiro.sVersao
'    objCusteioRoteiro.sUMedida = tCusteioRoteiro.sUMedida
'    objCusteioRoteiro.dQuantidade = tCusteioRoteiro.dQuantidade
'    objCusteioRoteiro.dtDataCusteio = tCusteioRoteiro.dtDataCusteio
'    objCusteioRoteiro.dtDataValidade = tCusteioRoteiro.dtDataValidade
'    objCusteioRoteiro.dCustoTotalInsumosKit = tCusteioRoteiro.dCustoTotalInsumosKit
'    objCusteioRoteiro.dCustoTotalInsumosMaq = tCusteioRoteiro.dCustoTotalInsumosMaq
'    objCusteioRoteiro.dCustoTotalMaoDeObra = tCusteioRoteiro.dCustoTotalMaoDeObra
'    objCusteioRoteiro.dPrecoTotalRoteiro = tCusteioRoteiro.dPrecoTotalRoteiro
'    objCusteioRoteiro.sObservacao = tCusteioRoteiro.sObservacao
'
'    lErro = CusteioRoteiro_Le_InsumosKit(objCusteioRoteiro)
'    If lErro <> SUCESSO Then gError 137943
'
'    lErro = CusteioRoteiro_Le_InsumosMaquina(objCusteioRoteiro)
'    If lErro <> SUCESSO Then gError 137944
'
'    lErro = CusteioRoteiro_Le_MaoDeObra(objCusteioRoteiro)
'    If lErro <> SUCESSO Then gError 137945
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    CusteioRoteiro_Le = SUCESSO
'
'    Exit Function
'
'Erro_CusteioRoteiro_Le:
'
'    CusteioRoteiro_Le = gErr
'
'    Select Case gErr
'
'        Case 137939
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 137940, 137941
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTEIOROTEIRO", gErr)
'
'        Case 137942   'sem dados
'
'        Case 137943 To 137945
'            'erros tratados nas rotinas chamadas
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152481)
'
'    End Select
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'
'Public Function CusteioRoteiro_Le_InsumosKit(ByVal objCusteioRoteiro As ClassCusteioRoteiro) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objCusteioRotInsumosKit As ClassCusteioRotInsumosKit
'Dim tCusteioRotInsumosKit As typeCusteioRotInsumosKit
'
'On Error GoTo Erro_CusteioRoteiro_Le_InsumosKit
'
'    'Executa a abertura do Comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 137946
'
'    'Alocação de espaço no buffer
'    tCusteioRotInsumosKit.sProduto = String(STRING_PRODUTO, 0)
'    tCusteioRotInsumosKit.sUMedida = String(STRING_UM_SIGLA, 0)
'    tCusteioRotInsumosKit.sObservacao = String(STRING_PROJETO_OBSERVACAO, 0)
'
'    'Le a tabela CusteioRotInsumosKit
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocCusteioRot, Seq, Produto, UMedida, Quantidade, CustoUnitarioCalculado, CustoUnitarioInformado, Observacao FROM CusteioRotInsumosKit WHERE NumIntDocCusteioRot= ? Order by Seq ", _
'                tCusteioRotInsumosKit.lNumIntDoc, tCusteioRotInsumosKit.lNumIntDocCusteioRot, tCusteioRotInsumosKit.iSeq, tCusteioRotInsumosKit.sProduto, tCusteioRotInsumosKit.sUMedida, tCusteioRotInsumosKit.dQuantidade, tCusteioRotInsumosKit.dCustoUnitarioCalculado, tCusteioRotInsumosKit.dCustoUnitarioInformado, tCusteioRotInsumosKit.sObservacao, objCusteioRoteiro.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 137947
'
'    'Busca Primeiro
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137948
'
'    'Faça enquanto tiver dados
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objCusteioRotInsumosKit = New ClassCusteioRotInsumosKit
'
'        objCusteioRotInsumosKit.lNumIntDoc = tCusteioRotInsumosKit.lNumIntDoc
'        objCusteioRotInsumosKit.lNumIntDocCusteioRot = tCusteioRotInsumosKit.lNumIntDocCusteioRot
'        objCusteioRotInsumosKit.iSeq = tCusteioRotInsumosKit.iSeq
'        objCusteioRotInsumosKit.sProduto = tCusteioRotInsumosKit.sProduto
'        objCusteioRotInsumosKit.sUMedida = tCusteioRotInsumosKit.sUMedida
'        objCusteioRotInsumosKit.dQuantidade = tCusteioRotInsumosKit.dQuantidade
'        objCusteioRotInsumosKit.dCustoUnitarioCalculado = tCusteioRotInsumosKit.dCustoUnitarioCalculado
'        objCusteioRotInsumosKit.dCustoUnitarioInformado = tCusteioRotInsumosKit.dCustoUnitarioInformado
'        objCusteioRotInsumosKit.sObservacao = tCusteioRotInsumosKit.sObservacao
'
'        objCusteioRoteiro.colCusteioRotInsumosKit.Add objCusteioRotInsumosKit
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137949
'
'    Loop
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    CusteioRoteiro_Le_InsumosKit = SUCESSO
'
'    Exit Function
'
'Erro_CusteioRoteiro_Le_InsumosKit:
'
'    CusteioRoteiro_Le_InsumosKit = gErr
'
'    Select Case gErr
'
'        Case 137946
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 137947 To 137949
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTEIOROTINSUMOSKIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152482)
'
'    End Select
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Public Function CusteioRoteiro_Le_InsumosMaquina(ByVal objCusteioRoteiro As ClassCusteioRoteiro) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objCusteioRotInsumosMaq As ClassCusteioRotInsumosMaq
'Dim tCusteioRotInsumosMaq As typeCusteioRotInsumosMaq
'
'On Error GoTo Erro_CusteioRoteiro_Le_InsumosMaquina
'
'    'Executa a abertura do Comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 137950
'
'    'Alocação de espaço no buffer
'    tCusteioRotInsumosMaq.sProduto = String(STRING_PRODUTO, 0)
'    tCusteioRotInsumosMaq.sUMedida = String(STRING_UM_SIGLA, 0)
'    tCusteioRotInsumosMaq.sObservacao = String(STRING_PROJETO_OBSERVACAO, 0)
'
'    'Le a tabela CusteioRotInsumosMaq
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocCusteioRot, Seq, Produto, UMedida, Quantidade, CustoUnitarioCalculado, CustoUnitarioInformado, Observacao FROM CusteioRotInsumosMaq WHERE NumIntDocCusteioRot= ? Order by Seq ", _
'                tCusteioRotInsumosMaq.lNumIntDoc, tCusteioRotInsumosMaq.lNumIntDocCusteioRot, tCusteioRotInsumosMaq.iSeq, tCusteioRotInsumosMaq.sProduto, tCusteioRotInsumosMaq.sUMedida, tCusteioRotInsumosMaq.dQuantidade, tCusteioRotInsumosMaq.dCustoUnitarioCalculado, tCusteioRotInsumosMaq.dCustoUnitarioInformado, tCusteioRotInsumosMaq.sObservacao, objCusteioRoteiro.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 137951
'
'    'Busca Primeiro
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137952
'
'    'Faça enquanto tiver dados
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objCusteioRotInsumosMaq = New ClassCusteioRotInsumosMaq
'
'        objCusteioRotInsumosMaq.lNumIntDoc = tCusteioRotInsumosMaq.lNumIntDoc
'        objCusteioRotInsumosMaq.lNumIntDocCusteioRot = tCusteioRotInsumosMaq.lNumIntDocCusteioRot
'        objCusteioRotInsumosMaq.iSeq = tCusteioRotInsumosMaq.iSeq
'        objCusteioRotInsumosMaq.sProduto = tCusteioRotInsumosMaq.sProduto
'        objCusteioRotInsumosMaq.sUMedida = tCusteioRotInsumosMaq.sUMedida
'        objCusteioRotInsumosMaq.dQuantidade = tCusteioRotInsumosMaq.dQuantidade
'        objCusteioRotInsumosMaq.dCustoUnitarioCalculado = tCusteioRotInsumosMaq.dCustoUnitarioCalculado
'        objCusteioRotInsumosMaq.dCustoUnitarioInformado = tCusteioRotInsumosMaq.dCustoUnitarioInformado
'        objCusteioRotInsumosMaq.sObservacao = tCusteioRotInsumosMaq.sObservacao
'
'        objCusteioRoteiro.colCusteioRotInsumosMaq.Add objCusteioRotInsumosMaq
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137953
'
'    Loop
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    CusteioRoteiro_Le_InsumosMaquina = SUCESSO
'
'    Exit Function
'
'Erro_CusteioRoteiro_Le_InsumosMaquina:
'
'    CusteioRoteiro_Le_InsumosMaquina = gErr
'
'    Select Case gErr
'
'        Case 137950
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 137951 To 137953
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTEIOROTINSUMOSMAQ", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152483)
'
'    End Select
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Public Function CusteioRoteiro_Le_MaoDeObra(ByVal objCusteioRoteiro As ClassCusteioRoteiro) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objCusteioRotMaoDeObra As ClassCusteioRotMaoDeObra
'Dim tCusteioRotMaoDeObra As typeCusteioRotMaoDeObra
'
'On Error GoTo Erro_CusteioRoteiro_Le_MaoDeObra
'
'    'Executa a abertura do Comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 137954
'
'    'Alocação de espaço no buffer
'    tCusteioRotMaoDeObra.sUMedida = String(STRING_UM_SIGLA, 0)
'    tCusteioRotMaoDeObra.sObservacao = String(STRING_PROJETO_OBSERVACAO, 0)
'
'    'Le a tabela CusteioRotMaoDeObra
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocCusteioRot, Seq, CodMO, UMedida, Quantidade, CustoUnitarioCalculado, CustoUnitarioInformado, Observacao FROM CusteioRotMaoDeObra WHERE NumIntDocCusteioRot= ? Order by Seq ", _
'                tCusteioRotMaoDeObra.lNumIntDoc, tCusteioRotMaoDeObra.lNumIntDocCusteioRot, tCusteioRotMaoDeObra.iSeq, tCusteioRotMaoDeObra.iCodMO, tCusteioRotMaoDeObra.sUMedida, tCusteioRotMaoDeObra.dQuantidade, tCusteioRotMaoDeObra.dCustoUnitarioCalculado, tCusteioRotMaoDeObra.dCustoUnitarioInformado, tCusteioRotMaoDeObra.sObservacao, objCusteioRoteiro.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 137955
'
'    'Busca Primeiro
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 137956
'
'    'Faça enquanto tiver dados
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objCusteioRotMaoDeObra = New ClassCusteioRotMaoDeObra
'
'        objCusteioRotMaoDeObra.lNumIntDoc = tCusteioRotMaoDeObra.lNumIntDoc
'        objCusteioRotMaoDeObra.lNumIntDocCusteioRot = tCusteioRotMaoDeObra.lNumIntDocCusteioRot
'        objCusteioRotMaoDeObra.iSeq = tCusteioRotMaoDeObra.iSeq
'        objCusteioRotMaoDeObra.iCodMO = tCusteioRotMaoDeObra.iCodMO
'        objCusteioRotMaoDeObra.sUMedida = tCusteioRotMaoDeObra.sUMedida
'        objCusteioRotMaoDeObra.dQuantidade = tCusteioRotMaoDeObra.dQuantidade
'        objCusteioRotMaoDeObra.dCustoUnitarioCalculado = tCusteioRotMaoDeObra.dCustoUnitarioCalculado
'        objCusteioRotMaoDeObra.dCustoUnitarioInformado = tCusteioRotMaoDeObra.dCustoUnitarioInformado
'        objCusteioRotMaoDeObra.sObservacao = tCusteioRotMaoDeObra.sObservacao
'
'        objCusteioRoteiro.colCusteioRotMaoDeObra.Add objCusteioRotMaoDeObra
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 137957
'
'    Loop
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    CusteioRoteiro_Le_MaoDeObra = SUCESSO
'
'    Exit Function
'
'Erro_CusteioRoteiro_Le_MaoDeObra:
'
'    CusteioRoteiro_Le_MaoDeObra = gErr
'
'    Select Case gErr
'
'        Case 137954
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 137955 To 137957
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CUSTEIOROTMAODEOBRA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152484)
'
'    End Select
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function CentrodeTrabalho_Le_CTOperadores(ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long
Dim lComando As Long
Dim objCTOperadores As ClassCTOperadores
Dim tCTOperadores As typeCTOperadores

On Error GoTo Erro_CentrodeTrabalho_Le_CTOperadores

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 139097

    lErro = Comando_Executar(lComando, "SELECT TipoMaoDeObra, NumIntDocCT, Quantidade FROM CTOperadores WHERE NumIntDocCT = ? Order by TipoMaoDeObra", tCTOperadores.iCodTipoMO, tCTOperadores.lNumIntDocCT, tCTOperadores.iQuantidade, objCentrodeTrabalho.lNumIntDoc)
    If lErro <> SUCESSO Then gError 139098
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 139099

    'CTOperadores não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then gError 139100

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCTOperadores = New ClassCTOperadores

        objCTOperadores.iCodTipoMO = tCTOperadores.iCodTipoMO
        objCTOperadores.lNumIntDocCT = tCTOperadores.lNumIntDocCT
        objCTOperadores.iQuantidade = tCTOperadores.iQuantidade
        
        objCentrodeTrabalho.colOperadores.Add objCTOperadores

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 139101

    Loop
    
    Call Comando_Fechar(lComando)
    
    CentrodeTrabalho_Le_CTOperadores = SUCESSO

    Exit Function

Erro_CentrodeTrabalho_Le_CTOperadores:

   CentrodeTrabalho_Le_CTOperadores = gErr

   Select Case gErr
    
        Case 139097
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 139098, 139099, 139101
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTOPERADORES", gErr)
            
        Case 139100
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152485)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'##########################################################################
'Inserido por Wagner 07/11/2005
'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function CTMaquinasParadas_Le(ByVal objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCTMaquinasParadas As typeCTMaquinasParadas

On Error GoTo Erro_CTMaquinasParadas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140701

    'Alocação de espaço no buffer
    tCTMaquinasParadas.sObservacao = String(STRING_OBSERVACAO, 0)

    'Le a tabela CTMaquinasParadas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, FilialEmpresa, Data, NumIntDocCT, NumIntDocMaq, " & _
                "Tipo, Horas, QtdMaquinas, Observacao FROM CTMaquinasParadas WHERE Codigo= ?  AND FilialEmpresa= ? ", _
                tCTMaquinasParadas.lNumIntDoc, tCTMaquinasParadas.lCodigo, tCTMaquinasParadas.iFilialEmpresa, tCTMaquinasParadas.dtData, _
                tCTMaquinasParadas.lNumIntDocCT, tCTMaquinasParadas.lNumIntDocMaq, tCTMaquinasParadas.iTipo, tCTMaquinasParadas.dHoras, tCTMaquinasParadas.iQtdMaquinas, tCTMaquinasParadas.sObservacao, _
                objCTMaquinasParadas.lCodigo, objCTMaquinasParadas.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 140702

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140703

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 140704

    objCTMaquinasParadas.lNumIntDoc = tCTMaquinasParadas.lNumIntDoc
    objCTMaquinasParadas.lCodigo = tCTMaquinasParadas.lCodigo
    objCTMaquinasParadas.iFilialEmpresa = tCTMaquinasParadas.iFilialEmpresa
    objCTMaquinasParadas.dtData = tCTMaquinasParadas.dtData
    objCTMaquinasParadas.lNumIntDocCT = tCTMaquinasParadas.lNumIntDocCT
    objCTMaquinasParadas.lNumIntDocMaq = tCTMaquinasParadas.lNumIntDocMaq
    objCTMaquinasParadas.iTipo = tCTMaquinasParadas.iTipo
    objCTMaquinasParadas.dHoras = tCTMaquinasParadas.dHoras
    objCTMaquinasParadas.iQtdMaquinas = tCTMaquinasParadas.iQtdMaquinas
    objCTMaquinasParadas.sObservacao = tCTMaquinasParadas.sObservacao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CTMaquinasParadas_Le = SUCESSO

    Exit Function

Erro_CTMaquinasParadas_Le:

    CTMaquinasParadas_Le = gErr

    Select Case gErr

        Case 140701
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140702, 140703
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTMAQUINASPARADAS", gErr)

        Case 140704 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152486)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ItemOPOperacoesMaquinas_Le(ByVal objOPOperacao As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOPOperacoesMaquinas As typeItemOPOperacoesMaquinas
Dim objItemOPOperacoesMaquinas As ClassItemOPOperacoesMaquinas

On Error GoTo Erro_ItemOPOperacoesMaquinas_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140724

    'Alocação de espaço no buffer

    'Le a tabelaItemOPOperacoesMaquinas
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOper, NumIntDocMaq, Data, " & _
                "Horas, Quantidade FROM ItemOPOperacoesMaquinas WHERE NumIntDocOper = ? ", _
                tItemOPOperacoesMaquinas.lNumIntDoc, tItemOPOperacoesMaquinas.lNumIntDocOper, _
                tItemOPOperacoesMaquinas.lNumIntDocMaq, tItemOPOperacoesMaquinas.dtData, tItemOPOperacoesMaquinas.dHoras, tItemOPOperacoesMaquinas.iQuantidade, _
                objOPOperacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 140725

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140726

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOPOperacoesMaquinas = New ClassItemOPOperacoesMaquinas

        objItemOPOperacoesMaquinas.lNumIntDoc = tItemOPOperacoesMaquinas.lNumIntDoc
        objItemOPOperacoesMaquinas.lNumIntDocOper = tItemOPOperacoesMaquinas.lNumIntDocOper
        objItemOPOperacoesMaquinas.lNumIntDocMaq = tItemOPOperacoesMaquinas.lNumIntDocMaq
        objItemOPOperacoesMaquinas.dtData = tItemOPOperacoesMaquinas.dtData
        objItemOPOperacoesMaquinas.dHoras = tItemOPOperacoesMaquinas.dHoras
        objItemOPOperacoesMaquinas.iQuantidade = tItemOPOperacoesMaquinas.iQuantidade
        
        lErro = ItemOPOperacoesMO_Le(objItemOPOperacoesMaquinas)
        If lErro <> SUCESSO Then gError 140771
        
        objOPOperacao.colUsoMaquinas.Add objItemOPOperacoesMaquinas
        
        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140770
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOPOperacoesMaquinas_Le = SUCESSO

    Exit Function

Erro_ItemOPOperacoesMaquinas_Le:

    ItemOPOperacoesMaquinas_Le = gErr

    Select Case gErr

        Case 140724
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140725, 140726, 140770
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPOPERACOESMAQUINAS", gErr)

        Case 140727 'Sem dados -> Tratado na rotina chamadora
        
        Case 140771

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152487)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ItemOPOperacoesMO_Le(ByVal objItemOPOperacoesMaquina As ClassItemOPOperacoesMaquinas) As Long

Dim lErro As Long
Dim lComando As Long
Dim tItemOPOperacoesMO As typeItemOPOperacoesMO
Dim objItemOPOperacoesMO As ClassItemOPOperacoesMO

On Error GoTo Erro_ItemOPOperacoesMO_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 140747

    'Alocação de espaço no buffer

    'Le a tabelaItemOPOperacoesMO
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocItemOPMaq, TipoMO, Horas, Quantidade FROM ItemOPOperacoesMO WHERE NumIntDocItemOPMaq= ?  ", _
                tItemOPOperacoesMO.lNumIntDoc, tItemOPOperacoesMO.lNumIntDocItemOPMaq, tItemOPOperacoesMO.iTipoMO, tItemOPOperacoesMO.dHoras, tItemOPOperacoesMO.iQuantidade, _
                objItemOPOperacoesMaquina.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 140748

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140749

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objItemOPOperacoesMO = New ClassItemOPOperacoesMO

        objItemOPOperacoesMO.lNumIntDoc = tItemOPOperacoesMO.lNumIntDoc
        objItemOPOperacoesMO.lNumIntDocItemOPMaq = tItemOPOperacoesMO.lNumIntDocItemOPMaq
        objItemOPOperacoesMO.iTipoMO = tItemOPOperacoesMO.iTipoMO
        objItemOPOperacoesMO.dHoras = tItemOPOperacoesMO.dHoras
        objItemOPOperacoesMO.iQuantidade = tItemOPOperacoesMO.iQuantidade
        
        objItemOPOperacoesMaquina.colMO.Add objItemOPOperacoesMO
        
        'Busca Primeiro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 140772
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ItemOPOperacoesMO_Le = SUCESSO

    Exit Function

Erro_ItemOPOperacoesMO_Le:

    ItemOPOperacoesMO_Le = gErr

    Select Case gErr

        Case 140747
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 140748, 140749, 140772
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMOPOPERACOESMO", gErr)

        Case 140750 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152488)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'##########################################################################

Public Function CTMaquina_Le_Faixa(ByVal viFilialEmpresa As Variant, ByVal viCodMaqIni As Variant, ByVal viCodMaqFim As Variant, ByVal vlCodCTIni As Variant, ByVal vlCodCTFim As Variant, ByVal colCTMaquina As Collection) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vlNumIntDocMaq As Variant
Dim vlNumIntDocCT As Variant
Dim alComando(1 To 1) As Long
Dim objCTMaquina As ClassCTMaquinas
 
On Error GoTo Erro_CTMaquina_Le_Faixa
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141097
    Next
    
    vlNumIntDocMaq = 0
    vlNumIntDocCT = 0
    
    sSQL = "SELECT CTM.NumIntDocCT, CTM.NumIntDocMaq FROM CTMaquinas AS CTM, Maquinas AS M, CentrodeTrabalho AS CT WHERE CTM.NumIntDocCT = CT.NumIntDoc AND CTM.NumIntDocMaq = M.NumIntDoc AND CT.FilialEmpresa = ? "
 
    If viCodMaqIni <> 0 Then
        sWhere = sWhere & "AND M.Codigo >= ? "
    End If
    If viCodMaqFim <> 0 Then
        sWhere = sWhere & "AND M.Codigo <= ? "
    End If
 
    If vlCodCTIni <> 0 Then
        sWhere = sWhere & "AND CT.Codigo >= ? "
    End If
    If vlCodCTFim <> 0 Then
        sWhere = sWhere & "AND CT.Codigo <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY CT.Codigo, M.Codigo "
     
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141098
 
    lErro = Comando_BindVarInt(alComando(1), vlNumIntDocCT)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141099
 
    lErro = Comando_BindVarInt(alComando(1), vlNumIntDocMaq)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141100
 
    lErro = Comando_BindVarInt(alComando(1), viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141101
 
    If viCodMaqIni <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), viCodMaqIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 141102
    End If
    
    If viCodMaqFim <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), viCodMaqFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 141103
    End If
    
    If vlCodCTIni <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlCodCTIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 141104
    End If
    
    If vlCodCTFim <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), vlCodCTFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 141105
    End If
 
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 141106
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141107
        
    'Para cada CTMaquina
    Do While lErro = SUCESSO
    
        Set objCTMaquina = New ClassCTMaquinas
 
        objCTMaquina.lNumIntDocCT = vlNumIntDocCT
        objCTMaquina.lNumIntDocMaq = vlNumIntDocMaq
        
        colCTMaquina.Add objCTMaquina
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141108
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    CTMaquina_Le_Faixa = SUCESSO
 
    Exit Function
 
Erro_CTMaquina_Le_Faixa:
 
    CTMaquina_Le_Faixa = gErr
 
    Select Case gErr
 
        Case 141097
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 141098 To 141108
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152489)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

Function NecessidadeRoteiro_Le(ByVal objRoteiroNecessidade As ClassRoteiroNecessidade, Optional ByVal bRecursivo As Boolean = False, Optional iNiveis As Integer = 0) As Long

Dim lErro As Long
Dim dFator As Double
Dim dFatorQuantidade As Double
Dim dFatorTempo As Double
Dim objProduto As ClassProduto
Dim objProdutoInsumo As ClassProduto
Dim objRoteirosBD As ClassRoteirosDeFabricacao
Dim objOperacoes As New ClassOperacoes
Dim objPO As ClassPlanoOperacional
Dim objOPOperacoes As ClassOrdemProducaoOperacoes
Dim objPOMaquinas As New ClassPOMaquinas
Dim objMaquinas As ClassMaquinas
Dim objMaquinasInsumos As New ClassMaquinasInsumos
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objPMPItens As ClassPMPItens, objPMPAux As New ClassPMP
Dim objMaodeObra As ClassTiposDeMaodeObra
Dim objMaqOperador As ClassMaquinaOperadores
Dim objRoteiroMP As ClassRoteiroMP
Dim objRoteiroMaoDeObra As ClassRoteiroMaoDeObra
Dim objRoteiroMaquina As ClassRoteiroMaquina
Dim objRoteiroInsumosMaq As ClassRoteiroInsumosMaq
Dim colMaquinas As New Collection
Dim bAchou As Boolean
Dim objRoteiroNecessidadeFilho As ClassRoteiroNecessidade
Dim objKit As New ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim objKitRec As New ClassKit

On Error GoTo Erro_NecessidadeRoteiro_Le

    'Se a quantidade já chegou a 0.1 ou já desceu na recursividade n instâncias
    If (objRoteiroNecessidade.dQuantidade <= QTDE_ESTOQUE_DELTA And iNiveis > 0) Or iNiveis >= 20 Then
    
        NecessidadeRoteiro_Le = SUCESSO
        Exit Function

    End If
    
    Set objRoteirosBD = New ClassRoteirosDeFabricacao
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objRoteiroNecessidade.sProdutoRaiz

    'Lê o produto para descobrir as unidades de medidas associadas
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 141632
      
    objRoteirosBD.sProdutoRaiz = objRoteiroNecessidade.sProdutoRaiz
    objRoteirosBD.sVersao = objRoteiroNecessidade.sVersao
    
    'Le o roteiro de fabricação
    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosBD)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 141631
    
    If lErro = SUCESSO Then
    
        'Descobre o fator de conversão
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objRoteiroNecessidade.sUM, objRoteirosBD.sUM, dFator)
        If lErro <> SUCESSO Then gError 141633
    
        dFatorQuantidade = (objRoteiroNecessidade.dQuantidade * dFator) / objRoteirosBD.dQuantidade
        
        Set objPO = New ClassPlanoOperacional
        
        objPO.sProduto = objRoteiroNecessidade.sProdutoRaiz
        objPO.dQuantidade = objRoteiroNecessidade.dQuantidade
        objPO.sUM = objRoteiroNecessidade.sUM
        objPO.dtDataInicio = gdtDataAtual
        
        For Each objOperacoes In objRoteirosBD.colOperacoes
        
            'Preenche as informações necessária parta executar o cálculo de tempo
            Set objOPOperacoes = New ClassOrdemProducaoOperacoes
            
            objOPOperacoes.lNumIntDocCT = objOperacoes.lNumIntDocCT
            objOPOperacoes.lNumIntDocCompet = objOperacoes.lNumIntDocCompet
            objOPOperacoes.iConsideraCarga = MARCADO
            
            Set objOPOperacoes.objOperacoesTempo = objOperacoes.objOperacoesTempo
            
            Set objPMPItens = New ClassPMPItens
                        
            'Calcula o que está sendo gasto para fabricar o produto
            lErro = CF("PlanoOperacional_Calcula_Tempos", objPMPAux, objPMPItens, objPO, objOPOperacoes, MRP_ACERTA_POR_DATA_INICIO)
            If lErro <> SUCESSO Then gError 141634
            
            'Para cada máquina alocada
            For Each objPOMaquinas In objPO.colAlocacaoMaquinas
                
                bAchou = False
                For Each objMaquinas In colMaquinas
                
                    If objMaquinas.lNumIntDoc = objPOMaquinas.lNumIntDocMaq Then
                        bAchou = True
                        Exit For
                    End If
                
                Next
            
                If Not bAchou Then
                
                    Set objMaquinas = New ClassMaquinas
                
                    objMaquinas.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
                    
                    lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
                    If lErro <> SUCESSO Then gError 141635
                 
                    lErro = CF("Maquinas_Le_Itens", objMaquinas)
                    If lErro <> SUCESSO Then gError 141635
                    
                    colMaquinas.Add objMaquinas
                    
                End If
                
                For Each objRoteiroMaquina In objRoteiroNecessidade.colMaquinas
                
                    If objRoteiroMaquina.iMaquina = objMaquinas.iCodigo Then
                        
                        'Pega o maior número de máquinas usadas
                        If objRoteiroMaquina.iQuantidade < objPOMaquinas.iQuantidade Then
                            objRoteiroMaquina.iQuantidade = objPOMaquinas.iQuantidade
                        End If
                        
                        objRoteiroMaquina.dHoras = objRoteiroMaquina.dHoras + objPOMaquinas.dHorasMaquina
                        objRoteiroMaquina.dCustoTotal = objRoteiroMaquina.dCustoUnitario * objRoteiroMaquina.dHoras
                    
                        bAchou = True
                        
                        Exit For
                        
                    End If
                
                Next
            
                If Not bAchou Then
                
                    Set objRoteiroMaquina = New ClassRoteiroMaquina
                    
                    objRoteiroMaquina.dHoras = objPOMaquinas.dHorasMaquina
                    objRoteiroMaquina.dCustoTotal = objMaquinas.dCustoHora * objPOMaquinas.dHorasMaquina
                    objRoteiroMaquina.dCustoUnitario = objMaquinas.dCustoHora
                    objRoteiroMaquina.iQuantidade = objPOMaquinas.iQuantidade
                    objRoteiroMaquina.iMaquina = objMaquinas.iCodigo
                    
                    Set objRoteiroMaquina.objMaquina = objMaquinas
                
                    objRoteiroNecessidade.colMaquinas.Add objRoteiroMaquina
                    
                End If
                
                For Each objMaquinasInsumos In objMaquinas.colProdutos
                
                    bAchou = False
                    
                    For Each objRoteiroInsumosMaq In objRoteiroNecessidade.colInsumosMaquina
                    
                        If objRoteiroInsumosMaq.sProduto = objMaquinasInsumos.sProduto Then
                        
                            'Descobre o fator de conversão
                            lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, "h", objMaquinasInsumos.sUMTempo, dFatorTempo)
                            If lErro <> SUCESSO Then gError 141636
                            
                            objRoteiroInsumosMaq.dQuantidade = objRoteiroInsumosMaq.dQuantidade + objMaquinasInsumos.dQuantidade * objPO.dTempoGasto / dFatorTempo
                        
                            bAchou = True
                            
                            Exit For
                            
                        End If
                    
                    Next
                
                    If Not bAchou Then
                    
                        Set objRoteiroInsumosMaq = New ClassRoteiroInsumosMaq
                        
                        'Descobre o fator de conversão
                        lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, "h", objMaquinasInsumos.sUMTempo, dFatorTempo)
                        If lErro <> SUCESSO Then gError 141654
                        
                        objRoteiroInsumosMaq.dQuantidade = objMaquinasInsumos.dQuantidade * objPOMaquinas.dHorasMaquina / dFatorTempo
                        objRoteiroInsumosMaq.sProduto = objMaquinasInsumos.sProduto
                        objRoteiroInsumosMaq.sUM = objMaquinasInsumos.sUMProduto
                            
                        objRoteiroNecessidade.colInsumosMaquina.Add objRoteiroInsumosMaq
                        
                    End If
                                
                Next
                
                For Each objMaqOperador In objMaquinas.colTipoOperadores
                
                    bAchou = False
                    
                    For Each objRoteiroMaoDeObra In objRoteiroNecessidade.colMaoDeObra
                    
                        If objRoteiroMaoDeObra.iCodMO = objMaqOperador.iTipoMaoDeObra Then
                        
                            If objRoteiroMaoDeObra.iQuantidade < objMaqOperador.iQuantidade * objPOMaquinas.iQuantidade Then
                                objRoteiroMaoDeObra.iQuantidade = objMaqOperador.iQuantidade * objPOMaquinas.iQuantidade
                            End If
                            objRoteiroMaoDeObra.dHoras = objRoteiroMaoDeObra.dHoras + objMaqOperador.iQuantidade * objMaqOperador.dPercentualUso * objPOMaquinas.dHorasMaquina
                            objRoteiroMaoDeObra.dCustoTotal = objRoteiroMaoDeObra.dCustoUnitario * objRoteiroMaoDeObra.dHoras
                        
                            bAchou = True
                            
                            Exit For
                            
                        End If
                    
                    Next
                
                    If Not bAchou Then
                    
                        Set objRoteiroMaoDeObra = New ClassRoteiroMaoDeObra
                        Set objMaodeObra = New ClassTiposDeMaodeObra
                        
                        objMaodeObra.iCodigo = objMaqOperador.iTipoMaoDeObra
                        
                        lErro = TiposDeMaodeObra_Le(objMaodeObra)
                        If lErro <> SUCESSO And lErro <> 137598 Then gError 141637
                        
                        objRoteiroMaoDeObra.dHoras = objMaqOperador.iQuantidade * objMaqOperador.dPercentualUso * objPOMaquinas.dHorasMaquina
                        objRoteiroMaoDeObra.dCustoTotal = objMaodeObra.dCustoHora * objRoteiroMaoDeObra.dHoras
                        objRoteiroMaoDeObra.dCustoUnitario = objMaodeObra.dCustoHora
                        objRoteiroMaoDeObra.iCodMO = objMaodeObra.iCodigo
                        objRoteiroMaoDeObra.sUM = "h"
                        objRoteiroMaoDeObra.iQuantidade = objMaqOperador.iQuantidade * objPOMaquinas.iQuantidade
                    
                        Set objRoteiroMaoDeObra.objMaodeObra = objMaodeObra
                    
                        objRoteiroNecessidade.colMaoDeObra.Add objRoteiroMaoDeObra
                        
                    End If
                
                Next
            
            Next
            
            For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
            
                bAchou = False
                
                Set objProdutoInsumo = New ClassProduto
               
                objProdutoInsumo.sCodigo = objOperacaoInsumos.sProduto
            
                'Lê o produto para descobrir as unidades de medidas associadas
                lErro = CF("Produto_Le", objProdutoInsumo)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 141655
                            
                For Each objRoteiroMP In objRoteiroNecessidade.colMP
                
                    If objRoteiroMP.sProduto = objOperacaoInsumos.sProduto And objRoteiroMP.sVersao = objOperacaoInsumos.sVersaoKitComp Then
                        
                        'Descobre o fator de conversão
                        lErro = CF("UM_Conversao_Trans", objProdutoInsumo.iClasseUM, objOperacaoInsumos.sUMProduto, objRoteiroMP.sUM, dFator)
                        If lErro <> SUCESSO Then gError 141656
                        
                        objRoteiroMP.dQuantidade = objRoteiroMP.dQuantidade + objOperacaoInsumos.dQuantidade * dFatorQuantidade * dFator
                    
                        bAchou = True
                        
                        Exit For
                        
                    End If
                
                Next
            
                'Se não encontrou o produto na coleção
                'E o produto é comprável ou o algoritimo não é recursivo
                If Not bAchou And ((Not bRecursivo) Or objProdutoInsumo.iCompras = PRODUTO_COMPRAVEL) Then
                    
                    Set objRoteiroMP = New ClassRoteiroMP
                    
                    objRoteiroMP.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQuantidade / (1 - objOperacaoInsumos.dPercentualPerda)
                    objRoteiroMP.sProduto = objOperacaoInsumos.sProduto
                    objRoteiroMP.sUM = objOperacaoInsumos.sUMProduto
                    objRoteiroMP.sVersao = objOperacaoInsumos.sVersaoKitComp
                
                    objRoteiroNecessidade.colMP.Add objRoteiroMP
                    
                End If
                
                If bRecursivo And objProdutoInsumo.iCompras = PRODUTO_PRODUZIVEL Then
                
                    Set objRoteiroNecessidadeFilho = New ClassRoteiroNecessidade
                
                    objRoteiroNecessidadeFilho.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQuantidade / (1 - objOperacaoInsumos.dPercentualPerda)
                    objRoteiroNecessidadeFilho.iFilialEmpresa = objRoteiroNecessidade.iFilialEmpresa
                    objRoteiroNecessidadeFilho.sProdutoRaiz = objOperacaoInsumos.sProduto
                    objRoteiroNecessidadeFilho.sUM = objOperacaoInsumos.sUMProduto
                    objRoteiroNecessidadeFilho.sVersao = objOperacaoInsumos.sVersaoKitComp
                    
                    Set objRoteiroNecessidadeFilho.colInsumosMaquina = objRoteiroNecessidade.colInsumosMaquina
                    Set objRoteiroNecessidadeFilho.colMaoDeObra = objRoteiroNecessidade.colMaoDeObra
                    Set objRoteiroNecessidadeFilho.colMP = objRoteiroNecessidade.colMP
                    Set objRoteiroNecessidadeFilho.colMaquinas = objRoteiroNecessidade.colMaquinas
                    
                    lErro = NecessidadeRoteiro_Le(objRoteiroNecessidadeFilho, bRecursivo, iNiveis + 1)
                    If lErro <> SUCESSO Then gError 141960
                
                End If
                
            Next
        
        Next
        
    Else
        
        objKit.sProdutoRaiz = objRoteiroNecessidade.sProdutoRaiz
        objKit.sVersao = objRoteiroNecessidade.sVersao
            
        lErro = ProdutoKit_Le_PrimeiroNivel_Versao(objKit)
        If lErro <> SUCESSO And lErro <> 103203 Then gError 141961

        For Each objProdutoKit In objKit.colComponentes
        
            If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                
                'Descobre o fator de conversão
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objRoteiroNecessidade.sUM, objProdutoKit.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 141962
                            
                dFatorQuantidade = (objRoteiroNecessidade.dQuantidade * dFator) / objProdutoKit.dQuantidade
            
            End If
        
            bAchou = False
            
            Set objProdutoInsumo = New ClassProduto
           
            objProdutoInsumo.sCodigo = objProdutoKit.sProduto
        
            'Lê o produto para descobrir as unidades de medidas associadas
            lErro = CF("Produto_Le", objProdutoInsumo)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 141963
        
            ''Descobre o fator de conversão
            'lErro = CF("UM_Conversao_Trans", objProdutoInsumo.iClasseUM, objProdutoKit.sUnidadeMed, objRoteiroMP.sUM, dFator)
            'If lErro <> SUCESSO Then gError 141964
            
            For Each objRoteiroMP In objRoteiroNecessidade.colMP
            
                If objRoteiroMP.sProduto = objProdutoKit.sProduto Then
                    
                    'Descobre o fator de conversão
                    lErro = CF("UM_Conversao_Trans", objProdutoInsumo.iClasseUM, objProdutoKit.sUnidadeMed, objRoteiroMP.sUM, dFator)
                    If lErro <> SUCESSO Then gError 141964
                    
                    objRoteiroMP.dQuantidade = objRoteiroMP.dQuantidade + objProdutoKit.dQuantidade * dFatorQuantidade * dFator
                
                    bAchou = True
                    
                    Exit For
                    
                End If
            
            Next
        
            'Se não encontrou o produto na coleção
            'E o produto é comprável ou o algoritimo não é recursivo
            If Not bAchou And (Not bRecursivo Or objProdutoInsumo.iCompras <> PRODUTO_PRODUZIVEL) And objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                
                Set objRoteiroMP = New ClassRoteiroMP
                
                objRoteiroMP.dQuantidade = objProdutoKit.dQuantidade * dFatorQuantidade / (1 - objProdutoKit.dPercentualPerda)
                objRoteiroMP.sProduto = objProdutoKit.sProduto
                objRoteiroMP.sUM = objProdutoKit.sUnidadeMed
            
                objRoteiroNecessidade.colMP.Add objRoteiroMP
                
            End If
            
            If bRecursivo And objProdutoInsumo.iCompras = PRODUTO_PRODUZIVEL And objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
            
                Set objRoteiroNecessidadeFilho = New ClassRoteiroNecessidade
            
                objRoteiroNecessidadeFilho.dQuantidade = objProdutoKit.dQuantidade * dFatorQuantidade / (1 - objProdutoKit.dPercentualPerda)
                objRoteiroNecessidadeFilho.iFilialEmpresa = objRoteiroNecessidade.iFilialEmpresa
                objRoteiroNecessidadeFilho.sProdutoRaiz = objProdutoKit.sProduto
                objRoteiroNecessidadeFilho.sUM = objProdutoKit.sUnidadeMed
                objRoteiroNecessidadeFilho.sVersao = objProdutoKit.sVersaoKitComp
                
                If objRoteiroNecessidadeFilho.sVersao = "" Then
                    
                    Set objKitRec = New ClassKit
                    
                    objKitRec.sProdutoRaiz = objProdutoKit.sProduto
                    
                    lErro = CF("Kit_Le_Padrao", objKitRec)
                    If lErro <> SUCESSO And lErro <> 106304 Then gError 141965
                    
                    objRoteiroNecessidadeFilho.sVersao = objKitRec.sVersao
                
                End If
                
                Set objRoteiroNecessidadeFilho.colInsumosMaquina = objRoteiroNecessidade.colInsumosMaquina
                Set objRoteiroNecessidadeFilho.colMaoDeObra = objRoteiroNecessidade.colMaoDeObra
                Set objRoteiroNecessidadeFilho.colMP = objRoteiroNecessidade.colMP
                
                lErro = NecessidadeRoteiro_Le(objRoteiroNecessidadeFilho, bRecursivo, iNiveis + 1)
                If lErro <> SUCESSO Then gError 141965
            
            End If
            
        Next
    
    End If
    
    NecessidadeRoteiro_Le = SUCESSO
    
    Exit Function

Erro_NecessidadeRoteiro_Le:

    NecessidadeRoteiro_Le = gErr
    
    Select Case gErr
    
        Case 141631 To 141637, 141654, 141655, 141656, 141960 To 141965
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152490)
    
    End Select
    
    Exit Function

End Function

'########################################################################################
'Inserido por Wagner 06/02/2006
Public Function Roteiro_Le_Faixa(ByVal vsProdIni As Variant, ByVal vsProdFim As Variant, ByVal colRoteiro As Collection) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim vsProduto As Variant
Dim vsVersao As Variant
Dim alComando(1 To 1) As Long
Dim objRoteiro As ClassRoteirosDeFabricacao
 
On Error GoTo Erro_Roteiro_Le_Faixa
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141657
    Next
    
    sSQL = "SELECT ProdutoRaiz, Versao FROM RoteirosDeFabricacao "
 
    If vsProdIni <> "" Then
        sWhere = sWhere & "WHERE Produto >= ? "
    End If
    If vsProdFim <> "" Then
        If Len(Trim(sWhere)) <> 0 Then
            sWhere = sWhere & "AND "
        Else
            sWhere = "WHERE "
        End If
        sWhere = sWhere & "Produto <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY Produto, Versao "
    
    vsProduto = String(STRING_PRODUTO, 0)
    vsVersao = String(STRING_KIT_VERSAO, 0)
 
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141658
 
    lErro = Comando_BindVarInt(alComando(1), vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141659
 
    lErro = Comando_BindVarInt(alComando(1), vsVersao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 141660
 
    If vsProdIni <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsProdIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 141661
    End If
    
    If vsProdFim <> "" Then
        lErro = Comando_BindVarInt(alComando(1), vsProdFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 141662
    End If
 
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 141663
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141664
        
    'Para cada roteiro
    Do While lErro = SUCESSO
    
        Set objRoteiro = New ClassRoteirosDeFabricacao
 
        objRoteiro.sProdutoRaiz = vsProduto
        objRoteiro.sVersao = vsVersao
        
        'Lê o roteiro
        lErro = RoteirosDeFabricacao_Le(objRoteiro)
        If lErro <> SUCESSO And lErro <> 134617 Then gError 141665
     
        colRoteiro.Add objRoteiro
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141666
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Roteiro_Le_Faixa = SUCESSO
 
    Exit Function
 
Erro_Roteiro_Le_Faixa:
 
    Roteiro_Le_Faixa = gErr
 
    Select Case gErr
 
        Case 141657
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 141658 To 141665
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case 141666
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152491)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function
'########################################################################################

Function CustoRoteiroFabricacao(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal dQuantidade As Double, objRoteiroNecessidade As ClassRoteiroNecessidade, ByVal iRotinaOrigem As Integer) As Long

Dim lErro As Long, bRecursivo As Boolean
Dim lComando As Long, sVersao As String, sUM As String

On Error GoTo Erro_CustoRoteiroFabricacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 13480
    
    'obter a versao para formacao de precos
    sVersao = String(STRING_KIT_VERSAO, 0)
    sUM = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(lComando, "SELECT Versao, SiglaUMEstoque FROM Kit, Produtos WHERE Kit.ProdutoRaiz = Produtos.Codigo AND ProdutoRaiz = ? AND VersaoFormPreco = ?", _
        sVersao, sUM, sProduto, 1)
    If lErro <> AD_SQL_SUCESSO Then gError 130481
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130482
    
    If lErro = AD_SQL_SUCESSO Then
    
        With objRoteiroNecessidade
            .sProdutoRaiz = sProduto
            .sVersao = sVersao
            .dQuantidade = dQuantidade
            .iFilialEmpresa = iFilialEmpresa
            .sUM = sUM
        End With
            
        bRecursivo = True
        lErro = CF("NecessidadeRoteiro_Le", objRoteiroNecessidade, bRecursivo)
        If lErro <> SUCESSO Then gError 130483
    
    End If
    
    Call Comando_Fechar(lComando)
    
    CustoRoteiroFabricacao = SUCESSO
     
    Exit Function
    
Erro_CustoRoteiroFabricacao:

    CustoRoteiroFabricacao = gErr
     
    Select Case gErr
          
        Case 130480
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130481, 130482
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTO_KIT_FABR", gErr)
        
        Case 130483
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130484)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function CustoRoteiro_InsumosMaquina(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Produto
'retorna o custo dos insumos das maquinas necessarios para fabricar o produto,
'incluindo os insumos para os subprodutos, recursivamente

Dim lErro As Long, sProduto As String
Dim dCustoProd As Double
Dim dCusto As Double
Dim objRoteiroInsumosMaq As ClassRoteiroInsumosMaq
Dim lComando As Long
Dim objRoteiroNecessidade As New ClassRoteiroNecessidade

On Error GoTo Erro_CustoRoteiro_InsumosMaquina

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130476
    
    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    lErro = CF("CustoRoteiroFabricacao", objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno, sProduto, objExeExp.objContexto.dQuantidade, objRoteiroNecessidade, objExeExp.objContexto.iRotinaOrigem)
    If lErro <> SUCESSO Then gError 130477
    
    'para cada materia prima
    For Each objRoteiroInsumosMaq In objRoteiroNecessidade.colInsumosMaquina
            
        'obter o seu custo em custembmp.
        lErro = CF("CustoDiretoProd_ObtemValor", objExeExp.objContexto.iFilialFaturamento, objRoteiroInsumosMaq.sProduto, objExeExp.objContexto.dTaxaDescPadrao, objExeExp.objContexto.dTaxaValFut, objExeExp.objContexto.dDiasValFut, dCusto, lComando)
        If lErro <> SUCESSO Then gError 130478
        
        'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
        dCustoProd = dCustoProd + (dCusto * objRoteiroInsumosMaq.dQuantidade)
    
    Next
    
    Call Comando_Fechar(lComando)
    
    objMnemonicoValor.colValor.Add dCustoProd
    
    CustoRoteiro_InsumosMaquina = SUCESSO

    Exit Function

Erro_CustoRoteiro_InsumosMaquina:

    CustoRoteiro_InsumosMaquina = gErr

    Select Case gErr

        Case 130476
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130477, 130478
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130479)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CustoRoteiro_MaoDeObra(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal objExeExp As ClassExeExp) As Long
'parametros:Produto
'retorna o custo dos insumos das maquinas necessarios para fabricar o produto,
'incluindo os insumos para os subprodutos, recursivamente

Dim lErro As Long, sProduto As String
Dim objRoteiroMaoDeObra As ClassRoteiroMaoDeObra, dCustoProd As Double
Dim dCustoMO As Double
Dim lComando As Long
Dim objRoteiroNecessidade As New ClassRoteiroNecessidade

On Error GoTo Erro_CustoRoteiro_MaoDeObra

    sProduto = objExeExp.objContexto.sProduto
    If sProduto = "" Then sProduto = objMnemonicoValor.vParam(1)

    lErro = CF("CustoRoteiroFabricacao", objExeExp.objContexto.iFilialFaturamento, objExeExp.objContexto.iAno, sProduto, objExeExp.objContexto.dQuantidade, objRoteiroNecessidade, objExeExp.objContexto.iRotinaOrigem)
    If lErro <> SUCESSO Then gError 130485
        
    'para cada mao de obra
    For Each objRoteiroMaoDeObra In objRoteiroNecessidade.colMaoDeObra
        
        'acumulo o custo
        dCustoProd = dCustoProd + (objRoteiroMaoDeObra.dCustoTotal)
    
    Next
    
    objMnemonicoValor.colValor.Add dCustoProd
    
    CustoRoteiro_MaoDeObra = SUCESSO

    Exit Function

Erro_CustoRoteiro_MaoDeObra:

    CustoRoteiro_MaoDeObra = gErr

    Select Case gErr

        Case 130485
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130486)

    End Select

    Exit Function

End Function

Public Function RoteirosDeFabricacao_Le_Versoes(ByVal objRoteirosDeFabricacao As ClassRoteirosDeFabricacao, ByVal colRotFab As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRoteirosDeFabricacao As typeRoteirosDeFabricacao
Dim objRotFab As ClassRoteirosDeFabricacao


On Error GoTo Erro_RoteirosDeFabricacao_Le_Versoes

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193797

    'Alocação de espaço no buffer
    tRoteirosDeFabricacao.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tRoteirosDeFabricacao.sVersao = String(STRING_KIT_VERSAO, 0)
    tRoteirosDeFabricacao.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tRoteirosDeFabricacao.sUM = String(STRING_UM_SIGLA, 0)
    tRoteirosDeFabricacao.sAutor = String(STRING_AUTOR_ROT_FAB, 0)

    'Le a tabelaRoteirosDeFabricacao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, ProdutoRaiz, Versao, Descricao, DataCriacao, DataUltModificacao, " & _
                "Quantidade, UM, Autor, NumMaxMaqPorOper FROM RoteirosDeFabricacao WHERE ProdutoRaiz= ?  ", _
                tRoteirosDeFabricacao.lNumIntDoc, tRoteirosDeFabricacao.sProdutoRaiz, tRoteirosDeFabricacao.sVersao, tRoteirosDeFabricacao.sDescricao, _
                tRoteirosDeFabricacao.dtDataCriacao, tRoteirosDeFabricacao.dtDataUltModificacao, tRoteirosDeFabricacao.dQuantidade, tRoteirosDeFabricacao.sUM, tRoteirosDeFabricacao.sAutor, tRoteirosDeFabricacao.iNumMaxMaqPorOper, _
                objRoteirosDeFabricacao.sProdutoRaiz)
    If lErro <> AD_SQL_SUCESSO Then gError 193798

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193799

    If lErro <> AD_SQL_SUCESSO Then gError 193800

    Do While lErro = AD_SQL_SUCESSO

        Set objRotFab = New ClassRoteirosDeFabricacao

        objRotFab.lNumIntDoc = tRoteirosDeFabricacao.lNumIntDoc
        objRotFab.sProdutoRaiz = tRoteirosDeFabricacao.sProdutoRaiz
        objRotFab.sVersao = tRoteirosDeFabricacao.sVersao
        objRotFab.sDescricao = tRoteirosDeFabricacao.sDescricao
        objRotFab.dtDataCriacao = tRoteirosDeFabricacao.dtDataCriacao
        objRotFab.dtDataUltModificacao = tRoteirosDeFabricacao.dtDataUltModificacao
        objRotFab.dQuantidade = tRoteirosDeFabricacao.dQuantidade
        objRotFab.sUM = tRoteirosDeFabricacao.sUM
        objRotFab.sAutor = tRoteirosDeFabricacao.sAutor
        objRotFab.iNumMaxMaqPorOper = tRoteirosDeFabricacao.iNumMaxMaqPorOper
        
        colRotFab.Add objRotFab

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193799
    
    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    RoteirosDeFabricacao_Le_Versoes = SUCESSO

    Exit Function

Erro_RoteirosDeFabricacao_Le_Versoes:

    RoteirosDeFabricacao_Le_Versoes = gErr

    Select Case gErr

        Case 193797
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 193798, 193799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTEIROSDEFABRICACAO", gErr)

        Case 193800

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193801)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Almoxarifado_Le_Faixa(ByVal viAlmIni As Variant, ByVal viAlmFim As Variant, ByVal colAlm As Collection) As Long
 
Dim lErro As Long
Dim iIndice As Integer
Dim sSQL As String
Dim sWhere As String
Dim viAlm As Variant
Dim alComando(1 To 1) As Long
Dim objAlm As ClassAlmoxarifado
 
On Error GoTo Erro_Almoxarifado_Le_Faixa
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 194840
    Next
    
    viAlm = CInt(0)
    
    sSQL = "SELECT Codigo FROM Almoxarifado "
 
    If viAlmIni <> 0 Then
        sWhere = sWhere & "WHERE Codigo >= ? "
    End If
    If viAlmFim <> 0 Then
        If Len(Trim(sWhere)) <> 0 Then
            sWhere = sWhere & "AND "
        Else
            sWhere = "WHERE "
        End If
        sWhere = sWhere & "Codigo <= ? "
    End If
    
    sSQL = sSQL & sWhere
    
    sSQL = sSQL & "ORDER BY Codigo "
 
    lErro = Comando_PrepararInt(alComando(1), sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 194841
 
    lErro = Comando_BindVarInt(alComando(1), viAlm)
    If (lErro <> AD_SQL_SUCESSO) Then gError 194842
 
    If viAlmIni <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), viAlmIni)
        If (lErro <> AD_SQL_SUCESSO) Then gError 194843
    End If
    
    If viAlmFim <> 0 Then
        lErro = Comando_BindVarInt(alComando(1), viAlmFim)
        If (lErro <> AD_SQL_SUCESSO) Then gError 194844
    End If
 
    lErro = Comando_ExecutarInt(alComando(1))
    If (lErro <> AD_SQL_SUCESSO) Then gError 194845
 
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194846
        
    'Para cada Produto
    Do While lErro = AD_SQL_SUCESSO
    
        Set objAlm = New ClassAlmoxarifado
 
        objAlm.iCodigo = viAlm
        
        'Lê o produto
        lErro = CF("Almoxarifado_Le", objAlm)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 194847
     
        colAlm.Add objAlm
     
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194848
 
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Almoxarifado_Le_Faixa = SUCESSO
 
    Exit Function
 
Erro_Almoxarifado_Le_Faixa:
 
    Almoxarifado_Le_Faixa = gErr
 
    Select Case gErr
 
        Case 194840
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 194841 To 194846, 194848
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case 194847
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194849)
 
    End Select
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
 
    Exit Function
 
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Certificados_Le(ByVal objCertificado As ClassCertificados) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCertificado As typeCertificados
Dim sSQL As String, vFiltro As Variant

On Error GoTo Erro_Certificados_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213310

    'Alocação de espaço no buffer
    tCertificado.sDescricao = String(STRING_MAXIMO, 0)
    tCertificado.sSigla = String(STRING_MAXIMO, 0)
    
    sSQL = "SELECT Codigo, Descricao, Sigla, Validade FROM Certificados WHERE "

    'Le a tabelaCertificados
    If objCertificado.lCodigo <> 0 Then
        sSQL = sSQL & "Codigo= ? "
        vFiltro = objCertificado.lCodigo
    Else
        sSQL = sSQL & "Sigla= ? "
        vFiltro = objCertificado.sSigla
    End If
    lErro = Comando_Executar(lComando, sSQL, tCertificado.lCodigo, tCertificado.sDescricao, tCertificado.sSigla, tCertificado.lValidade, vFiltro)
    If lErro <> AD_SQL_SUCESSO Then gError 213311

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213312

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCertificado.lCodigo = tCertificado.lCodigo
    objCertificado.sDescricao = tCertificado.sDescricao
    objCertificado.sSigla = tCertificado.sSigla
    objCertificado.lValidade = tCertificado.lValidade

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Certificados_Le = SUCESSO

    Exit Function

Erro_Certificados_Le:

    Certificados_Le = gErr

    Select Case gErr

        Case 213310
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213311, 213312
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CERTIFICADOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213313)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function Cursos_Le(ByVal objCurso As ClassCursos) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 3) As Long
Dim tCurso As typeCursos
Dim sSQL As String, vFiltro1 As Variant, vFiltro2 As Variant
Dim objMO As ClassCursoMO, tMO As typeCursoMO
Dim objCertificado As ClassCursoCertificados, tCertificado As typeCursoCertificados

On Error GoTo Erro_Cursos_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213343
    Next

    'Alocação de espaço no buffer
    tCurso.sDetalhamento = String(STRING_MAXIMO, 0)
    tCurso.sResponsavel = String(STRING_MAXIMO, 0)
    
    sSQL = "SELECT NumIntDoc, Codigo, FilialEmpresa, Detalhamento, Responsavel, DataInicio, DataConclusao FROM Cursos WHERE "

    If objCurso.lNumIntDoc <> 0 Then
        sSQL = sSQL & " NumIntDoc = ? AND FilialEmpresa <> ? "
        vFiltro1 = objCurso.lNumIntDoc
        vFiltro2 = CInt(0)
    Else
        sSQL = sSQL & " Codigo= ? AND FilialEmpresa= ? "
        vFiltro1 = objCurso.lCodigo
        vFiltro2 = objCurso.iFilialEmpresa
    End If

    'Le a tabelaCursos
    With tCurso
        lErro = Comando_Executar(alComando(1), sSQL, .lNumIntDoc, .lCodigo, .iFilialEmpresa, .sDetalhamento, .sResponsavel, .dtDataInicio, .dtDataConclusao, vFiltro1, vFiltro2)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 213344

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213345

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCurso.lNumIntDoc = tCurso.lNumIntDoc
    objCurso.lCodigo = tCurso.lCodigo
    objCurso.iFilialEmpresa = tCurso.iFilialEmpresa
    objCurso.sDetalhamento = tCurso.sDetalhamento
    objCurso.sResponsavel = tCurso.sResponsavel
    objCurso.dtDataInicio = tCurso.dtDataInicio
    objCurso.dtDataConclusao = tCurso.dtDataConclusao
    
    Set objCurso.colCertificados = New Collection
    Set objCurso.colMOCursos = New Collection
    
    lErro = Comando_Executar(alComando(2), "SELECT CodCertificado FROM CursoCertificados WHERE NumIntDocCurso = ?", tCertificado.lCodCertificado, objCurso.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 213346

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213347
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCertificado = New ClassCursoCertificados
        objCertificado.lCodCertificado = tCertificado.lCodCertificado
        objCertificado.lNumIntDocCurso = objCurso.lNumIntDoc
        objCurso.colCertificados.Add objCertificado
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213348
    
    Loop
    
    tMO.sAvaliacao = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(3), "SELECT CodMO, Aprovado, Avaliacao FROM CursoMO WHERE NumIntDocCurso = ? ORDER BY Seq", tMO.iCodMO, tMO.iAprovado, tMO.sAvaliacao, objCurso.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 213349

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213350
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objMO = New ClassCursoMO
        objMO.iCodMO = tMO.iCodMO
        objMO.iAprovado = tMO.iAprovado
        objMO.sAvaliacao = tMO.sAvaliacao
        objMO.lNumIntDocCurso = objCurso.lNumIntDoc
        
        objCurso.colMOCursos.Add objMO
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213351
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cursos_Le = SUCESSO

    Exit Function

Erro_Cursos_Le:

    Cursos_Le = gErr

    Select Case gErr

        Case 213343
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213344 To 213351
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CURSOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213352)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function Certificados_Lista_Validade_Texto(ByVal iNumDias As Integer, iNumCertif As Integer, sLista As String) As Long

Dim lErro As Long
Dim lComando As Long, iIndice As Integer
Dim dtDataLimite As Date, sListaLinha As String, sListaAux As String
Dim iCodMO As Integer, sDescMO As String, lCodCertificado As Long
Dim sSigla As String, dtDataValidade As Date

On Error GoTo Erro_Certificados_Lista_Validade_Texto

    sLista = ""
    sListaAux = ""
    iNumCertif = 0
    
    dtDataLimite = DateAdd("d", iNumDias, Date)

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213357

    'Alocação de espaço no buffer
    sDescMO = String(STRING_MAXIMO, 0)
    sSigla = String(STRING_MAXIMO, 0)

    lErro = Comando_Executar(lComando, "SELECT CodMO, DescMO, CodCertificado, Sigla, DataValidade FROM CertificadoMO WHERE DataValidade <= ? ORDER BY DataValidade", iCodMO, sDescMO, lCodCertificado, sSigla, dtDataValidade, dtDataLimite)
    If lErro <> AD_SQL_SUCESSO Then gError 213358

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213359

    iIndice = 0
    Do While lErro <> AD_SQL_SEM_DADOS
    
        iIndice = iIndice + 1

        If iIndice > 1 Then sListaAux = sListaAux & vbNewLine
        
        sListaLinha = CStr(lCodCertificado) & SEPARADOR & sSigla & " em " & Format(dtDataValidade, "dd/mm/yyyy") & " p/ " & left(CStr(iCodMO) & SEPARADOR & sDescMO, 15)
        
        sListaAux = sListaAux & sListaLinha

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213360

    Loop
    
    If Len(sListaAux) > 150 Then sListaAux = left(sListaAux, 150) & "..."
    
    sLista = "Existem " & CStr(iIndice) & " certificados a vencer em " & CStr(iNumDias) & "dias." & vbNewLine & sListaAux

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Certificados_Lista_Validade_Texto = SUCESSO

    Exit Function

Erro_Certificados_Lista_Validade_Texto:

    Certificados_Lista_Validade_Texto = gErr

    Select Case gErr

        Case 213357
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213358 To 213360
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CERTIFICADOS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213361)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
