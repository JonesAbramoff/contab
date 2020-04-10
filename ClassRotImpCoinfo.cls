VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotImpCoinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? na importacao só gravar os registros do estabelecimento devido (TVA, TVI,...)
'??? apagar registros qdo arquivos texto já tiverem sido importados antes

'info arq texto:
'todos divididos em 3 seções: dados, parametros e tamanho de dados

'pendencias jones:
    '
    
'pendencias mario:
    'cartao de credito vs retaguarda

'duvidas:

'onde pego o "nosso numero" da parcela da fatura ? cobrebem, num do docto = num da fatura

'guardar copias dos dados importados em tabelas do sgedados ? evitaria modificar bd do corporator e facilitaria link com coinfo
    'EXPORTACAO

'conferir chaves, relacionamento, flags,...
    'como liga o item da fatura (ou a fatura) com as ocorrencias
    'as ocorrencias poderao estar "soltas" ? sim, qdo recnsica = 0
    
'-existe cadastro de tipos de empresas, se nao for cliente é fornecedor ? vou cadastrar como os dois

'-se eu conseguir gravar no bd intermediario, posso marcar os registros como já lidos (ou exclui-los)
'ou as tabelas terem um identificador de atualizacao

'valor em moeda vs R$: travel só usa R$

'o numero da fatura é independente da filial ? por empresa (travel ace e asistbras)

'há "interacao" entre filiais, tipo uma fatura de uma sendo cobrada na nf da outra, conta corrente, etc. ?

'codest = cod de empresa

'a numeracao de nf nao se repete ? nao

'qdo a fatura entra "paga" ?

'será que devo cadastrar os clientes tb como fornecedor ? sim

'incluir nas tabelas de filiaisclientes e filiaisfornecedores o codemp da coinfo (como CodExterno)

'pré-cadastrar categorias de cliente ou criar novos campos ou usar copia dos dados importados:
    '- tipo de cobranca (carteira/bancaria)
    '- envfat, envnota, envocor (email ou impre)
    '

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCoInfoLogAtualizacao
    lNumIntArq As Long
    iTabela As Integer
    lNumRegArqTexto As Long
    iSeq As Integer
    sCritica As String
End Type

Dim glConexao1 As Long
Dim glConexao2 As Long
Dim glConexao3 As Long

Const TABELA_COINFOFATURA = 1
Const TABELA_COINFONF = 2
Const TABELA_COINFOITEMFATURA = 3

'Const TAMANHO_SEGMENTO_PRODUTO = 10

'Const FATOR_SOMA_COD_EMISSOR = 500000 'Código do fornecedor(emissor no corporator = código no sigav + FATOR_SOMA_COD_EMISSOR)

'Informações sobre os arquivos importados, para adicinar uma nova importação
'alterar NUM_ARQS, incluir NOME_ARQ e POS_ARQ e alterar as rotinas
'Coinfo_ImportarDados_Trata_Diretorio e Coinfo_ImportarDados_Trata_Diretorio2
Const NOME_ARQ_TIPOEMPR = "tipoempr.d"
Const NOME_ARQ_EMPRESA = "empresa.d"
Const NOME_ARQ_PRODUTO = "produtos.d" 'Agora está certo 'Está errado mesmo para não importar (Não é necessário e dá erro por causa de trocas de código de produto mantendo o nome reduzido)
Const NOME_ARQ_FATURA = "cabecfat.d"
Const NOME_ARQ_ITEMFATURA = "itemfatu.d"
Const NOME_ARQ_OCORRENCIA = "ocorrencia.d"
Const NOME_ARQ_NF = "nf.d"
Const NOME_ARQ_ITEMNF = "itemnf.d"
Const NOME_ARQ_GRUPO = "grupo.d"
Const NOME_ARQ_OVER = "overemis.d"
Const NOME_ARQ_EMISSOR = "emissor.d"
Const NOME_ARQ_COMISSAO = "comissao.d"
Const NOME_ARQ_ACORDO = "acordos.d"
Const NOME_ARQ_ACORDOCOMIS = "acordoscomis.d"
Const NOME_ARQ_DESTINO = "destinos.d"

Const POS_ARQ_TIPOEMPR = 1
Const POS_ARQ_EMPRESA = 2
Const POS_ARQ_PRODUTO = 3
Const POS_ARQ_FATURA = 4
Const POS_ARQ_ITEMFATURA = 5
Const POS_ARQ_OCORRENCIA = 6
Const POS_ARQ_NF = 7
Const POS_ARQ_ITEMNF = 8
Const POS_ARQ_GRUPO = 9
Const POS_ARQ_OVER = 10
Const POS_ARQ_EMISSOR = 11
Const POS_ARQ_COMISSAO = 12
Const POS_ARQ_ACORDO = 13
Const POS_ARQ_ACORDOCOMIS = 14
Const POS_ARQ_DESTINO = 15

Const NUM_ARQS = 15 'Número do arquivo

Private Function StrParaDblCoinfo(ByVal sNumeroComoTexto As String) As Double

    If sNumeroComoTexto = "?" Or sNumeroComoTexto = "NULL" Then
        StrParaDblCoinfo = 0
    Else
        StrParaDblCoinfo = StrParaDbl(sNumeroComoTexto)
    End If

End Function

Private Function StrParaDateCoinfo(ByVal sDataComoTexto As String) As Date

    If sDataComoTexto = "?" Or sDataComoTexto = "NULL" Then
        StrParaDateCoinfo = DATA_NULA
    Else
        StrParaDateCoinfo = StrParaDate(sDataComoTexto)
    End If
    
End Function

Private Function Coinfo_ImportarArqTexto(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.

Dim lErro As Long
Dim objArquivo As File
Dim sPosicao As String

On Error GoTo Erro_Coinfo_ImportarArqTexto

    'importar tipos de Grupo
    If objImpCoinfo.sArqGrupoHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto0"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqGrupoHora
    
        lErro = Coinfo_ImportarArqTexto0(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130593 And lErro <> 53 Then gError 130572
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
        
    End If

    'importar tipos de empresa
    If objImpCoinfo.sArqTipoEmprHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto1"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqTipoEmprHora
    
        lErro = Coinfo_ImportarArqTexto1(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130593 And lErro <> 53 Then gError 130572
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
        
    End If
    
    'importar empresas
    If objImpCoinfo.sArqEmprHora <> "" Then

        sPosicao = "Coinfo_ImportarArqTexto2"

        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqEmprHora

        lErro = Coinfo_ImportarArqTexto2(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130604 And lErro <> 53 Then gError 130608
    
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
    
    End If
    
    'importar produtos
    If objImpCoinfo.sArqProdHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto3"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqProdHora
        
        lErro = Coinfo_ImportarArqTexto3(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130619 And lErro <> 53 Then gError 130623
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
    
    End If
    
    'importar faturas (cabecalho e itens)
    If objImpCoinfo.sArqFatHora <> "" Or objImpCoinfo.sArqItemFatHora <> "" Then
        
        sPosicao = "Coinfo_ImportarArqTexto4"
        
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqFatHora & " e " & objImpCoinfo.sArqItemFatHora
        
        lErro = Coinfo_ImportarArqTexto4(objImpCoinfo)
        If lErro <> SUCESSO Then gError 130643
    
        If objImpCoinfo.sArqFatHora <> "" Then
            lErro = Rotina_Importar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 182709
        End If
        If objImpCoinfo.sArqItemFatHora <> "" Then
            lErro = Rotina_Importar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 182709
        End If
    
    End If
    
    'importar ocorrencias
    If objImpCoinfo.sArqOcorHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto5"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqOcorHora
        
        lErro = Coinfo_ImportarArqTexto5(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130673 And lErro <> 53 Then gError 130677
    
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
    
    End If
    
    'importar nfs e seus itens
    If objImpCoinfo.sArqNFHora <> "" Or objImpCoinfo.sArqItemNFHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto6"
        
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqNFHora & " e " & objImpCoinfo.sArqItemNFHora
        
        lErro = Coinfo_ImportarArqTexto6(objImpCoinfo)
        If lErro <> SUCESSO Then gError 130677
    
        If objImpCoinfo.sArqNFHora <> "" Then
            lErro = Rotina_Importar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 182709
        End If
        If objImpCoinfo.sArqItemNFHora <> "" Then
            lErro = Rotina_Importar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 182709
        End If
        
    End If
    
    'importar Emissores
    If objImpCoinfo.sArqEmissHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto7"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqEmissHora
    
        lErro = Coinfo_ImportarArqTexto7(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
        
    End If
    
    'importar Comissões Over
    If objImpCoinfo.sArqComisOverHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto8"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqComisOverHora
    
        lErro = Coinfo_ImportarArqTexto8(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
        
    End If
    
    '&&&&&&&############$$$$$$$$$$$%%%%%%%%%%%++++++++?????????
    'DESCOMENTAR AQUI QUANDO LIBERADO A PARTE DE COMISSÃO SOBRE CARTÃO
    
    'importar Comissões
    If objImpCoinfo.sArqComissaoHora <> "" Then

        sPosicao = "Coinfo_ImportarArqTexto9"

        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqComissaoHora

        lErro = Coinfo_ImportarArqTexto9(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572

        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709

    End If
    
    'DESCOMENTAR AQUI QUANDO LIBERADO A PARTE DE COMISSÃO SOBRE CARTÃO
    '&&&&&&&############$$$$$$$$$$$%%%%%%%%%%%++++++++?????????
    
    'importar Acordos
    If objImpCoinfo.sArqAcordoHora <> "" And objImpCoinfo.sArqAcordoComisHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto10"
    
        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqAcordoHora
        
        If right(objImpCoinfo.sArqAcordoHora, 14) <> right(objImpCoinfo.sArqAcordoComisHora, 14) Then gError 130572
       
        lErro = Coinfo_ImportarArqTexto10(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
        
    End If
    
    'importar Destinos
    If objImpCoinfo.sArqDestinoHora <> "" Then
    
        sPosicao = "Coinfo_ImportarArqTexto11"

        TelaAcompanhaBatchFAT.NomeArqImp.Caption = objImpCoinfo.sArqDestinoHora
       
        lErro = Coinfo_ImportarArqTexto11(objImpCoinfo)
        If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572
        
        lErro = Rotina_Importar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182709
        
    End If
    
    TelaAcompanhaBatchFAT.NomeArqImp.Caption = ""
    
    Coinfo_ImportarArqTexto = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto:

    Coinfo_ImportarArqTexto = gErr
     
    Select Case gErr
    
        Case 130572
            Call Rotina_Erro(vbOKOnly, "Os arquivos de acordos não estão integros. Acordo: " & right(objImpCoinfo.sArqAcordoHora, 14) & " e Comissões dos Acordos: " & right(objImpCoinfo.sArqAcordoComisHora, 14), gErr, Error)
          
        Case 130608, 130623, 130643, 130677, 182709
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130561)
     
    End Select
    
    Call Rotina_Erro(vbOKOnly, "A rotina de importação Coinfo_ImportarArqTexto apresentou um erro próximo a posição " & sPosicao, gErr, Error)
     
    Exit Function

End Function

Function Copia_Arquivo(ByVal sDir As String, ByVal sDirDestino As String) As Long

On Error GoTo Erro_Copia_Arquivo

    FileCopy sDir, sDirDestino

    Exit Function
    
Erro_Copia_Arquivo:

    Copia_Arquivo = gErr
     
    Select Case gErr
          
        Case 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182654)
     
    End Select
     
    Exit Function

End Function

Function Cria_Diretorio(ByVal sDir As String) As Long

On Error GoTo Erro_Cria_Diretorio

    MkDir sDir

    Exit Function
    
Erro_Cria_Diretorio:

    Cria_Diretorio = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182655)
     
    End Select
     
    Exit Function

End Function

Function Coinfo_ImportarDados_Trata_Diretorio(ByVal objImpCoinfo As ClassImpCoinfo, ByVal colcolFile As Collection, iContador As Integer) As Long
'Monta uma coleção com todos os arquivos que tem que ser importados

Dim lErro As Long
Dim sDiretorio As String
Dim sDiretorioNovo As String
Dim colFile As Collection
Dim objFSO As New FileSystemObject
Dim objFolder As Folder
Dim objFile As File
Dim iPOS As Integer
Dim sFileName As String
Dim iIndice As Integer
Dim bCopiar As Boolean
Dim iContadorAux As Integer
Dim lRetorno As Long
Dim colSaida As Collection
Dim colCampos As Collection
Dim iIndiceAux As Integer
Dim iTotalArquivo As Integer

On Error GoTo Erro_Coinfo_ImportarDados_Trata_Diretorio

    iTotalArquivo = 0

    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    If iSistemaIntegrado = SISTEMA_INTEGRADO_SIGAV Then
        lRetorno = GetPrivateProfileString("Geral", "dirArqImpCoinfo ", "c:\travelace\dadosfp\", sDiretorio, 512, "ADM100.INI")
    Else
        lRetorno = GetPrivateProfileString("Geral", "dirArqImpKogut ", "c:\travelace\dadosfp\", sDiretorio, 512, "ADM100.INI")
    End If
    sDiretorio = left(sDiretorio, lRetorno)
    
    sDiretorioNovo = sDiretorio & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "hhmmss")
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Lendo arquivos da pasta " & sDiretorio & vbNewLine
    
    'Cria um diretório novo
    Call Cria_Diretorio(sDiretorioNovo)
    
    sDiretorioNovo = sDiretorioNovo & "\"
    
    objImpCoinfo.sDiretorio = sDiretorio
    objImpCoinfo.sDiretorioNovo = sDiretorioNovo
    
    'Pega todos os aquivos da pasta
    Set objFolder = objFSO.GetFolder(sDiretorio)
    
    'Cria uma coleção para cada tipo de arquivo
    For iIndice = 1 To NUM_ARQS
        Set colFile = New Collection
        colcolFile.Add colFile
    Next
    
    'Para cada arquivo
    For Each objFile In objFolder.Files
        
        'Pega o Nome sem a data
        sFileName = objFile.Name
        For iIndice = 1 To Len(objFile.Name)
            If IsNumeric(Mid(objFile.Name, iIndice, 1)) Then
                sFileName = left(objFile.Name, iIndice - 1) & ".d"
                Exit For
            End If
        
        Next
        
        bCopiar = True
        iPOS = 0
    
        'Verifica qual é o tipo de arquivo
        Select Case sFileName
        
            Case NOME_ARQ_TIPOEMPR
                iPOS = POS_ARQ_TIPOEMPR
            
            Case NOME_ARQ_EMPRESA
                iPOS = POS_ARQ_EMPRESA
            
            Case NOME_ARQ_PRODUTO
                iPOS = POS_ARQ_PRODUTO
            
            Case NOME_ARQ_FATURA
                iPOS = POS_ARQ_FATURA
            
            Case NOME_ARQ_ITEMFATURA
                iPOS = POS_ARQ_ITEMFATURA
            
            Case NOME_ARQ_OCORRENCIA
                iPOS = POS_ARQ_OCORRENCIA
            
            Case NOME_ARQ_NF
                iPOS = POS_ARQ_NF
            
            Case NOME_ARQ_ITEMNF
                iPOS = POS_ARQ_ITEMNF
                
            Case NOME_ARQ_GRUPO
                iPOS = POS_ARQ_GRUPO
        
            Case NOME_ARQ_OVER
                iPOS = POS_ARQ_OVER
        
            Case NOME_ARQ_EMISSOR
                iPOS = POS_ARQ_EMISSOR
        
            Case NOME_ARQ_COMISSAO
                iPOS = POS_ARQ_COMISSAO
        
            Case NOME_ARQ_ACORDO
                iPOS = POS_ARQ_ACORDO
        
            Case NOME_ARQ_ACORDOCOMIS
                iPOS = POS_ARQ_ACORDOCOMIS
        
            Case NOME_ARQ_DESTINO
                iPOS = POS_ARQ_DESTINO
        
            Case Else
                bCopiar = False
            
        End Select
    
        'Se é um dos que vão ser importados, copia e coloca na coleção
        If bCopiar Then
            Call Copia_Arquivo(sDiretorio & objFile.Name, sDiretorioNovo & objFile.Name)
        
            Set colFile = colcolFile.Item(iPOS)
            
            TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Arquivo " & objFile.Name & " encontrado." & vbNewLine
            
            iTotalArquivo = iTotalArquivo + 1
            
            colFile.Add objFile
        
        End If
    
    Next
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & CStr(iTotalArquivo) & " arquivos encontrados." & vbNewLine
    TelaAcompanhaBatchFAT.dValorTotalImp = iTotalArquivo
    
    Set colCampos = New Collection
    
    colCampos.Add "Name"
    
    'Obtém a quantidade de vezes que vai ter que ocorrer a importação
    iContadorAux = 0
    For iIndice = 1 To NUM_ARQS
        If colcolFile.Item(iIndice).Count > iContadorAux Then iContadorAux = colcolFile.Item(iIndice).Count
    
        If colcolFile.Item(iIndice).Count > 1 Then
        
            Set colSaida = New Collection
            
            'Ordena a coleção para importar os arquivos mais antigos primeiro
            Call Ordena_Colecao(colcolFile.Item(iIndice), colSaida, colCampos)
            
            For iIndiceAux = colcolFile.Item(iIndice).Count To 1 Step -1
                colcolFile.Item(iIndice).Remove iIndiceAux
            Next
            
            For Each objFile In colSaida
                colcolFile.Item(iIndice).Add objFile
            Next
        
        End If
    
    Next
    iContador = iContadorAux
    
    Coinfo_ImportarDados_Trata_Diretorio = SUCESSO

    Exit Function
    
Erro_Coinfo_ImportarDados_Trata_Diretorio:

    Coinfo_ImportarDados_Trata_Diretorio = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182655)
     
    End Select
     
    Exit Function
    
End Function

Function Coinfo_ImportarDados_Trata_Diretorio2(ByVal objImpCoinfo As ClassImpCoinfo, ByVal colcolFile As Collection, ByVal iContador As Integer, ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim sArquivo As Variant
Dim objFile As File

On Error GoTo Erro_Coinfo_ImportarDados_Trata_Diretorio2

    'Para cada arquivo
    For iTipo = 1 To NUM_ARQS
    
        'Se ele existe para importação
        If iIndice <= colcolFile.Item(iTipo).Count Then
            Set objFile = colcolFile.Item(iTipo).Item(iIndice)
            sArquivo = objFile.Name
        Else
            sArquivo = ""
        End If
        
        'De acordo com o tipo obtém o nome
        Select Case iTipo
        
            Case POS_ARQ_TIPOEMPR
                objImpCoinfo.sArqTipoEmpr = NOME_ARQ_TIPOEMPR
                objImpCoinfo.sArqTipoEmprHora = sArquivo
        
            Case POS_ARQ_EMPRESA
                objImpCoinfo.sArqEmpr = NOME_ARQ_EMPRESA
                objImpCoinfo.sArqEmprHora = sArquivo
        
            Case POS_ARQ_PRODUTO
                objImpCoinfo.sArqProd = NOME_ARQ_PRODUTO
                objImpCoinfo.sArqProdHora = sArquivo
        
            Case POS_ARQ_FATURA
                objImpCoinfo.sArqFat = NOME_ARQ_FATURA
                objImpCoinfo.sArqFatHora = sArquivo
        
            Case POS_ARQ_ITEMFATURA
                objImpCoinfo.sArqItemFat = NOME_ARQ_ITEMFATURA
                objImpCoinfo.sArqItemFatHora = sArquivo
        
            Case POS_ARQ_OCORRENCIA
                objImpCoinfo.sArqOcor = NOME_ARQ_OCORRENCIA
                objImpCoinfo.sArqOcorHora = sArquivo
        
            Case POS_ARQ_NF
                objImpCoinfo.sArqNF = NOME_ARQ_NF
                objImpCoinfo.sArqNFHora = sArquivo
        
            Case POS_ARQ_ITEMNF
                objImpCoinfo.sArqItemNF = NOME_ARQ_ITEMNF
                objImpCoinfo.sArqItemNFHora = sArquivo
        
            Case POS_ARQ_GRUPO
                objImpCoinfo.sArqGrupo = NOME_ARQ_GRUPO
                objImpCoinfo.sArqGrupoHora = sArquivo
        
            Case POS_ARQ_OVER
                objImpCoinfo.sArqComisOver = NOME_ARQ_OVER
                objImpCoinfo.sArqComisOverHora = sArquivo
        
            Case POS_ARQ_EMISSOR
                objImpCoinfo.sArqEmiss = NOME_ARQ_EMISSOR
                objImpCoinfo.sArqEmissHora = sArquivo
        
            Case POS_ARQ_COMISSAO
                objImpCoinfo.sArqComissao = NOME_ARQ_COMISSAO
                objImpCoinfo.sArqComissaoHora = sArquivo
        
            Case POS_ARQ_ACORDO
                objImpCoinfo.sArqAcordo = NOME_ARQ_ACORDO
                objImpCoinfo.sArqAcordoHora = sArquivo
        
            Case POS_ARQ_ACORDOCOMIS
                objImpCoinfo.sArqAcordoComis = NOME_ARQ_ACORDOCOMIS
                objImpCoinfo.sArqAcordoComisHora = sArquivo
        
            Case POS_ARQ_DESTINO
                objImpCoinfo.sArqDestino = NOME_ARQ_DESTINO
                objImpCoinfo.sArqDestinoHora = sArquivo
        
        End Select
        
    Next
        
    Coinfo_ImportarDados_Trata_Diretorio2 = SUCESSO

    Exit Function
    
Erro_Coinfo_ImportarDados_Trata_Diretorio2:

    Coinfo_ImportarDados_Trata_Diretorio2 = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182655)
     
    End Select
     
    Exit Function
    
End Function

'Function Coinfo_ImportarDados(ByVal objFiltro As Object) As Long
''Função inicial de importação dos dados da Coinfo para o Corporator
'
'Dim lErro As Long
'Dim objImpCoinfo As New ClassImpCoinfo
'Dim colcolFile As New Collection
'Dim iIndice As Integer
'Dim iContador As Integer
'Dim vbResult As VbMsgBoxResult
'Dim sResult As String
'Dim iAno As Integer
'Dim iMes As Integer
'
'On Error GoTo Erro_Coinfo_ImportarDados
'
'    Set objImpCoinfo.objFiltro = objFiltro
'
'    'Obtém a sigla da empresa
'    'TVI - Travel Ace
'    'TVA - AssistBras
'    If glEmpresa = 2 Then
'        objImpCoinfo.sSiglaEmpresa = "TVI"
'    ElseIf glEmpresa = 1 Then
'        objImpCoinfo.sSiglaEmpresa = "TVA"
'    Else
'        Coinfo_ImportarDados = SUCESSO
'        Exit Function
'    End If
'
'    'Só importa da TVA
'    If glEmpresa <> 1 Then
'        Coinfo_ImportarDados = SUCESSO
'        Exit Function
'    End If
'
'    'Cria conexões com o SGEDados glConexao1 e glConexao2
'    lErro = Abre_Conexoes
'    If lErro <> SUCESSO Then gError 182724
'
'    If objFiltro.iSoArqExport = DESMARCADO Then
'
'        If iForcaSistemaIntegrado = DESMARCADO Then
'
'            If objFiltro.iImportarNVLRetroativo = MARCADO Then
'                'vbResult = MsgBox("Rotina Interna. Associa produtos aos almoxarifados. Deseja continuar?", vbYesNo, "Corporator")
'                vbResult = MsgBox("Rotina Interna. Regerar exportação para o sisweb. Deseja continuar?", vbYesNo, "Corporator")
'                If vbResult = vbYes Then
'                    Call Regera_Log_Vou
'                    'Call Coinfo_Produto_Associa_Alm
'                End If
'            End If
'
'            If objFiltro.iAtualizarClientesRetroativo = MARCADO Then
'                vbResult = MsgBox("Monta a hierarquia de vendedores dos voucher não pagos ?", vbYesNo, "Corporator")
'                If vbResult = vbYes Then Call Regera_Hierarquia_Vou
'            End If
'
'            If objFiltro.iGerarComissaoRetroativo = MARCADO Then
'                vbResult = MsgBox("Deseja gerar o comissionamento a partir do Corporator  ?", vbYesNo, "Corporator")
'                If vbResult = vbYes Then Call Coinfo_ImportaComissao_150908(objImpCoinfo)
'
''                sResult = InputBox("Deseja gerar o comissionamento a partir de quando pelo Corporator  ?", "Corporator", "01/2008")
''                If Len(Trim(sResult)) <> 0 Then
''                    For iAno = Right(sResult, 4) To Year(Date)
''                        For iMes = 1 To 12
''                            If iAno = Year(Date) And iMes > Month(Date) Then Exit For
''                            If iAno <> Right(sResult, 4) Or iMes >= Left(sResult, 2) Then
''                                'vbResult = MsgBox("Deseja importar o comissionamento gerado pelo Sigav ?", vbYesNo, "Corporator")
''                                'If vbResult = vbYes Then
''                                Call Coinfo_ImportaComissao_150908(objImpCoinfo, iAno, iMes)
''                            End If
''                        Next
''                    Next
''                End If
'
'
'            End If
'
'        End If
'
'        If objFiltro.iImportArqsNovos = MARCADO Then
'
'            'Gera a coleção de arquivos a serem atualizados
'            lErro = Coinfo_ImportarDados_Trata_Diretorio(objImpCoinfo, colcolFile, iContador)
'            If lErro <> SUCESSO Then gError 182686
'
'            'Para cada conjunto de importação
'            For iIndice = 1 To iContador
'
'                'Coloca as informações do arquivo em questão em objImpCoinfo
'                lErro = Coinfo_ImportarDados_Trata_Diretorio2(objImpCoinfo, colcolFile, iContador, iIndice)
'                If lErro <> SUCESSO Then gError 182687
'
'                'importar dados de arquivos texto para tabelas de importacao
'                lErro = Coinfo_ImportarArqTexto(objImpCoinfo)
'                If lErro <> SUCESSO Then gError 130573
'
'            Next
'
'        End If
'
'        '&&&&&&&############$$$$$$$$$$$%%%%%%%%%%%++++++++?????????
'        'EXECUTAR APENAS UMA VEZ
'        'GERA A COMISSÃO SOBRE TÍTULOS A RECEBER JÁ IMPORTADOS
'        'Call Coinfo_GerarComissaoRetroativo(objImpCoinfo)
'        'EXECUTAR APENAS UMA VEZ
'        '&&&&&&&############$$$$$$$$$$$%%%%%%%%%%%++++++++?????????
'
'        '&&&&&&&############$$$$$$$$$$$%%%%%%%%%%%++++++++?????????
'        'EXECUTAR APENAS UMA VEZ
'        'GERA A ASSOCIAÇÃO ENTRE CLIENTES E ALMOXARIFADOS
'    '    lErro = Coinfo_Acerta_Cliente_Alm_Retroativo
'    '    If lErro <> SUCESSO Then
'    '        gError 130574
'    '    End If
'        'EXECUTAR APENAS UMA VEZ
'        '&&&&&&&############$$$$$$$$$$$%%%%%%%%%%%++++++++?????????
'
'        If iForcaSistemaIntegrado = DESMARCADO Then
'
'            'atualiza tabela de tipos de empresas (Tipos de Cliente e Tipos de Fornecedores)
'            lErro = Coinfo_AtualizarTiposEmp(objImpCoinfo)
'            If lErro <> SUCESSO Then gError 130574
'
'            'atualiza tabela de empresas (Clientes e Fornecedores)
'            lErro = Coinfo_AtualizarEmpresas(objImpCoinfo)
'            If lErro <> SUCESSO Then gError 130575
'
'            'atualizar tabela de emissores (Fornecedores)
'            lErro = Coinfo_AtualizarDestinos(objImpCoinfo)
'            If lErro <> SUCESSO Then gError 130578
'
'            'atualizar tabela de produtos
'            lErro = Coinfo_AtualizarProdutos(objImpCoinfo)
'            If lErro <> SUCESSO Then gError 130576
'
'            'atualizar tabela de emissores (Fornecedores)
'            lErro = Coinfo_AtualizarAcordos(objImpCoinfo)
'            If lErro <> SUCESSO Then gError 130578
'
'            'atualizar tabela de faturas (Comissões (NFsPag), Reembolsos (TitPag), Títulos a Receber (Cartão e Normal)
'            'e Comissões sobre os títulos a receber)
'        '    lErro = Coinfo_AtualizarFaturas(objImpCoinfo)
'        '    If lErro <> SUCESSO Then gError 130578
'
'            If objFiltro.iContabilizarFaturas = MARCADO Then
'                lErro = Coinfo_AtualizaContabFaturas
'                'If lErro <> SUCESSO Then gError 130577
'            End If
'
'            If objFiltro.iContabilizarNFs = MARCADO Then
'                lErro = Coinfo_AtualizaContabNFs
'                'If lErro <> SUCESSO Then gError 130578
'            End If
'
'            'atualizar tabela de notas fiscais
'        '    lErro = Coinfo_AtualizarNFs(objImpCoinfo)
'        '    If lErro <> SUCESSO Then gError 130577
'
'            'Gera os títulos a receber das administradoras de cartão de crédito
'            'segundo o modelo antigo do Sigav utilizado até 31/10/2006
'        '    lErro = Coinfo_AtualizarItensFatu(objImpCoinfo)
'        '    If lErro <> SUCESSO Then gError 130578
'
'            'atualizar tabela de emissores (Fornecedores)
'            lErro = Coinfo_AtualizarEmissores(objImpCoinfo)
'            If lErro <> SUCESSO Then gError 130578
'
''            If objFiltro.iGerarOver = MARCADO Then
''                'atualizar tabela de comissões over (Títulos Pagar)
''                lErro = Coinfo_AtualizarOver(objImpCoinfo)
''                'If lErro <> SUCESSO Then gError 130578
''            End If
'
'            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'            'Função que gera a comissão de cartão de crédito com base em informações do Sigav
'            'Existe um filtro dentro dela para não gerar nada apóps a data de implantação do
'            'comissionamento pelo Corporator
'            'atualizar tabela de comissao (NFs a Pagar)
'            lErro = Coinfo_AtualizarComissao(objImpCoinfo)
'            'If lErro <> SUCESSO Then gError 130578
'
'            If objFiltro.iGerarMovEst = MARCADO Then
'                'Atualiza os movimentos de estoque
'                lErro = Coinfo_AtualizarVouchers(objImpCoinfo)
'                'If lErro <> SUCESSO Then gError 189312
'            End If
'
'            If objFiltro.iContabilizarVouchers = MARCADO Then
'                'Atualiza a contabilização dos vouchers
'                lErro = Coinfo_AtualizarVouchers2(objImpCoinfo)
'                'If lErro <> SUCESSO Then gError 189312
'            End If
'
'            If Hour(Now) < 8 Or Hour(Now) >= 22 Then
'                Call RelLcto_Exclui
'            End If
'
'        End If
'
'        If iSistemaIntegrado = SISTEMA_INTEGRADO_SIGAV Then
'            Call Extrair_Vouchers_Sigav
'        ElseIf iSistemaIntegrado = SISTEMA_INTEGRADO_KOGUT Then
'            Call Simula_Extracao_Vouchers_Sigav
'        End If
'
''        If iForcaSistemaIntegrado = DESMARCADO Then
''
''            'Atualiza os títulos a receber a partir dos vouchers
''            'lErro = Coinfo_AtualizarVouchers3(objImpCoinfo)
''            'If lErro <> SUCESSO Then gError 189312
''
''            lErro = Coinfo_AtualizarOcr(objImpCoinfo)
''            'If lErro <> SUCESSO Then gError 189312
''
''        End If
'
'    End If
'
'    Call Importa_Ocrs(objImpCoinfo)
'
'    Call Importa_OcrsCasos(objImpCoinfo)
'
'    'Gera os arquivos de baixa e cancelamento de Baixa
'    lErro = Coinfo_Gera_Arquivos(objImpCoinfo)
'    If lErro <> SUCESSO Then gError 130579
'
'    'Fecha glConexao1 e glConexao2
'    Call Fecha_Conexoes
'
'    'Call Rotina_Aviso(vbOKOnly, "AVISO_IMPORTACAO_FEITA_COM_SUCESSO")
'
'    Coinfo_ImportarDados = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_ImportarDados:
'
'    Coinfo_ImportarDados = gErr
'
'    Select Case gErr
'
'        Case 130573 To 130579, 182686, 182687, 182724, 189312
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130555)
'
'    End Select
'
'    Call Fecha_Conexoes
'
'    Exit Function
'
'End Function


Function Coinfo_ImportarDados(ByVal objFiltro As Object) As Long
'Função inicial de importação dos dados da Coinfo para o Corporator

Dim lErro As Long
Dim objImpCoinfo As New ClassImpCoinfo
Dim colcolFile As New Collection
Dim iIndice As Integer
Dim iContador As Integer
Dim vbResult As VbMsgBoxResult
Dim sResult As String
Dim iAno As Integer
Dim iMes As Integer
Dim sPosicao As String

On Error GoTo Erro_Coinfo_ImportarDados

    Set objImpCoinfo.objFiltro = objFiltro

    'Obtém a sigla da empresa
    'TVI - Travel Ace
    'TVA - AssistBras
    If glEmpresa = 2 Then
        objImpCoinfo.sSiglaEmpresa = "TVI"
    ElseIf glEmpresa = 1 Then
        objImpCoinfo.sSiglaEmpresa = "TVA"
    End If
    
    'Só importa da TVA
    If glEmpresa <> 1 Then
        Coinfo_ImportarDados = SUCESSO
        Exit Function
    End If
    
    sPosicao = "Abre_Conexoes"
    
    'Cria conexões com o SGEDados glConexao1 e glConexao2
    lErro = Abre_Conexoes
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If objFiltro.iSoArqExport = DESMARCADO Then

        If iForcaSistemaIntegrado = DESMARCADO Then

            If objFiltro.iImportarNVLRetroativo = MARCADO Then
                vbResult = MsgBox("Rotina Interna. Regerar exportação para o sisweb. Deseja continuar?", vbYesNo, "Corporator")
                If vbResult = vbYes Then
                    Call Regera_Log_Vou
                End If
            End If

            If objFiltro.iAtualizarClientesRetroativo = MARCADO Then
                vbResult = MsgBox("Monta a hierarquia de vendedores dos voucher não pagos ?", vbYesNo, "Corporator")
                If vbResult = vbYes Then Call Regera_Hierarquia_Vou
            End If

            If objFiltro.iGerarComissaoRetroativo = MARCADO Then
                vbResult = MsgBox("Deseja gerar o comissionamento a partir do Corporator  ?", vbYesNo, "Corporator")
                If vbResult = vbYes Then Call Coinfo_ImportaComissao_150908(objImpCoinfo)
            End If

        End If

        If objFiltro.iImportArqsNovos = MARCADO Then

            
            sPosicao = "Coinfo_ImportarDados_Trata_Diretorio"
            
            'Gera a coleção de arquivos a serem atualizados
            lErro = Coinfo_ImportarDados_Trata_Diretorio(objImpCoinfo, colcolFile, iContador)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            'Para cada conjunto de importação
            For iIndice = 1 To iContador

                sPosicao = "Coinfo_ImportarDados_Trata_Diretorio2 Contador:" & CStr(iIndice) & " de " & CStr(iContador)
                
                'Coloca as informações do arquivo em questão em objImpCoinfo
                lErro = Coinfo_ImportarDados_Trata_Diretorio2(objImpCoinfo, colcolFile, iContador, iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                sPosicao = "Coinfo_ImportarArqTexto Contador:" & CStr(iIndice) & " de " & CStr(iContador)

                'importar dados de arquivos texto para tabelas de importacao
                lErro = Coinfo_ImportarArqTexto(objImpCoinfo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Next

        End If

        If iForcaSistemaIntegrado = DESMARCADO Then

            sPosicao = "Coinfo_AtualizarTiposEmp"

            'atualiza tabela de tipos de empresas (Tipos de Cliente e Tipos de Fornecedores)
            lErro = Coinfo_AtualizarTiposEmp(objImpCoinfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            sPosicao = "Coinfo_AtualizarEmpresas"

            'atualiza tabela de empresas (Clientes e Fornecedores)
            lErro = Coinfo_AtualizarEmpresas(objImpCoinfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            sPosicao = "Coinfo_AtualizarDestinos"

            'atualizar tabela de emissores (Fornecedores)
            lErro = Coinfo_AtualizarDestinos(objImpCoinfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            sPosicao = "Coinfo_AtualizarProdutos"

            'atualizar tabela de produtos
            lErro = Coinfo_AtualizarProdutos(objImpCoinfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            sPosicao = "Coinfo_AtualizarAcordos"

            'atualizar tabela de emissores (Fornecedores)
            lErro = Coinfo_AtualizarAcordos(objImpCoinfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            sPosicao = "Coinfo_AtualizarEmissores"

            'atualizar tabela de emissores (Fornecedores)
            lErro = Coinfo_AtualizarEmissores(objImpCoinfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            sPosicao = "Coinfo_AtualizaContab"

            'Atualiza a contabilização
            If objFiltro.iContabilizarNFs = MARCADO Then Call Coinfo_AtualizaContabNFs
            If objFiltro.iContabilizarFaturas = MARCADO Then Call Coinfo_AtualizaContabFaturas
            If objFiltro.iContabilizarVouchers = MARCADO Then Call Coinfo_AtualizarVouchers2(objImpCoinfo)

            'Gera o movimento de saida dos itens do impresso
            If objFiltro.iGerarMovEst = MARCADO Then Call Coinfo_AtualizarVouchers(objImpCoinfo)
            

            If Hour(Now) < 8 Or Hour(Now) >= 22 Then
                sPosicao = "RelLcto_Exclui"
                Call RelLcto_Exclui
            End If

        End If

        sPosicao = "Simula_Extracao_Vouchers_Sigav"
        Call Simula_Extracao_Vouchers_Sigav

    End If
    
    If objFiltro.iGerarBol = MARCADO Then
        sPosicao = "TRVVou_Trata_Boletos_Pendentes"
        lErro = CF("TRVVou_Trata_Boletos_Pendentes")
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    sPosicao = "Cria_Clientes_Cartao"
    Call Cria_Clientes_Cartao

    sPosicao = "Importa_Ocrs"
    Call Importa_Ocrs(objImpCoinfo)
        
    sPosicao = "Importa_Cobertura"
    lErro = Importa_Cobertura(objImpCoinfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    sPosicao = "Importa_OcrsCasos"
    lErro = Importa_OcrsCasos(objImpCoinfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Gera os arquivos de baixa e cancelamento de Baixa
    sPosicao = "Coinfo_Gera_Arquivos"
    lErro = Coinfo_Gera_Arquivos(objImpCoinfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    'Fecha glConexao1 e glConexao2
    sPosicao = "Fecha_Conexoes"
    Call Fecha_Conexoes
        
    Coinfo_ImportarDados = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarDados:

    Coinfo_ImportarDados = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130555)
     
    End Select
    
    Call Fecha_Conexoes
     
    Call Rotina_Erro(vbOKOnly, "A rotina de importação Coinfo_ImportarDados apresentou um erro próximo a posição " & sPosicao, gErr, Error)
     
    Exit Function

End Function

Private Function Coinfo_AtualizarTiposEmp(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza tipos de empresa

Dim lErro As Long, lTransacao As Long, alComando(0 To 2) As Long, iIndice As Integer, lNumIntArq As Long
Dim objTipoCli As ClassTipoCliente, itipoempr As Integer, sdesctpem As String
Dim objTipoForn As ClassTipoFornecedor
Dim iTotalRegCertos As Integer
Dim iTotalRegErr As Integer, lNumRegArqTexto As Long
Dim sChave As String

On Error GoTo Erro_Coinfo_AtualizarTiposEmp

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130727
    Next
    
    'Obtém o número de registros que vai ser atualizado
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fptipoempr")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130715
    
    'Lê os arquivos que ainda não foram atualizados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fptipoempr", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 130716
    
    'Busca o Primeiro
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130717
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Le os tipos de empresa
        sdesctpem = String(255, 0)
        lErro = Comando_Executar(alComando(1), "SELECT NumRegArqTexto, tipoempr, desctpem FROM CoinfoTipoEmpr WHERE NumIntArq = ? ORDER BY NumRegArqTexto", lNumRegArqTexto, itipoempr, sdesctpem, lNumIntArq)
        If lErro <> AD_SQL_SUCESSO Then gError 130718
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130719
        
        Do While lErro = AD_SQL_SUCESSO
        
            sChave = "Tipo: " & CStr(itipoempr) & " Descrição: " & sdesctpem
        
            Set objTipoCli = New ClassTipoCliente
            objTipoCli.iCodigo = itipoempr
        
            'faz a leitura para carregar info já existente que nao tenha vindo na importacao
            lErro = CF("TipoDeCliente_Le", objTipoCli)
            If lErro <> SUCESSO And lErro <> 28943 Then gError ERRO_SEM_MENSAGEM
            
            objTipoCli.sDescricao = sdesctpem
            
            'faz a gravacao, criando ou alterando o cadastro
            lErro = CF("TiposDeCliente_Grava_EmTrans", objTipoCli)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set objTipoForn = New ClassTipoFornecedor
            objTipoForn.iCodigo = itipoempr
            
            'faz a leitura para carregar info já existente que nao tenha vindo na importacao
            lErro = CF("TipoFornecedor_Le", objTipoForn)
            If lErro <> SUCESSO And lErro <> 12765 Then gError ERRO_SEM_MENSAGEM
            
            objTipoForn.sDescricao = sdesctpem
            
            lErro = CF("TipoFornecedor_Grava_EmTrans", objTipoForn)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130724
        
        Loop
        
        iTotalRegCertos = iTotalRegCertos + 1
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComando(0), gdtDataHoje)
        If lErro <> SUCESSO Then gError 130728
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130725
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130726
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Tipos de Cliente e Tipos de Fornecedor foram atualizados com sucesso. " & CStr(iTotalRegCertos) & " registros." & vbNewLine
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Coinfo_AtualizarTiposEmp = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarTiposEmp:

    Coinfo_AtualizarTiposEmp = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 130727
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130726
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130716, 130717, 130725
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130728
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130718, 130719, 130724
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOTIPOEMPR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130556)
     
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_TRV_ATUALIZACAO_ARQUIVO", gErr, "TipoEmpresa", lNumIntArq, lNumRegArqTexto, sChave)

    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Coinfo_AtualizarEmpresas1(ByVal objImpCoinfo As ClassImpCoinfo) As Long

Dim lErro As Long, lTransacao As Long, alComando(0 To 4) As Long, iIndice As Integer, lNumIntArq As Long
Dim objCli As ClassCliente, objForn As ClassFornecedor
Dim lcodemp As Long, snchave As String, sNome As String, sEndereco As String, sBairro As String, sCidade As String
Dim sUF As String, sCEP As String, sfone1 As String, sfone2 As String, sFax As String, sCgc As String
Dim sie As String, sContato As String, sCargo As String, sendercob As String, dlimcred As Double, scidcob As String
Dim sufcob As String, dDesconto As Double, sobs1 As String, sobs2 As String, sobs3 As String, sobs4 As String, sobs5 As String
Dim sobs6 As String, itipoempr As Integer, sCPF As String, sEmail As String, sbairrcob As String, scepcob As String
Dim iAtivo As Integer, inumban As Integer, lagenban As Long, lnconta As Long, scodccm As String
Dim colEndereco As Collection, objEndereco As ClassEndereco, objEnderecoCob As ClassEndereco, objEnderecoEntrega As ClassEndereco, objEnderecoForn As ClassEndereco
Dim sSelectCidade As String, sCidadeNova As String, objCidades As ClassCidades, iCodigoCidadeNovo As Integer, lAux As Long, lRegsLidos As Long, sNomeRedAlternativo As String

On Error GoTo Erro_Coinfo_AtualizarEmpresas1

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130729
        
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130730
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fpempresa", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 130732
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130733
    
    Do While lErro = AD_SQL_SUCESSO
    
        'cadastrar cidades que estejam faltando
        sSelectCidade = "SELECT cidade FROM CoinfoEmpresa WHERE cidade <> '' AND NumIntArq = ? AND NOT EXISTS (SELECT * FROM Cidades WHERE Descricao = cidade)" & _
            "UNION SELECT cidcob FROM CoinfoEmpresa WHERE cidcob <> '' AND NumIntArq = ? AND NOT EXISTS (SELECT * FROM Cidades WHERE Descricao = cidcob)"
        sCidadeNova = String(STRING_CIDADE, 0)
        lErro = Comando_Executar(alComando(3), sSelectCidade, sCidadeNova, lNumIntArq, lNumIntArq)
        If lErro <> AD_SQL_SUCESSO Then gError 130747
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130748
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigoCidadeNovo)
            If lErro <> SUCESSO Then gError 130749
            
            Set objCidades = New ClassCidades
        
            objCidades.iCodigo = iCodigoCidadeNovo
            objCidades.sDescricao = left(sCidadeNova, STRING_CIDADE)
            
            lErro = CF("Cidade_Grava_EmTrans", objCidades)
            If lErro <> SUCESSO Then gError 130750
            
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
                gError 130751
            End If
        
        Loop
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130742
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130743
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Coinfo_AtualizarEmpresas1 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarEmpresas1:

    Coinfo_AtualizarEmpresas1 = gErr
     
    Select Case gErr
          
        Case 130736, 130737, 130738, 130749, 130750
        
        Case 130729
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130730
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130743
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130732, 130733, 130742
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130741
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130734, 130735, 130740, 130747, 130748, 130751, 130752, 130753
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOEMPRESA", gErr)
        
        Case 130731, 130744
            Call Rotina_Erro(vbOKOnly, "ERRO_ERATIVAR_LOCKS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130754)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Coinfo_AtualizarEmpresas(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza clientes e fornecedores

Dim lErro As Long, lTransacaoExt As Long, alComandoExt(0 To 2) As Long, iIndice As Integer, lNumIntArq As Long
Dim objCli As ClassCliente, objForn As ClassFornecedor
Dim lcodemp As Long, snchave As String, sNome As String, sEndereco As String, sBairro As String, sCidade As String
Dim sUF As String, sCEP As String, sfone1 As String, sfone2 As String, sFax As String, sCgc As String
Dim sie As String, sContato As String, sCargo As String, sendercob As String, dlimcred As Double, scidcob As String
Dim sufcob As String, dDesconto As Double, sobs1 As String, sobs2 As String, sobs3 As String, sobs4 As String, sobs5 As String
Dim sobs6 As String, itipoempr As Integer, sCPF As String, sEmail As String, sbairrcob As String, scepcob As String
Dim iAtivo As Integer, inumban As Integer, lagenban As Long, lnconta As Long, scodccm As String
Dim colEndereco As Collection, objEndereco As ClassEndereco, objEnderecoCob As ClassEndereco, objEnderecoEntrega As ClassEndereco, objEnderecoForn As ClassEndereco
Dim sSelectCidade As String, sCidadeNova As String, objCidades As ClassCidades, iCodigoCidadeNovo As Integer, lAux As Long, lRegsLidos As Long, sNomeRedAlternativo As String
Dim lTransacao As Long, alComando(1 To 3) As Long, lNumRegArqTexto As Long
Dim colClientesContatos As Collection, objClienteContato As ClassClienteContatos
Dim bAchou As Boolean, iTotalRegCertos As Integer, iTotalRegErr As Integer
Dim iAlmoxarifado As Integer, sEmail2 As String, alCmdForaTrans(1 To 1) As Long
Dim objEnderecoTRV As ClassEnderecoTRV, objFilialCliente As ClassFilialCliente
Dim objFilialCliCategoria As ClassFilialCliCategoria, colCategoriaItem As New Collection
Dim iTipoCliAnt As Integer, scodgru As String, objFilialForn As ClassFilialFornecedor
Dim colCategoriaItemForn As New Collection, icodprom As Integer, iAux As Integer
Dim sChave As String

On Error GoTo Erro_Coinfo_AtualizarEmpresas
   
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 130729
    Next
    
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        alCmdForaTrans(iIndice) = Comando_AbrirExt(glConexao3)
        If alCmdForaTrans(iIndice) = 0 Then gError 130729
    Next
        
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpempresa")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abertura de transação 'Aquivo
    lTransacaoExt = Transacao_AbrirExt(glConexao1)
    If lTransacaoExt = 0 Then gError 130730
    
    lErro = Coinfo_AtualizarEmpresas1(objImpCoinfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Le os arquivos não atualizados
    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, lNumIntArq, "fpempresa", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 130732
    
    lErro = Comando_BuscarProximo(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130733
    
    Do While lErro = AD_SQL_SUCESSO
    
        'inicializar buffers de leitura p/strings
        snchave = String(255, 0)
        sNome = String(255, 0)
        sEndereco = String(255, 0)
        sBairro = String(255, 0)
        sCidade = String(255, 0)
        sUF = String(255, 0)
        sCEP = String(255, 0)
        sfone1 = String(255, 0)
        sfone2 = String(255, 0)
        sFax = String(255, 0)
        sCgc = String(255, 0)
        sie = String(255, 0)
        sContato = String(255, 0)
        sCargo = String(255, 0)
        sendercob = String(255, 0)
        scidcob = String(255, 0)
        sufcob = String(255, 0)
        sobs1 = String(255, 0)
        sobs2 = String(255, 0)
        sobs3 = String(255, 0)
        sobs4 = String(255, 0)
        sobs5 = String(255, 0)
        sobs6 = String(255, 0)
        sCPF = String(255, 0)
        sEmail = String(255, 0)
        sbairrcob = String(255, 0)
        scepcob = String(255, 0)
        scodccm = String(255, 0)
        sEmail2 = String(255, 0)
        scodgru = String(255, 0)
        
        'Usa uma outra conexão fora de transação para não perder o próximo no commit da transação do cliente
        'e não travar o update
        lErro = Comando_Executar(alCmdForaTrans(1), "SELECT NumRegArqTexto, codemp, nchave, nome, endereco, bairro, cidade, UF, cep, fone1, fone2, fax, cgc, ie, contato, cargo, endercob, limcred, cidcob, ufcob, desconto, obs1, obs2, obs3, obs4, obs5, obs6, tipoempr, cpf, email, bairrcob, cepcob, Ativo, numban, agenban, nconta, codccm, Email2, codgru, codprom FROM CoinfoEmpresa WHERE NumIntArq = ? AND NoCorporatorEm = ? ORDER BY NumRegArqTexto", _
            lNumRegArqTexto, lcodemp, snchave, sNome, sEndereco, sBairro, sCidade, sUF, sCEP, sfone1, sfone2, sFax, sCgc, sie, sContato, sCargo, sendercob, dlimcred, scidcob, sufcob, dDesconto, sobs1, sobs2, sobs3, sobs4, sobs5, sobs6, itipoempr, sCPF, sEmail, sbairrcob, scepcob, iAtivo, inumban, lagenban, lnconta, scodccm, sEmail2, scodgru, icodprom, lNumIntArq, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 130734
        
        lErro = Comando_BuscarProximo(alCmdForaTrans(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130735
        
        Do While lErro = AD_SQL_SUCESSO
        
            sChave = "Código: " & CStr(lcodemp) & " Nome: " & sNome
        
            lRegsLidos = lRegsLidos + 1

            'Abertura de transação para gravação cliente a cliente
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 130730

            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 130729
            Next

            'se já houver um cliente com este nome reduzido mas com codigo diferente entao concatenar com o codigo da empresa
            lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Fornecedores WHERE Codigo <> ? AND NomeReduzido = ?", lAux, lcodemp, snchave)
            If lErro <> AD_SQL_SUCESSO Then gError 130752
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130753
            
            sNomeRedAlternativo = snchave
            If lErro = AD_SQL_SUCESSO Then
                'Se achou com nome reduzido igual -> altera
                sNomeRedAlternativo = left(snchave, 40) & " (" & CStr(lcodemp) & ")"
            Else
                'se já houver um cliente com este nome reduzido mas com codigo diferente entao concatenar com o codigo da empresa
                lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Clientes WHERE Codigo <> ? AND NomeReduzido = ?", lAux, lcodemp, snchave)
                If lErro <> AD_SQL_SUCESSO Then gError 130752
                
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130753
                
                If lErro = AD_SQL_SUCESSO Then
                    'Se achou com nome reduzido igual -> altera
                    sNomeRedAlternativo = left(snchave, 40) & " (" & CStr(lcodemp) & ")"
                End If
            End If
                        
            Set objCli = New ClassCliente
            Set colEndereco = New Collection
            Set objEndereco = New ClassEndereco
            Set objEnderecoTRV = New ClassEnderecoTRV
            Set objEnderecoCob = New ClassEndereco
            Set objEnderecoEntrega = New ClassEndereco
            Set objFilialCliente = New ClassFilialCliente
            Set colCategoriaItem = New Collection
            
            objCli.lCodigo = lcodemp
        
            'faz a leitura para carregar info já existente que nao tenha vindo na importacao
            lErro = CF("Cliente_Le", objCli)
            If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                       
            iTipoCliAnt = objCli.iTipo
            
            objFilialCliente.lCodCliente = objCli.lCodigo
            objFilialCliente.iCodFilial = FILIAL_MATRIZ

            If lErro = SUCESSO Then
            
                'faz a leitura para carregar info já existente que nao tenha vindo na importacao
                lErro = CF("Cliente_Le_Customizado", objCli, True)
                If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM

                'Lê as Categoria da Filial do Cliente
                lErro = CF("FilialCliente_Le_Categorias", objFilialCliente, colCategoriaItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Set objCli.colCategoriaItem = colCategoriaItem
                              
            Else
            
                Set objFilialCliCategoria = New ClassFilialCliCategoria
                
                objFilialCliCategoria.sCategoria = TRV_CATEGORIA_CONDFAT
                objFilialCliCategoria.sItem = TRV_CATEGORIA_CONDFAT_SEMANALMENTE
                objFilialCliCategoria.iFilial = FILIAL_MATRIZ
                objFilialCliCategoria.lCliente = objCli.lCodigo
                
                objCli.colCategoriaItem.Add objFilialCliCategoria
                
                Set objFilialCliCategoria = New ClassFilialCliCategoria
                
                objFilialCliCategoria.sCategoria = TRV_CATEGORIA_COBRANCA
                objFilialCliCategoria.sItem = TRV_CATEGORIA_COBRANCA_NORMAL
                objFilialCliCategoria.iFilial = FILIAL_MATRIZ
                objFilialCliCategoria.lCliente = objCli.lCodigo
                
                objCli.colCategoriaItem.Add objFilialCliCategoria
                
                Set objFilialCliCategoria = New ClassFilialCliCategoria
                
                objFilialCliCategoria.sCategoria = TRV_CATEGORIA_ENVIODECOBRANCA
                objFilialCliCategoria.sItem = TRV_CATEGORIA_ENVIODECOBRANCA_SIM
                objFilialCliCategoria.iFilial = FILIAL_MATRIZ
                objFilialCliCategoria.lCliente = objCli.lCodigo
                
                objCli.colCategoriaItem.Add objFilialCliCategoria
                
                If objCli.iTipo = 10 Then
                    objCli.iCondicaoPagto = 1
                Else
                    objCli.iCondicaoPagto = 10
                End If
            
            End If
            
            bAchou = False
            For Each objFilialCliCategoria In objCli.colCategoriaItem
                If objFilialCliCategoria.sCategoria = TRV_CATEGORIA_RESPONSAVEL Then
                    'se era 10 ou 15 e continua 10 ou 15 ou não era os mesmo e contnua não sendo NÃO TROCA
                    If Not (((iTipoCliAnt = 15 Or iTipoCliAnt = 10 Or iTipoCliAnt = 7) And (itipoempr = 15 Or itipoempr = 10 Or itipoempr = 7)) Or ((iTipoCliAnt <> 15 And iTipoCliAnt <> 10 And iTipoCliAnt <> 7) And (itipoempr <> 15 And itipoempr <> 10 And itipoempr <> 7))) Then
                        If itipoempr <> 15 And itipoempr <> 10 And itipoempr <> 7 Then
                            objFilialCliCategoria.sItem = TRV_CATEGORIA_RESPONSAVEL_PROMOTOR
                        Else
                            objFilialCliCategoria.sItem = TRV_CATEGORIA_RESPONSAVEL_CALLCENTER
                        End If
                    End If
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objFilialCliCategoria = New ClassFilialCliCategoria
                
                objFilialCliCategoria.sCategoria = TRV_CATEGORIA_RESPONSAVEL
                If itipoempr <> 15 And itipoempr <> 10 And itipoempr <> 7 Then
                    objFilialCliCategoria.sItem = TRV_CATEGORIA_RESPONSAVEL_PROMOTOR
                Else
                    objFilialCliCategoria.sItem = TRV_CATEGORIA_RESPONSAVEL_CALLCENTER
                End If
                objFilialCliCategoria.iFilial = FILIAL_MATRIZ
                objFilialCliCategoria.lCliente = objCli.lCodigo
                
                objCli.colCategoriaItem.Add objFilialCliCategoria
            End If
            
            objEnderecoTRV.sEmail2 = left(sEmail2, STRING_EMAIL)
            Set objEndereco.objInfoUsu = objEnderecoTRV
                       
            With objEndereco
                .sEndereco = left(sEndereco, STRING_ENDERECO)
                .sBairro = left(sBairro, STRING_BAIRRO)
                .sCidade = left(sCidade, STRING_CIDADE)
                .sSiglaEstado = left(Trim(sUF), 2)
                .sCEP = left(Replace(sCEP, "-", ""), 8)
                .sTelefone1 = left(sfone1, STRING_TELEFONE)
                .sTelefone2 = left(sfone2, STRING_TELEFONE)
                .sFax = left(sFax, STRING_FAX)
                .sContato = left(sContato, STRING_CONTATO)
                .sEmail = left(sEmail, STRING_EMAIL)
                .sEmail2 = left(sEmail2, STRING_EMAIL)
                .iCodigoPais = PAIS_BRASIL
            End With
            
            'Se a região de venda ainda não foi preenchida
            If objCli.iRegiao = 0 Then
                Select Case sUF
                    Case "PR", "SC"
                        objCli.iRegiao = 1
                    Case "MG", "BH", "MS", "MT"
                        objCli.iRegiao = 2
                    Case "RS"
                        objCli.iRegiao = 3
                    Case "PE", "AL", "AM", "AR", "BA", "CE", "MA", "PA", "PB", "PI", "RN", "RO", "RR", "SE"
                        objCli.iRegiao = 4
                    Case "DF"
                        objCli.iRegiao = 2
                    Case "SP", "", "US"
                        objCli.iRegiao = 6
                    Case "RJ", "ES"
                        objCli.iRegiao = 7
                    Case Else
                        objCli.iRegiao = 4
                End Select
                If scodgru = "251" Then
                    objCli.iRegiao = 10
                End If
            End If
            
            With objEnderecoEntrega
                .sSiglaEstado = left(Trim(sUF), 2)
                .iCodigoPais = PAIS_BRASIL
            End With
            
            'Mantém o endereço de entrega que não vem via importação
            If objCli.lEnderecoEntrega <> 0 Then
                objEnderecoEntrega.lCodigo = objCli.lEnderecoEntrega
                
                lErro = CF("Endereco_Le", objEnderecoEntrega)
                If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                
                lErro = CF("Endereco_Le_Customizado", objEnderecoEntrega)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                
            End If
            
            With objEnderecoCob
                .sEndereco = left(sendercob, STRING_ENDERECO)
                .sBairro = left(sbairrcob, STRING_BAIRRO)
                .sCidade = left(scidcob, STRING_CIDADE)
                .sSiglaEstado = sufcob
                .sCEP = left(Replace(scepcob, "-", ""), 8)
                .iCodigoPais = PAIS_BRASIL
            End With
            
            colEndereco.Add objEndereco
            colEndereco.Add objEnderecoEntrega
            colEndereco.Add objEnderecoCob
            
            With objCli
                .sNomeReduzido = left(sNomeRedAlternativo, STRING_CLIENTE_NOME_REDUZIDO)
                .sRazaoSocial = left(sNome, STRING_CLIENTE_RAZAO_SOCIAL)
                .sCgc = IIf(Len(sCgc) <> 0, sCgc, sCPF)
                .sInscricaoEstadual = left(sie, STRING_INSCR_EST)
                .iTipo = itipoempr
                .sInscricaoMunicipal = left(scodccm, STRING_INSCR_MUN)
                .dLimiteCredito = dlimcred
                .dDesconto = dDesconto / 100
                .iAtivo = iAtivo
                If objImpCoinfo.sSiglaEmpresa = "TVA" Then .iVendedor = icodprom
            End With
 
            Set objForn = New ClassFornecedor
            objForn.lCodigo = lcodemp
            
            'faz a leitura para carregar info já existente que nao tenha vindo na importacao
            lErro = CF("Fornecedor_Le", objForn)
            If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
            
                Set objFilialForn = New ClassFilialFornecedor
                Set colCategoriaItemForn = New Collection
                
                'Lê as categoria deste fornecedor
                 objFilialForn.lCodFornecedor = objForn.lCodigo
                 objFilialForn.iCodFilial = FILIAL_MATRIZ
                 
                 'Alterado em 03/03/04 por Luiz Nogueira
                 lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilialForn, colCategoriaItemForn)
                 If lErro <> SUCESSO And lErro <> 127085 Then gError ERRO_SEM_MENSAGEM
                 
                 Set objForn.colCategoriaItem = colCategoriaItemForn
            
            End If
            
            'atribuir os campos vindos do arquivo texto
            Set objEnderecoForn = New ClassEndereco
            
            With objEnderecoForn
                .sEndereco = left(sEndereco, STRING_ENDERECO)
                .sBairro = left(sBairro, STRING_BAIRRO)
                .sCidade = left(sCidade, STRING_CIDADE)
                .sSiglaEstado = left(Trim(sUF), 2)
                .sCEP = left(Replace(sCEP, "-", ""), 8)
                .sTelefone1 = left(sfone1, STRING_TELEFONE)
                .sTelefone2 = left(sfone2, STRING_TELEFONE)
                .sFax = left(sFax, STRING_FAX)
                .sContato = left(sContato, STRING_CONTATO)
                .sEmail = left(sEmail, STRING_EMAIL)
                .sEmail2 = left(sEmail2, STRING_EMAIL)
                .iCodigoPais = PAIS_BRASIL
            End With
            
            With objForn
                .sNomeReduzido = left(sNomeRedAlternativo, STRING_CLIENTE_NOME_REDUZIDO)
                .sRazaoSocial = left(sNome, STRING_CLIENTE_RAZAO_SOCIAL)
                .sCgc = left(IIf(Len(sCgc) <> 0, sCgc, sCPF), STRING_CGC)
                .sInscricaoEstadual = left(sie, STRING_INSCR_EST)
                .dDesconto = dDesconto
                .iTipo = itipoempr
                .iAtivo = iAtivo
                 If inumban <> 0 Then .iBanco = inumban
                 If lagenban <> 0 Then .sAgencia = CStr(lagenban)
                 If lnconta <> 0 Then .sContaCorrente = CStr(lnconta)
                
                .sInscricaoMunicipal = left(scodccm, STRING_INSCR_MUN)
            End With
            
            Set colClientesContatos = New Collection
            Set objClienteContato = New ClassClienteContatos

            objClienteContato.lCliente = objCli.lCodigo
            objClienteContato.iFilialCliente = FILIAL_MATRIZ

            'Le os contatos já cadastrados
            lErro = CF("ClienteContatos_Le_Todos", colClientesContatos, objClienteContato)
            If lErro <> SUCESSO And lErro <> 102574 Then gError ERRO_SEM_MENSAGEM

            For Each objClienteContato In colClientesContatos
                objClienteContato.lCliente = objCli.lCodigo
                objClienteContato.iFilialCliente = FILIAL_MATRIZ
            Next

            'Se o contato do endereço estiver preenchido
            If Len(Trim(objEndereco.sContato)) > 0 Or Len(Trim(objEndereco.sTelefone1)) > 0 Then

                'Verifica se o contato do endereço está incluso
                bAchou = False
                For Each objClienteContato In colClientesContatos
                    If (objClienteContato.sContato = objEndereco.sContato Or (objClienteContato.sContato = "Sem nome" And objEndereco.sContato = "")) And objClienteContato.sTelefone = objEndereco.sTelefone1 Then
                        bAchou = True
                        Exit For
                    End If
                Next
                
                iAux = 0
                For Each objClienteContato In colClientesContatos
                    If iAux <= objClienteContato.iCodigo Then iAux = objClienteContato.iCodigo
                Next
    
                'Se não -> Insere
                If Not bAchou Then

                    Set objClienteContato = New ClassClienteContatos

                    objClienteContato.lCliente = objCli.lCodigo
                    objClienteContato.iFilialCliente = FILIAL_MATRIZ

                    If Len(Trim(objEndereco.sContato)) = 0 Then
                        objClienteContato.sContato = "Sem nome"
                    Else
                        objClienteContato.sContato = objEndereco.sContato
                    End If
                    objClienteContato.sTelefone = objEndereco.sTelefone1

                    If colClientesContatos.Count = 0 Then
                        objClienteContato.iPadrao = MARCADO
                    End If

                    colClientesContatos.Add objClienteContato

                    objClienteContato.iCodigo = iAux + 1

                End If

            End If
            
            Set objCli.objInfoUsu = Nothing
                           
            'faz a gravacao, criando ou alterando o cadastro
            lErro = CF("Cliente_Grava_EmTrans", objCli, colEndereco)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            If colClientesContatos.Count > 0 Then
                lErro = CF("ClienteContatos_Grava_EmTrans", colClientesContatos)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
            lErro = CF("TRVCliente_Grava", objCli)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Coinfo_Associa_Cliente_Alm(objCli, iAlmoxarifado)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = CF("Fornecedor_Grava_EmTrans", objForn, objEnderecoForn)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            iTotalRegCertos = iTotalRegCertos + 1
            
            lErro = Comando_Executar(alComando(1), "UPDATE CoinfoEmpresa SET NoCorporatorEm = ? WHERE NumIntArq = ? AND codemp = ? AND NumRegArqTexto = ?", gdtDataHoje, lNumIntArq, lcodemp, lNumRegArqTexto)
            If lErro <> SUCESSO Then gError 130739
            
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
            
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 130743
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_BuscarProximo(alCmdForaTrans(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130740
        
        Loop
        
        lErro = Comando_ExecutarPos(alComandoExt(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComandoExt(0), gdtDataHoje)
        If lErro <> SUCESSO Then gError 130741
    
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130742
        
    Loop
    
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        Call Comando_Fechar(alCmdForaTrans(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 130743
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Clientes e Fornecedores foram atualizados com sucesso. " & CStr(iTotalRegCertos) & " registros." & vbNewLine
                
    Coinfo_AtualizarEmpresas = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarEmpresas:

    Coinfo_AtualizarEmpresas = gErr
        
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 130729
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130730
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130743
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130732, 130733, 130742
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130741
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130734, 130735, 130740, 130747, 130748, 130751, 130752, 130753
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOEMPRESA", gErr)
        
        Case 130731, 130744
            Call Rotina_Erro(vbOKOnly, "ERRO_REATIVAR_LOCKS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130557)
     
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_TRV_ATUALIZACAO_ARQUIVO", gErr, "Empresa", lNumIntArq, lNumRegArqTexto, sChave)
     
    Call Transacao_RollbackExt(lTransacaoExt)
    Call Transacao_Rollback
    
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        Call Comando_Fechar(alCmdForaTrans(iIndice))
    Next
        
    Exit Function

End Function

Private Function Coinfo_AtualizarProdutos(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza clientes e fornecedores

Dim lErro As Long, lTransacao As Long, alComando(0 To 5) As Long, iIndice As Integer, lNumIntArq As Long
Dim objProduto As ClassProduto, lNumRegArqTexto As Long, sCodPro As String, sDescricao As String
Dim colTabelaPrecoItem As Collection
Dim iTotalRegCertos As Integer
Dim iTotalRegErr As Integer
Dim sChave As String, sCodProdAux As String
Dim iCount As Integer, bContinuar As Boolean

On Error GoTo Erro_Coinfo_AtualizarProdutos

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next
    
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpproduto")
    If lErro <> SUCESSO Then gError 182714
    
    bContinuar = True
    Do While bContinuar
        
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 130804
        
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
            lNumIntArq, "fpproduto", DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 130805
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
        
        Do While lErro = AD_SQL_SUCESSO
        
            sCodProdAux = String(255, 0)
            sDescricao = String(255, 0)
            lErro = Comando_ExecutarPos(alComando(1), "SELECT CodPro, Descricao FROM CoinfoProduto WHERE NumIntArq = ? AND DataAtualizado = ? ORDER BY NumRegArqTexto", 0, sCodProdAux, sDescricao, lNumIntArq, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 130807
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808
            
            Do While lErro = AD_SQL_SUCESSO
            
                sChave = "Produto: " & sCodPro & " Descrição: " & sDescricao
            
                sCodPro = sCodProdAux
                Call TRVConverte_Prod(sCodPro)
    '            sCodPro = Replace(sCodPro, "/", " ")
    '            sCodPro = Replace(sCodPro, "\", " ")
    '            sCodPro = Replace(sCodPro, "-", " ")
    '            sCodPro = Replace(sCodPro, "_", " ")
    '            sCodPro = Replace(sCodPro, "&", " ")
            
                Set objProduto = New ClassProduto
                
                'Coloca o produto no formato do BD
                objProduto.sCodigo = sCodPro
    '            objProduto.sCodigo = objProduto.sCodigo & String(TAMANHO_SEGMENTO_PRODUTO - Len(objProduto.sCodigo), 32)
                
                'faz a leitura para carregar info já existente que nao tenha vindo na importacao
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 130809
        
                'se o produto já existe
                If lErro = SUCESSO Then
                
                    If objProduto.sDescricao <> sDescricao Then
                    
                        objProduto.sDescricao = sDescricao
                        
                        lErro = CF("Produto_Le_Customizado", objProduto)
                        If lErro <> SUCESSO Then gError 130810
                        
                        lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
                        If lErro <> SUCESSO Then gError 130810
                    
                        Set colTabelaPrecoItem = New Collection
                        lErro = CF("Produto_Le_TabelaPrecoItem", objProduto, colTabelaPrecoItem, giFilialEmpresa)
                        If lErro <> SUCESSO Then gError 130811
                        
                        lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
                        If lErro <> SUCESSO Then gError 130816
                                
                    End If
                
                Else 'produto novo
                
                    Set colTabelaPrecoItem = New Collection
                    
                    objProduto.sDescricao = left(sDescricao, STRING_PRODUTO_DESCRICAO)
                    objProduto.sNomeReduzido = objProduto.sCodigo & SEPARADOR & left(sDescricao, STRING_PRODUTO_NOME_REDUZIDO - Len(objProduto.sCodigo & SEPARADOR))
                    objProduto.iClasseUM = 2
                    objProduto.iNatureza = 7 'Outros
                    objProduto.sSiglaUMCompra = "un"
                    objProduto.sSiglaUMEstoque = "un"
                    objProduto.sSiglaUMVenda = "un"
                    objProduto.iTipo = 1
                    objProduto.iFaturamento = MARCADO
                    objProduto.iNivel = 1
                    
                    iTotalRegCertos = iTotalRegCertos + 1
                    
                    lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
                    If lErro <> SUCESSO Then gError 130817
                    
                End If
                
                lErro = Comando_ExecutarPos(alComando(3), "UPDATE CoinfoProduto SET DataAtualizado = ?", alComando(1), gdtDataHoje)
                If lErro <> SUCESSO Then gError 130813
                
                lErro = Rotina_Atualizar_AtualizaTelaBatch
                If lErro <> SUCESSO Then gError 182720
                
                iCount = iCount + 1
                If iCount = 100 Then
                    Exit Do
                End If
        
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130812
            
            Loop
            
            If iCount = 100 Then
                Exit Do
            End If
                
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComando(0), gdtDataHoje)
            If lErro <> SUCESSO Then gError 130813
        
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130814
            
        Loop
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 130815
        
        If iCount <> 100 Then
            iCount = 0
            bContinuar = False
        End If
        
    Loop
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Produtos foram atualizados com sucesso. " & CStr(iTotalRegCertos) & " registros." & vbNewLine
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_AtualizarProdutos = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarProdutos:

    Coinfo_AtualizarProdutos = gErr
     
    Select Case gErr
          
        Case 130809, 130810, 130811, 130816, 130817, 182714, 182720
        
        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130804
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130815
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130805, 130806, 130814
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130813
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130807, 130808, 130812
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOPRODUTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select
  
    Call Rotina_Erro(vbOKOnly, "ERRO_TRV_ATUALIZACAO_ARQUIVO", gErr, "Produto", lNumIntArq, lNumRegArqTexto, sChave)
  
    Exit Function

End Function

'Private Function Coinfo_AtualizarFaturasRec(ByVal objImpCoinfo As ClassImpCoinfo, ByVal lNumIntArq As Long, alComando() As Long, alComando2() As Long, bTemErro As Boolean, iTotalRegCertos As Integer, iTotalRegErr As Integer) As Long
'
''atualiza faturas no cr
'Dim lNumRegArqTexto As Long, lNumFat As Long, dtDataFat As Date, lcodemp As Long, dValor As Double, dtdataven As Date, iCancelado As Integer, dvaldcbrl As Double, dvalbru As Double, dtDataCanc As Date, iPago As Integer, dtDataPag As Date
'Dim lNumIntDocTit As Long, lClienteTit As Long, dtDataEmissaoTit As Date, iStatusTit As Integer, dValorTit As Double, dSaldoTit As Double
'Dim lErro As Long, iAchou As Integer, objTitRec As ClassTituloReceber
''??? criar colecao de objetos que irao gerar os titulos a receber dos cartoes
'Dim objParcelaReceber As ClassParcelaReceber
'Dim colParcelaReceber As colParcelaReceber
'Dim colComissaoEmissao As colComissao
'Dim colcolComissao As colcolComissao
'Dim colComissao As colComissao
'Dim colcolDesconto As colcolDesconto
'Dim colDesconto As colDesconto
'Dim objContabil As ClassContabil
'Dim colBaixaParcReceber As ColBaixaParcRec
'Dim objBaixaReceber As ClassBaixaReceber
'Dim lrecno As Long
'Dim lTransacao As Long
'Dim objImpCoInfoLog As ClassImpCoinfoLogAtu
'Dim colImpCoInfoLog As Collection
'Dim sCritica As String, objCtb As ClassImpCoinfoFatCRCtb
'Dim iComErro As Integer, objContabAutomatica As ClassContabAutomatica
'Dim bErro As Boolean
'Dim iFilialCorporator As Integer
'Dim iFilialCoinfo As Integer
'Dim sSigla As String
'Dim objImpCoinfoCartoes As ClassImpCoinfoCartoes
'Dim sCiaCart As String
'Dim iQuantParc As Integer
'Dim dComissao As Double
'Dim dTaxa As Double
'Dim dDeducoes As Double
'Dim iTipoDocCorporatorNFPag As Integer
'Dim lNumIntDocCorporatorNFPag As Long
'Dim objNFPag As New ClassNFsPag
'Dim objTitRecTRV As ClassTitulosRecTRV
'
'On Error GoTo Erro_Coinfo_AtualizarFaturasRec
'
'    'Le as faturas a receber
'    lErro = Comando_ExecutarPos(alComando(3), "SELECT RecNo, NumRegArqTexto, numfat, datafat, codemp, valor, dataven, cancelado, valdcbrl, valbru, datacanc, pago, datapag FROM CoinfoFatura WHERE CoinfoFatura.NumIntArq = ? AND CoinfoFatura.cpcr = ? AND CoinfoFatura.codest = ? AND CoinfoFatura.NoCorporatorEm = ? AND CoinfoFatura.Usuario <> ? ORDER BY CoinfoFatura.NumRegArqTexto", 0, _
'        lrecno, lNumRegArqTexto, lNumFat, dtDataFat, lcodemp, dValor, dtdataven, iCancelado, dvaldcbrl, dvalbru, dtDataCanc, iPago, dtDataPag, lNumIntArq, 0, objImpCoinfo.sSiglaEmpresa, DATA_NULA, "teste")
'    If lErro <> AD_SQL_SUCESSO Then gError 182519
'
'    'Busca a primeira
'    lErro = Comando_BuscarProximo(alComando(3))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182520
'
'    'Enquanto houver faturas não atualizadas
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        dTaxa = 0
'        dDeducoes = 0
'        dComissao = 0
'
'        lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", lcodemp, iFilialCorporator, iFilialCoinfo)
'        If lErro <> SUCESSO Then gError 182743
'
'        bErro = False
'
'        iAchou = 0
'
'        Set colImpCoInfoLog = New Collection
'        Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
'
'        objImpCoInfoLog.iTabela = TABELA_COINFOFATURA
'        objImpCoInfoLog.lNumIntArq = lNumIntArq
'        objImpCoInfoLog.lNumRegArqTexto = lNumRegArqTexto
'
'        'Se não tem data de vencimento é um título de cartão
'        If dtdataven = DATA_NULA Then
'
'            sSigla = TIPODOC_CARTAO_CRED_DEB
'
'            sCiaCart = String(255, 0)
'
'            'Le as faturas a receber para obter a Administradora e a quantidade de parcelas
'            lErro = Comando_Executar(alComando(12), "SELECT CiaCart, QuantParc FROM CoinfoItemFatura WHERE codest = ? AND Numfat = ? AND CodEmp = ? AND Cartao = 1 AND QuantParc <> 0 AND CiaCart <> '' ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
'                sCiaCart, iQuantParc, objImpCoinfo.sSiglaEmpresa, lNumFat, lcodemp)
'            If lErro <> AD_SQL_SUCESSO Then gError 182519
'
'            'Busca a primeira
'            lErro = Comando_BuscarProximo(alComando(12))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182520
'
'            If lErro <> AD_SQL_SUCESSO And iCancelado = DESMARCADO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_CARTAO_CRED_DEB & " e Número " & CStr(lNumFat) & " não possui itens de fatura associados.", objImpCoInfoLog, colImpCoInfoLog)
'            End If
'
'        Else
'            sSigla = TIPODOC_FATURA_SERVICO_CR
'        End If
'
'        'buscar a fatura como titulo a receber aberto
'        lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc, Cliente, DataEmissao, Status, Valor, Saldo FROM TitulosRec WHERE FilialEmpresa = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ?", 0, _
'            lNumIntDocTit, lClienteTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, sSigla, lNumFat, STATUS_EXCLUIDO)
'        If lErro <> AD_SQL_SUCESSO Then gError 182521
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComando(4))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182522
'
'        If lErro <> AD_SQL_SUCESSO Then
'
'            'buscar a fatura como titulo a receber baixado
'            lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc, Cliente, DataEmissao, Status, Valor, Saldo FROM TitulosRecBaixados WHERE FilialEmpresa = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ?", 0, _
'                lNumIntDocTit, lClienteTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, sSigla, lNumFat, STATUS_EXCLUIDO)
'            If lErro <> AD_SQL_SUCESSO Then gError 182523
'
'            'Busca o primeiro
'            lErro = Comando_BuscarProximo(alComando(5))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182524
'            If lErro = AD_SQL_SUCESSO Then iAchou = 2
'
'        Else
'            iAchou = 1
'        End If
'
'        'Título já pago (baixado) não pode ser alterado
'        'If iAchou = 2 Then gError 182532
'        If iAchou = 2 Then
'            'bErro = True
'            Call CoInfo_Adiciona_Log("O Título a receber Tipo " & sSigla & " e Número " & CStr(lNumFat) & " já está baixado e não pode ser alterado.", objImpCoInfoLog, colImpCoInfoLog, True)
'        End If
'
'        'Abertura de transação
'        lTransacao = Transacao_Abrir()
'        If lTransacao = 0 Then gError 182525
'
'        'preencher o objeto e mandar gravar
'        Set objTitRec = New ClassTituloReceber
'
'        If iAchou <> 0 Then
'            objTitRec.lNumIntDoc = lNumIntDocTit
'        End If
'
'        objTitRec.lCliente = lcodemp
'        objTitRec.iFilial = FILIAL_MATRIZ
'        objTitRec.lNumTitulo = lNumFat
'        objTitRec.dtDataEmissao = dtDataFat
'        objTitRec.dValor = dValor 'dvalbru
'        objTitRec.dSaldo = dValor 'dvalbru
'        objTitRec.iFilialEmpresa = iFilialCorporator  'FILIAL_MATRIZ 'FIXO -> ATÉ SEGUNDA ORDEM
'
'        'se não existir no sistema e não estiver cancelado
'        'Insere o título
'        If iAchou = 0 And iCancelado = DESMARCADO And Not bErro Then
'
'            'Se for um título normal (sem ser de cartão)
'            If sSigla = TIPODOC_FATURA_SERVICO_CR Then
'
'                objTitRec.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CR
'                objTitRec.iNumParcelas = 1
'
'                Set objParcelaReceber = New ClassParcelaReceber
'                Set colParcelaReceber = New colParcelaReceber
'                Set colComissaoEmissao = New colComissao
'                Set colcolComissao = New colcolComissao
'                Set colComissao = New colComissao
'                Set colcolDesconto = New colcolDesconto
'                Set colDesconto = New colDesconto
'
'                objParcelaReceber.iNumParcela = 1
'                objParcelaReceber.dtDataVencimento = dtdataven
'                objParcelaReceber.dtDataVencimentoReal = dtdataven
'                objParcelaReceber.dValor = dValor 'dvalbru
'                objParcelaReceber.dValorOriginal = dValor 'dvalbru
'                objParcelaReceber.dtDataCredito = DATA_NULA
'                objParcelaReceber.dtDataDepositoCheque = DATA_NULA
'                objParcelaReceber.dtDataEmissaoCheque = DATA_NULA
'                objParcelaReceber.dtDataPrevReceb = DATA_NULA
'                objParcelaReceber.dtDataProxCobr = DATA_NULA
'                objParcelaReceber.dtDataTransacaoCartao = DATA_NULA
'                objParcelaReceber.dtDesconto1Ate = DATA_NULA
'                objParcelaReceber.dtDesconto2Ate = DATA_NULA
'                objParcelaReceber.dtDesconto3Ate = DATA_NULA
'                objParcelaReceber.dtValidadeCartao = DATA_NULA
'
'                colParcelaReceber.AddObj objParcelaReceber
'
'                'Desconsiderar o desconto
'    '            If dvaldcbrl <> 0 Then
'    '
'    '                'Desconto tipo valor fixo
'    '                colDesconto.Add 1, dtdataven, dvaldcbrl
'    '
'    '            End If
'
'
'                colcolDesconto.Add colDesconto
'                colcolComissao.Add colComissao
'
'                'Grava o Título a receber
'                lErro = CF("TituloReceber_Grava_EmTrans", objTitRec, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
'    '            If lErro <> SUCESSO Then gError 182526
'
'                If lErro = SUCESSO Then
'
'                    If objTitRec.dtDataEmissao > #10/31/2006# Then
'
'                        'Coloca o desconto no COFINS
'                        objTitRec.dCOFINSRetido = dvaldcbrl
'
'                        'Coloca a comissão no CSLL
'                        Call Coinfo_Obtem_Comissao_Cliente(objTitRec.lNumTitulo, objImpCoinfo.sSiglaEmpresa, dComissao)
'                        objTitRec.dCSLLRetido = dComissao
'
'                        Set objContabAutomatica = New ClassContabAutomatica
'                        Set objCtb = New ClassImpCoinfoFatCRCtb
'
'                        'gravar a contabilizacao
'                        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.GeraContabilizacao(objTitRec, Nothing)
'                            If lErro = SUCESSO Then
'
'                                lErro = objContabAutomatica.Finaliza_Contab()
'
'
'                            Else
'                                bErro = True
'                                Call CoInfo_Adiciona_Log("Erro na contabilização do Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'                            End If
'
'                        Else
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Erro na contabilização do Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                        End If
'
'                        Set objContabAutomatica = Nothing
'                        Call objCtb.Encerrando
'                        Set objCtb = Nothing
'
'                    End If
'
'                End If
'
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & " não pode ser inserido.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & " foi gravado com sucesso.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            Else
'                'Tratar cartão
'
'                Set objImpCoinfoCartoes = New ClassImpCoinfoCartoes
'
'                objImpCoinfoCartoes.dtDataDep = dtDataFat
'                objImpCoinfoCartoes.iFilialEmpresa = iFilialCorporator
'                objImpCoinfoCartoes.dValor = dValor
'                objImpCoinfoCartoes.lNumTitulo = lNumFat
'
'                objImpCoinfoCartoes.sCiaCart = sCiaCart
'                objImpCoinfoCartoes.iQuantParc = iQuantParc
'
'                'Gera o parcelamento do título e grava
'                lErro = Coinfo_GerarTitRec_Cartao(objImpCoinfoCartoes)
'                If lErro <> SUCESSO Then 'gError 182600
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("Não foi possível gerar o Título a Receber para Data " & Format(objImpCoinfoCartoes.dtDataDep, "dd/mm/yyyy") & " Cia " & objImpCoinfoCartoes.sCiaCart & " e Quantidade de Parcelas " & CStr(objImpCoinfoCartoes.iQuantParc), objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("Foi gerado o Título a Receber Tipo " & TIPODOC_CARTAO_CRED_DEB & " número " & CStr(objImpCoinfoCartoes.lNumTitulo) & " para Data " & Format(objImpCoinfoCartoes.dtDataDep, "dd/mm/yyyy") & " Cia " & objImpCoinfoCartoes.sCiaCart & " e Quantidade de Parcelas " & CStr(objImpCoinfoCartoes.iQuantParc), objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'                objTitRec.lNumIntDoc = objImpCoinfoCartoes.lNumIntTitulo
'                objTitRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB
'                objTitRec.dValor = objImpCoinfoCartoes.dValorTitulo
'
'                dTaxa = objImpCoinfoCartoes.dValorTaxa
'                dDeducoes = objImpCoinfoCartoes.dDeducaoValorPresente
'
'                If objTitRec.dtDataEmissao > #10/31/2006# Then
'
'                    Set objContabAutomatica = New ClassContabAutomatica
'                    Set objCtb = New ClassImpCoinfoFatCRCtb
'
'                    'gravar a contabilizacao
'                    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
'                    If lErro = SUCESSO Then
'
'                        lErro = objContabAutomatica.GeraContabilizacao(objTitRec, objImpCoinfoCartoes)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.Finaliza_Contab()
'
'
'                        Else
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Erro na contabilização da comissão Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'                        End If
'
'                    Else
'                        bErro = True
'                        Call CoInfo_Adiciona_Log("Erro na contabilização da comissão Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                    End If
'
'                    Set objContabAutomatica = Nothing
'                    Call objCtb.Encerrando
'                    Set objCtb = Nothing
'
'                End If
'
'                'Gera a comissão a pagar
'                objTitRec.lCliente = lcodemp
'                objTitRec.lNumTitulo = lNumFat
'
'                Set objTitRecTRV = New ClassTitulosRecTRV
'
'                objTitRecTRV.lNumIntDocTitRec = objTitRec.lNumIntDoc
'                objTitRecTRV.dValorBruto = objImpCoinfoCartoes.dValor
'                objTitRecTRV.dValorDeducoes = objImpCoinfoCartoes.dDeducaoValorPresente
'                objTitRecTRV.dValorTarifa = objImpCoinfoCartoes.dValorTaxa
'
'                lErro = Coinfo_GerarComissaoCartao(objImpCoinfo, objTitRec, objTitRecTRV)
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & " não pode gerar a comissão.", objImpCoInfoLog, colImpCoInfoLog)
'                End If
'
'                lErro = CF("TitulosRecTRV_Grava", objTitRecTRV)
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & " não pode gerar a comissão.", objImpCoInfoLog, colImpCoInfoLog)
'                End If
'
'            End If
'
'        End If
'
'        'Se o Título tiver sido cancelado já existe sem estar baixado (não existiu a importação)
'        'Excluir
'        If iCancelado = MARCADO And iAchou = 1 And Not bErro Then
'
'            lErro = CF("TituloReceber_Le", objTitRec)
'            'If lErro <> SUCESSO And lErro <> 26061 Then gError 182533
'            If lErro <> SUCESSO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("O Título Tipo " & sSigla & " e Número " & CStr(lNumFat) & " não pode ser lido.", objImpCoInfoLog, colImpCoInfoLog)
'           Else
'                Set objContabil = New ClassContabil
'                Call objContabil.Contabil_Inicializa_Contabilidade4(23, MODULO_CONTASARECEBER)
'
'                lErro = CF("TituloReceber_Exclui_EmTransacao", objTitRec, objContabil)
'                'If lErro <> SUCESSO Then gError 182534
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título Tipo " & sSigla & " e Número " & CStr(lNumFat) & " não pode ser excluído.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("O Título Tipo " & sSigla & " e Número " & CStr(lNumFat) & " foi excluido por estar com status de cancelado.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'                Set objContabil = Nothing
'
'            End If
'
'            'Se for referente a cartão de crédito
'            If objTitRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
'
'                '##########################################################
'                'Tem que excluir a comissão ?????
'                '##########################################################
'
'                'Le as faturas a receber para obter a Administradora e a quantidade de parcelas
'                lErro = Comando_Executar(alComando(14), "SELECT C.TipoDocCorporator, C.NumIntDocCorporator  FROM CoinfoItemFatura AS I, CoinfoComissao AS C WHERE C.Serie = I.Serie AND I.NumVou = C.NumVou AND C.TipVou = I.TipoDoc AND C.CodEmp <> 0 AND I.codest = ? AND I.Numfat = ? AND I.CodEmp = ? AND I.Cartao = 1 AND I.QuantParc <> 0 AND I.CiaCart <> '' AND C.TipoDocCorporator <> 0 AND C.NumIntDocCorporator <> 0 ORDER BY C.NumIntArq DESC, C.NumRegArqTexto DESC", _
'                    iTipoDocCorporatorNFPag, lNumIntDocCorporatorNFPag, objImpCoinfo.sSiglaEmpresa, lNumFat, lcodemp)
'                If lErro <> AD_SQL_SUCESSO Then gError 182519
'
'                'Busca a primeira
'                lErro = Comando_BuscarProximo(alComando(14))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182520
'
'                'Se já teve uma comissão gerada
'                If lErro = AD_SQL_SUCESSO And Not bErro Then
'
'                    Set objNFPag = New ClassNFsPag
'
'                    objNFPag.lNumIntDoc = lNumIntDocCorporatorNFPag
'
'                    'Le a nota de comissão
'                    lErro = CF("NFPag_Le", objNFPag)
'                    If lErro = SUCESSO Then
'
'                        'Se ela ainda não foi faturada
'                        If objNFPag.lNumIntTitPag = 0 Then
'
'                            Set objContabil = New ClassContabil
'                            Call objContabil.Contabil_Inicializa_Contabilidade4(14, MODULO_CONTASAPAGAR)
'
'                            objNFPag.iExclusao = DESMARCADO
'
'                            lErro = CF("NFPag_Exclui_EmTransacao", objNFPag, objContabil)
'                            If lErro <> SUCESSO Then
'                                bErro = True
'                                Call CoInfo_Adiciona_Log("A comissão de número " & CStr(lNumFat) & " não pode ser excluída.", objImpCoInfoLog, colImpCoInfoLog)
'                            End If
'                            Set objContabil = Nothing
'
'                        End If
'
'                    Else
'                        Call CoInfo_Adiciona_Log("A comissão de número " & CStr(lNumFat) & " já foi faturada e por isso não pode ser cancelada.", objImpCoInfoLog, colImpCoInfoLog, True)
'                    End If
'
'                End If
'
'            End If
'
'        End If
'
'        'Se é um título novo ou se ele não estiver baixado e for para baixar
'        If iPago = MARCADO And iCancelado = DESMARCADO And iAchou <> 2 And Not bErro Then
'
'            If objTitRec.dtDataEmissao > #10/31/2006# Then
'                Call CoInfo_Adiciona_Log("O Título a Receber Tipo " & sSigla & " e Número " & CStr(lNumFat) & " possui data de emissão posterior a 01/11/2006, está baixado pelo Sigav e não está baixado pelo Corporator.", objImpCoInfoLog, colImpCoInfoLog, True)
'            Else
'
'                If sSigla = TIPODOC_FATURA_SERVICO_CR Then
'
'                    '++++++++++++++++ BAIXA POR PERDA
'                    Set colBaixaParcReceber = New ColBaixaParcRec
'                    Set objBaixaReceber = New ClassBaixaReceber
'
'                    If iAchou = 1 Then
'
'                        Set colParcelaReceber = New colParcelaReceber
'
'                        'Lê as Parcelas a Receber vinculadas ao Título
'                        lErro = CF("ParcelasReceber_Le", objTitRec, colParcelaReceber)
'                        If lErro <> SUCESSO Then gError 26380
'
'                        Set objParcelaReceber = colParcelaReceber.Item(1)
'
'                    End If
'
'                    'Inicializa BaixaReceber
'                    objBaixaReceber.iStatus = STATUS_LANCADO
'                    objBaixaReceber.iMotivo = MOTIVO_PERDA
'                    objBaixaReceber.dtData = dtDataPag
'                    objBaixaReceber.dtDataContabil = dtDataPag
'                    objBaixaReceber.dtDataRegistro = gdtDataHoje
'                    objBaixaReceber.sHistorico = "Importação de Título Baixado da CoInfo"
'
'                    colBaixaParcReceber.Add 0, 0, objParcelaReceber.lNumIntDoc, 0, STATUS_LANCADO, 0, 0, 0, dValor, dValor, 0
'
'                    'Grava BaixaReceber e BaixasReceberParcela associadas
'                    lErro = CF("BaixaReceber_Grava", objBaixaReceber, colBaixaParcReceber, objContabil, objTitRec.lCliente, objTitRec.iFilial)
'                    'If lErro <> SUCESSO Then gError 182527
'                    If lErro <> SUCESSO Then
'                        bErro = True
'                        Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & " não pode ser baixado.", objImpCoInfoLog, colImpCoInfoLog)
'                    Else
'                        Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CR & " e Número " & CStr(lNumFat) & " foi baixado por perda.", objImpCoInfoLog, colImpCoInfoLog, True)
'                    End If
'
'                Else
'                    'Tratar baixa dos cartões
'                    lErro = Coinfo_Gerar_BaixaParcRec(dValor, lNumIntDocTit, lNumFat, sCiaCart, iQuantParc, dtDataCanc)
'                     If lErro <> SUCESSO Then
'                         bErro = True
'                         Call CoInfo_Adiciona_Log("Não foi possível gerar uma baixa por perda para o Título " & CStr(lNumFat) & " decorrente de um cancelamento de pagamento em cartão.", objImpCoInfoLog, colImpCoInfoLog)
'                     Else
'                         Call CoInfo_Adiciona_Log("Foi gerada uma baixa por perda para o Título " & CStr(lNumFat) & " decorrente de um cancelamento de pagamento em cartão.", objImpCoInfoLog, colImpCoInfoLog, True)
'                     End If
'
'                End If
'
'            End If
'        End If
'
'        'Insere o Log de erros
'        lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
'        If lErro <> SUCESSO Then gError 182544
'
'        'Atualiza o titulo para informar que ele foi importado com sucesso
'        If iComErro = DESMARCADO Then
'            iTotalRegCertos = iTotalRegCertos + 1
'
'            lErro = Comando_Executar(alComando2(0), "UPDATE CoinfoFatura SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ?, ComErro = ?, ValorTaxaCartao = ?, ValorDEducoes = ?, ValorComissaoRetida = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", TIPO_TITULO_RECEBER_TRV, objTitRec.lNumIntDoc, gdtDataHoje, iComErro, dTaxa, dDeducoes, dComissao, lNumIntArq, lNumRegArqTexto)
'            If lErro <> SUCESSO Then gError 182528
'
'            If sSigla = TIPODOC_CARTAO_CRED_DEB Then
'                lErro = Comando_Executar(alComando2(3), "UPDATE CoinfoItemFatura SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ?, ComErro = ? WHERE codest = ? AND Numfat = ? AND CodEmp = ? AND Cartao = 1 AND QuantParc <> 0 AND CiaCart <> ''", TIPO_TITULO_RECEBER_TRV, objTitRec.lNumIntDoc, gdtDataHoje, iComErro, objImpCoinfo.sSiglaEmpresa, lNumFat, lcodemp)
'                If lErro <> SUCESSO Then gError 182528
'            End If
'        Else
'            iTotalRegErr = iTotalRegErr + 1
'            bTemErro = True
'
'            'Desfaz Transação para cancelar as coisas pela metade
'            Call Transacao_Rollback
'
'            'Abertura de transação só para marcar que houve erro
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182525
'
'            lErro = Comando_Executar(alComando2(1), "UPDATE CoinfoFatura SET ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iComErro, lNumIntArq, lNumRegArqTexto)
'            If lErro <> SUCESSO Then gError 182528
'
'            If sSigla = TIPODOC_CARTAO_CRED_DEB Then
'                lErro = Comando_Executar(alComando2(4), "UPDATE CoinfoItemFatura SET ComErro = ? WHERE codest = ? AND Numfat = ? AND CodEmp = ? AND Cartao = 1 AND QuantParc <> 0 AND CiaCart <> ''", iComErro, objImpCoinfo.sSiglaEmpresa, lNumFat, lcodemp)
'                If lErro <> SUCESSO Then gError 182528
'            End If
'        End If
'
'        'Confirma a transação
'        lErro = Transacao_Commit()
'        If lErro <> AD_SQL_SUCESSO Then gError 182529
'
'        lErro = Rotina_Atualizar_AtualizaTelaBatch
'        If lErro <> SUCESSO Then gError 182720
'
'        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
'        lErro = Comando_BuscarProximo(alComando(3))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182530
'
'    Loop
'
'    Coinfo_AtualizarFaturasRec = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarFaturasRec:
'
'    Coinfo_AtualizarFaturasRec = gErr
'
'    Select Case gErr
'
'        Case 182519, 182520, 182530
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
'
'        Case 182521, 182522
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", gErr)
'
'        Case 182523, 182524
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADOS", gErr)
'
'        Case 182525
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182526, 182527, 182533, 182534, 182543, 182544, 182720, 182743, 26380
'
'        Case 182528
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)
'
'        Case 182529
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 182532
'            Call Rotina_Erro(vbOKOnly, "ERRO_TITREC_BAIXADO_NAO_ALTERAVEL", gErr, TIPODOC_FATURA_SERVICO_CR, lNumFat)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182531)
'
'    End Select
'
'    'Desfaz Transação
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function
'
'Private Function Coinfo_AtualizarFaturasPag(ByVal objImpCoinfo As ClassImpCoinfo, ByVal lNumIntArq As Long, alComando() As Long, alComando2() As Long, bTemErro As Boolean, iTotalRegCertos As Integer, iTotalRegErr As Integer) As Long
'
'Dim lNumRegArqTexto As Long, lNumFat As Long, dtDataFat As Date, lcodemp As Long
'Dim dValor As Double, dtdataven As Date, iCancelado As Integer, dvaldcbrl As Double
'Dim dvalbru As Double, dtDataCanc As Date, iPago As Integer, dtDataPag As Date
'Dim lNumIntDocTit As Long, lFornecedorTit As Long, dtDataEmissaoTit As Date
'Dim iStatusTit As Integer, dValorTit As Double, dSaldoTit As Double
'Dim lErro As Long, iAchou As Integer, objTitPag As ClassTituloPagar
'Dim objParcelaPagar As ClassParcelaPagar
'Dim colParcelaPagar As colParcelaPagar
'Dim objContabil As ClassContabil
'Dim colBaixaParcPagar As ColBaixaParcRec
'Dim objBaixaPagar As ClassBaixaPagar
'Dim lrecno As Long
'Dim lTransacao As Long
'Dim objImpCoInfoLog As ClassImpCoinfoLogAtu
'Dim colImpCoInfoLog As Collection
'Dim sCritica As String, objCtb As ClassImpCoinfoNFPagCtb
'Dim iComErro As Integer, objContabAutomatica As ClassContabAutomatica
'Dim bErro As Boolean
'Dim iFilialCorporator As Integer
'Dim iFilialCoinfo As Integer
'Dim objNFPag As ClassNFsPag
'Dim iTabela As Integer
'Dim lNumIntTabela As Long
'Dim sTipoDoc As String
'Dim bReembolso As Boolean, objCtb2 As ClassImpCoinfoFatCPCtb
'Dim dComissao As Double
'
'On Error GoTo Erro_Coinfo_AtualizarFaturasPag
'
'    'Le as faturas a Pagar
'    lErro = Comando_ExecutarPos(alComando(7), "SELECT RecNo, NumRegArqTexto, numfat, datafat, codemp, valor, dataven, cancelado, valdcbrl, valbru, datacanc, pago, datapag FROM CoinfoFatura WHERE CoinfoFatura.NumIntArq = ? AND CoinfoFatura.cpcr = ? AND CoinfoFatura.codest = ? AND CoinfoFatura.NoCorporatorEm = ? AND CoinfoFatura.Usuario <> ? ORDER BY CoinfoFatura.NumRegArqTexto", 0, _
'        lrecno, lNumRegArqTexto, lNumFat, dtDataFat, lcodemp, dValor, dtdataven, iCancelado, dvaldcbrl, dvalbru, dtDataCanc, iPago, dtDataPag, lNumIntArq, 1, objImpCoinfo.sSiglaEmpresa, DATA_NULA, "teste")
'    If lErro <> AD_SQL_SUCESSO Then gError 182532
'
'    'Busca a primeira
'    lErro = Comando_BuscarProximo(alComando(7))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'    'Enquanto houver faturas não atualizadas
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        iAchou = 0
'        bErro = False
'
'        lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", lcodemp, iFilialCorporator, iFilialCoinfo)
'        If lErro <> SUCESSO Then gError 182743
'
'        Set colImpCoInfoLog = New Collection
'        Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
'
'        objImpCoInfoLog.iTabela = TABELA_COINFOFATURA
'        objImpCoInfoLog.lNumIntArq = lNumIntArq
'        objImpCoInfoLog.lNumRegArqTexto = lNumRegArqTexto
'
'        sTipoDoc = String(255, 0)
'
'        'Le as faturas a pagar
'        lErro = Comando_Executar(alComando(12), "SELECT TipoDoc  FROM CoinfoItemFatura WHERE codest = ? AND Numfat = ? AND CodEmp = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
'            sTipoDoc, objImpCoinfo.sSiglaEmpresa, lNumFat, lcodemp)
'        If lErro <> AD_SQL_SUCESSO Then gError 182519
'
'        'Busca a primeira
'        lErro = Comando_BuscarProximo(alComando(12))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182520
'
'        If lErro <> AD_SQL_SUCESSO Then
'            If iCancelado <> MARCADO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("O Título número " & CStr(lNumFat) & " não possui itens de fatura associados.", objImpCoInfoLog, colImpCoInfoLog)
'            End If
'        End If
'
'        'Se o item de fatura for do tipo NVL ou OCR é porque é um reembolso
'        'senão é uma comissão
'        If sTipoDoc = "NVL" Or sTipoDoc = "OCR" Then
'            bReembolso = True
'        Else
'            bReembolso = False
'        End If
'
'        'buscar a fatura como titulo a Pagar aberto
'        lErro = Comando_Executar(alComando(8), "SELECT NumIntDoc, Fornecedor, DataEmissao, Status, ValorTotal, Saldo FROM TitulosPag WHERE FilialEmpresa = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ?", _
'            lNumIntDocTit, lFornecedorTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, TIPODOC_FATURA_SERVICO_CP, lNumFat, STATUS_EXCLUIDO)
'        If lErro <> AD_SQL_SUCESSO Then gError 182534
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComando(8))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182535
'
'        If lErro <> AD_SQL_SUCESSO Then
'
'            'buscar a fatura como titulo a Pagar baixado
'            lErro = Comando_Executar(alComando(9), "SELECT NumIntDoc, Fornecedor, DataEmissao, Status, ValorTotal, Saldo FROM TitulosPagBaixados WHERE FilialEmpresa = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? ", _
'                lNumIntDocTit, lFornecedorTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, TIPODOC_FATURA_SERVICO_CP, lNumFat, STATUS_EXCLUIDO)
'            If lErro <> AD_SQL_SUCESSO Then gError 182536
'
'            'Busca o primeiro
'            lErro = Comando_BuscarProximo(alComando(9))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182537
'            If lErro = AD_SQL_SUCESSO Then
'                iAchou = 2
'            Else
'
'                'buscar a fatura como NFPagar aberto
'                lErro = Comando_Executar(alComando(10), "SELECT NumIntDoc, Fornecedor, DataEmissao, Status, ValorTotal, ValorTotal FROM NFsPag WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ?", _
'                    lNumIntDocTit, lFornecedorTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, lNumFat, STATUS_EXCLUIDO)
'                If lErro <> AD_SQL_SUCESSO Then gError 182534
'
'                'Busca o primeiro
'                lErro = Comando_BuscarProximo(alComando(10))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182535
'
'                If lErro <> AD_SQL_SUCESSO Then
'
'                    'buscar a fatura como titulo a Pagar baixado
'                    lErro = Comando_Executar(alComando(11), "SELECT NumIntDoc, Fornecedor, DataEmissao, Status, ValorTotal, ValorTotal FROM NFsPagBaixadas WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ? ", _
'                        lNumIntDocTit, lFornecedorTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, lNumFat, STATUS_EXCLUIDO)
'                    If lErro <> AD_SQL_SUCESSO Then gError 182536
'
'                    'Busca o primeiro
'                    lErro = Comando_BuscarProximo(alComando(11))
'                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182537
'                    If lErro = AD_SQL_SUCESSO Then
'                        iAchou = 4
'                    End If
'
'                Else
'                    iAchou = 3
'                End If
'
'            End If
'
'        Else
'            iAchou = 1
'        End If
'
'        'Título já pago (baixado) não pode ser alterado
'        'If iAchou = 2 Then gError 182532
'        If iAchou = 2 Then
'            'bErro = True
'            Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " já está baixado e não pode ser alterado.", objImpCoInfoLog, colImpCoInfoLog, True)
'        End If
'
'        'NFPag já pago (baixado) não pode ser alterado
'        'If iAchou = 2 Then gError 182532
'        If iAchou = 4 Then
'            'bErro = True
'            Call CoInfo_Adiciona_Log("A Nota Fiscal a Pagar Número " & CStr(lNumFat) & " já está baixado e não pode ser alterado.", objImpCoInfoLog, colImpCoInfoLog, True)
'        End If
'
'         'Abertura de transação
'        lTransacao = Transacao_Abrir()
'        If lTransacao = 0 Then gError 182538
'
''        'Exclui o Log Anterior para o Registro
''        lErro = CoInfoLogAtualizacao_Exclui(alComando2(5), alComando2(6), objImpCoInfoLog)
''        If lErro <> SUCESSO Then gError 182543
'
'        'preencher o objeto e mandar gravar
'        Set objTitPag = New ClassTituloPagar
'        Set objNFPag = New ClassNFsPag
'
'        objTitPag.lNumIntDoc = lNumIntDocTit
'        objNFPag.lNumIntDoc = lNumIntDocTit
'
'        'se não existir no sistema
'        If iAchou = 0 And iCancelado = DESMARCADO And Not bErro Then
'
'            'Se é um reembolso vira um título a pagar
'            'Comissão vira NFPag
'            If bReembolso Then
'
'                objTitPag.lFornecedor = lcodemp
'                objTitPag.iFilial = FILIAL_MATRIZ
'                objTitPag.iFilialPedCompra = FILIAL_MATRIZ
'                objTitPag.lNumTitulo = lNumFat
'                objTitPag.dtDataEmissao = dtDataFat
'                objTitPag.dValorTotal = dValor * -1
'                objTitPag.dValorProdutos = dValor * -1
'                objTitPag.dSaldo = dValor * -1
'                objTitPag.iFilialEmpresa = iFilialCorporator 'FILIAL_MATRIZ 'FIXO -> ATÉ SEGUNDA ORDEM
'                objTitPag.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CP
'                objTitPag.iNumParcelas = 1
'
'                Set objParcelaPagar = New ClassParcelaPagar
'                Set colParcelaPagar = New colParcelaPagar
'
'                objParcelaPagar.iNumParcela = 1
'                objParcelaPagar.dtDataVencimento = dtdataven
'                objParcelaPagar.dtDataVencimentoReal = dtdataven
'                objParcelaPagar.dValor = dValor * -1
'                objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
'                objParcelaPagar.iStatus = STATUS_ABERTO
'
'                With objParcelaPagar
'                    colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
'                End With
'
'                'Grava o Título a Pagar
'                lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
'                If lErro = SUCESSO Then
'
'                    If objTitPag.dtDataEmissao > #10/31/2006# Then
'
'                        Set objContabAutomatica = New ClassContabAutomatica
'                        Set objCtb2 = New ClassImpCoinfoFatCPCtb
'                        objCtb2.sModelo = "Reembolso"
'
'                        'Coloca a comissão no CSLL
'                        Call Coinfo_Obtem_Comissao_Cliente(objTitPag.lNumTitulo, objImpCoinfo.sSiglaEmpresa, dComissao)
'                        objTitPag.dCSLLRetido = dComissao
'
'                        'gravar a contabilizacao
'                        lErro = objContabAutomatica.Inicializa_Contab(objCtb2, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
'                            If lErro = SUCESSO Then
'
'                                lErro = objContabAutomatica.Finaliza_Contab()
'
'                            Else
'                                bErro = True
'                                Call CoInfo_Adiciona_Log("Erro na contabilização do Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'                            End If
'
'                        Else
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Erro na contabilização do Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                        End If
'
'                        Set objContabAutomatica = Nothing
'                        Call objCtb2.Encerrando
'                        Set objCtb2 = Nothing
'
'                    End If
'
'                End If
'
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " não pode ser gravado.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " foi gravado com sucesso.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            Else
'
'                objNFPag.lFornecedor = lcodemp
'                objNFPag.iFilial = FILIAL_MATRIZ
'                objNFPag.lNumNotaFiscal = lNumFat
'                objNFPag.dtDataEmissao = dtDataFat
'                objNFPag.dValorTotal = dValor * -1
'                objNFPag.dValorProdutos = dValor * -1
'                objNFPag.iFilialEmpresa = iFilialCorporator 'FILIAL_MATRIZ 'FIXO -> ATÉ SEGUNDA ORDEM
'                objNFPag.dtDataVencimento = dtdataven
'
'                'Grava a NF Pagar
'                lErro = CF("NFPag_Grava_EmTrans", objNFPag, objContabil)
'                If lErro = SUCESSO Then
'
'                    If objNFPag.dtDataEmissao > #10/31/2006# Then
'
'                        Set objContabAutomatica = New ClassContabAutomatica
'                        Set objCtb = New ClassImpCoinfoNFPagCtb
'
'                        'gravar a contabilizacao
'                        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objNFPag.dtDataEmissao, objNFPag.dtDataEmissao)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.GeraContabilizacao(objNFPag, False)
'                            If lErro = SUCESSO Then
'
'                                lErro = objContabAutomatica.Finaliza_Contab()
'
'                            Else
'                                bErro = True
'                                Call CoInfo_Adiciona_Log("Erro na contabilização da Nota Fiscal a Pagar Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                            End If
'
'                        Else
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Erro na contabilização da Nota Fiscal a Pagar Número " & CStr(lNumFat) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                        End If
'
'                        Set objContabAutomatica = Nothing
'                        Call objCtb.Encerrando
'                        Set objCtb = Nothing
'
'                    End If
'
'                End If
'
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal a Pagar Número " & CStr(lNumFat) & " não pode ser gravado.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal a Pagar Número " & CStr(lNumFat) & " foi gravado com sucesso.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            End If
'
'        End If
'
'        'Se o Título tiver sido cancelado e já existe sem estar baixado
'        'Excluir
'        If iCancelado = MARCADO And iAchou = 1 And Not bErro Then
'
'            Set objTitPag = New ClassTituloPagar
'
'            objTitPag.lNumIntDoc = lNumIntDocTit
'
'            lErro = CF("TituloPagar_Le", objTitPag)
'            'If lErro <> SUCESSO And lErro <> 18372 Then gError 182533
'            If lErro <> SUCESSO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("O Título Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " não pode ser lido.", objImpCoInfoLog, colImpCoInfoLog)
'            Else
'
'                Set objContabil = New ClassContabil
'                Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
'
'                lErro = CF("TituloPagar_Exclui_EmTransacao", objTitPag, objContabil)
'                'If lErro <> SUCESSO Then gError 182534
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " não pode ser excluído.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " foi excluido por estar com status de cancelado.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'                Set objContabil = Nothing
'            End If
'
'        End If
'
'        'Se a NF tiver sido cancelado e já existe sem estar baixado
'        'Excluir
'        If iCancelado = MARCADO And iAchou = 3 And Not bErro Then
'
'            Set objNFPag = New ClassNFsPag
'
'            objNFPag.lNumIntDoc = lNumIntDocTit
'
'            lErro = CF("NFPag_Le", objNFPag)
'            If lErro <> SUCESSO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("A Nota Fiscal a Pagar Número " & CStr(lNumFat) & " não pode ser lido.", objImpCoInfoLog, colImpCoInfoLog)
'            Else
'
'                Set objContabil = New ClassContabil
'                Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
'
'                lErro = CF("NFPag_Exclui_EmTransacao", objNFPag, objContabil)
'                'If lErro <> SUCESSO Then gError 182534
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal a Pagar Número " & CStr(lNumFat) & " não pode ser excluído.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal a Pagar Número " & CStr(lNumFat) & " foi excluido por estar com status de cancelado.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'                Set objContabil = Nothing
'            End If
'
'        End If
'
'        'Se é um título novo ou se ele não estiver baixado e for para baixar
'        If iPago = MARCADO And iCancelado = DESMARCADO And iAchou <> 2 And iAchou <> 4 Then
'
'            If objTitPag.dtDataEmissao > #10/31/2006# Then
'                Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_SERVICO_CP & " e Número " & CStr(lNumFat) & " possui data de emissão posterior a 01/11/2006, está baixado pelo Sigav e não está baixado pelo Corporator.", objImpCoInfoLog, colImpCoInfoLog, True)
'            End If
'            '++++++++++++++++ BAIXA
'
'        End If
'
'        'Insere o Log de erros
'        lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
'        If lErro <> SUCESSO Then gError 182544
'
'        If (iAchou = 0 And Not bReembolso) Or iAchou = 3 Or iAchou = 4 Then
'            iTabela = TIPO_NF_PAGAR_TRV
'            lNumIntTabela = objNFPag.lNumIntDoc
'        Else
'            iTabela = TIPO_TITULO_PAGAR_TRV
'            lNumIntTabela = objTitPag.lNumIntDoc
'        End If
'
'        'Atualiza a nota para informar que ela foi importada com sucesso
'        If iComErro = DESMARCADO Then
'            iTotalRegCertos = iTotalRegCertos + 1
'            lErro = Comando_Executar(alComando2(2), "UPDATE CoinfoFatura SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ?, ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iTabela, lNumIntTabela, gdtDataHoje, iComErro, lNumIntArq, lNumRegArqTexto)
'        Else
'            iTotalRegErr = iTotalRegErr + 1
'            bTemErro = True
'
'            'Desfaz Transação para canecalr as coisas pela metade
'            Call Transacao_Rollback
'
'            'Abertura de transação só para marcar que houve erro
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182538
'
'            lErro = Comando_Executar(alComando2(3), "UPDATE CoinfoFatura SET ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iComErro, lNumIntArq, lNumRegArqTexto)
'        End If
'        If lErro <> SUCESSO Then gError 182540
'
'        'Confirma a transação
'        lErro = Transacao_Commit()
'        If lErro <> AD_SQL_SUCESSO Then gError 182541
'
'        lErro = Rotina_Atualizar_AtualizaTelaBatch
'        If lErro <> SUCESSO Then gError 182720
'
'        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
'        lErro = Comando_BuscarProximo(alComando(7))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
'
'    Loop
'
'    Coinfo_AtualizarFaturasPag = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarFaturasPag:
'
'    Coinfo_AtualizarFaturasPag = gErr
'
'    Select Case gErr
'
'        Case 182532, 182533, 182542
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
'
'        Case 182534, 182535
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
'
'        Case 182536, 182537
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAG_BAIXADO", gErr)
'
'        Case 182538
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182539, 182543, 182544, 182720, 182743
'
'        Case 182540
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)
'
'        Case 182541
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
'
'    End Select
'
'    'Desfaz Transação
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function

Public Function Coinfo_ExcluiFaturasPag() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os TitulosPag com tipo FSCP

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iAchou As Integer
Dim lNumIntDocTit As Long
Dim objTitPag As ClassTituloPagar
Dim objCtb As ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil

On Error GoTo Erro_Coinfo_ExcluiFaturasPag

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    'buscar a fatura como titulo a Pagar aberto
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM TitulosPag WHERE SiglaDocumento = ? ", _
        lNumIntDocTit, TIPODOC_FATURA_SERVICO_CP)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS

         Set objTitPag = New ClassTituloPagar
         
         objTitPag.lNumIntDoc = lNumIntDocTit
         
         lErro = CF("TituloPagar_Le1", objTitPag)
         If lErro <> SUCESSO And lErro <> 18372 Then gError 182533
         If lErro = SUCESSO Then
    
             Set objContabil = New ClassContabil
             Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
             
             objTitPag.iExclusao = MARCADO
             
             lErro = CF("TituloPagar_Exclui_EmTransacao", objTitPag, objContabil)
             If lErro <> SUCESSO Then gError 182534
             Set objContabil = Nothing
         End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ExcluiFaturasPag = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ExcluiFaturasPag:

    Coinfo_ExcluiFaturasPag = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAG_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Coinfo_ExcluiCreditosPag() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os créditos a pagar do tipo NCP

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iAchou As Integer
Dim lNumIntDocTit As Long
Dim objCredPag As ClassCreditoPagar
Dim objCtb As ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil

On Error GoTo Erro_Coinfo_ExcluiCreditosPag

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    'buscar a fatura como titulo a Pagar aberto
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM CreditosPagForn WHERE SiglaDocumento = ? ", _
        lNumIntDocTit, "NCP")
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS

         Set objCredPag = New ClassCreditoPagar
         
         objCredPag.lNumIntDoc = lNumIntDocTit
         
         lErro = CF("CreditoPagar_Le", objCredPag, 1)
         If lErro <> SUCESSO And lErro <> 17071 Then gError 182533
         If lErro = SUCESSO Then
    
             Set objContabil = New ClassContabil
             Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
             
             objCredPag.iExclusao = MARCADO
             
             lErro = CF("CreditoPagar_Exclui_EmTransacao", objCredPag, objContabil)
             If lErro <> SUCESSO Then gError 182534
             Set objContabil = Nothing
         End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ExcluiCreditosPag = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ExcluiCreditosPag:

    Coinfo_ExcluiCreditosPag = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAG_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function
Public Function Coinfo_ExcluiFaturasOver() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os TitulosPag com tipo OVER

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iAchou As Integer
Dim lNumIntDocTit As Long
Dim objTitPag As ClassTituloPagar
Dim objCtb As ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil

On Error GoTo Erro_Coinfo_ExcluiFaturasOver

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    'buscar a fatura como titulo a Pagar aberto
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM TitulosPag WHERE SiglaDocumento = ? ", _
        lNumIntDocTit, TIPODOC_FATURA_OVER)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS

         Set objTitPag = New ClassTituloPagar
         
         objTitPag.lNumIntDoc = lNumIntDocTit
         
         lErro = CF("TituloPagar_Le1", objTitPag)
         If lErro <> SUCESSO And lErro <> 18372 Then gError 182533
         If lErro = SUCESSO Then
    
             Set objContabil = New ClassContabil
             Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
             
             objTitPag.iExclusao = MARCADO
             
             lErro = CF("TituloPagar_Exclui_EmTransacao", objTitPag, objContabil)
             If lErro <> SUCESSO Then gError 182534
             Set objContabil = Nothing
         End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ExcluiFaturasOver = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ExcluiFaturasOver:

    Coinfo_ExcluiFaturasOver = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAG_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Coinfo_ExcluiNFPagReembolso() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os NFsPag que foram importados da Coinfo oriundos de items de fatura do tipo
'NVL e OCR (Reembolso)

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iAchou As Integer
Dim lNumIntDocTit As Long
Dim objNFPag As ClassNFsPag
Dim objCtb As ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil

On Error GoTo Erro_Coinfo_ExcluiNFPagReembolso

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    'buscar a fatura como titulo a Pagar aberto
    lErro = Comando_Executar(alComando(0), "SELECT DISTINCT NF.NumIntDoc FROM NFsPag AS NF, CoinfoItemFatura AS I WHERE I.NumFat = NF.NumNotaFiscal AND I.Codemp = NF.Fornecedor AND I.TipoDoc IN ('NVL', 'OCR') ", _
        lNumIntDocTit)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS

         Set objNFPag = New ClassNFsPag
         
         objNFPag.lNumIntDoc = lNumIntDocTit
         
         lErro = CF("NFPag_Le", objNFPag)
         If lErro = SUCESSO Then
    
             Set objContabil = New ClassContabil
             Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
             
             objNFPag.iExclusao = MARCADO
             
             lErro = CF("NFPag_Exclui_EmTransacao", objNFPag, objContabil)
             If lErro <> SUCESSO Then gError 182534
             Set objContabil = Nothing
         End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ExcluiNFPagReembolso = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ExcluiNFPagReembolso:

    Coinfo_ExcluiNFPagReembolso = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAG_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Coinfo_ExcluiFaturasRec() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os TitulosRec sem data de vencimento

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iAchou As Integer
Dim lNumIntDocTit As Long
Dim objTitRec As ClassTituloReceber
Dim objCtb As ClassImpCoinfoFatCRCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil
Dim objFaturaTRV As New ClassFaturaTRV

On Error GoTo Erro_Coinfo_ExcluiFaturasRec

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    'buscar a fatura como titulo a receber sem data de vencimento
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM TitulosRec WHERE NumTitulo BETWEEN ? AND ? AND Status <> 5 ", lNumIntDocTit, 413540, 414021)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver títulos
    Do While lErro <> AD_SQL_SEM_DADOS

         Set objTitRec = New ClassTituloReceber
         Set objFaturaTRV = New ClassFaturaTRV
         
         objFaturaTRV.sMotivo = "Faturado voucher a voucher"
         
         objTitRec.lNumIntDoc = lNumIntDocTit
         Set objTitRec.objInfoUsu = objFaturaTRV
         
         lErro = CF("TituloReceber_Le", objTitRec)
         If lErro <> SUCESSO And lErro <> 26061 Then gError 182533
         If lErro = SUCESSO Then
    
             Set objContabil = New ClassContabil
             Call objContabil.Contabil_Inicializa_Contabilidade4(23, MODULO_CONTASARECEBER)
             
             objTitRec.iExclusao = MARCADO
             
             lErro = CF("TituloReceber_Exclui_EmTransacao", objTitRec, objContabil)
             If lErro <> SUCESSO Then gError 182534
             Set objContabil = Nothing
         End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ExcluiFaturasRec = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ExcluiFaturasRec:

    Coinfo_ExcluiFaturasRec = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_RECEBER", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function RelLcto_Exclui() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os TitulosRec sem data de vencimento

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim iAchou As Integer
Dim lNumIntRel As Long

On Error GoTo Erro_RelLcto_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 190060
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 190061

    'buscar a fatura como titulo a receber sem data de vencimento
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntRel FROM RelLcto ", 0, _
        lNumIntRel)
    If lErro <> AD_SQL_SUCESSO Then gError 190062
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190063
    
    'Enquanto houver títulos
    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RelLcto", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 190064
        
        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 190065
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 190066
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelLcto_Exclui = SUCESSO
     
    Exit Function
    
Erro_RelLcto_Exclui:

    RelLcto_Exclui = gErr
     
    Select Case gErr
    
        Case 190060
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 190061
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 190062, 190063.190065
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELLCTO", gErr)
        
        Case 190064
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_RELLCTO", gErr)

        Case 190066
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190067)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

'Public Function Coinfo_GerarComissaoRetroativo(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''Le os Títulos a receber de cartão de crédito que estão no BD e manda gerar a comissão
''referente a eles
'
'Dim lErro As Long
'Dim lTransacao As Long
'Dim iIndice As Integer
'Dim alComando(0 To 1) As Long
'Dim iAchou As Integer
'Dim lNumIntDocTit As Long
'Dim objTitRec As ClassTituloReceber
'Dim objCtb As ClassImpCoinfoFatCRCtb
'Dim objContabAutomatica As ClassContabAutomatica
'Dim objContabil As ClassContabil
'Dim objTitRecTRV As ClassTitulosRecTRV
'
'On Error GoTo Erro_Coinfo_GerarComissaoRetroativo
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 182534
'    Next
'
'    'Abertura de transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 182459
'
'    'Lê os titulos a receber de cartão
'    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM TitulosRecTodos WHERE SiglaDocumento = ? AND NumTitulo > ? ", _
'        lNumIntDocTit, TIPODOC_CARTAO_CRED_DEB, 3000)
'    If lErro <> AD_SQL_SUCESSO Then gError 182534
'
'    'Busca a primeira
'    lErro = Comando_BuscarProximo(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'    'Enquanto houver títulos
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        Set objTitRec = New ClassTituloReceber
'
'        objTitRec.lNumIntDoc = lNumIntDocTit
'
'        'Lê
'        lErro = CF("TituloReceber_Le", objTitRec)
'        If lErro <> SUCESSO And lErro <> 26061 Then gError 182533
'        If lErro = SUCESSO Then
'
'            Set objTitRecTRV = New ClassTitulosRecTRV
'
'            objTitRecTRV.lNumIntDocTitRec = objTitRec.lNumIntDoc
'
'            lErro = CF("TitulosRecTRV_Le", objTitRecTRV)
'            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182542
'
'            'Gera a comissão
'            lErro = Coinfo_GerarComissaoCartao(objImpCoinfo, objTitRec, objTitRecTRV)
'            If lErro <> SUCESSO Then gError 182542
'
'            lErro = CF("TitulosRecTRV_Grava", objTitRecTRV)
'            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182542
'
'        End If
'
'        'Lê o próximo
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 182541
'
'    'Fecha a transação
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Coinfo_GerarComissaoRetroativo = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_GerarComissaoRetroativo:
'
'    Coinfo_GerarComissaoRetroativo = gErr
'
'    Select Case gErr
'
'        Case 182532, 182533, 182542
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
'
'        Case 182534, 182535
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_RECEBER", gErr)
'
'        Case 182536, 182537
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADO", gErr)
'
'        Case 182538
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182539, 182543, 182544, 182720, 182743
'
'        Case 182540
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)
'
'        Case 182541
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function

'Private Function Coinfo_AtualizarFaturas(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''atualiza faturas no cp e no cr
'
'Dim lErro As Long, lTransacao As Long
'Dim alComando(0 To 14) As Long
'Dim alComando2(0 To 5) As Long
'Dim iIndice As Integer
'Dim lNumIntArq As Long
'Dim bTemErro As Boolean
'Dim iTotalRegCertosTR As Integer
'Dim iTotalRegErrTR As Integer
'Dim iTotalRegCertosTP As Integer
'Dim iTotalRegErrTP As Integer
'
'On Error GoTo Erro_Coinfo_AtualizarFaturas
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_AbrirExt(glConexao1)
'        If alComando(iIndice) = 0 Then gError 182508
'    Next
'
'    'Abre o comando
'    For iIndice = LBound(alComando2) To UBound(alComando2)
'        alComando2(iIndice) = Comando_Abrir()
'        If alComando2(iIndice) = 0 Then gError 182509
'    Next
'
'    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpcabecfat")
'    If lErro <> SUCESSO Then gError 182714
'
'    'Abertura de transação
'    lTransacao = Transacao_AbrirExt(glConexao1)
'    If lTransacao = 0 Then gError 182510
'
'    'Le os arquivos não atualizados
'    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
'        lNumIntArq, "fpcabecfat", DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 182511
'
'    'Busca o Primeiro
'    lErro = Comando_BuscarProximo(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182512
'
'    'Enquando existirem arquivos
'    Do While lErro = AD_SQL_SUCESSO
'
'        bTemErro = False
'
'        'gerar titulos a receber
'        lErro = Coinfo_AtualizarFaturasRec(objImpCoinfo, lNumIntArq, alComando, alComando2, bTemErro, iTotalRegCertosTR, iTotalRegErrTR)
'        If lErro <> SUCESSO Then gError 182513
'
'        'gerar titulos a pagar
'        lErro = Coinfo_AtualizarFaturasPag(objImpCoinfo, lNumIntArq, alComando, alComando2, bTemErro, iTotalRegCertosTP, iTotalRegErrTP)
'        If lErro <> SUCESSO Then gError 182514
'
'        'Não atualiza o Arquivo e não busca o próximo
'        If bTemErro Then Exit Do
'
'        'Atualiza o arquivo
'        lErro = Comando_ExecutarPos(alComando(1), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComando(0), gdtDataHoje)
'        If lErro <> SUCESSO Then gError 182515
'
'        'Busca o próximo
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182516
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacao)
'    If lErro <> AD_SQL_SUCESSO Then gError 182517
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Receber foram atualizados. " & CStr(iTotalRegCertosTR) & " registros com sucesso e " & CStr(iTotalRegErrTR) & " registros com erro." & vbNewLine
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Pagar foram atualizados. " & CStr(iTotalRegCertosTP) & " registros com sucesso e " & CStr(iTotalRegErrTP) & " registros com erro." & vbNewLine
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'Fecha o comando
'    For iIndice = LBound(alComando2) To UBound(alComando2)
'        Call Comando_Fechar(alComando2(iIndice))
'    Next
'
'    Coinfo_AtualizarFaturas = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarFaturas:
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Receber foram atualizados. " & CStr(iTotalRegCertosTR) & " registros com sucesso e " & CStr(iTotalRegErrTR) & " registros com erro." & vbNewLine
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Pagar foram atualizados. " & CStr(iTotalRegCertosTP) & " registros com sucesso e " & CStr(iTotalRegErrTP) & " registros com erro." & vbNewLine
'
'    Coinfo_AtualizarFaturas = gErr
'
'    Select Case gErr
'
'        Case 182508, 182509
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 182510
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182511, 182512, 182516
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 182513, 182514, 182714
'
'        Case 182515
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
'
'        Case 182517
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182518)
'
'    End Select
'
'    Call Transacao_RollbackExt(lTransacao)
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    For iIndice = LBound(alComando2) To UBound(alComando2)
'        Call Comando_Fechar(alComando2(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Private Function Filial_Coinfo_Retorna_Corporator(ByVal lFilialCoinfo As Long) As Integer
''Obtém a filial no corporator com base na filial do Sigav
'
''São Paulo - Matriz - Código 01
''Av. Ipiranga, 345 sl Centro - 01046-923
''07.139.957/0001-62
'
''Filial Recife - Código 9969
''Av. Engenheiro Domingos Ferreira, 2391 sala 703 - Boa Viagem 51020-031
''07.139.957/0002-43
'
''Filial Curitiba - Código 06
''Rua Tibagi, 294 - cjs. 801/802
''07.139.957/0003-24
'
''Filial Rio de Janeiro - Código 1117
''Avenida Rio Branco, 181 Sala 2005
''07.139.957/0004-05
'
''Filial Porto Alegre - Código 2842
''Av. Alberto Bins, 392 Sala 1203
''07.139.957/0005-96
'
''Filial Belo Horizonte - Código 36
''Rua Viçosa, 43 Loja 05 - São Pedro 30330-160
''07.139.957/0006-77
'
'    Select Case lFilialCoinfo
'
'        Case 1
'            Filial_Coinfo_Retorna_Corporator = 1
'        Case 9969
'            Filial_Coinfo_Retorna_Corporator = 2
'        Case 6
'            Filial_Coinfo_Retorna_Corporator = 3
'        Case 1117
'            Filial_Coinfo_Retorna_Corporator = 4
'        Case 2842
'            Filial_Coinfo_Retorna_Corporator = 5
'        Case 36
'            Filial_Coinfo_Retorna_Corporator = 6
'
'    End Select
'
'End Function
'
'Private Function Coinfo_AtualizarNFs(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''atualiza NFs
''percorrer os registros de fpnf e fpitemnf e marcá-los como lidos
''preenche objeto e grava
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim alComandoExt(0 To 4) As Long
'Dim alComando(1 To 2) As Long
'Dim lTransacao As Long
'Dim lTransacaoExt As Long
'Dim lNumIntArq As Long
'Dim objNF As ClassNFiscal
'Dim objItemNF As ClassItemNF
'Dim objProduto As New ClassProduto
'Dim objCliente As ClassCliente
'Dim objTribTab As ClassTribTab
'Dim objTipoDocInfo As New ClassTipoDocInfo
'Dim sSQL1 As String
'Dim sSQL2 As String
'Dim sSQL3 As String
'Dim sSQL4 As String
'Dim sSQL5 As String
'Dim lNumRegArqTexto As Long
'Dim lNumNF As Long
'Dim dtdataemi As Date
'Dim iCancelado As Integer
'Dim dValor As Double
'Dim sDescricao1 As String
'Dim sDescricao2 As String
'Dim sdescricao3 As String
'Dim lcodnat As Long
'Dim sdiscriminacao1 As String
'Dim sdiscriminacao2 As String
'Dim sdiscriminacao3 As String
'Dim sdiscriminacao4 As String
'Dim lcodemp As Long
'Dim idescacre1 As Integer
'Dim idescacre2 As Integer
'Dim idescacre3 As Integer
'Dim dvalcompl1 As Double
'Dim dvalcompl2 As Double
'Dim dvalcompl3 As Double
'Dim imanuauto As Integer
'Dim lcdfilial As Long
'Dim dValorItem As Double
'Dim lNumFat As Long
'Dim stexto1 As String
'Dim objContabil As New ClassContabil
'Dim colClientes As New Collection
'Dim bAchou As Boolean
'Dim iNumNotas As Integer
'Dim iContatorNotas As Integer
'Dim objSerie As ClassSerie
'Dim lNumIntDocNF As Long
'Dim dvalorAux As Double
'Dim bForcaItemNF As Boolean
'Dim objImpCoInfoLog As ClassImpCoinfoLogAtu
'Dim colImpCoInfoLog As Collection
'Dim sCritica As String
'Dim iComErro As Integer, objContabAutomatica As ClassContabAutomatica
'Dim bTemErro As Boolean, objCtb As ClassImpCoinfoNFCtb
'Dim bErro As Boolean
'Dim iTotalRegCertos As Integer
'Dim iTotalRegErr As Integer
'Dim iFilialEmpAux As Integer
'Dim dComissao As Double
'Dim dComissaoAcumulada As Double
'Dim sNomeArquivo As String
'Dim sNomeArquivoItem As String
'
'Const NATUREZAOP_TRV = "5949"
'Const TIPO_TRIBUTACAO_TRV = 3
'Const SERIE_TRV = "1"
'Const PRODUTO_TRV = "FAT       "
'Dim iFilialEmpresaAnt As Integer
'
'On Error GoTo Erro_Coinfo_AtualizarNFs
'
'    iFilialEmpresaAnt = giFilialEmpresa
'
'    'Obtém o total de registros que vão ser atualizados.
'    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpnf")
'    If lErro <> SUCESSO Then gError 182714
'
'    'Le o tipo de nota que vai ser usado
'    objTipoDocInfo.iCodigo = DOCINFO_NFISS
'
'    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 31415 Then gError 182456
'
'    'Le o produto que vai ser usado nos itens da Nota
'    objProduto.sCodigo = PRODUTO_TRV
'
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 182457
'
'    'Se o produto não existe => Erro
'    If lErro <> SUCESSO Then gError 182458
'
'    'Se for uma nota fiscal de serviço
'    If objTipoDocInfo.iCodigo = DOCINFO_NFISS Then
'        If Not (objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iFaturamento = PRODUTO_VENDAVEL And objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) Then gError 182489
'    End If
'
'    'SQL de leitura de CoinfoArqsLidos não atualizados
'    sSQL1 = "SELECT NumIntImportacao, IdentificacaoArq FROM CoinfoArqsLidos WHERE NomeArq = ? AND " & _
'            "DataAtualizado = ? ORDER BY IdentificacaoArq"
'
'    'SQL de leitura de CoinfoNF não atualizados
'    sSQL2 = "SELECT NumRegArqTexto,numnf,dataemi,cancelado,valor,descricao1, " & _
'            "descricao2,descricao3,codnat,discriminacao1,discriminacao2,discriminacao3, " & _
'            "discriminacao4,codemp,descacre1,descacre2,descacre3,valcompl1,valcompl2,valcompl3, " & _
'            "manuauto,cdfilial FROM CoinfoNF WHERE NumIntArq = ? AND NoCorporatorEm = ?  AND codest = ? " & _
'            "ORDER BY cdfilial, NumNF"
'
'    'SQL de leitura de CoInfoItemNF
'    sSQL3 = "SELECT valor, numfat, texto1 FROM CoInfoItemNF WHERE NumIntArq = ? AND CodEst = ? AND NumNF = ? AND cdFilial = ? ORDER BY Ordem"
'
'    'SQL de leitura de CoInfoItemNF
'    sSQL5 = "SELECT I.valor, I.numfat, I.texto1 FROM CoInfoItemNF AS I, CoInfoArqsLidos AS A WHERE A.NumIntImportacao = I.NumIntArq AND A.IdentificacaoArq = ? AND I.CodEst = ? AND I.NumNF = ? AND I.cdFilial = ? ORDER BY Ordem"
'
'    'SQL de verificação de existência da nota
'    sSQL4 = "SELECT NumIntDoc FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao > ? AND DataEmissao < ? AND Status <> ? AND (Tipo = ? OR Tipo = ?)"
'
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 182459
'
'    'abertura de comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
'        If alComandoExt(iIndice) = 0 Then gError 182460
'    Next
'
'    'abertura de comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 182460
'    Next
'
'    sNomeArquivo = String(250, 0)
'
'    'Le os Arquivos não atualizados
'    lErro = Comando_ExecutarPos(alComandoExt(0), sSQL1, 0, lNumIntArq, sNomeArquivo, "fpnf", DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 182461
'
'    'Busca o primeiro
'    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182462
'
'    iContatorNotas = 0
'
'    'Enquanto houver arquivos não atualizados
'    Do While lErro = AD_SQL_SUCESSO
'
'        bTemErro = False
'
'        'Aloca espaço no buffer
'        sDescricao1 = String(255, 0)
'        sDescricao2 = String(255, 0)
'        sdescricao3 = String(255, 0)
'
'        sdiscriminacao1 = String(255, 0)
'        sdiscriminacao2 = String(255, 0)
'        sdiscriminacao3 = String(255, 0)
'        sdiscriminacao4 = String(255, 0)
'
'        'Le as notas não atualizadas
'        lErro = Comando_ExecutarPos(alComandoExt(1), sSQL2, 0, lNumRegArqTexto, lNumNF, _
'        dtdataemi, iCancelado, dValor, sDescricao1, sDescricao2, sdescricao3, lcodnat, _
'        sdiscriminacao1, sdiscriminacao2, sdiscriminacao3, sdiscriminacao4, lcodemp, idescacre1, _
'        idescacre2, idescacre3, dvalcompl1, dvalcompl2, dvalcompl3, imanuauto, lcdfilial, _
'        lNumIntArq, DATA_NULA, objImpCoinfo.sSiglaEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then gError 182463
'
'        'Busca a primeira
'        lErro = Comando_BuscarPrimeiro(alComandoExt(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182464
'
'        iNumNotas = 0
'
'        'Enquanto houver notas não atualizadas
'        Do While lErro = AD_SQL_SUCESSO
'
'            bErro = False
'
'            iNumNotas = iNumNotas + 1
'            iContatorNotas = iContatorNotas + 1
'
'            Set colImpCoInfoLog = New Collection
'            Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
'
'            objImpCoInfoLog.iTabela = TABELA_COINFONF
'            objImpCoInfoLog.lNumIntArq = lNumIntArq
'            objImpCoInfoLog.lNumRegArqTexto = lNumRegArqTexto
'
'            Set objNF = New ClassNFiscal
'
'            objNF.iFilialEmpresa = Filial_Coinfo_Retorna_Corporator(lcdfilial)
'            objNF.sSerie = SERIE_TRV
'            objNF.lNumNotaFiscal = lNumNF
'            objNF.lCliente = lcodemp
'            objNF.iFilialCli = FILIAL_MATRIZ
'            objNF.iFilialEntrega = FILIAL_MATRIZ
'            objNF.dtDataEmissao = dtdataemi
'            objNF.iTipoNFiscal = objTipoDocInfo.iCodigo
'            objNF.iTipoDocInfo = objTipoDocInfo.iTipo
'            objNF.iStatus = STATUS_LANCADO
'
'            giFilialEmpresa = objNF.iFilialEmpresa
'
'            'Verifica se já existe nota com esses dados
'            lErro = Comando_Executar(alComandoExt(3), sSQL4, lNumIntDocNF, objNF.iFilialEmpresa, objNF.sSerie, objNF.lNumNotaFiscal, objNF.dtDataEmissao - PERIODO_EMISSAO, objNF.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
'            If lErro <> AD_SQL_SUCESSO Then gError 182490
'
'            'Busca a primeira
'            lErro = Comando_BuscarPrimeiro(alComandoExt(3))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182491
'
'            'Abertura de transação
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182459
'
''            'Exclui o Log Anterior para o Registro
''            lErro = CoInfoLogAtualizacao_Exclui(alComando(1), alComando(2), objImpCoInfoLog)
''            If lErro <> SUCESSO Then gError 182543
'
'            'Se Existe => Exclui para inserir de novo
'            If lErro = AD_SQL_SUCESSO Then
'
'                objNF.lNumIntDoc = lNumIntDocNF
'
'                iFilialEmpAux = objNF.iFilialEmpresa
'                lErro = CF("NFiscal_Le", objNF)
'                objNF.iFilialEmpresa = iFilialEmpAux
'                'If lErro <> SUCESSO And lErro <> 31442 Then gError 182495
'                If lErro <> SUCESSO Then
'                    If lErro = 31442 Then
'                        Call CoInfo_Adiciona_Log("A Nota Fiscal Número " & CStr(lNumNF) & " não existe no sistema.", objImpCoInfoLog, colImpCoInfoLog, True)
'                    Else
'                        bErro = True
'                        Call CoInfo_Adiciona_Log("A Nota Fiscal Número " & CStr(lNumNF) & " não pode ser lida.", objImpCoInfoLog, colImpCoInfoLog)
'                    End If
'                Else
'                    'Exclui a Nota
'                    Set objContabil = New ClassContabil
'
'                    Call objContabil.Contabil_Inicializa_Contabilidade4(4, MODULO_FATURAMENTO)
'                    lErro = CF("NotaFiscalSaida_Excluir_EmTrans", objNF, objContabil)
'                    'If lErro <> SUCESSO Then gError 182467
'                    If lErro <> SUCESSO Then
'                        bErro = True
'                        Call CoInfo_Adiciona_Log("A Nota Fiscal Número " & CStr(lNumNF) & " não pode ser excluida.", objImpCoInfoLog, colImpCoInfoLog)
'                    Else
'                        Call CoInfo_Adiciona_Log("A Nota Fiscal Número " & CStr(lNumNF) & " já existia no sistema, por isso foi excluída.", objImpCoInfoLog, colImpCoInfoLog, True)
'                    End If
'                    Set objContabil = Nothing
'                End If
'
'            End If
'
'            'Obtém dados fixos
'            objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
'            objNF.dtDataReferencia = DATA_NULA
'            objNF.iFreteRespons = FRETE_EMITENTE
'            objNF.dtDataRegistro = gdtDataAtual
'            objNF.dtDataEntrada = DATA_NULA
'            objNF.dtDataVencimento = DATA_NULA
'            objNF.sNaturezaOp = NATUREZAOP_TRV
'
'            'Obtém dados do Registro
'            objNF.dtDataSaida = dtdataemi
'            objNF.dValorTotal = dValor
'            objNF.dValorProdutos = dValor
'
'            objNF.sMensagemCorpoNota = left(sdiscriminacao1 & sdiscriminacao2 & sdiscriminacao3 & sdiscriminacao4, 250)
'
'            'Verifica a existência do cliente
'            bAchou = False
'            For Each objCliente In colClientes
'                If objCliente.lCodigo = objNF.lCliente Then
'                    bAchou = True
'                    Exit For
'                End If
'            Next
'
'            'Se não achou o cliente na coleção
'            If Not bAchou Then
'
'                Set objCliente = New ClassCliente
'
'                objCliente.lCodigo = objNF.lCliente
'
'                'Le o cliente
'                lErro = CF("Cliente_Le", objCliente)
'                'If lErro <> SUCESSO And lErro <> 12293 Then gError 182468
'                If lErro <> SUCESSO And lErro <> 12293 Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Cliente com código " & CStr(objNF.lCliente) & " não pode ser lido.", objImpCoInfoLog, colImpCoInfoLog)
'                End If
'
'                'Se não existe cliente = > Erro
'                'If lErro <> SUCESSO Then gError 182469
'                If lErro = 12293 Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Cliente com código " & CStr(objNF.lCliente) & " não está cadastrado no sistema.", objImpCoInfoLog, colImpCoInfoLog)
'                End If
'
'            End If
'
'            'Inicializa a Parte tributária da Nota
'            Call objNF.Inicializa_Tributacao
'
'            objNF.objTributacaoNF.iTipoTributacao = TIPO_TRIBUTACAO_TRV
'            objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
'            objNF.objTributacaoNF.sNaturezaOpInterna = NATUREZAOP_TRV
'            objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
'
'            With objNF.objTributacaoNF
'
''                .objTributacaoDesconto.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoDesconto.iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoFrete.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoFrete.iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoOutras.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoOutras.iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoSeguro.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
''                .objTributacaoSeguro.iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
'
'                .iICMSBaseManual = VAR_PREENCH_AUTOMATICO
'                .iICMSSubstBaseManual = VAR_PREENCH_AUTOMATICO
'                .iICMSSubstValorManual = VAR_PREENCH_AUTOMATICO
'                .iICMSValorManual = VAR_PREENCH_AUTOMATICO
'                .iIPIBaseManual = VAR_PREENCH_AUTOMATICO
'                .iIPIValorManual = VAR_PREENCH_AUTOMATICO
'
'                .iIRRFAliquotaManual = VAR_PREENCH_AUTOMATICO
'                .iIRRFValorManual = VAR_PREENCH_AUTOMATICO
'                .iISSAliquotaManual = VAR_PREENCH_AUTOMATICO
'                .iISSValorManual = VAR_PREENCH_AUTOMATICO
'
'                .iINSSBaseManual = VAR_PREENCH_AUTOMATICO
'                .iINSSDeducoesManual = VAR_PREENCH_AUTOMATICO
'                .iINSSValorManual = VAR_PREENCH_AUTOMATICO
'                .iINSSRetidoManual = VAR_PREENCH_AUTOMATICO
'                .iISSRetidoManual = VAR_PREENCH_AUTOMATICO
'
'                .iPISRetidoManual = VAR_PREENCH_AUTOMATICO
'                .iCOFINSRetidoManual = VAR_PREENCH_AUTOMATICO
'                .iCSLLRetidoManual = VAR_PREENCH_AUTOMATICO
'
'            End With
'
'            'Inicializa o rastreamento
'            Set objNF.objRastreamento = New ClassRastreamento
'
'            'Aloca espaço no buffer
'            stexto1 = String(255, 0)
'
'            'Le os itens da nota
'            lErro = Comando_Executar(alComandoExt(2), sSQL3, dValorItem, lNumFat, stexto1, lNumIntArq, objImpCoinfo.sSiglaEmpresa, objNF.lNumNotaFiscal, lcdfilial)
'            If lErro <> AD_SQL_SUCESSO Then gError 182470
'
'            'Busca o primeiro
'            lErro = Comando_BuscarPrimeiro(alComandoExt(2))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182471
'
'            'Se não tem itens
'            If lErro = AD_SQL_SEM_DADOS Then
'
'                sNomeArquivoItem = "itemnf" & Mid(sNomeArquivo, 3, Len(sNomeArquivo) - 2)
'
'                'Le os itens da nota
'                lErro = Comando_Executar(alComandoExt(4), sSQL5, dValorItem, lNumFat, stexto1, sNomeArquivoItem, objImpCoinfo.sSiglaEmpresa, objNF.lNumNotaFiscal, lcdfilial)
'                If lErro <> AD_SQL_SUCESSO Then gError 182470
'
'                'Busca o primeiro
'                lErro = Comando_BuscarPrimeiro(alComandoExt(4))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182471
'
'                If lErro = AD_SQL_SEM_DADOS Then
'
'                    'Força ter um
'                    stexto1 = objNF.sMensagemCorpoNota
'                    dValorItem = objNF.dValorTotal
'                    If Len(Trim(stexto1)) = 0 Then stexto1 = objProduto.sDescricao
'                    bForcaItemNF = True
'                Else
'
'                    alComandoExt(2) = alComandoExt(4)
'
'                    bForcaItemNF = False
'                End If
'            Else
'                bForcaItemNF = False
'            End If
'
'            iIndice = 0
'
'            dComissaoAcumulada = 0
'
'            'Enquanto houver itens
'            Do While (lErro = AD_SQL_SUCESSO Or bForcaItemNF) And Not bErro
'
'                If lNumFat <> 0 Then
'
'                    lErro = Coinfo_Obtem_Comissao_Cliente(lNumFat, objImpCoinfo.sSiglaEmpresa, dComissao, True)
'                    If lErro <> SUCESSO Then gError 182473
'
'                    dComissaoAcumulada = dComissaoAcumulada + dComissao
'
'                End If
'
'                iIndice = iIndice + 1
'
'                Set objItemNF = New ClassItemNF
'
'                'Obtém dados fixos
'                objItemNF.dtDataEntrega = DATA_NULA
'
'                objItemNF.sProduto = objProduto.sCodigo
'                objItemNF.sDescricaoItem = stexto1
'                objItemNF.dQuantidade = 1
'                objItemNF.dPrecoUnitario = dValorItem
'
'                'Verifica se o valor do item pode ser 0
'                If objTipoDocInfo.iCodigo <> DOCINFO_NFISICM And objTipoDocInfo.iCodigo <> DOCINFO_NFISIPI Then
'                    If dValorItem = 0 Then
'                        gError 182484
'                    End If
'                End If
'
'                objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
'                objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
'                objItemNF.sUMVenda = objProduto.sSiglaUMVenda
'                objItemNF.iClasseUM = objProduto.iClasseUM
'
'                objItemNF.iItem = iIndice
'
'                objNF.ColItensNF.Add1 objItemNF
'
'                'Inicializa parte tributária do item da Nota
'                Call objItemNF.Inicializa_Tributacao
'
'                objItemNF.objTributacaoItemNF.iTipoTributacao = TIPO_TRIBUTACAO_TRV
'                objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'                objItemNF.objTributacaoItemNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
'                objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
'
'                With objItemNF.objTributacaoItemNF
'                    .iIPITipoManual = VAR_PREENCH_AUTOMATICO
'                    .iIPIBaseManual = VAR_PREENCH_AUTOMATICO
'                    .iIPIPercRedBaseManual = VAR_PREENCH_AUTOMATICO
'                    .iIPIAliquotaManual = VAR_PREENCH_AUTOMATICO
'                    .iIPIValorManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSTipoManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSBaseManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSPercRedBaseManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSAliquotaManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSValorManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSSubstBaseManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSSubstAliquotaManual = VAR_PREENCH_AUTOMATICO
'                    .iICMSSubstValorManual = VAR_PREENCH_AUTOMATICO
'                End With
'
'                'Se foi forçado a ter um item => Sai do Loop
'                If bForcaItemNF Then Exit Do
'
'                lErro = Comando_BuscarProximo(alComandoExt(2))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182472
'
'            Loop
'
'            Set objTribTab = New ClassTribTab
'
'            lErro = objTribTab.TributacaoNF_Reset(objNF)
'            If lErro <> SUCESSO Then gError 182473
'
'            'Atualiza a Tributação
'            Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
'            Call objTribTab.DataEmissao_Alterada(objNF.dtDataEmissao)
'            Call objTribTab.Serie_Alterada(objNF.sSerie)
'
'            objTribTab.objNFTributacao.iISSInclusoManual = VAR_PREENCH_AUTOMATICO
'
'            lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
'            If lErro <> SUCESSO Then gError 182474
'
'            lErro = objTribTab.PreencherNF(objNF)
'            If lErro <> SUCESSO Then gError 182475
'
'            Set objTribTab = Nothing
'
'            objNF.sNaturezaOp = objNF.objTributacaoNF.sNaturezaOpInterna
'
'            'acertar valor total considerando os tributos, retenções,...
'            objNF.dValorTotal = objNF.dValorTotal + objNF.objTributacaoNF.dICMSSubstValor + objNF.objTributacaoNF.dIPIValor + objNF.dValorSeguro + objNF.dValorOutrasDespesas + IIf(objNF.objTributacaoNF.iISSIncluso = 0, objNF.objTributacaoNF.dISSValor, 0)
'
'            'Verifica se é obrigatório ter itens
'            If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
'                'If objNF.ColItensNF.Count = 0 Then gError 182495
'                If objNF.ColItensNF.Count = 0 Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("A Nota " & CStr(objNF.lNumNotaFiscal) & " não tem itens cadastrados.", objImpCoInfoLog, colImpCoInfoLog)
'                End If
'            End If
'
'            'Obtém o valor total dos itens
'            dvalorAux = 0
'            For Each objItemNF In objNF.ColItensNF
'                dvalorAux = dvalorAux + objItemNF.dValorTotal
'            Next
'
'            'Verifica se o valor total dos itens e igual ao valor total da nota e o valor
'            'da nota original
'            'If Abs(dValorAux - dValor) > DELTA_VALORMONETARIO Or Abs(objNF.dValorTotal - dValor) > DELTA_VALORMONETARIO Then gError 182500
'
'            If Abs(dvalorAux - dValor) > DELTA_VALORMONETARIO Or Abs(objNF.dValorTotal - dValor) > DELTA_VALORMONETARIO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("O Valor da Nota " & CStr(objNF.lNumNotaFiscal) & " gerada está diferente das informações passadas. Valor Da Nota Gerada: " & Format(objNF.dValorTotal, "STANDARD") & ", Valor da Nota Importada: " & Format(dValor, "STANDARD") & " e Valor total dos itens: " & Format(dvalorAux, "STANDARD") & ".", objImpCoInfoLog, colImpCoInfoLog)
'            End If
'
'            Set objSerie = New ClassSerie
'
'            objSerie.iFilialEmpresa = objNF.iFilialEmpresa
'            objSerie.sSerie = objNF.sSerie
'
'            'Lê a Serie no BD
'            lErro = CF("Serie_Le", objSerie)
'            'If lErro <> SUCESSO And lErro <> 22202 Then gError 182485
'            If lErro <> SUCESSO Then
'                Call CoInfo_Adiciona_Log("A Série " & objNF.sSerie & " não está cadastrada no sistema.", objImpCoInfoLog, colImpCoInfoLog, True)
'            End If
'
'            'Altera o próximo número da Nota
'            objSerie.lProxNumNFiscal = objNF.lNumNotaFiscal
'
'            If Not bErro Then
'
'                'Grava a Serie no BD
'                lErro = CF("Serie_Grava_Trans", objSerie, False)
'                'If lErro <> SUCESSO Then gError 182487
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("A Série " & objNF.sSerie & " não pode ser gravada.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("A Série " & objNF.sSerie & " teve seu próximo número de nota alterado para " & objSerie.lProxNumNFiscal & ".", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            End If
'
'            If Not bErro Then
'
'                'Tenta Gravar a Nota
'                objNF.lNumNotaFiscal = 0
'                lErro = CF("NFiscal_Grava_EmTrans", objNF, Nothing)
'                'If lErro <> SUCESSO Then gError 182483
'                If lErro = SUCESSO Then
'
'                    If objNF.dtDataEmissao > #10/31/2006# Then
'
'                        Set objContabAutomatica = New ClassContabAutomatica
'                        Set objCtb = New ClassImpCoinfoNFCtb
'
'                        objNF.dTaxaMoeda = dComissaoAcumulada
'
'                        'gravar a contabilizacao
'                        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_FATURAMENTO, "FAT", objNF.dtDataEmissao, objNF.dtDataEmissao)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.GeraContabilizacao(objNF)
'                            If lErro = SUCESSO Then
'
'                                lErro = objContabAutomatica.Finaliza_Contab()
'
'                            Else
'                                bErro = True
'                                Call CoInfo_Adiciona_Log("Erro na contabilização da Nota Fiscal " & CStr(objNF.lNumNotaFiscal) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                            End If
'
'                        Else
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Erro na contabilização da Nota Fiscal " & CStr(objNF.lNumNotaFiscal) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                        End If
'
'                        Set objContabAutomatica = Nothing
'                        Call objCtb.Encerrando
'                        Set objCtb = Nothing
'
'                    End If
'
'                End If
'
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal " & CStr(objNF.lNumNotaFiscal) & " não pode ser gravada.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal " & CStr(objNF.lNumNotaFiscal) & " foi gravada com sucesso.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            End If
'
'            'Se a nota estiver cancelada
'            If iCancelado = MARCADO And Not bErro Then
'
'                'Cancela a Nota lançada
'                lErro = CF("NotaFiscalSaida_Cancelar_EmTrans", objNF, objNF.dtDataEmissao)
'                'If lErro <> SUCESSO Then gError 182493
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal " & CStr(objNF.lNumNotaFiscal) & " não pode ser cancelada.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("A Nota Fiscal " & CStr(objNF.lNumNotaFiscal) & " foi cancelada com sucesso.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            End If
'
'            'Insere o Log de erros
'            lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
'            If lErro <> SUCESSO Then gError 182544
'
'            If iComErro = DESMARCADO Then
'                iTotalRegCertos = iTotalRegCertos + 1
'                'Atualiza a nota para informar que ela foi importada com sucesso
'                lErro = Comando_Executar(alComando(1), "UPDATE CoinfoNF SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", TIPO_NOTA_FISCAL_TRV, objNF.lNumIntDoc, gdtDataHoje, lNumIntArq, lNumRegArqTexto)
'            Else
'                iTotalRegErr = iTotalRegErr + 1
'                bTemErro = True
'
'                'Desfaz Transação para canecalr as coisas pela metade
'                Call Transacao_Rollback
'
'                'Abertura de transação só para marcar que houve erro
'                lTransacao = Transacao_Abrir()
'                If lTransacao = 0 Then gError 182459
'
'                lErro = Comando_Executar(alComando(2), "UPDATE CoinfoNF SET ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iComErro, lNumIntArq, lNumRegArqTexto)
'            End If
'            If lErro <> SUCESSO Then gError 182476
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 182481
'
'            lErro = Rotina_Atualizar_AtualizaTelaBatch
'            If lErro <> SUCESSO Then gError 182720
'
'            'Busca a próxima nota não atualizada
'            lErro = Comando_BuscarProximo(alComandoExt(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182478
'
'        Loop
'
'        'Não atualiza o Arquivo e não busca o próximo
'        If bTemErro Then Exit Do
'
'        'Se todos as NFs foram atualizadas, então marca a data no arquivo
'        lErro = Comando_ExecutarPos(alComandoExt(1), "UPDATE CoinfoArqsLidos SET DataAtualizado = ? ", alComandoExt(0), gdtDataHoje)
'        If lErro <> SUCESSO Then gError 182479
'
'        'Busca o próximo arquivo não atualizado
'        lErro = Comando_BuscarProximo(alComandoExt(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182480
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacaoExt)
'    If lErro <> AD_SQL_SUCESSO Then gError 182481
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Notas Fiscais foram atualizadas. " & CStr(iTotalRegCertos) & " registros com sucesso e " & CStr(iTotalRegErr) & " registros com erro." & vbNewLine
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'Fecha o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    giFilialEmpresa = iFilialEmpresaAnt
'
'    Coinfo_AtualizarNFs = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarNFs:
'
'    giFilialEmpresa = iFilialEmpresaAnt
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Notas Fiscais foram atualizadas. " & CStr(iTotalRegCertos) & " registros com sucesso e " & CStr(iTotalRegErr) & " registros com erro." & vbNewLine
'
'    Coinfo_AtualizarNFs = gErr
'
'    Select Case gErr
'
'        Case 182456, 182457, 182466, 182467, 182468, 182473, 182474, 182475, 182483, 182485, 182486, 182487, 182493, 182543, 182544, 182714, 182720
'
'        Case 182458
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
'
'        Case 182459, 182465
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182460
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 182461, 182462, 182480
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 182463, 182464, 182478
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFONF", gErr)
'
'        Case 182469
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
'
'        Case 182470, 182471, 182472
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMNF", gErr)
'
'        Case 182476
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFONF", gErr)
'
'        Case 182477, 182481
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 182479
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
'
'        Case 182495
'            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF2", gErr, objNF.lNumNotaFiscal)
'
'        Case 182484
'              Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO2", gErr, objItemNF.iItem, objNF.lNumNotaFiscal)
'
'        Case 182489
'             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOTA_SERVICO", gErr, objProduto.sCodigo)
'
'        Case 182490, 182491
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
'
'        Case 182500
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALORES_IMPORT_NF_DIFERENTES", gErr, objNF.dValorTotal, dValor, dvalorAux)
'
'        Case 182501 'Tratado na Rotina Chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182482)
'
'    End Select
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'Fecha o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    'Desfaz Transação
'    Call Transacao_Rollback
'    Call Transacao_RollbackExt(lTransacaoExt)
'
'    Exit Function
'
'End Function

Private Function Coinfo_ImportarArqTexto0(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar tipos de empresa

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objGrupo As ClassImpCoinfoGrupo, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto0

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130580
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_GRUPO", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpgrupo", objImpCoinfo.sArqGrupoHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqGrupoHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqGrupoHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqGrupoHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqGrupo = lNumIntArqAnt
        gError 130593
    End If
    
    objImpCoinfo.lNumIntArqGrupo = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqGrupoHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
        Set objGrupo = New ClassImpCoinfoGrupo
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objGrupo.sCodigo = StrParaStrCoinfo(sCampoValor)
                        Case 2
                            objGrupo.sNome = StrParaStrCoinfo(sCampoValor)
                        Case 3
                            objGrupo.iMarca = IIf(sCampoValor = "yes", 1, 0)
                        Case 4
                            objGrupo.iFilial = StrParaInt(sCampoValor)
                        Case 5
                            objGrupo.iCodpes = StrParaInt(sCampoValor)
                        Case 6
                            objGrupo.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 7 Then gError 130592
        objGrupo.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objGrupo
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoGrupo (NumIntArq, NumRegArqTexto, Codigo, Nome, Marca, Filial, Codpes) VALUES (?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .sCodigo, .sNome, .iMarca, .iFilial, .iCodpes)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130583
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqGrupoHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto0 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto0:

    Coinfo_ImportarArqTexto0 = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM, 53
               
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 130592
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqGrupoHora, CStr(lNumRegs))
          
        Case ERRO_SEM_MENSAGEM, 53
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOGRUPO", gErr, objGrupo.sCodigo, objGrupo.sNome)
       
        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 130593
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqGrupoHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqGrupoHora)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select

    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqGrupoHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
    
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto1(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar tipos de empresa

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objTipoEmp As ClassImpCoinfoTipoEmp, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto1

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130580
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_TIPO_EMPR", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fptipoempr", objImpCoinfo.sArqTipoEmprHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqTipoEmprHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqTipoEmprHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqTipoEmprHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqTipoEmpr = lNumIntArqAnt
        gError 130593
    End If
    
    objImpCoinfo.lNumIntArqTipoEmpr = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqTipoEmprHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1
        
        lNumRegs = lNumRegs + 1
        
        Set objTipoEmp = New ClassImpCoinfoTipoEmp
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objTipoEmp.itipoempr = StrParaInt(sCampoValor)
                        Case 2
                            objTipoEmp.sdesctpem = left(StrParaStrCoinfo(sCampoValor), STRING_TIPO_CLIENTE_DESCRICAO)
                        Case 3
                            objTipoEmp.dtdatareg = StrParaDateCoinfo(sCampoValor)
                        Case 4
                            objTipoEmp.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 5
                            objTipoEmp.icliente = IIf(sCampoValor = "yes", 1, 0)
                        Case 6
                            objTipoEmp.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 7 Then gError 130592
        objTipoEmp.dtDataExp = StrParaDateCoinfo(sCampoValor)

                
        'inserir registro
        With objTipoEmp
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoTipoEmpr (NumIntArq, NumRegArqTexto, tipoempr, desctpem, datareg, Usuario, cliente) VALUES (?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .itipoempr, .sdesctpem, .dtdatareg, .sUsuario, .icliente)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130583
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
       
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqTipoEmprHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
   
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto1 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto1:

    Coinfo_ImportarArqTexto1 = gErr
     
    Select Case gErr
          
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 130592
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqTipoEmprHora, CStr(lNumRegs))
          
        Case ERRO_SEM_MENSAGEM, 53
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOTIPOEMPR", gErr, objTipoEmp.itipoempr, objTipoEmp.sdesctpem)
       
        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130593
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqTipoEmprHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqTipoEmprHora)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
     
    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqTipoEmprHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto2(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar empresas

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objEmp As ClassImpCoinfoEmp, sRegContinuacao As String
Dim objFSO As New FileSystemObject
Dim objTS As TextStream, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto2

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130595
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130596
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_EMPR", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpempresa", objImpCoinfo.sArqEmprHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqEmprHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqEmprHora = "" Then
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqEmprHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqEmpr = lNumIntArqAnt
        gError 130604
    End If
    
    objImpCoinfo.lNumIntArqEmpr = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqEmprHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
    
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
aqui:
        
        Set objEmp = New ClassImpCoinfoEmp
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objEmp.lcodemp = StrParaLong(sCampoValor)
                        Case 2
                            objEmp.snchave = left(StrParaStrCoinfo(sCampoValor), 50)
                        Case 3
                            objEmp.sNome = left(StrParaStrCoinfo(sCampoValor), 50)
                        Case 4
                            objEmp.sEndereco = left(StrParaStrCoinfo(sCampoValor), 65)
                        Case 5
                            objEmp.sBairro = left(StrParaStrCoinfo(sCampoValor), 40)
                        Case 6
                            objEmp.sCidade = left(StrParaStrCoinfo(sCampoValor), 30)
                        Case 7
                            objEmp.sUF = StrParaStrCoinfo(sCampoValor)
                            If Len(Trim(objEmp.sUF)) <> 2 Then gError 208886
                        Case 8
                            objEmp.sCEP = left(Replace(Replace(sCampoValor, "-", ""), ".", ""), 9)
                        Case 9
                            objEmp.sfone1 = left(StrParaStrCoinfo(sCampoValor), 30)
                        Case 10
                            objEmp.sfone2 = left(StrParaStrCoinfo(sCampoValor), 30)
                        Case 11
                            objEmp.sFax = left(StrParaStrCoinfo(sCampoValor), 30)
                        Case 12
                            objEmp.sCgc = left(StrParaStrCoinfo(sCampoValor), 14)
                        Case 13
                            objEmp.sie = left(StrParaStrCoinfo(sCampoValor), 18)
                        Case 14
                            objEmp.iNatureza = IIf(sCampoValor = "yes", 1, 0)
                        Case 15
                            objEmp.sContato = StrParaStrCoinfo(sCampoValor)
                        Case 16
                            objEmp.sCargo = left(StrParaStrCoinfo(sCampoValor), 20)
                        Case 17
                            objEmp.icob = IIf(sCampoValor = "yes", 1, 0)
                        Case 18
                            objEmp.spraca = left(StrParaStrCoinfo(sCampoValor), 30)
                        Case 19
                            objEmp.sendercob = left(StrParaStrCoinfo(sCampoValor), 65)
                        Case 20
                            objEmp.dlimcred = StrParaDblCoinfo(sCampoValor)
                        Case 21
                            objEmp.dtDataCad = StrParaDateCoinfo(sCampoValor)
                        Case 22
                            objEmp.scidcob = left(StrParaStrCoinfo(sCampoValor), 30)
                        Case 23
                            objEmp.sufcob = StrParaStrCoinfo(sCampoValor)
                        Case 24
                            objEmp.dtDataAlt = StrParaDateCoinfo(sCampoValor)
                        Case 25
                            objEmp.iCodpes = StrParaInt(sCampoValor)
                        Case 26
                            objEmp.dDesconto = StrParaDblCoinfo(sCampoValor)
                        Case 27
                            objEmp.sobs1 = StrParaStrCoinfo(sCampoValor)
                        Case 28
                            objEmp.sobs2 = StrParaStrCoinfo(sCampoValor)
                        Case 29
                            objEmp.sobs3 = StrParaStrCoinfo(sCampoValor)
                        Case 30
                            objEmp.sobs4 = StrParaStrCoinfo(sCampoValor)
                        Case 31
                            objEmp.sobs5 = StrParaStrCoinfo(sCampoValor)
                        Case 32
                            objEmp.sobs6 = StrParaStrCoinfo(sCampoValor)
                        Case 33
                            objEmp.lcodred = StrParaLong(sCampoValor)
                        Case 34
                            objEmp.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 35
                            objEmp.itipoempr = StrParaInt(sCampoValor)
                        Case 36
                            objEmp.dtdatareg = StrParaDateCoinfo(sCampoValor)
                        Case 37
                            objEmp.ifaturado = IIf(sCampoValor = "yes", 1, 0)
                        Case 38
                            objEmp.icondpag = StrParaInt(sCampoValor)
                        Case 39
                            objEmp.sCPF = left(StrParaStrCoinfo(sCampoValor), 14)
                        Case 40
                            objEmp.sEmail = left(StrParaStrCoinfo(sCampoValor), 100)
                        Case 41
                            objEmp.smoecred = left(StrParaStrCoinfo(sCampoValor), 3)
                        Case 42
                            objEmp.icodprom = StrParaInt(sCampoValor)
                        Case 43
                            objEmp.sbairrcob = left(StrParaStrCoinfo(sCampoValor), 44)
                        Case 44
                            objEmp.scepcob = left(Replace(Replace(sCampoValor, "-", ""), ".", ""), 9)
                        Case 45
                            objEmp.iAtivo = IIf(sCampoValor = "yes", 1, 0)
                        Case 46
                            objEmp.dtexpirlim = StrParaDateCoinfo(sCampoValor)
                        Case 47
                            objEmp.ifixapraz = IIf(sCampoValor = "yes", 1, 0)
                        Case 48
                            objEmp.dredcomis = StrParaDblCoinfo(sCampoValor)
                        Case 49
                            objEmp.icliente = IIf(sCampoValor = "yes", 1, 0)
                        Case 50
                            objEmp.icomissiona = IIf(sCampoValor = "yes", 1, 0)
                        Case 51
                            objEmp.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case 52
                            objEmp.inumban = StrParaInt(sCampoValor)
                        Case 53
                            objEmp.dISS = StrParaDblCoinfo(sCampoValor)
                        Case 54
                            objEmp.lagenban = StrParaLong(sCampoValor)
                        Case 55
                            objEmp.lnconta = StrParaLong(sCampoValor)
                        Case 56
                            objEmp.stextcon1 = StrParaStrCoinfo(sCampoValor)
                        Case 57
                            objEmp.stextcon2 = StrParaStrCoinfo(sCampoValor)
                        Case 58
                            objEmp.stextcon3 = StrParaStrCoinfo(sCampoValor)
                        Case 59
                            objEmp.ienvfat = IIf(sCampoValor = "yes", 1, 0)
                        Case 60
                            objEmp.ienvnota = IIf(sCampoValor = "yes", 1, 0)
                        Case 61
                            objEmp.ienvocor = IIf(sCampoValor = "yes", 1, 0)
                        Case 62
                            objEmp.scodest = StrParaStrCoinfo(sCampoValor)
                        Case 63
                            'If Len(sCampoValor) > 1000 Then MsgBox ("erro stexto")
                            objEmp.stexto1 = StrParaStrCoinfo(left(sCampoValor, 250))
                            If Len(sCampoValor) > 250 Then
                                objEmp.stexto2 = StrParaStrCoinfo(Mid(sCampoValor, 251, 250))
                                If Len(sCampoValor) > 500 Then
                                    objEmp.stexto3 = StrParaStrCoinfo(Mid(sCampoValor, 501, 250))
                                    If Len(sCampoValor) > 750 Then
                                        objEmp.stexto4 = StrParaStrCoinfo(Mid(sCampoValor, 751, 250))
                                    End If
                                End If
                            End If
                        Case 64
                            objEmp.sEmail2 = StrParaStrCoinfo(sCampoValor)
                        Case 65
                            objEmp.scodgru = StrParaStrCoinfo(sCampoValor)
                        Case 66
                            objEmp.lEmpPai = StrParaLongCoinfo(sCampoValor)
                        Case 67
                            objEmp.icorrent = IIf(sCampoValor = "yes", 1, 0)
                            
                        Case 68
                            objEmp.lempagt = StrParaLongCoinfo(sCampoValor)
                            If objEmp.lempagt = 1000000 Then objEmp.lempagt = 0 '??? o sisweb nao consegue tratar isso
                            
                        Case 69
                            objEmp.dcomis = StrParaDblCoinfo(sCampoValor)
                            
                        Case 70
                            objEmp.lcodrep = StrParaLongCoinfo(sCampoValor)
                            If objEmp.lcodrep = 1000000 Then objEmp.lcodrep = 0 '??? o sisweb nao consegue tratar isso
                            
                        Case 71
                            objEmp.sSerie = StrParaStrCoinfo(sCampoValor)
                        Case 72
                            objEmp.iAporte = IIf(sCampoValor = "yes", 1, 0)
                        Case 73
                            objEmp.inf = IIf(sCampoValor = "yes", 1, 0)
                        Case 74
                            objEmp.scodccm = StrParaStrCoinfo(sCampoValor)
                        Case 75
                            objEmp.iusaccre = IIf(sCampoValor = "yes", 1, 0)
                        Case 76
                            objEmp.isoprdexc = IIf(sCampoValor = "yes", 1, 0)
                        Case 77
                            objEmp.llovouc = StrParaLong(sCampoValor)

                        Case Else
                            gError 130609
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 78 Then
            sRegContinuacao = ts.ReadLine
            sRegistro = sRegistro & sRegContinuacao
            lTamanhoDados = lTamanhoDados + Len(sRegContinuacao) + 1 ' + 1 p/'\0x0A'
            GoTo aqui
        End If
        
        objEmp.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objEmp
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoEmpresa (NumIntArq, NumRegArqTexto, codemp, nchave, nome, endereco, bairro, cidade, UF, cep, fone1, fone2, fax, " & _
                "cgc , ie, Natureza, contato, Cargo, cob, praca, endercob, limcred, datacad, cidcob, ufcob, " & _
                "dataalt , codpes, desconto, obs1, obs2, obs3, obs4, obs5, obs6, codred, Usuario, tipoempr, " & _
                "datareg , faturado, condpag, cpf, email, moecred, codprom, bairrcob, cepcob, Ativo, expirlim, " & _
                "fixapraz , redcomis, cliente, comissiona, exportado, numban, iss, agenban, nconta, textcon1, " & _
                "textcon2 , textcon3, envfat, envnota, envocor, codest, texto1, texto2, texto3, texto4, email2, codgru, emppai, corrent, " & _
                "empagt, comis, codrep, serie, aporte, nf, codccm, usaccre, soprdexc, lovouc, NoCorporatorEm) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .lcodemp, left(.snchave, 50), left(.sNome, 50), left(.sEndereco, 65), left(.sBairro, 20), left(.sCidade, 30), .sUF, .sCEP, _
                .sfone1, .sfone2, .sFax, .sCgc, .sie, .iNatureza, .sContato, .sCargo, .icob, .spraca, .sendercob, .dlimcred, _
                .dtDataCad, .scidcob, .sufcob, .dtDataAlt, .iCodpes, .dDesconto, .sobs1, .sobs2, .sobs3, .sobs4, .sobs5, .sobs6, _
                .lcodred, .sUsuario, .itipoempr, .dtdatareg, .ifaturado, .icondpag, .sCPF, .sEmail, .smoecred, .icodprom, .sbairrcob, _
                .scepcob, .iAtivo, .dtexpirlim, .ifixapraz, .dredcomis, .icliente, .icomissiona, .iExportado, .inumban, .dISS, _
                .lagenban, .lnconta, .stextcon1, .stextcon2, .stextcon3, .ienvfat, .ienvnota, .ienvocor, .scodest, .stexto1, .stexto2, .stexto3, .stexto4, _
                .sEmail2, .scodgru, .lEmpPai, .icorrent, .lempagt, .dcomis, .lcodrep, .sSerie, .iAporte, .inf, .scodccm, .iusaccre, _
                .isoprdexc, .llovouc, DATA_NULA)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130599
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130607
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqEmprHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto2 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto2:

    Coinfo_ImportarArqTexto2 = gErr
     
    Select Case gErr
                 
        Case 130595
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130596
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130604
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqEmprHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqEmprHora)

        Case 130609
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqEmprHora, CStr(lNumRegs))
                      
        Case 130599
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOEMPRESA", gErr, objEmp.lcodemp, objEmp.sNome)
       
        Case 130607
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 208886
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_INVALIDA", gErr, objEmp.sUF)
              
        Case ERRO_SEM_MENSAGEM, 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130563)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    If bAbriuArq Then
    
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqEmprHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    
        Set objTS = objFSO.OpenTextFile("C:\SGE_Log_" & objImpCoinfo.sArqEmprHora, ForWriting, True)
    
        objTS.WriteLine "NumReg: " & CStr(lNumRegs)
        objTS.WriteLine "Pos: " & CStr(iCampoPos)
        objTS.WriteLine "ValorCampo : " & sCampoValor
    
        If Not (objEmp Is Nothing) Then
            With objEmp
                objTS.WriteLine "Código: " & objEmp.lcodemp
                objTS.WriteLine "Nome: " & objEmp.sNome & " Tamanho: " & CStr(Len(objEmp.sNome))
            End With
        End If
        
        objTS.Close
     
    End If
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto3(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar produtos

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objProd As ClassImpCoinfoProd, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto3

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130610
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130611
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_PROD", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpproduto", objImpCoinfo.sArqProdHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqProdHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqProdHora = "" Then
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqProdHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqProd = lNumIntArqAnt
        gError 130619
    End If
    
    objImpCoinfo.lNumIntArqProd = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqProdHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine

        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
        Set objProd = New ClassImpCoinfoProd
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objProd.sCodPro = StrParaStrCoinfo(sCampoValor)
                        Case 2
                            objProd.sDescricao = left(StrParaStrCoinfo(sCampoValor), STRING_PRODUTO_DESCRICAO_TELA)
                        Case 3
                            objProd.iCodSeg = StrParaInt(sCampoValor)
                        Case 4
                            objProd.sAlcance = StrParaStrCoinfo(sCampoValor)
                        Case 5
                            objProd.sMoeda = StrParaStrCoinfo(sCampoValor)
                        Case 6
                            objProd.sTipProd = StrParaStrCoinfo(sCampoValor)
                        Case 7
                            objProd.dISS = StrParaDblCoinfo(sCampoValor)
                        Case 8
                            objProd.dExento = StrParaDblCoinfo(sCampoValor)
                        Case 9
                            objProd.dVlPadrao = StrParaDblCoinfo(sCampoValor)
                        Case 10
                            objProd.dPercDesc = StrParaDblCoinfo(sCampoValor)
                        Case 11
                            objProd.iLFamiliar = IIf(sCampoValor = "yes", 1, 0)
                        Case 12
                            objProd.dPercTitular = StrParaDblCoinfo(sCampoValor)
                        Case 13
                            objProd.dPercAcomp = StrParaDblCoinfo(sCampoValor)
                        Case 14
                            objProd.dcomismax = StrParaDblCoinfo(sCampoValor)
                        Case 15
                            objProd.iLimIdade = StrParaInt(sCampoValor)
                        Case 16
                            objProd.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 17
                            objProd.dtdatareg = StrParaDateCoinfo(sCampoValor)
                        Case 18
                            objProd.iValidade = StrParaInt(sCampoValor)
                        Case 19
                            objProd.dcomispro = StrParaDblCoinfo(sCampoValor)
                        Case 20
                            objProd.iLcatalogo = IIf(sCampoValor = "yes", 1, 0)
                        Case 21
                            objProd.lEmpPai = StrParaLong(sCampoValor)
                        Case 22
                            objProd.iDiasVal = StrParaInt(sCampoValor)
                        Case 23
                            objProd.imaxfamil = StrParaInt(sCampoValor)
                        Case 24
                            objProd.isupinf = IIf(sCampoValor = "yes", 1, 0)
                        Case 25
                            objProd.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130624
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 26 Then gError 130624
        
        objProd.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objProd
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoProduto (NumIntArq,NumRegArqTexto,CodPro,Descricao,CodSeg,Alcance,Moeda,TipProd,ISS,Exento,VlPadrao,PercDesc," & _
                "LFamiliar,PercTitular,PercAcomp,comismax,LimIdade,usuario,datareg,Validade,comispro," & _
                "Lcatalogo , EmpPai, DiasVal, maxfamil, supinf) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .sCodPro, .sDescricao, .iCodSeg, .sAlcance, .sMoeda, .sTipProd, .dISS, .dExento, .dVlPadrao, _
                .dPercDesc, .iLFamiliar, .dPercTitular, .dPercAcomp, .dcomismax, .iLimIdade, .sUsuario, .dtdatareg, .iValidade, _
                .dcomispro, .iLcatalogo, .lEmpPai, .iDiasVal, .imaxfamil, .isupinf)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130614
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130622
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqProdHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto3 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto3:

    Coinfo_ImportarArqTexto3 = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM, 53
        
        Case 130619
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqProdHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqProdHora)

        Case 130614
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOPRODUTO", gErr, objProd.sCodPro, objProd.sDescricao)
        
        Case 130610
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130611
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130622
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130624
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqProdHora, CStr(lNumRegs))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130564)
     
    End Select
    
    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqProdHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto4(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar faturas (cabecalho e itens)

Dim lErro As Long, lTransacao As Long
Dim fso As New FileSystemObject
Dim bTemArq1 As Boolean
Dim bTemArq2 As Boolean
Dim lNumRegsFat As Long
Dim lNumRegsItemFat As Long

On Error GoTo Erro_Coinfo_ImportarArqTexto4

    bTemArq1 = False
    bTemArq2 = False

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130644
    
'    lErro = Coinfo_ImportarArqTexto4a(objImpCoinfo, lNumRegsFat)
'    If lErro <> SUCESSO And lErro <> 130640 And lErro <> 53 Then gError 130645
'    If lErro <> 53 Then bTemArq1 = True
    
    lErro = Coinfo_ImportarArqTexto4b(objImpCoinfo, lNumRegsItemFat)
    If lErro <> SUCESSO And lErro <> 130651 And lErro <> 53 Then gError 130646
    If lErro <> 53 Then bTemArq2 = True
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130647
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqFatHora & " teve " & CStr(lNumRegsFat) & " registros importados com sucesso." & vbNewLine
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqItemFatHora & " teve " & CStr(lNumRegsItemFat) & " registros importados com sucesso." & vbNewLine
   
    Coinfo_ImportarArqTexto4 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto4:

    Coinfo_ImportarArqTexto4 = gErr
     
    Select Case gErr
          
        Case 130645, 130646
        
        Case 130644
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130647
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130648)
     
    End Select
     
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Coinfo_AtualizaArqsLidos(ByVal objImpCoinfo As ClassImpCoinfo, ByVal sNomeArq As String, ByVal sIdentificacaoArq As String, lNumIntArqAnt As Long, ByVal lNumRegs As Long, ByVal lTamanhoDados As Long, ByVal lNumIntArq As Long) As Long
'se o arquivo já havia sido lido antes, obtem o lNumIntArq usado anteriormente e coloca-o em lNumIntArqAnt
'senao apenas insere no controle

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 1) As Long

On Error GoTo Erro_Coinfo_AtualizaArqsLidos

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130709
        
    Next
    
    Select Case sNomeArq
    
        Case "fptipoempr"
            objImpCoinfo.lNumRegsArqTipoEmpr = lNumRegs
            
        Case "fpempresa"
            objImpCoinfo.lNumRegsArqEmpr = lNumRegs
                    
        Case "fpproduto"
            objImpCoinfo.lNumRegsArqProd = lNumRegs
            
        Case "fpcabecfat"
            objImpCoinfo.lNumRegsArqFat = lNumRegs
            
        Case "fpitemfatu"
            objImpCoinfo.lNumRegsArqItemFat = lNumRegs
            
        Case "fpocorrencia"
            objImpCoinfo.lNumRegsArqOcor = lNumRegs
            
        Case "fpnf"
            objImpCoinfo.lNumRegsArqNF = lNumRegs
            
        Case "fpitemnf"
            objImpCoinfo.lNumRegsArqItemNF = lNumRegs
            
        Case "fpgrupo"
            objImpCoinfo.lNumRegsArqGrupo = lNumRegs
                    
    End Select
    
    'se o arquivo já tiver sido importado anteriormente, registrar isto no log
    lErro = Comando_Executar(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND IdentificacaoArq =? AND Sistema = ?", lNumIntArqAnt, sNomeArq, sIdentificacaoArq, iSistemaIntegrado)
    If lErro <> AD_SQL_SUCESSO Then gError 130710
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130711
    
    'se o arquivo ainda nao foi importado
    If lErro <> AD_SQL_SUCESSO Then
        lNumIntArqAnt = 0
        lErro = Comando_Executar(alComando(1), "INSERT INTO CoinfoArqsLidos (NomeArq, IdentificacaoArq, NumIntImportacao, DataImportacao, HoraImportacao, Usuario, NumRegs, TamanhoDados,DataAtualizado,Sistema) VALUES (?,?,?,?,?,?,?,?,?,?)", _
            sNomeArq, sIdentificacaoArq, lNumIntArq, Date, CDbl(CDate(Format(Now, "HH:MM:SS"))), gsUsuario, lNumRegs, lTamanhoDados, DATA_NULA, iSistemaIntegrado)
        If lErro <> AD_SQL_SUCESSO Then gError 130712
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_AtualizaArqsLidos = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaArqsLidos:

    Coinfo_AtualizaArqsLidos = gErr
     
    Select Case gErr
          
        Case 130709
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130710, 130711
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130712
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOARQSLIDOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130571)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Coinfo_AtualizarLogImportacao(ByVal objImpCoinfo As ClassImpCoinfo) As Long

Dim lErro As Long

On Error GoTo Erro_Coinfo_AtualizarLogImportacao

    Coinfo_AtualizarLogImportacao = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarLogImportacao:

    Coinfo_AtualizarLogImportacao = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130591)
     
    End Select
     
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto4a(ByVal objImpCoinfo As ClassImpCoinfo, lNumRegs As Long) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar cabecalhos das faturas

Dim lErro As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objFat As ClassImpCoinfoFat

On Error GoTo Erro_Coinfo_ImportarArqTexto4a

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130632
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_FATURA", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 130633

    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpcabecfat", objImpCoinfo.sArqFatHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError 130639

    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqFatHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine

    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqFatHora = "" Then
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqFatHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqFat = lNumIntArqAnt
        gError 130640
    End If
    
    objImpCoinfo.lNumIntArqFat = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqFatHora, 1, 0)
    bArqAberto = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 2
'            Exit Do
'        End If
        
        lNumRegs = lNumRegs + 1
        
        Set objFat = New ClassImpCoinfoFat
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objFat.lNumFat = StrParaLong(sCampoValor)
                        Case 2
                            objFat.dtDataFat = StrParaDateCoinfo(sCampoValor)
                        Case 3
                            objFat.lcodemp = StrParaLong(sCampoValor)
                        Case 4
                            objFat.sMoeda = StrParaStrCoinfo(sCampoValor)
                        Case 5
                            objFat.dValor = StrParaDblCoinfo(sCampoValor)
                        Case 6
                            objFat.scodest = StrParaStrCoinfo(sCampoValor)
                        Case 7
                            objFat.dtdataven = StrParaDateCoinfo(sCampoValor)
                        Case 8
                            objFat.iCancelado = IIf(sCampoValor = "yes", 1, 0)
                        Case 9
                            objFat.icpcr = IIf(sCampoValor = "yes", 1, 0)
                        Case 10
                            objFat.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 11
                            objFat.dtdatareg = StrParaDateCoinfo(sCampoValor)
                        Case 12
                            objFat.lrecno = StrParaLong(sCampoValor)
                        
                        'acho que nao está no arquivo
'                        Case 13
'                            objFat.lrecnsica = StrParaLong(sCampoValor)
                        Case 13
                            objFat.dvalisscl = StrParaDblCoinfo(sCampoValor)
                        Case 14
                            objFat.dPercDesc = StrParaDblCoinfo(sCampoValor)
                        Case 15
                            objFat.dvaldcbrl = StrParaDblCoinfo(sCampoValor)
                        Case 16
                            objFat.dvalbru = StrParaDblCoinfo(sCampoValor)
                        Case 17
                            objFat.dtDataCanc = StrParaDateCoinfo(sCampoValor)
                        Case 18
                            objFat.iPago = IIf(sCampoValor = "yes", 1, 0)
                        Case 19
                            objFat.dtDataPag = StrParaDateCoinfo(sCampoValor)
                        Case 20
                            objFat.lNumNF = StrParaLong(sCampoValor)
                        Case 21
                            objFat.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130634
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 22 Then gError 130663
        
        objFat.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objFat
        
            If Year(.dtDataPag) = 206 Then
                .dtDataPag = StrParaDate(Day(.dtDataPag) & "/" & Month(.dtDataPag) & "/2006")
            End If
        
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoFatura (NumIntArq, NumRegArqTexto, numfat, datafat, codemp, moeda, valor, codest, dataven, cancelado, cpcr, " & _
                "usuario, datareg, recno, recnsica, valisscl, percdesc, valdcbrl, valbru, datacanc, " & _
                "pago, datapag, numnf, NoCorporatorEm) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .lNumFat, .dtDataFat, .lcodemp, .sMoeda, .dValor, .scodest, .dtdataven, .iCancelado, _
                .icpcr, .sUsuario, .dtdatareg, .lrecno, .lrecnsica, .dvalisscl, .dPercDesc, .dvaldcbrl, .dvalbru, _
                .dtDataCanc, .iPago, .dtDataPag, .lNumNF, DATA_NULA)
        End With
        If lErro <> AD_SQL_SUCESSO Then
            gError 130635
        End If
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
'    If iSecao = 1 Then gError 130636
'
'    'leitura da seção de controle
'
'    'Até chegar ao fim do arquivo
'    Do While True
'
'        'Busca o próximo registro do arquivo
'        sRegistro = ts.ReadLine
'
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 3
'            Exit Do
'        End If
'
'        If sRegistro = "filename=fpcabecfat" Then iTestesOk = iTestesOk + 1
'        If Left(sRegistro, 8) = "records=" Then
'            If lNumRegs <> StrParaLong(Mid(sRegistro, 9)) Then gError 130637
'            iTestesOk = iTestesOk + 1
'        End If
'        If Left(sRegistro, 10) = "timestamp=" Then
'            sArqTimeStamp = Mid(sRegistro, 11)
'        End If
'
'    Loop
'
'    If iTestesOk <> 2 Then gError 130638
    
'    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
'    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpcabecfat", sArqTimeStamp, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
'    If lErro <> SUCESSO Then gError 130639
'
'    'se o arquivo já havia sido importado anteriormente
'    If lNumIntArqAnt <> 0 Then
'        objImpCoinfo.lNumIntArqFat = lNumIntArqAnt
'        gError 130640
'    End If
'
'    objImpCoinfo.lNumIntArqFat = lNumIntArq
    
'    If iSecao = 2 Then gError 130641
'
'    'leitura da seção com o tamanho dos dados
'
'    'Busca o próximo registro do arquivo
'    sRegistro = ts.ReadLine
'
'    If StrParaLong(sRegistro) <> lTamanhoDados Then gError 130642
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto4a = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto4a:

    Coinfo_ImportarArqTexto4a = gErr
     
    Select Case gErr
          
        Case 130633, 130639, 130640, 53
        
        Case 130634, 130636, 130637, 130638, 130641, 130642, 130663
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_COINFOFATURA", gErr)
        
        Case 130632
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130635
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOFATURA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130631)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto4b(ByVal objImpCoinfo As ClassImpCoinfo, lNumRegs As Long) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar itens das faturas

Dim lErro As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objItemFat As ClassImpCoinfoItemFat
Dim iModelo As Integer
Dim sData As String
Dim sDataAux As String
Dim objVoucher As ClassTRVVouchers
Dim objOCR As ClassTRVOcorrencias
Dim bCondicao As Boolean
Dim objFSO As New FileSystemObject
Dim objTS As TextStream, sCodPro As String, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto4b

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130650
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_ITEMFATURA", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpitemfatu", objImpCoinfo.sArqItemFatHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqItemFatHora = "" Then gError 130651
        
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqItemFatHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
    
    iModelo = 3
    
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1
        
        lNumRegs = lNumRegs + 1
        
        Set objItemFat = New ClassImpCoinfoItemFat
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objItemFat.lcodemp = StrParaLong(sCampoValor)
                        Case 2
                            objItemFat.dtData = StrParaDateCoinfo(sCampoValor)
                        Case 3
                            objItemFat.lNumVou = StrParaLong(sCampoValor)
                        Case 4
                            objItemFat.sMoeda = StrParaStrCoinfo(sCampoValor)
                        Case 5
                            objItemFat.dValor = StrParaDblCoinfo(sCampoValor)
                        Case 6
                            objItemFat.scodest = StrParaStrCoinfo(sCampoValor)
                        Case 7
                            objItemFat.lNumFat = StrParaLong(sCampoValor)
                        Case 8
                            objItemFat.dtdataven = StrParaDateCoinfo(sCampoValor)
                        Case 9
                            objItemFat.lnumdoc = StrParaLong(sCampoValor)
                        Case 10
                            objItemFat.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 11
                            objItemFat.dtdatareg = StrParaDateCoinfo(sCampoValor)
                        Case 12
                            objItemFat.iMarca = IIf(sCampoValor = "yes", 1, 0)
                        Case 13
                            objItemFat.sTipoDoc = StrParaStrCoinfo(sCampoValor)
                        Case 14 'QtdPax
                            objItemFat.iQtdPax = StrParaInt(sCampoValor)
                        Case 15
                            objItemFat.lrecnsica = StrParaLong(sCampoValor)
                        Case 16
                            objItemFat.icondpag = StrParaInt(sCampoValor)
                        Case 17
                            objItemFat.sSerie = StrParaStrCoinfo(sCampoValor)
                        Case 18
                            objItemFat.sTipVou = StrParaStrCoinfo(sCampoValor)
                        Case 19
                            objItemFat.scodgru = StrParaStrCoinfo(sCampoValor)
                        Case 20
                            objItemFat.sControle = StrParaStrCoinfo(sCampoValor)
                        Case 21
                            objItemFat.iCartao = IIf(sCampoValor = "yes", 1, 0)
                        Case 22
                            objItemFat.sCodPro = StrParaStrCoinfo(sCampoValor)
                        Case 23 'Grupo
                            objItemFat.sGrupo = StrParaStrCoinfo(sCampoValor)
                        Case 24
                            objItemFat.sNumCCred = StrParaStrCoinfo(sCampoValor)
                        Case 25
                            If Len(sCampoValor) <= 7 Then
                                objItemFat.sValidCC = StrParaStrCoinfo(sCampoValor)
                            Else
                                objItemFat.sValidCC = left(StrParaStrCoinfo(sCampoValor), 2) & "/" & right(StrParaStrCoinfo(sCampoValor), 4)
                            End If
                        Case 26
                            objItemFat.iQuantParc = StrParaInt(sCampoValor)
                        Case 27
                            objItemFat.dVlCartao = StrParaDblCoinfo(sCampoValor)
                        Case 28
                            objItemFat.sNumAuto = StrParaStrCoinfo(sCampoValor)
                            If objItemFat.sNumAuto = "0" Then objItemFat.sNumAuto = ""
                        Case 29
                            objItemFat.sCiaCart = StrParaStrCoinfo(sCampoValor)
                        Case 30
                            objItemFat.iExcluido = IIf(sCampoValor = "yes", 1, 0)
                        Case 31
                            objItemFat.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case 32
                            objItemFat.dtDataExp = StrParaDateCoinfo(sCampoValor)
                        Case 33 'diasantc
                            objItemFat.idiasantc = IIf(sCampoValor = "yes", 1, 0)
                        Case 34 'master
                            objItemFat.imaster = IIf(sCampoValor = "yes", 1, 0)
                        Case 35 'nummstr
                            objItemFat.lnummstr = StrParaLong(sCampoValor)
                        Case Else
                            If iSistemaIntegrado = SISTEMA_INTEGRADO_SIGAV Then gError 130652
                            
                            Select Case iCampoPos
                            
                                Case 36
                                    objItemFat.lKit = StrParaLong(sCampoValor)
                                Case 37
                                    objItemFat.dtDataRecepcao = StrParaDateCoinfo(sCampoValor)
                                Case 38
                                    objItemFat.lEmissor = StrParaLong(sCampoValor)
                                Case 39
                                    objItemFat.sDestino = left(StrParaStrCoinfo(sCampoValor), 50)
                                Case 40
                                    objItemFat.dtdatainiciovigencia = StrParaDateCoinfo(sCampoValor)
                                Case 41
                                    objItemFat.dtdatafimvigencia = StrParaDateCoinfo(sCampoValor)
                                Case 42
                                    objItemFat.sIdioma = StrParaStrCoinfo(sCampoValor)
                                Case 43
                                    objItemFat.sDestinoVou = left(StrParaStrCoinfo(sCampoValor), 50)
                                Case 44
                                    objItemFat.dtarifamoeda = StrParaDblCoinfo(sCampoValor)
                                Case 45
                                    objItemFat.dCambio = StrParaDblCoinfo(sCampoValor)
                                Case 46
                                    objItemFat.dtarifareal = StrParaDblCoinfo(sCampoValor)
                                Case 47
                                    objItemFat.stitularcartao = StrParaStrCoinfo(sCampoValor)
                                Case 48
                                    objItemFat.stitularcartaocpf = StrParaStrCoinfo(sCampoValor)
                                Case 49
                                    objItemFat.spaxcartaofid = StrParaStrCoinfo(sCampoValor)
                                Case 50
                                    objItemFat.spaxsobrenome = left(StrParaStrCoinfo(sCampoValor), 50)
                                Case 51
                                    objItemFat.spaxnome = left(StrParaStrCoinfo(sCampoValor), 50)
                                Case 52
                                    objItemFat.dtpaxdatanasc = StrParaDateCoinfo(sCampoValor)
                                Case 53
                                    objItemFat.spaxsexo = StrParaStrCoinfo(sCampoValor)
                                Case 54
                                    objItemFat.spaxtipodoc = StrParaStrCoinfo(sCampoValor)
                                Case 55
                                    objItemFat.spaxnumdoc = StrParaStrCoinfo(sCampoValor)
                                Case 56
                                    objItemFat.spaxendereco = left(StrParaStrCoinfo(sCampoValor), STRING_ENDERECO)
                                Case 57
                                    objItemFat.spaxbairro = left(StrParaStrCoinfo(sCampoValor), STRING_BAIRRO)
                                Case 58
                                    objItemFat.spaxcidade = left(StrParaStrCoinfo(sCampoValor), STRING_CIDADE)
                                Case 59
                                    objItemFat.spaxcep = left(StrParaStrCoinfo(sCampoValor), STRING_CEP)
                                Case 60
                                    objItemFat.spaxuf = StrParaStrCoinfo(sCampoValor)
                                    If objItemFat.spaxuf <> "" And Len(Trim(objItemFat.spaxuf)) <> 2 Then gError 208885
                                Case 61
                                    objItemFat.spaxemail = left(StrParaStrCoinfo(sCampoValor), 64)
                                Case 62
                                    objItemFat.spaxcontato = left(StrParaStrCoinfo(sCampoValor), STRING_CONTATO)
                                Case 63
                                    objItemFat.spaxtelefone1 = left(StrParaStrCoinfo(sCampoValor), STRING_TELEFONE)
                                Case 64
                                    objItemFat.spaxtelefone2 = left(StrParaStrCoinfo(sCampoValor), STRING_TELEFONE)
                                Case 65
                                    objItemFat.sConvenio = StrParaStrCoinfo(sCampoValor)
                                Case 66
                                    objItemFat.lCodEmpVou = StrParaLong(sCampoValor)
                                Case 67
                                    objItemFat.dTarifaUnitaria = StrParaDblCoinfo(sCampoValor)
                                Case Else
                                    gError 130652
                                
                            End Select
                    
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        If iSistemaIntegrado = SISTEMA_INTEGRADO_SIGAV Then
            If iCampoPos <> 36 Then gError 130652 'nummstr
            objItemFat.lKit = StrParaLong(sCampoValor)
        Else
            If iCampoPos <> 68 Then gError 130652 'nummstr
            objItemFat.dTarifaUNNormal = StrParaDblCoinfo(sCampoValor)
        End If
                
        'inserir registro
        With objItemFat
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoItemFatura (NumIntArq, NumRegArqTexto, codemp, Data, numvou, moeda, valor, codest, numfat, dataven, numdoc, " & _
                "usuario, datareg, marca, tipodoc, recnsica, condpag, Serie, TipVou, codgru, Controle, " & _
                "cartao, CodPro, numccred, validcc, quantparc, vlcartao, numauto, ciacart, NoCorporatorEm, Excluido,QtdPax,Grupo,diasantc,master,nummstr,kit, " & _
                "datarecepcao,emissor,destino,datainiciovigencia,datafimvigencia,idioma,destinovou,tarifamoeda,cambio,tarifareal,titularcartao,titularcartaocpf,paxcartaofid,paxsobrenome,paxnome,paxdatanasc,paxsexo,paxtipodoc,paxnumdoc,paxendereco,paxbairro,paxcidade,paxcep,paxuf,paxemail,paxcontato,paxtelefone1,paxtelefone2, convenio,codempvou,TarifaUnitaria, TarifaUNNormal) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .lcodemp, .dtData, .lNumVou, .sMoeda, .dValor, .scodest, .lNumFat, .dtdataven, _
                .lnumdoc, .sUsuario, .dtdatareg, .iMarca, .sTipoDoc, .lrecnsica, .icondpag, .sSerie, .sTipVou, _
                .scodgru, .sControle, .iCartao, .sCodPro, .sNumCCred, .sValidCC, .iQuantParc, .dVlCartao, .sNumAuto, .sCiaCart, DATA_NULA, .iExcluido, .iQtdPax, .sGrupo, .idiasantc, .imaster, .lnummstr, .lKit, _
                .dtDataRecepcao, .lEmissor, .sDestino, .dtdatainiciovigencia, .dtdatafimvigencia, .sIdioma, .sDestinoVou, .dtarifamoeda, .dCambio, .dtarifareal, .stitularcartao, .stitularcartaocpf, .spaxcartaofid, .spaxsobrenome, .spaxnome, .dtpaxdatanasc, .spaxsexo, .spaxtipodoc, .spaxnumdoc, .spaxendereco, .spaxbairro, .spaxcidade, .spaxcep, .spaxuf, .spaxemail, .spaxcontato, .spaxtelefone1, .spaxtelefone2, .sConvenio, .lCodEmpVou, .dTarifaUnitaria, .dTarifaUNNormal)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130654
        
        If giVersaoTRV >= TRV_VERSAO_COMISSIONAMENTO Then
                       
            'Se é um voucher ou comissão anterior a entrada em produção do Corporator
            bCondicao = objItemFat.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Or ((objItemFat.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or objItemFat.sTipoDoc = TRV_TIPODOC_CMC_TEXTO) And gdtDataInicioComisCorp > objItemFat.dtData)
               
        Else
               
            'Se é um voucher
            bCondicao = objItemFat.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Or objItemFat.sTipoDoc = TRV_TIPODOC_CMR_TEXTO Or objItemFat.sTipoDoc = TRV_TIPODOC_CMC_TEXTO
        
        End If
        
        If bCondicao Then
        
            Set objVoucher = New ClassTRVVouchers
            
            objVoucher.dtData = objItemFat.dtData
            If objItemFat.iCartao = MARCADO And objItemFat.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
                objVoucher.dValor = objItemFat.dVlCartao
            Else
                objVoucher.dValor = objItemFat.dValor
            End If
            objVoucher.iCartao = objItemFat.iCartao
            objVoucher.iCondPagto = objItemFat.icondpag
            objVoucher.lCliente = objItemFat.lcodemp
            objVoucher.lNumFatCoinfo = 0 'objItemFat.lNumFat
            objVoucher.lNumIntArq = lNumIntArq
            objVoucher.lNumRegArqTexto = lNumRegs
            objVoucher.lNumVou = objItemFat.lNumVou
            objVoucher.lrecnsica = objItemFat.lrecnsica
            objVoucher.sCodGrupo = objItemFat.scodgru
            objVoucher.sSerie = objItemFat.sSerie
            objVoucher.sTipoDoc = objItemFat.sTipoDoc
            objVoucher.sTipVou = objItemFat.sTipVou
            objVoucher.iPax = objItemFat.iQtdPax
            objVoucher.sControle = objItemFat.sControle
            objVoucher.sMoeda = objItemFat.sMoeda
            
            sCodPro = objItemFat.sCodPro
            Call TRVConverte_Prod(sCodPro)
'            sCodPro = Replace(sCodPro, "/", " ")
'            sCodPro = Replace(sCodPro, "\", " ")
'            sCodPro = Replace(sCodPro, "-", " ")
'            sCodPro = Replace(sCodPro, "_", " ")
'            sCodPro = Replace(sCodPro, "&", " ")
'            sCodPro = left(sCodPro, TAMANHO_SEGMENTO_PRODUTO)
'            sCodPro = sCodPro & String(TAMANHO_SEGMENTO_PRODUTO - Len(sCodPro), 32)
            
            objVoucher.sProduto = sCodPro
            objVoucher.sCiaCart = objItemFat.sCiaCart
            objVoucher.sNumCCred = objItemFat.sNumCCred
            objVoucher.sNumAuto = objItemFat.sNumAuto
            objVoucher.iQuantParc = objItemFat.iQuantParc
            objVoucher.idiasantc = objItemFat.idiasantc
            objVoucher.iKit = objItemFat.lKit
            objVoucher.iStatus = STATUS_TRV_OCR_LIBERADO
            objVoucher.dTarifaUNNormal = objItemFat.dTarifaUNNormal
            objVoucher.lClienteVou = objItemFat.lCodEmpVou
            objVoucher.sUsuarioWeb = objItemFat.sUsuario
            
            'Indica que foi gerado boleto por fora do sistema
            If objItemFat.lNumFat > TRV_INI_SEQ_BOLETO_GER_FORA_CORP Then
                'objVoucher.lNumIntDocDestino = -1 'Para não permitir ser faturado normalmente
                objVoucher.iTrataBoleto = 1 'Indica que precisa gravar o boleto do voucher
                'Dados para gerar o boleto
                objVoucher.lNumBoleto = objItemFat.lNumFat
                objVoucher.dtDataVencBoleto = objItemFat.dtdataven
                objVoucher.dValorBoleto = objItemFat.dVlCartao
            End If
            
            If objItemFat.scodest = gsEmpresaTRV Then
                lErro = CF("TRVVouchers_Grava", objVoucher, MARCADO)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
            If objItemFat.iExcluido = MARCADO Or objItemFat.scodest <> gsEmpresaTRV Then
                lErro = CF("TRVVouchers_Cancela", objVoucher, True)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
        
        'Se é na inativação da própria empresa
        ElseIf objItemFat.sTipoDoc = TRV_TIPODOC_NVL_TEXTO And objItemFat.scodest = objImpCoinfo.sSiglaEmpresa Then
                
            Set objVoucher = New ClassTRVVouchers
                
            objVoucher.lNumVou = objItemFat.lNumVou
            objVoucher.sSerie = objItemFat.sSerie
            objVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO
            objVoucher.sTipVou = objItemFat.sTipVou
            objVoucher.lCliente = objItemFat.lcodemp
            objVoucher.dValor = objItemFat.dValor
            objVoucher.iCartao = objItemFat.iCartao
            
            If objItemFat.iExcluido = MARCADO Then
                
                Set objOCR = New ClassTRVOcorrencias
                
                objOCR.sSerie = objVoucher.sSerie
                objOCR.sTipoDoc = objVoucher.sTipVou
                objOCR.lNumVou = objVoucher.lNumVou
                objOCR.lCliente = objVoucher.lCliente
            
                'Exclui a inativação
                lErro = CF("TRVOcorrencias_Exclui_EmTrans", objOCR, True)
            Else
                lErro = CF("TRVVouchers_Cancela", objVoucher, True, objVoucher.dValor, objItemFat.dtData)
            End If
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        ElseIf objItemFat.sTipoDoc = TRV_TIPODOC_OCR_TEXTO And objItemFat.scodest = objImpCoinfo.sSiglaEmpresa Then
            
            If objItemFat.iExcluido = MARCADO Then
                
                Set objOCR = New ClassTRVOcorrencias
                
                objOCR.lCodigo = objItemFat.lnumdoc + FATOR_COD_OCR_IMPORTACAO
            
                'Exclui a inativação
                lErro = CF("TRVOcorrencias_Exclui_EmTrans", objOCR, False, True)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End If
            
        End If
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto4b = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto4b:

    Coinfo_ImportarArqTexto4b = gErr
     
    Select Case gErr
        
        Case 130650
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 130651
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqItemFatHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqItemFatHora)
                        
        Case 130652
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqItemFatHora, CStr(lNumRegs))
          
        Case 130654
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOITEMFATURA", gErr, objItemFat.sTipVou, objItemFat.sSerie, objItemFat.lNumVou)
            
        Case 208885
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_INVALIDA", gErr, objItemFat.spaxuf)
        
        Case ERRO_SEM_MENSAGEM, 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130649)
     
    End Select
    
    If bAbriuArq Then
    
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqItemFatHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    
        Set objTS = objFSO.OpenTextFile("C:\SGE\SGE_Log_" & objImpCoinfo.sArqItemFatHora, ForWriting, True)
    
        objTS.WriteLine "NumReg: " & CStr(lNumRegs)
        objTS.WriteLine "Pos: " & CStr(iCampoPos)
        objTS.WriteLine "ValorCampo : " & sCampoValor
    
        If Not (objItemFat Is Nothing) Then
            With objItemFat
                objTS.WriteLine "Tipo: " & .sTipVou & " Tamanho: " & CStr(Len(.sTipVou))
                objTS.WriteLine "Serie: " & .sSerie & " Tamanho: " & CStr(Len(.sSerie))
                objTS.WriteLine "NumVou: " & CStr(.lNumVou)
                objTS.WriteLine "DataEmissao: " & CStr(.dtData)
                objTS.WriteLine "DataRecepcao: " & CStr(.dtDataRecepcao)
                objTS.WriteLine "Cliente: " & CStr(.lCodEmpVou)
                objTS.WriteLine "Emissor: " & CStr(.lEmissor)
                objTS.WriteLine "Produto: " & .sCodPro & " Tamanho: " & CStr(Len(.sCodPro))
                objTS.WriteLine "Destino: " & .sDestino & " Tamanho: " & CStr(Len(.sDestino))
                objTS.WriteLine "DataInicio: " & CStr(.dtdatainiciovigencia)
                objTS.WriteLine "DataTermino: " & CStr(.dtdatafimvigencia)
            End With
        End If
        
        objTS.Close
     
    End If
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto5(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar ocorrencias

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objOcor As ClassImpCoinfoOcor, sRegContinuacao As String

On Error GoTo Erro_Coinfo_ImportarArqTexto5

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130664
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130665
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_OCOR", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 130666
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpocorrencia", objImpCoinfo.sArqOcorHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError 130672
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqOcorHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqOcorHora = "" Then
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqOcorHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqOcor = lNumIntArqAnt
        gError 130673
    End If
    
    objImpCoinfo.lNumIntArqOcor = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqOcorHora, 1, 0)
    bArqAberto = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 2
'            Exit Do
'        End If
        
        lNumRegs = lNumRegs + 1
        
aqui:
        
        Set objOcor = New ClassImpCoinfoOcor
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objOcor.lnumocorr = StrParaLong(sCampoValor)
                        Case 2
                            objOcor.dtData = StrParaDateCoinfo(sCampoValor)
                        Case 3
                            objOcor.dValor = StrParaDblCoinfo(sCampoValor)
                        Case 4
                            objOcor.scodest = StrParaStrCoinfo(sCampoValor)
                        Case 5
                            objOcor.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 6
                            objOcor.dtdatareg = StrParaDateCoinfo(sCampoValor)
                        Case 7
                            objOcor.lrecnsica = StrParaLong(sCampoValor)
                        Case 8
                            objOcor.sTipVou = StrParaStrCoinfo(sCampoValor)
                        Case 9
                            objOcor.sSerie = StrParaStrCoinfo(sCampoValor)
                        Case 10
                            objOcor.lNumVou = StrParaLong(sCampoValor)
                        Case 11
                            objOcor.idc = IIf(sCampoValor = "yes", 1, 0)
                        Case 12
                            objOcor.lcodemp = StrParaLong(sCampoValor)
                        Case 13
                            If Len(sCampoValor) > 1500 Then MsgBox ("erro stexto")
                            objOcor.stexto1 = left(sCampoValor, 250)
                            If Len(sCampoValor) > 250 Then
                                objOcor.stexto2 = Mid(sCampoValor, 251, 250)
                                If Len(sCampoValor) > 500 Then
                                    objOcor.stexto3 = Mid(sCampoValor, 501, 250)
                                    If Len(sCampoValor) > 750 Then
                                        objOcor.stexto4 = Mid(sCampoValor, 751, 250)
                                        If Len(sCampoValor) > 1000 Then
                                            objOcor.stexto5 = Mid(sCampoValor, 1001, 250)
                                            If Len(sCampoValor) > 1250 Then
                                                objOcor.stexto6 = Mid(sCampoValor, 1251, 250)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Case 14
                            objOcor.iliberado = IIf(sCampoValor = "yes", 1, 0)
                        Case 15
                            objOcor.sControle = StrParaStrCoinfo(sCampoValor)
                        Case 16
                            objOcor.lordem = StrParaLong(sCampoValor)
                        Case 17
                            objOcor.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130667
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 18 Then
            sRegContinuacao = ts.ReadLine
            sRegistro = sRegistro & sRegContinuacao
            lTamanhoDados = lTamanhoDados + Len(sRegContinuacao) + 1 ' + 1 p/'\0x0A'
            GoTo aqui
        End If
        
        objOcor.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objOcor
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoOcorrencia " & _
                "(NumIntArq,NumRegArqTexto,numocorr,Data,valor,codest,usuario,datareg,recnsica,TipVou,Serie, " & _
                "NumVou,dc,CodEmp,texto1,texto2,texto3,texto4,texto5,texto6,liberado," & _
                "Controle, ordem) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .lnumocorr, .dtData, .dValor, .scodest, .sUsuario, .dtdatareg, .lrecnsica, _
                .sTipVou, .sSerie, .lNumVou, .idc, .lcodemp, .stexto1, .stexto2, .stexto3, .stexto4, _
                .stexto5, .stexto6, .iliberado, .sControle, .lordem)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130668
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
'    If iSecao = 1 Then gError 130669
'
'    'leitura da seção de controle
'
'    'Até chegar ao fim do arquivo
'    Do While True
'
'        'Busca o próximo registro do arquivo
'        sRegistro = ts.ReadLine
'
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 3
'            Exit Do
'        End If
'
'        If sRegistro = "filename=fpocorrencia" Then iTestesOk = iTestesOk + 1
'        If Left(sRegistro, 8) = "records=" Then
'            If lNumRegs <> StrParaLong(Mid(sRegistro, 9)) Then gError 130670
'            iTestesOk = iTestesOk + 1
'        End If
'        If Left(sRegistro, 10) = "timestamp=" Then
'            sArqTimeStamp = Mid(sRegistro, 11)
'        End If
'
'    Loop
'
'    If iTestesOk <> 2 Then gError 130671
    
'    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
'    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpocorrencia", sArqTimeStamp, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
'    If lErro <> SUCESSO Then gError 130672
'
'    'se o arquivo já havia sido importado anteriormente
'    If lNumIntArqAnt <> 0 Then
'        objImpCoinfo.lNumIntArqOcor = lNumIntArqAnt
'        gError 130673
'    End If
'
'    objImpCoinfo.lNumIntArqOcor = lNumIntArq
'
'    If iSecao = 2 Then gError 130674
'
'    'leitura da seção com o tamanho dos dados
'
'    'Busca o próximo registro do arquivo
'    sRegistro = ts.ReadLine
'
'    If StrParaLong(sRegistro) <> lTamanhoDados Then gError 130675
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130676
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqOcorHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    If objImpCoinfo.sArqOcorHora <> "" Then
'        Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqOcorHora, True)
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto5 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto5:

    Coinfo_ImportarArqTexto5 = gErr
     
    Select Case gErr
          
        Case 130666, 130672, 53
        
        Case 130673
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqOcorHora <> "" Then
'                Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqOcorHora, True)
            End If
        Case 130668
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOOCORRENCIA", gErr)
        
        Case 130664
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130665
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130676
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130667, 130669, 130670, 130671, 130674, 130675
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_COINFOOCORRENCIA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130563)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto6(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar notas fiscais (cabecalho e itens)

Dim lErro As Long, lTransacao As Long
Dim fso As New FileSystemObject
Dim bTemArq1 As Boolean
Dim bTemArq2 As Boolean
Dim lNumRegsItemNF As Long
Dim lNumRegsNF As Long

On Error GoTo Erro_Coinfo_ImportarArqTexto6

    bTemArq1 = False
    bTemArq2 = False

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130678
    
    lErro = Coinfo_ImportarArqTexto6a(objImpCoinfo, lNumRegsNF)
    If lErro <> SUCESSO And lErro <> 130693 And lErro <> 53 Then gError 130679
    If lErro <> 53 Then bTemArq1 = True
    
    lErro = Coinfo_ImportarArqTexto6b(objImpCoinfo, lNumRegsItemNF)
    If lErro <> SUCESSO And lErro <> 130705 And lErro <> 53 Then gError 130680
    If lErro <> 53 Then bTemArq2 = True
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130681
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqNFHora & " teve " & CStr(lNumRegsNF) & " registros importados com sucesso." & vbNewLine
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqItemNFHora & " teve " & CStr(lNumRegsItemNF) & " registros importados com sucesso." & vbNewLine

    If objImpCoinfo.sArqNFHora <> "" Then
'        If bTemArq1 Then Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqNFHora, True)
    End If
    If objImpCoinfo.sArqItemNFHora <> "" Then
'        If bTemArq2 Then Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqItemNFHora, True)
    End If

    Coinfo_ImportarArqTexto6 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto6:

    Coinfo_ImportarArqTexto6 = gErr
     
    Select Case gErr
          
        Case 130679, 130680
        
        Case 130678
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130681
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130682)
     
    End Select
     
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto6a(ByVal objImpCoinfo As ClassImpCoinfo, lNumRegs As Long) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar cabecalhos das notas fiscais

Dim lErro As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objNF As ClassImpCoinfoNF

On Error GoTo Erro_Coinfo_ImportarArqTexto6a

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130684
        
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_NF", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 130685
        
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpnf", objImpCoinfo.sArqNFHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError 130692
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqNFHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqNFHora = "" Then
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqNFHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqNF = lNumIntArqAnt
        gError 130693
    End If
    
    objImpCoinfo.lNumIntArqNF = lNumIntArq
        
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqNFHora, 1, 0)
    bArqAberto = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 2
'            Exit Do
'        End If
        
        lNumRegs = lNumRegs + 1
        
        Set objNF = New ClassImpCoinfoNF
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objNF.scodest = StrParaStrCoinfo(sCampoValor)
                        Case 2
                            objNF.lNumNF = StrParaLong(sCampoValor)
                        Case 3
                            objNF.dtDataEmi = StrParaDateCoinfo(sCampoValor)
                        Case 4
                            objNF.iCancelado = IIf(sCampoValor = "yes", 1, 0)
                        Case 5
                            objNF.dValor = StrParaDblCoinfo(sCampoValor)
                        Case 6
                            objNF.sDescricao1 = StrParaStrCoinfo(sCampoValor)
                        Case 7
                            objNF.sDescricao2 = StrParaStrCoinfo(sCampoValor)
                        Case 8
                            objNF.sdescricao3 = StrParaStrCoinfo(sCampoValor)
                        Case 9
                            objNF.lcodnat = StrParaLong(sCampoValor)
                        Case 10
                            objNF.sdiscriminacao1 = StrParaStrCoinfo(sCampoValor)
                        Case 11
                            objNF.sdiscriminacao2 = StrParaStrCoinfo(sCampoValor)
                        Case 12
                            objNF.sdiscriminacao3 = StrParaStrCoinfo(sCampoValor)
                        Case 13
                            objNF.sdiscriminacao4 = StrParaStrCoinfo(sCampoValor)
                        Case 14
                            objNF.lcodemp = StrParaLong(sCampoValor)
                        Case 15
                            objNF.idescacre1 = IIf(sCampoValor = "yes", 1, 0)
                        Case 16
                            objNF.idescacre2 = IIf(sCampoValor = "yes", 1, 0)
                        Case 17
                            objNF.idescacre3 = IIf(sCampoValor = "yes", 1, 0)
                        Case 18
                            objNF.dvalcompl1 = StrParaDblCoinfo(sCampoValor)
                        Case 19
                            objNF.dvalcompl2 = StrParaDblCoinfo(sCampoValor)
                        Case 20
                            objNF.dvalcompl3 = StrParaDblCoinfo(sCampoValor)
                        Case 21
                            objNF.imanuauto = IIf(sCampoValor = "yes", 1, 0)
                        Case 22
                            objNF.lcdfilial = StrParaLong(sCampoValor)
                        Case 23
                            objNF.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130686
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 24 Then gError 130687
    
        objNF.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objNF
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoNF (NumIntArq, NumRegArqTexto, codest, numnf, dataemi, cancelado, valor, descricao1, descricao2, descricao3, " & _
            "codnat, discriminacao1, discriminacao2, discriminacao3, discriminacao4, codemp, " & _
            "descacre1 , descacre2, descacre3, valcompl1, valcompl2, valcompl3, manuauto, cdfilial, NoCorporatorEm) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntArq, lNumRegs, .scodest, .lNumNF, .dtDataEmi, .iCancelado, .dValor, .sDescricao1, .sDescricao2, .sdescricao3, _
            .lcodnat, .sdiscriminacao1, .sdiscriminacao2, .sdiscriminacao3, .sdiscriminacao4, .lcodemp, _
            .idescacre1, .idescacre2, .idescacre3, .dvalcompl1, .dvalcompl2, .dvalcompl3, .imanuauto, .lcdfilial, DATA_NULA)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130688
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
'    If iSecao = 1 Then gError 130689
'
'    'leitura da seção de controle
'
'    'Até chegar ao fim do arquivo
'    Do While True
'
'        'Busca o próximo registro do arquivo
'        sRegistro = ts.ReadLine
'
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 3
'            Exit Do
'        End If
'
'        If sRegistro = "filename=fpnf" Then iTestesOk = iTestesOk + 1
'        If Left(sRegistro, 8) = "records=" Then
'            If lNumRegs <> StrParaLong(Mid(sRegistro, 9)) Then gError 130690
'            iTestesOk = iTestesOk + 1
'        End If
'        If Left(sRegistro, 10) = "timestamp=" Then
'            sArqTimeStamp = Mid(sRegistro, 11)
'        End If
'
'    Loop
'
'    If iTestesOk <> 2 Then gError 130691
'
'    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
'    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpnf", sArqTimeStamp, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
'    If lErro <> SUCESSO Then gError 130692
'
'    'se o arquivo já havia sido importado anteriormente
'    If lNumIntArqAnt <> 0 Then
'        objImpCoinfo.lNumIntArqNF = lNumIntArqAnt
'        gError 130693
'    End If
'
'    objImpCoinfo.lNumIntArqNF = lNumIntArq
'
'    If iSecao = 2 Then gError 130694
'
'    'leitura da seção com o tamanho dos dados
'
'    'Busca o próximo registro do arquivo
'    sRegistro = ts.ReadLine
'
'    If StrParaLong(sRegistro) <> lTamanhoDados Then gError 130695
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto6a = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto6a:

    Coinfo_ImportarArqTexto6a = gErr
     
    Select Case gErr
          
        Case 130685, 130692, 130693, 53
        
        Case 130686, 130687, 130689, 130690, 130691, 130694, 130695
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_COINFONF", gErr)
        
        Case 130684
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130688
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFONF", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130683)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto6b(ByVal objImpCoinfo As ClassImpCoinfo, lNumRegs As Long) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar itens das notas fiscais

Dim lErro As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objItemNF As ClassImpCoinfoItemNF

On Error GoTo Erro_Coinfo_ImportarArqTexto6b

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130696
    
    'Mantém a correlação com a nota
    lNumIntArq = objImpCoinfo.lNumIntArqNF

    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpitemnf", objImpCoinfo.sArqItemNFHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError 130704
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqItemNFHora = "" Then
        gError 130705
    End If
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqItemNFHora, 1, 0)
    bArqAberto = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 2
'            Exit Do
'        End If
        
        lNumRegs = lNumRegs + 1
        
        Set objItemNF = New ClassImpCoinfoItemNF
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    sCampoValor = Trim(sCampoValor)
                    
                    Select Case iCampoPos
                    
                        Case 1
                            objItemNF.scodest = StrParaStrCoinfo(sCampoValor)
                        Case 2
                            objItemNF.lNumNF = StrParaLong(sCampoValor)
                        Case 3
                            objItemNF.dValor = StrParaDblCoinfo(sCampoValor)
                        Case 4
                            objItemNF.idescacre1 = IIf(sCampoValor = "yes", 1, 0)
                        Case 5
                            objItemNF.idescacre2 = IIf(sCampoValor = "yes", 1, 0)
                        Case 6
                            objItemNF.idescacre3 = IIf(sCampoValor = "yes", 1, 0)
                        Case 7
                            objItemNF.lNumFat = StrParaLong(sCampoValor)
                        Case 8
                            objItemNF.stexto1 = StrParaStrCoinfo(sCampoValor)
                        Case 9
                            objItemNF.lordem = StrParaLong(sCampoValor)
                        Case 10
                            objItemNF.lcdfilial = StrParaLong(sCampoValor)
                        Case 11
                            objItemNF.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130698
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 12 Then gError 130699
    
        objItemNF.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objItemNF
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoItemNF (NumIntArq, NumRegArqTexto, codest, numnf, valor, descacre1, descacre2, descacre3, numfat, texto1, ordem, cdfilial) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntArq, lNumRegs, .scodest, .lNumNF, .dValor, .idescacre1, .idescacre2, .idescacre3, .lNumFat, .stexto1, .lordem, .lcdfilial)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130700
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
'    If iSecao = 1 Then gError 130701
'
'    'leitura da seção de controle
'
'    'Até chegar ao fim do arquivo
'    Do While True
'
'        'Busca o próximo registro do arquivo
'        sRegistro = ts.ReadLine
'
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 3
'            Exit Do
'        End If
'
'        If sRegistro = "filename=fpitemnf" Then iTestesOk = iTestesOk + 1
'        If Left(sRegistro, 8) = "records=" Then
'            If lNumRegs <> StrParaLong(Mid(sRegistro, 9)) Then gError 130702
'            iTestesOk = iTestesOk + 1
'        End If
'        If Left(sRegistro, 10) = "timestamp=" Then
'            sArqTimeStamp = Mid(sRegistro, 11)
'        End If
'
'    Loop
'
'    If iTestesOk <> 2 Then gError 130703
'
'    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
'    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpitemnf", sArqTimeStamp, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
'    If lErro <> SUCESSO Then gError 130704
'
'    'se o arquivo já havia sido importado anteriormente
'    If lNumIntArqAnt <> 0 Then
''        objImpCoinfo.lNumIntArqItemNF = lNumIntArqAnt
'        gError 130705
'    End If
'
''    objImpCoinfo.lNumIntArqItemNF = lNumIntArq
'
'    If iSecao = 2 Then gError 130706
'
'    'leitura da seção com o tamanho dos dados
'
'    'Busca o próximo registro do arquivo
'    sRegistro = ts.ReadLine
'
'    If StrParaLong(sRegistro) <> lTamanhoDados Then gError 130707
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto6b = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto6b:

    Coinfo_ImportarArqTexto6b = gErr
     
    Select Case gErr
          
        Case 130697, 130704, 130705, 53
        
        Case 130698, 130699, 130701, 130702, 130703, 130706, 130707
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_COINFOITEMNF", gErr)
        
        Case 130696
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130700
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOITEMNF", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130708)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto7(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar tipos de empresa

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objEmissor As ClassImpCoinfoEmissor
Dim objFSO As New FileSystemObject
Dim objTS As TextStream, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto7

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130580
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_EMISSOR", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpemissor", objImpCoinfo.sArqEmissHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqEmissHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqEmissHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqEmissHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqEmiss = lNumIntArqAnt
        gError 130585
    End If
    
    objImpCoinfo.lNumIntArqEmiss = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqEmissHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
        Set objEmissor = New ClassImpCoinfoEmissor
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objEmissor.lCodEmi = StrParaLong(sCampoValor)
                        Case 2
                            objEmissor.sapelido = StrParaStrCoinfo(sCampoValor)
                        Case 3
                            objEmissor.sNome = StrParaStrCoinfo(sCampoValor)
                        Case 4
                            objEmissor.lcodemp = StrParaLong(sCampoValor)
                        Case 5
                            objEmissor.sEmail = StrParaStrCoinfo(sCampoValor)
                        Case 6
                            objEmissor.sObs = left(StrParaStrCoinfo(sCampoValor), 100)
                        Case 7
                            objEmissor.sTipoFat = StrParaStrCoinfo(sCampoValor)
                        Case 8
                            objEmissor.iAtivo = IIf(sCampoValor = "yes", 1, 0)
                        Case 9
                            objEmissor.sUsuario = StrParaStrCoinfo(sCampoValor)
                        Case 10
                            objEmissor.dpercover = StrParaDblCoinfo(sCampoValor)
                        Case 11
                            objEmissor.slogin = StrParaStrCoinfo(sCampoValor)
                        Case 12
                            objEmissor.sSenha = StrParaStrCoinfo(sCampoValor)
                        Case 13
                            objEmissor.ialterar = IIf(sCampoValor = "yes", 1, 0)
                        Case 14
                            objEmissor.dtDataAlt = StrParaDateCoinfo(sCampoValor)
                        Case 15
                            objEmissor.icondpag = StrParaInt(sCampoValor)
                        Case 16
                            objEmissor.lemiger = StrParaLong(sCampoValor)
                        Case 17
                            objEmissor.lemisup = StrParaLong(sCampoValor)
                        Case 18
                            objEmissor.lcodimas = StrParaLong(sCampoValor)
                        Case 19
                            objEmissor.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 20 Then gError 130592
        objEmissor.dtDataExp = StrParaDateCoinfo(sCampoValor)
                
        'inserir registro
        With objEmissor
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoEmissor (NumIntArq,NumRegArqTexto,CodEmi,apelido,nome,codemp,email,obs,TipoFat,Ativo,Usuario,percover,login,senha,alterar,dataalt,condpag,emiger,emisup,codimas,TipoDocCorporator,NumIntDocCorporator,NoCorporatorEm,ComErro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                lNumIntArq, lNumRegs, .lCodEmi, .sapelido, .sNome, .lcodemp, .sEmail, .sObs, .sTipoFat, .iAtivo, .sUsuario, .dpercover, .slogin, .sSenha, .ialterar, .dtDataAlt, .icondpag, .lemiger, .lemisup, .lcodimas, 0, 0, DATA_NULA, 0)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130583
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
           
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqEmissHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto7 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto7:

    Coinfo_ImportarArqTexto7 = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM, 53
        
        Case 130585 'O arquivo já havia sido importado
            'fechar aquivo texto
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqEmissHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqEmissHora)
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOEMISSOR", gErr, objEmissor.lcodemp, objEmissor.lCodEmi, objEmissor.sNome)
        
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130592
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqEmissHora, CStr(lNumRegs))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select

    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqEmissHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
    
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    If 130582 <> gErr And 130584 <> gErr And 53 <> gErr And 130585 <> gErr Then
    
        Set objTS = objFSO.OpenTextFile("C:\SGE_Log_" & objImpCoinfo.sArqEmissHora, ForWriting, True)
    
        objTS.WriteLine "NumReg: " & CStr(lNumRegs)
        objTS.WriteLine "Pos: " & CStr(iCampoPos)
        objTS.WriteLine "ValorCampo : " & sCampoValor
    
        If Not (objEmissor Is Nothing) Then
            With objEmissor
                objTS.WriteLine "Código: " & objEmissor.lCodEmi
                objTS.WriteLine "Apelido: " & objEmissor.sapelido & " Tamanho: " & CStr(Len(objEmissor.sapelido))
            End With
        End If
        
        objTS.Close
     
    End If
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto8(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar tipos de empresa

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objComisOver As ClassImpCoinfoComisOver

On Error GoTo Erro_Coinfo_ImportarArqTexto8

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130580
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_OVER", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 130582
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpover", objImpCoinfo.sArqComisOverHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError 130584
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqComisOverHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqComisOverHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqComisOverHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqComisOver = lNumIntArqAnt
        gError 130585
    End If
    
    objImpCoinfo.lNumIntArqComisOver = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqComisOverHora, 1, 0)
    bArqAberto = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 2
'            Exit Do
'        End If
        
        lNumRegs = lNumRegs + 1
        
        Set objComisOver = New ClassImpCoinfoComisOver
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objComisOver.scodest = Trim(sCampoValor)
                        Case 2
                            objComisOver.dtDatEmi = StrParaDateCoinfo(sCampoValor)
                        Case 3
                            objComisOver.lCodEmi = StrParaLong(sCampoValor)
                        Case 4
                            objComisOver.lNumVou = StrParaLong(sCampoValor)
                        Case 5
                            objComisOver.sTipVou = Trim(sCampoValor)
                        Case 6
                            objComisOver.sSerie = Trim(sCampoValor)
                        Case 7
                            objComisOver.dtdataven = StrParaDateCoinfo(sCampoValor)
                        Case 8
                            objComisOver.dValor = StrParaDblCoinfo(sCampoValor)
                        Case 9
                            objComisOver.iPago = IIf(sCampoValor = "yes", 1, 0)
                        Case 10
                            objComisOver.dtDataPag = StrParaDateCoinfo(sCampoValor)
                        Case 11
                            objComisOver.lrecnsica = StrParaLong(sCampoValor)
                        Case 12
                            objComisOver.lnumpgcm = StrParaLong(sCampoValor)
                        Case 13
                            objComisOver.sMoeda = Trim(sCampoValor)
                        Case 14
                            objComisOver.sTipoFat = Trim(sCampoValor)
                        Case 15
                            objComisOver.ieliminado = IIf(sCampoValor = "yes", 1, 0)
                        Case 16
                            objComisOver.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 17 Then gError 130593
        objComisOver.dtDataExp = StrParaDateCoinfo(sCampoValor)
        
        'inserir registro
        With objComisOver
        
            If Year(.dtDataPag) = 206 Then
                .dtDataPag = StrParaDate(Day(.dtDataPag) & "/" & Month(.dtDataPag) & "/2006")
            End If
            
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoComissoesOver (NumIntArq,NumRegArqTexto,codest,DatEmi,CodEmi,NumVou,TipVou,Serie,dataven,valor,pago,datapag,recnsica,numpgcm,moeda,tipofat,eliminado,TipoDocCorporator,NumIntDocCorporator,NoCorporatorEm,ComErro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                lNumIntArq, lNumRegs, .scodest, .dtDatEmi, .lCodEmi, .lNumVou, .sTipVou, .sSerie, .dtdataven, .dValor, .iPago, .dtDataPag, .lrecnsica, .lnumpgcm, .sMoeda, .sTipoFat, .ieliminado, 0, 0, DATA_NULA, 0)
        End With
        If lErro <> AD_SQL_SUCESSO Then
            gError 130583
        End If
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
       
'    If iSecao = 1 Then gError 130594
'
'    'leitura da seção de controle
'
'    'Até chegar ao fim do arquivo
'    Do While True
'
'        'Busca o próximo registro do arquivo
'         sRegistro = ts.ReadLine
'
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 3
'            Exit Do
'        End If
'
'        If sRegistro = "filename=fptipoempr" Then iTestesOk = iTestesOk + 1
'        If Left(sRegistro, 8) = "records=" Then
'            If lNumRegs <> StrParaLong(Mid(sRegistro, 9)) Then gError 130589
'            iTestesOk = iTestesOk + 1
'        End If
'        If Left(sRegistro, 10) = "timestamp=" Then
'            sArqTimeStamp = Mid(sRegistro, 11)
'        End If
'
'    Loop
'
'    If iTestesOk <> 2 Then gError 130590
    
'    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
'    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fptipoempr", objImpCoinfo.sArqTipoEmprHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
'    If lErro <> SUCESSO Then gError 130584
'
'    'se o arquivo já havia sido importado anteriormente
'    If lNumIntArqAnt <> 0 Then
'        objImpCoinfo.lNumIntArqTipoEmpr = lNumIntArqAnt
'        gError 130585
'    End If
'
'    objImpCoinfo.lNumIntArqTipoEmpr = lNumIntArq
    
'    If iSecao = 2 Then gError 130586
'
'    'leitura da seção com o tamanho dos dados
'
'    'Busca o próximo registro do arquivo
'    sRegistro = ts.ReadLine
'
'    If StrParaLong(sRegistro) <> lTamanhoDados Then gError 130587
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqComisOverHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    If objImpCoinfo.sArqComisOverHora <> "" Then
'        Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqgrupoHora, True)
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto8 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto8:

    Coinfo_ImportarArqTexto8 = gErr
     
    Select Case gErr
          
        Case 130582, 130584, 53
        
        Case 130585 'O arquivo já havia sido importado
            'fechar aquivo texto
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqComisOverHora <> "" Then
'                Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqgrupoHora, True)
            End If
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOGRUPO", gErr)
        
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130590, 130586, 130587, 130592, 130593, 130594
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_COINFOGRUPO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto9(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar tipos de empresa

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objComissao As ClassImpCoinfoComissao

On Error GoTo Erro_Coinfo_ImportarArqTexto9

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130580
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_COMISSAO", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 130582
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "comissao", objImpCoinfo.sArqComissaoHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError 130584
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqComissaoHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqComissaoHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqComissaoHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqComissao = lNumIntArqAnt
        gError 130585
    End If
    
    objImpCoinfo.lNumIntArqComissao = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqComissaoHora, 1, 0)
    bArqAberto = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 2
'            Exit Do
'        End If
        
        lNumRegs = lNumRegs + 1
        
        Set objComissao = New ClassImpCoinfoComissao
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objComissao.sTipVou = Trim(sCampoValor)
                        Case 2
                            objComissao.sSerie = Trim(sCampoValor)
                        Case 3
                            objComissao.lNumVou = StrParaLong(sCampoValor)
                        Case 4
                            objComissao.sTipVou = Trim(sCampoValor)
                        Case 5
                            objComissao.dValBase = StrParaDblCoinfo(sCampoValor)
                        Case 6
                            objComissao.dValComi = StrParaDblCoinfo(sCampoValor)
                        Case 7
                            objComissao.lrecnsica = StrParaLong(sCampoValor)
                        Case 8
                            objComissao.dPercComi = StrParaDblCoinfo(sCampoValor)
                        Case 9
                            objComissao.iExportado = IIf(sCampoValor = "yes", 1, 0)
                        Case 10
                            objComissao.dtDataExp = StrParaDateCoinfo(sCampoValor)
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 11 Then gError 130593
        objComissao.lcodemp = StrParaLong(sCampoValor)
        
        'inserir registro
        With objComissao
            
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoComissao (NumIntArq,NumRegArqTexto,TipVou,Serie,NumVou,tipdoc,valbase,valcomi,recnsica,percomi,TipoDocCorporator,NumIntDocCorporator,NoCorporatorEm,ComErro,CodEmp) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                lNumIntArq, lNumRegs, .sTipVou, .sSerie, .lNumVou, .sTipDoc, .dValBase, .dValComi, .lrecnsica, .dPercComi, 0, 0, DATA_NULA, 0, .lcodemp)
        End With
        If lErro <> AD_SQL_SUCESSO Then
            gError 130583
        End If
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
       
'    If iSecao = 1 Then gError 130594
'
'    'leitura da seção de controle
'
'    'Até chegar ao fim do arquivo
'    Do While True
'
'        'Busca o próximo registro do arquivo
'         sRegistro = ts.ReadLine
'
'        'se "." entao trocou de seção
'        If sRegistro = "." Then
'            iSecao = 3
'            Exit Do
'        End If
'
'        If sRegistro = "filename=fptipoempr" Then iTestesOk = iTestesOk + 1
'        If Left(sRegistro, 8) = "records=" Then
'            If lNumRegs <> StrParaLong(Mid(sRegistro, 9)) Then gError 130589
'            iTestesOk = iTestesOk + 1
'        End If
'        If Left(sRegistro, 10) = "timestamp=" Then
'            sArqTimeStamp = Mid(sRegistro, 11)
'        End If
'
'    Loop
'
'    If iTestesOk <> 2 Then gError 130590
    
'    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
'    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fptipoempr", objImpCoinfo.sArqTipoEmprHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
'    If lErro <> SUCESSO Then gError 130584
'
'    'se o arquivo já havia sido importado anteriormente
'    If lNumIntArqAnt <> 0 Then
'        objImpCoinfo.lNumIntArqTipoEmpr = lNumIntArqAnt
'        gError 130585
'    End If
'
'    objImpCoinfo.lNumIntArqTipoEmpr = lNumIntArq
    
'    If iSecao = 2 Then gError 130586
'
'    'leitura da seção com o tamanho dos dados
'
'    'Busca o próximo registro do arquivo
'    sRegistro = ts.ReadLine
'
'    If StrParaLong(sRegistro) <> lTamanhoDados Then gError 130587
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqComissaoHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    If objImpCoinfo.sArqComissaoHora <> "" Then
'        Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqComissaoHora, True)
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto9 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto9:

    Coinfo_ImportarArqTexto9 = gErr
     
    Select Case gErr
          
        Case 130582, 130584, 53
        
        Case 130585 'O arquivo já havia sido importado
            'fechar aquivo texto
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqComissaoHora <> "" Then
'                Call fso.DeleteFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqComissaoHora, True)
            End If
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOGRUPO", gErr)
        
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130590, 130586, 130587, 130592, 130593, 130594
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_COINFOGRUPO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function CoInfo_Adiciona_Log(ByVal sCritica As String, ByVal objImpCoinfoLogAtu As ClassImpCoinfoLogAtu, colImpCoinfoLogAtu As Collection, Optional ByVal bAviso As Boolean = False) As Long

Dim objImpCoinfoLogAtuNovo As New ClassImpCoinfoLogAtu
Dim iSeq As Integer

    iSeq = 0
    For Each objImpCoinfoLogAtuNovo In colImpCoinfoLogAtu
        If objImpCoinfoLogAtuNovo.iTabela = objImpCoinfoLogAtu.iTabela And _
            objImpCoinfoLogAtuNovo.lNumIntArq = objImpCoinfoLogAtu.lNumIntArq And _
            objImpCoinfoLogAtuNovo.lNumRegArqTexto = objImpCoinfoLogAtu.lNumRegArqTexto Then
            If iSeq < objImpCoinfoLogAtuNovo.iSeq Then
                iSeq = objImpCoinfoLogAtuNovo.iSeq
            End If
        End If
    Next
    
    objImpCoinfoLogAtuNovo.iSeq = iSeq + 1
    objImpCoinfoLogAtuNovo.iTabela = objImpCoinfoLogAtu.iTabela
    objImpCoinfoLogAtuNovo.lNumIntArq = objImpCoinfoLogAtu.lNumIntArq
    objImpCoinfoLogAtuNovo.lNumRegArqTexto = objImpCoinfoLogAtu.lNumRegArqTexto
    objImpCoinfoLogAtuNovo.sCritica = sCritica
    
    If bAviso Then
        objImpCoinfoLogAtuNovo.iAviso = MARCADO
    End If
    
    colImpCoinfoLogAtu.Add objImpCoinfoLogAtuNovo

End Function

Public Function CoInfoLogAtualizacao_Insere(ByVal colImpCoinfoLogAtu As Collection, iComErro As Integer, ByVal objImpCoInfoLog As ClassImpCoinfoLogAtu, Optional ByVal bNaoInsereNada As Boolean = False) As Long
'Grava o Log

Dim lErro As Long
Dim iIndice As Integer
Dim objImpCoinfoLogAtu As ClassImpCoinfoLogAtu
Dim alComandoExt(1 To 3) As Long
Dim lTransacaoExt As Long

On Error GoTo Erro_CoInfoLogAtualizacao_Insere

    'Abre o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao2)
        If alComandoExt(iIndice) = 0 Then gError 182690
    Next
    
    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao2)
    If lTransacaoExt = 0 Then gError 182691
    
    'Exclui o último log para o registro
    lErro = CoInfoLogAtualizacao_Exclui(alComandoExt(1), alComandoExt(2), objImpCoInfoLog)
    If lErro <> SUCESSO Then gError 182692
    
    iComErro = DESMARCADO

    'Para cada mensagem
    For Each objImpCoinfoLogAtu In colImpCoinfoLogAtu
    
        If objImpCoinfoLogAtu.iAviso = DESMARCADO Then iComErro = MARCADO
        
        TelaAcompanhaBatchFAT.Log = objImpCoinfoLogAtu.sCritica & vbNewLine

        lErro = Comando_Executar(alComandoExt(3), "INSERT INTO CoInfoLogAtualizacao( NumIntArq, Tabela, NumRegArqTexto, Seq, Critica, Aviso, Data) VALUES (?,?,?,?,?,?,?)", _
                objImpCoinfoLogAtu.lNumIntArq, objImpCoinfoLogAtu.iTabela, objImpCoinfoLogAtu.lNumRegArqTexto, objImpCoinfoLogAtu.iSeq, objImpCoinfoLogAtu.sCritica, objImpCoinfoLogAtu.iAviso, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 182536
        
    Next
        
    If iComErro = DESMARCADO And Not bNaoInsereNada Then
    
        Set objImpCoinfoLogAtu = objImpCoInfoLog
        objImpCoinfoLogAtu.sCritica = "Registro atualizado com Sucesso."
        objImpCoinfoLogAtu.iSeq = colImpCoinfoLogAtu.Count + 1
    
        TelaAcompanhaBatchFAT.Log = objImpCoinfoLogAtu.sCritica & vbNewLine
    
        lErro = Comando_Executar(alComandoExt(3), "INSERT INTO CoInfoLogAtualizacao( NumIntArq, Tabela, NumRegArqTexto, Seq, Critica, Aviso, Data) VALUES (?,?,?,?,?,?,?)", _
                objImpCoinfoLogAtu.lNumIntArq, objImpCoinfoLogAtu.iTabela, objImpCoinfoLogAtu.lNumRegArqTexto, objImpCoinfoLogAtu.iSeq, objImpCoinfoLogAtu.sCritica, MARCADO, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182536
    
    End If
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 182693
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    CoInfoLogAtualizacao_Insere = SUCESSO

    Exit Function

Erro_CoInfoLogAtualizacao_Insere:

    CoInfoLogAtualizacao_Insere = gErr

    Select Case gErr

        Case 182536
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOLOGATUALIZACAO", gErr)
            
        Case 182690
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 182691
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 182692
        
        Case 182693
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182537)

    End Select
    
    Call Transacao_RollbackExt(lTransacaoExt)
    
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next

    Exit Function

End Function

Public Function CoInfoLogAtualizacao_Exclui(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objImpCoinfoLogAtu As ClassImpCoinfoLogAtu) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CoInfoLogAtualizacao_Exclui
    
    'Le a tabelaCoInfoLogAtualizacao
    lErro = Comando_ExecutarPos(lComando1, "SELECT 1 FROM CoInfoLogAtualizacao WHERE NumIntArq= ?  AND Tabela= ?  AND NumRegArqTexto= ? ", _
                0, iAux, objImpCoinfoLogAtu.lNumIntArq, objImpCoinfoLogAtu.iTabela, objImpCoinfoLogAtu.lNumRegArqTexto)
    If lErro <> AD_SQL_SUCESSO Then gError 182538

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182539

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> SUCESSO Then gError 182540
    
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM CoInfoLogAtualizacao", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 182541
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CoInfoLogAtualizacao_Exclui = SUCESSO

    Exit Function

Erro_CoInfoLogAtualizacao_Exclui:

    CoInfoLogAtualizacao_Exclui = gErr

    Select Case gErr
        
        Case 182538, 182539, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOLOGATUALIZACAO", gErr)

        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_COINFOLOGATUALIZACAO", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COINFOLOGATUALIZACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182542)

    End Select
    
    Exit Function

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++
'Só para testes
Private Function Limpa_Tabelas()

'Dim lErro As Long
'Dim iIndice As Integer
'Dim objNF As ClassNFiscal
'Dim objTitPag As ClassTituloPagar
'Dim objTitRec As ClassTituloReceber
'
'    For iIndice = 1 To 2781
'
'        Set objNF = New ClassNFiscal
'
'        objNF.lNumIntDoc = iIndice
'
'        lErro = CF("NFiscal_Le", objNF)
'        If lErro = SUCESSO Then
'            Call CF("NotaFiscalSaida_Excluir", objNF, Nothing)
'        End If
'
'    Next
'
'    For iIndice = 182 To 181
'
'        Set objTitPag = New ClassTituloPagar
'
'        objTitPag.lNumIntDoc = iIndice
'
'        lErro = CF("TituloPagar_Le", objTitPag)
'        If lErro = SUCESSO Then
'            Call CF("TituloPagar_Exclui", objTitPag, Nothing)
'        End If
'
'    Next
'
'    For iIndice = 2977 To 2976
'
'        Set objTitRec = New ClassTituloReceber
'
'        objTitRec.lNumIntDoc = iIndice
'
'        lErro = CF("TituloReceber_Le", objTitRec)
'        If lErro = SUCESSO Then
'            Call CF("TituloReceber_Exclui", objTitRec, Nothing)
'        End If
'
'    Next
                
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++

''ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
''LEITURA
'Public Function Coinfo_AtualizarItensFatu(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim colItensFatu As Collection
'Dim colItensAlterados As Collection
'Dim colJaAtualizado As Collection
'Dim alComando(1 To 9) As Long
'Dim alComandoExt(0 To 4) As Long
'Dim lTransacao As Long
'Dim lTransacaoExt As Long
'Dim sCiaCartAnt As String
'Dim dtDataDepAnt As Date
'Dim iQuantParcAnt As Integer
'Dim objImpCoinfoItemFat As ClassImpCoinfoItemFat
'Dim objImpCoinfoItemFatAux As ClassImpCoinfoItemFat
'Dim objImpCoinfoCartoes As ClassImpCoinfoCartoes
'Dim colCartoes As New Collection
'Dim colPerdas As New Collection
'Dim lNumIntArq As Long
'Dim objImpCoInfoLog As ClassImpCoinfoLogAtu
'Dim colImpCoInfoLog As Collection
'Dim lNumTitulo As Long
'Dim bFazParte As Boolean
'Dim iComErro As Integer
'Dim bErro As Boolean
'Dim bTeveErro As Boolean
'Dim iTotalRegCertos As Integer
'Dim iTotalRegErr As Integer
'Dim iFilialEmpresaAnt As Integer
'Dim bAchou As Boolean
'
'On Error GoTo Erro_Coinfo_AtualizarItensFatu
'
'    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpitemfatu")
'    If lErro <> SUCESSO Then gError 182714
'
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 182593
'
'    'Executa a abertura do Comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
'        If alComandoExt(iIndice) = 0 Then gError 182594
'    Next
'
'    'Executa a abertura do Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 182595
'    Next
'
'    bTeveErro = False
'
'    'Le os arquivos não atualizados
'    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
'        lNumIntArq, "fpitemfatu", DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 182596
'
'    'Busca o Primeiro
'    lErro = Comando_BuscarProximo(alComandoExt(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182597
'
'    'Enquando existirem arquivos
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set colItensFatu = New Collection
'        Set colItensAlterados = New Collection
'        Set colJaAtualizado = New Collection
'        Set colCartoes = New Collection
'
'        lErro = Coinfo_AtualizarItensFatu2(lNumIntArq, alComandoExt(1), alComandoExt(4), objImpCoinfo, colItensFatu, colItensAlterados, colJaAtualizado)
'        If lErro <> SUCESSO Then gError 182598
'
'        sCiaCartAnt = ""
'        iQuantParcAnt = 0
'        dtDataDepAnt = DATA_NULA
'        iFilialEmpresaAnt = 0
'
'        For Each objImpCoinfoItemFat In colItensFatu
'
'            If objImpCoinfoItemFat.sCiaCart <> "" And objImpCoinfoItemFat.iQuantParc <> 0 And objImpCoinfoItemFat.dtData <> DATA_NULA And objImpCoinfoItemFat.iFilialEmpresa <> 0 Then
'
'                'If objImpCoinfoItemFat.sCiaCart <> sCiaCartAnt Or objImpCoinfoItemFat.iQuantParc <> iQuantParcAnt Or objImpCoinfoItemFat.dtData <> dtDataDepAnt Or iFilialEmpresaAnt <> objImpCoinfoItemFat.iFilialEmpresa Then
'
'                    Set objImpCoinfoCartoes = New ClassImpCoinfoCartoes
'
'                    objImpCoinfoCartoes.sCiaCart = objImpCoinfoItemFat.sCiaCart
'                    objImpCoinfoCartoes.iQuantParc = objImpCoinfoItemFat.iQuantParc
'                    objImpCoinfoCartoes.dtDataDep = objImpCoinfoItemFat.dtData
'                    objImpCoinfoCartoes.iFilialEmpresa = objImpCoinfoItemFat.iFilialEmpresa
'
'                    colCartoes.Add objImpCoinfoCartoes
'
''                    sCiaCartAnt = objImpCoinfoItemFat.sCiaCart
''                    iQuantParcAnt = objImpCoinfoItemFat.iQuantParc
''                    dtDataDepAnt = objImpCoinfoItemFat.dtData
''                    iFilialEmpresaAnt = objImpCoinfoItemFat.iFilialEmpresa
''
''                End If
'
'                objImpCoinfoCartoes.dValor = objImpCoinfoCartoes.dValor + objImpCoinfoItemFat.dVlCartao
'
'                objImpCoinfoCartoes.colItens.Add objImpCoinfoItemFat
'
'            Else
'
'                'Verifica se é uma alteração
'                bAchou = False
'                For Each objImpCoinfoItemFatAux In colItensAlterados
'                    If objImpCoinfoItemFatAux.lcodemp = objImpCoinfoItemFat.lcodemp And objImpCoinfoItemFatAux.lrecnsica = objImpCoinfoItemFat.lrecnsica Then
'                        bAchou = True
'                        Exit For
'                    End If
'                Next
'
'                'Se não, marca como já atualizado
'                If Not bAchou Then
'                    colJaAtualizado.Add objImpCoinfoItemFat
'                End If
'
'            End If
'
'        Next
'
'        For Each objImpCoinfoCartoes In colCartoes
'
'            bErro = False
'
'            Set colImpCoInfoLog = New Collection
'            Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
'
'            objImpCoInfoLog.iTabela = TABELA_COINFOFATURA
'            objImpCoInfoLog.lNumIntArq = lNumIntArq
'
'             'Abertura de transação
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182599
'
'            For Each objImpCoinfoItemFat In objImpCoinfoCartoes.colItens
'                'Insere o Log vazio para apagar o que foi feito anteriormente
'                objImpCoInfoLog.lNumRegArqTexto = objImpCoinfoItemFat.lNumRegArqTexto
'                lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog, True)
'                If lErro <> SUCESSO Then gError 182544
'            Next
'
'            objImpCoinfoCartoes.lNumTitulo = 0
'            lErro = Coinfo_GerarTitRec_Cartao(objImpCoinfoCartoes)
'            If lErro <> SUCESSO Then 'gError 182600
'                bErro = True
'
'                For Each objImpCoinfoItemFat In objImpCoinfoCartoes.colItens
'                    objImpCoInfoLog.lNumRegArqTexto = objImpCoinfoItemFat.lNumRegArqTexto
'                    Call CoInfo_Adiciona_Log("Não foi possível gerar o Título a Receber para Data " & Format(objImpCoinfoCartoes.dtDataDep, "dd/mm/yyyy") & " Cia " & objImpCoinfoCartoes.sCiaCart & " e Quantidade de Parcelas " & CStr(objImpCoinfoCartoes.iQuantParc), objImpCoInfoLog, colImpCoInfoLog)
'                Next
'            Else
'                For Each objImpCoinfoItemFat In objImpCoinfoCartoes.colItens
'                    objImpCoInfoLog.lNumRegArqTexto = objImpCoinfoItemFat.lNumRegArqTexto
'                    Call CoInfo_Adiciona_Log("Foi gerado o Título a Receber Tipo " & TIPODOC_CARTAO_CRED_DEB & " número " & CStr(objImpCoinfoCartoes.lNumTitulo) & " para Data " & Format(objImpCoinfoCartoes.dtDataDep, "dd/mm/yyyy") & " Cia " & objImpCoinfoCartoes.sCiaCart & " e Quantidade de Parcelas " & CStr(objImpCoinfoCartoes.iQuantParc), objImpCoInfoLog, colImpCoInfoLog, True)
'                Next
'            End If
'
'            If Not bErro Then
'
'                For iIndice = colItensAlterados.Count To 1 Step -1
'
'                    Set objImpCoinfoItemFat = colItensAlterados.Item(iIndice)
'
'                    objImpCoInfoLog.lNumRegArqTexto = objImpCoinfoItemFat.lNumRegArqTexto
'
'                    bFazParte = False
'                    For Each objImpCoinfoItemFatAux In objImpCoinfoCartoes.colItens
'                        If objImpCoinfoItemFat.lrecnsica = objImpCoinfoItemFatAux.lrecnsica And _
'                        objImpCoinfoItemFat.lcodemp = objImpCoinfoItemFatAux.lcodemp Then
'                            bFazParte = True
'                            Exit For
'                        End If
'                    Next
'
'                    'Se esse registro de alteração faz parte de algum dos que vão ser incluidos
'                    If bFazParte Then
'
'                        lErro = Coinfo_Gerar_BaixaParcRec(objImpCoinfoItemFat.dValor, objImpCoinfoItemFat.lNumIntTitulo, lNumTitulo, objImpCoinfoItemFat.sCiaCart, objImpCoinfoItemFat.iQuantParc, gdtDataAtual)
'                        If lErro <> SUCESSO Then
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Não foi possível gerar uma baixa por perda para o Título " & CStr(lNumTitulo) & " decorrente de um cancelamento de pagamento em cartão.", objImpCoInfoLog, colImpCoInfoLog)
'                        Else
'                            Call CoInfo_Adiciona_Log("Foi gerada uma baixa por perda para o Título " & CStr(lNumTitulo) & " decorrente de um cancelamento de pagamento em cartão.", objImpCoInfoLog, colImpCoInfoLog, True)
'                        End If
'
'                        colItensAlterados.Remove iIndice
'                    End If
'
'                Next
'
'            End If
'
'            objImpCoInfoLog.lNumRegArqTexto = 0
'
'            'Insere o Log de erros
'            lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
'            If lErro <> SUCESSO Then gError 182544
'
'            'Atualiza a nota para informar que ela foi importada com sucesso
'            If iComErro = MARCADO Then
'
'                bTeveErro = True
'
'                'Desfaz Transação para canecalr as coisas pela metade
'                Call Transacao_Rollback
'
'                'Abertura de transação só para marcar que houve erro
'                lTransacao = Transacao_Abrir()
'                If lTransacao = 0 Then gError 182525
'
'            End If
'
'            lErro = Coinfo_AtualizarItensFatu3(alComando(1), alComando(2), alComando(3), objImpCoinfoCartoes, iComErro, iTotalRegCertos, iTotalRegErr)
'            If lErro <> SUCESSO Then gError 182601
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 182602
'
''            lErro = Rotina_Atualizar_AtualizaTelaBatch
''            If lErro <> SUCESSO Then gError 182720
'
'        Next
'
'        For Each objImpCoinfoItemFat In colJaAtualizado
'
'            Set objImpCoinfoCartoes = New ClassImpCoinfoCartoes
'
'            objImpCoinfoCartoes.colItens.Add objImpCoinfoItemFat
'
'            'Abertura de transação só para marcar que houve erro
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182525
'
'            lErro = Coinfo_AtualizarItensFatu3(alComando(7), alComando(8), alComando(9), objImpCoinfoCartoes, DESMARCADO, iTotalRegCertos, iTotalRegErr)
'            If lErro <> SUCESSO Then gError 182601
'
''            lErro = Rotina_Atualizar_AtualizaTelaBatch
''            If lErro <> SUCESSO Then gError 182720
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 182602
'
'        Next
'
'        'Trata as perdas que não tiveram geração de título a receber nessa leva
'        For Each objImpCoinfoItemFat In colItensAlterados
'
'            bErro = False
'
'            Set colImpCoInfoLog = New Collection
'            Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
'
'            objImpCoInfoLog.iTabela = TABELA_COINFOFATURA
'            objImpCoInfoLog.lNumIntArq = lNumIntArq
'            objImpCoInfoLog.lNumRegArqTexto = objImpCoinfoItemFat.lNumRegArqTexto
'
'            Set objImpCoinfoCartoes = New ClassImpCoinfoCartoes
'
'            objImpCoinfoCartoes.colItens.Add objImpCoinfoItemFat
'
'            'Abertura de transação
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182525
'
'            lErro = Coinfo_Gerar_BaixaParcRec(objImpCoinfoItemFat.dValor, objImpCoinfoItemFat.lNumIntTitulo, lNumTitulo, objImpCoinfoItemFat.sCiaCart, objImpCoinfoItemFat.iQuantParc, gdtDataAtual)
'            If lErro <> SUCESSO Then
'                bErro = True
'                Call CoInfo_Adiciona_Log("Não foi possível gerar uma baixa por perda para o Título " & CStr(lNumTitulo) & " decorrente de um cancelamento de pagamento em cartão.", objImpCoInfoLog, colImpCoInfoLog)
'            Else
'                Call CoInfo_Adiciona_Log("Foi gerada uma baixa por perda para o Título " & CStr(lNumTitulo) & " decorrente de um cancelamento de pagamento em cartão.", objImpCoInfoLog, colImpCoInfoLog, True)
'            End If
'
'            'Insere o Log de erros
'            lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
'            If lErro <> SUCESSO Then gError 182544
'
'            'Atualiza a nota para informar que ela foi importada com sucesso
'            If iComErro = MARCADO Then
'
'                bTeveErro = True
'
'                'Desfaz Transação para canecalr as coisas pela metade
'                Call Transacao_Rollback
'
'                'Abertura de transação só para marcar que houve erro
'                lTransacao = Transacao_Abrir()
'                If lTransacao = 0 Then gError 182525
'
'            End If
'
'            lErro = Coinfo_AtualizarItensFatu3(alComando(4), alComando(5), alComando(6), objImpCoinfoCartoes, iComErro, iTotalRegCertos, iTotalRegErr)
'            If lErro <> SUCESSO Then gError 182601
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 182602
'
''            lErro = Rotina_Atualizar_AtualizaTelaBatch
''            If lErro <> SUCESSO Then gError 182720
'
'        Next
'
'        If bTeveErro Then Exit Do
'
'        'Atualiza o arquivo
'        lErro = Comando_ExecutarPos(alComandoExt(3), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComandoExt(0), gdtDataHoje)
'        If lErro <> SUCESSO Then gError 182606
'
'        'Busca o Primeiro
'        lErro = Comando_BuscarProximo(alComandoExt(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182603
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacaoExt)
'    If lErro <> AD_SQL_SUCESSO Then gError 182604
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Receber de Cartões foram atualizadas. " & CStr(iTotalRegCertos) & " registros com sucesso e " & CStr(iTotalRegErr) & " registros com erro." & vbNewLine
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'Fecha o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    Coinfo_AtualizarItensFatu = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarItensFatu:
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Receber de Cartões foram atualizadas. " & CStr(iTotalRegCertos) & " registros com sucesso e " & CStr(iTotalRegErr) & " registros com erro." & vbNewLine
'
'    Coinfo_AtualizarItensFatu = gErr
'
'    Select Case gErr
'
'        Case 182593, 182599
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182594, 182595
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'       Case 182596, 182597, 182603
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 182598, 182600, 182601
'
'        Case 182602, 182604
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 182606
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOARQSLIDOS", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182605)
'
'    End Select
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'Fecha o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    'Desfaz Transação
'    Call Transacao_Rollback
'    Call Transacao_RollbackExt(lTransacaoExt)
'
'    Exit Function
'
'End Function
'
'Public Function Coinfo_AtualizarItensFatu2(ByVal lNumIntArq As Long, ByVal lComando As Long, ByVal lComando2 As Long, ByVal objImpCoinfo As ClassImpCoinfo, ByVal colImpCoinfoItemFat As Collection, ByVal colItensAlterados As Collection, ByVal colJaAtualizado As Collection) As Long
'
'Dim lErro As Long
'Dim tImpCoinfoItemFat As typeImpCoinfoItemFat
'Dim tImpCoinfoItemFatAux As typeImpCoinfoItemFat
'Dim objImpCoinfoItemFat As ClassImpCoinfoItemFat
'Dim bNovo As Boolean
'Dim bExclusao As Boolean
'Dim bAlteracao As Boolean
'Dim bColEsta As Boolean
'Dim iFilialEmpCorp As Integer
'Dim iFilialEmpCoinfo As Integer
'Dim colSaida As New Collection
'Dim colCampos As New Collection
'Dim iIndice As Integer
'
'On Error GoTo Erro_Coinfo_AtualizarItensFatu2
'
'    'Alocação de espaço no buffer
'    tImpCoinfoItemFat.sCiaCart = String(250, 0)
'
'    'Le a tabelaCoinfoItemFatura
'    lErro = Comando_Executar(lComando, "SELECT NumIntArq, NumRegArqTexto, QuantParc, VlCartao, " & _
'                "CiaCart, Data, recnsica, codemp, Excluido FROM CoinfoItemFatura WHERE NumIntArq= ? AND NoCorporatorEm = ? AND codest = ? AND Cartao = ? AND Data > ? AND Data <= ? ORDER BY CiaCart, DataDep, QuantParc ", _
'                tImpCoinfoItemFat.lNumIntArq, tImpCoinfoItemFat.lNumRegArqTexto, tImpCoinfoItemFat.iQuantParc, _
'                tImpCoinfoItemFat.dVlCartao, tImpCoinfoItemFat.sCiaCart, tImpCoinfoItemFat.dtData, _
'                tImpCoinfoItemFat.lrecnsica, tImpCoinfoItemFat.lcodemp, tImpCoinfoItemFat.iExcluido, lNumIntArq, DATA_NULA, objImpCoinfo.sSiglaEmpresa, MARCADO, #6/30/2006#, #10/31/2006#)
'    If lErro <> AD_SQL_SUCESSO Then gError 182580
'
'    'Busca Primeiro
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182581
'
'    'Sem Dados
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        bNovo = False
'        bExclusao = False
'        bAlteracao = False
'
'        lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", tImpCoinfoItemFat.lcodemp, iFilialEmpCorp, iFilialEmpCoinfo)
'        If lErro <> SUCESSO Then gError 182744
'
'        tImpCoinfoItemFatAux.sCiaCart = String(250, 0)
'
'        'Busca a última atualização feita para o registro
'        lErro = Comando_Executar(lComando2, "SELECT NumIntArq, NumRegArqTexto, NumIntTitulo, QuantParc, VlCartao, CiaCart, Data, Excluido FROM CoinfoItemFaturaImport WHERE  recnsica = ? AND codemp = ? AND NumIntArq <= ? ", _
'        tImpCoinfoItemFatAux.lNumIntArq, tImpCoinfoItemFatAux.lNumRegArqTexto, tImpCoinfoItemFatAux.lNumIntTitulo, tImpCoinfoItemFatAux.iQuantParc, tImpCoinfoItemFatAux.dVlCartao, tImpCoinfoItemFatAux.sCiaCart, tImpCoinfoItemFatAux.dtData, tImpCoinfoItemFatAux.iExcluido, _
'        tImpCoinfoItemFat.lrecnsica, tImpCoinfoItemFat.lcodemp, lNumIntArq)
'        If lErro <> AD_SQL_SUCESSO Then gError 182580
'
'        'Busca Primeiro
'        lErro = Comando_BuscarPrimeiro(lComando2)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182581
'
'        If lErro = AD_SQL_SUCESSO Then
'
'            bNovo = False
'
'            'Se ele foi excluido
'            If tImpCoinfoItemFat.iExcluido = MARCADO Then
'
'                'Se o registro já tinha sido excluido antes
'                If tImpCoinfoItemFatAux.iExcluido = MARCADO Then
'                    bExclusao = False
'                Else
'                    bExclusao = True
'                End If
'
'            End If
'
'            'Se mudou alguma informação relevante
'            If tImpCoinfoItemFatAux.iQuantParc <> tImpCoinfoItemFat.iQuantParc Or _
'                tImpCoinfoItemFatAux.dVlCartao <> tImpCoinfoItemFat.dVlCartao Or _
'                tImpCoinfoItemFatAux.sCiaCart <> tImpCoinfoItemFat.sCiaCart Or _
'                tImpCoinfoItemFatAux.dtData <> tImpCoinfoItemFat.dtData Then
'
'                 bAlteracao = True
'            Else
'                 bAlteracao = False
'            End If
'
'        Else
'            bNovo = True
'            bAlteracao = False
'            If tImpCoinfoItemFat.iExcluido = MARCADO Then
'                bExclusao = True
'            Else
'                bExclusao = False
'            End If
'
'        End If
'
'        bColEsta = False
'
'        'Se é um novo registro e não estiver marcado como excluido
'        'ou se um registro com alguma alteração relevante
'        If (bNovo And Not bExclusao) Or bAlteracao Then
'
'            Set objImpCoinfoItemFat = New ClassImpCoinfoItemFat
'
'            objImpCoinfoItemFat.lrecnsica = tImpCoinfoItemFat.lrecnsica
'            objImpCoinfoItemFat.lcodemp = tImpCoinfoItemFat.lcodemp
'
'            objImpCoinfoItemFat.lNumIntArq = tImpCoinfoItemFat.lNumIntArq
'            objImpCoinfoItemFat.lNumRegArqTexto = tImpCoinfoItemFat.lNumRegArqTexto
'            objImpCoinfoItemFat.iQuantParc = tImpCoinfoItemFat.iQuantParc
'            objImpCoinfoItemFat.dVlCartao = tImpCoinfoItemFat.dVlCartao
'            objImpCoinfoItemFat.sCiaCart = tImpCoinfoItemFat.sCiaCart
'            objImpCoinfoItemFat.dtData = tImpCoinfoItemFat.dtData
'            objImpCoinfoItemFat.iFilialEmpresa = iFilialEmpCorp
'
'            colImpCoinfoItemFat.Add objImpCoinfoItemFat
'
'            bColEsta = True
'
'        End If
'
'        'Se é um registro já existente com uma exclusão atual ou
'        'Se é um registro existente com alguma alteração
'        If (Not bNovo And bExclusao) Or bAlteracao Then
'
'            Set objImpCoinfoItemFat = New ClassImpCoinfoItemFat
'
'            objImpCoinfoItemFat.lrecnsica = tImpCoinfoItemFat.lrecnsica
'            objImpCoinfoItemFat.lcodemp = tImpCoinfoItemFat.lcodemp
'
'            objImpCoinfoItemFat.lNumIntArq = tImpCoinfoItemFat.lNumIntArq
'            objImpCoinfoItemFat.lNumRegArqTexto = tImpCoinfoItemFat.lNumRegArqTexto
'            objImpCoinfoItemFat.iQuantParc = tImpCoinfoItemFatAux.iQuantParc
'            objImpCoinfoItemFat.dVlCartao = tImpCoinfoItemFatAux.dVlCartao
'            objImpCoinfoItemFat.sCiaCart = tImpCoinfoItemFatAux.sCiaCart
'            objImpCoinfoItemFat.dtData = tImpCoinfoItemFatAux.dtData
'            objImpCoinfoItemFat.lNumIntTitulo = tImpCoinfoItemFatAux.lNumIntTitulo
'            objImpCoinfoItemFat.iFilialEmpresa = iFilialEmpCorp
'
'            colItensAlterados.Add objImpCoinfoItemFat
'
'            bColEsta = True
'
'        End If
'
'        If Not bColEsta Then
'
'            Set objImpCoinfoItemFat = New ClassImpCoinfoItemFat
'
'            objImpCoinfoItemFat.lrecnsica = tImpCoinfoItemFat.lrecnsica
'            objImpCoinfoItemFat.lcodemp = tImpCoinfoItemFat.lcodemp
'            objImpCoinfoItemFat.iFilialEmpresa = iFilialEmpCorp
'
'            objImpCoinfoItemFat.lNumIntArq = tImpCoinfoItemFat.lNumIntArq
'            objImpCoinfoItemFat.lNumRegArqTexto = tImpCoinfoItemFat.lNumRegArqTexto
'
'            colJaAtualizado.Add objImpCoinfoItemFat
'
'        End If
'
'        'Busca Próximo
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182582
'
'    Loop
'
'    'Reordena por causa da Filial Empresa para poder grupar corretamente
'    colCampos.Add "dtData"
'    colCampos.Add "sCiaCart"
'    colCampos.Add "iQuantParc"
'    colCampos.Add "iFilialEmpresa"
'
'    lErro = Ordena_Colecao(colImpCoinfoItemFat, colSaida, colCampos)
'    If lErro <> SUCESSO Then gError 182745
'
'    For iIndice = colImpCoinfoItemFat.Count To 1 Step -1
'        colImpCoinfoItemFat.Remove iIndice
'    Next
'
'    For Each objImpCoinfoItemFat In colSaida
'        colImpCoinfoItemFat.Add objImpCoinfoItemFat
'    Next
'
'    Coinfo_AtualizarItensFatu2 = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarItensFatu2:
'
'    Coinfo_AtualizarItensFatu2 = gErr
'
'    Select Case gErr
'
'        Case 182580, 182581, 182582
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMFATURA", gErr)
'
'        Case 182744, 182745
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182583)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Coinfo_AtualizarItensFatu3(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objImpCoinfoCartoes As ClassImpCoinfoCartoes, ByVal iComErro As Integer, iTotalRegCertos As Integer, iTotalRegErr As Integer) As Long
''Tem que estar dentro de transação
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim objImpCoinfoItemFat As ClassImpCoinfoItemFat
'Dim iAux As Integer
'Dim alComando(1 To 3) As Long
'
'On Error GoTo Erro_Coinfo_AtualizarItensFatu3
'
'    For Each objImpCoinfoItemFat In objImpCoinfoCartoes.colItens
'
'        'Executa a abertura do Comando
'        For iIndice = LBound(alComando) To UBound(alComando)
'            alComando(iIndice) = Comando_Abrir()
'            If alComando(iIndice) = 0 Then gError 182595
'        Next
'
'        'Le a tabelaCoinfoItemFatura
'        lErro = Comando_ExecutarPos(alComando(1), "SELECT cartao FROM CoinfoItemFatura WHERE NumIntArq = ? AND NumRegArqTexto = ? ", 0, iAux, objImpCoinfoItemFat.lNumIntArq, objImpCoinfoItemFat.lNumRegArqTexto)
'        If lErro <> AD_SQL_SUCESSO Then gError 182587
'
'        'Busca Primeiro
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182588
'
'        If lErro <> SUCESSO Then gError 182589
'
'        'Trava registro contra alterações/Leituras
'        lErro = Comando_LockExclusive(alComando(1))
'        If lErro <> SUCESSO Then gError 182590
'
'        If iComErro = MARCADO Then
'            iTotalRegErr = iTotalRegErr + 1
'            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CoinfoItemFatura SET ComErro = ? ", alComando(1), MARCADO)
'        Else
'            iTotalRegCertos = iTotalRegCertos + 1
'            lErro = Comando_ExecutarPos(alComando(3), "UPDATE CoinfoItemFatura SET TipoDocCorporator = ? , NumIntDocCorporator = ?, NoCorporatorEm = ?, ComErro = ?", alComando(1), TIPO_TITULO_RECEBER_TRV, objImpCoinfoCartoes.lNumIntTitulo, gdtDataAtual, DESMARCADO)
'        End If
'        If lErro <> AD_SQL_SUCESSO Then gError 182591
'
'        'Fecha o comando
'        For iIndice = LBound(alComando) To UBound(alComando)
'            Call Comando_Fechar(alComando(iIndice))
'        Next
'
'        lErro = Rotina_Atualizar_AtualizaTelaBatch
'        If lErro <> SUCESSO Then gError 182720
'
'     Next
'
'    Coinfo_AtualizarItensFatu3 = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarItensFatu3:
'
'    Coinfo_AtualizarItensFatu3 = gErr
'
'    Select Case gErr
'
'        Case 182587, 182588
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMFATURA", gErr)
'
'        Case 182589
'            Call Rotina_Erro(vbOKOnly, "ERRO_COINFOITEMFATURA_NAO_CADASTRADO", gErr, objImpCoinfoItemFat.lNumRegArqTexto, objImpCoinfoItemFat.lNumIntArq)
'
'        Case 182590
'            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COINFOITEMFATURA", gErr)
'
'        Case 182591
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOITEMFATURA", gErr)
'
'        Case 182720
'
'        Case 182595
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182592)
'
'    End Select
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function Coinfo_GerarTitRec_Cartao(objImpCoinfoCartoes As ClassImpCoinfoCartoes) As Long

Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objRede As New ClassRede
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objTitRecCartao As New ClassTituloReceber
Dim colParcelasCartao As New ColParcelaReceber
Dim iAchou As Integer
Dim objAdmMeioPagtoParcela As ClassAdmMeioPagtoParcelas
Dim colComissaoEmissao As New colComissao
Dim colcolComissao As New colcolComissao
Dim colComissaoFantasma As colComissao
Dim dtDataVencimento1 As Date
Dim dValor As Double
Dim lNumTitOriginal As Long
Dim lNumAuto As Long
Dim objParcelaRecCartao As ClassParcelaReceber
Dim dtDataVencimentoReal As Date
Dim iIndice As Integer
Dim dValorPresente As Double
Dim dValorTaxa As Double, dDeducaoValorPresente As Double, objParcelaRecCartao1 As ClassParcelaReceber

On Error GoTo Erro_Coinfo_GerarTitRec_Cartao

    lErro = CF("AdmMeioPagto_Le_Todas1", giFilialEmpresa, colAdmMeioPagto)
    If lErro <> SUCESSO Then gError 183111
    
    For Each objAdmMeioPagto In colAdmMeioPagto
        If objAdmMeioPagto.sNome = objImpCoinfoCartoes.sCiaCart Then
            objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then gError 183112
    
    objRede.iCodigo = objAdmMeioPagto.iRede
    
    lErro = CF("Rede_Le", objRede)
    If lErro <> SUCESSO And lErro <> 104244 Then gError 183113
    
    If lErro <> SUCESSO Then gError 183114
    
    Set colComissaoFantasma = New colComissao
    Set colcolComissao = New colcolComissao

    'Lê em colCodPagtoLoja as Codiçoes de Pagamento relacionados a Admnistradora
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104086 Then gError 183120

    iAchou = 0

    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
        
        If objAdmMeioPagtoCondPagto.iNumParcelas = objImpCoinfoCartoes.iQuantParc Then
            iAchou = 1
            Exit For
        End If
    Next

    If iAchou = 0 Then gError 183121

    'se a coleção retornar vazia-> erro
    If objAdmMeioPagtoCondPagto.colParcelas.Count = 0 Then gError 183122


    If objImpCoinfoCartoes.lNumTitulo <> 0 Then
        objTitRecCartao.lNumTitulo = objImpCoinfoCartoes.lNumTitulo ' lNumAuto
    Else
        'gera um número para o título a receber
        lErro = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_NUMTITULO_CARTAO", lNumAuto, objAdmMeioPagtoCondPagto.colParcelas.Count, objImpCoinfoCartoes.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 183123
    
        objTitRecCartao.lNumTitulo = lNumAuto
    
    End If
    
    objTitRecCartao.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB
    objTitRecCartao.dtDataEmissao = objImpCoinfoCartoes.dtDataDep
    dValor = objImpCoinfoCartoes.dValor * (1 - objAdmMeioPagtoCondPagto.dTaxa - objAdmMeioPagtoCondPagto.dDesconto)
    objTitRecCartao.iNumParcelas = objAdmMeioPagtoCondPagto.iParcelasRecebto
    objTitRecCartao.iStatus = STATUS_LANCADO
    objTitRecCartao.dtDataRegistro = gdtDataHoje
    objTitRecCartao.lCliente = objRede.lCliente
    objTitRecCartao.iFilial = objRede.iFilialCli
    objTitRecCartao.iFilialEmpresa = objImpCoinfoCartoes.iFilialEmpresa
    objTitRecCartao.iMoeda = MOEDA_REAL
    objTitRecCartao.dtReajusteBase = DATA_NULA
    objTitRecCartao.dtReajustadoAte = DATA_NULA

    dtDataVencimento1 = objImpCoinfoCartoes.dtDataDep

    dValorTaxa = objImpCoinfoCartoes.dValor * (objAdmMeioPagtoCondPagto.dTaxa - objAdmMeioPagtoCondPagto.dDesconto)

    'varre a coleção de parcelas "modelo do admmeiopagtocondpagto
    For Each objAdmMeioPagtoParcela In objAdmMeioPagtoCondPagto.colParcelas
        
        'cria uma coleção de comissão fantasmas para a parcela em questão
        colcolComissao.Add colComissaoFantasma

        'cria uma nova parcela a receber
        Set objParcelaRecCartao = New ClassParcelaReceber
        
        'o valor da parcela é o valor total multiplicado pelo percentual da parcela
        objParcelaRecCartao.dValor = dValor * objAdmMeioPagtoParcela.dPercRecebimento
        objParcelaRecCartao.iNumParcela = objAdmMeioPagtoParcela.iParcela
        
        dtDataVencimento1 = dtDataVencimento1 + objAdmMeioPagtoParcela.iIntervaloRecebimento
        objParcelaRecCartao.dtDataVencimento = dtDataVencimento1

        'calcula seu vencimento real
        lErro = CF("DataVencto_Real", dtDataVencimento1, dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError 183124

        objParcelaRecCartao.dtDataVencimentoReal = dtDataVencimentoReal
        objParcelaRecCartao.iStatus = STATUS_ABERTO
        objParcelaRecCartao.iCobrador = COBRADOR_PROPRIA_EMPRESA
        objParcelaRecCartao.iCarteiraCobranca = CARTEIRA_CARTEIRA
        
        'adiciona à coleção de parcelas a receber
        With objParcelaRecCartao
        '#################################################
        'ALTERADO POR WAGNER
            colParcelasCartao.Add 0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, DATA_NULA, .dDesconto1Valor, .iDesconto2Codigo, DATA_NULA, .dDesconto2Valor, .iDesconto3Codigo, DATA_NULA, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValor
        '#################################################
        End With

        objTitRecCartao.dValor = objTitRecCartao.dValor + objParcelaRecCartao.dValor

    Next

    If objAdmMeioPagto.dTaxaVista > 0 Then
    
        Set colComissaoFantasma = New colComissao
        Set colcolComissao = New colcolComissao
    
        'cria uma coleção de comissão fantasmas para a parcela em questão
        colcolComissao.Add colComissaoFantasma
    
        'cria uma nova parcela a receber
        Set objParcelaRecCartao = New ClassParcelaReceber

        objParcelaRecCartao.iNumParcela = 1
        objTitRecCartao.iNumParcelas = 1

        objParcelaRecCartao.dtDataVencimento = dtDataVencimento1

        'calcula seu vencimento real
        lErro = CF("DataVencto_Real", dtDataVencimento1, dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError 183124

        objParcelaRecCartao.dtDataVencimentoReal = dtDataVencimentoReal
        objParcelaRecCartao.iStatus = STATUS_ABERTO
        objParcelaRecCartao.iCobrador = COBRADOR_PROPRIA_EMPRESA
        objParcelaRecCartao.iCarteiraCobranca = CARTEIRA_CARTEIRA
        
        '******     INICIO DO CÁLCULO DO VALOR PRESENTE  **********
        dValorPresente = 0
    
        For Each objParcelaRecCartao1 In colParcelasCartao
            
            dValorPresente = dValorPresente + (objParcelaRecCartao1.dValor / ((1 + objAdmMeioPagto.dTaxaVista) ^ ((objParcelaRecCartao1.dtDataVencimento - objParcelaRecCartao.dtDataVencimento) / 30)))
            dDeducaoValorPresente = dDeducaoValorPresente + objParcelaRecCartao1.dValor - (objParcelaRecCartao1.dValor / ((1 + objAdmMeioPagto.dTaxaVista) ^ ((objParcelaRecCartao1.dtDataVencimento - objParcelaRecCartao.dtDataVencimento) / 30)))
        Next
    
        objParcelaRecCartao.dValor = dValorPresente
    
        '***********  FIM DO CÁLCULO DO VALOR PRESENTE  ***********
        
        Set colParcelasCartao = New ColParcelaReceber
        
        'adiciona à coleção de parcelas a receber
        With objParcelaRecCartao
        '#################################################
        'ALTERADO POR WAGNER
            colParcelasCartao.Add 0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, DATA_NULA, .dDesconto1Valor, .iDesconto2Codigo, DATA_NULA, .dDesconto2Valor, .iDesconto3Codigo, DATA_NULA, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValor
        '#################################################
        End With

        objTitRecCartao.dValor = objParcelaRecCartao.dValor

    End If


    objTitRecCartao.dSaldo = Round(objTitRecCartao.dValor, 2)

    dValor = 0

    For iIndice = 1 To colParcelasCartao.Count
        
        Set objParcelaRecCartao = colParcelasCartao.Item(iIndice)
        
        If iIndice = colParcelasCartao.Count Then
            objParcelaRecCartao.dValor = objTitRecCartao.dSaldo - dValor
        Else
        
            objParcelaRecCartao.dValor = Round(objParcelaRecCartao.dValor, 2)
            dValor = dValor + objParcelaRecCartao.dValor
        End If
    
        objParcelaRecCartao.dSaldo = objParcelaRecCartao.dValor
    
    Next

    objTitRecCartao.dSaldo = objTitRecCartao.dValor

    lErro = CF("TituloReceber_Insere", objTitRecCartao, colParcelasCartao, colComissaoEmissao, colcolComissao)
    If lErro <> SUCESSO Then gError 183125
    
    objImpCoinfoCartoes.lNumIntTitulo = objTitRecCartao.lNumIntDoc
    'objImpCoinfoCartoes.lNumTitulo = objTitRecCartao.lNumTitulo
    
    objImpCoinfoCartoes.dValorTaxa = dValorTaxa
    objImpCoinfoCartoes.dDeducaoValorPresente = dDeducaoValorPresente
    objImpCoinfoCartoes.dValorTitulo = objTitRecCartao.dValor

    Coinfo_GerarTitRec_Cartao = SUCESSO

    Exit Function

Erro_Coinfo_GerarTitRec_Cartao:

    Coinfo_GerarTitRec_Cartao = gErr

    Select Case gErr
            
        Case 183111, 183113, 183120, 183123 To 183125

        Case 183112
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTAO_INEXISTENTE", gErr, objImpCoinfoCartoes.sCiaCart)
        
        Case 183114
            Call Rotina_Erro(vbOKOnly, "ERRO_REDE_NAO_ENCONTRADA", gErr, objRede.iCodigo)

        Case 183121
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_INEXISTENTE", gErr, objImpCoinfoCartoes.iQuantParc, objImpCoinfoCartoes.sCiaCart)

        Case 183122
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOPARCELAS_NAOENCONTRADA", gErr, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto, objAdmMeioPagtoCondPagto.iParcelamento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183126)

    End Select

    Exit Function

End Function

Function Coinfo_Gerar_BaixaParcRec(ByVal dValorParam As Double, ByVal lNumIntTitulo As Long, lNumTitulo As Long, ByVal sCiaCart As String, ByVal iQuantParc As Integer, ByVal dtDataBaixa As Date) As Long

Dim objTituloReceber As New ClassTituloReceber
Dim lErro As Long
Dim colAdmMeioPagto As New Collection
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim objTitRecCartao As New ClassTituloReceber
Dim colParcelasCartao As New ColParcelaReceber
Dim objParcelaRecCartao As ClassParcelaReceber
Dim iAchou As Integer
Dim objAdmMeioPagtoParcela As ClassAdmMeioPagtoParcelas
Dim dValor As Double
Dim colBaixaParcReceber As New ColBaixaParcRec
Dim objBaixaParcRec As ClassBaixaParcRec
Dim colParcelasRec As New ColParcelaReceber
Dim objParcRec As ClassParcelaReceber
Dim iIndice As Integer
Dim objContabil As ClassContabil
Dim sHistorico As String
Dim dtDataContabil As Date
Dim dtDataVencimento1 As Date
Dim dtDataVencimentoReal As Date

On Error GoTo Erro_Coinfo_Gerar_BaixaParcRec

    objTituloReceber.lNumIntDoc = lNumIntTitulo

    'Lê Titulo a Receber com o Número Interno passado como parâmetro
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 183127

    If lErro <> SUCESSO Then gError 183128

    lNumTitulo = objTituloReceber.lNumTitulo

    'Lê as Parcelas a Receber do Titulo cujo NumIntDoc foi passado em objTituloReceber
    lErro = CF("ParcelasReceber_Le", objTituloReceber, colParcelasRec)
    If lErro <> SUCESSO Then gError 183132

    lErro = CF("AdmMeioPagto_Le_Todas1", giFilialEmpresa, colAdmMeioPagto)
    If lErro <> SUCESSO Then gError 183129
    
    For Each objAdmMeioPagto In colAdmMeioPagto
    
        If objAdmMeioPagto.sNome = sCiaCart Then
            objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
            iAchou = 1
            Exit For
        End If
    
    Next
    
    If iAchou = 0 Then gError 183130
    
    'Lê em colCodPagtoLoja as Codiçoes de Pagamento relacionados a Admnistradora
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104086 Then gError 183131

    iAchou = 0

    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
        
        If objAdmMeioPagtoCondPagto.iNumParcelas = iQuantParc Then
            iAchou = 1
            Exit For
        End If
    Next

    If iAchou = 0 Then gError 183132
    
    dValor = dValorParam * (1 - objAdmMeioPagtoCondPagto.dTaxa - objAdmMeioPagtoCondPagto.dDesconto)
    
    'varre a coleção de parcelas "modelo do admmeiopagtocondpagto
    For Each objAdmMeioPagtoParcela In objAdmMeioPagtoCondPagto.colParcelas
        
        'cria uma nova parcela a receber
        Set objParcelaRecCartao = New ClassParcelaReceber
        
        'o valor da parcela é o valor total multiplicado pelo percentual da parcela
        objParcelaRecCartao.dValor = dValor * objAdmMeioPagtoParcela.dPercRecebimento
        objParcelaRecCartao.iNumParcela = objAdmMeioPagtoParcela.iParcela
        
        dtDataVencimento1 = dtDataVencimento1 + objAdmMeioPagtoParcela.iIntervaloRecebimento
        objParcelaRecCartao.dtDataVencimento = dtDataVencimento1

        'calcula seu vencimento real
        lErro = CF("DataVencto_Real", dtDataVencimento1, dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError 183124

        objParcelaRecCartao.dtDataVencimentoReal = dtDataVencimentoReal
        objParcelaRecCartao.iStatus = STATUS_ABERTO
        objParcelaRecCartao.iCobrador = COBRADOR_PROPRIA_EMPRESA
        objParcelaRecCartao.iCarteiraCobranca = CARTEIRA_CARTEIRA
        
        'adiciona à coleção de parcelas a receber
        With objParcelaRecCartao
        '#################################################
        'ALTERADO POR WAGNER
            colParcelasCartao.Add 0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, DATA_NULA, .dDesconto1Valor, .iDesconto2Codigo, DATA_NULA, .dDesconto2Valor, .iDesconto3Codigo, DATA_NULA, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValor
        '#################################################
        End With

        objTitRecCartao.dValor = objTitRecCartao.dValor + objParcelaRecCartao.dValor

    Next

    objTitRecCartao.dSaldo = Round(objTitRecCartao.dValor, 2)

'    If objTitRecCartao.dSaldo > objTituloReceber.dSaldo Then objTitRecCartao.dSaldo = objTituloReceber.dSaldo

    dValor = 0

    For iIndice = 1 To colParcelasCartao.Count
        
        Set objParcelaRecCartao = colParcelasCartao.Item(iIndice)
        
        If iIndice = colParcelasCartao.Count Then
            objParcelaRecCartao.dValor = objTitRecCartao.dSaldo - dValor
        Else
        
            objParcelaRecCartao.dValor = Round(objParcelaRecCartao.dValor, 2)
            dValor = dValor + objParcelaRecCartao.dValor
        End If
    
        objParcelaRecCartao.dSaldo = objParcelaRecCartao.dValor
    
        Set objParcRec = colParcelasRec.Item(iIndice)
    
        Call colBaixaParcReceber.Add(0, 0, objParcRec.lNumIntDoc, 0, STATUS_LANCADO, 0, 0, 0, objParcelaRecCartao.dValor, objParcelaRecCartao.dValor, COBRADOR_PROPRIA_EMPRESA)
    
    Next


    'Insere BaixaReceber e BaixasParcelaReceber correspondentes no BD.
    lErro = CF("BaixaRec_Grava_Perda_Trans", colBaixaParcReceber, dtDataBaixa, objContabil, sHistorico, dtDataContabil, objTituloReceber.lCliente, objTituloReceber.iFilial)
    If lErro <> SUCESSO Then gError 183133

    Coinfo_Gerar_BaixaParcRec = SUCESSO

    Exit Function

Erro_Coinfo_Gerar_BaixaParcRec:

    Coinfo_Gerar_BaixaParcRec = gErr

    Select Case gErr
            
        Case 183127, 183129, 183133

        Case 183128
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSREC_NAO_EXISTENTE", gErr, objTituloReceber.lNumIntDoc)

        Case 183130
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTAO_INEXISTENTE", gErr, sCiaCart)
        
        Case 183132
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_INEXISTENTE", gErr, iQuantParc, sCiaCart)

        Case 183122
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOPARCELAS_NAOENCONTRADA", gErr, objAdmMeioPagtoCondPagto.iFilialEmpresa, objAdmMeioPagtoCondPagto.iAdmMeioPagto, objAdmMeioPagtoCondPagto.iParcelamento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183126)

    End Select

    Exit Function

End Function

Private Function Coinfo_Total_Registros(ByVal objImpCoinfo As ClassImpCoinfo, ByVal sArq As String) As Long
'Calcula o total de registros que vai ser atualizado e acerta a tela de acompanhamento

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntArq As Long
Dim lTotalRegAux As Long
Dim sTabela As String
Dim sWhere As String
Dim sNomeTabela As String
Dim lTotalRegistros As Long
Dim vValorFiltro1 As Variant
Dim vValorFiltro2 As Variant
Dim sDataAtualizado As String

On Error GoTo Erro_Coinfo_Total_Registros

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182710
    Next
    
    lTotalRegistros = 0
    
    sDataAtualizado = "DataAtualizado"
    
    'Dependendo do arquivo, aplica o mesmo filtro da rotina de atualização
    Select Case sArq
    
        Case "fptipoempr", "fpproduto"
            If sArq = "fpproduto" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "4 de 15"
                sNomeTabela = "Produtos"
                sTabela = "CoinfoProduto"
            Else
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "1 de 15"
                sNomeTabela = "Tipos de Cliente e Tipos de Fornecedor"
                sTabela = "CoinfoTipoEmpr"
            End If
            sWhere = "WHERE NumIntArq = ? AND 1=? AND 1=? "
            vValorFiltro1 = 1
            vValorFiltro2 = 1
            
        Case "fpempresa", "fpemissor", "fpocorrencia", "fpdestino", "fpacordo"
            If sArq = "fpempresa" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "2 de 15"
                sTabela = "CoinfoEmpresa"
                sNomeTabela = "Clientes e Fornecedores"
            ElseIf sArq = "fpemissor" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "8 de 15"
                sTabela = "CoinfoEmissor"
                sNomeTabela = "Fornecedores (Emissores)"
            ElseIf sArq = "fpdestino" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "3 de 15"
                sTabela = "CoinfoDestino"
                sNomeTabela = "Destinos"
            ElseIf sArq = "fpacordo" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "5 de 15"
                sTabela = "CoinfoAcordo"
                sNomeTabela = "Acordos"
            Else
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "15 de 15"
                sTabela = "CoinfoOcorrencia"
                sNomeTabela = "Ocorrências"
            End If
            sWhere = "WHERE NumIntArq = ? AND 1=? AND NoCorporatorEm = ? "
            vValorFiltro1 = 1
            vValorFiltro2 = DATA_NULA
            
        Case "fpcabecfat", "fpnf", "fpitemfatu", "fpover"
            If sArq = "fpcabecfat" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "X de 15"
                sTabela = "CoinfoFatura"
                sNomeTabela = "Títulos a Receber e Títulos a Pagar"
                sWhere = "WHERE NumIntArq = ? AND CodEst=? AND NoCorporatorEm = ? "
            ElseIf sArq = "fpnf" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "X de 15"
                sTabela = "CoinfoNF"
                sNomeTabela = "Notas Fiscais"
                sWhere = "WHERE NumIntArq = ? AND CodEst=? AND NoCorporatorEm = ? "
            ElseIf sArq = "fpitemfatu" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "14 de 15"
                sTabela = "CoinfoItemFatura"
                sNomeTabela = "Títulos a Receber de Cartões"
                sWhere = "WHERE NumIntArq = ? AND CodEst=? AND NoCorporatorEm = ? AND Cartao = 1 AND Data > {d '2006-06-30'} "
            ElseIf sArq = "fpover" Then
                TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "9 de 15"
                sTabela = "CoinfoComissoesOver"
                sNomeTabela = "Títulos a Pagar Over"
                sWhere = "WHERE NumIntArq = ? AND CodEst=? AND NoCorporatorEm = ? AND DatEmi > {d '2006-04-30'} AND Numpgcm <>0 "
            End If
            vValorFiltro1 = objImpCoinfo.sSiglaEmpresa
            vValorFiltro2 = DATA_NULA
            
        Case "comissao"
            TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "9 de 15"
            sTabela = "CoinfoItemFaturaImport2"
            sNomeTabela = "Nfs a Pagar (Comissão)"
            sWhere = "WHERE (? = (SELECT MAX(NumIntImportacao) FROM CoinfoArqsLidos WHERE NomeArq = 'Comissao' AND DataAtualizado = {d '1822-09-07'})) AND CodEst=? AND SiglaDocumento = ? AND DataEmissao > {d '2006-10-31'} AND EXISTS (SELECT C.NumVou FROM CoinfoComissao AS C WHERE C.ValComi <> 0 AND C.Serie=CoinfoItemFaturaImport2.Serie AND C.NumVou = CoinfoItemFaturaImport2.NumVou) AND NOT EXISTS (SELECT C.NumVou FROM CoinfoComissao AS C WHERE C.Serie=CoinfoItemFaturaImport2.Serie AND C.NumVou = CoinfoItemFaturaImport2.NumVou AND C.NumIntDocCorporator <>0) "
            vValorFiltro1 = objImpCoinfo.sSiglaEmpresa
            vValorFiltro2 = TIPODOC_CARTAO_CRED_DEB
        Case "fpitemfatu2"
            TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "11 de 15"
            sTabela = "CoinfoItemFatura"
            sNomeTabela = "Vouchers emissão de Impressos"
            sWhere = "WHERE NumIntArq = ? AND CodEst=? AND DataAtualizacaoMovEst = ? AND Data > {d '2007-01-01'} AND Grupo <> '' AND TipoDoc = 'VOU'  AND DiasAntc = 0 "
            vValorFiltro1 = objImpCoinfo.sSiglaEmpresa
            vValorFiltro2 = DATA_NULA
            sDataAtualizado = "DataAtualizado2"
            sArq = "fpitemfatu"
        Case "fpitemfatu3"
            TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = "12 de 15"
            sTabela = "CoinfoItemFaturaCTB"
            sNomeTabela = "Vouchers (Contabilização)"
            sWhere = "WHERE ((NumIntArq = ? AND CodEst=?) OR ForcaCTB=1) AND DataAtualizacaoContab = ? AND Data > {d '2007-01-01'} "
            vValorFiltro1 = objImpCoinfo.sSiglaEmpresa
            vValorFiltro2 = DATA_NULA
            sDataAtualizado = "DataAtualizado3"
            sArq = "fpitemfatu"

    End Select
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Início da importação de " & sNomeTabela & vbNewLine
    
    'Le os arquivos que vão ser atualizados
    lErro = Comando_Executar(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND " & sDataAtualizado & " = ? ORDER BY IdentificacaoArq", lNumIntArq, sArq, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 182711
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182712
    
    'Enquando existirem arquivos não atualizados
    Do While lErro = AD_SQL_SUCESSO
    
        'Le o total de registros não atualizados desse arquivo
        lErro = Comando_Executar(alComando(1), "SELECT Count(*) FROM " & sTabela & " " & sWhere, lTotalRegAux, lNumIntArq, vValorFiltro1, vValorFiltro2)
        If lErro <> AD_SQL_SUCESSO Then
            gError 182713
        End If
        
        'Busca
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182712
        
        'Totaliza os registros não atualizados
        lTotalRegistros = lTotalRegistros + lTotalRegAux

        'Busca o próximo registro
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182712

    Loop
    
    TelaAcompanhaBatchFAT.NomeArqAtu = sNomeTabela
    
    TelaAcompanhaBatchFAT.dValorAtualAtu = 0
    If lTotalRegistros = 0 Then
        TelaAcompanhaBatchFAT.dValorTotalAtu = 1
        
        'Atualiza a barra de progressão
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182727
    
    Else
        TelaAcompanhaBatchFAT.dValorTotalAtu = lTotalRegistros
    End If
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Coinfo_Total_Registros = SUCESSO

    Exit Function

Erro_Coinfo_Total_Registros:

    Coinfo_Total_Registros = gErr

    Select Case gErr
    
        Case 182710
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 182711, 182712, 182713
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
            
        Case 182727

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182715)

    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Coinfo_Total_Registros2(ByVal sNomeTabela As String, ByVal sEtapa As String, ByVal lTotalRegistros As Long) As Long
'Calcula o total de registros que vai ser atualizado e acerta a tela de acompanhamento

Dim lErro As Long

On Error GoTo Erro_Coinfo_Total_Registros2
    
    TelaAcompanhaBatchFAT.NomeArqAtu2.Caption = sEtapa
    TelaAcompanhaBatchFAT.NomeArqAtu = sNomeTabela
    
    TelaAcompanhaBatchFAT.dValorAtualAtu = 0
    If lTotalRegistros = 0 Then
        TelaAcompanhaBatchFAT.dValorTotalAtu = 1
        
        'Atualiza a barra de progressão
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 182727
    
    Else
        TelaAcompanhaBatchFAT.dValorTotalAtu = lTotalRegistros
    End If
    
    Coinfo_Total_Registros2 = SUCESSO

    Exit Function

Erro_Coinfo_Total_Registros2:

    Coinfo_Total_Registros2 = gErr

    Select Case gErr
    
        Case 182727

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182715)

    End Select

    Exit Function

End Function

Private Function Abre_Conexoes() As Long

Dim lErro As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim objEmpresa As New ClassDicEmpresa

On Error GoTo Erro_Abre_Conexoes

    objEmpresa.lCodigo = glEmpresa
    
    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError 182723

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    glConexao1 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If glConexao1 = 0 Then gError 182724
    
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    glConexao2 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If glConexao2 = 0 Then gError 182725
    
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    glConexao3 = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If glConexao3 = 0 Then gError 182725
    
    Abre_Conexoes = SUCESSO
    
    Exit Function

Erro_Abre_Conexoes:

    Abre_Conexoes = gErr

    Select Case gErr
    
        Case 182723, 182724, 182725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182726)

    End Select

    Exit Function

End Function

Private Function Fecha_Conexoes() As Long

On Error GoTo Erro_Fecha_Conexoes

    Call Conexao_FecharExt(glConexao1)
    Call Conexao_FecharExt(glConexao2)
    Call Conexao_FecharExt(glConexao3)
    
    Fecha_Conexoes = SUCESSO
    
    Exit Function

Erro_Fecha_Conexoes:

    Fecha_Conexoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182727)

    End Select

    Exit Function

End Function

'Private Function Coinfo_Obtem_FilialEmp_Cliente(ByVal lcodemp As Long, iFilialEmpCorporator As Integer, iFilialEmpCoinfo As Integer) As Long
''Obtém a FilialEmpresa com base no Cliente (lcodemp)
'
'Dim lErro As Long
'Dim alComando(0 To 0) As Long
'Dim iIndice As Integer
'Dim iFilialCorp As Integer
'Dim iFilialCoinfo As Integer
'Dim lEmpresa As Long
'
'On Error GoTo Erro_Coinfo_Obtem_FilialEmp_Cliente
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 130803
'    Next
'
'    lEmpresa = lcodemp
'
'    'Le o Grupo
'    lErro = Comando_Executar(alComando(0), "SELECT Filial FROM CoInfoEmpresaGrupo WHERE CodEmp = ? ", _
'        iFilialCoinfo, lEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 130805
'
'    'Busca o primeiro
'    lErro = Comando_BuscarProximo(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
'
'    If lErro = SUCESSO Then
'        iFilialEmpCorporator = Filial_Coinfo_Retorna_Corporator(iFilialCoinfo)
'        iFilialEmpCoinfo = iFilialCoinfo
'    Else
'        iFilialEmpCorporator = FILIAL_MATRIZ
'        iFilialEmpCoinfo = FILIAL_MATRIZ
'    End If
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Coinfo_Obtem_FilialEmp_Cliente = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_Obtem_FilialEmp_Cliente:
'
'    Coinfo_Obtem_FilialEmp_Cliente = gErr
'
'    Select Case gErr
'
'        Case 130803
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 130805, 130806
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Coinfo_Obtem_Comissao_Cliente(ByVal lNumFat As Long, ByVal scodest As String, dComissao As Double, Optional ByVal bExcluiCartao As Boolean = False) As Long
'Obtém a FilialEmpresa com base no Cliente (lcodemp)

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim sSerie As String
Dim lNumVou As Long
Dim sSerieAnt As String
Dim lNumVouAnt As Long
Dim dValComi As Double
Dim iCartao As Integer

On Error GoTo Erro_Coinfo_Obtem_Comissao_Cliente

    dComissao = 0

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next

    sSerie = String(255, 0)
    
    'Le os vouchers
    lErro = Comando_Executar(alComando(0), "SELECT Serie, NumVou, Cartao FROM CoinfoItemFatura AS I WHERE I.NumFat = ? AND I.CodEST = ? ORDER BY I.Serie DESC, I.NumVou DESC, NUmIntArq DESC, NumRegArqTexto DESC ", _
        sSerie, lNumVou, iCartao, lNumFat, scodest)
    If lErro <> AD_SQL_SUCESSO Then gError 130805
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
    'Enquanto existirem vouchers
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If sSerieAnt <> sSerie Or lNumVouAnt <> lNumVou Then
            
            sSerieAnt = sSerie
            lNumVouAnt = lNumVou

            If Not bExcluiCartao Or iCartao <> MARCADO Then
            
                'Le a comissão do voucher
                lErro = Comando_Executar(alComando(1), "SELECT C.ValComi FROM CoinfoComissao AS C WHERE C.Serie = ? AND C.NumVou = ? ORDER BY C.NumIntArq DESC, C.NumRegArqTexto DESC ", _
                    dValComi, sSerie, lNumVou)
                If lErro <> AD_SQL_SUCESSO Then gError 130805
                
                'Busca o primeiro
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
                
                dComissao = dComissao + dValComi
                
            End If
            
        End If
    
        'Busca o próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_Obtem_Comissao_Cliente = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Obtem_Comissao_Cliente:

    Coinfo_Obtem_Comissao_Cliente = gErr
     
    Select Case gErr

        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130805, 130806
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select
     
    Exit Function

End Function

Private Function Coinfo_Gera_Arquivos(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'Gera os arquivos de baixa e de cancelamento de baixa

Dim lErro As Long
Dim dtData As Date

Const DATA_INI_GERAR_ARQ_BAIXAS = #7/7/2008#

On Error GoTo Erro_Coinfo_Gera_Arquivos

    If iSistemaIntegrado = SISTEMA_INTEGRADO_SIGAV Then

        dtData = DATA_INI_GERAR_ARQ_BAIXAS
    
        'Pecorre desde DATA_INI_GERAR_ARQ_BAIXAS até ontem
        Do While dtData < DateAdd("d", -1, Date)
    
            dtData = DateAdd("d", 1, dtData)
    
            'Gera os arquivos que estão faltando
            lErro = Coinfo_Gera_Arquivos2(objImpCoinfo, dtData)
            If lErro <> SUCESSO Then gError 182758
    
            'Gera os arquivos que estão faltando
            lErro = Coinfo_Gera_Arquivos3(objImpCoinfo, dtData)
            If lErro <> SUCESSO Then gError 182758
    
        Loop
        
        lErro = CF("TRVTitulos_Exporta_Sigav")
        If lErro <> SUCESSO Then gError 196242
        
    End If
    
    If iSistemaIntegrado = SISTEMA_INTEGRADO_KOGUT Then
        lErro = CF("TRVOcorrencias_Exporta_Sigav")
        If lErro <> SUCESSO Then gError 196242
    End If
    
    If iSistemaIntegrado = SISTEMA_INTEGRADO_KOGUT Then
        lErro = CF("TRVVouchers_Exporta_Sigav")
        If lErro <> SUCESSO Then gError 196242
    End If
        
    Coinfo_Gera_Arquivos = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Gera_Arquivos:

    Coinfo_Gera_Arquivos = gErr
     
    Select Case gErr
    
        Case 182758, 196242

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182759)
     
    End Select
    
    Exit Function

End Function

Private Function Coinfo_Gera_Arquivos2(ByVal objImpCoinfo As ClassImpCoinfo, ByVal dtData As Date) As Long
'Gera os arquivos de baixas que ainda não foram gerados na pasta indicada no ADM100.INI

'2006-11-29
'Andre Tatsuhiko Korosue: 11:44:16 OK, mas gostaria que vocês fossem já estudando o seguinte:
'preciso que os arquivos de baixa, independente de ser TVI ou TVA caiam em /var/forprint/baixa
'e os arquivos de cancelamento caiam em /var/forprint/cancbaixa do nosso servidor.
'Andre Tatsuhiko Korosue: 11:44:51... Tudo bem?
'Wagner Luis Carvalhosa de Mello: 11:47:23 A gente não tem como gerar os arquivos das duas
'empresas juntas porque a sua geração ocorre em momentos diferentes. Podemos colocar os
'dois arquivos na mesma pasta mas com nomes diferentes, quanto a localidades dessas
'pastas não tem problema algum.
'Andre Tatsuhiko Korosue: 11:48:27 OK, vocês podem gerar de modo separado porém gostaria
'que caissem na mesma pasta.
'Wagner Luis Carvalhosa de Mello: 11:49:24 OK. Vou colocar a sigla da empresa antes do nome
'do arquivo.

Dim lErro As Long
Dim alComando(1 To 6) As Long
Dim iIndice As Integer
Dim sSQL1 As String
Dim sSQL2 As String
Dim sDiretorio As String
Dim lRetorno As Long
Dim sArquivo1 As String
Dim sHistorico As String
Dim dtDataCancelamento As Date
Dim dtDataBaixa As Date
Dim dValorBaixado As Double
Dim dValorRecebido As Double
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim iFilial As Integer
Dim lCliente As Long
Dim lNumTitulo As Long
Dim dtDataRef As Date
Dim bArq1Aberto As Boolean
Dim lTransacao As Long
Dim lNumIntDocParcela As Long
Dim objImpCoinfoArqBaixa As ClassImpCoinfoArqBaixa
Dim iNumParcela As Integer
Dim dHora As Double
Dim colImpCoinfoArqBaixa As New Collection
Dim dValorBaixadoCanc As Double
Dim dvalorRecebidoCanc As Double
Dim dValorMultaCanc As Double
Dim dValorJurosCanc As Double
Dim dValorDescontoCanc As Double
Dim colCampos As New Collection
Dim colSaida As New Collection
Dim iStatus As Integer
Dim lNumIntBaixaParcRec As Long

Const TIPO_ARQ_BAIXAS = 1
Const TIPO_ARQ_CANC = 2

On Error GoTo Erro_Coinfo_Gera_Arquivos2

    bArq1Aberto = False

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182750
    Next
        
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqImpCorp", "c:\travelace\corporator\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    'sDiretorio = sDiretorio & objImpCoinfo.sSiglaEmpresa & "\"
    
    'Leitura das Baixas
    lErro = Comando_Executar(alComando(3), "SELECT DataRef FROM CoinfoArqsGerados WHERE TipoArquivo = ? AND DataRef = ?", dtDataRef, TIPO_ARQ_BAIXAS, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 182760
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182761
    
    'Trata Baixas
    If lErro <> SUCESSO Then
    
        sSQL1 = "SELECT BPR.NumIntDoc, BPR.Status, PR.NumParcela, BR.HoraRegistro, BPR.NumIntParcela, BR.Data, BR.Historico, BPR.ValorBaixado, BPR.ValorRecebido, TR.FilialEmpresa, TR.Cliente, TR.NumTitulo,BPR.ValorMulta, BPR.ValorJuros, BPR.ValorDesconto  FROM BaixasRec AS BR, BaixasParcRec AS BPR, ParcelasRecTodas AS PR, TitulosRecTodos AS TR " & _
                "WHERE BR.NumIntBaixa = BPR.NumIntBaixa AND BPR.NumIntParcela = PR.NumIntDoc AND PR.NumIntTitulo = TR.NumIntDoc AND BR.DataRegistro = ? AND (TR.SiglaDocumento = ? OR (TR.SiglaDocumento = ? AND TR.DataEmissao > ?))"
        
        'sArquivo1 = "Baixas" & Format(dtData, "YYYYMMDD") & ".txt"
        sArquivo1 = objImpCoinfo.sSiglaEmpresa & "BaixasECancBaixas" & Format(dtData, "YYYYMMDD") & ".txt"
                
        sHistorico = String(255, 0)
        
        'Leitura das Baixas
        lErro = Comando_Executar(alComando(1), sSQL1, lNumIntBaixaParcRec, iStatus, iNumParcela, dHora, lNumIntDocParcela, dtDataBaixa, sHistorico, dValorBaixado, dValorRecebido, iFilial, lCliente, lNumTitulo, dValorMulta, dValorJuros, dValorDesconto, dtData, TIPODOC_FATURA_SERVICO_CR, TIPODOC_CARTAO_CRED_DEB, #10/31/2006#)
        If lErro <> AD_SQL_SUCESSO Then gError 182751
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182752
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Se a baixa tiver sido cancelada o valor que havia sido baixado deve ser
            'obtido dos cancelamentos
            If iStatus = STATUS_EXCLUIDO Then
            
                'Leitura dos Cancelamentos
                lErro = Comando_Executar(alComando(6), "SELECT SUM(ValorBaixadoCancelado), SUM(MultaCancelada), SUM(JurosCancelado), SUM(DescontoCancelado), SUM(TotalCancelado) FROM BaixasParcRecCancItem WHERE NumIntBaixaParcRec = ?", dValorBaixadoCanc, dValorMultaCanc, dValorJurosCanc, dValorDescontoCanc, dvalorRecebidoCanc, lNumIntBaixaParcRec)
                If lErro <> AD_SQL_SUCESSO Then gError 182754
                
                lErro = Comando_BuscarProximo(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182755
            
            End If
        
            Set objImpCoinfoArqBaixa = New ClassImpCoinfoArqBaixa
            
            objImpCoinfoArqBaixa.dHora = dHora
            objImpCoinfoArqBaixa.dtData = dtDataBaixa
            If iStatus = STATUS_EXCLUIDO Then
                objImpCoinfoArqBaixa.dValor = dValorBaixadoCanc
                objImpCoinfoArqBaixa.dValorDesconto = dValorDescontoCanc
                objImpCoinfoArqBaixa.dValorJuros = dValorJurosCanc
                objImpCoinfoArqBaixa.dValorMulta = dValorMultaCanc
                objImpCoinfoArqBaixa.dValorRecebido = dvalorRecebidoCanc
            Else
                objImpCoinfoArqBaixa.dValor = dValorBaixado
                objImpCoinfoArqBaixa.dValorDesconto = dValorDesconto
                objImpCoinfoArqBaixa.dValorJuros = dValorJuros
                objImpCoinfoArqBaixa.dValorMulta = dValorMulta
                objImpCoinfoArqBaixa.dValorRecebido = dValorRecebido
            End If
            objImpCoinfoArqBaixa.iNumParcela = iNumParcela
            objImpCoinfoArqBaixa.iTipo = TIPO_ARQ_BAIXAS
            objImpCoinfoArqBaixa.lNumTitulo = lNumTitulo
            
            colImpCoinfoArqBaixa.Add objImpCoinfoArqBaixa
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182753
            
        Loop
    
        sSQL2 = "SELECT PR.NumParcela, BPR.HoraRegCancelamento, BPR.NumIntParcela, BPR.DataCancelamento, BR.Historico, BPRCI.ValorBaixadoCancelado, BPR.ValorRecebido, TR.FilialEmpresa, TR.Cliente, TR.NumTitulo,BPRCI.MultaCancelada, BPRCI.JurosCancelado, BPRCI.DescontoCancelado, BPRCI.TotalCancelado  FROM BaixasRec AS BR, BaixasParcRec AS BPR, ParcelasRecTodas AS PR, TitulosRecTodos AS TR, BaixasParcRecCancItem AS BPRCI " & _
                "WHERE BPRCI.NumIntBaixaParcRec = BPR.NumIntDoc AND BR.NumIntBaixa = BPR.NumIntBaixa AND BPR.NumIntParcela = PR.NumIntDoc AND PR.NumIntTitulo = TR.NumIntDoc AND BPR.DataRegCancelamento = ? AND (TR.SiglaDocumento = ? OR (TR.SiglaDocumento = ? AND TR.DataEmissao > ?))"
                
        sHistorico = String(255, 0)
        
        'Leitura dos Cancelamentos
        lErro = Comando_Executar(alComando(2), sSQL2, iNumParcela, dHora, lNumIntDocParcela, dtDataCancelamento, sHistorico, dValorBaixadoCanc, dValorRecebido, iFilial, lCliente, lNumTitulo, dValorMultaCanc, dValorJurosCanc, dValorDescontoCanc, dvalorRecebidoCanc, dtData, TIPODOC_FATURA_SERVICO_CR, TIPODOC_CARTAO_CRED_DEB, #10/31/2006#)
        If lErro <> AD_SQL_SUCESSO Then gError 182754
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182755
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objImpCoinfoArqBaixa = New ClassImpCoinfoArqBaixa
            
            objImpCoinfoArqBaixa.dHora = dHora
            objImpCoinfoArqBaixa.dtData = dtDataCancelamento
            objImpCoinfoArqBaixa.dValor = dValorBaixadoCanc
            objImpCoinfoArqBaixa.dValorDesconto = dValorDescontoCanc
            objImpCoinfoArqBaixa.dValorJuros = dValorJurosCanc
            objImpCoinfoArqBaixa.dValorMulta = dValorMultaCanc
            objImpCoinfoArqBaixa.dValorRecebido = dvalorRecebidoCanc
            objImpCoinfoArqBaixa.iNumParcela = iNumParcela
            objImpCoinfoArqBaixa.iTipo = TIPO_ARQ_CANC
            objImpCoinfoArqBaixa.lNumTitulo = lNumTitulo
            
            colImpCoinfoArqBaixa.Add objImpCoinfoArqBaixa
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182756
            
        Loop
        
        colCampos.Add "dHora"
        
        Call Ordena_Colecao(colImpCoinfoArqBaixa, colSaida, colCampos)
        
        Open sDiretorio & sArquivo1 For Output As #1
        bArq1Aberto = True
        
        For Each objImpCoinfoArqBaixa In colSaida
            Print #1, FormataCpoData(objImpCoinfoArqBaixa.dtData) & objImpCoinfo.sSiglaEmpresa & FormataCpoNum(objImpCoinfoArqBaixa.lNumTitulo, 10) & objImpCoinfoArqBaixa.iTipo & FormataCpoNum(objImpCoinfoArqBaixa.iNumParcela, 2) & FormataCpoValor(objImpCoinfoArqBaixa.dValor, 10) & FormataCpoValor(objImpCoinfoArqBaixa.dValorJuros, 10) & FormataCpoValor(objImpCoinfoArqBaixa.dValorMulta, 10) & FormataCpoValor(objImpCoinfoArqBaixa.dValorDesconto, 10) & FormataCpoValor(objImpCoinfoArqBaixa.dValorRecebido, 10)
        Next
        
        Close #1
        bArq1Aberto = False
        
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 182768
        
        lErro = Comando_Executar(alComando(5), "INSERT INTO CoinfoArqsGerados (DataRef, TipoArquivo, NomeArquivo, DataGeracao) VALUES (?,?,?,?)", dtData, TIPO_ARQ_BAIXAS, sArquivo1, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 182769
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 182770
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Coinfo_Gera_Arquivos2 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Gera_Arquivos2:

    Coinfo_Gera_Arquivos2 = gErr
     
    Select Case gErr

        Case 182750
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 182751 To 182756
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASREC", gErr)
            
        Case 182760 To 182763
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSGERADOS", gErr)
        
         Case 182765, 182768
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182766, 182769
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOARQSGERADOS", gErr)
            
        Case 182767, 182770
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 185057
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182757)
     
    End Select

    If bArq1Aberto Then Close #1
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Coinfo_Gera_Arquivos3(ByVal objImpCoinfo As ClassImpCoinfo, ByVal dtData As Date) As Long
'Gera os arquivos de baixas que ainda não foram gerados na pasta indicada no ADM100.INI

Dim lErro As Long
Dim alComando(1 To 7) As Long
Dim iIndice As Integer
Dim sSQL1 As String
Dim sSQL2 As String
Dim sDiretorio As String
Dim lRetorno As Long
Dim sArquivo1 As String
Dim sHistorico As String
Dim dtDataCancelamento As Date
Dim dtDataBaixa As Date
Dim dValorBaixado As Double
Dim dValorRecebido As Double
Dim iFilial As Integer
Dim lCliente As Long
Dim lNumTitulo As Long
Dim dtDataRef As Date
Dim bArq1Aberto As Boolean
Dim lTransacao As Long
Dim lNumIntDocParcela As Long
Dim objImpCoinfoArqBaixa As ClassImpCoinfoArqBaixa
Dim iNumParcela As Integer
Dim dHora As Double
Dim colImpCoinfoArqBaixa As New Collection
Dim dValorBaixadoCanc As Double
Dim colCampos As New Collection
Dim colSaida As New Collection
Dim iStatus As Integer
Dim lNumIntBaixaParcRec As Long
Dim lNumIntTitulo As Long
Dim sNumCCred As String
Dim lNumVou As Long

Const TIPO_ARQ_BAIXAS = 1
Const TIPO_ARQ_CANC = 2
Const TIPO_ARQ_BAIXAS_CCANTIGO = 3

On Error GoTo Erro_Coinfo_Gera_Arquivos3

    bArq1Aberto = False

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182750
    Next
        
    'Obtém o diretório onde estão os arquivos
    sDiretorio = String(512, 0)
    lRetorno = GetPrivateProfileString("Geral", "dirArqImpCorpCC", "c:\travelace\cartao\", sDiretorio, 512, "ADM100.INI")
    sDiretorio = left(sDiretorio, lRetorno)
    
    'sDiretorio = sDiretorio & objImpCoinfo.sSiglaEmpresa & "\"
    
    'Leitura das Baixas
    lErro = Comando_Executar(alComando(3), "SELECT DataRef FROM CoinfoArqsGerados WHERE TipoArquivo = ? AND DataRef = ?", dtDataRef, TIPO_ARQ_BAIXAS_CCANTIGO, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 182760
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182761
    
    'Trata Baixas
    If lErro <> SUCESSO Then
    
        sSQL1 = "SELECT TR.NumIntDoc, BPR.NumIntDoc, BPR.Status, PR.NumParcela, BR.HoraRegistro, BPR.NumIntParcela, BR.Data, BR.Historico, BPR.ValorBaixado, BPR.ValorRecebido, TR.FilialEmpresa, TR.Cliente, TR.NumTitulo  FROM BaixasRec AS BR, BaixasParcRec AS BPR, ParcelasRecTodas AS PR, TitulosRecTodos AS TR " & _
                "WHERE BR.NumIntBaixa = BPR.NumIntBaixa AND BPR.NumIntParcela = PR.NumIntDoc AND PR.NumIntTitulo = TR.NumIntDoc AND BR.DataRegistro = ? AND ((TR.SiglaDocumento = ? AND TR.DataEmissao <= ?))"
        
        'sArquivo1 = "Baixas" & Format(dtData, "YYYYMMDD") & ".txt"
        sArquivo1 = objImpCoinfo.sSiglaEmpresa & "BaixasECancBaixasCCAntigo" & Format(dtData, "YYYYMMDD") & ".txt"
                
        sHistorico = String(255, 0)
        
        'Leitura das Baixas
        lErro = Comando_Executar(alComando(1), sSQL1, lNumIntTitulo, lNumIntBaixaParcRec, iStatus, iNumParcela, dHora, lNumIntDocParcela, dtDataBaixa, sHistorico, dValorBaixado, dValorRecebido, iFilial, lCliente, lNumTitulo, dtData, TIPODOC_CARTAO_CRED_DEB, #10/31/2006#)
        If lErro <> AD_SQL_SUCESSO Then gError 182751
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182752
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Se a baixa tiver sido cancelada o valor que havia sido baixado deve ser
            'obtido dos cancelamentos
            If iStatus = STATUS_EXCLUIDO Then
            
                'Leitura dos Cancelamentos
                lErro = Comando_Executar(alComando(6), "SELECT SUM(ValorBaixadoCancelado) FROM BaixasParcRecCancItem WHERE NumIntBaixaParcRec = ?", dValorBaixadoCanc, lNumIntBaixaParcRec)
                If lErro <> AD_SQL_SUCESSO Then gError 182754
                
                lErro = Comando_BuscarProximo(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182755
            
            End If
        
            Set objImpCoinfoArqBaixa = New ClassImpCoinfoArqBaixa
            
            objImpCoinfoArqBaixa.dHora = dHora
            objImpCoinfoArqBaixa.dtData = dtDataBaixa
            If iStatus = STATUS_EXCLUIDO Then
                objImpCoinfoArqBaixa.dValor = dValorBaixadoCanc
            Else
                objImpCoinfoArqBaixa.dValor = dValorBaixado
            End If
            objImpCoinfoArqBaixa.iNumParcela = iNumParcela
            objImpCoinfoArqBaixa.iTipo = TIPO_ARQ_BAIXAS
            objImpCoinfoArqBaixa.lNumTitulo = lNumTitulo
            objImpCoinfoArqBaixa.lNumIntBaixaParcRec = lNumIntBaixaParcRec
            objImpCoinfoArqBaixa.lNumIntTitulo = lNumIntTitulo
            objImpCoinfoArqBaixa.lNumIntParcela = lNumIntDocParcela
            
            colImpCoinfoArqBaixa.Add objImpCoinfoArqBaixa
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182753
            
        Loop
    
        sSQL2 = "SELECT TR.NumIntDoc, BPR.NumIntDoc, PR.NumParcela, BPR.HoraRegCancelamento, BPR.NumIntParcela, BPR.DataCancelamento, BR.Historico, BPRCI.ValorBaixadoCancelado, BPR.ValorRecebido, TR.FilialEmpresa, TR.Cliente, TR.NumTitulo  FROM BaixasRec AS BR, BaixasParcRec AS BPR, ParcelasRecTodas AS PR, TitulosRecTodos AS TR, BaixasParcRecCancItem AS BPRCI " & _
                "WHERE BPRCI.NumIntBaixaParcRec = BPR.NumIntDoc AND BR.NumIntBaixa = BPR.NumIntBaixa AND BPR.NumIntParcela = PR.NumIntDoc AND PR.NumIntTitulo = TR.NumIntDoc AND BPR.DataRegCancelamento = ? AND ((TR.SiglaDocumento = ? AND TR.DataEmissao <= ?))"
                
        sHistorico = String(255, 0)
        
        'Leitura dos Cancelamentos
        lErro = Comando_Executar(alComando(2), sSQL2, lNumIntTitulo, lNumIntBaixaParcRec, iNumParcela, dHora, lNumIntDocParcela, dtDataCancelamento, sHistorico, dValorBaixadoCanc, dValorRecebido, iFilial, lCliente, lNumTitulo, dtData, TIPODOC_CARTAO_CRED_DEB, #10/31/2006#)
        If lErro <> AD_SQL_SUCESSO Then gError 182754
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182755
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objImpCoinfoArqBaixa = New ClassImpCoinfoArqBaixa
            
            objImpCoinfoArqBaixa.dHora = dHora
            objImpCoinfoArqBaixa.dtData = dtDataCancelamento
            objImpCoinfoArqBaixa.dValor = dValorBaixadoCanc
            objImpCoinfoArqBaixa.iNumParcela = iNumParcela
            objImpCoinfoArqBaixa.iTipo = TIPO_ARQ_CANC
            objImpCoinfoArqBaixa.lNumTitulo = lNumTitulo
            objImpCoinfoArqBaixa.lNumIntBaixaParcRec = lNumIntBaixaParcRec
            objImpCoinfoArqBaixa.lNumIntTitulo = lNumIntTitulo
            objImpCoinfoArqBaixa.lNumIntParcela = lNumIntDocParcela
            
            colImpCoinfoArqBaixa.Add objImpCoinfoArqBaixa
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182756
            
        Loop
        
        colCampos.Add "dHora"
        colCampos.Add "lNumIntBaixaParcRec"
        colCampos.Add "iTipo"
        
        Call Ordena_Colecao(colImpCoinfoArqBaixa, colSaida, colCampos)
        
        For Each objImpCoinfoArqBaixa In colSaida
        
            sNumCCred = String(255, 0)
        
            lErro = Comando_Executar(alComando(7), "SELECT NumVou, CartaoCompleto FROM CoinfoItemFaturaImport2 WHERE NumIntParc = ?", lNumVou, sNumCCred, objImpCoinfoArqBaixa.lNumIntParcela)
            If lErro <> AD_SQL_SUCESSO Then gError 182834
        
            lErro = Comando_BuscarPrimeiro(alComando(7))
            If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182835
            
            objImpCoinfoArqBaixa.sNumCC = sNumCCred
            objImpCoinfoArqBaixa.lNumVou = lNumVou
        
        Next
        
        Open sDiretorio & sArquivo1 For Output As #1
        bArq1Aberto = True
        
        For Each objImpCoinfoArqBaixa In colSaida
            Print #1, FormataCpoData(objImpCoinfoArqBaixa.dtData) & objImpCoinfo.sSiglaEmpresa & FormataCpoNum(objImpCoinfoArqBaixa.sNumCC, 20) & FormataCpoNum(objImpCoinfoArqBaixa.lNumVou, 10) & objImpCoinfoArqBaixa.iTipo & FormataCpoNum(objImpCoinfoArqBaixa.iNumParcela, 2) & FormataCpoValor(objImpCoinfoArqBaixa.dValor, 10)
        Next
        
        Close #1
        bArq1Aberto = False
        
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 182768
        
        lErro = Comando_Executar(alComando(5), "INSERT INTO CoinfoArqsGerados (DataRef, TipoArquivo, NomeArquivo, DataGeracao) VALUES (?,?,?,?)", dtData, TIPO_ARQ_BAIXAS_CCANTIGO, sArquivo1, gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 182769
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 182770
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Coinfo_Gera_Arquivos3 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Gera_Arquivos3:

    Coinfo_Gera_Arquivos3 = gErr
     
    Select Case gErr

        Case 182750
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 182751 To 182756
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASREC", gErr)
            
        Case 182760 To 182763
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSGERADOS", gErr)
        
         Case 182765, 182768
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182766, 182769
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOARQSGERADOS", gErr)
            
        Case 182767, 182770
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 185057
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182757)
     
    End Select

    If bArq1Aberto Then Close #1
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Coinfo_AtualizarEmissores(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza os fornecedores

Dim lErro As Long
Dim lTransacaoExt As Long, lTransacao As Long
Dim alComandoExt(0 To 2) As Long, alComando(1 To 3) As Long, alCmdForaTrans(1 To 1) As Long
Dim iIndice As Integer, lNumIntArq As Long
Dim objForn As ClassFornecedor, objFornAux As ClassFornecedor
Dim objEnderecoForn As ClassEndereco, lNumRegArqTexto As Long
Dim bAchou As Boolean, iTotalRegCertos As Integer
Dim iTotalRegErr As Integer, lCodEmi As Long, sapelido As String
Dim sNome As String, lcodemp As Long, sEmail As String
Dim sObs As String, lRegsLidos As Long, lAux As Long
Dim snchave As String, sNomeRedAlternativo As String, objCli As ClassCliente
Dim objFilialForn As New ClassFilialFornecedor, colCategoriaItemForn As New Collection
Dim sChave As String

On Error GoTo Erro_Coinfo_AtualizarEmissores

    'Abre o comando para o arquivo
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 130729
    Next
    
    'Abre o comando para as leituras fora de transação
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        alCmdForaTrans(iIndice) = Comando_AbrirExt(glConexao3)
        If alCmdForaTrans(iIndice) = 0 Then gError 130729
    Next
        
    'Obtém o total de registros que vão ser atualizados
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpemissor")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao1)
    If lTransacaoExt = 0 Then gError 130730
    
    'Lê os arquivos que não foram atualizados
    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fpemissor", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 130732
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130733
    
    Do While lErro = AD_SQL_SUCESSO
    
        'inicializar buffers de leitura p/strings
        sapelido = String(255, 0)
        sNome = String(255, 0)
        sEmail = String(255, 0)
        sObs = String(255, 0)
        
        'Le o emissor
        lErro = Comando_Executar(alCmdForaTrans(1), "SELECT NumRegArqTexto, CodEmi, apelido, nome, codemp, email,obs  FROM CoinfoEmissor WHERE NumIntArq = ? AND NoCorporatorEm = ? ORDER BY NumRegArqTexto", _
            lNumRegArqTexto, lCodEmi, sapelido, sNome, lcodemp, sEmail, sObs, lNumIntArq, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 130734
        
        'Busca o primeiro
        lErro = Comando_BuscarProximo(alCmdForaTrans(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130735
        
        'Enquando existirem registros
        Do While lErro = AD_SQL_SUCESSO
        
            sChave = "Emissor: " & CStr(lCodEmi) & " Nome: " & sNome & " Empresa: " & CStr(lcodemp)
        
            lRegsLidos = lRegsLidos + 1
            
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 130730
            
            'Abre o comando para os registros
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 130729
            Next
            
            snchave = sapelido

            'se já houver um cliente com este nome reduzido mas com codigo diferente entao concatenar com o codigo da empresa
            lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM Fornecedores WHERE Codigo <> ? AND NomeReduzido = ?", lAux, lCodEmi + FATOR_SOMA_COD_EMISSOR, snchave)
            If lErro <> AD_SQL_SUCESSO Then gError 130752
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130753
            
            sNomeRedAlternativo = snchave
            If lErro = AD_SQL_SUCESSO Then
                'duplicidade de nome reduzido
                sNomeRedAlternativo = left(snchave, 40) & " (" & CStr(lCodEmi + FATOR_SOMA_COD_EMISSOR) & ")"
            End If
            
            Set objForn = New ClassFornecedor
            objForn.lCodigo = lCodEmi + FATOR_SOMA_COD_EMISSOR
            
            'faz a leitura para carregar info já existente que nao tenha vindo na importacao
            lErro = CF("Fornecedor_Le", objForn)
            If lErro <> SUCESSO And lErro <> 12729 Then gError 130738
            
            If lErro <> SUCESSO Then
            
                'atribuir os campos vindos do arquivo texto
                Set objEnderecoForn = New ClassEndereco
                Set objFornAux = New ClassFornecedor
                objFornAux.lCodigo = lcodemp
                
                'faz a leitura para carregar info já existente que nao tenha vindo na importacao
                lErro = CF("Fornecedor_Le", objFornAux)
                If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
            
                objEnderecoForn.lCodigo = objFornAux.lEndereco
                'Lê os dados do endereço do Fornecedor
                lErro = CF("Endereco_Le", objEnderecoForn)
                If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                
                If lErro <> SUCESSO Then
                    objEnderecoForn.iCodigoPais = PAIS_BRASIL
                    objEnderecoForn.sSiglaEstado = "SP"
                End If
            
                objEnderecoForn.sEmail = left(sEmail, 50)
                
            Else
            
                Set objFilialForn = New ClassFilialFornecedor
                Set colCategoriaItemForn = New Collection
                
                'Lê as categoria deste fornecedor
                objFilialForn.lCodFornecedor = objForn.lCodigo
                objFilialForn.iCodFilial = FILIAL_MATRIZ
                 
                'Alterado em 03/03/04 por Luiz Nogueira
                lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilialForn, colCategoriaItemForn)
                If lErro <> SUCESSO And lErro <> 127085 Then gError ERRO_SEM_MENSAGEM
                 
                Set objForn.colCategoriaItem = colCategoriaItemForn
            
                'atribuir os campos vindos do arquivo texto
                Set objEnderecoForn = New ClassEndereco
                
                objEnderecoForn.lCodigo = objForn.lEndereco
                'Lê os dados do endereço do Fornecedor
                lErro = CF("Endereco_Le", objEnderecoForn)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                objEnderecoForn.sEmail = left(sEmail, 50)
                
            End If
            
            With objForn
                .sNomeReduzido = left(sNomeRedAlternativo, STRING_FORNECEDOR_NOME_REDUZIDO)
                If Len(Trim(sNome)) = 0 Then
                    .sRazaoSocial = left(sNomeRedAlternativo, STRING_FORNECEDOR_RAZAO_SOC)
                Else
                    .sRazaoSocial = left(sNome, STRING_FORNECEDOR_RAZAO_SOC)
                End If
                .iTipo = 11
                .sObservacao = left(sObs, STRING_FORNECEDOR_OBS)
                .iAtivo = FORN_ATIVO
            End With
                        
            lErro = CF("Fornecedor_Grava_EmTrans", objForn, objEnderecoForn)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set objCli = New ClassCliente
     
            If lcodemp <> 0 Then
            
                objCli.lCodigo = lcodemp
            
                'faz a leitura para carregar info já existente que nao tenha vindo na importacao
                lErro = CF("Cliente_Le", objCli)
                If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                
                lErro = CF("TRVCliente_Grava", objCli)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
            
            iTotalRegCertos = iTotalRegCertos + 1
            
            lErro = Comando_Executar(alComando(1), "UPDATE CoinfoEmissor SET NoCorporatorEm = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", gdtDataHoje, lNumIntArq, lNumRegArqTexto)
            If lErro <> SUCESSO Then gError 130739
            
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
            
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 130743
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_BuscarProximo(alCmdForaTrans(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130740
        
        Loop
        
        lErro = Comando_ExecutarPos(alComandoExt(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComandoExt(0), gdtDataHoje)
        If lErro <> SUCESSO Then gError 130741
    
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130742
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 130743
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Emissores (Fornecedores) foram atualizados com sucesso. " & CStr(iTotalRegCertos) & " registros." & vbNewLine
        
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
        
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        Call Comando_Fechar(alCmdForaTrans(iIndice))
    Next
    
    Coinfo_AtualizarEmissores = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarEmissores:

    Coinfo_AtualizarEmissores = gErr
    
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case 130729
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130730
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130743
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130732, 130733, 130742
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130741
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130734, 130735, 130740, 130747, 130748, 130751, 130752, 130753
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOEMISSOR", gErr)
        
        Case 130731, 130744
            Call Rotina_Erro(vbOKOnly, "ERRO_REATIVAR_LOCKS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130557)
     
    End Select
     
    Call Rotina_Erro(vbOKOnly, "ERRO_TRV_ATUALIZACAO_ARQUIVO", gErr, "Emissor", lNumIntArq, lNumRegArqTexto, sChave)
     
    Call Transacao_RollbackExt(lTransacaoExt)
    Call Transacao_Rollback
    
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        Call Comando_Fechar(alCmdForaTrans(iIndice))
    Next
        
    Exit Function

End Function

'Private Function Coinfo_AtualizarOver(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''Gera comissões over como Títulos a Pagar Tipo OVER e número igual a Numpgcm
''Numpgcm é o número dado pelo Sigav ao grupamento de N vouchers e estorno de vourchers
''que vão ser pagos juntos. Os vouchers estornados vem com valores negativos no próximo
''pagamento ou no mesmo caso ele ainda não tenha sido pago.
'
''A liberação do over ocorre ou na baixa ou na emissão do títulos receber associado
''Os fornecedores cuja liberação é na emissão (exceção) tem uma Categoria chamada Over setada
''com o item Emissão
'
''O vencimento é na sexta da semana seguinte a liberação
'
''2006-12-04
''Fredy Willian Nowaski: 14:04:57 A emissão da comissão over é na data de recebimento da fatura,
''com vencimento na sexta feira da semana seguinte, exceto às agências do RJ que a Denyse te passou.
'
''2006-12-04
''Wagner Luis Carvalhosa de Mello: 19:01:38 O que a Denyse que passou é que vocês tem um relatório
''que pega tudo que ainda não foi pago para um emissor e agrupa dando uma numeração (numpgcm) ao
''emitir o relatório. Pelo que eu vi os estornos já são levados em conta na hora do sigav gerar
''outro grupamento, ou seja, além do que se tem a pagar e abate o que foi estornado. Eu quero
''saber se é mesmo assim que funciona....
''Andre Tatsuhiko Korosue: 19:02:43 É assim que funciona ...
'
'Dim lErro As Long
'Dim lTransacaoExt As Long
'Dim lTransacao As Long
'Dim alComandoExt(0 To 8) As Long
'Dim alComando(0 To 3) As Long
'Dim iIndice As Integer
'Dim lNumIntArq As Long
'Dim lNumRegArqTexto As Long
'Dim bTemErro As Boolean
'Dim iTotalRegCertos As Integer
'Dim iTotalRegErr As Integer
'Dim dtDatEmi As Date
'Dim lCodEmi As Long
'Dim lNumVou As Long
'Dim dtdataven As Date
'Dim dValor As Double
'Dim iPago As Integer
'Dim dtDataPag As Date
'Dim ieliminado As Integer
'Dim iAchou As Integer
'Dim iFilialCorporator As Integer
'Dim iFilialCoinfo As Integer
'Dim objTitPag As ClassTituloPagar
'Dim objParcelaPagar As ClassParcelaPagar
'Dim colParcelaPagar As colParcelaPagar
'Dim objContabil As ClassContabil
'Dim colBaixaParcPagar As ColBaixaParcRec
'Dim objBaixaPagar As ClassBaixaPagar
'Dim objImpCoInfoLog As ClassImpCoinfoLogAtu
'Dim colImpCoInfoLog As Collection
'Dim sCritica As String
'Dim objCtb As ClassImpCoinfoFatCPCtb
'Dim iComErro As Integer
'Dim objContabAutomatica As ClassContabAutomatica
'Dim bErro As Boolean
'Dim lNumIntDocTit As Long, lFornecedorTit As Long, dtDataEmissaoTit As Date
'Dim iStatusTit As Integer, dValorTit As Double, dSaldoTit As Double
'Dim lcodemp As Long
'Dim objNF As ClassNFiscal
'Dim lNumIntDocCredito As Long
'Dim iTabela As Integer
'Dim lNumIntDocTab As Long
'Dim dtDataVencimento As Date
'Dim bPendencia As Boolean
'Dim lNumIntDocTitRec As Long
'Dim dtDataBaixaTitRec As Date
'Dim dtDataEmissaoTitRec As Date
'Dim dtDataAux As Date
'Dim sItem As String
'Dim iStatusTitRec As Integer
'Dim sTipoTitRec As String
'Dim bTipoAntigo As Boolean
'Dim lnumpgcm As Long
'Dim lNumpgcmAnt As Long
'Dim dvalorAux As Double
'Dim bTerminarExecucao As Boolean
'Dim bUltimoRegistro As Boolean
'Dim colcolOver As New Collection
'Dim colOver As Collection
'Dim objImpCoinfoOver As ClassImpCoinfoComisOver
'Dim sSerie As String
'Dim lrecnsica As Long
'Dim sSerieAnt As String
'Dim lRecnsicaAnt As Long
'Dim lNumVouAnt As Long
'Dim sTipVou As String
'
'On Error GoTo Erro_Coinfo_AtualizarOver
'
'    'Abre o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
'        If alComandoExt(iIndice) = 0 Then gError 182508
'    Next
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 182509
'    Next
'
'    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpover")
'    If lErro <> SUCESSO Then gError 182714
'
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 182510
'
'    'Le os arquivos não atualizados
'    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
'        lNumIntArq, "fpover", DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 182511
'
'    'Busca o Primeiro
'    lErro = Comando_BuscarProximo(alComandoExt(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182512
'
'    'Enquando existirem arquivos
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        bTemErro = False
'        Set colcolOver = New Collection
'
'        sSerie = String(255, 0)
'        sTipVou = String(255, 0)
'
'        'Le as comissões over
'        lErro = Comando_ExecutarPos(alComandoExt(1), "SELECT TipVou, Recnsica, Serie, Numpgcm, NumRegArqTexto,DatEmi,CodEmi,NumVou,DataVen,Valor,Pago,DataPag,Eliminado FROM CoinfoComissoesOver WHERE NumIntArq = ? AND codest = ? AND NoCorporatorEm = ? AND DatEmi > ? AND Numpgcm <> 0 ORDER BY Numpgcm,Recnsica, Serie, NumVou, NumRegArqTexto", 0, _
'            sTipVou, lrecnsica, sSerie, lnumpgcm, lNumRegArqTexto, dtDatEmi, lCodEmi, lNumVou, dtdataven, dValor, iPago, dtDataPag, ieliminado, lNumIntArq, objImpCoinfo.sSiglaEmpresa, DATA_NULA, #1/1/2008#)
'        If lErro <> AD_SQL_SUCESSO Then gError 182532
'
'        'Busca a primeira
'        lErro = Comando_BuscarProximo(alComandoExt(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'        'Enquanto houver comissões não atualizadas
'        Do While lErro <> AD_SQL_SEM_DADOS
'
'            'Em versões posteriores a versão de faturamento somente over emitidos no Sigav anteriores
'            'a data de início do comissionamento pelo Corporator serão considerados
'            If giVersaoTRV = TRV_VERSAO_FATURAMENTO Or dtDatEmi < gdtDataInicioComisCorp Then
'
'                'Se não for o mesmo item
'                If lNumVou <> lNumVouAnt Or sSerie <> sSerieAnt Or lrecnsica <> lRecnsicaAnt Or lnumpgcm <> lNumpgcmAnt Then
'
'                    'Se for um novo grupamento
'                    If lnumpgcm <> lNumpgcmAnt Then
'                        Set colOver = New Collection
'                        colcolOver.Add colOver
'                        lNumpgcmAnt = lnumpgcm
'                    End If
'
'                    Set objImpCoinfoOver = New ClassImpCoinfoComisOver
'
'                    objImpCoinfoOver.lrecnsica = lrecnsica
'                    objImpCoinfoOver.lNumIntArq = lNumIntArq
'                    objImpCoinfoOver.dtDataPag = dtDataPag
'                    objImpCoinfoOver.dtdataven = dtdataven
'                    objImpCoinfoOver.dtDatEmi = dtDatEmi
'                    objImpCoinfoOver.dValor = dValor
'                    objImpCoinfoOver.iPago = iPago
'                    objImpCoinfoOver.ieliminado = ieliminado
'                    objImpCoinfoOver.sSerie = sSerie
'                    objImpCoinfoOver.lnumpgcm = lnumpgcm
'                    objImpCoinfoOver.lNumRegArqTexto = lNumRegArqTexto
'                    objImpCoinfoOver.lCodEmi = lCodEmi
'                    objImpCoinfoOver.lNumVou = lNumVou
'                    objImpCoinfoOver.sTipVou = sTipVou
'
'                    colOver.Add objImpCoinfoOver
'
'                    lNumVouAnt = lNumVou
'                    sSerieAnt = sSerie
'                    lRecnsicaAnt = lrecnsica
'
'                End If
'
'            End If
'
'            'Busca o próximo
'            lErro = Comando_BuscarProximo(alComandoExt(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
'
'        Loop
'
'        'Para cada grupamento de vouchers -> Gera um título a pagar
'        For Each colOver In colcolOver
'
'            'O Valor do Título é a soma dos vouchers com o valor final arredondado (inteiro)
'            dvalorAux = 0
'            For Each objImpCoinfoOver In colOver
'                dvalorAux = dvalorAux + objImpCoinfoOver.dValor
'            Next
'
'            dvalorAux = Round(dvalorAux, 0)
'
'            Set objImpCoinfoOver = colOver.Item(1)
'
'            iAchou = 0
'            bErro = False
'            bPendencia = False
'
'            'Le a empresa do emissor
'            lErro = Comando_Executar(alComandoExt(4), "SELECT CodEmp FROM CoinfoEmissor WHERE CodEmi = ? ORDER BY NumIntArq DESC,  NumRegArqTexto DESC", lcodemp, objImpCoinfoOver.lCodEmi)
'            If lErro <> AD_SQL_SUCESSO Then gError 182532
'
'            'Busca a primeira
'            lErro = Comando_BuscarProximo(alComandoExt(4))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'            'Se tem uma empresa associada
'            If lErro = AD_SQL_SUCESSO Then
'
'                'Busca a filial
'                lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", lcodemp, iFilialCorporator, iFilialCoinfo)
'                If lErro <> SUCESSO Then gError 182743
'
'            Else
'
'                iFilialCorporator = FILIAL_MATRIZ
'                iFilialCoinfo = FILIAL_MATRIZ
'
'            End If
'
'            'Para cada voucher do grupamento
'            For Each objImpCoinfoOver In colOver
'
'                bTipoAntigo = False
'
''                sTipoTitRec = String(STRING_SIGLA_DOCUMENTO, 0)
''
''                'Verifica se o título a receber associado a essa comissão existe
''                lErro = Comando_Executar(alComandoExt(5), "SELECT DISTINCT TR.NumIntDoc, TR.DataEmissao, TR.Status, TR.SiglaDocumento FROM CoinfoComissoesOver AS O, CoinfoItemFatura AS I, CoinfoFatura AS F, TitulosRecTodos AS TR " & _
''                                                            "WHERE O.CodEst = I.CodEst AND O.NumVou = I.NumVou AND O.Serie = I.Serie AND F.CodEst = I.CodEst AND F.NumFat = I.NumFat AND F.NumFat <> 0 AND " & _
''                                                            "F.TipoDocCorporator = ? AND F.NumIntDocCorporator <> 0 AND F.NumIntDocCorporator = TR.NumIntDoc AND O.CodEst = ? AND O.NumVou = ? AND O.Serie = ? ", lNumIntDocTitRec, dtDataEmissaoTitRec, iStatusTitRec, sTipoTitRec, TIPO_TITULO_RECEBER_TRV, objImpCoinfo.sSiglaEmpresa, objImpCoinfoOver.lNumVou, objImpCoinfoOver.sSerie)
''                If lErro <> AD_SQL_SUCESSO Then gError 182532
''
''                'Busca a primeira
''                lErro = Comando_BuscarProximo(alComandoExt(5))
''                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
''
''                'Se não encontrou => Verifica se o pagamento foi feito em cartão não tratado pela Coinfo
''                If lErro <> AD_SQL_SUCESSO Then
''
''                    'Verifica se é um título de cartão de crédito no formato antigo
''                    lErro = Comando_Executar(alComandoExt(8), "SELECT DISTINCT TR.NumIntDoc, TR.DataEmissao, TR.Status, TR.SiglaDocumento FROM CoinfoComissoesOver AS O, CoinfoItemFatura AS I, TitulosRecTodos AS TR " & _
''                                                                "WHERE O.CodEst = I.CodEst AND O.NumVou = I.NumVou AND O.Serie = I.Serie AND I.Cartao = 1 AND Data <= {d '2006-10-31'} AND " & _
''                                                                "I.TipoDocCorporator = ? AND I.NumIntDocCorporator <> 0 AND I.NumIntDocCorporator = TR.NumIntDoc AND O.CodEst = ? AND O.NumVou = ? AND O.Serie = ? ", lNumIntDocTitRec, dtDataEmissaoTitRec, iStatusTitRec, sTipoTitRec, TIPO_TITULO_RECEBER_TRV, objImpCoinfo.sSiglaEmpresa, objImpCoinfoOver.lNumVou, objImpCoinfoOver.sSerie)
''                    If lErro <> AD_SQL_SUCESSO Then gError 182532
''
''                    'Busca a primeira
''                    lErro = Comando_BuscarProximo(alComandoExt(8))
''                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
''
''                    If lErro = AD_SQL_SUCESSO Then bTipoAntigo = True
''
''                End If
'
'                sTipoTitRec = String(STRING_SIGLA_DOCUMENTO, 0)
'
'                'Verifica se o título a receber associado a essa comissão existe
'                lErro = Comando_Executar(alComandoExt(5), "SELECT DISTINCT TR.NumIntDoc, TR.DataEmissao, TR.Status, TR.SiglaDocumento FROM TRVVouchers AS V, TitulosRecTodos AS TR " & _
'                                                            "WHERE V.TipoDoc = 'VOU'  AND V.NumVou = ? AND V.Serie = ? AND V.TipVou = ? AND V.TipoDocDestino = 3 AND V.NumIntDocDestino = TR.NumIntDoc ", lNumIntDocTitRec, dtDataEmissaoTitRec, iStatusTitRec, sTipoTitRec, objImpCoinfoOver.lNumVou, objImpCoinfoOver.sSerie, objImpCoinfoOver.sTipVou)
'                If lErro <> AD_SQL_SUCESSO Then gError 182532
'
'                'Busca a primeira
'                lErro = Comando_BuscarProximo(alComandoExt(5))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'                If lErro = AD_SQL_SUCESSO Then Exit For
'
'            Next
'
'            'Se não encontrou seta o objImpCoinfoOver como o primeiro voucher
'            If objImpCoinfoOver Is Nothing Then Set objImpCoinfoOver = colOver.Item(1)
'
'            'Não existe o título a receber associado a essa comissão no corporator => Marca como pendente
'            If lErro <> AD_SQL_SUCESSO Then
'
'                bPendencia = True
'                dtDataEmissaoTitRec = DATA_NULA
'                dtDataBaixaTitRec = DATA_NULA
'
'            Else
'
'                'Verifica se houve alguma baixa
'                lErro = Comando_Executar(alComandoExt(6), "SELECT MIN(BR.Data) FROM TitulosRecTodos AS TR, ParcelasRecTodas AS PR, BaixasRec AS BR, BaixasParcRec AS BPR " & _
'                                                            "WHERE PR.NumIntTitulo = TR.NumIntDoc AND BPR.NumIntParcela = PR.NumIntDoc AND BR. NumIntBaixa = BPR.NumIntBaixa AND " & _
'                                                            "BPR.Status <> ? AND TR.NumIntDoc = ? ", dtDataBaixaTitRec, STATUS_CANCELADO, lNumIntDocTitRec)
'                If lErro <> AD_SQL_SUCESSO Then gError 182532
'
'                'Busca a primeira
'                lErro = Comando_BuscarProximo(alComandoExt(6))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'                'Se já houve baixa no a receber a comissão tem que ser liberada
'                If lErro = AD_SQL_SUCESSO And dtDataBaixaTitRec <> DATA_NULA Then
'                    bPendencia = False
'                Else
'                    'Se não houve baixa a comissão só será liberada se o fornecedor
'                    'estiver com a categoria over marcada para pagamento na emissão ou se
'                    'for um título de cartão emitido a mais de 30 dias
'                    bPendencia = True
'                    dtDataBaixaTitRec = DATA_NULA
'
'                    'Se for um título de cartão gerado da forma antiga, mesmo que não haja baixa
'                    'Se já passou um mês da emissão, libera o título
'                    If bTipoAntigo Then
'
'                        If gdtDataAtual >= DateAdd("d", 30, dtDataEmissaoTitRec) Then
'                            bPendencia = False
'                            dtDataBaixaTitRec = DateAdd("d", 30, dtDataEmissaoTitRec)
'                        End If
'
'                    End If
'                End If
'
'            End If
'
'            sItem = String(255, 0)
'
'            'Verifica o tipo de tratamento (Se o pagamento é na baixa ou na emissão)
'            lErro = Comando_Executar(alComandoExt(7), "SELECT Item FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Filial = ? AND Categoria = ? ", sItem, lcodemp, FILIAL_MATRIZ, "OVER")
'            If lErro <> AD_SQL_SUCESSO Then gError 182532
'
'            'Busca a primeira
'            lErro = Comando_BuscarProximo(alComandoExt(7))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
'
'            'Se não encontrou o tipo de tratamento então é porque é na baixa
'            If lErro <> AD_SQL_SUCESSO Then
'                sItem = "Baixa"
'            End If
'
'            'Se o tratamento é na baixa seta a data (Se não houve a baixa => bPendencia = True )
'            If InStr(1, UCase(sItem), "BAIXA") <> 0 Then
'                dtDataAux = dtDataBaixaTitRec
'            Else
'                'Se for na emissão não pode ficar pendente se tiver o título a receber associado
'                If dtDataEmissaoTitRec <> DATA_NULA Then
'                    dtDataAux = dtDataEmissaoTitRec
'                    bPendencia = False
'                End If
'            End If
'
'            Set colImpCoInfoLog = New Collection
'            Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
'
'            objImpCoInfoLog.iTabela = TABELA_COINFOFATURA
'            objImpCoInfoLog.lNumIntArq = lNumIntArq
'            objImpCoInfoLog.lNumRegArqTexto = objImpCoinfoOver.lNumRegArqTexto
'
'            'buscar a fatura como titulo a Pagar aberto
'            lErro = Comando_Executar(alComandoExt(2), "SELECT NumIntDoc, Fornecedor, DataEmissao, Status, ValorTotal, Saldo FROM TitulosPag WHERE FilialEmpresa = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ?", _
'                lNumIntDocTit, lFornecedorTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, TIPODOC_FATURA_OVER, objImpCoinfoOver.lnumpgcm, STATUS_EXCLUIDO)
'            If lErro <> AD_SQL_SUCESSO Then gError 182534
'
'            'Busca o primeiro
'            lErro = Comando_BuscarProximo(alComandoExt(2))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182535
'
'            If lErro <> AD_SQL_SUCESSO Then
'
'                'buscar a fatura como titulo a Pagar baixado
'                lErro = Comando_Executar(alComandoExt(3), "SELECT NumIntDoc, Fornecedor, DataEmissao, Status, ValorTotal, Saldo FROM TitulosPagBaixados WHERE FilialEmpresa = ? AND SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? ", _
'                    lNumIntDocTit, lFornecedorTit, dtDataEmissaoTit, iStatusTit, dValorTit, dSaldoTit, iFilialCorporator, TIPODOC_FATURA_OVER, objImpCoinfoOver.lnumpgcm, STATUS_EXCLUIDO)
'                If lErro <> AD_SQL_SUCESSO Then gError 182536
'
'                'Busca o primeiro
'                lErro = Comando_BuscarProximo(alComandoExt(3))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182537
'                If lErro = AD_SQL_SUCESSO Then iAchou = 2
'
'            Else
'                iAchou = 1
'            End If
'
'            'Título já pago (baixado) não pode ser alterado
'            'If iAchou = 2 Then gError 182532
'            If iAchou = 2 Then
'                'bErro = True
'                Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_OVER & " e Número " & CStr(objImpCoinfoOver.lnumpgcm) & " já está baixado e não pode ser alterado.", objImpCoInfoLog, colImpCoInfoLog, True)
'            End If
'
'             'Abertura de transação
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 182538
'
'            'preencher o objeto e mandar gravar
'            Set objTitPag = New ClassTituloPagar
'
'            If iAchou <> 0 Then
'                objTitPag.lNumIntDoc = lNumIntDocTit
'            End If
'
'            lNumIntDocTab = objTitPag.lNumIntDoc
'            iTabela = TIPO_TITULO_PAGAR_TRV
'
'            'se não existir no sistema e não for um estorno e não estiver pendente
'            If iAchou = 0 And Not bErro And dvalorAux > 0 And Not bPendencia Then
'
'                dtDataAux = Date
'
'                objTitPag.lFornecedor = objImpCoinfoOver.lCodEmi + FATOR_SOMA_COD_EMISSOR
'                objTitPag.iFilial = FILIAL_MATRIZ
'                objTitPag.iFilialPedCompra = FILIAL_MATRIZ
'                objTitPag.lNumTitulo = objImpCoinfoOver.lnumpgcm
'                objTitPag.dValorTotal = dvalorAux
'                objTitPag.dValorProdutos = dvalorAux
'                objTitPag.dSaldo = dvalorAux
'                objTitPag.iFilialEmpresa = iFilialCorporator
'                objTitPag.sSiglaDocumento = TIPODOC_FATURA_OVER
'                objTitPag.iNumParcelas = 1
'
'                'O Vencimento ocorre na sexta da semana seguinte
'                Select Case Weekday(dtDataAux)
'
'                    Case vbMonday
'                        dtDataVencimento = DateAdd("d", 11, dtDataAux)
'                    Case vbTuesday
'                        dtDataVencimento = DateAdd("d", 10, dtDataAux)
'                    Case vbWednesday
'                        dtDataVencimento = DateAdd("d", 9, dtDataAux)
'                    Case vbThursday
'                        dtDataVencimento = DateAdd("d", 8, dtDataAux)
'                    Case vbFriday
'                        dtDataVencimento = DateAdd("d", 7, dtDataAux)
'                    Case vbSaturday
'                        dtDataVencimento = DateAdd("d", 6, dtDataAux)
'                    Case vbSunday
'                        dtDataVencimento = DateAdd("d", 5, dtDataAux)
'
'                End Select
'
'                objTitPag.dtDataEmissao = dtDataAux
'
'                Set objParcelaPagar = New ClassParcelaPagar
'                Set colParcelaPagar = New colParcelaPagar
'
'                objParcelaPagar.iNumParcela = 1
'                objParcelaPagar.dtDataVencimento = dtDataVencimento
'                objParcelaPagar.dtDataVencimentoReal = dtDataVencimento
'                objParcelaPagar.dValor = dvalorAux
'                objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
'                objParcelaPagar.iStatus = STATUS_ABERTO
'
'                With objParcelaPagar
'                    colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
'                End With
'
'                'Grava o Título a Pagar
'                lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
'                If lErro = SUCESSO Then
'
'                    If objTitPag.dtDataEmissao > #10/31/2006# Then
'
'                        Set objContabAutomatica = New ClassContabAutomatica
'                        Set objCtb = New ClassImpCoinfoFatCPCtb
'                        objCtb.sModelo = "Comissão Over"
'
'                        'gravar a contabilizacao
'                        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
'                            If lErro = SUCESSO Then
'
'                                lErro = objContabAutomatica.Finaliza_Contab()
'
'                            Else
'                                bErro = True
'                                Call CoInfo_Adiciona_Log("Erro na contabilização do Título a Pagar Tipo " & TIPODOC_FATURA_OVER & " e Número " & CStr(lNumVou) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                            End If
'
'                        Else
'                            bErro = True
'                            Call CoInfo_Adiciona_Log("Erro na contabilização do Título a Pagar Tipo " & TIPODOC_FATURA_OVER & " e Número " & CStr(lNumVou) & ".", objImpCoInfoLog, colImpCoInfoLog)
'
'                        End If
'
'                        Set objContabAutomatica = Nothing
'                        Call objCtb.Encerrando
'                        Set objCtb = Nothing
'
'                    End If
'
'                    lNumIntDocTab = objTitPag.lNumIntDoc
'                    iTabela = TIPO_TITULO_PAGAR_TRV
'
'                End If
'
'                If lErro <> SUCESSO Then
'                    bErro = True
'                    Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_OVER & " e Número " & CStr(lNumVou) & " não pode ser gravado.", objImpCoInfoLog, colImpCoInfoLog)
'                Else
'                    Call CoInfo_Adiciona_Log("O Título a Pagar Tipo " & TIPODOC_FATURA_OVER & " e Número " & CStr(lNumVou) & " foi gravado com sucesso.", objImpCoInfoLog, colImpCoInfoLog, True)
'                End If
'
'            End If
'
'            'Insere o Log de erros
'            lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
'            If lErro <> SUCESSO Then gError 182544
'
'            'Atualiza a nota para informar que ela foi importada com sucesso
'            If iComErro = DESMARCADO And Not bPendencia Then
'
'                'Para cada voucher
'                For Each objImpCoinfoOver In colOver
'                    iTotalRegCertos = iTotalRegCertos + 1
'                    lErro = Comando_Executar(alComando(2), "UPDATE CoinfoComissoesOver SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ?, ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iTabela, lNumIntDocTab, gdtDataHoje, iComErro, objImpCoinfoOver.lNumIntArq, objImpCoinfoOver.lNumRegArqTexto)
'                    If lErro <> SUCESSO Then gError 182540
'                Next
'            Else
'                If iComErro = MARCADO Then
'                    bTemErro = True
'                Else
'                    iComErro = 2
'                End If
'
'                'Desfaz Transação para cancelar as coisas pela metade
'                Call Transacao_Rollback
'
'                'Abertura de transação só para marcar que houve erro
'                lTransacao = Transacao_Abrir()
'                If lTransacao = 0 Then gError 182538
'
'                'Para cada voucher
'                For Each objImpCoinfoOver In colOver
'
'                    If iComErro = MARCADO Then
'                        iTotalRegErr = iTotalRegErr + 1
'                    End If
'                    lErro = Comando_Executar(alComando(3), "UPDATE CoinfoComissoesOver SET ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iComErro, objImpCoinfoOver.lNumIntArq, objImpCoinfoOver.lNumRegArqTexto)
'                    If lErro <> SUCESSO Then gError 182540
'                Next
'
'            End If
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 182541
'
'            'Para cada voucher
'            For Each objImpCoinfoOver In colOver
'                lErro = Rotina_Atualizar_AtualizaTelaBatch
'                If lErro <> SUCESSO Then gError 182720
'            Next
'        Next
'
'        'Não atualiza o Arquivo e não busca o próximo
'        If bTemErro Then Exit Do
'
'        'Atualiza o arquivo
'        lErro = Comando_ExecutarPos(alComandoExt(1), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComandoExt(0), gdtDataHoje)
'        If lErro <> SUCESSO Then gError 182515
'
'        'Busca o próximo
'        lErro = Comando_BuscarProximo(alComandoExt(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182516
'
'    Loop
'
'    'Desfaz a Transação
'    Call Transacao_RollbackExt(lTransacaoExt)
'    Call Transacao_Rollback
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Pagar (Over) foram atualizados. " & CStr(iTotalRegCertos) & " registros com sucesso e " & CStr(iTotalRegErr) & " registros com erro." & vbNewLine
'
'    'Fecha o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    'Fecha o comando
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    Coinfo_AtualizarOver = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarOver:
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Títulos a Pagar (Over) foram atualizados. " & CStr(iTotalRegCertos) & " registros com sucesso e " & CStr(iTotalRegErr) & " registros com erro." & vbNewLine
'
'    Coinfo_AtualizarOver = gErr
'
'    Select Case gErr
'
'        Case 182508, 182509
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 182510
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 182511, 182512, 182516
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 182513, 182514, 182714
'
'        Case 182515
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
'
'        Case 182517
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182518)
'
'    End Select
'
'    Call Transacao_RollbackExt(lTransacaoExt)
'
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Coinfo_GerarComissaoCartao(ByVal objImpCoinfo As ClassImpCoinfo, ByVal objTitRec As ClassTituloReceber, ByVal objTitRecTRV As ClassTitulosRecTRV, Optional ByVal bVersaoNova As Boolean = False) As Long
''Gera um NFPag de comissão para um título a receber de cartão porque não vem na tabela de
''Faturas da Coinfo
'
''Vai em CoinfoComissao com o número e série dos vouches contidos na fatura
'
''2006-11-29
''Andre Tatsuhiko Korosue: 11:37:12 Quanto ao pagamento de comissão ao agente, você deverá
''gerar uma conta a pagar baseado na tabela comissão. Conforme conversado com o pessoal,
''o nosso sistema já cria o comissionamento a pagar no momento da emissão do voucher.
''Wagner Luis Carvalhosa de Mello: 11:38:26O que vocês criam não chega até nós ?
''Andre Tatsuhiko Korosue: 11:39:33 Não, diretamente. A informação do comissionamento a pagar
''está na tabela comissao e deverá ser trabalhado por vocês.
'
''2006-12-04
''Fredy Willian Nowaski: 14:08:24Então é diferente. A emissão da comissão é no momento da
''venda, porém o vencimento é na próxima sexta feira do efetivo recebimento por nós das
''operadoras de cartão.
''Wagner Luis Carvalhosa de Mello: 14:09:46Do recebimento da primeira parcela ou da última ?
''Fredy Willian Nowaski: 14:10:24Da primeira parcela. Pagaremos toda a comissão.
''Wagner Luis Carvalhosa de Mello: 14:11:19OK. Então o vencimento é após 30 dias a sexta
''seguinte, ok ?
''Wagner Luis Carvalhosa de Mello: 14:11:49Porque a primeira parcela é paga após 30 dias.
''Fredy Willian Nowaski: 14:12:27exato.
'
''++++++++++++> PS-> Para acertar o que já foi importado de Títulos de Cartão sem a parte
''de geração da comissão existe uma função chamada Coinfo_GerarComissaoRetroativo
'
'Dim lErro As Long
'Dim alComando(0 To 5) As Long
'Dim iIndice As Integer
'Dim sSerie As String
'Dim sSerieAnt As String
'Dim lNumVou As Long
'Dim lNumVouAnt As Long
'Dim objNFPag As ClassNFsPag
'Dim colNFPag As New Collection
'Dim dValor As Double
'Dim lcodemp As Long
'Dim objContabil As ClassContabil
'Dim objContabAutomatica As ClassContabAutomatica
'Dim objCtb As ClassImpCoinfoNFPagCtb
'Dim iAchou As Integer
'Dim lNumIntDocNF As Long
'Dim dtDataVencimento As Date
'Dim dtDataAux As Date
'Dim bAchou As Boolean
'Dim colComissoes As New Collection
'Dim objImpCoinfoComissao As ClassImpCoinfoComissao
'Dim lNumIntArq As Long
'Dim lNumRegArqTexto As Long
'Dim sTipoDoc As String
'Dim sTipoDocAnt As String
'Dim dValorComissao As Double
'Dim dtDataVou As Date
'Dim objTRVTitulos As ClassTRVTitulos
'
'On Error GoTo Erro_Coinfo_GerarComissaoCartao
'
'    'Abre os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 185040
'    Next
'
'    iAchou = 0
'
'    'buscar a fatura como NFPagar aberto
'    lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc, ValorTotal FROM NFsPag WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ?", _
'        lNumIntDocNF, dValorComissao, objTitRec.iFilialEmpresa, objTitRec.lNumTitulo, STATUS_EXCLUIDO)
'    If lErro <> AD_SQL_SUCESSO Then gError 182534
'
'    'Busca o primeiro
'    lErro = Comando_BuscarProximo(alComando(3))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182535
'
'    If lErro <> AD_SQL_SUCESSO Then
'
'        'buscar a fatura como titulo a Pagar baixado
'        lErro = Comando_Executar(alComando(4), "SELECT NumIntDoc, ValorTotal FROM NFsPagBaixadas WHERE FilialEmpresa = ? AND NumNotaFiscal = ? AND Status <> ? ", _
'            lNumIntDocNF, dValorComissao, objTitRec.iFilialEmpresa, objTitRec.lNumTitulo, STATUS_EXCLUIDO)
'        If lErro <> AD_SQL_SUCESSO Then gError 182536
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComando(4))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182537
'        If lErro = AD_SQL_SUCESSO Then
'            iAchou = 2
'        End If
'
'    Else
'        iAchou = 1
'    End If
'
'    'Se não é novo
'    If iAchou <> 0 Then
'
'        objTitRecTRV.lNumIntDocNFPagComi = lNumIntDocNF
'        objTitRecTRV.dValorComissao = dValorComissao
'
'    Else
'
'        sSerie = String(255, 0)
'        sTipoDoc = String(255, 0)
'
'        'Le os itens de fatura relacionados ao título a receber
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Serie, NumVou, TipoDoc, Data FROM TRVVouchers WHERE NumFatCoinfo = ? AND Cliente = ?", 0, sSerie, lNumVou, sTipoDoc, dtDataVou, objTitRec.lNumTitulo, objTitRec.lCliente)
'        If lErro <> AD_SQL_SUCESSO Then gError 185041
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185042
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            'Em versões posteriores a versão de faturamento somente vouchers emitidos no Sigav anteriores
'            'a data de início do comissionamento pelo Corporator serão considerados
'            If giVersaoTRV = TRV_VERSAO_FATURAMENTO Or dtDataVou < gdtDataInicioComisCorp Then
'
'                'Le as comissões associadas a esse item de fatura que não foram geradas
'                lErro = Comando_Executar(alComando(1), "SELECT C1.CodEmp, C1.ValComi,C1.NumIntArq,C1.NumRegArqTexto FROM CoinfoComissao AS C1 WHERE C1.NumVou = ? AND C1.Serie = ? AND C1.TipVou = ? AND C1.CodEmp <> 0 AND NOT EXISTS (SELECT C2.NumVou FROM CoinfoComissao AS C2 , NFsPag_Todas AS N WHERE N.Status <> 5 AND C2.TipoDocCorporator = 2 AND C2.NumIntDocCorporator=N.NumIntDoc AND C1.NumVou = C2.NumVou AND C1.Serie = C2.Serie AND C2.NoCorporatorEm <> ?) ORDER BY C1.NumIntArq DESC, C1.NumRegArqTexto DESC", _
'                lcodemp, dValor, lNumIntArq, lNumRegArqTexto, lNumVou, sSerie, sTipoDoc, DATA_NULA)
'                If lErro <> AD_SQL_SUCESSO Then gError 185043
'
'                'Busca a primeira
'                lErro = Comando_BuscarProximo(alComando(1))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185044
'
'                'Se encontrou
'                If lErro = AD_SQL_SUCESSO Then
'
'                    Set objImpCoinfoComissao = New ClassImpCoinfoComissao
'
'                    objImpCoinfoComissao.lNumIntArq = lNumIntArq
'                    objImpCoinfoComissao.lNumRegArqTexto = lNumRegArqTexto
'
'                    colComissoes.Add objImpCoinfoComissao
'
'                    'Se já existe uma comissão para esse emissor, atualiza só o valor
'                    bAchou = False
'                    For Each objNFPag In colNFPag
'                        If objNFPag.lFornecedor = lcodemp Then
'                            bAchou = True
'                            objNFPag.dValorProdutos = objNFPag.dValorProdutos + dValor
'                            objNFPag.dValorTotal = objNFPag.dValorTotal + dValor
'                            Exit For
'                        End If
'                    Next
'
'                    'Se não existe coloca na coleção
'                    If Not bAchou Then
'
'                        Set objNFPag = New ClassNFsPag
'
'                        objNFPag.lFornecedor = lcodemp
'                        objNFPag.iFilial = FILIAL_MATRIZ
'                        objNFPag.lNumNotaFiscal = objTitRec.lNumTitulo
'                        objNFPag.dtDataEmissao = objTitRec.dtDataEmissao
'                        objNFPag.dValorTotal = dValor
'                        objNFPag.dValorProdutos = dValor
'                        objNFPag.iFilialEmpresa = objTitRec.iFilialEmpresa
'
'                        dtDataAux = DateAdd("d", 30, objTitRec.dtDataEmissao)
'
'                        'O Vencimento ocorre até a sexta 30 dias após
'                        Select Case Weekday(dtDataAux)
'
'                            Case vbMonday
'                                dtDataVencimento = DateAdd("d", 4, dtDataAux)
'                            Case vbTuesday
'                                dtDataVencimento = DateAdd("d", 3, dtDataAux)
'                            Case vbWednesday
'                                dtDataVencimento = DateAdd("d", 2, dtDataAux)
'                            Case vbThursday
'                                dtDataVencimento = DateAdd("d", 1, dtDataAux)
'                            Case vbFriday
'                                dtDataVencimento = DateAdd("d", 7, dtDataAux)
'                            Case vbSaturday
'                                dtDataVencimento = DateAdd("d", 6, dtDataAux)
'                            Case vbSunday
'                                dtDataVencimento = DateAdd("d", 5, dtDataAux)
'
'                        End Select
'
'                        objNFPag.dtDataVencimento = dtDataVencimento
'
'                        colNFPag.Add objNFPag
'
'                    End If
'
'                    'Trava registro contra alterações/Leituras
'                    lErro = Comando_LockExclusive(alComando(0))
'                    If lErro <> SUCESSO Then gError 185042
'
'                    lErro = Comando_ExecutarPos(alComando(5), "UPDATE TRVVouchers SET ValorComissao = ? ", alComando(0), dValor)
'                    If lErro <> AD_SQL_SUCESSO Then gError 185042
'
'                End If
'
'            End If
'
'            'Busca o próximo
'            lErro = Comando_BuscarProximo(alComando(0))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185042
'
'        Loop
'
'        'Para cada nota gerada
'        For Each objNFPag In colNFPag
'
'            'Se existe valor a ser pago
'            If objNFPag.dValorTotal > DELTA_VALORMONETARIO Then
'
'                'Grava a NF Pagar
'                lErro = CF("NFPag_Grava_EmTrans", objNFPag, objContabil)
'                If lErro = SUCESSO Then
'
'                    objTitRecTRV.lNumIntDocNFPagComi = objNFPag.lNumIntDoc
'                    objTitRecTRV.dValorComissao = objNFPag.dValorTotal
'                    objNFPag.dtDataEmissao = Date
'
'                    If objNFPag.dtDataEmissao > #10/31/2006# Then
'
'                        Set objContabAutomatica = New ClassContabAutomatica
'                        Set objCtb = New ClassImpCoinfoNFPagCtb
'
'                        'gravar a contabilizacao
'                        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objNFPag.dtDataEmissao, objNFPag.dtDataEmissao)
'                        If lErro = SUCESSO Then
'
'                            lErro = objContabAutomatica.GeraContabilizacao(objNFPag, True)
'                            If lErro = SUCESSO Then
'
'                                lErro = objContabAutomatica.Finaliza_Contab()
'
'                            Else
'                                gError 189432
'
'                            End If
'
'                        Else
'                            gError 189432
'
'                        End If
'
'                        Set objContabAutomatica = Nothing
'                        Call objCtb.Encerrando
'                        Set objCtb = Nothing
'
'                    End If
'
'                End If
'                If lErro <> SUCESSO Then gError 185046
'
'                For Each objImpCoinfoComissao In colComissoes
'                    lErro = Comando_Executar(alComando(2), "UPDATE CoinfoComissao SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ?  WHERE NumIntArq = ? AND NumRegArqTexto = ? ", TIPO_NF_PAGAR_TRV, objNFPag.lNumIntDoc, gdtDataAtual, objImpCoinfoComissao.lNumIntArq, objImpCoinfoComissao.lNumRegArqTexto)
'                    If lErro <> AD_SQL_SUCESSO Then gError 185045
'                Next
'
'                Set objTRVTitulos = New ClassTRVTitulos
'
'                objTRVTitulos.iTipoDocOrigem = TRV_TIPO_DOC_DESTINO_NFSPAG
'                objTRVTitulos.lNumIntDocOrigem = objNFPag.lNumIntDoc
'                objTRVTitulos.sTipoDoc = TRV_TIPODOC_CMCC_TEXTO
'
'                lErro = CF("TRVTitulos_Insere", objTRVTitulos)
'                If lErro <> SUCESSO Then gError 196240
'
'            End If
'
'        Next
'
'    End If
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Coinfo_GerarComissaoCartao = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_GerarComissaoCartao:
'
'    Coinfo_GerarComissaoCartao = gErr
'
'    Select Case gErr
'
'        Case 185040
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 185041 To 185045
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 185046, 189432
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185047)
'
'    End Select
'
'    'Fecha o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Public Function Coinfo_AtualizaContabFaturasRec(ByVal objImpCoinfo As ClassImpCoinfo, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long
'Função auxiliar que não é chamada em nenhum ponto
'Refaz a contabilização em um determinado período

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 5) As Long
Dim iAchou As Integer
Dim lNumIntDocTit As Long
Dim objTitRec As ClassTituloReceber
Dim objCtb As ClassImpCoinfoFatCRCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil
Dim dComissao As Double
Dim objImpCoinfoCartoes As ClassImpCoinfoCartoes
Dim dvaldcbrl As Double
Dim dValor As Double
Dim sNome As String
Dim dTaxa As Double

On Error GoTo Erro_Coinfo_AtualizaContabFaturasRec

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    'buscar a fatura como titulo a receber no período
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM TitulosRecTodos WHERE DataEmissao BETWEEN ? AND ? AND SiglaDocumento IN (?,?)", _
        lNumIntDocTit, dtDataIni, dtDataFim, TIPODOC_CARTAO_CRED_DEB, TIPODOC_FATURA_SERVICO_CR)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver títulos
    Do While lErro <> AD_SQL_SEM_DADOS

         Set objTitRec = New ClassTituloReceber
         
         objTitRec.lNumIntDoc = lNumIntDocTit
         
         'Le o título
         lErro = CF("TituloReceber_Le", objTitRec)
         If lErro <> SUCESSO And lErro <> 26061 Then gError 182533
         If lErro <> SUCESSO Then
            'Lê o título baixado
            lErro = CF("TituloReceberBaixado_Le", objTitRec)
            If lErro <> SUCESSO And lErro <> 56570 Then gError 182533
         End If
         If lErro = SUCESSO Then
         
            'Lê o valor bruto e desconto
            lErro = Comando_Executar(alComando(2), "SELECT valor, valdcbrl FROM CoinfoFatura WHERE TipoDocCorporator = ? AND NumIntDocCorporator = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                dValor, dvaldcbrl, TIPO_TITULO_RECEBER_TRV, objTitRec.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
            
            'Busca a primeira
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
            
            If lErro <> AD_SQL_SEM_DADOS Then
    
                Set objContabil = New ClassContabil
                Call objContabil.Contabil_Inicializa_Contabilidade4(23, MODULO_CONTASARECEBER)
                 
                If Not (objContabil Is Nothing) Then
                         
                    'Exclui os lançamentos contábeis (se houverem) associados ao titulo em questao
                    lErro = objContabil.Contabil_Exclui(objTitRec.lNumIntDoc, CONTABIL_EXCLUI_EFETIVO, , , , objTitRec.iFilialEmpresa)
                    If lErro <> SUCESSO Then gError 83949
            
                    Set objContabil = Nothing
                End If
                
                If objTitRec.dtDataEmissao > #10/31/2006# Then
                
                    Set objContabAutomatica = New ClassContabAutomatica
                    Set objCtb = New ClassImpCoinfoFatCRCtb
                    
                    'gravar a contabilizacao
                    lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
                    If lErro = SUCESSO Then
                    
                        'Se não for cartão de crédito
                        If objTitRec.sSiglaDocumento <> TIPODOC_CARTAO_CRED_DEB Then
                            
                            'Coloca o desconto no COFINS
                            objTitRec.dCOFINSRetido = dvaldcbrl
                            
                            'Coloca a comissão no CSLL
                            Call Coinfo_Obtem_Comissao_Cliente(objTitRec.lNumTitulo, objImpCoinfo.sSiglaEmpresa, dComissao)
                            objTitRec.dCSLLRetido = dComissao
                        
                            'Gera a contabilização
                            lErro = objContabAutomatica.GeraContabilizacao(objTitRec, Nothing)
                        
                        Else 'Se é título de cartão
                            
                            sNome = String(255, 0)
                            
                            'Lê o item
                            lErro = Comando_Executar(alComando(3), "SELECT CiaCart FROM CoinfoItemFatura WHERE NumFat = ? AND CodEst = ? ORDER BY NumIntArq DESC, NumRegArqTexto DESC", _
                                sNome, objTitRec.lNumTitulo, objImpCoinfo.sSiglaEmpresa)
                            If lErro <> AD_SQL_SUCESSO Then gError 182534
                            
                            'Busca a primeira
                            lErro = Comando_BuscarProximo(alComando(3))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
                            
                            'Lê a taxa
                            lErro = Comando_Executar(alComando(4), "SELECT Taxa FROM AdmMeioPagtoCondPagto AS C, AdmMeioPagto AS M WHERE C.AdmMeioPagto = M.Codigo AND C.FilialEmpresa = M.FilialEmpresa AND M.Nome = ? AND C.FilialEmpresa = 1 AND Parcelamento = ?", _
                                dTaxa, sNome, objTitRec.iNumParcelas)
                            If lErro <> AD_SQL_SUCESSO Then gError 182534
                            
                            'Busca a primeira
                            lErro = Comando_BuscarProximo(alComando(4))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
                            
                            Set objImpCoinfoCartoes = New ClassImpCoinfoCartoes
                            
                            'Se for da amex
                            If sNome = "AX" Then
                                       
                                'Calcula a taxa
                                objImpCoinfoCartoes.dValorTaxa = dValor * dTaxa
                            
                                'A dedução é o que sobrou do valor bruto - o valor líquido - taxa
                                objImpCoinfoCartoes.dDeducaoValorPresente = dValor - objTitRec.dValor - (dValor * dTaxa)
                            
                            Else
                            
                                'Verifica se a difernça entre valor bruto e líquido é a taxa
                                If Abs(dValor - dTaxa * dValor - objTitRec.dValor) > DELTA_VALORMONETARIO Then
                                    gError 99999
                                End If
                                
                                'Não tem deduções
                                objImpCoinfoCartoes.dDeducaoValorPresente = 0
                                
                                'Calcula a taxa
                                objImpCoinfoCartoes.dValorTaxa = dValor - objTitRec.dValor
                            
                            End If
                        
                            'Gera a contabilização
                            lErro = objContabAutomatica.GeraContabilizacao(objTitRec, objImpCoinfoCartoes)
                        
                        End If
                        
                        If lErro = SUCESSO Then
                    
                            lErro = objContabAutomatica.Finaliza_Contab()
                            
                        Else
                        
                            gError 189433
            
                        End If
                        
                    Else
                    
                        gError 189433
                        
                    End If
        
                    Set objContabAutomatica = Nothing
                    Call objCtb.Encerrando
                    Set objCtb = Nothing
                
                End If
                  
            End If
            
        End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_AtualizaContabFaturasRec = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaContabFaturasRec:

    Coinfo_AtualizaContabFaturasRec = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_RECEBER", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743, 189433
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

'Private Function Coinfo_AtualizarComissao(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''atualiza os fornecedores
'
'Dim lErro As Long
'Dim lTransacaoExt As Long
'Dim lTransacao As Long
'Dim alComandoExt(0 To 4) As Long
'Dim alComando(0 To 4) As Long
'Dim iIndice As Integer
'Dim lNumIntArq As Long
'Dim objTitRec As ClassTituloReceber
'Dim sSelect As String
'Dim sSubSelect1 As String
'Dim sSubSelect2 As String
'Dim sWhere As String
'Dim lNumIntTitulo As Long
'Dim objTitRecTRV As ClassTitulosRecTRV
'
'On Error GoTo Erro_Coinfo_AtualizarComissao
'
'    'Abre o arquivo
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
'        If alComandoExt(iIndice) = 0 Then gError 185496
'    Next
'
'    'Abre o arquivo
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 185497
'    Next
'
'    'Obtém o total de registros que vão ser atualizados
'    lErro = Coinfo_Total_Registros(objImpCoinfo, "comissao")
'    If lErro <> SUCESSO Then gError 185498
'
'    'Buscar o NumIntDoc dos Títulos
'    sSelect = "SELECT DISTINCT I.NumIntTitulo FROM CoinfoItemFaturaImport2 AS I "
'
'    'Sendo títulos de cartão, dessa empresa e importados
'    sWhere = "WHERE I.SiglaDocumento = ? AND DataEmissao > ? AND I.CodEst = ? "
'
'    'e onde exista comissão
'    sSubSelect1 = "AND EXISTS (SELECT C.NumVou FROM CoinfoComissao AS C WHERE C.ValComi <> 0 AND C.Serie=I.Serie AND C.NumVou = I.NumVou AND C.NumIntArq = ?) "
'
'    'e onde não exista comissão que já gerou um a pagar para esse voucher
'    sSubSelect2 = "AND NOT EXISTS (SELECT C.NumVou FROM CoinfoComissao AS C WHERE C.Serie=I.Serie AND C.NumVou = I.NumVou AND C.NumIntDocCorporator <>0) "
'
'    sSelect = sSelect & sWhere & sSubSelect1 & sSubSelect2
'
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 185499
'
'    'Lê os arquivos que não foram atualizados
'    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
'        lNumIntArq, "comissao", DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 185500
'
'    'Busca o primeiro
'    lErro = Comando_BuscarProximo(alComandoExt(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185501
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Le o títulos que deveriam ter comissões e não as tem
'        lErro = Comando_Executar(alComandoExt(1), sSelect, lNumIntTitulo, TIPODOC_CARTAO_CRED_DEB, #10/31/2006#, objImpCoinfo.sSiglaEmpresa, lNumIntArq)
'        If lErro <> AD_SQL_SUCESSO Then gError 185502
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComandoExt(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185503
'
'        'Enquando existirem registros
'        Do While lErro = AD_SQL_SUCESSO
'
'            'Abertura de transação
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 185504
'
'            Set objTitRec = New ClassTituloReceber
'
'            objTitRec.lNumIntDoc = lNumIntTitulo
'
'            'Le o título
'            lErro = CF("TituloReceber_Le", objTitRec)
'            If lErro <> SUCESSO And lErro <> 26061 Then gError 185505
'            If lErro <> SUCESSO Then
'               'Lê o título baixado
'               lErro = CF("TituloReceberBaixado_Le", objTitRec)
'               If lErro <> SUCESSO And lErro <> 56570 Then gError 185506
'            End If
'            If lErro = SUCESSO Then
'
'                Set objTitRecTRV = New ClassTitulosRecTRV
'
'                objTitRecTRV.lNumIntDocTitRec = objTitRec.lNumIntDoc
'
'                lErro = Coinfo_GerarComissaoCartao(objImpCoinfo, objTitRec, objTitRecTRV)
'                If lErro <> SUCESSO Then gError 185507
'
'                lErro = CF("TitulosRecTRV_Grava", objTitRecTRV)
'                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 182542
'
'            End If
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 185508
'
'            lErro = Rotina_Atualizar_AtualizaTelaBatch
'            If lErro <> SUCESSO Then gError 185509
'
'            lErro = Comando_BuscarProximo(alComandoExt(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185510
'
'        Loop
'
'        lErro = Comando_ExecutarPos(alComandoExt(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComandoExt(0), gdtDataHoje)
'        If lErro <> SUCESSO Then gError 185511
'
'        lErro = Comando_BuscarProximo(alComandoExt(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185512
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacaoExt)
'    If lErro <> AD_SQL_SUCESSO Then gError 185513
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Comissões (Nfs a Pagar) foram atualizadas com sucesso. " & vbNewLine
'
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Coinfo_AtualizarComissao = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarComissao:
'
'    Coinfo_AtualizarComissao = gErr
'
'
'    Select Case gErr
'
'        Case 185496, 185497
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 185498, 185505, 185506, 185507, 185509
'
'        Case 185499, 185504
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 185500, 185501, 185512
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 185502, 185503, 185510
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOCOMISSAO", gErr)
'
'        Case 185508, 185513
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 185511
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185514)
'
'    End Select
'
'    Call Transacao_RollbackExt(lTransacaoExt)
'    Call Transacao_Rollback
'
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Public Function Coinfo_Associa_Cliente_Alm(ByVal objCliente As ClassCliente, iAlmoxarifado As Integer) As Long
'Associa um cliente a um almoxarifado
'Tem que estar dentro de transação

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEndereco As New ClassEndereco
Dim lNumAuto As Long, iPOS As Integer, sNomeRed As String

Const NUM_ALM_RESERVADOS = 20

On Error GoTo Erro_Coinfo_Associa_Cliente_Alm

    'Abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 185773
    Next
    
    'Le a associação entre cliente x almoxarifado
    lErro = Comando_Executar(alComando(0), "SELECT Almoxarifado FROM CoinfoClienteAlm WHERE Cliente = ?", iAlmoxarifado, objCliente.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 185774
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185775
    
    'Se não existir
    If lErro = AD_SQL_SEM_DADOS Then
    
        'faz a leitura do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 185776
        
'        If objCliente.iTipo <> 1 Then

            objEndereco.lCodigo = objCliente.lEndereco
        
            'Lê o endereço à partir do Código
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 185777
            
            iAlmoxarifado = 0
            
            Do
    
                'seleciona o codigo no bd e verifica se já existe
                lErro = CF("Config_ObterAutomatico_EmTrans", "CPRConfig", "NUM_PROX_ALMOXARIFADO", "Almoxarifado", "Codigo", lNumAuto)
                If lErro <> SUCESSO And lErro <> 25191 Then gError 185778
                
                iAlmoxarifado = CInt(lNumAuto)
            
            Loop Until iAlmoxarifado > NUM_ALM_RESERVADOS
            
            'If iAlmoxarifado <= NUM_ALM_RESERVADOS Then gError 185779
        
            objAlmoxarifado.iFilialEmpresa = giFilialEmpresa
            objAlmoxarifado.iCodigo = iAlmoxarifado
            objAlmoxarifado.sDescricao = left(objCliente.sRazaoSocial, STRING_ALMOXARIFADO_DESCRICAO)
            
            'If objCliente.lCodigo >= 1000000 Or Len(objCliente.sNomeReduzido) > STRING_ALMOXARIFADO_NOMEREDUZIDO Then
                          
'            sNomeRed = objCliente.sNomeReduzido
'            If Len(objCliente.sNomeReduzido) > STRING_ALMOXARIFADO_NOMEREDUZIDO Then
                iPOS = InStr(1, objCliente.sNomeReduzido, "(")
                'Se não tem o código do cliente ou ele está além da capacidade
                If iPOS = 0 Or iPOS > STRING_ALMOXARIFADO_NOMEREDUZIDO - 9 Then
                    sNomeRed = left(objCliente.sNomeReduzido, STRING_ALMOXARIFADO_NOMEREDUZIDO - 9) & "(" & objCliente.lCodigo & ")"
                Else
                    sNomeRed = left(objCliente.sNomeReduzido, iPOS - 1) & "(" & objCliente.lCodigo & ")"
                End If
'            End If
            objAlmoxarifado.sNomeReduzido = sNomeRed
            
            'Else
            '    objAlmoxarifado.sNomeReduzido = left(objCliente.sNomeReduzido, STRING_ALMOXARIFADO_NOMEREDUZIDO)
            'End If
            'MsgBox (objAlmoxarifado.sNomeReduzido)
            lErro = Comando_Executar(alComando(1), "INSERT INTO CoinfoClienteAlm (Cliente, Almoxarifado) VALUES (?,?)", objCliente.lCodigo, objAlmoxarifado.iCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 185781
        
            lErro = CF("Almoxarifado_Grava_EmTrans", objAlmoxarifado, objEndereco)
            If lErro <> SUCESSO Then
                gError 185782
            End If
            
'        End If
        
    End If
    
    'Fechamento de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_Associa_Cliente_Alm = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Associa_Cliente_Alm:

    Coinfo_Associa_Cliente_Alm = gErr
     
    Select Case gErr
    
        Case 185773
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 185774, 185775
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOCLIENTEALM", gErr)
            
        Case 185776, 185777, 185778, 185782
            
        Case 185779
            Call Rotina_Erro(vbOKOnly, "ERRO_COINFOCLIENTEALM_CODMENOR", gErr, NUM_ALM_RESERVADOS)

        Case 185781
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_COINFOCLIENTEALM", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185784)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Public Function Coinfo_Acerta_Cliente_Alm_Retroativo() As Long
'Associa um cliente a um almoxarifado

Dim lErro As Long
Dim lTransacao As Long

Dim obClienteFaixa As New ClassClienteFaixa
Dim colClientes As New Collection
Dim colClientesEst As New Collection
Dim objCliente As ClassCliente
Dim iAlmoxarifado As Integer

On Error GoTo Erro_Coinfo_Acerta_Cliente_Alm_Retroativo

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 185780
        
    lErro = CF("Cliente_Le_Faixa", obClienteFaixa, colClientes, colClientesEst)
    If lErro <> SUCESSO Then gError 185785
    
    For Each objCliente In colClientes
    
        lErro = Coinfo_Associa_Cliente_Alm(objCliente, iAlmoxarifado)
        If lErro <> SUCESSO Then gError 185786
    
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 185783
    
    Coinfo_Acerta_Cliente_Alm_Retroativo = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Acerta_Cliente_Alm_Retroativo:

    Coinfo_Acerta_Cliente_Alm_Retroativo = gErr
     
    Select Case gErr

        Case 185780
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 185783
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 185785, 185786
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185787)
     
    End Select

    Call Transacao_Rollback

    Exit Function

End Function


Private Function Coinfo_AtualizarVouchers(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza os movimentos de estoque

Dim lErro As Long
Dim lTransacaoExt As Long, lTransacao As Long
Dim alComandoExt(0 To 2) As Long, alComando(0 To 3) As Long, alCmdForaTrans(1 To 1) As Long
Dim iIndice As Integer, lNumIntArq As Long, lNumRegArqTexto As Long, dtData As Date
Dim lcodemp As Long, lNumVou As Long, dValor As Double, sSerie As String
Dim dtDataAux As Date, lcodempAux As Long, dvalorAux As Double, lNumIntDoc As Long
Dim sGrupo As String, iNumPassageiros As Integer, bIgnora As Boolean
Dim bMovAlteradaData As Boolean, lCodigoMovEst As Long
Dim objMovEstoque As New ClassMovEstoque, iFatorMultiplicacao As Integer
Dim sCodPro As String, objKit As ClassKit, objProdutoKit As New ClassProdutoKit
Dim objItemMovEstoque As New ClassItemMovEstoque, colRatreamentoMovto As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado, objCliente As ClassCliente
Dim iFilialCorporator As Integer, iFilialCoinfo As Integer
Dim lCodigo As Long, objImpCoInfoLog As ClassImpCoinfoLogAtu
Dim colImpCoInfoLog As Collection, sCritica As String
Dim objCtb As ClassImpCoinfoFatCRCtb, iComErro As Integer
Dim objContabAutomatica As ClassContabAutomatica, objContabil As ClassContabil
Dim bErro As Boolean, iAlmoxarifado As Integer, objEstoqueProduto As ClassEstoqueProduto
Dim colRastreamento As Collection, sTipVou As String
Dim iKit As Integer, iKitAnt As Integer, sCodProAnt As String
Dim bExcluiMov As Boolean, objProduto As ClassProduto, objProdutoCat As ClassProdutoCategoria
Dim bAchou As Boolean

Const TIPO_VOUCHER_INDIVIDUAL = "I"
Const TIPO_VOUCHER_FAMILIAR = "F"
Const TIPO_VOUCHER_GRUPO = "G"
Const CODIGO_IMPRESSO = "IMP       "

On Error GoTo Erro_Coinfo_AtualizarVouchers

    'Abre o arquivo
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 185788
    Next
    
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        alCmdForaTrans(iIndice) = Comando_AbrirExt(glConexao3)
        If alCmdForaTrans(iIndice) = 0 Then gError 185788
    Next
        
    'Obtém o total de registros que vão ser atualizados
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpitemfatu2")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao1)
    If lTransacaoExt = 0 Then gError 185791
    
    'Lê os arquivos que não foram atualizados
    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado2 = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fpitemfatu", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 185792
    
    'Busca o primeiro
    lErro = Comando_BuscarProximo(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185793
    
    'Enquanto existirem arquivos
    Do While lErro = AD_SQL_SUCESSO
    
        sSerie = String(255, 0)
        sCodPro = String(255, 0)
        sGrupo = String(255, 0)
        sTipVou = String(255, 0)
        
        'Le os vouchers que ainda não foram atualizados
        lErro = Comando_Executar(alCmdForaTrans(1), "SELECT NumIntDoc, NumRegArqTexto, Data, NumVou, CodEmp, Valor, Serie, CodPro, Grupo, QtdPax, TipVou, Kit FROM CoinfoItemFatura WHERE CodEst = ? AND NumIntArq = ? AND DataAtualizacaoMovEst = ? AND Data > ? AND Grupo <> '' AND TipoDoc = 'VOU' AND DiasAntc = 0 ORDER BY NumRegArqTexto", lNumIntDoc, lNumRegArqTexto, dtData, lNumVou, lcodemp, dValor, sSerie, sCodPro, sGrupo, iNumPassageiros, sTipVou, iKit, objImpCoinfo.sSiglaEmpresa, lNumIntArq, DATA_NULA, #11/30/2007#)
        If lErro <> AD_SQL_SUCESSO Then gError 185794
        
        'Busca o primeiro
        lErro = Comando_BuscarProximo(alCmdForaTrans(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185795
        
        'Enquando existirem vouchers não atualizados
        Do While lErro = AD_SQL_SUCESSO
        
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 185800
        
            'Abre o arquivo
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 185789
            Next
        
            Call TRVConverte_Prod(sCodPro)
        
            bIgnora = False
            
            bMovAlteradaData = False
            bExcluiMov = False
            
            bErro = False
            
            'Inicializa o Log
            Set colImpCoInfoLog = New Collection
            Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
            
            objImpCoInfoLog.iTabela = TABELA_COINFOITEMFATURA
            objImpCoInfoLog.lNumIntArq = lNumIntArq
            objImpCoInfoLog.lNumRegArqTexto = lNumRegArqTexto
            
            sCodProAnt = String(255, 0)

            'Le o vouchers para verificar se a atualização não ocorreu em outro registro
            lErro = Comando_Executar(alComando(0), "SELECT Data, CodEmp, Valor, CodigoMovEst, Kit,CodPro FROM CoinfoItemFatura WHERE CodEst = ? AND NumVou = ? AND Serie = ? AND DataAtualizacaoMovEst <> ? AND CodigoMovEst <> 0 ORDER BY NumIntArq DESC, NumRegArqTexto DESC", dtDataAux, lcodempAux, dvalorAux, lCodigoMovEst, iKitAnt, sCodProAnt, objImpCoinfo.sSiglaEmpresa, lNumVou, sSerie, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 185796
            
            'Busca o primeiro
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185797
            
            'Se o registro já foi atualizado
            If lErro = AD_SQL_SUCESSO Then
            
                Call TRVConverte_Prod(sCodProAnt)
            
                'Se mudou a quantidades de impressos que devem sair, a data da saída ou o produto vendido
                If dtData <> dtDataAux Or UCase(sCodProAnt) <> UCase(sCodPro) Or iKit <> iKitAnt Then
                    If lCodigoMovEst <> 0 Then
                        If UCase(sCodProAnt) <> UCase(sCodPro) Or iKit <> iKitAnt Then
                            bExcluiMov = True
                        Else
                            bMovAlteradaData = True
                        End If
                    End If
                Else
                    bIgnora = True
                End If
                
            End If
            
            If iKit = 0 Then bIgnora = True
            
            'Se é necessário tratar registro (não é uma repetição sem alteração de data\kit\Produto)
            If Not bIgnora Then
                        
                Set objCliente = New ClassCliente
                Set objKit = New ClassKit
                Set objAlmoxarifado = New ClassAlmoxarifado
                Set objMovEstoque = New ClassMovEstoque

                'Le a filial associada ao cliente
                lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", lcodemp, iFilialCorporator, iFilialCoinfo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                'Le o cliente
                objCliente.lCodigo = lcodemp
            
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                
                If lErro <> SUCESSO Then
                    bErro = True
                    Call CoInfo_Adiciona_Log("O Cliente " & CStr(lcodemp) & " não está cadastrado.", objImpCoInfoLog, colImpCoInfoLog)
                End If
                
                'Se a associação já existir somente lê o almoxarifado, senão cria
                lErro = Coinfo_Associa_Cliente_Alm(objCliente, iAlmoxarifado)
                If lErro <> SUCESSO Then
                    bErro = True
                    Call CoInfo_Adiciona_Log("Não foi possível ler o almoxarifado associado ao cliente", objImpCoInfoLog, colImpCoInfoLog)
                End If
                
                If bMovAlteradaData Or bExcluiMov Then
                    
                    objMovEstoque.lCodigo = lCodigoMovEst
                    objMovEstoque.iFilialEmpresa = giFilialEmpresa
                    
                    lErro = CF("MovEstoque_Le", objMovEstoque)
                    If lErro <> SUCESSO And lErro <> 30128 Then gError ERRO_SEM_MENSAGEM
                
                    If lErro <> SUCESSO Then
                        bErro = True
                        Call CoInfo_Adiciona_Log("O movimento de estoque original do voucher com código " & CStr(lCodigoMovEst) & " não está cadastrado.", objImpCoInfoLog, colImpCoInfoLog)
                    Else
                                    
                        'Lê os ítens do Movimento de Estoque
                        lErro = CF("MovEstoqueInternoItens_Le", objMovEstoque)
                        If lErro <> SUCESSO And lErro <> 30565 Then gError ERRO_SEM_MENSAGEM
                    
                        If lErro <> SUCESSO Then
                            bErro = True
                            Call CoInfo_Adiciona_Log("O movimento de estoque original do voucher com código " & CStr(lCodigoMovEst) & " não tem itens cadastrados.", objImpCoInfoLog, colImpCoInfoLog)
                        End If
                        
                    End If
                    
                End If
                
                If bExcluiMov Then
                
                    Set objMovEstoque = New ClassMovEstoque
               
                    objMovEstoque.lCodigo = lCodigoMovEst
                    objMovEstoque.iFilialEmpresa = giFilialEmpresa
                    
                    'trata a exclusão do movimento de estoque
                    lErro = CF("MovEstoque_Exclui_Trans", objMovEstoque, Nothing)
                    If lErro <> SUCESSO Then
                        bErro = True
                        Call CoInfo_Adiciona_Log("Não conseguiu excluir o movimento de estoque com código " & CStr(lCodigoMovEst) & " não tem itens cadastrados.", objImpCoInfoLog, colImpCoInfoLog)
                    Else
                    
                        'EXCLUI CONTABILIZAÇÃO
                    
                    End If
                    
                    Set objMovEstoque = New ClassMovEstoque
                    
                End If
                
                If bMovAlteradaData Then
                    'Altera a data do movimento
                    objMovEstoque.dtData = dtData
                Else
                
                    Select Case sGrupo
                    
                        Case TIPO_VOUCHER_INDIVIDUAL
                            iFatorMultiplicacao = 1
                        
                        Case TIPO_VOUCHER_FAMILIAR
                            iFatorMultiplicacao = iNumPassageiros
                    
                        Case TIPO_VOUCHER_GRUPO
                            iFatorMultiplicacao = Round((iNumPassageiros / 10) + 0.4999)
                            
                        Case Else
                            bErro = True
                            Call CoInfo_Adiciona_Log("O tipo " & sGrupo & " de voucher desconhecido.", objImpCoInfoLog, colImpCoInfoLog)
                    
                    End Select
                    
                    If Not bErro Then
                    
                        lErro = CF("MovEstoque_Automatico_EmTransacao", giFilialEmpresa, lCodigo)
                        If lErro <> SUCESSO Then
                            bErro = True
                            Call CoInfo_Adiciona_Log("Não foi possível gerar o número do próximo movimento de estoque.", objImpCoInfoLog, colImpCoInfoLog)
                        End If
                        
                    End If
                    
                    objMovEstoque.iFilialEmpresa = giFilialEmpresa
                    objMovEstoque.dtData = dtData
                    objMovEstoque.lCodigo = lCodigo
                    
                    'CRIAR CÓDIGO PARA BUSCAR VERSÃO DE UMA CATEGORIA
                    Set objProduto = New ClassProduto
                    
                    objProduto.sCodigo = sCodPro & String(10 - Len(sCodPro), Asc(" "))
                    
                    'Carrega a colCategoriaItem do objProduto
                    lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
                    If lErro <> SUCESSO Then
                        bErro = True
                        Call CoInfo_Adiciona_Log("Não foi possível ler a categoria do produto " & objProduto.sCodigo & ".", objImpCoInfoLog, colImpCoInfoLog)
                    End If
                    
                    bAchou = False
                    For Each objProdutoCat In objProduto.colCategoriaItem
                        If objProdutoCat.sCategoria = CATEGORIA_PRODUTO_VERSAO_KIT Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        bErro = True
                        Call CoInfo_Adiciona_Log("Não informado o kit a ser emitido pelo produto " & objProduto.sCodigo & ".", objImpCoInfoLog, colImpCoInfoLog)
                    Else
                        objKit.sVersao = objProdutoCat.sItem
                    End If
                    
                    'Obtém listagem de produtos do kit
                    objKit.sProdutoRaiz = CODIGO_IMPRESSO
                    
                    'Le o Kit
                    lErro = CF("Kit_Le", objKit)
                    If lErro <> SUCESSO And lErro <> 21826 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro <> SUCESSO Then
                        bErro = True
                        Call CoInfo_Adiciona_Log("O kit para o produto " & objKit.sProdutoRaiz & " e versão " & objKit.sVersao & " não está cadastrado.", objImpCoInfoLog, colImpCoInfoLog)
                    End If
                    
                    'Le os insumos do kit
                    lErro = CF("Kit_Le_Componentes", objKit)
                    If lErro <> SUCESSO And lErro <> 21831 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro <> SUCESSO Then
                        bErro = True
                        Call CoInfo_Adiciona_Log("O kit para o produto " & objKit.sProdutoRaiz & " e versão " & objKit.sVersao & " não tem itens.", objImpCoInfoLog, colImpCoInfoLog)
                    End If
                    
                    For Each objProdutoKit In objKit.colComponentes
                    
                        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                    
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, MOV_EST_OUTRAS_REQ, 0, 0, objProdutoKit.sProduto, "", objProdutoKit.sUnidadeMed, objProdutoKit.dQuantidade * iKit, iAlmoxarifado, "", 0, "", 0, "", "", "", "", 0, colRatreamentoMovto, Nothing, DATA_NULA)
                            
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
                                    
                            'VERIFICA SE TEM ASSOCIACAO COM O ALMOXARIFADO, SENÃO CRIA
                            Set objEstoqueProduto = New ClassEstoqueProduto
                            Set colRastreamento = New Collection
                                    
                            'Preenche objEstoqueProduto
                            objEstoqueProduto.sProduto = objProdutoKit.sProduto
                            objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
                            objEstoqueProduto.dtDataInicial = dtData
                            objEstoqueProduto.iFilialEmpresa = FILIAL_MATRIZ
                        
                            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 21306 Then
                                bErro = True
                                Call CoInfo_Adiciona_Log("Não foi possível ler a Tabela EstoqueProduto.", objImpCoInfoLog, colImpCoInfoLog)
                            End If
                            'Não cadastrado
                            If lErro = 21306 Then
                                bErro = True
                                Call CoInfo_Adiciona_Log("Não existe estoque inicial cadastrado para o produto " & objEstoqueProduto.sProduto & " e o almoxarifado " & CStr(objEstoqueProduto.iAlmoxarifado) & ".", objImpCoInfoLog, colImpCoInfoLog)
                            
'                                lErro = CF("EstoqueInicial_Grava1", objEstoqueProduto, MARCADO, colRastreamento)
'                                If lErro <> SUCESSO Then
'                                    bErro = True
'                                    Call CoInfo_Adiciona_Log("Não foi possível associar o produto " & objProdutoKit.sProduto & " ao almoxarifado " & CStr(iAlmoxarifado) & ".", objImpCoInfoLog, colImpCoInfoLog)
'                                End If
                            End If
                        End If
                                
                    Next
                
                End If
                
                If Not bErro Then
                
                    'Se a data do movimento é posterior ao estoque inicial -> Grava o movimento
                    If objEstoqueProduto.dtDataInicial <= objMovEstoque.dtData Then
                    
                        objMovEstoque.sObservacao = "Emissão do voucher Tipo " & sTipVou & " Série " & sSerie & " Número " & CStr(lNumVou) & " - Qtd Impressos = " & CStr(iKit)
                
                        'Grava no BD (inclusive os dados contabeis) (contabilidade)
                        lErro = CF("MovEstoque_Interno_Grava0", objMovEstoque, Nothing)
                        If lErro <> SUCESSO Then
                            bErro = True
                            Call CoInfo_Adiciona_Log("O movimento de estoque para o voucher " & CStr(lNumVou) & " e série " & sSerie & " não pode ser gravado.", objImpCoInfoLog, colImpCoInfoLog)
                        Else
                        
                            'GRAVA CONTABILIZAÇÃO
                            
                            'NÃO TEM CONTABILIZAÇÃO, OS VOUCHERS JÁ SÃO CONTABILIZADOS EM OUTRO PONTO
                            
                        End If
                        
                    End If
                
                End If
                
                'Insere o Log de erros
                lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Atualiza o voucher para informar que ele foi importado com sucesso
                If iComErro = DESMARCADO Then
                
                    'Atualiza os dados no vouchers
                    lErro = Comando_Executar(alComando(1), "UPDATE CoinfoItemFatura SET DataAtualizacaoMovEst = ?, CodigoMovEst = ?, ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", Date, objMovEstoque.lCodigo, iComErro, lNumIntArq, lNumRegArqTexto)
                    If lErro <> SUCESSO Then gError 185806
                
                Else
                
                    bErro = True
                    
                    'Fecha os comandos
                    For iIndice = LBound(alComando) To UBound(alComando)
                        Call Comando_Fechar(alComando(iIndice))
                    Next
                
                    Call Transacao_Rollback

                    'Abertura de transação só para marcar que houve erro
                    lTransacao = Transacao_Abrir()
                    If lTransacao = 0 Then gError 185807
                    
                    'Abre o arquivo
                    For iIndice = LBound(alComando) To UBound(alComando)
                        alComando(iIndice) = Comando_Abrir()
                        If alComando(iIndice) = 0 Then gError 185789
                    Next
            
                    'Atualiza os dados no vouchers
                    lErro = Comando_Executar(alComando(2), "UPDATE CoinfoItemFatura SET ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iComErro, lNumIntArq, lNumRegArqTexto)
                    If lErro <> SUCESSO Then gError 185808
                
                End If
                
            Else
        
                'Atualiza os dados no vouchers
                lErro = Comando_Executar(alComando(3), "UPDATE CoinfoItemFatura SET ComErro = ?, DataAtualizacaoMovEst = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", 0, Date, lNumIntArq, lNumRegArqTexto)
                If lErro <> SUCESSO Then gError 185808
                            
            End If
            
            'Fecha os comandos
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
            
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 185809
            
            'Atualiza tela de acompanhamento
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            'Busca próximo vouchers
            lErro = Comando_BuscarProximo(alCmdForaTrans(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185811
        
        Loop
        
        If Not bErro Then
        
            'Marca o arquivo como atualizado
            lErro = Comando_ExecutarPos(alComandoExt(2), "UPDATE CoinfoArqsLidos SET DataAtualizado2 = ?", alComandoExt(0), gdtDataHoje)
            If lErro <> SUCESSO Then gError 185812
            
        End If
    
        'Busca o próximo arquivo
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185813
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 185814
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "A emissão de vouchers foi atualizada com sucesso (contabilização e movimentação de estoque)." & vbNewLine
    
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
        
    Coinfo_AtualizarVouchers = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarVouchers:

    Coinfo_AtualizarVouchers = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
          
        Case 185788, 185789
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 185791, 185800, 185807
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 185792, 185793, 185813
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 185794 To 185797, 185811
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMFATURA", gErr)
        
        Case 185806, 185808
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOITEMFATURA", gErr)
        
        Case 185809, 185814
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 185813
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185815)
     
    End Select
    
    'Desfaz a transação
    Call Transacao_RollbackExt(lTransacaoExt)
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Private Function Coinfo_AtualizarVouchers2(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'Atualiza a contabilização dos vouchers

Dim lErro As Long
Dim lTransacaoExt As Long, lTransacao As Long
Dim alComandoExt(0 To 2) As Long, alComando(0 To 5) As Long, alCmdForaTrans(1 To 1) As Long
Dim iIndice As Integer, lNumIntDoc As Long, lNumIntArq As Long
Dim lNumRegArqTexto As Long, sTipVou As String, sSerie As String
Dim lNumVou As Long, lClienteVou As Long, lClienteAnt As Long, lClienteCtb As Long
Dim iFilialCorporatorAnt As Integer, iFilialCorporatorCtb As Integer
Dim dtDataVou As Date, dtDataAnt As Date, dtDataCtb As Date, dValorVou As Double
Dim dValorAnt As Double, dValorCtb As Double, bIgnora As Boolean, dComissao As Double
Dim iCartaoCtb As Integer, iCartaoAnt As Integer, bDesfazContabAnterior As Boolean
Dim iFilialCoinfo As Integer, objImpCoInfoLog As ClassImpCoinfoLogAtu, colImpCoInfoLog As Collection
Dim sCritica As String, objCtb As New ClassImpCoinfoItemCRCtb, iComErro As Integer
Dim objContabAutomatica As ClassContabAutomatica, objContabil As ClassContabil
Dim bErro As Boolean, objImpVoucher As ClassImpCoinfoItemFat, bTemErro As Boolean
Dim iExcluido As Integer, bPulaArquivo As Boolean, bPulaRegistro As Boolean
Dim iForcaCTB As Integer, lNumFat As Long, bSabeArq As Boolean, bSabeVou As Boolean, sPonto As String

On Error GoTo Erro_Coinfo_AtualizarVouchers2

    'Abre o arquivo
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 185788
    Next
    
    For iIndice = LBound(alCmdForaTrans) To UBound(alCmdForaTrans)
        alCmdForaTrans(iIndice) = Comando_AbrirExt(glConexao3)
        If alCmdForaTrans(iIndice) = 0 Then gError 185788
    Next
    
    bSabeArq = False
    bSabeVou = False
    sPonto = ""
    
    giDesconsideraFechamentoPeriodo = MARCADO
    
    'Obtém o total de registros que vão ser atualizados
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpitemfatu3")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Abertura de transação
    lTransacaoExt = Transacao_AbrirExt(glConexao1)
    If lTransacaoExt = 0 Then gError 185791
    
    'Lê os arquivos que não foram atualizados
    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado3 = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fpitemfatu", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 185792
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185793
    
    'Enquanto existirem arquivos
    Do While lErro = AD_SQL_SUCESSO
        
        bSabeArq = True
    
        bTemErro = False
        bPulaArquivo = False
    
        sSerie = String(255, 0)
        sTipVou = String(255, 0)
        
        'Le os vouchers que ainda não foram atualizados
        lErro = Comando_Executar(alCmdForaTrans(1), "SELECT NumIntArq, NumRegArqTexto, TipVou, Serie, NumVou, Excluido, Valor, Data, CodEmpVou, Cartao FROM CoinfoItemFaturaCTB WHERE ((CodEst = ? AND NumIntArq = ?) OR ForcaCTB=1) AND DataAtualizacaoContab = ? AND Data > = ? ORDER BY NumIntArq, NumRegArqTexto", lNumIntArq, lNumRegArqTexto, sTipVou, sSerie, lNumVou, iExcluido, dValorCtb, dtDataCtb, lClienteCtb, iCartaoCtb, objImpCoinfo.sSiglaEmpresa, lNumIntArq, DATA_NULA, #1/1/2010#)
        If lErro <> AD_SQL_SUCESSO Then gError 185794
        
        'Busca o primeiro
        lErro = Comando_BuscarProximo(alCmdForaTrans(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185795
        
        'Enquando existirem vouchers não atualizados
        Do While lErro = AD_SQL_SUCESSO
        
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 185800
        
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 185789
            Next
    
            bSabeVou = True
        
            bIgnora = False
            bPulaRegistro = False
            lNumFat = 0
            
            'Le o voucher em TRVVouchers
            lErro = Comando_Executar(alComando(3), "SELECT Data, Valor, ValorComissao, ClienteVou, ForcaCTB, NumFatCoinfo FROM TRVVouchers WHERE NumVou = ? AND TipoDoc = 'VOU' AND Serie = ? AND TipVou = ? AND ExtraiInfoSigav = 0 ", dtDataVou, dValorVou, dComissao, lClienteVou, iForcaCTB, lNumFat, lNumVou, sSerie, sTipVou)
            If lErro <> AD_SQL_SUCESSO Then gError 185794
            
            'Busca o primeiro
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185795
            
            If lErro <> AD_SQL_SUCESSO Then bPulaRegistro = True 'Não tem as informações necessárias importadas ainda
            'If lNumFat <> 0 Then bIgnora = True
            
            'Pula por ainda não ter informações suficientes
            If Not bPulaRegistro Then
            
                Set objImpVoucher = New ClassImpCoinfoItemFat
                
                objImpVoucher.dtData = dtDataCtb
                objImpVoucher.lNumVou = lNumVou
                objImpVoucher.dValor = dValorVou
                objImpVoucher.sSerie = sSerie
                objImpVoucher.iCartao = iCartaoCtb
                objImpVoucher.lcodemp = lClienteCtb
                If objImpVoucher.iCartao = MARCADO Then
                    objImpVoucher.dComissao = 0
                Else
                    objImpVoucher.dComissao = dComissao
                End If
                
                sPonto = "1"
                
                'Le a filial associada ao cliente
                lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", lClienteCtb, iFilialCorporatorCtb, iFilialCoinfo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                sPonto = "2"
                
                objImpVoucher.iFilialEmpresa = iFilialCorporatorCtb
                
                bDesfazContabAnterior = False
                
                bErro = False
                
                'Inicializa o Log
                Set colImpCoInfoLog = New Collection
                Set objImpCoInfoLog = New ClassImpCoinfoLogAtu
                
                objImpCoInfoLog.iTabela = TABELA_COINFOITEMFATURA
                objImpCoInfoLog.lNumIntArq = lNumIntArq
                objImpCoInfoLog.lNumRegArqTexto = lNumRegArqTexto
    
                'Le o vouchers para verificar se a atualização não ocorreu em outro registro
                lErro = Comando_Executar(alComando(4), "SELECT NumIntDoc, Data, CodEmp, Valor, Cartao FROM CoinfoItemFatura WHERE CodEst = ? AND NumVou = ? AND Serie = ? AND TipVou = ? AND NumIntDoc <> 0 AND TipoDoc = 'VOU' AND Excluido = 0 ORDER BY NumIntArq DESC, NumRegArqTexto DESC", lNumIntDoc, dtDataAnt, lClienteAnt, dValorAnt, iCartaoAnt, objImpCoinfo.sSiglaEmpresa, lNumVou, sSerie, sTipVou)
                If lErro <> AD_SQL_SUCESSO Then gError 185796
                
                'Busca o primeiro
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185797
                
                sPonto = "3"
                
                'Se o registro já foi atualizado
                If lErro = AD_SQL_SUCESSO Then
                
                    objImpVoucher.lNumIntDoc = lNumIntDoc
                
                    'Le a filial associada ao cliente
                    lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", lClienteAnt, iFilialCorporatorAnt, iFilialCoinfo)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    'Se ocorreu uma mudança de valor ou data ou filial empresa
                    If Abs(dValorCtb - dValorAnt) > DELTA_VALORMONETARIO Or dtDataCtb <> dtDataAnt Or iExcluido = MARCADO Or iFilialCorporatorCtb <> iFilialCorporatorAnt Or iCartaoCtb <> iCartaoAnt Or iForcaCTB = MARCADO Then
                        bDesfazContabAnterior = True
                    Else
                        bIgnora = True
                    End If
                Else
                    If iExcluido = MARCADO Then bIgnora = True
                End If
                
                sPonto = "4"
                                                            
                'Se é necessário tratar registro (não é uma repetição sem alteração de data\valor\filialempresa)
                If Not bIgnora Then
                
                    sPonto = "4.1"
                    
                    If objImpVoucher.lNumIntDoc = 0 Then
                
                        lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ITEM_FATURA", lNumIntDoc, 1, EMPRESA_TODA)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        objImpVoucher.lNumIntDoc = lNumIntDoc
                        
                    End If
                    
                    sPonto = "4.2"
                    
                    'Se é para desfazer a contabilização
                    If bDesfazContabAnterior And Not bErro Then
                    
                        Set objContabil = New ClassContabil
                        Call objContabil.Contabil_Inicializa_Contabilidade4(objCtb.iTransacao, MODULO_CONTABILIDADE)
                         
                        If Not (objContabil Is Nothing) Then
                                 
                            'Exclui os lançamentos contábeis (se houverem) associados ao titulo em questao
                            lErro = objContabil.Contabil_Exclui(objImpVoucher.lNumIntDoc, CONTABIL_EXCLUI_EFETIVO, , , , objImpVoucher.iFilialEmpresa)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                            Set objContabil = Nothing
                        End If
                    
                    End If
                    
                    sPonto = "4.3"
                    
                    If Not bErro And iExcluido = DESMARCADO Then
    
                        Set objContabAutomatica = New ClassContabAutomatica
                        Set objCtb = New ClassImpCoinfoItemCRCtb
                        
                        'gravar a contabilizacao
                        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objImpVoucher.dtData, objImpVoucher.dtData)
                        If lErro = SUCESSO Then
                        
                            lErro = objContabAutomatica.GeraContabilizacao(objImpVoucher)
                            If lErro = SUCESSO Then
                                lErro = objContabAutomatica.Finaliza_Contab()
                            Else
                                bErro = True
                                Call CoInfo_Adiciona_Log("Erro na contabilização. GeraContabilizacao", objImpCoInfoLog, colImpCoInfoLog)
                            End If
                            
                        Else
                            bErro = True
                            Call CoInfo_Adiciona_Log("Erro na contabilização. Inicializa_Contab", objImpCoInfoLog, colImpCoInfoLog)
                            
                        End If
            
                        Set objContabAutomatica = Nothing
                        Call objCtb.Encerrando
                        Set objCtb = Nothing
                        
                    End If
                    sPonto = "4.4"
                Else
                    objImpVoucher.lNumIntDoc = 0
                End If
                
                sPonto = "5"
                    
                'Insere o Log de erros
                lErro = CoInfoLogAtualizacao_Insere(colImpCoInfoLog, iComErro, objImpCoInfoLog)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Atualiza o voucher para informar que ele foi importado com sucesso
                If iComErro = DESMARCADO Then
                
                    'Atualiza os dados no vouchers
                    lErro = Comando_Executar(alComando(5), "UPDATE TRVVouchers SET ForcaCTB = ?, NumIntDoc = ? WHERE NumVou = ? AND TipoDoc = 'VOU' AND Serie = ? AND TipVou = ?", 0, objImpVoucher.lNumIntDoc, lNumVou, sSerie, sTipVou)
                    If lErro <> SUCESSO Then gError 185806
                    
                    'Atualiza os dados no vouchers
                    lErro = Comando_Executar(alComando(1), "UPDATE CoinfoItemFatura SET DataAtualizacaoContab = ?, NumIntDoc = ?, ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", gdtDataHoje, objImpVoucher.lNumIntDoc, iComErro, lNumIntArq, lNumRegArqTexto)
                    If lErro <> SUCESSO Then gError 185806
                
                Else
                
                    bErro = True
                    bTemErro = True
                    
                    'Fecha os comandos
                    For iIndice = LBound(alComando) To UBound(alComando)
                        Call Comando_Fechar(alComando(iIndice))
                    Next
                    
                    Call Transacao_Rollback
    
                    'Abertura de transação só para marcar que houve erro
                    lTransacao = Transacao_Abrir()
                    If lTransacao = 0 Then gError 185807
                    
                    'Abre o arquivo
                    For iIndice = LBound(alComando) To UBound(alComando)
                        alComando(iIndice) = Comando_Abrir()
                        If alComando(iIndice) = 0 Then gError 185789
                    Next
            
                    'Atualiza os dados no vouchers
                    lErro = Comando_Executar(alComando(2), "UPDATE CoinfoItemFatura SET ComErro = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", iComErro, lNumIntArq, lNumRegArqTexto)
                    If lErro <> SUCESSO Then gError 185808
                
                End If
                
                sPonto = "6"
                                   
            End If
            
            'Fecha os comandos
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
                
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 185809
            
            'Atualiza tela de acompanhamento
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If bPulaRegistro Then bPulaArquivo = True
            
            'Busca próximo vouchers
            lErro = Comando_BuscarProximo(alCmdForaTrans(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185811
        
        Loop
        
        If Not bTemErro And Not bPulaArquivo Then
        
            'Marca o arquivo como atualizado
            lErro = Comando_ExecutarPos(alComandoExt(2), "UPDATE CoinfoArqsLidos SET DataAtualizado3 = ?", alComandoExt(0), gdtDataHoje)
            If lErro <> SUCESSO Then gError 185812
            
        End If
    
        'Busca o próximo arquivo
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185813
        
    Loop
    
    bSabeArq = False
    
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_CommitExt(lTransacaoExt)
    If lErro <> AD_SQL_SUCESSO Then gError 185814
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "A contabilização dos vouchers foi feita com sucesso." & vbNewLine
                
    giDesconsideraFechamentoPeriodo = DESMARCADO
        
    Coinfo_AtualizarVouchers2 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarVouchers2:

    giDesconsideraFechamentoPeriodo = DESMARCADO

    Coinfo_AtualizarVouchers2 = gErr
    
    Select Case gErr
          
        Case 185788, 185789
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case 185791, 185800, 185807
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 185792, 185793, 185813
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 185794 To 185797, 185811
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMFATURA", gErr)
        
        Case 185806, 185808
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOITEMFATURA", gErr)
        
        Case 185809, 185814
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 185813
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185815)
     
    End Select
    
    If bSabeArq Then
    
        If bSabeVou Then
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_CONTAB_VOU1", gErr, lNumIntArq, sTipVou, sSerie, lNumVou, sPonto)
        Else
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_CONTAB_VOU2", gErr, lNumIntArq, sPonto)
        End If
    
    End If
    
    'Desfaz a transação
    Call Transacao_RollbackExt(lTransacaoExt)
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function

End Function

Private Function Coinfo_DesfazCTBVou(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza os fornecedores

Dim lErro As Long
Dim lTransacaoExt As Long
Dim lTransacao As Long

Dim alComandoExt(0 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iFilialEmpresa As Integer
Dim objCtb As New ClassImpCoinfoItemCRCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil

On Error GoTo Erro_Coinfo_DesfazCTBVou
    
    'Abre o arquivo
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 185788
    Next
    
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 185791

    'Lê os arquivos que não foram atualizados
    lErro = Comando_Executar(alComandoExt(0), "SELECT NumIntDoc, FilialEmpresa FROM Lancamentos WHERE Transacao = ? ORDER BY Data, NumIntDoc ", lNumIntDoc, iFilialEmpresa, objCtb.iTransacao)
    If lErro <> AD_SQL_SUCESSO Then gError 185792
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185793
    
    'Enquanto existirem registros
    Do While lErro = AD_SQL_SUCESSO
        
        'Abertura de transação só para marcar que houve erro
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 185807
                
        Set objContabil = New ClassContabil
        Call objContabil.Contabil_Inicializa_Contabilidade4(objCtb.iTransacao, MODULO_CONTABILIDADE)
         
        If Not (objContabil Is Nothing) Then
                 
            'Exclui os lançamentos contábeis (se houverem) associados ao doc em questao
            lErro = objContabil.Contabil_Exclui(lNumIntDoc, CONTABIL_EXCLUI_EFETIVO, , , , iFilialEmpresa)
            If lErro <> SUCESSO Then gError 83949
    
            Set objContabil = Nothing
        End If
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 185809
            
        'Busca o próximo
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185793
            
    Loop
                
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacaoExt)
'    If lErro <> AD_SQL_SUCESSO Then gError 185814
        
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
                
    Coinfo_DesfazCTBVou = SUCESSO
     
    Exit Function
    
Erro_Coinfo_DesfazCTBVou:

    Coinfo_DesfazCTBVou = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185815)
     
    End Select
    
    'Desfaz a transação
    Call Transacao_RollbackExt(lTransacaoExt)
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
        
    Exit Function

End Function

Public Function Coinfo_Produto_Associa_Alm() As Long
'Associa um produto a um almoxarifado

Dim lErro As Long
Dim lTransacao As Long
Dim colProdutos As New Collection
Dim colAlmoxarifados As New Collection
Dim objEstoqueProduto As ClassEstoqueProduto
Dim objEstoqueProdutoBD As ClassEstoqueProduto
Dim objProd As ClassProduto
Dim objAlm As ClassAlmoxarifado
Dim colRastreamento As New Collection
Dim iAlmoxarifadoPadrao As Integer

On Error GoTo Erro_Coinfo_Produto_Associa_Alm

'    'Abertura de transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 194834
'
'    lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
'    If lErro <> SUCESSO Then gError 194850
    
    lErro = CF("Produto_Le_Faixa", "", "", colProdutos)
    If lErro <> SUCESSO Then gError 194835
    
    lErro = CF("Almoxarifado_Le_Faixa", 0, 0, colAlmoxarifados)
    If lErro <> SUCESSO Then gError 194836
    
    For Each objProd In colProdutos
    
        If objProd.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And objProd.sCodigo <> "IMP       " And objProd.iTipo = 3 Then
        
            For Each objAlm In colAlmoxarifados
            
                Set objEstoqueProduto = New ClassEstoqueProduto
                Set objEstoqueProdutoBD = New ClassEstoqueProduto
    
                objEstoqueProdutoBD.iFilialEmpresa = giFilialEmpresa
                objEstoqueProdutoBD.iAlmoxarifado = objAlm.iCodigo
                objEstoqueProdutoBD.sProduto = objProd.sCodigo
                
                lErro = CF("EstoqueProduto_Le", objEstoqueProdutoBD)
                If lErro <> SUCESSO And lErro <> 21306 Then gError 196235
    
                If lErro <> SUCESSO Then
                
                    objEstoqueProduto.dtDataInicial = gdtDataHoje
                    objEstoqueProduto.iFilialEmpresa = giFilialEmpresa
                    objEstoqueProduto.iAlmoxarifado = objAlm.iCodigo
                    objEstoqueProduto.sProduto = objProd.sCodigo
                    objEstoqueProduto.dQuantidadeInicial = 0 '50
                    
                    iAlmoxarifadoPadrao = MARCADO
                    
                    'Chama EstoqueInicial_Grava
                    'lErro = CF("EstoqueInicial_Grava_Trans", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
                    lErro = CF("EstoqueInicial_Grava", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
                    If lErro <> SUCESSO Then gError 194837
                    
                End If
                
            Next
            
        End If
        
    Next
    
'    lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
'    If lErro <> SUCESSO Then gError 194851
'
'    'Confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 194838


    Call Rotina_Aviso(vbOKOnly, "Rotina finalizada com sucesso.")
    
    Coinfo_Produto_Associa_Alm = SUCESSO
     
    Exit Function
    
Erro_Coinfo_Produto_Associa_Alm:

    Coinfo_Produto_Associa_Alm = gErr
     
    Select Case gErr

        Case 194834
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 194835 To 194837, 194850, 194851, 196235
            
        Case 194838
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194839)
     
    End Select
    
'    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
'
'    Call Transacao_Rollback

    Exit Function

End Function

'Private Function Coinfo_AtualizarVouchers3(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''Atualiza os Títulos de Cartão de Crédito a partir dos vouchers pagos com cartão
'
'Dim lErro As Long
'Dim lTransacao As Long
'Dim alComando(0 To 12) As Long
'Dim iIndice As Integer
'Dim sNumAuto As String
'Dim sNumAutoAnt As String
'Dim sCia As String
'Dim sCiaAnt As String
'Dim sNumCCred As String
'Dim sNumCCredAnt As String
'Dim dtData As Date
'Dim dtDataAnt As Date
'Dim lNumVou As Long
'Dim sSerie As String
'Dim sTipo As String
'Dim objVoucher As ClassTRVVouchers
'Dim colVoucher As New Collection
'Dim colcolVoucher As New Collection
'Dim lCountRegs As Long
'Dim objTitRec As ClassTituloReceber
'Dim dValor As Double
'Dim objImpCoinfoCartoes As ClassImpCoinfoCartoes
'Dim iFilialCorporator As Integer
'Dim iFilialCoinfo As Integer
'Dim sConteudo As String
'Dim iComErro As Integer
'Dim objCtb As ClassImpCoinfoFatCRCtb
'Dim objContabAutomatica As ClassContabAutomatica
'Dim objTRVConfig As ClassTRVConfig
'Dim lProxTitRec As Long
'Dim colConfig As Collection
'Dim objTitRecTRV As ClassTitulosRecTRV
'Dim sModeloFat As String
'Dim sNomeDiretorio As String
'Dim objTRVTitulosExp As ClassTRVTitulosExp
'Dim objVoucherInfo As ClassTRVVoucherInfo
'Dim dtDataEmissaoAux As Date
'Dim objVouExp As ClassTRVVouExp
'Dim objTRVTitulos As ClassTRVTitulos
'
'On Error GoTo Erro_Coinfo_AtualizarVouchers3
'
'    'Abre o comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 192704
'    Next
'
'    sCia = String(STRING_TRV_VOU_CIACART, 0)
'    sNumCCred = String(STRING_TRV_VOU_NUMCCRED, 0)
'    sSerie = String(STRING_TRV_OCR_SERIE, 0)
'    sTipo = String(STRING_TRV_OCR_TIPOVOU, 0)
'    sNumAuto = String(STRING_TRV_VOU_NUMAUTO, 0)
'
'    'Lê vouchers pagos com cartão de crédito não atualizados emitidos a pelo menos 2 dias
'    lErro = Comando_Executar(alComando(0), "SELECT CiaCart, NumCCred, Data, NumAuto, NumVou, Serie, TipVou FROM TRVVouchers WHERE Cartao = 1 AND NumFatCoinfo = 0 AND CiaCart <> '' AND NumAuto <> 0 AND TipoDoc = ? AND Data < ? AND Status <> 7 ORDER BY CiaCart, NumCCred, Data, NumAuto", sCia, sNumCCred, dtData, sNumAuto, lNumVou, sSerie, sTipo, TRV_TIPODOC_VOU_TEXTO, DateAdd("d", -2, Date))
'    If lErro <> AD_SQL_SUCESSO Then gError 192705
'
'    'Busca o primeiro
'    lErro = Comando_BuscarPrimeiro(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192706
'
'    lCountRegs = 0
'
'    'Enquanto existirem arquivos
'    Do While lErro = AD_SQL_SUCESSO
'
'        lCountRegs = lCountRegs + 1
'
'        'Se modificou o número do carão de crédito, o administradora, a data ou o número de autorização deve gerar um novo título
'        If sCiaAnt <> sCia Or sNumCCredAnt <> sNumCCred Or dtData <> dtDataAnt Or sNumAutoAnt <> sNumAuto Then
'            sCiaAnt = sCia
'            sNumCCredAnt = sNumCCred
'            dtDataAnt = dtData
'            sNumAutoAnt = sNumAuto
'
'            If colVoucher.Count > 0 Then colcolVoucher.Add colVoucher
'
'            Set colVoucher = New Collection
'        End If
'
'        Set objVoucher = New ClassTRVVouchers
'
'        objVoucher.lNumVou = lNumVou
'        objVoucher.sSerie = sSerie
'        objVoucher.sTipVou = sTipo
'        objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
'
'        colVoucher.Add objVoucher
'
'        'Busca o próximo arquivo
'        lErro = Comando_BuscarProximo(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192707
'
'        If lErro = AD_SQL_SEM_DADOS Then
'            If colVoucher.Count > 0 Then colcolVoucher.Add colVoucher
'        End If
'
'    Loop
'
'    lErro = Coinfo_Total_Registros2("Títulos CRT", "14 de 15", lCountRegs)
'    If lErro <> SUCESSO Then gError 192708
'
''    lErro = CF("TRVConfig_Le", TRVCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
''    If lErro <> SUCESSO Then gError 192712
''
''    lProxTitRec = StrParaLong(sConteudo)
'
'    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sModeloFat)
'    If lErro <> SUCESSO Then gError 192712
'
'    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sNomeDiretorio)
'    If lErro <> SUCESSO Then gError 192712
'
'    For Each colVoucher In colcolVoucher
'
'        lTransacao = Transacao_Abrir()
'        If lTransacao = 0 Then gError 192709
'
'        lErro = CF("TRVConfig_Le_Lockado", alComando(5), TRVCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
'        If lErro <> SUCESSO Then gError 192712
'
'        lProxTitRec = StrParaLong(sConteudo)
'
'        Set objTitRec = New ClassTituloReceber
'
'        dValor = 0
'        For Each objVoucher In colVoucher
'
'            lErro = CF("TRVVouchers_Le", objVoucher)
'            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192710
'
'            dValor = dValor + objVoucher.dValor
'
'        Next
'
'        Set objVoucher = colVoucher.Item(1)
'        Set objVoucherInfo = New ClassTRVVoucherInfo
'
'        objVoucherInfo.lNumVou = objVoucher.lNumVou
'        objVoucherInfo.sSerie = objVoucher.sSerie
'        objVoucherInfo.sTipo = objVoucher.sTipVou
'
'        lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo)
'        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 196541
'
'        If lErro <> SUCESSO Then gError 196542 'ERRO_VOUCHER_SEM_DADOS_SIGAV
'
'        lErro = CF("Coinfo_Obtem_FilialEmp_Cliente", objVoucherInfo.lCliente, iFilialCorporator, iFilialCoinfo)
'        If lErro <> SUCESSO Then gError 192711
'
'        Set objImpCoinfoCartoes = New ClassImpCoinfoCartoes
'
'        objImpCoinfoCartoes.dtDataDep = objVoucher.dtData
'        objImpCoinfoCartoes.iFilialEmpresa = iFilialCorporator
'        objImpCoinfoCartoes.dValor = dValor
'        objImpCoinfoCartoes.lNumTitulo = lProxTitRec
'        lProxTitRec = lProxTitRec + 1
'        objImpCoinfoCartoes.sCiaCart = objVoucher.sCiaCart
'        objImpCoinfoCartoes.iQuantParc = objVoucher.iQuantParc
'
'        'Gera o parcelamento do título e grava
'        lErro = Coinfo_GerarTitRec_Cartao(objImpCoinfoCartoes)
'        If lErro <> SUCESSO Then gError 192713
'
'        objTitRec.lNumIntDoc = objImpCoinfoCartoes.lNumIntTitulo
'
'        lErro = CF("TituloReceber_Le", objTitRec)
'        If lErro <> SUCESSO And lErro <> 26061 Then gError 192724
'
'        dtDataEmissaoAux = objTitRec.dtDataEmissao
'        objTitRec.dtDataEmissao = Date
'
''        dTaxa = objImpCoinfoCartoes.dValorTaxa
''        dDeducoes = objImpCoinfoCartoes.dDeducaoValorPresente
'
'        Set objContabAutomatica = New ClassContabAutomatica
'        Set objCtb = New ClassImpCoinfoFatCRCtb
'
'        'gravar a contabilizacao
'        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
'        If lErro <> SUCESSO Then gError 192714
'
'        lErro = objContabAutomatica.GeraContabilizacao(objTitRec, objImpCoinfoCartoes)
'        If lErro <> SUCESSO Then gError 192715
'
'        Set objContabAutomatica = Nothing
'        Call objCtb.Encerrando
'        Set objCtb = Nothing
'
'        Set objTitRecTRV = New ClassTitulosRecTRV
'
'        objTitRecTRV.lNumIntDocTitRec = objTitRec.lNumIntDoc
'        objTitRecTRV.dValorBruto = objImpCoinfoCartoes.dValor
'        objTitRecTRV.dValorDeducoes = objImpCoinfoCartoes.dDeducaoValorPresente
'        objTitRecTRV.dValorTarifa = objImpCoinfoCartoes.dValorTaxa
'
'        'Atualiza a informação da Fatura no Voucher
'        For Each objVoucher In colVoucher
'
'            sNumAuto = String(STRING_TRV_VOU_NUMAUTO, 0)
'
'            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumAuto FROM TRVVouchers WHERE NumVou = ? AND Serie = ? AND TipVou = ? AND TipoDoc = ? ", 0, sNumAuto, objVoucher.lNumVou, objVoucher.sSerie, objVoucher.sTipVou, objVoucher.sTipoDoc)
'            If lErro <> AD_SQL_SUCESSO Then gError 192717
'
'            'Busca o primeiro
'            lErro = Comando_BuscarPrimeiro(alComando(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192718
'
'            'Marca o arquivo como atualizado
'            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TRVVouchers SET NumFatCoinfo = ?, TipoDocDestino = ?, NumIntDocDestino = ?", alComando(1), objTitRec.lNumTitulo, TRV_TIPO_DOC_DESTINO_TITREC, objTitRec.lNumIntDoc)
'            If lErro <> SUCESSO Then gError 192719
'
'            If objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO Then
'                Set objVouExp = New ClassTRVVouExp
'                objVouExp.sSerie = objVoucher.sSerie
'                objVouExp.sTipVou = objVoucher.sTipVou
'                objVouExp.lNumVou = objVoucher.lNumVou
'                objVouExp.lNumFat = objTitRec.lNumTitulo
'                objVouExp.dtDataFat = objTitRec.dtDataEmissao
'                objVouExp.iTransacao = TRV_EXPORT_VOU_TRANS_FATURAMENTO
'                objVouExp.lCliente = objVoucher.lClienteVou
'                lErro = CF("TRVVouExp_Insere", alComando(7), alComando(8), alComando(9), alComando(10), alComando(11), alComando(12), objVouExp)
'                If lErro <> SUCESSO Then gError 200212
'            End If
'
'        Next
'
'        objTitRec.dtDataEmissao = dtDataEmissaoAux
'
'        lErro = Coinfo_GerarComissaoCartao(objImpCoinfo, objTitRec, objTitRecTRV, True)
'        If lErro <> SUCESSO Then gError 192716
'
'        lErro = CF("TitulosRecTRV_Grava", objTitRecTRV)
'        If lErro <> SUCESSO Then gError 192716
'
'        Set objTRVTitulosExp = New ClassTRVTitulosExp
'
'        objTRVTitulosExp.iExcluido = DESMARCADO
'        objTRVTitulosExp.iExportado = DESMARCADO
'        objTRVTitulosExp.iTipoDocOrigem = TRV_TIPO_DOC_DESTINO_TITREC
'        objTRVTitulosExp.lNumIntDocOrigem = objTitRec.lNumIntDoc
'        objTRVTitulosExp.lNumTitulo = objTitRec.lNumTitulo
'        objTRVTitulosExp.dValorDeducoes = objImpCoinfoCartoes.dDeducaoValorPresente
'        objTRVTitulosExp.dValorTarifa = objImpCoinfoCartoes.dValorTaxa
'
'        lErro = CF("TRVTitulosExp_Insere", objTRVTitulosExp)
'        If lErro <> SUCESSO Then gError 196240
'
'        Set objTRVTitulos = New ClassTRVTitulos
'
'        objTRVTitulos.iTipoDocOrigem = TRV_TIPO_DOC_DESTINO_TITREC
'        objTRVTitulos.lNumIntDocOrigem = objTitRec.lNumIntDoc
'        objTRVTitulos.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
'
'        lErro = CF("TRVTitulos_Insere", objTRVTitulos)
'        If lErro <> SUCESSO Then gError 196240
'
'        Set objTRVConfig = New ClassTRVConfig
'        'Set colConfig = New Collection
'
'        objTRVConfig.sCodigo = TRVCONFIG_PROX_NUM_TITREC
'        objTRVConfig.sConteudo = CStr(lProxTitRec)
'
'        'colConfig.Add objTRVConfig
'
'        'gera o documento html
'        lErro = CF("TRVFaturas_Regera_Html", objTitRec.lNumTitulo, objTitRec.lNumTitulo, sModeloFat, sNomeDiretorio, objTitRec.sSiglaDocumento)
'        If lErro <> SUCESSO Then gError 192713
'
'        'lErro = CF("TRVConfig_Grava_Trans", colConfig)
'        lErro = CF("TRVConfig_Grava_Lockado", alComando(5), alComando(6), objTRVConfig)
'        If lErro <> SUCESSO Then gError 192723
'
'        'Confirma a transação
'        lErro = Transacao_Commit()
'        If lErro <> AD_SQL_SUCESSO Then gError 182720
'
'        lErro = Rotina_Atualizar_AtualizaTelaBatch
'        If lErro <> SUCESSO Then gError 192721
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Coinfo_AtualizarVouchers3 = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarVouchers3:
'
'    Coinfo_AtualizarVouchers3 = gErr
'
'    Select Case gErr
'
'        Case 192704
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 192705 To 192707, 192717, 192718
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
'
'        Case 192708, 192710 To 192716, 182721, 182723, 182724, 196240, 200212
'
'        Case 192709
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 192719
'            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVVOUCHERS", gErr)
'
'        Case 192720
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 196541
'
'        Case 196542
'            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_DADOS_SIGAV", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182722)
'
'    End Select
'
'    'Desfaz a transação
'    Call Transacao_Rollback
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Private Function Extrair_Vouchers_Sigav()

Dim lErro As Long
Dim objTRVVoucherInfo As ClassTRVVoucherInfo
Dim objTRVVoucher As ClassTRVVouchers
Dim colVoucher As New Collection
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Extrair_Vouchers_Sigav
    
    lErro = CF("TRVVouchers_Le_Periodo", DATA_NULA, DATA_NULA, colVoucher, True)
    If lErro <> SUCESSO Then gError 192863

    lErro = Coinfo_Total_Registros2("Extraindo dados do Sigav", "13 de 15", colVoucher.Count)
    If lErro <> SUCESSO Then gError 192863

    For Each objTRVVoucher In colVoucher
    
        Set objTRVVoucherInfo = New ClassTRVVoucherInfo
        
        objTRVVoucherInfo.sTipo = objTRVVoucher.sTipVou
        objTRVVoucherInfo.lNumVou = objTRVVoucher.lNumVou
        objTRVVoucherInfo.sSerie = objTRVVoucher.sSerie
            
        lErro = Obter_Dados_Sigav(objTRVVoucherInfo, "forprint")
        If lErro = SUCESSO Then
        
            lErro = CF("TRVVoucherInfoSigav_Grava", objTRVVoucherInfo)
            'If lErro <> SUCESSO Then gError 192865
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 192863
        
        End If
        
    Next
    
    Extrair_Vouchers_Sigav = SUCESSO

    Exit Function

Erro_Extrair_Vouchers_Sigav:

    Extrair_Vouchers_Sigav = gErr

    Select Case gErr
            
        Case 192863 To 192865

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192866)

    End Select

    Exit Function

End Function

'Private Function Coinfo_AtualizarOcr(ByVal objImpCoinfo As ClassImpCoinfo) As Long
''atualiza os fornecedores
'
'Dim lErro As Long
'Dim lTransacaoExt As Long
'Dim lTransacao As Long
'Dim alComandoExt(0 To 4) As Long
'Dim alComando(0 To 4) As Long
'Dim iIndice As Integer
'Dim lNumIntArq As Long
'Dim objImpCoinfoOcor As ClassImpCoinfoOcor
'Dim sSelect As String
'Dim tImpCoinfoOcor As typeImpCoinfoOcor
'
'On Error GoTo Erro_Coinfo_AtualizarOcr
'
'    'Abre o arquivo
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
'        If alComandoExt(iIndice) = 0 Then gError 185496
'    Next
'
'    'Abre o arquivo
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 185497
'    Next
'
'    'Obtém o total de registros que vão ser atualizados
'    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpocorrencia")
'    If lErro <> SUCESSO Then gError 185498
'
'    'Buscar o NumIntDoc dos Títulos
'    sSelect = "SELECT NumRegArqTexto, numocorr, data, valor, codest, " & _
'                "usuario, datareg, recnsica, TipVou, Serie, NumVou, dc, " & _
'                "CodEmp, texto1, texto2, texto3, texto4, texto5, texto6, " & _
'                "liberado, Controle, ordem FROM CoinfoOcorrencia WHERE NumIntArq= ? AND CodEst = ? AND NoCorporatorEm = ? AND EXISTS (SELECT NumVou FROM TRVVouchers AS V WHERE V.NumVou = CoinfoOcorrencia.NumVou AND V.TipVou = CoinfoOcorrencia.TipVou AND V.Serie =CoinfoOcorrencia.Serie AND V.TipoDoc = 'VOU' ) ORDER BY NumRegArqTexto"
'
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 185499
'
'    'Lê os arquivos que não foram atualizados
'    lErro = Comando_ExecutarPos(alComandoExt(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
'        lNumIntArq, "fpocorrencia", DATA_NULA)
'    If lErro <> AD_SQL_SUCESSO Then gError 185500
'
'    'Busca o primeiro
'    lErro = Comando_BuscarProximo(alComandoExt(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185501
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        tImpCoinfoOcor.scodest = String(255, 0)
'        tImpCoinfoOcor.sUsuario = String(255, 0)
'        tImpCoinfoOcor.sTipVou = String(255, 0)
'        tImpCoinfoOcor.sSerie = String(255, 0)
'        tImpCoinfoOcor.stexto1 = String(255, 0)
'        tImpCoinfoOcor.stexto2 = String(255, 0)
'        tImpCoinfoOcor.stexto3 = String(255, 0)
'        tImpCoinfoOcor.stexto4 = String(255, 0)
'        tImpCoinfoOcor.stexto5 = String(255, 0)
'        tImpCoinfoOcor.stexto6 = String(255, 0)
'        tImpCoinfoOcor.sControle = String(255, 0)
'
'        'Le o títulos que deveriam ter comissões e não as tem
'        lErro = Comando_Executar(alComandoExt(1), sSelect, _
'                tImpCoinfoOcor.lNumRegArqTexto, tImpCoinfoOcor.lnumocorr, tImpCoinfoOcor.dtData, _
'                tImpCoinfoOcor.dValor, tImpCoinfoOcor.scodest, tImpCoinfoOcor.sUsuario, tImpCoinfoOcor.dtdatareg, tImpCoinfoOcor.lrecnsica, _
'                tImpCoinfoOcor.sTipVou, tImpCoinfoOcor.sSerie, tImpCoinfoOcor.lNumVou, tImpCoinfoOcor.idc, tImpCoinfoOcor.lcodemp, _
'                tImpCoinfoOcor.stexto1, tImpCoinfoOcor.stexto2, tImpCoinfoOcor.stexto3, tImpCoinfoOcor.stexto4, tImpCoinfoOcor.stexto5, _
'                tImpCoinfoOcor.stexto6, tImpCoinfoOcor.iliberado, tImpCoinfoOcor.sControle, tImpCoinfoOcor.lordem, _
'                 lNumIntArq, objImpCoinfo.sSiglaEmpresa, DATA_NULA)
'        If lErro <> AD_SQL_SUCESSO Then gError 185502
'
'        'Busca o primeiro
'        lErro = Comando_BuscarProximo(alComandoExt(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185503
'
'        'Enquando existirem registros
'        Do While lErro = AD_SQL_SUCESSO
'
'            'Abertura de transação
'            lTransacao = Transacao_Abrir()
'            If lTransacao = 0 Then gError 185504
'
'            Set objImpCoinfoOcor = New ClassImpCoinfoOcor
'
'            objImpCoinfoOcor.lnumocorr = tImpCoinfoOcor.lnumocorr
'            objImpCoinfoOcor.dtData = tImpCoinfoOcor.dtData
'            objImpCoinfoOcor.dValor = tImpCoinfoOcor.dValor
'            objImpCoinfoOcor.scodest = tImpCoinfoOcor.scodest
'            objImpCoinfoOcor.sUsuario = tImpCoinfoOcor.sUsuario
'            objImpCoinfoOcor.dtdatareg = tImpCoinfoOcor.dtdatareg
'            objImpCoinfoOcor.lrecnsica = tImpCoinfoOcor.lrecnsica
'            objImpCoinfoOcor.sTipVou = tImpCoinfoOcor.sTipVou
'            objImpCoinfoOcor.sSerie = tImpCoinfoOcor.sSerie
'            objImpCoinfoOcor.lNumVou = tImpCoinfoOcor.lNumVou
'            objImpCoinfoOcor.idc = tImpCoinfoOcor.idc
'            objImpCoinfoOcor.lcodemp = tImpCoinfoOcor.lcodemp
'            objImpCoinfoOcor.stexto1 = tImpCoinfoOcor.stexto1
'            objImpCoinfoOcor.stexto2 = tImpCoinfoOcor.stexto2
'            objImpCoinfoOcor.stexto3 = tImpCoinfoOcor.stexto3
'            objImpCoinfoOcor.stexto4 = tImpCoinfoOcor.stexto4
'            objImpCoinfoOcor.stexto5 = tImpCoinfoOcor.stexto5
'            objImpCoinfoOcor.stexto6 = tImpCoinfoOcor.stexto6
'            objImpCoinfoOcor.iliberado = tImpCoinfoOcor.iliberado
'            objImpCoinfoOcor.sControle = tImpCoinfoOcor.sControle
'            objImpCoinfoOcor.lordem = tImpCoinfoOcor.lordem
'
'            If objImpCoinfoOcor.dtData > #12/31/2006# Then
'
'                If objImpCoinfo.objFiltro.iConsiderarFatSigav = MARCADO Then
'                    lErro = CF("TRVOcorrencias_Importa", objImpCoinfoOcor, True)
'                Else
'                    lErro = CF("TRVOcorrencias_Importa", objImpCoinfoOcor)
'                End If
'                If lErro <> SUCESSO Then gError 185509
'
'            End If
'
'            lErro = Comando_Executar(alComando(0), "UPDATE CoinfoOcorrencia SET TipoDocCorporator = ?, NumIntDocCorporator = ?, NoCorporatorEm = ? WHERE NumIntArq = ? AND NumRegArqTexto = ?", TIPO_OCORRENCIA_TRV, objImpCoinfoOcor.lNumIntDoc, gdtDataHoje, lNumIntArq, tImpCoinfoOcor.lNumRegArqTexto)
'            If lErro <> SUCESSO Then gError 182540
'
'            'Confirma a transação
'            lErro = Transacao_Commit()
'            If lErro <> AD_SQL_SUCESSO Then gError 185508
'
'            lErro = Rotina_Atualizar_AtualizaTelaBatch
'            If lErro <> SUCESSO Then gError 185509
'
'            lErro = Comando_BuscarProximo(alComandoExt(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185510
'
'        Loop
'
'        lErro = Comando_ExecutarPos(alComandoExt(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComandoExt(0), gdtDataHoje)
'        If lErro <> SUCESSO Then gError 185511
'
'        lErro = Comando_BuscarProximo(alComandoExt(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185512
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacaoExt)
'    If lErro <> AD_SQL_SUCESSO Then gError 185513
'
'    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Ocorrências foram atualizadas com sucesso. " & vbNewLine
'
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Coinfo_AtualizarOcr = SUCESSO
'
'    Exit Function
'
'Erro_Coinfo_AtualizarOcr:
'
'    Coinfo_AtualizarOcr = gErr
'
'    Select Case gErr
'
'        Case 185496, 185497
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 185498, 185505, 185506, 185507, 185509
'
'        Case 185499, 185504
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 185500, 185501, 185512
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
'
'        Case 185502, 185503, 185510
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOCOMISSAO", gErr)
'
'        Case 185508, 185513
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 185511
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185514)
'
'    End Select
'
'    Call Transacao_RollbackExt(lTransacaoExt)
'    Call Transacao_Rollback
'
'    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
'        Call Comando_Fechar(alComandoExt(iIndice))
'    Next
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Public Function Coinfo_NVLRetroativo(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'Função auxiliar que não é chamada em nenhum ponto
'Cancela os vouchers que tem documento NVL

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim lNumVou As Long
Dim sSerie As String
Dim sTipo As String
Dim objVoucher As ClassTRVVouchers
Dim lNumFat As Long
Dim iCancelado As Integer
Dim iTipoNumIntDocDest As Integer
Dim lNumIntDocDest As Long
Dim dValor As Double
Dim lNumIntDocOCR As Long
Dim lCodigoOcr As Long
Dim iTipoDestino As Integer
Dim dtData As Date
Dim lCliente As Long
Dim iCartao As Integer

On Error GoTo Erro_Coinfo_NVLRetroativo

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459
    
    sSerie = String(255, 0)
    sTipo = String(255, 0)

    'buscar a fatura como titulo a receber sem data de vencimento
    lErro = Comando_Executar(alComando(0), "SELECT NumVou, Serie, TipVou, NumFat, Valor, Data, CodEmp, Cartao FROM VouchersTodos WHERE TipoDoc = 'NVL' AND CodEst = ? AND Data >= ? AND Excluido <> 1 AND NumFat = 0 ORDER BY NumIntArq, NumRegArqTexto ", _
        lNumVou, sSerie, sTipo, lNumFat, dValor, dtData, lCliente, iCartao, objImpCoinfo.sSiglaEmpresa, #1/1/2007#)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver títulos
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objVoucher = New ClassTRVVouchers
        
        objVoucher.lNumVou = lNumVou
        objVoucher.sSerie = sSerie
        objVoucher.sTipVou = sTipo
        objVoucher.sTipoDoc = TRV_TIPODOC_NVL_TEXTO
        objVoucher.lCliente = lCliente
        objVoucher.dValor = dValor
        objVoucher.iCartao = iCartao
        
        lNumIntDocOCR = 0
        
        lErro = CF("TRVVouchers_Cancela", objVoucher, True, dValor, dtData, lNumIntDocOCR)
        If lErro <> SUCESSO Then gError 99999
        
        If lNumFat <> 0 Then
        
            'Le a tabela TitulosRecTRV
            lErro = Comando_Executar(alComando(1), "SELECT Cancelado, TipoDocCorporator, NumIntDocCorporator FROM CoinfoFatura WHERE CodEst = ? AND NumFat = ? AND NumIntDocCorporator <> 0 ORDER BY  NumIntArq DESC,  NumRegArqTexto DESC", iCancelado, iTipoNumIntDocDest, lNumIntDocDest, objImpCoinfo.sSiglaEmpresa, lNumFat)
            If lErro <> AD_SQL_SUCESSO Then gError 194439
            
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194440
        
            If iCancelado = DESMARCADO And lNumIntDocOCR <> 0 Then
            
                lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM TRVOcorrencias WHERE NumIntDoc= ? ", _
                            0, lCodigoOcr, lNumIntDocOCR)
                If lErro <> AD_SQL_SUCESSO Then gError 194441
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 194442
                
                If lErro = AD_SQL_SEM_DADOS Then gError 194433
            
                'Trava registro contra alterações/Leituras
                lErro = Comando_LockExclusive(alComando(2))
                If lErro <> SUCESSO Then gError 194434
                
                Select Case iTipoNumIntDocDest
                
                    Case TIPO_TITULO_PAGAR_TRV
                        iTipoDestino = TRV_TIPO_DOC_DESTINO_TITPAG
                        
                    Case TIPO_TITULO_RECEBER_TRV
                        iTipoDestino = TRV_TIPO_DOC_DESTINO_TITREC
                
                    Case TIPO_NF_PAGAR_TRV
                        iTipoDestino = TRV_TIPO_DOC_DESTINO_NFSPAG
                        
                    Case Else
                        iTipoDestino = 0
                        lNumIntDocDest = 0
                
                End Select
        
                lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ?, Status = ? ", alComando(2), _
                iTipoDestino, lNumIntDocDest, lNumFat, STATUS_TRV_OCR_LIBERADO)
                If lErro <> AD_SQL_SUCESSO Then gError 194435
            
            End If
            
        End If
            
        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    
    MsgBox "Rotina finalizada com sucesso"
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_NVLRetroativo = SUCESSO
     
    Exit Function
    
Erro_Coinfo_NVLRetroativo:

    Coinfo_NVLRetroativo = gErr
     
    Select Case gErr
    
        Case 182532, 182533, 182542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOFATURA", gErr)
            
        Case 182534, 182535
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_RECEBER", gErr)
        
        Case 182536, 182537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADO", gErr)

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182539, 182543, 182544, 182720, 182743
        
        Case 182540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COINFOFATURA", gErr)

        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRVCliente_Grava_Retroativo() As Long

Dim lErro As Long
Dim objClienteFaixa As New ClassClienteFaixa
Dim objCliente As ClassCliente
Dim colClientes As New Collection
Dim colCliEst As New Collection
Dim lTransacao As Long

On Error GoTo Erro_TRVCliente_Grava_Retroativo

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 196341
    
    'objClienteFaixa.lClienteAte = 3003
    'objClienteFaixa.lClienteDe = 3003

    lErro = CF("Cliente_Le_Faixa", objClienteFaixa, colClientes, colCliEst)
    If lErro <> SUCESSO Then gError 196338
    
    For Each objCliente In colClientes
    
        lErro = CF("TRVCliente_Grava", objCliente)
        If lErro <> SUCESSO Then
            gError 196339
        End If
    
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 196342
    
    MsgBox "Rotina executada com sucesso."

    TRVCliente_Grava_Retroativo = SUCESSO

    Exit Function

Erro_TRVCliente_Grava_Retroativo:

    TRVCliente_Grava_Retroativo = gErr

    Select Case gErr

        Case 196338, 196339

        Case 196341
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 196342
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196340)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Public Function TRVGera_Comissao_Retroativo(Optional ByVal iAno As Integer = 0, Optional ByVal iMes As Integer = 0) As Long
'Função auxiliar que não é chamada em nenhum ponto
'Cancela os vouchers que tem documento NVL

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 5) As Long
Dim lNumVou As Long
Dim sSerie As String
Dim sTipo As String
Dim objVou As ClassTRVVouchers
Dim lNumIntDocVouInfo As Long

On Error GoTo Erro_TRVGera_Comissao_Retroativo

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459
    
    sSerie = String(255, 0)
    sTipo = String(255, 0)
    
    'buscar os vouchers de 2008
    If iAno = 0 Or iMes = 0 Then
        lErro = Comando_Executar(alComando(0), "SELECT V1.NumVou, V1.Serie, V1.TipVou FROM TRVVouchers AS V1, TRVVoucherInfoSigav AS V2 WHERE V1.NumVou = V2.numvou AND V2.Tipo = V1.TipVou AND V1.TipoDoc = 'VOU' AND V1.Serie = V2.Serie AND Data >= {d '2008-01-01'} AND V1.Status <> 7 ", _
            lNumVou, sSerie, sTipo)
    Else
         lErro = Comando_Executar(alComando(0), "SELECT V1.NumVou, V1.Serie, V1.TipVou FROM TRVVouchers AS V1, TRVVoucherInfoSigav AS V2 WHERE V1.NumVou = V2.numvou AND V2.Tipo = V1.TipVou AND V1.TipoDoc = 'VOU' AND V1.Serie = V2.Serie AND YEAR(Data) = ? AND MONTH(Data) = ? AND V1.Status <> 7 ", _
            lNumVou, sSerie, sTipo, iAno, iMes)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRVVouchers
        
        objVou.lNumVou = lNumVou
        objVou.sSerie = sSerie
        objVou.sTipVou = sTipo
        objVou.sTipoDoc = "VOU"
        
        'Busca o histórico relacionado ao voucher para liberá-lo para o comissionamento
        lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc FROM TRVVoucherInfo WHERE TipVou = ? AND Serie= ? AND NumVou = ? ", 0, lNumIntDocVouInfo, sTipo, sSerie, lNumVou)
        If lErro <> AD_SQL_SUCESSO Then gError 196362
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196363
        
        'Enquanto existirem regsitros
        Do While lErro <> AD_SQL_SEM_DADOS
        
            'Trava registro contra alterações/Leituras
            lErro = Comando_LockExclusive(alComando(4))
            If lErro <> SUCESSO Then gError 196364
    
            lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM TRVVoucherInfo ", alComando(4))
            If lErro <> AD_SQL_SUCESSO Then gError 196365
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196366
        
        Loop
        
        lErro = CF("TRVVouchers_Le", objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192874

        lErro = CF("TRVVoucher_Gera_Comissao", objVou, objVou.dValor, objVou.dtData, "Voucher emitido", TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO)
        'If lErro <> SUCESSO Then gError 196314
    
        'Busca a próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    
    'MsgBox "Rotina finalizada com sucesso"
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TRVGera_Comissao_Retroativo = SUCESSO
     
    Exit Function
    
Erro_TRVGera_Comissao_Retroativo:

    TRVGera_Comissao_Retroativo = gErr
     
    Select Case gErr
    
        Case 185496
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRV_Acerta_NVL() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Cancela os vouchers que tem documento NVL

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim lCliente As Long
Dim dtDataEmissao As Date
Dim dValor As Double
Dim iFilialEmpresa As Integer
Dim dtDataVencimento As Date
Dim dValorComissao As Double
Dim lNumIntDocOCR As Long
Dim lNumIntDocNFsPag As Long
Dim lNumIntDocNFsPagAnt As Long
Dim lNumTitulo As Long
Dim objTitPag As ClassTituloPagar
Dim objNFPag As ClassNFsPag
Dim objContabil As ClassContabil
Dim lNumIntDocAux As Long
Dim dtDataVenctoReal As Date
Dim objParcelaPagar As ClassParcelaPagar
Dim colParcelaPagar As colParcelaPagar
Dim objContabAutomatica As ClassContabAutomatica
Dim objCtb As New ClassImpCoinfoFatCPCtb

On Error GoTo Erro_TRV_Acerta_NVL

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    
    'buscar os vouchers de 2008
    lErro = Comando_Executar(alComando(0), "SELECT F.NumIntDoc, F.NumintDocDestino,F.Cliente, F.DataEmissao, F.ValorTitulo, F.FilialEmpresa, N.DataVencimento, F.ValorComissao, F.NumTitulo FROM DocFaturados AS F, NfsPag AS N WHERE F.TipoDocDestino = 5 AND F.TipoDoc IN ('OCR','NVL') AND F.ValorTitulo <> 0 AND F.DataEmissao >= {d '2008-02-01'} AND N.NumIntDoc = F.NumIntDocDestino ORDER BY F.NumTitulo", _
    lNumIntDocOCR, lNumIntDocNFsPag, lCliente, dtDataEmissao, dValor, iFilialEmpresa, dtDataVencimento, dValorComissao, lNumTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If lNumIntDocNFsPagAnt <> lNumIntDocNFsPag Then
        
            lNumIntDocNFsPagAnt = lNumIntDocNFsPag
    
            '################################################
            'Exclui o NFPAG
            Set objNFPag = New ClassNFsPag
         
            objNFPag.lNumIntDoc = lNumIntDocNFsPag
         
            lErro = CF("NFPag_Le", objNFPag)
            If lErro = SUCESSO Then
    
                Set objContabil = New ClassContabil
                Call objContabil.Contabil_Inicializa_Contabilidade4(14, MODULO_CONTASAPAGAR)
             
                objNFPag.iExclusao = MARCADO
             
                lErro = CF("NFPag_Exclui_EmTransacao", objNFPag, objContabil)
                If lErro <> SUCESSO Then gError 182534
                Set objContabil = Nothing
                
            End If
            '#################################################
            
            
            '#################################################
            'Gera o TITPAG
            Set objTitPag = New ClassTituloPagar
            Set objParcelaPagar = New ClassParcelaPagar
            Set colParcelaPagar = New colParcelaPagar
            Set objContabAutomatica = New ClassContabAutomatica
            
            objTitPag.lFornecedor = lCliente
            objTitPag.iFilial = FILIAL_MATRIZ
            objTitPag.iFilialPedCompra = FILIAL_MATRIZ
            objTitPag.lNumTitulo = lNumTitulo
            objTitPag.dtDataEmissao = dtDataEmissao
            objTitPag.dValorTotal = dValor
            objTitPag.dValorProdutos = dValor
            objTitPag.dSaldo = dValor
            objTitPag.iFilialEmpresa = iFilialEmpresa
            objTitPag.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CP
            objTitPag.iNumParcelas = 1
        
            objParcelaPagar.iNumParcela = 1
            objParcelaPagar.dtDataVencimento = dtDataVencimento
            
            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then gError 190997
            
            objParcelaPagar.dtDataVencimentoReal = dtDataVenctoReal
            objParcelaPagar.dValor = dValor
            objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
            objParcelaPagar.iStatus = STATUS_ABERTO
        
            With objParcelaPagar
                colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
            End With
        
            'Grava o Título a Pagar
            lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
            If lErro <> SUCESSO Then gError 190998
            
            objCtb.sModelo = "Reembolso"
            
            'gravar a contabilizacao
            lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
            If lErro <> SUCESSO Then gError 190999
            
            lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
            If lErro <> SUCESSO Then gError 192000
        
            Set objContabAutomatica = Nothing
            Call objCtb.Encerrando
            Set objCtb = Nothing
            '#################################################
            
        End If
    
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM TRVOcorrencias WHERE NumIntDoc= ? ", _
                    0, lNumIntDocAux, lNumIntDocOCR)
        If lErro <> AD_SQL_SUCESSO Then gError 192007
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192008
        
        If lErro = AD_SQL_SEM_DADOS Then gError 192018
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> SUCESSO Then gError 192009

        lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ? ", alComando(2), _
        4, objTitPag.lNumIntDoc, lNumTitulo)
        If lErro <> AD_SQL_SUCESSO Then gError 192010
    
        'Busca a próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    MsgBox "Rotina finalizada com sucesso"
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TRV_Acerta_NVL = SUCESSO
     
    Exit Function
    
Erro_TRV_Acerta_NVL:

    TRV_Acerta_NVL = gErr
     
    Select Case gErr
    
        Case 185496
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function


Public Function TRV_Acerta_NVL2() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Cancela os vouchers que tem documento NVL

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim lCliente As Long
Dim dtDataEmissao As Date
Dim dValor As Double
Dim iFilialEmpresa As Integer
Dim dtDataVencimento As Date
Dim dValorComissao As Double
Dim lNumIntDocOCR As Long
Dim lNumIntDocNFsPag As Long
Dim lNumIntDocNFsPagAnt As Long
Dim lNumTitulo As Long
Dim objTitPag As ClassTituloPagar
Dim objNFPag As ClassNFsPag
Dim objContabil As ClassContabil
Dim lNumIntDocAux As Long
Dim dtDataVenctoReal As Date
Dim objParcelaPagar As ClassParcelaPagar
Dim colParcelaPagar As colParcelaPagar
Dim objContabAutomatica As ClassContabAutomatica
Dim objCtb As New ClassImpCoinfoFatCPCtb

On Error GoTo Erro_TRV_Acerta_NVL2

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459

    
    'buscar os vouchers de 2008
    lErro = Comando_Executar(alComando(0), "SELECT F.NumIntDoc, F.NumintDocDestino,F.Cliente, F.DataEmissao, F.ValorTitulo, F.FilialEmpresa, N.DataVencimento, F.ValorComissao, F.NumTitulo FROM DocFaturados AS F, TRVFaturasAuxCP AS N WHERE F.TipoDocDestino = 4 AND F.TipoDoc IN ('OCR','NVL') AND F.ValorTitulo < 0 AND F.DataEmissao >= {d '2008-02-01'} AND N.NumIntTitulo = F.NumintDocDestino ORDER BY F.NumTitulo", _
    lNumIntDocOCR, lNumIntDocNFsPag, lCliente, dtDataEmissao, dValor, iFilialEmpresa, dtDataVencimento, dValorComissao, lNumTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533

    Do While lErro <> AD_SQL_SEM_DADOS
    
        dValor = Abs(dValor)
    
        If lNumIntDocNFsPagAnt <> lNumIntDocNFsPag Then
        
            lNumIntDocNFsPagAnt = lNumIntDocNFsPag
    
             Set objTitPag = New ClassTituloPagar
             
             objTitPag.lNumIntDoc = lNumIntDocNFsPag
             
             lErro = CF("TituloPagar_Le1", objTitPag)
             If lErro <> SUCESSO And lErro <> 18372 Then gError 182533
             If lErro = SUCESSO Then
        
                 Set objContabil = New ClassContabil
                 Call objContabil.Contabil_Inicializa_Contabilidade4(14, MODULO_CONTASAPAGAR)
                 
                 objTitPag.iExclusao = MARCADO
                 
                 lErro = CF("TituloPagar_Exclui_EmTransacao", objTitPag, objContabil)
                 If lErro <> SUCESSO Then gError 182534
                 Set objContabil = Nothing
                 
             End If
            
            '#################################################
            'Gera o TITPAG
            Set objTitPag = New ClassTituloPagar
            Set objParcelaPagar = New ClassParcelaPagar
            Set colParcelaPagar = New colParcelaPagar
            Set objContabAutomatica = New ClassContabAutomatica
            
            objTitPag.lFornecedor = lCliente
            objTitPag.iFilial = FILIAL_MATRIZ
            objTitPag.iFilialPedCompra = FILIAL_MATRIZ
            objTitPag.lNumTitulo = lNumTitulo
            objTitPag.dtDataEmissao = dtDataEmissao
            objTitPag.dValorTotal = dValor
            objTitPag.dValorProdutos = dValor
            objTitPag.dSaldo = dValor
            objTitPag.iFilialEmpresa = iFilialEmpresa
            objTitPag.sSiglaDocumento = TIPODOC_FATURA_SERVICO_CP
            objTitPag.iNumParcelas = 1
        
            objParcelaPagar.iNumParcela = 1
            objParcelaPagar.dtDataVencimento = dtDataVencimento
            
            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then gError 190997
            
            objParcelaPagar.dtDataVencimentoReal = dtDataVenctoReal
            objParcelaPagar.dValor = dValor
            objParcelaPagar.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
            objParcelaPagar.iStatus = STATUS_ABERTO
        
            With objParcelaPagar
                colParcelaPagar.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
            End With
        
            'Grava o Título a Pagar
            lErro = CF("NFFatPag_Grava_EmTrans", objTitPag, colParcelaPagar, objContabil)
            If lErro <> SUCESSO Then gError 190998
            
            objCtb.sModelo = "Reembolso"
            
            'gravar a contabilizacao
            lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
            If lErro <> SUCESSO Then gError 190999
            
            lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
            If lErro <> SUCESSO Then gError 192000
        
            Set objContabAutomatica = Nothing
            Call objCtb.Encerrando
            Set objCtb = Nothing
            '#################################################
            
        End If
    
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM TRVOcorrencias WHERE NumIntDoc= ? ", _
                    0, lNumIntDocAux, lNumIntDocOCR)
        If lErro <> AD_SQL_SUCESSO Then gError 192007
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192008
        
        If lErro = AD_SQL_SEM_DADOS Then gError 192018
    
        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> SUCESSO Then gError 192009

        lErro = Comando_ExecutarPos(alComando(3), "UPDATE TRVOcorrencias SET TipoDocDestino= ?, NumIntDocDestino=?, NumTitulo = ? ", alComando(2), _
        4, objTitPag.lNumIntDoc, lNumTitulo)
        If lErro <> AD_SQL_SUCESSO Then gError 192010
    
        'Busca a próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    MsgBox "Rotina finalizada com sucesso"
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TRV_Acerta_NVL2 = SUCESSO
     
    Exit Function
    
Erro_TRV_Acerta_NVL2:

    TRV_Acerta_NVL2 = gErr
     
    Select Case gErr
    
        Case 185496
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function


Public Function TRV_Acerta_NVL_Contab() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Refaz a contabilização em um determinado período

Dim lErro As Long
Dim lTransacaoExt As Long
Dim lTransacao As Long

Dim alComandoExt(0 To 4) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iFilialEmpresa As Integer
Dim objCtb As New ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As ClassContabAutomatica
Dim objContabil As ClassContabil
Dim objTitPag As New ClassTituloPagar

On Error GoTo Erro_TRV_Acerta_NVL_Contab
    
    'Abre o arquivo
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 185788
    Next
    
'    'Abertura de transação
'    lTransacaoExt = Transacao_AbrirExt(glConexao1)
'    If lTransacaoExt = 0 Then gError 185791

    'Lê os arquivos que não foram atualizados
    lErro = Comando_Executar(alComandoExt(0), "SELECT NumIntDoc, FilialEmpresa FROM TitulosPagTodos WHERE DataEmissao > {d '2008-02-01'} AND SiglaDocumento = 'FSCP' ", lNumIntDoc, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 185792
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185793
    
    'Enquanto existirem registros
    Do While lErro = AD_SQL_SUCESSO
        
        'Abertura de transação só para marcar que houve erro
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 185807
                
        Set objContabil = New ClassContabil
        Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTABILIDADE)
         
        If Not (objContabil Is Nothing) Then
                 
            'Exclui os lançamentos contábeis (se houverem) associados ao doc em questao
            lErro = objContabil.Contabil_Exclui(lNumIntDoc, CONTABIL_EXCLUI_EFETIVO, , , , iFilialEmpresa)
            If lErro <> SUCESSO Then gError 83949
    
            Set objContabil = Nothing
            
        End If
        
        Set objTitPag = New ClassTituloPagar
        Set objContabAutomatica = New ClassContabAutomatica
        objCtb.sModelo = "Reembolso"
        
        objTitPag.lNumIntDoc = lNumIntDoc
        
        lErro = CF("TituloPagar_Le1", objTitPag)
        If lErro <> SUCESSO And lErro <> 18372 Then gError 182533
        If lErro = SUCESSO Then
        
            'gravar a contabilizacao
            lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
            If lErro <> SUCESSO Then gError 190999
            
            lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
            If lErro <> SUCESSO Then gError 192000
            
        End If
    
        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 185809
            
        'Busca o próximo
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185793
            
    Loop
                
'    'Confirma a transação
'    lErro = Transacao_CommitExt(lTransacaoExt)
'    If lErro <> AD_SQL_SUCESSO Then gError 185814
    
    MsgBox "Rotina finalizada com sucesso"
        
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
                
    TRV_Acerta_NVL_Contab = SUCESSO
     
    Exit Function
    
Erro_TRV_Acerta_NVL_Contab:

    TRV_Acerta_NVL_Contab = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185815)
     
    End Select
    
    'Desfaz a transação
'    Call Transacao_RollbackExt(lTransacaoExt)
    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
        
    Exit Function

End Function

Private Function Coinfo_AtualizarCTBNFs(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza NFs
'percorrer os registros de fpnf e fpitemnf e marcá-los como lidos
'preenche objeto e grava
        
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim alComandoExt(0 To 4) As Long
Dim lTransacao As Long
Dim objNF As ClassNFiscal
Dim lNumIntDocNF As Long
Dim iFilialEmpresaAnt As Integer
Dim iFilialEmpAux As Integer
Dim objItemNF As ClassItemNF
Dim lNumIntDocTitulo As Long
Dim sSerie As String
Dim sTipoDoc As String
Dim lNumVou As Long
Dim dComissaoAcumulada As Double
Dim dComissao As Double
Dim objContabil As New ClassContabil
Dim objContabAutomatica As ClassContabAutomatica
Dim objCtb As ClassImpCoinfoNFCtb
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Coinfo_AtualizarCTBNFs

    iFilialEmpresaAnt = giFilialEmpresa

    'abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182460
    Next
    
    'abertura de comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 182460
    Next
    
    'Le as NFs sem contabilização
    lErro = Comando_Executar(alComandoExt(0), "SELECT NF.NumIntDoc, NF.FilialEmpresa FROM NFiscal AS NF WHERE NF.Cliente <> 0 AND NF.DataEmissao > {d '2008-02-01'} AND NOT EXISTS (SELECT L.NumIntDoc FROM Lancamentos AS L WHERE L.Transacao = 4 AND NF.NumIntDoc = L.NumIntDoc)", lNumIntDocNF, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 182461
    
    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182462
    
    'Enquanto houver arquivos não atualizados
    Do While lErro = AD_SQL_SUCESSO
    
        Set objNF = New ClassNFiscal
    
        objNF.lNumIntDoc = lNumIntDocNF
        
        iFilialEmpAux = iFilialEmpresa
        lErro = CF("NFiscal_Le", objNF)
        If lErro <> SUCESSO Then gError 99999
        
        objNF.iFilialEmpresa = iFilialEmpAux
        giFilialEmpresa = iFilialEmpAux
        
        'Lê os Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNF)
        If lErro <> SUCESSO Then gError 35705
    
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNF)
        If lErro <> SUCESSO Then gError 39342
    
        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNF)
        If lErro <> SUCESSO Then gError 39343
        
        dComissaoAcumulada = 0
        For Each objItemNF In objNF.ColItensNF
    
            'Busca os vouchers da fatura para poder calcular a comissão
            lErro = Comando_Executar(alComando(0), "SELECT X.NumIntDocTitulo FROM TRVItemNotaTitulo AS X WHERE  X.NumIntDocItemNF = ? ", lNumIntDocTitulo, objItemNF.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 192300
             
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192301
             
            If lErro <> AD_SQL_SUCESSO Then gError 192302
            
            sSerie = String(STRING_TRV_OCR_SERIE, 0)
            sTipoDoc = String(STRING_TRV_OCR_TIPODOC, 0)

            'Busca os vouchers da fatura para poder calcular a comissão
            lErro = Comando_Executar(alComando(2), "SELECT Numvou, Serie, TipoDoc FROM TRVVouchers WHERE TipoDocDestino = 3 AND NumIntDocDestino = ? ", lNumVou, sSerie, sTipoDoc, lNumIntDocTitulo)
            If lErro <> AD_SQL_SUCESSO Then gError 192292

            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192293

            Do While lErro <> AD_SQL_SEM_DADOS

                'Le a comissão da tabela CoinfoComissao (A comissão da tabela TRVVouchers é válida somente para faturas geradas pelo corporator)
                lErro = CF("Coinfo_Obtem_Comissao_Cliente2", lNumVou, sSerie, sTipoDoc, gsEmpresaTRV, dComissao)
                If lErro <> SUCESSO Then gError 192294

                dComissaoAcumulada = dComissaoAcumulada + dComissao

                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 192295

            Loop
        
        Next
        
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 182459
        
        Set objContabAutomatica = New ClassContabAutomatica
        Set objCtb = New ClassImpCoinfoNFCtb

        objNF.dTaxaMoeda = dComissaoAcumulada

        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_FATURAMENTO, "FAT", objNF.dtDataEmissao, objNF.dtDataEmissao)
        If lErro <> SUCESSO Then gError 192296

        lErro = objContabAutomatica.GeraContabilizacao(objNF)
        If lErro <> SUCESSO Then gError 192297

        lErro = objContabAutomatica.Finaliza_Contab()
        If lErro <> SUCESSO Then gError 192298

        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 182481
        
        'Busca o próximo arquivo não atualizado
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182480
        
    Loop
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    giFilialEmpresa = iFilialEmpresaAnt
    
    Call Rotina_Aviso(vbOKOnly, "Rotina Finalizada com sucesso")
    
    Coinfo_AtualizarCTBNFs = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarCTBNFs:

    giFilialEmpresa = iFilialEmpresaAnt
    
    Coinfo_AtualizarCTBNFs = gErr
     
    Select Case gErr
    
     
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182482)
     
    End Select
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Coinfo_AtualizaContabFaturas() As Long
'Gera a contabilizaçlão em cima das faturas geradas pela tela de faturamento

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim alComandoExt(0 To 1) As Long
Dim lNumIntDocTRVTitExp As Long
Dim iTipoDoc As Integer
Dim lNumIntDoc As Long
Dim lNumFat As Long
Dim dValorComissao As Double
Dim dValorTaxa   As Double, dDeducoes As Double
Dim lCountRegs As Long


On Error GoTo Erro_Coinfo_AtualizaContabFaturas

    'abertura de comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 196746
    Next
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196747
    Next

    'Busca os títulos não contabilizados
    lErro = Comando_Executar(alComandoExt(1), "SELECT COUNT(NumIntDoc) FROM TRVTitulosExp WHERE TemQueContabilizar = ? ", _
    lCountRegs, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 196758
     
    lErro = Comando_BuscarPrimeiro(alComandoExt(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196759
    
    lErro = Coinfo_Total_Registros2("Contabilizando Faturas", "6 de 15", lCountRegs)
    If lErro <> SUCESSO Then gError 196760
    
    'Busca os títulos não contabilizados
    lErro = Comando_Executar(alComandoExt(0), "SELECT NumIntDoc, TipoDocOrigem, NumIntDocOrigem, NumTitulo, ValorComissao, ValorTarifa, ValorDeducoes FROM TRVTitulosExp WHERE TemQueContabilizar = ? ", _
    lNumIntDocTRVTitExp, iTipoDoc, lNumIntDoc, lNumFat, dValorComissao, dValorTaxa, dDeducoes, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 196748
     
    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196749
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 196750
        
        Select Case iTipoDoc
        
            Case TRV_TIPO_DOC_DESTINO_TITPAG
            
                lErro = Coinfo_AtualizaContabFaturas_TitPag(lNumIntDoc, lNumFat)
            
            Case TRV_TIPO_DOC_DESTINO_TITREC
            
'                'Busca os títulos não contabilizados
'                lErro = Comando_Executar(alComando(1), "SELECT SUM(ValorComissao), SUM(Valor) FROM TRVVouchers WHERE TipoDocDestino = ? AND NumIntDocDestino = ? ", _
'                dValorComissao, dValorVoucher, iTipoDoc, lNumIntDoc)
'                If lErro <> AD_SQL_SUCESSO Then gError 196751
'
'                lErro = Comando_BuscarPrimeiro(alComando(1))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196752
            
                lErro = Coinfo_AtualizaContabFaturas_TitRec(lNumIntDoc, lNumFat, dValorComissao, dValorTaxa, dDeducoes)
            
            Case TRV_TIPO_DOC_DESTINO_NFSPAG
        
                lErro = Coinfo_AtualizaContabFaturas_NFPag(lNumIntDoc, lNumFat)
                
            Case Else
            
                gError 196753
                
        End Select
        If lErro <> SUCESSO Then gError 196754
        
        lErro = Comando_Executar(alComando(0), "UPDATE TRVTitulosExp SET TemQueContabilizar = ? WHERE NumIntDoc = ? ", DESMARCADO, lNumIntDocTRVTitExp)
        If lErro <> AD_SQL_SUCESSO Then gError 196755
        
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 196756
        
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 196761
    
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196757
    
    Loop
       
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    Coinfo_AtualizaContabFaturas = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaContabFaturas:

    Coinfo_AtualizaContabFaturas = gErr
     
    Select Case gErr
    
        Case 196746, 196747
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196748, 196749, 196757, 196758, 196759
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVTITULOSEXP", gErr)
        
        Case 196750
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 196751, 196752
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 196753
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_DESTINO_NAO_CADASTRADO", gErr, iTipoDoc)
            
        Case 196754, 196760, 196761
        
        Case 196755
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVTITULOSEXP", gErr)
        
        Case 196756
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196758)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Coinfo_AtualizaContabFaturas_TitRec(ByVal lNumIntDoc As Long, ByVal lNumFat As Long, ByVal dValorComissao As Double, ByVal dValorTaxa As Double, ByVal dDeducaoValorPresente As Double) As Long

Dim lErro As Long
Dim objTitRec As New ClassTituloReceber
Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoFatCRCtb
Dim objContabAutomatica As New ClassContabAutomatica
Dim objImpCoinfoCartoes As New ClassImpCoinfoCartoes

On Error GoTo Erro_Coinfo_AtualizaContabFaturas_TitRec
    
    objTitRec.lNumIntDoc = lNumIntDoc
    
    'Le o título
    lErro = CF("TituloReceber_Le", objTitRec)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 196730
    If lErro <> SUCESSO Then
       'Lê o título baixado
       lErro = CF("TituloReceberBaixado_Le", objTitRec)
       If lErro <> SUCESSO And lErro <> 56570 Then gError 196731
    End If
    
    'Se o título não existe (foi excluido), não é preciso fazer nada
    If lErro = SUCESSO Then
    
        If objTitRec.lNumTitulo <> lNumFat Then gError 196732
    
        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASARECEBER, "CR", objTitRec.dtDataEmissao, objTitRec.dtDataEmissao)
        If lErro <> SUCESSO Then gError 196733
    
        If objTitRec.sSiglaDocumento = TIPODOC_CARTAO_CRED_DEB Then
            objImpCoinfoCartoes.dValorTaxa = dValorTaxa
            objImpCoinfoCartoes.dDeducaoValorPresente = dDeducaoValorPresente
            lErro = objContabAutomatica.GeraContabilizacao(objTitRec, objImpCoinfoCartoes)
        Else
            'Coloca a comissão no CSLL
            objTitRec.dCSLLRetido = dValorComissao
            lErro = objContabAutomatica.GeraContabilizacao(objTitRec, Nothing)
        End If
        If lErro <> SUCESSO Then gError 196734
    
        lErro = objContabAutomatica.Finaliza_Contab()
        If lErro <> SUCESSO Then gError 196735
    
        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
        
    End If
    
    Coinfo_AtualizaContabFaturas_TitRec = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaContabFaturas_TitRec:

    Coinfo_AtualizaContabFaturas_TitRec = gErr
     
    Select Case gErr
    
        Case 196730, 196731, 196733, 196734, 196735
        
        Case 196732
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DO_TITULO_DIFERENTE", gErr, objTitRec.lNumTitulo, lNumFat)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196736)
     
    End Select
       
    Exit Function

End Function

Private Function Coinfo_AtualizaContabFaturas_TitPag(ByVal lNumIntDoc As Long, ByVal lNumFat As Long) As Long

Dim lErro As Long
Dim objTitPag As New ClassTituloPagar
Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoFatCPCtb
Dim objContabAutomatica As New ClassContabAutomatica

On Error GoTo Erro_Coinfo_AtualizaContabFaturas_TitPag
    
    objTitPag.lNumIntDoc = lNumIntDoc
    
    lErro = CF("TituloPagar_Le_Todos", objTitPag, True)
    If lErro <> SUCESSO Then gError 196737
    
    If objTitPag.iStatus <> STATUS_EXCLUIDO Then
    
        If objTitPag.lNumTitulo <> lNumFat Then gError 196741
        
'        Set objContabil = New ClassContabil
'        Call objContabil.Contabil_Inicializa_Contabilidade4(1, MODULO_CONTASAPAGAR)
'
'        If Not (objContabil Is Nothing) Then
'
'            'Exclui os lançamentos contábeis (se houverem) associados ao doc em questao
'            lErro = objContabil.Contabil_Exclui(lNumIntDoc, CONTABIL_EXCLUI_EFETIVO, , , , objTitPag.iFilialEmpresa)
'            If lErro <> SUCESSO Then gError 83949
'
'            Set objContabil = Nothing
'        End If
    
        If objTitPag.sSiglaDocumento = "OVER" Then
            objCtb.sModelo = "Comissão Over"
        Else
            objCtb.sModelo = "Reembolso"
        End If
    
        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objTitPag.dtDataEmissao, objTitPag.dtDataEmissao)
        If lErro <> SUCESSO Then gError 196738
    
        lErro = objContabAutomatica.GeraContabilizacao(objTitPag)
        If lErro <> SUCESSO Then gError 196739
    
        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
    
    End If

    Coinfo_AtualizaContabFaturas_TitPag = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaContabFaturas_TitPag:

    Coinfo_AtualizaContabFaturas_TitPag = gErr
     
    Select Case gErr
    
        Case 196737 To 196739
    
        Case 196741
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DO_TITULO_DIFERENTE", gErr, objTitPag.lNumTitulo, lNumFat)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196740)
     
    End Select
       
    Exit Function

End Function

Private Function Coinfo_AtualizaContabFaturas_NFPag(ByVal lNumIntDoc As Long, ByVal lNumFat As Long) As Long

Dim lErro As Long
Dim objNFPag As New ClassNFsPag
Dim objContabil As ClassContabil
Dim objCtb As New ClassImpCoinfoNFPagCtb
Dim objContabAutomatica As New ClassContabAutomatica

On Error GoTo Erro_Coinfo_AtualizaContabFaturas_NFPag

    objNFPag.lNumIntDoc = lNumIntDoc

    'Lê o Título
    lErro = CF("NFPag_Le_Todos2", objNFPag, MARCADO)
    If lErro <> SUCESSO And lErro <> 18257 Then gError 196742
    
    If lErro = SUCESSO Then
    
        If objNFPag.lNumNotaFiscal <> lNumFat Then gError 196745
                
        'gravar a contabilizacao
        lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_CONTASAPAGAR, "CP", objNFPag.dtDataEmissao, objNFPag.dtDataEmissao)
        If lErro <> SUCESSO Then gError 196743
    
        lErro = objContabAutomatica.GeraContabilizacao(objNFPag, False)
        If lErro <> SUCESSO Then gError 196744
    
        Set objContabAutomatica = Nothing
        Call objCtb.Encerrando
        Set objCtb = Nothing
        
    End If

    Coinfo_AtualizaContabFaturas_NFPag = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaContabFaturas_NFPag:

    Coinfo_AtualizaContabFaturas_NFPag = gErr
     
    Select Case gErr
    
        Case 196742 To 196744
        
        Case 196745
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DO_TITULO_DIFERENTE", gErr, objNFPag.lNumNotaFiscal, lNumFat)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196740)
     
    End Select
       
    Exit Function

End Function

Private Function Coinfo_AtualizaContabNFs() As Long
'Gera a contabilizaçlão em cima das faturas geradas pela tela de faturamento

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim alComandoExt(0 To 1) As Long
Dim lNumIntDocTRVNFExp As Long
Dim iTipoDoc As Integer
Dim lNumIntDoc As Long
Dim lNumNF As Long
Dim dValorComissao As Double
Dim lCountRegs As Long
Dim iFilialEmpresaAnt As Integer
Dim objNF As New ClassNFiscal
Dim objContabAutomatica As ClassContabAutomatica
Dim objCtb As ClassImpCoinfoNFCtb

On Error GoTo Erro_Coinfo_AtualizaContabNFs

    iFilialEmpresaAnt = giFilialEmpresa

    'abertura de comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        alComandoExt(iIndice) = Comando_AbrirExt(glConexao1)
        If alComandoExt(iIndice) = 0 Then gError 196867
    Next
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 196868
    Next

    'Busca ass nfs não contabilizadas
    lErro = Comando_Executar(alComandoExt(1), "SELECT COUNT(NumIntDoc) FROM TRVNFsExp WHERE TemQueContabilizar = ? ", _
    lCountRegs, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 196869
     
    lErro = Comando_BuscarPrimeiro(alComandoExt(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196870
    
    lErro = Coinfo_Total_Registros2("Contabilizando NFs", "7 de 15", lCountRegs)
    If lErro <> SUCESSO Then gError 196871
    
    'Busca os títulos não contabilizados
    lErro = Comando_Executar(alComandoExt(0), "SELECT NumIntDoc, NumIntDocNF, NumNF, ValorComissao FROM TRVNFsExp WHERE TemQueContabilizar = ? ORDER BY NumIntDocNF ", _
    lNumIntDocTRVNFExp, lNumIntDoc, lNumNF, dValorComissao, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 196872
     
    lErro = Comando_BuscarPrimeiro(alComandoExt(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196873
    
    Do While lErro <> AD_SQL_SEM_DADOS
           
        Set objNF = New ClassNFiscal
    
        objNF.lNumIntDoc = lNumIntDoc
        
        lErro = CF("NFiscal_Le", objNF)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 196875
        
        If lErro = SUCESSO And objNF.iStatus <> STATUS_CANCELADO Then
        
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 196874
            
            giFilialEmpresa = objNF.iFilialEmpresa
            
            'Lê os Ítens da Nota Fiscal
            lErro = CF("NFiscalItens_Le", objNF)
            If lErro <> SUCESSO Then gError 196876
        
            'Lê a parte de Tributação
            lErro = CF("NFiscal_Le_Tributacao", objNF)
            If lErro <> SUCESSO Then gError 196877
        
            'Lê a parte de tributacao dos Ítens da Nota Fiscal
            lErro = CF("NFiscalItens_Le_Trib", objNF)
            If lErro <> SUCESSO Then gError 196878
            
            Set objContabAutomatica = New ClassContabAutomatica
            Set objCtb = New ClassImpCoinfoNFCtb
    
            objNF.dTaxaMoeda = dValorComissao

            'gravar a contabilizacao
            lErro = objContabAutomatica.Inicializa_Contab(objCtb, MODULO_FATURAMENTO, "FAT", objNF.dtDataEmissao, objNF.dtDataEmissao)
            If lErro <> SUCESSO Then gError 196879
    
            lErro = objContabAutomatica.GeraContabilizacao(objNF)
            If lErro <> SUCESSO Then gError 196880
    
            lErro = objContabAutomatica.Finaliza_Contab()
            If lErro <> SUCESSO Then gError 196881
    
            Set objContabAutomatica = Nothing
            Call objCtb.Encerrando
            Set objCtb = Nothing
            
            lErro = Comando_Executar(alComando(0), "UPDATE TRVNFsExp SET TemQueContabilizar = ? WHERE NumIntDoc = ? ", DESMARCADO, lNumIntDocTRVNFExp)
            If lErro <> AD_SQL_SUCESSO Then gError 196882
            
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 196883
            
        ElseIf lErro = SUCESSO And objNF.iStatus = STATUS_CANCELADO Then
            
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 196874
            
            lErro = Comando_Executar(alComando(0), "UPDATE TRVNFsExp SET TemQueContabilizar = ? WHERE NumIntDoc = ? ", DESMARCADO, lNumIntDocTRVNFExp)
            If lErro <> AD_SQL_SUCESSO Then gError 196882
            
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 196883
            
        End If
        
        lErro = Rotina_Atualizar_AtualizaTelaBatch
        If lErro <> SUCESSO Then gError 196884
    
        lErro = Comando_BuscarProximo(alComandoExt(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196885
    
    Loop
       
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    giFilialEmpresa = iFilialEmpresaAnt

    
    Coinfo_AtualizaContabNFs = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizaContabNFs:

    giFilialEmpresa = iFilialEmpresaAnt

    Coinfo_AtualizaContabNFs = gErr
     
    Select Case gErr
    
        Case 196867, 196868
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196869, 196870, 196872, 196873, 196885
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVNFSEXP", gErr)
            
        Case 196871, 196875 To 196881, 196884
        
        Case 196874
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 196882
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TRVNFSEXP", gErr)
        
        Case 196883
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196886)
     
    End Select
    
    Call Rotina_Erro(vbOKOnly, "Não foi possível contabilizar o registro da tabela TRVNFsExp com ID " & CStr(lNumIntDocTRVNFExp), gErr)
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha o comando
    For iIndice = LBound(alComandoExt) To UBound(alComandoExt)
        Call Comando_Fechar(alComandoExt(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Coinfo_ImportaComissao_010808(ByVal objImpCoinfo As ClassImpCoinfo, Optional ByVal iAno As Integer, Optional ByVal iMes As Integer) As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os créditos a pagar do tipo NCP

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 27) As Long
Dim tVou As typeTRVVouchers
Dim objVou As ClassTRVVouchers
Dim lNumIntDocNFPag As Long

On Error GoTo Erro_Coinfo_ImportaComissao_010808

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459
    
'=============================>
    '*************************
    'VALOR BRUTO DE VOUCHERS NÃO CANCELADOS
    '*************************
    tVou.sSerie = String(255, 0)
    tVou.sTipVou = String(255, 0)
    
    If iAno <> 0 Then
        lErro = Comando_Executar(alComando(0), "SELECT V1.Cartao, V1.ValorComissao, V1.NumVou, V1.TipVou, V1.Serie, V1.Data, V1.TipoDocDestino, V1.NumIntDocDestino, V2.TarifaValorReal, V1.NumFatCoinfo, V2.Cliente " & _
                                                "FROM TRVVouchers AS V1, TRVVoucherInfoSigav AS V2 WHERE V1.Status <> 7 AND V1.TipoDoc = 'VOU' AND V1.Serie = V2.Serie AND V1.TipVou = V2.Tipo AND V1.NumVou = V2.NumVou AND YEAR(V1.Data) = ? AND MONTH(V1.Data) = ?", _
            tVou.iCartao, tVou.dValorComissao, tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.iTipoDocDestino, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lNumFatCoinfo, tVou.lCliente, iAno, iMes)
    Else
        lErro = Comando_Executar(alComando(0), "SELECT V1.Cartao, V1.ValorComissao, V1.NumVou, V1.TipVou, V1.Serie, V1.Data, V1.TipoDocDestino, V1.NumIntDocDestino, V2.TarifaValorReal, V1.NumFatCoinfo, V2.Cliente " & _
                                                "FROM TRVVouchers AS V1, TRVVoucherInfoSigav AS V2 WHERE V1.Status <> 7 AND V1.TipoDoc = 'VOU' AND V1.Serie = V2.Serie AND V1.TipVou = V2.Tipo AND V1.NumVou = V2.NumVou AND V1.Data >= {d '2007-01-01'}", _
            tVou.iCartao, tVou.dValorComissao, tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.iTipoDocDestino, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lNumFatCoinfo, tVou.lCliente)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRVVouchers
    
        objVou.sSerie = tVou.sSerie
        objVou.sTipVou = tVou.sTipVou
        objVou.lNumVou = tVou.lNumVou
        
        lNumIntDocNFPag = 0
        
        If tVou.iCartao = MARCADO Then
        
            'Le as comissões associadas a esse item de fatura que não foram geradas
            lErro = Comando_Executar(alComando(21), "SELECT C1.ValComi, C1.NumIntDocCorporator FROM CoinfoComissao AS C1 WHERE C1.NumVou = ? AND C1.Serie = ? AND C1.TipVou = ? AND C1.CodEmp <> 0 ORDER BY C1.NumIntDocCorporator DESC, C1.NumIntArq DESC, C1.NumRegArqTexto DESC", _
            tVou.dValorComissao, lNumIntDocNFPag, objVou.lNumVou, objVou.sSerie, TRV_TIPODOC_VOU_TEXTO)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
            
            'Busca a primeira
            lErro = Comando_BuscarProximo(alComando(21))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
        
        End If
        
        lErro = TRVVou_Gera_Comis_Insere_Info(alComando(1), alComando(2), alComando(3), objVou, TRV_TIPODOC_BRUTO_TEXTO, tVou.dtData, tVou.dValorBruto, "Importação do Sigav", TRV_VOU_INFO_STATUS_BLOQUEADO, 0, 0, 0, TRV_CLIENTEINFO_TIPO_CLIENTE, tVou.lCliente, TRV_TIPO_ESTORNO_LANC_NORMAL, 0, 1, 0, 0, 0)
        If lErro <> SUCESSO Then gError 99999
        
        '#####################################################
        'VALOR CMA OU CMCC
        lErro = TRVVou_Gera_Comis_Insere_Info(alComando(1), alComando(2), alComando(3), objVou, IIf(tVou.iCartao = MARCADO, TRV_TIPODOC_CMCC_TEXTO, TRV_TIPODOC_CMA_TEXTO), tVou.dtData, tVou.dValorComissao, "Importação do Sigav", IIf(tVou.iCartao = MARCADO, TRV_VOU_INFO_STATUS_ANTIGA, TRV_VOU_INFO_STATUS_BLOQUEADO), 0, 0, 0, TRV_CLIENTEINFO_TIPO_CLIENTE, tVou.lCliente, TRV_TIPO_ESTORNO_LANC_NORMAL, 0, 2, IIf(lNumIntDocNFPag <> 0, 5, 0), lNumIntDocNFPag, IIf(lNumIntDocNFPag <> 0, tVou.lNumFatCoinfo, 0))
        If lErro <> SUCESSO Then gError 99999
        '#####################################################
        
        If tVou.dValorBruto <> 0 Then
            lErro = Comando_Executar(alComando(22), "UPDATE TRVVouchers SET ClienteComissao = ClienteVou ,ComissaoAg = ?  WHERE NumVou= ?  AND TipoDoc= ?  AND Serie= ?  AND TipVou= ? AND ValorBruto <> 0 ", _
                tVou.dValorComissao / tVou.dValorBruto, tVou.lNumVou, TRV_TIPODOC_VOU_TEXTO, tVou.sSerie, tVou.sTipVou)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
        End If
    
        'Busca a primeira
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    Loop
'=============================>
        
'=============================>
    '*************************
    'VALOR CMR e CMC
    '*************************
    tVou.sSerie = String(255, 0)
    tVou.sTipVou = String(255, 0)
    tVou.sTipoDoc = String(255, 0)
    
    If iAno = 0 Then
        lErro = Comando_Executar(alComando(4), "SELECT V1.TipoDoc, V1.Numvou , V1.TipVou, V1.Serie, V1.Data, V1.TipoDocDestino, V1.NumIntDocDestino, abs(V1.Valor) ,v1.Numfatcoinfo,v1.cliente " & _
                                                "FROM TRVVouchers AS V1 WHERE V1.TipoDoc IN ('CMR', 'CMC') AND (V1.Data >= {d '2007-01-01'} OR V1.NumFatCoinfo = 0) AND (V1.Status <>7 OR V1.NumfatCoinfo <> 0) and V1.Valor <> 0 ", _
            tVou.sTipoDoc, tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.iTipoDocDestino, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lNumFatCoinfo, tVou.lCliente)
    Else
        lErro = Comando_Executar(alComando(4), "SELECT V1.TipoDoc, V1.Numvou , V1.TipVou, V1.Serie, V1.Data, V1.TipoDocDestino, V1.NumIntDocDestino, abs(V1.Valor) ,v1.Numfatcoinfo,v1.cliente " & _
                                                "FROM TRVVouchers AS V1 WHERE V1.TipoDoc IN ('CMR', 'CMC') AND ((YEAR(V1.Data) = ? AND MONTH(V1.Data) = ?)) AND (V1.Status <>7 OR V1.NumfatCoinfo <> 0) and V1.Valor <> 0 ", _
            tVou.sTipoDoc, tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.iTipoDocDestino, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lNumFatCoinfo, tVou.lCliente, iAno, iMes)
    End If

    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRVVouchers
    
        objVou.sSerie = tVou.sSerie
        objVou.sTipVou = tVou.sTipVou
        objVou.lNumVou = tVou.lNumVou

        lErro = TRVVou_Gera_Comis_Insere_Info(alComando(1), alComando(2), alComando(3), objVou, tVou.sTipoDoc, tVou.dtData, tVou.dValorBruto, "Importação do Sigav", TRV_VOU_INFO_STATUS_LIBERADO, 0, 0, 0, TRV_CLIENTEINFO_TIPO_CLIENTE, tVou.lCliente, TRV_TIPO_ESTORNO_LANC_NORMAL, 0, IIf(tVou.sTipoDoc = TRV_TIPODOC_CMR_TEXTO, 3, 4), tVou.iTipoDocDestino, tVou.lNumIntDocDestino, tVou.lNumFatCoinfo)
        If lErro <> SUCESSO Then gError 99999

        If tVou.sTipoDoc = TRV_TIPODOC_CMC_TEXTO Then
            lErro = Comando_Executar(alComando(23), "UPDATE TRVVouchers SET Correntista = ? ,ComissaoCorr = (? / ValorBruto) WHERE NumVou= ?  AND TipoDoc= ?  AND Serie= ?  AND TipVou= ? AND ValorBruto <> 0 ", _
                tVou.lCliente, tVou.dValorBruto, tVou.lNumVou, TRV_TIPODOC_VOU_TEXTO, tVou.sSerie, tVou.sTipVou)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
        Else
            lErro = Comando_Executar(alComando(24), "UPDATE TRVVouchers SET Representante = ? ,ComissaoRep = (? / ValorBruto) WHERE NumVou= ?  AND TipoDoc= ?  AND Serie= ?  AND TipVou= ? AND ValorBruto <> 0 ", _
                tVou.lCliente, tVou.dValorBruto, tVou.lNumVou, TRV_TIPODOC_VOU_TEXTO, tVou.sSerie, tVou.sTipVou)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
        End If

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
'=============================>

'=============================>
    '*************************
    'VALOR OVER
    '*************************
    tVou.sSerie = String(255, 0)
    tVou.sTipVou = String(255, 0)
    
    If iAno = 0 Then
        lErro = Comando_Executar(alComando(5), "SELECT O.NumVou, O.TipVou, O.Serie, O.DatEmi,  O.NumIntDocCorporator, O.valor, O.CodEmi AS Cliente, T.NumTitulo  " & _
                                                "FROM  OverUltimo AS O LEFT OUTER JOIN TitulosPagTodos AS T ON O.NumIntDocCorporator = T.NumIntDoc WHERE O.DatEmi >= {d '2007-01-01'} AND (O.valor > 0) AND (O.codest = ?) ", _
            tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lCliente, tVou.lNumFatCoinfo, objImpCoinfo.sSiglaEmpresa)
    Else
        lErro = Comando_Executar(alComando(5), "SELECT O.NumVou, O.TipVou, O.Serie, O.DatEmi,  O.NumIntDocCorporator, O.valor, O.CodEmi AS Cliente, T.NumTitulo  " & _
                                                "FROM  OverUltimo AS O LEFT OUTER JOIN TitulosPagTodos AS T ON O.NumIntDocCorporator = T.NumIntDoc WHERE (O.valor > 0) AND (O.codest = ?) AND YEAR(O.DatEmi) = ? AND MONTH(O.DatEmi) = ? ", _
            tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lCliente, tVou.lNumFatCoinfo, objImpCoinfo.sSiglaEmpresa, iAno, iMes)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRVVouchers
        
        tVou.lNumIntDocDestino = 4
        tVou.lCliente = tVou.lCliente + FATOR_SOMA_COD_EMISSOR
    
        objVou.sSerie = tVou.sSerie
        objVou.sTipVou = tVou.sTipVou
        objVou.lNumVou = tVou.lNumVou

        lErro = TRVVou_Gera_Comis_Insere_Info(alComando(1), alComando(2), alComando(3), objVou, TRV_TIPODOC_OVER_TEXTO, tVou.dtData, tVou.dValorBruto, "Importação do Sigav", TRV_VOU_INFO_STATUS_ANTIGA, 0, 0, 0, TRV_CLIENTEINFO_TIPO_FORNECEDOR, tVou.lCliente, TRV_TIPO_ESTORNO_LANC_NORMAL, 0, 5, IIf(tVou.lNumIntDocDestino <> 0, 4, 0), tVou.lNumIntDocDestino, tVou.lNumFatCoinfo)
        If lErro <> SUCESSO Then gError 99999
        
        lErro = Comando_Executar(alComando(25), "UPDATE TRVVouchers SET ComissaoEmissor = (? / ValorBruto) WHERE NumVou= ?  AND TipoDoc= ?  AND Serie= ?  AND TipVou= ? AND ValorBruto <> 0 ", _
            tVou.dValorBruto, tVou.lNumVou, TRV_TIPODOC_VOU_TEXTO, tVou.sSerie, tVou.sTipVou)
        If lErro <> AD_SQL_SUCESSO Then gError 182534

        'obter dados dos itens da fatura correspondentes para acumular dados de cartao
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182542
    
    Loop
'=============================>
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ImportaComissao_010808 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportaComissao_010808:

    Coinfo_ImportaComissao_010808 = gErr
     
    Select Case gErr

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function Coinfo_ImportaComissao_150908(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os créditos a pagar do tipo NCP

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 27) As Long
Dim tVou As typeTRVVouchers
Dim objVou As ClassTRVVouchers
Dim lNumIntDocNFPag As Long

On Error GoTo Erro_Coinfo_ImportaComissao_150908

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182534
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182459
    
'=============================>
    '*************************
    'VALOR BRUTO DE VOUCHERS NÃO CANCELADOS
    '*************************
    tVou.sSerie = String(255, 0)
    tVou.sTipVou = String(255, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT V1.Cartao, V1.ValorComissao, V1.NumVou, V1.TipVou, V1.Serie, V1.Data, V1.TipoDocDestino, V1.NumIntDocDestino, V1.ValorBruto, V1.NumFatCoinfo, V2.Cliente " & _
                                            "FROM TRVVouchers AS V1, TRVVoucherInfoSigav AS V2 WHERE V1.Status <> 7 AND V1.TipoDoc = 'VOU' AND V1.Serie = V2.Serie AND V1.TipVou = V2.Tipo AND V1.NumVou = V2.NumVou AND V1.Data >= {d '2008-09-15'} and not exists (select * from trvvoucherinfo as i where v1.numvou = i.numvou and v1.serie=i.serie and v1.tipvou = i.tipvou) and v1.valorbruto <> 0", _
        tVou.iCartao, tVou.dValorComissao, tVou.lNumVou, tVou.sTipVou, tVou.sSerie, tVou.dtData, tVou.iTipoDocDestino, tVou.lNumIntDocDestino, tVou.dValorBruto, tVou.lNumFatCoinfo, tVou.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 182534
    
    'Busca a primeira
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    'Enquanto houver faturas
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVou = New ClassTRVVouchers
    
        objVou.sSerie = tVou.sSerie
        objVou.sTipVou = tVou.sTipVou
        objVou.lNumVou = tVou.lNumVou
        
        lNumIntDocNFPag = 0
        
        If tVou.iCartao = MARCADO Then
        
            'Le as comissões associadas a esse item de fatura que não foram geradas
            lErro = Comando_Executar(alComando(21), "SELECT C1.ValComi, C1.NumIntDocCorporator FROM CoinfoComissao AS C1 WHERE C1.NumVou = ? AND C1.Serie = ? AND C1.TipVou = ? AND C1.CodEmp <> 0 ORDER BY C1.NumIntDocCorporator DESC, C1.NumIntArq DESC, C1.NumRegArqTexto DESC", _
            tVou.dValorComissao, lNumIntDocNFPag, objVou.lNumVou, objVou.sSerie, TRV_TIPODOC_VOU_TEXTO)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
            
            'Busca a primeira
            lErro = Comando_BuscarProximo(alComando(21))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
        
        End If
        
        lErro = TRVVou_Gera_Comis_Insere_Info(alComando(1), alComando(2), alComando(3), objVou, TRV_TIPODOC_BRUTO_TEXTO, tVou.dtData, tVou.dValorBruto, "Importação do Sigav", TRV_VOU_INFO_STATUS_BLOQUEADO, 0, 0, 0, TRV_CLIENTEINFO_TIPO_CLIENTE, tVou.lCliente, TRV_TIPO_ESTORNO_LANC_NORMAL, 0, 1, 0, 0, 0)
        If lErro <> SUCESSO Then gError 99999
        
        '#####################################################
        'VALOR CMA OU CMCC
        lErro = TRVVou_Gera_Comis_Insere_Info(alComando(1), alComando(2), alComando(3), objVou, IIf(tVou.iCartao = MARCADO, TRV_TIPODOC_CMCC_TEXTO, TRV_TIPODOC_CMA_TEXTO), tVou.dtData, tVou.dValorComissao, "Importação do Sigav", IIf(tVou.iCartao = MARCADO, TRV_VOU_INFO_STATUS_ANTIGA, TRV_VOU_INFO_STATUS_BLOQUEADO), 0, 0, 0, TRV_CLIENTEINFO_TIPO_CLIENTE, tVou.lCliente, TRV_TIPO_ESTORNO_LANC_NORMAL, 0, 2, IIf(lNumIntDocNFPag <> 0, 5, 0), lNumIntDocNFPag, IIf(lNumIntDocNFPag <> 0, tVou.lNumFatCoinfo, 0))
        If lErro <> SUCESSO Then gError 99999
        '#####################################################
        
        If tVou.dValorBruto <> 0 Then
            lErro = Comando_Executar(alComando(22), "UPDATE TRVVouchers SET ClienteComissao = ClienteVou ,ComissaoAg = ?  WHERE NumVou= ?  AND TipoDoc= ?  AND Serie= ?  AND TipVou= ? AND ValorBruto <> 0 ", _
                tVou.dValorComissao / tVou.dValorBruto, tVou.lNumVou, TRV_TIPODOC_VOU_TEXTO, tVou.sSerie, tVou.sTipVou)
            If lErro <> AD_SQL_SUCESSO Then gError 182534
        End If
    
        'Busca a primeira
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182533
    
    Loop
'=============================>
             
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182481
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_ImportaComissao_150908 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportaComissao_150908:

    Coinfo_ImportaComissao_150908 = gErr
     
    Select Case gErr

        Case 182538
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 182541
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182505)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function TRVVou_Gera_Comis_Insere_Info(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal objVou As ClassTRVVouchers, ByVal sTipoDoc As String, ByVal dtData As Date, ByVal dValor As Double, ByVal sHistorico As String, ByVal iStatus As Integer, ByVal iTipoLiberacao As Integer, ByVal lNumIntDocLiberacao As Long, ByVal lNumIntDocOCR As Long, ByVal iTipoCliForn As Integer, ByVal lCliForn As Long, ByVal iEstorno As Integer, ByVal lNumIntDocEstorno As Long, lContador As Long, ByVal iTipoDocDestino As Integer, ByVal lNumIntDocDestino As Long, ByVal lNumTitulo As Long) As Long
'Gera a comissão externa relacionada ao voucher

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_TRVVou_Gera_Comis_Insere_Info

    If Abs(dValor) > DELTA_VALORMONETARIO Then

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt2", lComando1, lComando2, "TRVConfig", "NUM_INT_PROX_TRVVOURCHERINFO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 198153
        
        'Insere em TRVVoucherInfo
        lErro = Comando_Executar(lComando3, "INSERT INTO TRVVoucherInfo (NumIntDoc,TipoDoc,TipVou,Serie,NumVou,Seq,Data,TipoDocDestino,NumIntDocDestino,NumIntDocComiInt,Valor,Historico,NumTitulo, Status,TipoLiberacao,NumIntDocLiberacao, NumIntDocOcr,TipoCliForn,CliForn, Estorno, NumIntDocEstorno, DataRegistro, HoraRegistro, Usuario) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
                     lNumIntDoc, sTipoDoc, objVou.sTipVou, objVou.sSerie, objVou.lNumVou, lContador, dtData, iTipoDocDestino, lNumIntDocDestino, 0, Arredonda_Moeda(dValor), sHistorico, lNumTitulo, iStatus, iTipoLiberacao, lNumIntDocLiberacao, lNumIntDocOCR, iTipoCliForn, lCliForn, iEstorno, lNumIntDocEstorno, Date, CDbl(Time), gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 198154
           
    End If
                       
    TRVVou_Gera_Comis_Insere_Info = SUCESSO

    Exit Function

Erro_TRVVou_Gera_Comis_Insere_Info:

    TRVVou_Gera_Comis_Insere_Info = gErr

    Select Case gErr

        Case 198153
            
        Case 198154
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRVVOUCHERINFO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198155)

    End Select

    Exit Function

End Function

Private Function Simula_Extracao_Vouchers_Sigav()

Dim lErro As Long
Dim objTRVVoucherInfo As ClassTRVVoucherInfo
Dim objTRVVoucher As ClassTRVVouchers
Dim colVoucher As New Collection
Dim vbResult As VbMsgBoxResult
Dim lComando As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim lConexaoTVA As Long
Dim objEmpresa As New ClassDicEmpresa

On Error GoTo Erro_Simula_Extracao_Vouchers_Sigav

    objEmpresa.lCodigo = 1 'TVA'

    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError 192863

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoTVA = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoTVA = 0 Then gError 192863

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 200170

    lErro = CF("TRVVouchers_Le_Periodo", DATA_NULA, DATA_NULA, colVoucher, True, iSistemaIntegrado)
    If lErro <> SUCESSO Then gError 192863

    lErro = Coinfo_Total_Registros2("Extraindo dados da Kogut", "13 de 15", colVoucher.Count)
    If lErro <> SUCESSO Then gError 192863

    For Each objTRVVoucher In colVoucher
    
        Set objTRVVoucherInfo = New ClassTRVVoucherInfo
        
        objTRVVoucherInfo.sTipo = objTRVVoucher.sTipVou
        objTRVVoucherInfo.lNumVou = objTRVVoucher.lNumVou
        objTRVVoucherInfo.sSerie = objTRVVoucher.sSerie
            
        lErro = Obter_Dados_Kogut(lComando, objTRVVoucher, objTRVVoucherInfo)
        If lErro = SUCESSO Then
        
            lErro = CF("TRVVoucherInfoSigav_Grava", objTRVVoucherInfo, lConexaoTVA)
            'If lErro <> SUCESSO Then gError 192865
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 192863
        
        End If
        
    Next
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    If lConexaoTVA <> 0 Then Call Conexao_FecharExt(lConexaoTVA)
    
    Simula_Extracao_Vouchers_Sigav = SUCESSO

    Exit Function

Erro_Simula_Extracao_Vouchers_Sigav:

    Simula_Extracao_Vouchers_Sigav = gErr

    Select Case gErr
            
        Case 192863 To 192865

        Case 200170
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192866)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    If lConexaoTVA <> 0 Then Call Conexao_FecharExt(lConexaoTVA)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Private Function Obter_Dados_Kogut(ByVal lComando As Long, ByVal objVou As ClassTRVVouchers, ByVal objVouInfo As ClassTRVVoucherInfo) As Long

Dim lErro As Long
Dim tIF As typeImpCoinfoItemFat
Dim sCodPro As String

On Error GoTo Erro_Obter_Dados_Kogut

    'Alocação de espaço no buffer
    tIF.sMoeda = String(255, 0)
    tIF.scodest = String(255, 0)
    tIF.sUsuario = String(255, 0)
    tIF.sTipoDoc = String(255, 0)
    tIF.sSerie = String(255, 0)
    tIF.sTipVou = String(255, 0)
    tIF.scodgru = String(255, 0)
    tIF.sControle = String(255, 0)
    tIF.sCodPro = String(255, 0)
    tIF.sNumCCred = String(255, 0)
    tIF.sValidCC = String(255, 0)
    tIF.sCiaCart = String(255, 0)
    tIF.sGrupo = String(255, 0)
    tIF.sDestino = String(255, 0)
    tIF.sIdioma = String(255, 0)
    tIF.sDestinoVou = String(255, 0)
    tIF.stitularcartao = String(255, 0)
    tIF.stitularcartaocpf = String(255, 0)
    tIF.spaxcartaofid = String(255, 0)
    tIF.spaxsobrenome = String(255, 0)
    tIF.spaxnome = String(255, 0)
    tIF.spaxsexo = String(255, 0)
    tIF.spaxtipodoc = String(255, 0)
    tIF.spaxnumdoc = String(255, 0)
    tIF.spaxendereco = String(255, 0)
    tIF.spaxbairro = String(255, 0)
    tIF.spaxcidade = String(255, 0)
    tIF.spaxcep = String(255, 0)
    tIF.spaxuf = String(255, 0)
    tIF.spaxemail = String(255, 0)
    tIF.spaxcontato = String(255, 0)
    tIF.spaxtelefone1 = String(255, 0)
    tIF.spaxtelefone2 = String(255, 0)
    tIF.sConvenio = String(255, 0)
    tIF.sNumAuto = String(255, 0)

    'Le a tabelacoinfoitemfatura
    lErro = Comando_Executar(lComando, "SELECT NumIntArq, NumRegArqTexto, codemp, data, numvou, moeda, valor, codest, numfat, dataven, numdoc, usuario, datareg, " & _
                "marca, tipodoc, recnsica, condpag, Serie, TipVou, codgru, Controle, cartao, CodPro, NumCCred, ValidCC, QuantParc, VlCartao, " & _
                "NumAuto, CiaCart, Exportado, DataExp, DataDep, TipoDocCorporator, NumIntDocCorporator, NoCorporatorEm, Excluido, ComErro, QtdPax, Grupo, diasantc, master, " & _
                "nummstr, NumIntDoc, DataAtualizacaoMovEst, CodigoMovEst, DataAtualizacaoContab, Kit, datarecepcao, emissor, destino, datainiciovigencia, datafimvigencia, idioma, destinovou, tarifamoeda, " & _
                "cambio, tarifareal, titularcartao, titularcartaocpf, paxcartaofid, paxsobrenome, paxnome, paxdatanasc, paxsexo, paxtipodoc, paxnumdoc, paxendereco, paxbairro, paxcidade, " & _
                "paxcep, paxuf, paxemail, paxcontato, paxtelefone1, paxtelefone2, CodEmpVou,Convenio, TarifaUnitaria FROM coinfoitemfatura WHERE TipoDoc = 'VOU' AND TipVou = ? AND Serie = ? AND NumVou = ? AND NumIntArq IN (SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE Sistema = 1 AND NomeArq = 'fpitemfatu')  ORDER BY NumIntArq DESC, NumRegArqTexto DESC ", _
                tIF.lNumIntArq, tIF.lNumRegArqTexto, tIF.lcodemp, tIF.dtData, tIF.lNumVou, tIF.sMoeda, tIF.dValor, tIF.scodest, tIF.lNumFat, _
                tIF.dtdataven, tIF.lnumdoc, tIF.sUsuario, tIF.dtdatareg, tIF.iMarca, tIF.sTipoDoc, tIF.lrecnsica, tIF.icondpag, tIF.sSerie, tIF.sTipVou, _
                tIF.scodgru, tIF.sControle, tIF.iCartao, tIF.sCodPro, tIF.sNumCCred, tIF.sValidCC, tIF.iQuantParc, tIF.dVlCartao, tIF.sNumAuto, tIF.sCiaCart, _
                tIF.iExportado, tIF.dtDataExp, tIF.dtDataDep, tIF.iTipoDocCorporator, tIF.lNumIntDocCorporator, tIF.dtNoCorporatorEm, tIF.iExcluido, tIF.iComErro, tIF.iQtdPax, tIF.sGrupo, _
                tIF.idiasantc, tIF.imaster, tIF.lnummstr, tIF.lNumIntDoc, tIF.dtDataAtualizacaoMovEst, tIF.lCodigoMovEst, tIF.dtDataAtualizacaoContab, tIF.lKit, tIF.dtDataRecepcao, tIF.lEmissor, _
                tIF.sDestino, tIF.dtdatainiciovigencia, tIF.dtdatafimvigencia, tIF.sIdioma, tIF.sDestinoVou, tIF.dtarifamoeda, tIF.dCambio, tIF.dtarifareal, tIF.stitularcartao, tIF.stitularcartaocpf, _
                tIF.spaxcartaofid, tIF.spaxsobrenome, tIF.spaxnome, tIF.dtpaxdatanasc, tIF.spaxsexo, tIF.spaxtipodoc, tIF.spaxnumdoc, tIF.spaxendereco, tIF.spaxbairro, tIF.spaxcidade, _
                tIF.spaxcep, tIF.spaxuf, tIF.spaxemail, tIF.spaxcontato, tIF.spaxtelefone1, tIF.spaxtelefone2, tIF.lCodEmpVou, tIF.sConvenio, tIF.dTarifaUnitaria, _
                objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> AD_SQL_SUCESSO Then gError 200171

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200172

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objVouInfo.sTipo = tIF.sTipVou
    objVouInfo.sSerie = tIF.sSerie
    objVouInfo.lNumVou = tIF.lNumVou
    objVouInfo.dtDataEmissao = tIF.dtData
    objVouInfo.dtDataRecepcao = tIF.dtDataRecepcao
    objVouInfo.lCliente = tIF.lCodEmpVou
    objVouInfo.lFornEmissor = tIF.lEmissor + FATOR_SOMA_COD_EMISSOR
    
    sCodPro = tIF.sCodPro
    Call TRVConverte_Prod(sCodPro)
'    sCodPro = Replace(sCodPro, "/", " ")
'    sCodPro = Replace(sCodPro, "\", " ")
'    sCodPro = Replace(sCodPro, "-", " ")
'    sCodPro = Replace(sCodPro, "_", " ")
'    sCodPro = Replace(sCodPro, "&", " ")
'    sCodPro = left(sCodPro, TAMANHO_SEGMENTO_PRODUTO)
'    sCodPro = sCodPro & String(TAMANHO_SEGMENTO_PRODUTO - Len(sCodPro), 32)
    
    objVouInfo.sProduto = sCodPro
    
    objVouInfo.sDestino = tIF.sDestino
    objVouInfo.dtDataInicio = tIF.dtdatainiciovigencia
    objVouInfo.dtDataTermino = tIF.dtdatafimvigencia
    objVouInfo.sVigencia = CStr(DateDiff("d", tIF.dtdatainiciovigencia, tIF.dtdatafimvigencia) + 1) & " dias"
    objVouInfo.sIdioma = tIF.sIdioma
    objVouInfo.iPax = tIF.iQtdPax
    objVouInfo.sDestinoVou = tIF.sDestinoVou
    objVouInfo.iAntc = tIF.idiasantc
    objVouInfo.sControle = tIF.sControle
    objVouInfo.sConvenio = tIF.sConvenio
    objVouInfo.dtDataPag = DATA_NULA
    objVouInfo.iCartao = tIF.iCartao
    objVouInfo.iPago = 0
    objVouInfo.lNumFat = tIF.lNumFat
    'objVouInfo.lCliPassageiro = tIF.lCliPassageiro
    objVouInfo.dtDataNasc = tIF.dtpaxdatanasc
    objVouInfo.sSexo = IIf(left(tIF.spaxsexo, 1) = "F", "Feminino", "Masculino")
    objVouInfo.sTipoDoc = tIF.spaxtipodoc
    objVouInfo.sCartaoFid = tIF.spaxcartaofid
    objVouInfo.sMoeda = tIF.sMoeda
    'objVouInfo.dTarifaUnitaria = tIF.dtarifamoeda
    objVouInfo.dCambio = tIF.dCambio
    objVouInfo.sValor = ""
    'objVouInfo.dTarifaPerc = tIF.dTarifaPerc
    objVouInfo.dTarifaValorMoeda = tIF.dtarifamoeda
    objVouInfo.dTarifaValorReal = tIF.dtarifareal
    'objVouInfo.dComissaoPerc = tIF.dComissaoPerc
    'objVouInfo.dComissaoValorMoeda = tIF.dComissaoValorMoeda
    'objVouInfo.dComissaoValorReal = tIF.dComissaoValorReal
    'objVouInfo.dCartaoPerc = tIF.dCartaoPerc
    'objVouInfo.dCartaoValorMoeda = tIF.dCartaoValorMoeda
    'objVouInfo.dCartaoValorReal = tIF.dCartaoValorReal
    'objVouInfo.dOverPerc = tIF.dOverPerc
    'objVouInfo.dOverValorMoeda = tIF.dOverValorMoeda
    'objVouInfo.dOverValorReal = tIF.dOverValorReal
    'objVouInfo.dCMRPerc = tIF.dCMRPerc
    'objVouInfo.dCMRValorMoeda = tIF.dCMRValorMoeda
    'objVouInfo.dCMRValorReal = tIF.dCMRValorReal
    objVouInfo.sCia = tIF.sCiaCart
    objVouInfo.sValidade = tIF.sValidCC
    objVouInfo.sNumeroCC = tIF.sNumCCred
    objVouInfo.sTitular = tIF.stitularcartao
    objVouInfo.sTitularCPF = tIF.stitularcartaocpf
    objVouInfo.dValorCartao = tIF.dVlCartao
    objVouInfo.lParcela = tIF.iQuantParc
    objVouInfo.sAprovacao = tIF.sNumAuto
    objVouInfo.sPassageiroSobreNome = tIF.spaxsobrenome
    objVouInfo.sPassageiroNome = tIF.spaxnome
    objVouInfo.sPassageiroCGC = tIF.spaxnumdoc
    objVouInfo.sPassageiroEndereco = tIF.spaxendereco
    objVouInfo.sPassageiroBairro = tIF.spaxbairro
    objVouInfo.sPassageiroCidade = tIF.spaxcidade
    objVouInfo.sPassageiroCEP = tIF.spaxcep
    objVouInfo.sPassageiroUF = tIF.spaxuf
    objVouInfo.sPassageiroEmail = tIF.spaxemail
    objVouInfo.sPassageiroContato = tIF.spaxcontato
    objVouInfo.sPassageiroTelefone1 = tIF.spaxtelefone1
    objVouInfo.sPassageiroTelefone2 = tIF.spaxtelefone2
    objVouInfo.sGrupo = tIF.sGrupo
    objVouInfo.sTitularCPF = tIF.stitularcartaocpf
    objVouInfo.dTarifaUnitaria = tIF.dTarifaUnitaria

    Obter_Dados_Kogut = SUCESSO

    Exit Function

Erro_Obter_Dados_Kogut:

    Obter_Dados_Kogut = gErr

    Select Case gErr

        Case 200171, 200172
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMFATURA", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200174)

    End Select

    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto11(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar Destinos

Dim lErro As Long, lTransacao As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim lComando As Long, iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String, lNumIntArq As Long
Dim fso As New FileSystemObject, ts As TextStream
Dim objDestino As ClassImpCoinfoDestino, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto11

    bAbriuArq = False

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130580
        
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_DESTINO", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpdestino", objImpCoinfo.sArqDestinoHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqDestinoHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqDestinoHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqDestinoHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqDestino = lNumIntArqAnt
        gError 130585
    End If
    
    objImpCoinfo.lNumIntArqDestino = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqDestinoHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
        Set objDestino = New ClassImpCoinfoDestino
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objDestino.sCodigo = StrParaStrCoinfo(sCampoValor)
                        Case 2
                            objDestino.sDescricao = StrParaStrCoinfo(sCampoValor)
                       
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 3 Then gError 130592
        objDestino.iExcluido = IIf(sCampoValor = "yes", 1, 0)
                
        'inserir registro
        With objDestino
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoDestino (NumIntArq, NumRegArqTexto, Codigo, Descricao, Excluido) VALUES (?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .sCodigo, .sDescricao, .iExcluido)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130583
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqDestinoHora & " teve " & CStr(lNumRegs) & " registros importados com sucesso." & vbNewLine
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Coinfo_ImportarArqTexto11 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto11:

    Coinfo_ImportarArqTexto11 = gErr
         
    Select Case gErr
        
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130585 'O arquivo já havia sido importado
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqDestinoHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqDestinoHora)

        Case 130592
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqDestinoHora, CStr(lNumRegs))
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFODESTINO", gErr, objDestino.sCodigo, objDestino.sDescricao)

        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM, 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
    
    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqDestinoHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto10(ByVal objImpCoinfo As ClassImpCoinfo) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lNumIntArq As Long

On Error GoTo Erro_Coinfo_ImportarArqTexto10

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130581
    
    'Abre o comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 130580
    
    'Abre o comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 130580

    lErro = CF("Config_ObterNumInt", "CoinfoConfig", "NUM_PROX_ARQ_ACORDO", lNumIntArq, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 130582
    
    lErro = Coinfo_ImportarArqTexto10a(lComando1, objImpCoinfo, lNumIntArq)
    If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572
    
    lErro = Coinfo_ImportarArqTexto10b(lComando2, objImpCoinfo, lNumIntArq)
    If lErro <> SUCESSO And lErro <> 130585 And lErro <> 53 Then gError 130572
    
    If lErro <> SUCESSO Then gError lErro
    
    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130588

    Coinfo_ImportarArqTexto10 = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto10:

    Coinfo_ImportarArqTexto10 = gErr
     
    Select Case gErr
          
        Case 130585, 53, 130572, 130582
        
        Case 130580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130588
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
    
    Call Transacao_Rollback
    
    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Private Function Coinfo_ImportarArqTexto10a(ByVal lComando, ByVal objImpCoinfo As ClassImpCoinfo, ByVal lNumIntArq As Long) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar Acordos

Dim lErro As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String
Dim fso As New FileSystemObject, ts As TextStream
Dim objAcordo As ClassImpCoinfoAcordo, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto10a

    bAbriuArq = False
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpacordo", objImpCoinfo.sArqAcordoHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqAcordoHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqAcordoHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqAcordoHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqAcordo = lNumIntArqAnt
        gError 130585
    End If
    
    objImpCoinfo.lNumIntArqAcordo = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqAcordoHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
         sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
        Set objAcordo = New ClassImpCoinfoAcordo
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objAcordo.lNumero = StrParaLong(sCampoValor)
                        Case 2
                            objAcordo.sContrato = left(StrParaStrCoinfo(sCampoValor), STRING_TRV_ACORDO_CONTRATO)
                        Case 3
                            objAcordo.lcodemp = StrParaLong(sCampoValor)
                        Case 4
                            objAcordo.dtdatavalidde = StrParaDateCoinfo(sCampoValor)
                        Case 5
                            objAcordo.dtdatavalidate = StrParaDateCoinfo(sCampoValor)
                        Case 6
                            objAcordo.sObservacao = StrParaStrCoinfo(sCampoValor)
                        Case 7
                            objAcordo.sDescricao = StrParaStrCoinfo(sCampoValor)
                       
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 8 Then gError 130592
        objAcordo.iExcluido = IIf(sCampoValor = "yes", 1, 0)
                
        'inserir registro
        With objAcordo
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoAcordo (NumIntArq, NumRegArqTexto, numero, contrato, codemp, datavalidde, datavalidate, observacao, descricao, excluido) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .lNumero, .sContrato, .lcodemp, .dtdatavalidde, .dtdatavalidate, .sObservacao, .sDescricao, .iExcluido)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130583
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
        
    Coinfo_ImportarArqTexto10a = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto10a:

    Coinfo_ImportarArqTexto10a = gErr
     
    Select Case gErr
        
        Case 130585 'O arquivo já havia sido importado
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqAcordoHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqAcordoHora)

        Case 130592
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqAcordoHora, CStr(lNumRegs))
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOACORDO", gErr, objAcordo.lNumero)

        Case ERRO_SEM_MENSAGEM, 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
    
    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqAcordoHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Exit Function

End Function

Private Function Coinfo_ImportarArqTexto10b(ByVal lComando, ByVal objImpCoinfo As ClassImpCoinfo, ByVal lNumIntArq As Long) As Long
'colocar o conteudo dos arquivos texto exportados pelo coinfo em tabelas do corporator.
'importar Acordos

Dim lErro As Long, iSecao As Integer, lTamanhoDados As Long, bArqAberto As Boolean
Dim sRegistro As String, iCampoPos As Integer, sCampoValor As String, iRegPos As Integer
Dim iIndice As Integer, iTestesOk As Integer, lNumRegs As Long
Dim sArqTimeStamp As String, lNumIntArqAnt As Long, bEmTexto As Boolean
Dim sAspas As String, sCaracter As String
Dim fso As New FileSystemObject, ts As TextStream
Dim objAcordoComis As ClassImpCoinfoAcordoComis, bAbriuArq As Boolean

On Error GoTo Erro_Coinfo_ImportarArqTexto10b

    bAbriuArq = False
    
    'inserir controle de que arquivo foi lido (se já havia sido lido, abortar)
    lErro = Coinfo_AtualizaArqsLidos(objImpCoinfo, "fpacordocomis", objImpCoinfo.sArqAcordoComisHora, lNumIntArqAnt, lNumRegs, lTamanhoDados, lNumIntArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Importando arquivo " & objImpCoinfo.sArqAcordoComisHora & " com NumRegArq " & CStr(lNumIntArq) & vbNewLine
    
    'se o arquivo já havia sido importado anteriormente
    If lNumIntArqAnt <> 0 Or objImpCoinfo.sArqAcordoComisHora = "" Then
    
        TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "O arquivo " & objImpCoinfo.sArqAcordoComisHora & " já havia sido importado com NumRegArq " & CStr(lNumIntArqAnt) & vbNewLine
        objImpCoinfo.lNumIntArqAcordoComis = lNumIntArqAnt
        gError 130585
    End If
    
    objImpCoinfo.lNumIntArqAcordoComis = lNumIntArq
    
    bArqAberto = False
    'abrir arquivo texto
    Set ts = fso.OpenTextFile(objImpCoinfo.sDiretorio & objImpCoinfo.sArqAcordoComisHora, 1, 0)
    bArqAberto = True
    bAbriuArq = True
        
    sAspas = Chr$(34)
    
    'leitura da seção de dados
    iSecao = 1
    
    'Até chegar ao fim do arquivo
    Do While Not ts.AtEndOfLine
    
        'Busca o próximo registro do arquivo
        sRegistro = ts.ReadLine
                
        lTamanhoDados = lTamanhoDados + Len(sRegistro) + 1 ' + 1 p/'\0x0A'
        
        lNumRegs = lNumRegs + 1
        
        Set objAcordoComis = New ClassImpCoinfoAcordoComis
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        For iRegPos = 1 To Len(sRegistro)
        
            sCaracter = Mid(sRegistro, iRegPos, 1)
        
            If bEmTexto Then
                If Mid(sRegistro, iRegPos, 1) = sAspas Then
                    bEmTexto = False
                Else
                    sCampoValor = sCampoValor & sCaracter
                End If
            Else
            
                If sCaracter = " " Then
                
                    Select Case iCampoPos
                    
                        Case 1
                            objAcordoComis.lnumeroacordo = StrParaLong(sCampoValor)
                        Case 2
                            objAcordoComis.iSequencial = StrParaInt(sCampoValor)
                        Case 3
                            objAcordoComis.sProduto = StrParaStrCoinfo(sCampoValor)
                        Case 4
                            objAcordoComis.sDestino = StrParaStrCoinfo(sCampoValor)
                        Case Else
                            gError 130592
                    End Select
                    
                    iCampoPos = iCampoPos + 1
                    sCampoValor = ""
                    
                Else
                    If sCaracter = Chr$(34) Then
                        bEmTexto = True
                    Else
                        sCampoValor = sCampoValor & sCaracter
                    End If
                End If
                
            End If
            
        Next
        
        If iCampoPos <> 5 Then gError 130592
        objAcordoComis.dPercComissao = StrParaDblCoinfo(sCampoValor)
                
        'inserir registro
        With objAcordoComis
            lErro = Comando_Executar(lComando, "INSERT INTO CoinfoAcordoComis (NumIntArq, NumRegArqTexto, numeroacordo, sequencial, produto, destino, perccomissao) VALUES (?,?,?,?,?,?,?)", _
                lNumIntArq, lNumRegs, .lnumeroacordo, .iSequencial, .sProduto, .sDestino, .dPercComissao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130583
        
        iCampoPos = 1
        sCampoValor = ""
        bEmTexto = False
        
    Loop
    
    'fechar arquivo texto
    ts.Close
    bArqAberto = False
        
    Coinfo_ImportarArqTexto10b = SUCESSO
     
    Exit Function
    
Erro_Coinfo_ImportarArqTexto10b:

    Coinfo_ImportarArqTexto10b = gErr
          
    Select Case gErr
        
        Case 130585 'O arquivo já havia sido importado
            If bArqAberto Then
                ts.Close
                bArqAberto = False
            End If
            If objImpCoinfo.sArqAcordoComisHora <> "" Then Call Kill(objImpCoinfo.sDiretorio & objImpCoinfo.sArqAcordoComisHora)

        Case 130592
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQ_COINFOITEMFATURA_LAYOUT_ERRADO", gErr, objImpCoinfo.sArqAcordoComisHora, CStr(lNumRegs))
        
        Case 130583
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COINFOACORDOCOMIS", gErr, objAcordoComis.lnumeroacordo, objAcordoComis.sProduto)

        Case ERRO_SEM_MENSAGEM, 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130562)
     
    End Select
    
    If bAbriuArq Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO", gErr, objImpCoinfo.sArqAcordoComisHora, CStr(lNumRegs), CStr(iCampoPos), sCampoValor)
    End If
     
    'fechar aquivo texto
    If bArqAberto Then ts.Close
    
    Exit Function

End Function

Private Function Coinfo_AtualizarDestinos(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza destinos

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim lNumIntArq As Long
Dim lNumRegArqTexto As Long
Dim colTabelaPrecoItem As Collection
Dim iTotalRegCertos As Integer
Dim iTotalRegErr As Integer
Dim sCodigo As String
Dim sDescricao As String
Dim lProxCodigo As Long
Dim lCodigoDestino As Long
Dim sChave As String

On Error GoTo Erro_Coinfo_AtualizarDestinos

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next
    
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpdestino")
    If lErro <> SUCESSO Then gError 182714
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130804
    
    lErro = Comando_Executar(alComando(3), "SELECT MAX(CodValor) FROM CamposGenericosValores WHERE CodCampo = 24", lProxCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 130807

    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fpdestino", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 130805
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
    Do While lErro = AD_SQL_SUCESSO
    
        iTotalRegCertos = iTotalRegCertos + 1
    
        sCodigo = String(255, 0)
        sDescricao = String(255, 0)
        
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, Descricao FROM CoinfoDestino WHERE NumIntArq = ? ORDER BY NumRegArqTexto", 0, sCodigo, sDescricao, lNumIntArq)
        If lErro <> AD_SQL_SUCESSO Then gError 130807
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808
        
        Do While lErro = AD_SQL_SUCESSO
        
            sChave = "Destino: " & sCodigo & " Descrição: " & sDescricao
        
            lErro = Comando_Executar(alComando(4), "SELECT CodValor FROM CamposGenericosValores WHERE Valor = ? AND CodCampo = 24", lCodigoDestino, sCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 130807
  
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808
            
            If lErro <> AD_SQL_SUCESSO Then
                
                lProxCodigo = lProxCodigo + 1
                lCodigoDestino = lProxCodigo
            
                lErro = Comando_Executar(alComando(5), "INSERT INTO CamposGenericosValores (CodCampo, CodValor, Valor, Complemento1) VALUES (?,?,?,?)", 24, lCodigoDestino, sCodigo, sDescricao)
                If lErro <> AD_SQL_SUCESSO Then gError 130813
            
            End If
            
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE CoinfoDestino SET NoCorporatorEm = ?, CodigoDestino = ?", alComando(1), Date, lCodigoDestino)
            If lErro <> AD_SQL_SUCESSO Then gError 130813
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 182720
    
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130812
        
        Loop
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComando(0), gdtDataHoje)
        If lErro <> SUCESSO Then gError 130813
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130814
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130815
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Os Destinos foram atualizados com sucesso. " & CStr(iTotalRegCertos) & " registros." & vbNewLine
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_AtualizarDestinos = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarDestinos:

    Coinfo_AtualizarDestinos = gErr
     
    Select Case gErr
          
        Case 130809, 130810, 130811, 130816, 130817, 182714, 182720
        
        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130804
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130815
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130805, 130806, 130814
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130813
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130807, 130808, 130812
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFODESTINO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_TRV_ATUALIZACAO_ARQUIVO", gErr, "Destinos", lNumIntArq, lNumRegArqTexto, sChave)

    Exit Function

End Function

Private Function Coinfo_AtualizarAcordos(ByVal objImpCoinfo As ClassImpCoinfo) As Long
'atualiza Acordos

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim lNumIntArq As Long
Dim lNumRegArqTexto As Long
Dim colTabelaPrecoItem As Collection
Dim iTotalRegCertos As Integer
Dim iTotalRegErr As Integer
Dim sCodigo As String
Dim objAcordo As ClassTRVAcordos
Dim objAcordoComis As ClassTRVAcordoComissao
Dim lNumero As Long
Dim sContrato As String
Dim lcodemp As Long
Dim dtdatavalidde As Date
Dim dtdatavalidate As Date
Dim sObservacao As String
Dim sDescricao As String
Dim iExcluido As Integer
Dim sProduto As String
Dim sDestino As String
Dim dPercComissao As Double
Dim iSeq As Integer
Dim sCodPro As String
Dim lDestino As Long
Dim sChave As String

On Error GoTo Erro_Coinfo_AtualizarAcordos

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next
    
    lErro = Coinfo_Total_Registros(objImpCoinfo, "fpacordo")
    If lErro <> SUCESSO Then gError 182714
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130804
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntImportacao FROM CoinfoArqsLidos WHERE NomeArq = ? AND DataAtualizado = ? ORDER BY IdentificacaoArq", 0, _
        lNumIntArq, "fpacordo", DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 130805
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130806
    
    Do While lErro = AD_SQL_SUCESSO
    
        iTotalRegCertos = iTotalRegCertos + 1
    
        sContrato = String(255, 0)
        sDescricao = String(255, 0)
        sObservacao = String(255, 0)
        
        lErro = Comando_ExecutarPos(alComando(1), "SELECT numero, contrato, codemp, datavalidde, datavalidate, observacao, descricao, excluido FROM CoinfoAcordo WHERE NumIntArq = ? ", 0, lNumero, sContrato, lcodemp, dtdatavalidde, dtdatavalidate, sObservacao, sDescricao, iExcluido, lNumIntArq)
        If lErro <> AD_SQL_SUCESSO Then gError 130807
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808
        
        Do While lErro = AD_SQL_SUCESSO
        
            sChave = "Acordo: " & CStr(lNumero)
        
            Set objAcordo = New ClassTRVAcordos
            
            objAcordo.dtValidadeAte = dtdatavalidate
            objAcordo.dtValidadeDe = dtdatavalidde
            objAcordo.iFilialCliente = 1
            objAcordo.lCliente = lcodemp
            objAcordo.lNumero = lNumero
            objAcordo.sContrato = sContrato
            objAcordo.sDescricao = sDescricao
            objAcordo.sObservacao = sObservacao
            
            If iExcluido = MARCADO Then
            
                lErro = CF("TRVAcordo_Exclui_EmTrans", objAcordo)
                If lErro <> SUCESSO Then gError 182714
            
            Else
            
                sProduto = String(255, 0)
                sDestino = String(255, 0)
            
                lErro = Comando_Executar(alComando(4), "SELECT Produto, Destino, PercComissao FROM CoinfoAcordoComis WHERE NumIntArq = ? AND NumeroAcordo = ? ORDER BY Sequencial", sProduto, sDestino, dPercComissao, lNumIntArq, lNumero)
                If lErro <> AD_SQL_SUCESSO Then gError 130807
                
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808
                
                iSeq = 0
            
                Do While lErro = AD_SQL_SUCESSO
                
                    iSeq = iSeq + 1
                    lDestino = 0
                    
                    lErro = Comando_Executar(alComando(5), "SELECT CodValor FROM CamposGenericosValores WHERE CodCampo = 24 AND Valor = ? ", lDestino, sDestino)
                    If lErro <> AD_SQL_SUCESSO Then gError 130807
                    
                    lErro = Comando_BuscarProximo(alComando(5))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808
                    
                    Set objAcordoComis = New ClassTRVAcordoComissao
                    objAcordo.colTRVAcordoComiss.Add objAcordoComis
                    
                    sCodPro = sProduto
                    Call TRVConverte_Prod(sCodPro)
'                    sCodPro = Replace(sCodPro, "/", " ")
'                    sCodPro = Replace(sCodPro, "\", " ")
'                    sCodPro = Replace(sCodPro, "-", " ")
'                    sCodPro = Replace(sCodPro, "_", " ")
'                    sCodPro = Replace(sCodPro, "&", " ")
'                    sCodPro = Trim(sCodPro) & String(10 - Len(Trim(sCodPro)), " ")
                    
                    objAcordoComis.dPercComissao = dPercComissao / 100
                    objAcordoComis.iDestino = lDestino
                    objAcordoComis.iSeq = iSeq
                    objAcordoComis.sProduto = sCodPro
            
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130812
                
                Loop
                    
                lErro = CF("TRVAcordo_Grava", objAcordo, False)
                If lErro <> SUCESSO Then gError 182714
        
            End If
            
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE CoinfoAcordo SET NoCorporatorEm = ?, NumIntDocAcordo = ?", alComando(1), Date, objAcordo.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 130813
            
            lErro = Rotina_Atualizar_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 182720
    
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130812
        
        Loop
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CoinfoArqsLidos SET DataAtualizado = ?", alComando(0), gdtDataHoje)
        If lErro <> SUCESSO Then gError 130813
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130814
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130815
    
    TelaAcompanhaBatchFAT.Log = TelaAcompanhaBatchFAT.Log & "Os acordos foram atualizados com sucesso. " & CStr(iTotalRegCertos) & " registros." & vbNewLine
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Coinfo_AtualizarAcordos = SUCESSO
     
    Exit Function
    
Erro_Coinfo_AtualizarAcordos:

    Coinfo_AtualizarAcordos = gErr
     
    Select Case gErr
          
        Case 130809, 130810, 130811, 130816, 130817, 182714, 182720
        
        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130804
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130815
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 130805, 130806, 130814
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOARQSLIDOS", gErr)
        
        Case 130813
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COINFOARQSLIDOS", gErr)
        
        Case 130807, 130808, 130812
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOACORDOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select
    
    Call Rotina_Erro(vbOKOnly, "ERRO_TRV_ATUALIZACAO_ARQUIVO", gErr, "Acordos", lNumIntArq, lNumRegArqTexto, sChave)
     
    Exit Function

End Function

Private Function Regera_Log_Vou() As Long
'atualiza Acordos

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 7) As Long
Dim objVouExp As ClassTRVVouExp
Dim tVou As typeTRVVouchers

On Error GoTo Erro_Regera_Log_Vou

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130804
    
    tVou.sSerie = String(255, 0)
    tVou.sTipVou = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT TipVou, Serie, NumVou, ClienteVou FROM TRVVouchers WHERE TipoDoc = 'VOU' AND Sistema = 1 AND Data >= {d '2008-09-13'} ", tVou.sTipVou, tVou.sSerie, tVou.lNumVou, tVou.lClienteVou)
    If lErro <> AD_SQL_SUCESSO Then gError 130807
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808

    Do While lErro = AD_SQL_SUCESSO

        Set objVouExp = New ClassTRVVouExp

        objVouExp.sSerie = tVou.sSerie
        objVouExp.sTipVou = tVou.sTipVou
        objVouExp.lNumVou = tVou.lNumVou
        objVouExp.lCliente = tVou.lClienteVou
        objVouExp.iTransacao = TRV_EXPORT_VOU_TRANS_ALT_COMI
        
        lErro = TRVVouExp_Insere(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), objVouExp)
        If lErro <> SUCESSO Then gError 200195

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808

    Loop

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130815
    
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Regera_Log_Vou = SUCESSO
     
    Exit Function
    
Erro_Regera_Log_Vou:

    Regera_Log_Vou = gErr
     
    Select Case gErr
          
        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130804
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130815
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function TRVVouExp_Insere(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal objVou As ClassTRVVouExp) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lNumIntTitulo As Long
Dim dtDataFat As Date
Dim lNumFat As Long
Dim dtDataPagto As Date
Dim lRepresentante As Long
Dim dComissaoRep As Double
Dim lCorrentista As Long
Dim dComissaoCorr As Double
Dim lClienteComissao As Long
Dim dComissaoAg As Double
Dim lEmissor As Long
Dim dComissaoEmissor As Double
Dim dtDataCanc As Date

On Error GoTo Erro_TRVVouExp_Insere

    If iSistemaIntegrado = SISTEMA_INTEGRADO_KOGUT Then
        
        lErro = objVou.Obtem_Dados_Geracao
        If lErro <> SUCESSO Then gError 200180
        
        'Precisa pegar os dados da comissão
        lErro = Comando_Executar(lComando6, "SELECT Representante, ComissaoRep, Correntista, ComissaoCorr, ClienteComissao, ComissaoAg, Emissor, ComissaoEmissor, DataCanc  FROM TRVVouchers AS V WHERE V.TipoDoc = ? AND V.TipVou= ? AND V.Serie = ? AND V.NumVou = ? ", lRepresentante, dComissaoRep, lCorrentista, dComissaoCorr, lClienteComissao, dComissaoAg, lEmissor, dComissaoEmissor, dtDataCanc, TRV_TIPODOC_VOU_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
        If lErro <> AD_SQL_SUCESSO Then gError 200208
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando6)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200209
        
        If lErro = AD_SQL_SUCESSO Then
            objVou.lRepresentanteComissao = lRepresentante
            objVou.lAgenciaComissao = lClienteComissao
            objVou.lCorrentistaComissao = lCorrentista
            objVou.lEmissorComissao = lEmissor - FATOR_SOMA_COD_EMISSOR
            
            If objVou.lEmissorComissao < 0 Then objVou.lEmissorComissao = 0
            
            objVou.dPercComiAg = dComissaoAg
            objVou.dPercComiCor = dComissaoCorr
            objVou.dPercComiEmi = dComissaoEmissor
            objVou.dPercComiRep = dComissaoRep
            objVou.dtDataCanc = dtDataCanc
        End If
            
        If objVou.iTransacao = TRV_EXPORT_VOU_TRANS_CANC_VOU Or objVou.iTransacao = TRV_EXPORT_VOU_TRANS_PAGTO Or objVou.iTransacao = TRV_EXPORT_VOU_TRANS_ALT_COMI Then
            
            'Precisa pegar os dados da fatura e da baixa se houver
            lErro = Comando_Executar(lComando4, "SELECT T.NumIntDoc, T.NumTitulo, T.DataEmissao FROM TRVVouchers AS V, TitulosRecTodos AS T WHERE V.TipoDocDestino = 3 AND V.NumIntDocDestino = T.NumIntDoc AND V.TipoDoc = ? AND V.TipVou= ? AND V.Serie = ? AND V.NumVou = ? ", lNumIntTitulo, lNumFat, dtDataFat, TRV_TIPODOC_VOU_TEXTO, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
            If lErro <> AD_SQL_SUCESSO Then gError 200208
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando4)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200209
            
            If lErro = AD_SQL_SUCESSO Then
                objVou.lNumFat = lNumFat
                objVou.dtDataFat = dtDataFat
            End If
            
            'Precisa pegar os dados da fatura e da baixa se houver
            lErro = Comando_Executar(lComando5, "SELECT B.Data FROM BaixasRec As B, BaixasParcRec AS BP, ParcelasRecTodas AS P WHERE P.NumParcela = 1 AND P.NumIntTitulo = ? AND P.NumIntDoc = BP.NumIntParcela AND BP.NumIntBaixa = B.NumIntBaixa AND BP.Status <> 5 AND B.Status <> 5 ORDER BY B.Data ", dtDataPagto, lNumIntTitulo)
            If lErro <> AD_SQL_SUCESSO Then gError 200210
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando5)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200211
            
            If lErro = AD_SQL_SUCESSO Then
                objVou.dtDataPag = dtDataPagto
            End If
        
        End If
    
        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt2", lComando2, lComando3, "TRVConfig", "NUM_INT_PROX_TRVVOUEXP", lNumIntDoc)
        If lErro <> SUCESSO Then gError 200181
        
        lErro = Comando_Executar(lComando1, "INSERT INTO TRVVouExp( NumIntDoc,Transacao, Usuario, Data, Hora, " & _
                "NumVou, TipVou, Serie, Cliente, NumFat, DataFat,DataCanc,DataPag, Exportado, " & _
                "UsuarioExportacao, DataExportacao, HoraExportacao, NomeArq,AgenciaComissao,RepresentanteComissao,EmissorComissao,CorrentistaComissao,PercComiAg,PergComiRep,PercComiEmi,PercComiCor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntDoc, objVou.iTransacao, objVou.sUsuario, objVou.dtData, objVou.dHora, _
                objVou.lNumVou, objVou.sTipVou, objVou.sSerie, objVou.lCliente, objVou.lNumFat, objVou.dtDataFat, objVou.dtDataCanc, objVou.dtDataPag, objVou.iExportado, _
                objVou.sUsuarioExportacao, objVou.dtDataExportacao, objVou.dHoraExportacao, objVou.sNomeArq, objVou.lAgenciaComissao, objVou.lRepresentanteComissao, objVou.lEmissorComissao, objVou.lCorrentistaComissao, objVou.dPercComiAg, objVou.dPercComiRep, objVou.dPercComiEmi, objVou.dPercComiCor)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200182

    End If

    TRVVouExp_Insere = SUCESSO

    Exit Function

Erro_TRVVouExp_Insere:

    TRVVouExp_Insere = gErr

    Select Case gErr
        
        Case 200180, 200181

        Case 200182
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVVOUEXP", gErr)
            
        Case 200208 To 200211
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200183)

    End Select
    
    Exit Function

End Function

Private Function Regera_Hierarquia_Vou() As Long
'atualiza Acordos

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim alComandoX(1 To 4) As Long
Dim tVou As typeTRVVouchers
Dim objVou As ClassTRVVouchers
Dim sSQL As String
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String
Dim iLenParamOut As Integer
Dim lConexaoTVA As Long
Dim objEmpresa As New ClassDicEmpresa

On Error GoTo Erro_Regera_Hierarquia_Vou

    objEmpresa.lCodigo = 1 'TVA'

    lErro = Empresa_Le(objEmpresa)
    If lErro <> SUCESSO Then gError 189756

    sDSN = objEmpresa.sStringConexao
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024

    lConexaoTVA = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoTVA = 0 Then gError 197890

    For iIndice = LBound(alComandoX) To UBound(alComandoX)
        alComandoX(iIndice) = Comando_AbrirExt(lConexaoTVA)
        If alComandoX(iIndice) = 0 Then gError 190857
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130803
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130804
    
    'Le os voucher sem hierarquia de vendedores sem estar cancelado
    'Não faturados, ou que a fatura não enha sido paga ou paga após 20/09/2008
    sSQL = "SELECT V.TipVou, V.Serie, V.NumVou, V.ClienteVou FROM TRVVouchers AS V WHERE TipoDoc = 'VOU' AND Status <> 7 AND NOT EXISTS " & _
        "(SELECT H.Vendedor FROM TRVVouVendedores AS H WHERE V.TipVou = H.TipVou AND V.Serie = H.Serie AND V.NumVou = H.NumVou) AND " & _
        "((NumFatCoinfo = 0) OR (Data >= {d '2008-09-01'}) OR (NumIntDocDestino NOT IN ( " & _
        "SELECT NumIntTitulo FROM ParcelasRecBaixadas )) OR " & _
        "(NumIntDocDestino IN ( " & _
        "SELECT NumIntTitulo FROM ParcelasRecTodas WHERE NumIntDoc IN ( " & _
        "SELECT NumIntParcela FROM BaixasParcRec WHERE NumIntBaixa IN ( " & _
        "SELECT NumIntBaixa FROM BaixasRec WHERE Data >= {d '2008-09-20'}))))) "

    tVou.sSerie = String(255, 0)
    tVou.sTipVou = String(255, 0)
    
    lErro = Comando_Executar(alComando(1), sSQL, tVou.sTipVou, tVou.sSerie, tVou.lNumVou, tVou.lClienteVou)
    If lErro <> AD_SQL_SUCESSO Then gError 130807
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808

    Do While lErro = AD_SQL_SUCESSO

        Set objVou = New ClassTRVVouchers
        
        objVou.sSerie = tVou.sSerie
        objVou.sTipVou = tVou.sTipVou
        objVou.lClienteVou = tVou.lClienteVou
        objVou.lNumVou = tVou.lNumVou

        lErro = TRVVouVendedores_Grava(objVou, alComandoX(1), alComandoX(2), alComandoX(3), alComando(2), alComandoX(4))
        If lErro <> SUCESSO Then gError 200399

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130808

    Loop
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComandoX) To UBound(alComandoX)
        Call Comando_Fechar(alComandoX(iIndice))
    Next
    
    Call Conexao_FecharExt(lConexaoTVA)

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130815

    Regera_Hierarquia_Vou = SUCESSO
     
    Exit Function
    
Erro_Regera_Hierarquia_Vou:

    Regera_Hierarquia_Vou = gErr
     
    Select Case gErr
          
        Case 130803
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130804
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130815
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130558)
     
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComandoX) To UBound(alComandoX)
        Call Comando_Fechar(alComandoX(iIndice))
    Next
    
    Call Conexao_FecharExt(lConexaoTVA)
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function TRVVouVendedores_Grava(ByVal objVou As ClassTRVVouchers, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iPromotor As Integer
Dim colVend As New Collection
Dim objVend As ClassVendedor
Dim iAux As Integer

On Error GoTo Erro_TRVVouVendedores_Grava

    'Le a Produto
    lErro = Comando_Executar(lComando1, "SELECT Vendedor FROM TRVProdutos WHERE Codigo = ? ", _
    iPromotor, objVou.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 200391
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200392
    
    If iPromotor = 0 Then
    
        'Le a Filiais Clientes
        lErro = Comando_Executar(lComando2, "SELECT Vendedor FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", _
        iPromotor, objVou.lClienteVou, FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 200393
        
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200394

    End If
    
    objVou.lPromotor = iPromotor
    
    lErro = TRVVendedor_Le_Recursivo(lComando3, colVend, iPromotor, 1)
    If lErro <> SUCESSO Then gError 200395
    
    'Le o Responsável
    lErro = Comando_Executar(lComando5, "SELECT Filial FROM FilialClienteCategorias WHERE Categoria = ? AND Cliente = ? AND Item = ?", _
    iAux, "Responsável", objVou.lClienteVou, "Call Center")
    If lErro <> AD_SQL_SUCESSO Then gError 200400
    
    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando5)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200401
    
    If lErro = AD_SQL_SUCESSO Then
        Set objVend = New ClassVendedor
        objVend.iCodigo = 9999
        colVend.Add objVend
    End If
    
    For Each objVend In colVend
    
        lErro = Comando_Executar(lComando4, "INSERT INTO TRVVouVendedores( TipVou, Serie, NumVou, Vendedor) VALUES (?,?,?,?)", objVou.sTipVou, objVou.sSerie, objVou.lNumVou, objVend.iCodigo)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200396

    Next

    TRVVouVendedores_Grava = SUCESSO

    Exit Function

Erro_TRVVouVendedores_Grava:

    TRVVouVendedores_Grava = gErr

    Select Case gErr

        Case 200391, 200392
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 200393, 200394
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr)
            
        Case 200395

        Case 200396
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRVVOUVENDEDORES", gErr)
        
        Case 200400, 200401
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200397)

    End Select

    Exit Function

End Function

Public Function TRVVendedor_Le_Recursivo(ByVal lComando As Long, ByVal colVendedores As Collection, ByVal iVendedorFilho As Integer, ByVal iNivel As Integer) As Long

Dim lErro As Long
Dim iVendedorPai  As Integer
Dim objVendedor As ClassVendedor

On Error GoTo Erro_TRVVendedor_Le_Recursivo

    If iVendedorFilho <> 0 Then

        If iNivel > 20 Then gError 200385
       
        'Le a tabelaTRVProdTarifa
        lErro = Comando_Executar(lComando, "SELECT Superior FROM VendedorTRV WHERE Vendedor = ? ", _
                    iVendedorPai, iVendedorFilho)
        If lErro <> AD_SQL_SUCESSO Then gError 200386
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200387
    
        'Sem Dados
        If lErro <> AD_SQL_SUCESSO Then gError 200388
        
        If iVendedorPai <> 0 Then
        
            Set objVendedor = New ClassVendedor
            
            objVendedor.iCodigo = iVendedorPai
            
            colVendedores.Add objVendedor
        
            lErro = TRVVendedor_Le_Recursivo(lComando, colVendedores, iVendedorPai, iNivel + 1)
            If lErro <> SUCESSO Then gError 200389
        
        End If
    
    End If
    
    TRVVendedor_Le_Recursivo = SUCESSO

    Exit Function

Erro_TRVVendedor_Le_Recursivo:

    TRVVendedor_Le_Recursivo = gErr

    Select Case gErr
    
        Case 200385 'ERRO_LIMITE_RECURSIVIDADE
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)
    
        Case 200386, 200387
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORES", gErr)

        Case 200388
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, iVendedorFilho)
            
        Case 200389
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200390)

    End Select

    Exit Function

End Function

Private Sub Importa_Ocrs(ByVal objImpCoinfo As ClassImpCoinfo)

Dim objFSO As New FileSystemObject, objTS As TextStream, lTransacao As Long
Dim sRegistro As String, sNomeArquivo As String, sDiretorio As String
Dim objOCR As ClassTRVOcorrencias, objOCRBD As ClassTRVOcorrencias
Dim objOcrDet As ClassTRVOcorrenciaDet, lID As Long
Dim objVou As ClassTRVVouchers, lCodigo As Long
Dim objFolder As Folder, lErro As Long
Dim objFile As File, sDate As String, iDir As Integer, sData As String

On Error GoTo Erro_Importa_Ocrs
   
    sDiretorio = "\\ARMAGEDOM\SGE\Import\OCR\"
    'sDiretorio = "C:\Clientes\Travel\Import\OCR\" 'TESTE
    
    sDate = Format(Now, "yyyymmddhhmmss")
    
    'Pega todos os aquivos da pasta genérica
    iDir = 1
    Set objFolder = objFSO.GetFolder(sDiretorio)
    iDir = 0
    
    'Para cada arquivo copia para pasta das empresas e apaga do genérico
    For Each objFile In objFolder.Files
    
        iDir = 2
        lErro = Copia_Arquivo(sDiretorio & objFile.Name, sDiretorio & "TVA\" & sDate & "_" & objFile.Name)
        If lErro <> SUCESSO Then gError 205812
        iDir = 0
        
        iDir = 3
        lErro = Copia_Arquivo(sDiretorio & objFile.Name, sDiretorio & "TVI\" & sDate & "_" & objFile.Name)
        If lErro <> SUCESSO Then gError 205813
        iDir = 0
        
        iDir = 4
        Call Kill(sDiretorio & objFile.Name)
        iDir = 0
    
    Next
    
    sDiretorio = sDiretorio & gsEmpresaTRV & "\"
    
    'Pega todos os aquivos da pasta da empresa
    iDir = 1
    Set objFolder = objFSO.GetFolder(sDiretorio)
    iDir = 0
    
    'Para cada arquivo da empresa
    For Each objFile In objFolder.Files

        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 205814
    
        iDir = 4
        Set objTS = objFSO.OpenTextFile(sDiretorio & objFile.Name, 1, 0)
        iDir = 0
        
        Do While Not objTS.AtEndOfLine
        
            Set objOCR = New ClassTRVOcorrencias
            Set objOcrDet = New ClassTRVOcorrenciaDet
            Set objVou = New ClassTRVVouchers
            Set objOCRBD = New ClassTRVOcorrencias
    
            'Busca o próximo registro do arquivo
            sRegistro = objTS.ReadLine
           
            objOCR.lID = StrParaLong(Mid(sRegistro, 1, 9))
            objOCR.sTipoDoc = Mid(sRegistro, 10, 1)
            objOCR.sSerie = Mid(sRegistro, 11, 1)
            objOCR.lNumVou = StrParaLong(Mid(sRegistro, 12, 9))
            objOCR.dValorTotal = StrParaDbl(Replace(Mid(sRegistro, 21, 10), ".", ","))
            sData = Mid(sRegistro, 31, 10) 'mm/dd/yyyy
            objOCR.dtDataEmissao = StrParaDate(Mid(sData, 4, 3) & left(sData, 3) & right(sData, 4))
            objOCR.sHistorico = Mid(sRegistro, 41, 255)
            objOCR.sObservacao = Mid(sRegistro, 296, 255)
            objOCR.iFormaPagto = FORMAPAGTO_TRV_OCR_FAT
           
            objVou.lNumVou = objOCR.lNumVou
            objVou.sTipVou = objOCR.sTipoDoc
            objVou.sSerie = objOCR.sSerie
            objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
            
            lErro = CF("TRVVouchers_Le", objVou)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205815
            
            If lErro = SUCESSO And objVou.iStatus <> STATUS_TRV_VOU_CANCELADO Then
            
                objOCR.dValorOCRCMAVou = objOCR.dValorTotal * IIf(objVou.iCartao = MARCADO, 0, objVou.dComissaoAg)
                objOCR.dValorOCRBrutoVou = objOCR.dValorTotal + objOCR.dValorOCRCMAVou
                           
                objOCR.lCliente = objVou.lClienteVou
                objOCR.iFilialCliente = FILIAL_MATRIZ
                
                objOCR.iOrigem = IMPORTACAO_OCORRENCIA_CODIGO
                
                objOCRBD.lID = objOCR.lID
                    
                lErro = CF("TRVOcorrencias_Le", objOCRBD)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205816
                
                If lErro = SUCESSO Then
                
                    If objOCRBD.lNumIntDocDestino <> 0 Then gError 205817 'JÁ FATURADO
                    If objOCRBD.lNumVou <> objOCR.lNumVou Then gError 205818  'NÃO PODE MUDAR DE VOUCHER
                
                    objOCR.lCodigo = objOCRBD.lCodigo 'Pegar próximo numero
                    
                Else
                    
                    lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_TRVOCORRENCIAS", "TRVOcorrencias", "Codigo", lCodigo)
                    If lErro <> SUCESSO Then gError 205819
                    
                    objOCR.lCodigo = lCodigo
                    
                End If
                
                objOCR.iStatus = STATUS_TRV_OCR_BLOQUEADO
                    
                objOcrDet.iTipo = IMPORTACAO_OCORRENCIA_TIPO_IMP_CODIGO
                objOcrDet.iSeq = 1
                objOcrDet.dValor = objOCR.dValorTotal
                
                objOCR.colDetalhamento.Add objOcrDet
            
                lErro = CF("TRVOcorrencias_Grava_EmTrans", objOCR)
                If lErro <> SUCESSO Then gError 205820
        
            End If
        
        Loop
                
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 205821
        
        'Copia para pasta de processados e apaga o anterior
        iDir = 5
        lErro = Copia_Arquivo(sDiretorio & objFile.Name, sDiretorio & "Proc\" & objFile.Name)
        If lErro <> SUCESSO Then gError 76
        iDir = 0
        
        objTS.Close
        
        iDir = 4
        Call Kill(sDiretorio & objFile.Name)
        iDir = 0
        
    Next
    
    Exit Sub

Erro_Importa_Ocrs:

    Select Case gErr
    
        Case 53, 70, 76
            If iDir = 1 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio, "Importa_Ocrs")
            ElseIf iDir = 2 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & "TVA\", "Importa_Ocrs")
            ElseIf iDir = 3 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & "TVI\", "Importa_Ocrs")
            ElseIf iDir = 4 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & objFile.Name, "Importa_Ocrs")
            ElseIf iDir = 5 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & "Proc\" & objFile.Name, "Importa_Ocrs")
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, "", "Importa_Ocrs")
            End If
    
        Case 205812, 205813
        
        Case 205814
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 205815, 205816, 205819, 205820, 205822
        
        Case 205817
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCR_FATURADA", gErr, objOCRBD.lID, objOCRBD.lCodigo)
        
        Case 205818
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCR_MUDOU_VOU", gErr, objOCRBD.lID, objOCRBD.lCodigo)
        
        Case 205821
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205823)

    End Select

    Call Transacao_Rollback

    Exit Sub

End Sub

Private Function Importa_OcrsCasos(ByVal objImpCoinfo As ClassImpCoinfo) As Long

Dim lTransacao As Long, alComando(1 To 7) As Long, iIndice As Integer
Dim sRegistro As String, sNomeArquivo As String, sDiretorio As String
Dim objFolder As Folder, lErro As Long, lNumIntArq As Long
Dim objFile As File, objFSO As New FileSystemObject
Dim objPastaTrabalho As Object 'Excel.Workbook
Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objOcrCasoImp As ClassTRVOcrCasoImport, sData As String
Dim iPOS As Integer, sNomeCompleto As String, iLinha As Integer
Dim objOcrCaso As ClassTRVOcrCasos, objVou As ClassTRVVouchers
Dim objVouAux As ClassTRVVoucherInfo, dCambio As Double, iDir As Integer
Dim objCidade As ClassCidades, iCodigoCidadeNovo As Integer, bArqAberto As Boolean

On Error GoTo Erro_Importa_OcrsCasos

    bArqAberto = False
        
    lErro = CF("TRVConfig_Le", TRVCONFIG_DIR_ASSISTENCIA, EMPRESA_TODA, sDiretorio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    If Len(Trim(sDiretorio)) > 0 Then
            
        'Pega todos os aquivos da pasta da empresa
        iDir = 1
        Set objFolder = objFSO.GetFolder(sDiretorio)
        iDir = 0
        
        'Para cada arquivo da empresa
        For Each objFile In objFolder.Files
    
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 208420
            
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 208421
            Next
            
            lErro = CF("Config_ObterNumInt", "TRVConfig", "NUM_PROX_ARQ_OCR_CASOS", lNumIntArq, 1, EMPRESA_TODA)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO TRVOcrCasoArqImport (NumIntArq, Data, Hora, NomeArq) VALUES (?,?,?,?)", _
                lNumIntArq, Date, CDbl(Time), objFile.Name)
            If lErro <> AD_SQL_SUCESSO Then gError 208422
            
            'Abre o excel
            lErro = CF("Excel_Abrir")
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            'Desabilita as mensagens do Excel
            GL_objExcelSist.DisplayAlerts = False
            
            iDir = 3
            Call GL_objExcelSist.Workbooks.Open(sDiretorio & objFile.Name)
            iDir = 0
            bArqAberto = True
            
            Set objPastaTrabalho = GL_objExcelSist.ActiveWorkBook
            
            'Seleciona a planilha ativa na pasta de trabalho criada
            Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
            'GL_objExcelSist.Visible = True
            
            iLinha = 1
            sData = objPlanilhaExcel.range("A" & CStr(iLinha))
            If IsDate(sData) Then
                iLinha = 0 'Não tem cabeçaho, começa da primeira linha
            End If
            TelaAcompanhaBatchFAT.TotReg = "0"
            Do While True
            
                iLinha = iLinha + 1
                           
                sData = objPlanilhaExcel.range("A" & CStr(iLinha))
                
                If Len(Trim(sData)) = 0 Then Exit Do
                
                Set objOcrCasoImp = New ClassTRVOcrCasoImport
                Set objOcrCaso = New ClassTRVOcrCasos
                
                If Not IsDate(sData) Then gError 208423
                
                objOcrCasoImp.lNumIntArq = lNumIntArq
                objOcrCasoImp.dtData = StrParaDate(sData)
                objOcrCasoImp.sCodigo = objPlanilhaExcel.range("B" & CStr(iLinha))
                objOcrCasoImp.sTCliente = objPlanilhaExcel.range("C" & CStr(iLinha))
                
                sNomeCompleto = objPlanilhaExcel.range("D" & CStr(iLinha))
                iPOS = InStr(1, sNomeCompleto, "/")
                If iPOS > 0 Then
                    objOcrCasoImp.sSobrenome = left(sNomeCompleto, iPOS - 1)
                    objOcrCasoImp.sNome = Mid(sNomeCompleto, iPOS + 1)
                Else
                    objOcrCasoImp.sNome = sNomeCompleto
                    objOcrCasoImp.sSobrenome = ""
                End If
                
                objOcrCasoImp.lSeq = iLinha - 1
                objOcrCasoImp.sChaveVou = objPlanilhaExcel.range("E" & CStr(iLinha))
                objOcrCasoImp.sCidadeOCR = objPlanilhaExcel.range("F" & CStr(iLinha))
                objOcrCasoImp.sEstadoOCR = objPlanilhaExcel.range("G" & CStr(iLinha))
                objOcrCasoImp.sPaisOCR = objPlanilhaExcel.range("H" & CStr(iLinha))
                objOcrCasoImp.sPrestador = objPlanilhaExcel.range("I" & CStr(iLinha))
                objOcrCasoImp.sCarater = objPlanilhaExcel.range("J" & CStr(iLinha))
                objOcrCasoImp.sGrauSatisfacao = objPlanilhaExcel.range("K" & CStr(iLinha))
                objOcrCasoImp.sTelefone = objPlanilhaExcel.range("L" & CStr(iLinha))
                
                objOcrCasoImp.sNumVouTexto = Trim(Mid(objOcrCasoImp.sChaveVou, 1, 18))
                'Chave = A055B000505768    001
                'Chave =        7736288    000
                If Not IsNumeric(left(objOcrCasoImp.sChaveVou, 1)) Then
                    If IsNumeric(Mid(objOcrCasoImp.sChaveVou, 6, 9)) Then
                        objOcrCasoImp.sTipVou = UCase(Trim(left(objOcrCasoImp.sChaveVou, 1)))
                        objOcrCasoImp.sSerie = UCase(Trim(Mid(objOcrCasoImp.sChaveVou, 5, 1)))
                        objOcrCasoImp.lNumVou = StrParaLong(Mid(objOcrCasoImp.sChaveVou, 6, 9))
                        objOcrCasoImp.sNumVouTexto = Mid(objOcrCasoImp.sChaveVou, 6, 9)
                    End If
                Else
                    If IsNumeric(Trim(Mid(objOcrCasoImp.sChaveVou, 1, 18))) Then
                        If Len(Trim(Mid(objOcrCasoImp.sChaveVou, 1, 18))) < 10 Then
                            objOcrCasoImp.lNumVou = StrParaLong(Trim(Mid(objOcrCasoImp.sChaveVou, 1, 18)))
                        Else
                        'Chave = 9520021430        000
                            objOcrCaso.sProduto = "CVC       "
                            objOcrCaso.lClienteVou = 3376 'CVC
                        End If
                    End If
                End If
                
                lErro = Comando_Executar(alComando(2), "INSERT INTO TRVOcrCasoImport( NumIntArq, Seq, Codigo, Data, TCliente, " & _
                        "Nome, Sobrenome, ChaveVou, TipVou, Serie, " & _
                        "NumVou, CidadeOCR, EstadoOCR, PaisOCR, Prestador, " & _
                        "Carater, GrauSatisfacao, Telefone) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        objOcrCasoImp.lNumIntArq, objOcrCasoImp.lSeq, objOcrCasoImp.sCodigo, objOcrCasoImp.dtData, objOcrCasoImp.sTCliente, _
                        objOcrCasoImp.sNome, objOcrCasoImp.sSobrenome, objOcrCasoImp.sChaveVou, objOcrCasoImp.sTipVou, objOcrCasoImp.sSerie, _
                        objOcrCasoImp.lNumVou, objOcrCasoImp.sCidadeOCR, objOcrCasoImp.sEstadoOCR, objOcrCasoImp.sPaisOCR, objOcrCasoImp.sPrestador, _
                        objOcrCasoImp.sCarater, objOcrCasoImp.sGrauSatisfacao, objOcrCasoImp.sTelefone)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208424
                               
                objOcrCaso.sCodigo = objOcrCasoImp.sCodigo
                objOcrCaso.sTipVou = objOcrCasoImp.sTipVou
                objOcrCaso.sSerie = objOcrCasoImp.sSerie
                objOcrCaso.lNumVou = objOcrCasoImp.lNumVou
                objOcrCaso.sNumVouTexto = objOcrCasoImp.sNumVouTexto
                objOcrCaso.spaxnome = objOcrCasoImp.sNome & " " & objOcrCasoImp.sSobrenome
                objOcrCaso.dtDataPriEvento = objOcrCasoImp.dtData
                objOcrCaso.objEndereco.iCodigoPais = PAIS_BRASIL
                
                'Tenta ler o voucher para pegar mais informações
                If Len(Trim(objOcrCaso.sSerie)) > 0 And Len(Trim(objOcrCaso.sTipVou)) > 0 Then
                
                    Set objVou = New ClassTRVVouchers
                
                    objVou.sTipVou = objOcrCasoImp.sTipVou
                    objVou.sSerie = objOcrCasoImp.sSerie
                    objVou.lNumVou = objOcrCasoImp.lNumVou
                    objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
                
                    lErro = CF("TRVVouchers_Le", objVou, alComando(4), alComando(5))
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then
                    
                        objOcrCaso.dtDataEmissao = objVou.dtData
                        objOcrCaso.iQtdPax = objVou.iPax
                        objOcrCaso.lClienteVou = objVou.lClienteVou
                        objOcrCaso.sTitularNome = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
                        objOcrCaso.sNomeFavorecido = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
                        objOcrCaso.sProduto = objVou.sProduto
                                        
                        Set objVouAux = New ClassTRVVoucherInfo
                    
                        objVouAux.lNumVou = objVou.lNumVou
                        objVouAux.sSerie = objVou.sSerie
                        objVouAux.sTipo = objVou.sTipVou
                        
                        lErro = CF("TRVVoucherInfoSigav_Le", objVouAux, alComando(6))
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                    
                        If lErro <> SUCESSO Then gError 208429
                    
                        objOcrCaso.dtDataIda = objVouAux.dtDataInicio
                        objOcrCaso.dtDataVolta = objVouAux.dtDataTermino
                        
                        If objVouAux.sTipoDoc = "CPF" And Len(Trim(objVouAux.sPassageiroCGC)) = 11 Then
                            objOcrCaso.sFavorecidoCGC = objVouAux.sPassageiroCGC
                        End If
                        
                        objOcrCaso.objEndereco.sBairro = left(objVouAux.sPassageiroBairro, STRING_BAIRRO)
                        objOcrCaso.objEndereco.sCEP = Replace(objVouAux.sPassageiroCEP, "-", "")
                        If Len(Trim(objOcrCaso.objEndereco.sCEP)) <> 8 Then objOcrCaso.objEndereco.sCEP = ""
                        objOcrCaso.objEndereco.sCidade = left(objVouAux.sPassageiroCidade, STRING_CIDADE)
                        objOcrCaso.objEndereco.sContato = objVouAux.sPassageiroContato
                        objOcrCaso.objEndereco.sEmail = left(objVouAux.sPassageiroEmail, STRING_EMAIL)
                        objOcrCaso.objEndereco.sEndereco = left(objVouAux.sPassageiroEndereco, STRING_ENDERECO)
                        objOcrCaso.objEndereco.sTelefone1 = left(objVouAux.sPassageiroTelefone1, STRING_TELEFONE)
                        objOcrCaso.objEndereco.sTelefone2 = left(objVouAux.sPassageiroTelefone2, STRING_TELEFONE)
                        objOcrCaso.objEndereco.sSiglaEstado = Trim(objVouAux.sPassageiroUF)
                        If Len(Trim(objOcrCaso.objEndereco.sSiglaEstado)) <> 2 Then objOcrCaso.objEndereco.sSiglaEstado = ""
                        
                        If Len(Trim(objOcrCaso.objEndereco.sCidade)) > 0 Then
                        
                            Set objCidade = New ClassCidades
                            
                            objCidade.sDescricao = objOcrCaso.objEndereco.sCidade
                        
                            lErro = CF("Cidade_Le_Nome", objCidade)
                            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
                        
                            If lErro <> SUCESSO Then
                                                
                                lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigoCidadeNovo)
                                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
                                Set objCidade = New ClassCidades
                            
                                objCidade.iCodigo = iCodigoCidadeNovo
                                objCidade.sDescricao = objOcrCaso.objEndereco.sCidade
                                
                                lErro = CF("Cidade_Grava_EmTrans", objCidade)
                                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
                            End If
                            
                        End If
                        
                    End If
               
                End If
                
                 'Le o cambio dos vouchers emitidos no dia
                lErro = Comando_Executar(alComando(7), "SELECT TOP 1 Cambio FROM TrvVouchers WHERE Data <= ? AND Moeda = 'USD' ORDER BY Data DESC ", _
                            dCambio, objOcrCaso.dtDataPriEvento)
                If lErro <> AD_SQL_SUCESSO Then gError 208426
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208427
                
                If lErro = AD_SQL_SUCESSO Then
                    objOcrCaso.dCambio = dCambio
                End If
                
                lErro = CF("TRVOcrCasos_Grava_EmTrans", objOcrCaso, True)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                TelaAcompanhaBatchFAT.TotReg = CStr(StrParaInt(TelaAcompanhaBatchFAT.TotReg) + 1)
                DoEvents
            Loop
            
            'Fecha o Excel
            Call CF("Excel_Fechar")
            bArqAberto = False
                    
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
                    
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 208425
            
            'Copia para pasta de processados e apaga o anterior
            iDir = 2
            lErro = Copia_Arquivo(sDiretorio & objFile.Name, sDiretorio & "Proc\" & objFile.Name)
            If lErro <> SUCESSO Then gError 76
            iDir = 0
            
            iDir = 3
            Call Kill(sDiretorio & objFile.Name)
            iDir = 0
            
        Next
        
    End If
    
    Importa_OcrsCasos = SUCESSO
    
    Exit Function

Erro_Importa_OcrsCasos:

    Importa_OcrsCasos = gErr

    Select Case gErr
        
        Case 53, 70, 76
            If iDir = 1 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio, "Importa_OcrsCasos")
            ElseIf iDir = 2 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & "Proc\" & objFile.Name, "Importa_OcrsCasos")
            ElseIf iDir = 3 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & objFile.Name, "Importa_OcrsCasos")
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, "", "Importa_OcrsCasos")
            End If
            
        Case 208420
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                
        Case 208421
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 208422
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRVOCRCASOARQIMPORT", gErr)
                
        Case 208423
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASO_DATA_INVALIDA", gErr, sData, objFile.Name, iLinha)
                
        Case 208424
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_TRVOCRCASOIMPORT", gErr)
                
        Case 208425
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 208426, 208427
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVVOUCHERS", gErr)
            
        Case 208429
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUINFOSEIGAV_INEXISTENTE", gErr, objVouAux.sTipo & objVouAux.sSerie & CStr(objVouAux.lNumVou))
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208428)

    End Select
    
    If bArqAberto Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRV_IMPORTACAO_ARQUIVO2", gErr, objFile.Name, iLinha)
        Call CF("Excel_Fechar")
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    Exit Function

End Function

Private Function Importa_Cobertura(ByVal objImpCoinfo As ClassImpCoinfo) As Long

Dim lTransacao As Long, alComando(1 To 2) As Long, iIndice As Integer
Dim sRegistro As String, sDiretorio As String, sCampo As String
Dim objFolder As Folder, lErro As Long, bArqAberto As Boolean
Dim objFile As File, objFSO As New FileSystemObject, objTS As TextStream
Dim objProdLim As ClassTRVProdLimites, lSeq As Long
Dim iDir As Integer

On Error GoTo Erro_Importa_Cobertura
       
    bArqAberto = False
        
    lErro = CF("TRVConfig_Le", TRVCONFIG_DIR_COBERTURA, EMPRESA_TODA, sDiretorio)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    If Len(Trim(sDiretorio)) > 0 Then
            
        'Pega todos os aquivos da pasta
        iDir = 1
        Set objFolder = objFSO.GetFolder(sDiretorio)
        iDir = 0
        
        'Para cada arquivo
        For Each objFile In objFolder.Files
    
            'Abertura de transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 209073
            
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 209074
            Next
            
            lErro = Comando_Executar(alComando(1), "DELETE FROM TRVProdLimites")
            If lErro <> AD_SQL_SUCESSO Then gError 209075
            
            bArqAberto = False
            Set objTS = objFSO.OpenTextFile(objFile.Path)
            bArqAberto = True
        
            'Até chegar ao fim do arquivo
            Do While Not objTS.AtEndOfLine
            
                'Busca o próximo registro do arquivo
                sRegistro = objTS.ReadLine
                 
                Set objProdLim = New ClassTRVProdLimites
                 
                lErro = Importa_Cobertura_Campo(sRegistro, 1, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                If sCampo <> "0" And sCampo <> "1" Then gError 209076
                objProdLim.iTipo = StrParaInt(sCampo)
                 
                lErro = Importa_Cobertura_Campo(sRegistro, 2, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                Call TRVConverte_Prod(sCampo)
                objProdLim.sProduto = sCampo
                
                lErro = Importa_Cobertura_Campo(sRegistro, 3, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                objProdLim.sDescProd = left(sCampo, 250)
                
                lErro = Importa_Cobertura_Campo(sRegistro, 4, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                If Not IsNumeric(sCampo) Then gError 209077
                objProdLim.lCodServ = StrParaLong(sCampo)
                 
                lErro = Importa_Cobertura_Campo(sRegistro, 5, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                If Not IsNumeric(sCampo) Then gError 209078
                objProdLim.dValorLimite = StrParaDbl(sCampo)
                 
                lErro = Importa_Cobertura_Campo(sRegistro, 6, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                If sCampo = "BRL" Then
                    objProdLim.iMoeda = MOEDA_REAL
                ElseIf sCampo = "EUR" Then
                    objProdLim.iMoeda = 2 'Euro
                Else
                    objProdLim.iMoeda = MOEDA_DOLAR
                End If
                 
                lErro = Importa_Cobertura_Campo(sRegistro, 7, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                objProdLim.sDescricao1 = left(sCampo, 250)
                 
                lErro = Importa_Cobertura_Campo(sRegistro, 8, sCampo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                objProdLim.sDescricao2 = left(sCampo, 250)
                
                lSeq = lSeq + 1
                 
                lErro = Comando_Executar(alComando(1), "INSERT INTO TRVProdLimites( Produto, CodServ, Seq, DescProd, ValorLimite, " & _
                        "Moeda, Descricao1, Descricao2, Tipo) VALUES (?,?,?,?,?,?,?,?,?)", _
                        objProdLim.sProduto, objProdLim.lCodServ, lSeq, objProdLim.sDescProd, objProdLim.dValorLimite, _
                        objProdLim.iMoeda, objProdLim.sDescricao1, objProdLim.sDescricao2, objProdLim.iTipo)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209079
                
            Loop
            objTS.Close
            bArqAberto = False
                    
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
                    
            'Confirma a transação
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 209080
            
            'Copia para pasta de processados e apaga o anterior
            iDir = 2
            lErro = Copia_Arquivo(sDiretorio & objFile.Name, sDiretorio & "Proc\" & objFile.Name)
            If lErro <> SUCESSO Then gError 76
            iDir = 0
             
            iDir = 3
            Call Kill(sDiretorio & objFile.Name)
            iDir = 0
            
            Exit For
            
        Next
        
    End If
    
    Importa_Cobertura = SUCESSO
    
    Exit Function

Erro_Importa_Cobertura:

    Importa_Cobertura = gErr

    Select Case gErr
    
        Case 53, 70, 76
            If iDir = 1 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio, "Importa_Cobertura")
            ElseIf iDir = 2 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & "Proc\" & objFile.Name, "Importa_Cobertura")
            ElseIf iDir = 3 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, sDiretorio & objFile.Name, "Importa_Cobertura")
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_PATH_NOT_FOUND", gErr, "", "Importa_Cobertura")
            End If
                    
        Case 209073
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                
        Case 209074
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 209075
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRVPRODLIMITES", gErr)
                
        Case 209076
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVPRODLIMITES_TIPO_INVALIDO", gErr, sCampo, sRegistro)
                
        Case 209077
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVPRODLIMITES_CODIGO_INVALIDO", gErr, sCampo, sRegistro)
                
        Case 209078
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVPRODLIMITES_VALORLIMITE_INVALIDO", gErr, sCampo, sRegistro)
                
        Case 209079
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRVPRODLIMITES", gErr)
                
        Case 209080
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209081)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    If bArqAberto Then
        Call Rotina_Erro(vbOKOnly, "ERRO_TRVPRODLIMITES_IMPORTACAO", gErr, objFile.Name, sRegistro)
        objTS.Close
    End If

    Exit Function

End Function

Private Function Importa_Cobertura_Campo(ByVal sRegistro As String, ByVal iCampo As Integer, sCampo As String) As String

Dim iPos1 As Integer
Dim iPos2 As Integer
Dim iIndice As Integer
Dim sAux As String

On Error GoTo Erro_Importa_Cobertura_Campo

    sCampo = ""

    iPos2 = 0
    iPos1 = InStr(iPos2 + 1, sRegistro, "|")
    For iIndice = 1 To iCampo - 1
        iPos2 = iPos1
        iPos1 = InStr(iPos2 + 1, sRegistro, "|")
        If iPos1 = 0 Then iPos1 = Len(sRegistro) + 1
    Next
    sAux = Mid(sRegistro, iPos2 + 1, iPos1 - iPos2 - 1)
    sCampo = sAux
    
    Importa_Cobertura_Campo = SUCESSO
    
    Exit Function

Erro_Importa_Cobertura_Campo:

    Importa_Cobertura_Campo = gErr

    Call Rotina_Erro(vbOKOnly, "ERRO_TRVPRODLIM_FALHA_CAMPO", gErr, iCampo, sRegistro)

    Exit Function

End Function

Public Function Cria_Clientes_Cartao() As Long
'Função auxiliar que não é chamada em nenhum ponto
'Exclui os TitulosRec sem data de vencimento

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lNumVou As Long, sTipo As String, sSerie As String
Dim objVoucherInfo As ClassTRVVoucherInfo

On Error GoTo Erro_Cria_Clientes_Cartao

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209810
    Next
    
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209811

    'TEM QUE OBTER UM DOS VOUCHERS
    sTipo = String(STRING_TRV_OCR_TIPOVOU, 0)
    sSerie = String(STRING_TRV_OCR_SERIE, 0)
    
    'Le todos os voucher de cartão a aprtir de 2010 que já foram faturados e o cliente ainda não foi criado
    lErro = Comando_Executar(alComando(0), "SELECT NumVou, TipVou, Serie FROM TRVVouchers WHERE Cartao = 1 AND CliPassageiro = 0 AND NumFatCoinfo <> 0 AND Data >= {d '2010-01-01'}", lNumVou, sTipo, sSerie)
    If lErro <> AD_SQL_SUCESSO Then gError 209812

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209813
    
    iIndice = 0
    
    'Enquanto vouchers
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objVoucherInfo = New ClassTRVVoucherInfo
        
        iIndice = iIndice + 1

        objVoucherInfo.lNumVou = lNumVou
        objVoucherInfo.sSerie = sSerie
        objVoucherInfo.sTipo = sTipo

        lErro = CF("TRVVoucherInfoSigav_Le", objVoucherInfo, alComando(1))
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

        'CADASTRAR O PASSAGEIRO E ACERTA A INFORMAÇÃO EM VOUCHERINFO
        lErro = CF("TRV_VoucherInfo_Cadastra_Cliente_EmTrans", objVoucherInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If iIndice > 1000 Then Exit Do
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209814
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209815
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Cria_Clientes_Cartao = SUCESSO
     
    Exit Function
    
Erro_Cria_Clientes_Cartao:

    Cria_Clientes_Cartao = gErr
     
    Select Case gErr
    
        Case 209810
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 209811
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 209812 To 209814
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELLCTO", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case 209815
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209816)
     
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function

End Function

