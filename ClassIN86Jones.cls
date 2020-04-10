VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function ContaCtb_Formata(ByVal sConta As String) As String
Dim sContaEnxuta As String, lErro As Long

    If Len(Trim(sConta)) <> 0 Then
    
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then
            ContaCtb_Formata = sConta
        Else
            ContaCtb_Formata = sContaEnxuta
        End If
    
    Else
        ContaCtb_Formata = sConta
    End If
    
End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Private Sub Prepara_IN86()
'Verifica os diretórios onde serão criados os arquivos, cria os diretórios inexistentes
'e totaliza o número de registros a serem processados

Dim iPosContraBarra As Integer
Dim objIN86Arquivo As ClassIN86Arquivos
Dim sPasta As String

On Error GoTo Erro_Prepara_IN86

    'Limpa as variáveis que controlam a qtde de registros processados
    TelaAcompanhaBatchFIS.dValorTotal = 0
    TelaAcompanhaBatchFIS.dValorAtual = 0

    'Para cada arquivo na coleção
    For Each objIN86Arquivo In TelaAcompanhaBatchFIS.objIN86Modelo.colIN86Arquivos
    
        'Se é para gerar o DUMP do arquivo => 'incrementa a contagem do total de registros a processar
        If objIN86Arquivo.iDUMP = MARCADO Then TelaAcompanhaBatchFIS.dValorTotal = TelaAcompanhaBatchFIS.dValorTotal + 1
        
        'Se é para gerar o relatório de acompanhamento do arquivo => 'incrementa a contagem do total de registros a processar
        If objIN86Arquivo.iRelatAcompanhamento = MARCADO Then TelaAcompanhaBatchFIS.dValorTotal = TelaAcompanhaBatchFIS.dValorTotal + 1
        
        'Se é para gerar o layout do arquivo => 'incrementa a contagem do total de registros a processar
        If objIN86Arquivo.iLayout = MARCADO Then TelaAcompanhaBatchFIS.dValorTotal = TelaAcompanhaBatchFIS.dValorTotal + 1
        
        'Se o arquivo foi selecionado
        If objIN86Arquivo.iSelecionado = MARCADO Then
        
            'Incrementa a contagem do total de registros a processar
            TelaAcompanhaBatchFIS.dValorTotal = TelaAcompanhaBatchFIS.dValorTotal + 1
            
            'Inicializa variável que controla a busca pelo caracter "\" no nome do arquivo
            iPosContraBarra = 1
            
            'Enquanto houver o caracter "\" no nome do arquivo
            Do
            
                'Idenfica a posição da próxima contra-barra
                iPosContraBarra = InStr(iPosContraBarra, objIN86Arquivo.sNome, "\")
                
                'se não encontrar uma contrabarra,
                'significa que existe a pasta determinada,
                'portanto, pode sair do loop
                If iPosContraBarra = 0 Then Exit Do
                
                'Guarda o nome do diretório encontrado
                sPasta = Mid(objIN86Arquivo.sNome, 1, iPosContraBarra)
                
                'Se essa pasta não existe no computador
                If Len(Trim(Dir(sPasta, vbDirectory))) = 0 Then
                    
                    'se estava na raiz e não encontrou o drive--> erro
                    If right(sPasta, 2) = ":\" Then gError 102988
                    
                    'Cria a pasta
                    Call MkDir(sPasta)
                    
                End If
                
                'Incrementa a variável de controle, para procurar a
                'próxima contra-barra imediatamente após a posição da última encontrada
                iPosContraBarra = iPosContraBarra + 1
            
            'esse loop SEMPRE sairá por exit do
            Loop
        
        End If
        
    Next
    
    Exit Sub
    
Erro_Prepara_IN86:

    Select Case gErr
        
        Case 102988
            Call Rotina_Erro(vbOKOnly, "ERRO_DRIVE_INEXISTENTE", gErr, sPasta)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150866)
            
    End Select
    
    Exit Sub

End Sub

'Revisada e alterada por Luiz Nogueira em 28/01/04
Private Function Rotina_GeraIN86_AtualizaTelaBatch() As Long
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_GeraIN86_AtualizaTelaBatch

    '*** Para depurar, usando o BatchFis como .dll, o trecho abaixo deve estar comentado
    'Atualiza tela de acompanhamento do Batch
    
    'Redesenha a tela
    DoEvents
    
    'Atualiza na variável de controle o número de registros processados
    TelaAcompanhaBatchFIS.dValorAtual = TelaAcompanhaBatchFIS.dValorAtual + 1
    
    'Redesenha a tela
    DoEvents
    
    'Atualiza na tela o número de registros processados
    TelaAcompanhaBatchFIS.TotReg.Caption = CStr(TelaAcompanhaBatchFIS.dValorAtual)
    
    'Redesenha a tela
    DoEvents
    
    'Atualiza a barra de progresso
    TelaAcompanhaBatchFIS.ProgressBar1.Value = CInt((TelaAcompanhaBatchFIS.dValorAtual / TelaAcompanhaBatchFIS.dValorTotal) * 100)

    'Se o usuário clicou no botão cancelar
    If TelaAcompanhaBatchFIS.iCancelaBatch = CANCELA_BATCH Then

        'Pede confirmação do cancelamento
        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_GERAIN86")

        'Se o usuário confirmou o cancelamento => sai da função
        If vbMsgBox = vbYes Then gError 103647

        'Se o cancelamento não foi confirmado, limpa a flag que indica click no botão cancelar
        TelaAcompanhaBatchFIS.iCancelaBatch = 0

    End If
    '***
    
    Rotina_GeraIN86_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_GeraIN86_AtualizaTelaBatch:

    Rotina_GeraIN86_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 103647

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150867)

    End Select
       
    Exit Function

End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
Public Function Rotina_Gera_IN86(ByVal objIN86Modelo As ClassIN86Modelos, ByVal colIN86TiposArquivo As Collection) As Long
'Dispara a geração dos arquivos e relatórios para o IN86

Dim lErro As Long
Dim colRegistrosDump As Collection
Dim sNomeArquivo As String
Dim sNomeArq As String
Dim sNomeArquivoGerado As String
Dim objIN86Arquivo As ClassIN86Arquivos
Dim objIN86TipoArquivo As ClassIN86TiposArquivos

On Error GoTo Erro_Rotina_Gera_IN86

    'Cria as pastas onde ficarão os arquivos e totaliza o número de registros a processar
    Call Prepara_IN86

    'preenche a string com o nome padrao do arquivo leiame
    sNomeArq = App.Path & PATH_IN86 & ARQUIVO_NOME_LEIAME
    
    'gera um nome de arquivo LEIAME
    lErro = CF("Gera_Nome_Arquivo", sNomeArq, EXTENSAO_ARQUIVO_TXT)
    If lErro <> SUCESSO Then gError 102989
   
    'para cada arquivo da colecao
    For Each objIN86Arquivo In objIN86Modelo.colIN86Arquivos
    
        'se o arquivo esta selecionado
        If objIN86Arquivo.iSelecionado = MARCADO Then
        
            'Limpa a coleção que guarda os registros a serem inseridos no arquivo
            Set colRegistrosDump = New Collection
            
            'Instancia objIN86TipoArquivo com os dados do tipo do arquivo em questão
            Set objIN86TipoArquivo = colIN86TiposArquivo.Item(objIN86Arquivo.iTipo)
            
            'Guarda o nome que de ser usado como base para o nome do arquivo
            sNomeArquivoGerado = objIN86Arquivo.sNome
            
            'gera o nome do arquivo
            lErro = CF("Gera_Nome_Arquivo", sNomeArquivoGerado, EXTENSAO_ARQUIVO_TXT)
            If lErro <> SUCESSO Then gError 102990
            
            'devolve o nome do arquivo completo para o obj
            objIN86Arquivo.sNome = sNomeArquivoGerado
            
            'se o arquivo que está sendo gerado é do tipo auxiliar
            If objIN86TipoArquivo.iAuxiliar = TIPO_ARQUIVO_AUXILIAR Then
            
                'Chama a rotina de geração do arquivo, passando os devidos parâmetros
                lErro = CF(objIN86TipoArquivo.sRotinaGeradora, objIN86Modelo, objIN86Arquivo.sNome, colRegistrosDump)
                If lErro <> SUCESSO Then gError 102991
             
            'senão
            Else
            
                'Chama a rotina de geração do arquivo, passando os devidos parâmetros
                lErro = CF(objIN86TipoArquivo.sRotinaGeradora, objIN86Modelo, objIN86Arquivo, colRegistrosDump)
                If lErro <> SUCESSO Then gError 102992
                
            End If
            
            'Atualiza o número de arquivos processados e a barra de progresso
            lErro = Rotina_GeraIN86_AtualizaTelaBatch
            If lErro <> SUCESSO Then gError 102993
            
            'se é para gerar o DUMP do arquivo
            If objIN86Arquivo.iDUMP = MARCADO Then
            
                'Chama a rotina de geração de DUMP
                lErro = Gera_IN86_Rel_Dump(colRegistrosDump, objIN86Arquivo)
                If lErro <> SUCESSO Then gError 102994
                
                'Atualiza o número de arquivos processados e a barra de progresso
                lErro = Rotina_GeraIN86_AtualizaTelaBatch
                If lErro <> SUCESSO Then gError 102995

            End If
            
            'se é para gerar o layout do arquivo
            If objIN86Arquivo.iLayout = MARCADO Then
            
                'Chama a rotina de geração de layout
                lErro = Gera_IN86_ArqRel_Layout(objIN86TipoArquivo, sNomeArq)
                If lErro <> SUCESSO Then gError 102996
                
                'Atualiza o número de arquivos processados e a barra de progresso
                lErro = Rotina_GeraIN86_AtualizaTelaBatch
                If lErro <> SUCESSO Then gError 102997

            End If
            
            'se o arquivo pede para gerar relatório de acompanhamento e se foi gerado algum arquivo anteriormente
            If objIN86Arquivo.iRelatAcompanhamento = MARCADO Then
            
                'Chama a rotina de geração de relatório de acompanhamento
                lErro = Gera_IN86_ArqRel_Acompanhamento(objIN86Arquivo, objIN86TipoArquivo, sNomeArq, objIN86Modelo.iMeioEntrega, colRegistrosDump.Count)
                If lErro <> SUCESSO Then gError 102998
                
                'Atualiza o número de arquivos processados e a barra de progresso
                lErro = Rotina_GeraIN86_AtualizaTelaBatch
                If lErro <> SUCESSO Then gError 102999
            
            End If
        
        End If
        
    Next
        
    Rotina_Gera_IN86 = SUCESSO

    Exit Function

Erro_Rotina_Gera_IN86:

    Rotina_Gera_IN86 = gErr
    
    Select Case gErr
    
        Case 102989 To 102999
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150868)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_FornecedorHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para o Fornecedor usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objFornecedorHistorico As ClassFornecedorHistorico
Dim colFornecedorHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_FornecedorHistorico

    'Faz a leitura na tabela de FornecedorHistorico
    lErro = FornecedorHistorico_Le_Todos(objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, colFornecedorHistorico)
    If lErro <> SUCESSO And lErro <> 97794 Then gError 97821

    'Se não encontrou => erro
    If lErro = 97794 Then gError 97822

    'Cria o arquivo de FornecedorHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objFornecedorHistorico In colFornecedorHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_FornecedorHistorico(sRegistro, objFornecedorHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    Close #1

    Gera_IN86_Arquivo_FornecedorHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_FornecedorHistorico:

    Gera_IN86_Arquivo_FornecedorHistorico = gErr

    Select Case gErr

        Case 97821

        Case 97822
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150869)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_FornecedorHistorico(sRegistro As String, ByVal objFornecedorHistorico As ClassFornecedorHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objFornecedorHistorico RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sFornecedor As String

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objFornecedorHistorico.dtDataAtualizacao, "DDMMYYYY")

    'sFornecedor = CodFornecedor + CodFilialFornecedor(Com zeros a esquerda)
    sFornecedor = CStr(objFornecedorHistorico.lCodigo) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (objFornecedorHistorico.iFilialForn))

    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sFornecedor)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CGC_CPF, " ", objFornecedorHistorico.sCgc)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_ESTADUAL, " ", objFornecedorHistorico.sInscricaoEstadual)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_MUNICIPAL, " ", objFornecedorHistorico.sInscricaoMunicipal)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_RAZAO_SOCIAL, " ", objFornecedorHistorico.sRazaoSocial)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ENDERECO, " ", objFornecedorHistorico.sEndereco)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_BAIRRO, " ", objFornecedorHistorico.sBairro)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MUNICIPIO, " ", objFornecedorHistorico.sCidade)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UF, " ", objFornecedorHistorico.sUF)

    'Se o país for o Brasil
    If objFornecedorHistorico.sPais <> PAIS_BRASIL_NOME Then
        
        'Monta o registro considerando o país em branco
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_PAIS, " ", objFornecedorHistorico.sPais)
    
    'Senão, ou seja, se for outro país
    Else
        
        'Monta o registro considerando o nome do país
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_PAIS, " ", "")
    
    End If

    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CEP, " ", objFornecedorHistorico.sCEP)

End Sub

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_ClienteHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 a partir da tabela ClienteHistorico usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objClienteHistorico As ClassClienteHistorico
Dim colClienteHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_ClienteHistorico

    'Faz a leitura na tabela de ClienteHistorico
    lErro = ClienteHistorico_Le_Todos(objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, colClienteHistorico)
    If lErro <> SUCESSO And lErro <> 97805 Then gError 97832

    'Se não encontrou => erro
    If lErro = 97805 Then gError 97833

    'Cria o arquivo de ClienteHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objClienteHistorico In colClienteHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_ClienteHistorico(sRegistro, objClienteHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro

        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    Close #1

    Gera_IN86_Arquivo_ClienteHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_ClienteHistorico:

    Gera_IN86_Arquivo_ClienteHistorico = gErr

    Select Case gErr

        Case 97832

        Case 97833
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO_2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150870)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_ClienteHistorico(sRegistro As String, ByVal objClienteHistorico As ClassClienteHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objClienteHistorico RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sCliente As String

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objClienteHistorico.dtDataAtualizacao, "DDMMYYYY")

    'sCliente = CodCliente + CodFilialCliente(Com zeros a esquerda)
    sCliente = CStr(objClienteHistorico.lCodigo) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (objClienteHistorico.iFilialCliente))

    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sCliente)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CGC_CPF, " ", objClienteHistorico.sCgc)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_ESTADUAL, " ", objClienteHistorico.sInscricaoEstadual)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_MUNICIPAL, " ", objClienteHistorico.sInscricaoMunicipal)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_RAZAO_SOCIAL, " ", objClienteHistorico.sRazaoSocial)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ENDERECO, " ", objClienteHistorico.sEndereco)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_BAIRRO, " ", objClienteHistorico.sBairro)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MUNICIPIO, " ", objClienteHistorico.sCidade)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UF, " ", objClienteHistorico.sUF)

    'Se o país for o Brasil
    If objClienteHistorico.sPais <> PAIS_BRASIL_NOME Then
        
        'Monta o registro considerando o país em branco
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_PAIS, " ", objClienteHistorico.sPais)
    
    'Senão, ou seja, se for outro país
    Else
        
        'Monta o registro incluindo o nome do país
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_PAIS, " ", "")
    
    End If

    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CEP, " ", objClienteHistorico.sCEP)

End Sub

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_TransportadoraHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para a Transportadora usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objTranspHistorico As ClassTranspHistorico
Dim colTranspHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_TransportadoraHistorico

    'Faz a leitura na tabela de TransportadoraHistorico
    lErro = TranspHistorico_Le_Todos(objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, colTranspHistorico)
    If lErro <> SUCESSO And lErro <> 97798 Then gError 97835

    'Se não encontrou => erro
    If lErro = 97798 Then gError 97836

    'Cria o arquivo de TransportadoraHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objTranspHistorico In colTranspHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_TransportadoraHistorico(sRegistro, objTranspHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro

        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    'Fecha o arquivo
    Close #1

    Gera_IN86_Arquivo_TransportadoraHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_TransportadoraHistorico:

    Gera_IN86_Arquivo_TransportadoraHistorico = gErr

    Select Case gErr

        Case 97835

        Case 97836
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPHISTORICO_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150871)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_TransportadoraHistorico(sRegistro As String, ByVal objTranspHistorico As ClassTranspHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objTranspHistorico RECEBE (Input) os valores que serão utilizados para montar o registro

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objTranspHistorico.dtDataAtualizacao, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", objTranspHistorico.iCodTransp)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CGC_CPF, " ", objTranspHistorico.sCgc)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_ESTADUAL, " ", objTranspHistorico.sInscricaoEstadual)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_MUNICIPAL, " ", objTranspHistorico.sInscricaoMunicipal)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_RAZAO_SOCIAL, " ", objTranspHistorico.sRazaoSocial)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ENDERECO, " ", objTranspHistorico.sEndereco)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_BAIRRO, " ", objTranspHistorico.sBairro)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MUNICIPIO, " ", objTranspHistorico.sCidade)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UF, " ", objTranspHistorico.sUF)

    'Se o país for o Brasil
    If objTranspHistorico.sPais <> PAIS_BRASIL_NOME Then
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_PAIS, " ", objTranspHistorico.sPais)
    
    Else 'Senão, ou seja, se for outro país
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_PAIS, " ", "")
    End If

    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CEP, " ", objTranspHistorico.sCEP)

End Sub

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_CclHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para o Ccl usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objCclHistorico As ClassCclHistorico
Dim colCclHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_CclHistorico

    'Faz a leitura na tabela de CclHistorico
    lErro = CF("CclHistorico_Le_Todos", colCclHistorico)
    If lErro <> SUCESSO And lErro <> 97818 Then gError 97842

    'Se não encontrou => erro
    If lErro = 97818 Then gError 97843

    'Cria o arquivo de CclHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objCclHistorico In colCclHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_CclHistorico(sRegistro, objCclHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    'Fecha o arquivo
    Close #1

    Gera_IN86_Arquivo_CclHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_CclHistorico:

    Gera_IN86_Arquivo_CclHistorico = gErr

    Select Case gErr

        Case 97842

        Case 97843
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150872)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_CclHistorico(sRegistro As String, ByVal objCclHistorico As ClassCclHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objCclHistorico RECEBE (Input) os valores que serão utilizados para montar o registro
    
    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objCclHistorico.dtDataAtualizacao, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CCL, " ", objCclHistorico.sCodigoCcl)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESCRICAO, " ", objCclHistorico.sDescCcl)

End Sub

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_PlanoContaHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para o PlanoConta usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objPlanoContaHistorico As ClassPlanoContaHistorico
Dim colPlanoContaHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_PlanoContaHistorico

    'Faz a leitura na tabela de PlanoContaHistorico
    lErro = CF("PlanoContaHistorico_Le_Todos", colPlanoContaHistorico)
    If lErro <> SUCESSO And lErro <> 97810 Then gError 97839

    'Se não encontrou => erro
    If lErro = 97810 Then gError 97840

    'Cria o arquivo de PlanoContaHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objPlanoContaHistorico In colPlanoContaHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_PlanoContaHistorico(sRegistro, objPlanoContaHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    'Fecha o arquivo
    Close #1

    Gera_IN86_Arquivo_PlanoContaHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_PlanoContaHistorico:

    Gera_IN86_Arquivo_PlanoContaHistorico = gErr

    Select Case gErr

        Case 97839

        Case 97840
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTAHISTORICO_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150873)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_PlanoContaHistorico(sRegistro As String, ByVal objPlanoContaHistorico As ClassPlanoContaHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objPlanoContaHistorico RECEBE (Input) os valores que serão utilizados para montar o registro

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objPlanoContaHistorico.dtDataAtualizacao, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(objPlanoContaHistorico.sCodigo))
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_TIPOCONTA_CTB, " ", IIf(objPlanoContaHistorico.iTipo = 1, "S", "A"))
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(objPlanoContaHistorico.sContaPai))
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESCRICAO, " ", objPlanoContaHistorico.sDescricao)

End Sub

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_NaturezaOPHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para o NaturezaOP usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objNaturezaOPHistorico As ClassNaturezaOPHistorico
Dim colNaturezaOPHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_NaturezaOPHistorico

    'Faz a leitura na tabela de NaturezaOPHistorico
    lErro = CF("NaturezaOPHistorico_Le_Todos", colNaturezaOPHistorico)
    If lErro <> SUCESSO And lErro <> 97800 Then gError 97845

    'Se não encontrou => erro
    If lErro = 97800 Then gError 97846

    'Cria o arquivo de NaturezaOPHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objNaturezaOPHistorico In colNaturezaOPHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_NaturezaOPHistorico(sRegistro, objNaturezaOPHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    'Fecha o arquivo
    Close #1

    Gera_IN86_Arquivo_NaturezaOPHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_NaturezaOPHistorico:

    Gera_IN86_Arquivo_NaturezaOPHistorico = gErr

    Select Case gErr

        Case 97845

        Case 97846
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOPHISTORICO_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150874)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_NaturezaOPHistorico(sRegistro As String, ByVal objNaturezaOPHistorico As ClassNaturezaOPHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objNaturezaOPHistorico RECEBE (Input) os valores que serão utilizados para montar o registro

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objNaturezaOPHistorico.dtDataAtualizacao, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NATUREZAOP, " ", objNaturezaOPHistorico.sCodigo)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESCRICAO, " ", objNaturezaOPHistorico.sDescricao)

End Sub

'Revisado - OK Luiz 19/02/02
Public Function Gera_IN86_Arquivo_ProdutoHistorico(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para o Produto usando o nome do arquivo passado como parametro
'sNomeArquivo RECEBE Nome Completo Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim sRegistro As String
Dim objProdutoHistorico As ClassProdutoHistorico
Dim colProdutoHistorico As New Collection

On Error GoTo Erro_Gera_IN86_Arquivo_ProdutoHistorico

    'Faz a leitura na tabela de ProdutoHistorico
    lErro = ProdutoHistorico_Le_Todos(objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, colProdutoHistorico)
    If lErro <> SUCESSO And lErro <> 97820 Then gError 97848

    'Se não encontrou => erro
    If lErro = 97820 Then gError 97849

    'Cria o arquivo de ProdutoHistorico
    Open sNomeArquivo For Output As #1

    'Para cada registro lido
    For Each objProdutoHistorico In colProdutoHistorico

        'Prepara a linha para o registro encontrado na tabela
        Call Monta_Registro_ProdutoHistorico(sRegistro, objProdutoHistorico)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Inclui o registro no arquivo
        Print #1, sRegistro

    Next

    'Fecha o arquivo
    Close #1

    Gera_IN86_Arquivo_ProdutoHistorico = SUCESSO

    Exit Function

Erro_Gera_IN86_Arquivo_ProdutoHistorico:

    Gera_IN86_Arquivo_ProdutoHistorico = gErr

    Select Case gErr

        Case 97848

        Case 97849
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOHISTORICO_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150875)

    End Select

End Function

'Revisado - OK Luiz 19/02/02
Private Sub Monta_Registro_ProdutoHistorico(sRegistro As String, ByVal objProdutoHistorico As ClassProdutoHistorico)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objProdutoHistorico RECEBE (Input) os valores que serão utilizados para montar o registro

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objProdutoHistorico.dtDataAtualizacao, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", objProdutoHistorico.sCodProduto)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESCRICAO, " ", objProdutoHistorico.sDescProduto)

End Sub

Public Function Gera_IN86_Arquivo_LancamentosCTB(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Lançamentos Contábeis
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim lComando As Long
Dim tLancamento As typeLancamento
Dim objLancamento As ClassLancamento_Detalhe
Dim objLancamentoCol As ClassLancamento_Detalhe
Dim colLancamento As New Collection
Dim sSelect As String

On Error GoTo Erro_Gera_IN86_Arquivo_LancamentosCTB

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97880
    
    With tLancamento
        
        'Inicializa as Strings do Type
        .sCcl = String(STRING_CCL, 0)
        .sConta = String(STRING_CONTA, 0)
        .sHistorico = String(STRING_HISTORICO, 0)
        .sOrigem = String(STRING_ORIGEM, 0)
    
    End With
    
    'Se for a empresa toda => NAO TEM FILTRO
    If objIN86Arquivos.iFilialEmpresa = EMPRESA_TODA Or gobjCTB.giContabCentralizada <> 0 Then
    
        'Carrega sSelect com a Expressão SQL
        sSelect = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Seq, Data, Conta, Ccl, SeqContraPartida, Valor, Doc, Historico FROM Lancamentos WHERE Data >= ? AND Data <= ? AND Aglutinado <> ? ORDER BY FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq"
        
        'Executa o Comando SQL
        With tLancamento
            lErro = Comando_Executar(lComando, sSelect, .iFilialEmpresa, .sOrigem, .iExercicio, .iPeriodoLan, .iSeq, .dtData, .sConta, .sCcl, .iSeqContraPartida, .dValor, .lDoc, .sHistorico, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, LANCAMENTO_AGLUTINADO)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97881
                    
    Else 'FILTRO POR FILIAL
    
        'Carrega sSelect com a Expressao SQL
        sSelect = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Seq, Data, Conta, Ccl, SeqContraPartida, Valor, Doc, Historico FROM Lancamentos WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? AND Aglutinado <> ? ORDER BY FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq"
        
        'Executa o Comando SQL
        With tLancamento
            lErro = Comando_Executar(lComando, sSelect, .iFilialEmpresa, .sOrigem, .iExercicio, .iPeriodoLan, .iSeq, .dtData, .sConta, .sCcl, .iSeqContraPartida, .dValor, .lDoc, .sHistorico, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa, LANCAMENTO_AGLUTINADO)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97882
        
    End If
            
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97883

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97884
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tem lancamentos...
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Cria um novo obj
        Set objLancamento = New ClassLancamento_Detalhe
        
        'Passa para do type para o Obj
        With objLancamento
        
             .sOrigem = tLancamento.sOrigem
             .dtData = tLancamento.dtData
             .sConta = tLancamento.sConta
             .sCcl = tLancamento.sCcl
             .iSeqContraPartida = tLancamento.iSeqContraPartida
             .dValor = tLancamento.dValor
             .lDoc = tLancamento.lDoc
             .sHistorico = tLancamento.sHistorico
             .iFilialEmpresa = objIN86Arquivos.iFilialEmpresa
             .iExercicio = tLancamento.iExercicio
             .iPeriodoLan = tLancamento.iPeriodoLan
             .iSeq = tLancamento.iSeq
        
        End With
        
        'Enquanto a Origem, FilialEmpresa, Exercicio, Periodo de Lançamento e Doc são iguais ao que já existe na colecao =>
        'significa que são lançamentos pertencentes ao mesmo documento, portanto, adiciona-os à mesma coleção.
        'Quando um desses campos for diferente, significa que o lançamento pertecente à um outro documento, e, portanto,
        'é preciso processar os lançamentos que já foram acumulados na coleção
        If colLancamento.Count <> 0 Then
            
            Set objLancamentoCol = colLancamento.Item(1)
            
            If objLancamentoCol.sOrigem <> objLancamento.sOrigem Or _
                objLancamentoCol.iFilialEmpresa <> objLancamento.iFilialEmpresa Or _
                objLancamentoCol.iExercicio <> objLancamento.iExercicio Or _
                objLancamentoCol.iPeriodoLan <> objLancamento.iPeriodoLan Or _
                objLancamentoCol.lDoc <> objLancamento.lDoc Then
        
                'Processa a colecao
                lErro = Processa_Colecao_Lancamentos_IN86(colLancamento, colRegistrosDump)
                If lErro <> SUCESSO Then gError 97889
                                
                'Cria uma nova colecao...
                Set colLancamento = New Collection
                    
            End If
        
        End If
        
        colLancamento.Add objLancamento
            
        'Busca o Próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97886
        
    Loop

    'Processa a última colecao
    lErro = Processa_Colecao_Lancamentos_IN86(colLancamento, colRegistrosDump)
    If lErro <> SUCESSO Then gError 97890
        
    Close #1
    
    Call Comando_Fechar(lComando)

    Gera_IN86_Arquivo_LancamentosCTB = SUCESSO
    
    Exit Function

Erro_Gera_IN86_Arquivo_LancamentosCTB:

    Gera_IN86_Arquivo_LancamentosCTB = gErr
    
    Select Case gErr
    
        Case 97880
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97881 To 97883, 97886
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
            
        Case 97884 'ERRO: SEM DADOS
        
        Case 97889, 97890
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150876)
    
    End Select
    
    Call Comando_Fechar(lComando)

End Function

Private Function Processa_Colecao_Lancamentos_IN86(colLancamento As Collection, ByVal colRegistrosDump As Collection) As Long

Dim lErro As Long
Dim sCtaContraPartida As String
Dim objLanc As ClassLancamento_Detalhe
Dim sRegistro As String

On Error GoTo Erro_Processa_Colecao_Lancamentos_IN86
    
    'Para cada obj na colecao
    For Each objLanc In colLancamento
        
        'Guarda a Conta de Contra Partida...
        If objLanc.iSeqContraPartida <> 0 Then
            sCtaContraPartida = colLancamento.Item(objLanc.iSeqContraPartida).sConta
        Else
            sCtaContraPartida = ""
        End If
        
        'Prepara sRegistro com a linha que vai ser inserida no arquivo
        Call Monta_Registro_Lancamento_IN86(sRegistro, objLanc, sCtaContraPartida)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Inclui o registro no arquivo
        Print #1, sRegistro
        
    Next
    
    Processa_Colecao_Lancamentos_IN86 = SUCESSO
    
    Exit Function
    
Erro_Processa_Colecao_Lancamentos_IN86:

    Processa_Colecao_Lancamentos_IN86 = gErr
    
    Select Case gErr
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150877)
    
    End Select

End Function

Private Sub Monta_Registro_Lancamento_IN86(sRegistro As String, ByVal objLancamento As ClassLancamento_Detalhe, sCtaContraPartida As String)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objLancamento RECEBE (Input) os valores que serão utilizados para montar o registro
'sCtaContraPartida RECEBE (Input) a conta de contra partida que fará parte do registro

Dim sArquivamentoTEMP As String

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objLancamento.dtData, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(objLancamento.sConta))
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CCL, " ", objLancamento.sCcl)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(sCtaContraPartida))
    sRegistro = sRegistro & FormataCpoValor(Format(Abs(objLancamento.dValor), "0.00"), STRING_IN86_VALOR)
    
    'Se o valor do lançamento for negativo => significa que é um débito
    If objLancamento.dValor < 0 Then
        
        'Inclui no registro a letra 'D' que indica que o lançamento é um débito
        sRegistro = sRegistro & "D"
    
    'Senão, ou seja, se o valor é positivo
    'significa que é um crédito
    Else
        
        'Inclui no registro a letra 'C' que indica que o lançamento é um crédito
        sRegistro = sRegistro & "C"
    End If
    
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ARQUIVAMENTO, " ", sArquivamentoTEMP)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_LANCAMENTO, " ", objLancamento.lDoc)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_HISTORICO, " ", objLancamento.sHistorico)
    
    'Os campos abaixo foram incluídos pela Forprint, ou seja, não estão previstos na
    'documentação do IN86
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIALEMPRESA, "0", objLancamento.iFilialEmpresa)
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ORIGEM, " ", objLancamento.sOrigem)
    
End Sub

Public Function Gera_IN86_Arquivo_SaldosMensais(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os SaldosMensais a partir do nome do arquivo, data inicial e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim dtDataPerIni As Date
Dim dtDataPerFim As Date
Dim iQtdPeriodo As Integer
Dim iIndice As Integer
Dim sSelecaoSQLConta As String
Dim sSelecaoSQLDadosConta As String
Dim alComando(1) As Long
Dim tSaldosMensaisIN86 As typeSaldosMensaisIN86
Dim iIndComando As Integer
Dim sRegistro As String
Dim objIN86SldMensais As New ClassIN86SldMensais

On Error GoTo Erro_Gera_IN86_Arquivo_SaldosMensais

    'Abertura dos Comandos
    For iIndComando = LBound(alComando) To UBound(alComando)
        alComando(iIndComando) = Comando_Abrir
        If alComando(iIndComando) = 0 Then gError 97920
    Next
    
    'Cria o arquivo de Saldos Mensais
    Open objIN86Arquivos.sNome For Output As #1
    
    'Obtem a quantidade de períodos (meses) do intervalo passado
    iQtdPeriodo = DateDiff("m", objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim) + 1
    
    'Guarda a DataInicial do Primeiro Período
    dtDataPerIni = objIN86Modelos.dtDataInicio
    
    'Se encontrou mais de um período para ser processado =>
    'significa que é necessário obter a data final do primeiro período
    If iQtdPeriodo > 1 Then
        dtDataPerFim = DateAdd("m", 1, "01/" & (Month(objIN86Modelos.dtDataInicio) & "/" & Year(objIN86Modelos.dtDataInicio))) - 1
    
    'Senão => significa que as datas estão no mesmo período,
    'portanto, a data final é a data passada no obj
    Else
        dtDataPerFim = objIN86Modelos.dtDataFim
    End If

    'Se a filial empresa for EMPRESA_TODA => significa que é para incluir no arquivo todos
    'os registros, exceto os pertencentes à EMPRESA_TODA
    'COALESCE é uma função da linguagem SQL que retorna o primeiro valor não nulo entre os valores passados
    If objIN86Arquivos.iFilialEmpresa = EMPRESA_TODA Then
        sSelecaoSQLConta = "SELECT FilialEmpresa, Conta, COALESCE( MAX(SldIni) + (SUM(Cre - Deb)), MAX(SldIni)) AS SaldoInicial FROM IN86SldMensais WHERE (FilialEmpresa <> ?) AND ((Data < ?) OR Data IS NULL ) GROUP BY FilialEmpresa, Conta ORDER BY Conta"
    
    'Senão, ou seja, se foi passada uma filial específica =>
    'significa que é para incluir apenas os registros da filial passada
    Else
        sSelecaoSQLConta = "SELECT FilialEmpresa, Conta, COALESCE( MAX(SldIni) + (SUM(Cre - Deb)), MAX(SldIni)) AS SaldoInicial FROM IN86SldMensais WHERE (FilialEmpresa = ?) AND ((Data < ?) OR Data IS NULL ) GROUP BY FilialEmpresa, Conta ORDER BY Conta"
    End If
    
    'Para cada mes do intervalo...
    For iIndice = 1 To iQtdPeriodo
        
        'Inicializa o Buffer
        tSaldosMensaisIN86.sConta = String(STRING_CONTA, 0)
        
        'Executa o Select para obter a conta e Saldo Inicial e...
        'Se for EMPRESA_TODA...
        If objIN86Arquivos.iFilialEmpresa = EMPRESA_TODA Then
            
            With tSaldosMensaisIN86
                'Prepara a Execucao do Comando SQL para todas as FiliaisEmpresas
                lErro = Comando_Executar(alComando(0), sSelecaoSQLConta, .iFilialEmpresa, .sConta, .dValorSaldoInicial, EMPRESA_TODA, dtDataPerIni)
            End With
            If lErro <> SUCESSO Then gError 97921
        Else 'Senão...
            With tSaldosMensaisIN86
                'Prepara a Execucao do Comando SQL para essa determinada FilialEmpresa
                lErro = Comando_Executar(alComando(0), sSelecaoSQLConta, .iFilialEmpresa, .sConta, .dValorSaldoInicial, objIN86Arquivos.iFilialEmpresa, dtDataPerIni)
            End With
            If lErro <> SUCESSO Then gError 97922
        End If
        
        'Busca a Conta e Saldo Inicial da mesma...
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97923
        
        'Se nao encontrou nenhuma conta => ERRO
        If lErro = AD_SQL_SEM_DADOS Then gError 97924
        
        'Guarda o Comando SQL para Buscar os Dados da Conta encontrada na Busca Acima
        sSelecaoSQLDadosConta = "SELECT MvDiaCta.FilialEmpresa, MIN(MvDiaCta.Data) AS DataSaldoInicial, SUM(MvDiaCta.Deb) AS TotalDebito, SUM(MvDiaCta.Cre) AS TotalCredito FROM MvDiaCta, PlanoConta WHERE MvDiaCta.Conta = PlanoConta.Conta AND MvDiaCta.FilialEmpresa = ? AND MvDiaCta.Data >= ? AND MvDiaCta.Data <= ? AND MvDiaCta.Conta = ? GROUP BY MvDiaCta.FilialEmpresa, MONTH(MvDiaCta.Data), MvDiaCta.Conta ORDER BY MvDiaCta.FilialEmpresa, MvDiaCta.Conta, MvDiaCta.DataSaldoInicial"
        
        'Enquanto for uma conta válida => gera registro
        Do While lErro <> AD_SQL_SEM_DADOS
            
            With tSaldosMensaisIN86
                'Prepara a Execucao do Comando SQL para Buscar os dados da Conta encontrada na Busca Acima
                lErro = Comando_Executar(alComando(1), sSelecaoSQLDadosConta, .iFilialEmpresa, .dtDataSldInicial, .dValorTotalDeb, .dValorTotalCred, objIN86Arquivos.iFilialEmpresa, dtDataPerIni, dtDataPerFim, .sConta)
            End With
            If lErro <> SUCESSO Then gError 97925
            
            'Busca os Dados da Conta...
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97926
            
            'Transfere do Type para o Obj
            With objIN86SldMensais
                .dSaldoInicial = tSaldosMensaisIN86.dValorSaldoInicial
                
                'Se não encontrou movimentação no Período => Guarda a Data Inicial
                'do Período em questão e Zera os Creditos/Debitos
                If lErro = AD_SQL_SEM_DADOS Then
                    .dtData = dtDataPerIni
                    tSaldosMensaisIN86.dValorTotalCred = 0
                    tSaldosMensaisIN86.dValorTotalDeb = 0
                Else
                    .dtData = tSaldosMensaisIN86.dtDataSldInicial
                End If
                
                .dTotalCreditos = tSaldosMensaisIN86.dValorTotalCred
                .dTotalDebitos = tSaldosMensaisIN86.dValorTotalDeb
                .iFilialEmpresa = tSaldosMensaisIN86.iFilialEmpresa
                .sConta = tSaldosMensaisIN86.sConta
                .dSaldoFinal = .dSaldoInicial + .dTotalCreditos - .dTotalDebitos
            
                'Se tem Saldo Inicial OU Debitos OU Creditos...
                If ((.dSaldoInicial <> 0) Or (.dTotalCreditos <> 0 Or .dTotalDebitos <> 0)) Then
                                   
                    'MONTA O REGISTRO
                    Call Monta_Registro_SldMensais_IN86(sRegistro, objIN86SldMensais)
                    
                    'Guarda o registro na coleção que será retornada à função chamadora
                    colRegistrosDump.Add sRegistro
                    
                    'Inclui o registro no arquivo
                    Print #1, sRegistro
                    
                End If
            
            End With
            
            'Busca a Próxima Conta e Saldo Inicial...
            lErro = Comando_BuscarProximo(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97928
            
        Loop
        
        'Obtem a data inicial para o próximo período
        dtDataPerIni = DateAdd("m", 1, dtDataPerIni)
    
        'Se for o último período => joga a data final
        If iIndice = iQtdPeriodo Then
            dtDataPerFim = objIN86Modelos.dtDataFim
        Else
            'Obtem a data final para o próximo período
            dtDataPerFim = DateAdd("m", 1, dtDataPerIni) - 1
        End If
    
    Next
    
    'Fecha os Comandos
    For iIndComando = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndComando))
    Next
    
    'Fecha o arquivo
    Close #1
    
    Gera_IN86_Arquivo_SaldosMensais = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_SaldosMensais:

    Gera_IN86_Arquivo_SaldosMensais = gErr
    
    Select Case gErr
    
        Case 97920
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 97921, 97922, 97923, 97925, 97926, 97928
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86SLDMENSAIS", gErr)
            
        Case 97924 'SEM DADOS
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86SLDMENSAIS_INEXISTENTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 150878)
    
    End Select
    
    'Fecha os Comandos
    For iIndComando = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndComando))
    Next
    
    'Fecha o arquivo
    Close #1

End Function

Private Sub Monta_Registro_SldMensais_IN86(sRegistro As String, ByVal objIN86SldMensais As ClassIN86SldMensais)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86SldMensais RECEBE (Input) os valores que serão utilizados para montar o registro

Dim dValorTEMP As Double

    'Formata todos os campos para o padrão IN86
    sRegistro = Format(objIN86SldMensais.dtData, "DDMMYYYY")
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(objIN86SldMensais.sConta))
    sRegistro = sRegistro & FormataCpoValor(Format(Abs(objIN86SldMensais.dSaldoInicial), "0.00"), STRING_IN86_VALOR)
    
    'Indica se é um Crédito ou Débito
    If objIN86SldMensais.dSaldoInicial < 0 Then
        'Inclui no registro a letra 'D' que indica um débito
        sRegistro = sRegistro & "D"
    Else
        'Inclui no registro a letra 'C' que indica um crédito
        sRegistro = sRegistro & "C"
    End If
    
    sRegistro = sRegistro & FormataCpoValor(Format(Abs(objIN86SldMensais.dTotalDebitos), "0.00"), STRING_IN86_VALOR)
        
    sRegistro = sRegistro & FormataCpoValor(Format(Abs(objIN86SldMensais.dTotalCreditos), "0.00"), STRING_IN86_VALOR)
    
    sRegistro = sRegistro & FormataCpoValor(Format(Abs(objIN86SldMensais.dSaldoFinal), "0.00"), STRING_IN86_VALOR)
    
    'Indica se é um Crédito ou Débito
    If objIN86SldMensais.dSaldoFinal < 0 Then
        'Inclui no registro a letra 'D' que indica um débito
        sRegistro = sRegistro & "D"
    Else
        'Inclui no registro a letra 'C' que indica um crédito
        sRegistro = sRegistro & "C"
    End If
    
    'O campo abaixo foi incluído pela Forprint, ou seja, não está previsto na
    'documentação do IN86
    sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIALEMPRESA, "0", objIN86SldMensais.iFilialEmpresa)

End Sub

'Alterada por Luiz Nogueira em 28/01/04
Public Function Gera_IN86_Arquivo_NFiscaisEmpresa(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para as NFs emitidas pela empresa (entrada e saída) a partir do nome do arquivo, data inicial e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim tNFiscaisEmpresaIN86 As typeNFiscaisEmpresaIN86
Dim objIN86NFiscaisEmpresa As New ClassIN86NFiscaisEmpresa
Dim sSelect As String
Dim lComando As Long
Dim sRegistro As String

On Error GoTo Erro_Gera_IN86_Arquivo_NFiscaisEmpresa

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97950
    
    'Inicializa os Buffers
    With tNFiscaisEmpresaIN86
    
        .sEspecieVolume = String(STRING_CAMPOSGENERICOS_DESCRICAO, 0)
        .sIDVeiculo = String(STRING_NFISCAL_PLACA, 0)
        .sIndicadorMov = String(STRING_IN86_INDICADORMOV, 0)
        .sIndicadorMovCMP = String(STRING_IN86_INDICADORMOV, 0)
        .sInscEstSubsTRB = String(STRING_IN86_INSCR_ESTADUAL, 0)
        .sObservacao = String(STRING_IN86_OBSERVACAO, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sTipoFatura = String(STRING_IN86_TIPOFATURA, 0)
        
    End With
    
    'Se a filial Empresa passada como parametro for uma filial especifica =>
    'faz o filtro por essa filial
    If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Then
        
        sSelect = "SELECT Tipo, EntradaOuSaida, EntradaSaidaCMP, ModeloArqICMS, Serie, NumNotaFiscal, " & _
                  "DataEmissao, Emitente, Destinatario, Origem, Fornecedor, FilialForn, " & _
                  "Cliente, FilialCli, DataEntrada, DataSaida, ValorProdutos, ValorDesconto, " & _
                  "ValorFrete, ValorSeguro, ValorOutrasDespesas, IPIValor, ICMSSubstValor, " & _
                  "ValorTotal, ViaTransporte, CodTransportadora, VolumeQuant, VolumeEspecie, " & _
                  "PesoBruto, PesoLiq, FreteRespons, Placa, Status, Observacao, FilialEmpresa, " & _
                  "ClasseDocCPR, NumIntDocCPR FROM IN86_NFiscais_Empresa WHERE Emitente=? AND " & _
                  "DataEmissao>=? AND DataEmissao<=? AND FilialEmpresa=? ORDER BY FilialEmpresa, " & _
                  "DataEmissao, NumNotaFiscal"
                  
        'Prepara a execucao do Comando SQL
        With tNFiscaisEmpresaIN86
            lErro = Comando_Executar(lComando, sSelect, .iTipoTipoDocInfo, .sIndicadorMov, .sIndicadorMovCMP, .iModeloDoc, .sSerie, .lNumeroDoc, .dtDataEmissao, .iEmitente, .iDestinatario, .iOrigem, .lFornecedor, .iFilialForn, .lCliente, .iFilialCliente, .dtDataEntrada, .dtDataSaida, .dValorMercadoria, .dValorTotDesconto, .dValorFrete, .dValorSeguro, .dValorOutrasDesp, .dValorTotalIPI, .dValorTotalICMSSubsTRB, .dValorTotalNF, .iViaTransporte, .iTransportador, .iQtdVolumes, .sEspecieVolume, .dPesoBruto, .dPesoLiquido, .iModalidadeFrete, .sIDVeiculo, .iStatus, .sObservacao, .iFilialEmpresa, .iClasseDocCPR, .lNumIntDocCPR, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97951
    
    Else
    'Se for a EMPRESA_TODA => Busca os registros para todas as filiais existentes
        
        sSelect = "SELECT Tipo, EntradaOuSaida, EntradaSaidaCMP, ModeloArqICMS, Serie, NumNotaFiscal, " & _
                  "DataEmissao, Emitente, Destinatario, Origem, Fornecedor, FilialForn, " & _
                  "Cliente, FilialCli, DataEntrada, DataSaida, ValorProdutos, ValorDesconto, " & _
                  "ValorFrete, ValorSeguro, ValorOutrasDespesas, IPIValor, ICMSSubstValor, " & _
                  "ValorTotal, ViaTransporte, CodTransportadora, VolumeQuant, VolumeEspecie, " & _
                  "PesoBruto, PesoLiq, FreteRespons, Placa, Status, Observacao, FilialEmpresa, " & _
                  "ClasseDocCPR, NumIntDocCPR FROM IN86_NFiscais_Empresa WHERE Emitente=? AND " & _
                  "DataEmissao>=? AND DataEmissao<=? ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal"
        
        'Prepara a execucao do Comando SQL
        With tNFiscaisEmpresaIN86
            lErro = Comando_Executar(lComando, sSelect, .sIndicadorMov, .sIndicadorMovCMP, .iModeloDoc, .sSerie, .lNumeroDoc, .dtDataEmissao, .iEmitente, .iDestinatario, .iOrigem, .lFornecedor, .iFilialForn, .lCliente, .iFilialCliente, .dtDataEntrada, .dtDataSaida, .dValorMercadoria, .dValorTotDesconto, .dValorFrete, .dValorSeguro, .dValorOutrasDesp, .dValorTotalIPI, .dValorTotalICMSSubsTRB, .dValorTotalNF, .iViaTransporte, .iTransportador, .iQtdVolumes, .sEspecieVolume, .dPesoBruto, .dPesoLiquido, .iModalidadeFrete, .sIDVeiculo, .iStatus, .sObservacao, .iFilialEmpresa, .iClasseDocCPR, .lNumIntDocCPR, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97952
    
    End If
    
    'Busca o Primeiro Registro para o Comando SQL Acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97953
    
    'Se nao encontrou nenhum registro => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 97954
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os Dados Lidos do Type para o Obj
        Call Move_IN86NFiscaisEmpresa_Obj(tNFiscaisEmpresaIN86, objIN86NFiscaisEmpresa)
    
        'Prepara o Registro para ser inserido no Arquivo
        lErro = Monta_Registro_NFiscaisEmpresa(sRegistro, objIN86NFiscaisEmpresa)
        If lErro <> SUCESSO Then gError 97956
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere no Arquivo
        Print #1, sRegistro
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97955
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_NFiscaisEmpresa = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_NFiscaisEmpresa:

    Gera_IN86_Arquivo_NFiscaisEmpresa = gErr
    
    Select Case gErr
    
        Case 97950
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97951, 97952, 97953, 97955
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAISEMPRESA", gErr)
        
        Case 97954
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAISEMPRESA_INEXISTENTE", gErr)
            
        Case 97956
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150879)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

'Alterada por Luiz Nogueira em 28/01/04
Private Sub Move_IN86NFiscaisEmpresa_Obj(tNFiscaisEmpresaIN86 As typeNFiscaisEmpresaIN86, objIN86NFiscaisEmpresa As ClassIN86NFiscaisEmpresa)
'Transfere os Dados do Type para o Obj

    With objIN86NFiscaisEmpresa
    
        .dPesoBruto = tNFiscaisEmpresaIN86.dPesoBruto
        .dPesoLiquido = tNFiscaisEmpresaIN86.dPesoLiquido
        .dtDataEmissao = tNFiscaisEmpresaIN86.dtDataEmissao
        .dtDataEntrada = tNFiscaisEmpresaIN86.dtDataEntrada
        .dtDataSaida = tNFiscaisEmpresaIN86.dtDataSaida
        .dValorFrete = tNFiscaisEmpresaIN86.dValorFrete
        .dValorMercadoria = tNFiscaisEmpresaIN86.dValorMercadoria
        .dValorOutrasDesp = tNFiscaisEmpresaIN86.dValorOutrasDesp
        .dValorSeguro = tNFiscaisEmpresaIN86.dValorSeguro
        .dValorTotalICMSSubsTRB = tNFiscaisEmpresaIN86.dValorTotalICMSSubsTRB
        .dValorTotalIPI = tNFiscaisEmpresaIN86.dValorTotalIPI
        .dValorTotalNF = tNFiscaisEmpresaIN86.dValorTotalNF
        .dValorTotDesconto = tNFiscaisEmpresaIN86.dValorTotDesconto
        .iClasseDocCPR = tNFiscaisEmpresaIN86.iClasseDocCPR
        .iDestinatario = tNFiscaisEmpresaIN86.iDestinatario
        .iEmitente = tNFiscaisEmpresaIN86.iEmitente
        .iFilialCliente = tNFiscaisEmpresaIN86.iFilialCliente
        .iFilialEmpresa = tNFiscaisEmpresaIN86.iFilialEmpresa
        .iFilialForn = tNFiscaisEmpresaIN86.iFilialForn
        .iOrigem = tNFiscaisEmpresaIN86.iOrigem
        .iQtdVolumes = tNFiscaisEmpresaIN86.iQtdVolumes
        .iStatus = tNFiscaisEmpresaIN86.iStatus
        .lCliente = tNFiscaisEmpresaIN86.lCliente
        .lFornecedor = tNFiscaisEmpresaIN86.lFornecedor
        .lNumeroDoc = tNFiscaisEmpresaIN86.lNumeroDoc
        .lNumIntDocCPR = tNFiscaisEmpresaIN86.lNumIntDocCPR
        .sEspecieVolume = tNFiscaisEmpresaIN86.sEspecieVolume
        .sIDVeiculo = tNFiscaisEmpresaIN86.sIDVeiculo
        .sIndicadorMov = tNFiscaisEmpresaIN86.sIndicadorMov
        .sIndicadorMovCMP = tNFiscaisEmpresaIN86.sIndicadorMovCMP
        .sInscEstSubsTRB = 0 '??? FALTA TRATAR A INSCRICAO ESTADUAL!!!
        .iModalidadeFrete = tNFiscaisEmpresaIN86.iModalidadeFrete
        .iModeloDoc = tNFiscaisEmpresaIN86.iModeloDoc
        If .iModeloDoc >= 9000 Then .iModeloDoc = 1 'para tratar tipos como MODREGFIS_SIMPLES_FAT em diante
        .sObservacao = tNFiscaisEmpresaIN86.sObservacao
        .sSerie = tNFiscaisEmpresaIN86.sSerie
        .sTipoFatura = 0 '???  FALTA TRATAR O TIPO DA FATURA!!!
        .iTransportador = tNFiscaisEmpresaIN86.iTransportador
        .iViaTransporte = tNFiscaisEmpresaIN86.iViaTransporte
        .iTipoTipoDocInfo = tNFiscaisEmpresaIN86.iTipoTipoDocInfo 'Incluído por Luiz Nogueira em 28/01/04
    End With

End Sub

'Alterada por Luiz Nogueira em 28/01/04
Private Function Monta_Registro_NFiscaisEmpresa(sRegistro As String, ByVal objIN86NFiscaisEmpresa As ClassIN86NFiscaisEmpresa) As Long
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86NFiscaisEmpresa RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sIndicadorMov As String
Dim sCliente As String
Dim sFornecedor As String
Dim sDescViaTransp As String
Dim lErro As Long

On Error GoTo Erro_Monta_Registro_NFiscaisEmpresa

    With objIN86NFiscaisEmpresa
        
        'Verifica o valor do campo tipo na tabela TipoDocInfo
        Select Case .iTipoTipoDocInfo
        
            'Se for nota fiscal interna de entrada ou nota fiscal externa
            Case TIPODOCINFO_TIPO_NFIE, TIPODOCINFO_TIPO_NFEXT
            
                'Preenche o indicador com o "E"
                sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", IN86_INDICADORMOV_ENTRADA)
                sIndicadorMov = IN86_INDICADORMOV_ENTRADA
            
            'Se for nota fiscal de interna de saída
            Case TIPODOCINFO_TIPO_NFIS
            
                'Preenche o indicador com o "E"
                sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", IN86_INDICADORMOV_SAIDA)
                sIndicadorMov = IN86_INDICADORMOV_SAIDA
            
            'Se for um outro caso
            Case Else
            
                'Preenche o indicador com a palavra "Erro". Usado para depurar a geração do arquivo, pois caso
                'haja alguma situação não tratada, será fácil descobrir
                sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, 4, " ", "Erro")
                sIndicadorMov = "Erro"
        
        End Select
        
'??? retirar
'        If .sIndicadorMov = "N" Then
'            sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", .sIndicadorMovCMP)
'            sIndicadorMov = .sIndicadorMovCMP
'        Else
'            sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", .sIndicadorMov)
'            sIndicadorMov = .sIndicadorMov
'        End If
    
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MODELODOC, " ", Format(.iModeloDoc, "00"))
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SERIE, " ", .sSerie)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_NUMERO_NF, "0", .lNumeroDoc)
        
        sRegistro = sRegistro & Format(.dtDataEmissao, "DDMMYYYY")
                
        'Se for uma Saída...
        If sIndicadorMov = IN86_INDICADORMOV_SAIDA Then
            
            'Se o destinatário é a própria empresa
            If .iDestinatario = DESTINATARIO_EMPRESA Then
                
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", DESTINATARIO_EMPRESA)
            
            'Se o destinatário é um cliente
            ElseIf .iDestinatario = DESTINATARIO_CLIENTE Then
                
                'Inclui o código do cliente no registro
                'sCliente = CodCliente + CodFilialCliente(Com zeros a esquerda)
                
                sCliente = CStr(.lCliente) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialCliente))
    
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sCliente)
                
            'Se o destinatário é um fornecedor
            ElseIf .iDestinatario = DESTINATARIO_FORNECEDOR Then
                
                'Inclui o código do fornecedor no registro
                'sFornecedor = CodFornecedor + CodFilialFornecedor(Com zeros a esquerda)
                
                sFornecedor = CStr(.lFornecedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialForn))
                
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sFornecedor)
            
            End If
            
            sRegistro = sRegistro & Format(.dtDataSaida, "DDMMYYYY")
            
        'Se for uma Entrada...
        ElseIf sIndicadorMov = IN86_INDICADORMOV_ENTRADA Then
        
            'Se a origem da nota é a própria empresa
            If .iOrigem = EMITENTE_EMPRESA Then
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", EMITENTE_EMPRESA)
            
            'Se a origem da nota é um cliente
            ElseIf .iOrigem = EMITENTE_CLIENTE Then
                
                'Inclui o código do cliente no registro
                'sCliente = CodCliente + CodFilialCliente(Com zeros a esquerda)
                
                sCliente = CStr(.lCliente) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialCliente))
    
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sCliente)
                
            'Se a origem da nota é um fornecedor
            ElseIf .iOrigem = EMITENTE_FORNECEDOR Then
                
                'Inclui o código do fornecedor no registro
                'sFornecedor = CodFornecedor + CodFilialFornecedor(Com zeros a esquerda)
                
                sFornecedor = CStr(.lFornecedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialForn))
                
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sFornecedor)
            
            End If
            
            sRegistro = sRegistro & Format(.dtDataEntrada, "DDMMYYYY")
        
        End If
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorMercadoria), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorMercadoria), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorTotDesconto), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotDesconto), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorFrete), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorFrete), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorSeguro), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorSeguro), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorOutrasDesp), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorOutrasDesp), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorTotalIPI), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotalIPI), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorTotalICMSSubsTRB), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotalICMSSubsTRB), "0.00"), STRING_IN86_VALOR)
        
        'sRegistro = sRegistro & FormataCpoValor(Abs(.dValorTotalNF), STRING_IN86_VALOR)
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotalNF), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_ESTADUAL, " ", .sInscEstSubsTRB)
        
        'obtem a descricao da Via de Transporte
        lErro = CF("Obtem_ViaTransp_Descricao", sDescViaTransp, .iViaTransporte)
        If lErro <> SUCESSO Then gError 97960
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_VIA_TRANSPORTE, " ", sDescViaTransp)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_TRANSPORTADOR, " ", .iTransportador)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR, "0", .iQtdVolumes)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ESPECIEVOLUME, " ", .sEspecieVolume)
        
        sRegistro = sRegistro & FormataCpoValor(.dPesoBruto, STRING_IN86_PESO - 1) & 0
        'sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dPesoBruto), "0.000"), STRING_IN86_PESO)
      
        sRegistro = sRegistro & FormataCpoValor(.dPesoLiquido, STRING_IN86_PESO - 1) & 0
        'sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dPesoLiquido), "0.000"), STRING_IN86_PESO)
        
        'Se for emitente...
        If .iModalidadeFrete = FRETE_EMITENTE Then
            sRegistro = sRegistro & TIPO_CIF
        Else
        'Se for destinatario...
            sRegistro = sRegistro & TIPO_FOB
        End If
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_IDVEICULO, " ", .sIDVeiculo)
                
        'Se o Status for CANCELADO => Concatena 'S' (SIM) ao Registro
        If .iStatus = STATUS_CANCELADO Then
            sRegistro = sRegistro & "S"
        Else 'Senão, Concatena 'N' (NAO) ao Registro
            sRegistro = sRegistro & "N"
        End If
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_TIPOFATURA, " ", .sTipoFatura)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_OBSERVACAO, " ", .sObservacao)
        
    End With

    Monta_Registro_NFiscaisEmpresa = SUCESSO
    
    Exit Function

Erro_Monta_Registro_NFiscaisEmpresa:

    Monta_Registro_NFiscaisEmpresa = gErr
    
    Select Case gErr
        
        Case 97960
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150880)
    
    End Select
    
End Function

'Alterada por Luiz Nogueira em 28/01/04
Public Function Gera_IN86_Arquivo_ItensNFiscaisEmpresa(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os itens das NFs emitidas pela empresa (entrada e saída) a partir do nome do arquivo, data inicial e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo
 
Dim lErro As Long
Dim tItensNFiscaisEmpresaIN86 As typeItensNFiscaisEmpresaIN86
Dim objIN86ItensNFEmpresa As New ClassIN86ItensNFEmpresa
Dim sSelect As String
Dim lComando As Long
Dim sRegistro As String

On Error GoTo Erro_Gera_IN86_Arquivo_ItensNFiscaisEmpresa

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97965
    
    'Inicializa os Buffers
    With tItensNFiscaisEmpresaIN86
    
        .sDescProd = String(STRING_ITEMNF_DESCRICAO, 0)
        .sIndicadorMov = String(STRING_IN86_INDICADORMOV, 0)
        .sIndicadorMovCMP = String(STRING_IN86_INDICADORMOV, 0)
        .sIPICodigo = String(STRING_IN86_IPICODIGO, 0)
        .sNaturezaOp = String(STRING_IN86_NATUREZAOP, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        
    End With
    
    'Se a filial Empresa passada como parametro for uma filial especifica =>
    'faz o filtro por essa filial
    If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Then
        
        sSelect = "SELECT Tipo, EntradaOuSaida, EntradaSaidaCMP, ModeloArqICMS, Serie, NumNotaFiscal, " & _
                  "DataEmissao, Item, Produto, Descricao, NaturezaOP, IPICodigo, Quantidade, " & _
                  "UnidadeMed, PrecoUnitario, ValorDesconto, IPITipo, IPIAliquota, IPIBaseCalculo, " & _
                  "IPIValor, OrigemMercadoria, TipoTribCST, ICMSTipo, ICMSAliquota, ICMSBase, " & _
                  "ICMSValor, ICMSSubstBase, ICMSSubstValor, TipoMovtoEstoque, FilialEmpresa, " & _
                  "ColunaNoLivro, ColunaNoLivroEntrada, ColunaNoLivroSaida " & _
                  "FROM IN86_ItensNFiscais_Empresa WHERE Emitente=? AND DataEmissao>=? AND DataEmissao<=? " & _
                  "AND FilialEmpresa=? ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal, Item"
                  
        'Prepara a execucao do Comando SQL
        With tItensNFiscaisEmpresaIN86
            lErro = Comando_Executar(lComando, sSelect, .iTipoTipoDocInfo, .sIndicadorMov, .sIndicadorMovCMP, .iModeloDoc, .sSerie, .lNumeroNF, .dtDataEmissao, .iItemNF, .sProduto, .sDescProd, .sNaturezaOp, .sIPICodigo, .dQuantidade, .sUM, .dPrecoUnitario, .dValorDesconto, .iIPITipo, .dIPIAliquota, .dIPIBaseCalculo, .dIPIValor, .iOrigemMerc, .iTipoTribCST, .iICMSTipo, .dICMSAliquota, .dICMSBase, .dICMSValor, .dICMSSubstBase, .dICMSSubstValor, .iTipoMovEstoque, .iFilialEmpresa, .iColunaNoLivro, .iColunaNoLivroEntrada, .iColunaNoLivroSaida, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97966
    
    Else
    'Se for a EMPRESA_TODA => Busca os registros para todas as filiais existentes
        
        sSelect = "SELECT Tipo, EntradaOuSaida, EntradaSaidaCMP, ModeloArqICMS, Serie, NumNotaFiscal, " & _
                  "DataEmissao, Item, Produto, Descricao, NaturezaOP, IPICodigo, Quantidade, " & _
                  "UnidadeMed, PrecoUnitario, ValorDesconto, IPITipo, IPIAliquota, IPIBaseCalculo, " & _
                  "IPIValor, OrigemMercadoria, TipoTribCST, ICMSTipo, ICMSAliquota, ICMSBase, " & _
                  "ICMSValor, ICMSSubstBase, ICMSSubstValor, TipoMovtoEstoque, FilialEmpresa, " & _
                  "ColunaNoLivro, ColunaNoLivroEntrada, ColunaNoLivroSaida " & _
                  "FROM IN86_ItensNFiscais_Empresa WHERE Emitente=? AND DataEmissao>=? AND DataEmissao<=? " & _
                  "ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal, Item"
                  
        'Prepara a execucao do Comando SQL
        With tItensNFiscaisEmpresaIN86
            lErro = Comando_Executar(lComando, sSelect, .iTipoTipoDocInfo, .sIndicadorMov, .sIndicadorMovCMP, .iModeloDoc, .sSerie, .lNumeroNF, .dtDataEmissao, .iItemNF, .sProduto, .sDescProd, .sNaturezaOp, .sIPICodigo, .dQuantidade, .sUM, .dPrecoUnitario, .dValorDesconto, .iIPITipo, .dIPIAliquota, .dIPIBaseCalculo, .dIPIValor, .iOrigemMerc, .iTipoTribCST, .iICMSTipo, .dICMSAliquota, .dICMSBase, .dICMSValor, .dICMSSubstBase, .dICMSSubstValor, .iTipoMovEstoque, .iFilialEmpresa, .iColunaNoLivro, .iColunaNoLivroEntrada, .iColunaNoLivroSaida, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97967
    
    End If
    
    'Busca o Primeiro Registro para o Comando SQL Acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97968
    
    'Se nao encontrou nenhum registro => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 97969
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os Dados Lidos do Type para o Obj
        Call Move_IN86ItensNFiscaisEmpresa_Obj(tItensNFiscaisEmpresaIN86, objIN86ItensNFEmpresa)
    
        'Prepara o Registro para ser inserido no Arquivo
        Call Monta_Registro_ItensNFiscaisEmpresa(sRegistro, objIN86ItensNFEmpresa)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere no Arquivo
        Print #1, sRegistro
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97971
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_ItensNFiscaisEmpresa = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_ItensNFiscaisEmpresa:

    Gera_IN86_Arquivo_ItensNFiscaisEmpresa = gErr
    
    Select Case gErr
    
        Case 97965
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97966, 97967, 97968, 97971
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAISEMPRESA", gErr)
        
        Case 97969
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNFISCAISEMPRESA_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150881)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

'Alterada por Luiz Nogueira em 28/01/04
Private Sub Move_IN86ItensNFiscaisEmpresa_Obj(tItensNFiscaisEmpresaIN86 As typeItensNFiscaisEmpresaIN86, objIN86ItensNFEmpresa As ClassIN86ItensNFEmpresa)
'Transfere o Conteúdo do type para o Obj

    With tItensNFiscaisEmpresaIN86
    
        objIN86ItensNFEmpresa.dICMSAliquota = .dICMSAliquota
        objIN86ItensNFEmpresa.dICMSBase = .dICMSBase
        objIN86ItensNFEmpresa.dICMSSubstBase = .dICMSSubstBase
        objIN86ItensNFEmpresa.dICMSSubstValor = .dICMSSubstValor
        objIN86ItensNFEmpresa.dICMSValor = .dICMSValor
        objIN86ItensNFEmpresa.dIPIAliquota = .dIPIAliquota
        objIN86ItensNFEmpresa.dIPIBaseCalculo = .dIPIBaseCalculo
        objIN86ItensNFEmpresa.dIPIValor = .dIPIValor
        objIN86ItensNFEmpresa.dPrecoUnitario = .dPrecoUnitario
        objIN86ItensNFEmpresa.dQuantidade = .dQuantidade
        objIN86ItensNFEmpresa.dtDataEmissao = .dtDataEmissao
        objIN86ItensNFEmpresa.dValorDesconto = .dValorDesconto
        objIN86ItensNFEmpresa.iFilialEmpresa = .iFilialEmpresa
        objIN86ItensNFEmpresa.iICMSTipo = .iICMSTipo
        objIN86ItensNFEmpresa.iIPITipo = .iIPITipo
        objIN86ItensNFEmpresa.iItemNF = .iItemNF
        objIN86ItensNFEmpresa.iModeloDoc = .iModeloDoc
        If objIN86ItensNFEmpresa.iModeloDoc >= 9000 Then objIN86ItensNFEmpresa.iModeloDoc = 1 'para tratar tipos como MODREGFIS_SIMPLES_FAT em diante
        objIN86ItensNFEmpresa.iOrigemMerc = .iOrigemMerc
        objIN86ItensNFEmpresa.iTipoMovEstoque = .iTipoMovEstoque
        objIN86ItensNFEmpresa.iTipoTribCST = .iTipoTribCST
        objIN86ItensNFEmpresa.lNumeroNF = .lNumeroNF
        objIN86ItensNFEmpresa.sDescProd = .sDescProd
        objIN86ItensNFEmpresa.sIndicadorMov = .sIndicadorMov
        objIN86ItensNFEmpresa.sIndicadorMovCMP = .sIndicadorMovCMP
        objIN86ItensNFEmpresa.sIPICodigo = .sIPICodigo
        objIN86ItensNFEmpresa.sNaturezaOp = .sNaturezaOp
        objIN86ItensNFEmpresa.sProduto = .sProduto
        objIN86ItensNFEmpresa.sSerie = .sSerie
        objIN86ItensNFEmpresa.sUM = .sUM
        objIN86ItensNFEmpresa.lFornecedor = .lFornecedor
        objIN86ItensNFEmpresa.iFilialForn = .iFilialForn
        objIN86ItensNFEmpresa.lCliente = .lCliente
        objIN86ItensNFEmpresa.iFilialCli = .iFilialCli
        objIN86ItensNFEmpresa.iOrigem = .iOrigem
        objIN86ItensNFEmpresa.iColunaNoLivro = .iColunaNoLivro
        objIN86ItensNFEmpresa.iColunaNoLivroEntrada = .iColunaNoLivroEntrada
        objIN86ItensNFEmpresa.iColunaNoLivroSaida = .iColunaNoLivroSaida
        objIN86ItensNFEmpresa.iTipoTipoDocInfo = .iTipoTipoDocInfo 'Incluído por Luiz Nogueira em 28/01/04

    End With

End Sub

'Alterada por Luiz Nogueira em 28/01/04
Private Sub Monta_Registro_ItensNFiscaisEmpresa(sRegistro As String, objIN86ItensNFEmpresa As ClassIN86ItensNFEmpresa)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86NFiscaisEmpresa RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sIndicadorMov As String
Dim sS1 As String
Dim sS2 As String
Dim iVirgula As Integer
Dim sStrTEMP As String

    With objIN86ItensNFEmpresa
        
        'Verifica o valor do campo tipo na tabela TipoDocInfo
        Select Case .iTipoTipoDocInfo
        
            'Se for nota fiscal interna de entrada ou nota fiscal externa
            Case TIPODOCINFO_TIPO_NFIE, TIPODOCINFO_TIPO_NFEXT
            
                'Preenche o indicador com o "E"
                sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", IN86_INDICADORMOV_ENTRADA)
                sIndicadorMov = IN86_INDICADORMOV_ENTRADA
            
            'Se for nota fiscal de interna de saída
            Case TIPODOCINFO_TIPO_NFIS
            
                'Preenche o indicador com o "E"
                sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", IN86_INDICADORMOV_SAIDA)
                sIndicadorMov = IN86_INDICADORMOV_SAIDA
            
            'Se for um outro caso
            Case Else
            
                'Preenche o indicador com a palavra "Erro". Usado para depurar a geração do arquivo, pois caso
                'haja alguma situação não tratada, será fácil descobrir
                sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, 4, " ", "Erro")
                sIndicadorMov = "Erro"
        
        End Select
        
'??? retirar
'        'Define o Tipo do Movimento
'        If .sIndicadorMov = "N" Then
'            sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", .sIndicadorMovCMP)
'            sIndicadorMov = .sIndicadorMovCMP
'        Else
'            sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INDICADORMOV, " ", .sIndicadorMov)
'            sIndicadorMov = .sIndicadorMov
'        End If
    
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MODELODOC, " ", Format(.iModeloDoc, "00"))
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SERIE, " ", .sSerie)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_NUMERO_NF, "0", .lNumeroNF)
        
        sRegistro = sRegistro & Format(.dtDataEmissao, "DDMMYYYY")
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_ITEMNF, "0", .iItemNF)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", .sProduto)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESCRICAO, " ", .sDescProd)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NATUREZAOP, " ", .sNaturezaOp)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NATUREZAOP, " ", .sNaturezaOp)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_IPICODIGO, " ", .sIPICodigo)
        
        'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQuantidade, "0.000"))
        
        'Encontra na Quantidade Retornada, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando a Quantidade Formatada (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UM, " ", .sUM)
        
        'Formata o PrecoUnitario com o tamanho da STRING_IN86_VALOR, sendo as 4 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dPrecoUnitario, "0.0000"))
        
        'Encontra no PrecoUnitario Retornado, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando o PrecoUnitario Formatado (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        'Formata o PrecoTotalItem  com o tamanho da STRING_IN86_VALOR, sendo as 2 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(((.dPrecoUnitario * .dQuantidade) - .dValorDesconto), "0.00"))
        
        'Encontra no PrecoTotalItem Retornado, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando o PrecoTotalItem Formatado (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & FormataCpoValor(.dValorDesconto, STRING_IN86_VALOR)
        
        '??? Luiz: posteriormente, alterar a forma de encontrar o indicador de ICMS
        '??? para funcionar de forma análoga à Livros Fiscais
        'Tipo da Tributacao do IPI
        sRegistro = sRegistro & .iColunaNoLivro
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dIPIAliquota * 100, "0.00"), STRING_IN86_ALIQUOTA)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dIPIBaseCalculo, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dIPIValor, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SITUACAO_TRB, "0", CStr(.iOrigemMerc) & CStr(.iTipoTribCST))
        
        '??? Luiz: posteriormente, alterar a forma de encontrar o indicador de ICMS
        '??? para funcionar de forma análoga à Livros Fiscais
        If sIndicadorMov = IN86_INDICADORMOV_SAIDA Then
            sRegistro = sRegistro & .iColunaNoLivroSaida
        ElseIf sIndicadorMov = IN86_INDICADORMOV_ENTRADA Then
            sRegistro = sRegistro & .iColunaNoLivroEntrada
        End If
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSAliquota * 100, "0.00"), STRING_IN86_ALIQUOTA)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSBase, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSValor, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSSubstBase, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSSubstValor, "0.00"), STRING_IN86_VALOR)
        
        If .iTipoMovEstoque <> 0 Then
            sRegistro = sRegistro & "S"
        Else
            sRegistro = sRegistro & "N"
        End If
                
    End With

End Sub

Public Function Gera_IN86_Arquivo_NFiscaisTerceiros(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para as NFs emitidas por Terceiros (entrada) a partir do nome do arquivo, data inicial e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim tNFiscaisTerceirosIN86 As typeNFiscaisEmpresaIN86
Dim objIN86NFiscaisTerceiros As New ClassIN86NFiscaisEmpresa
Dim sSelect As String
Dim lComando As Long
Dim sRegistro As String

On Error GoTo Erro_Gera_IN86_Arquivo_NFiscaisTerceiros

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97980
    
    'Inicializa os Buffers
    With tNFiscaisTerceirosIN86
    
        .sEspecieVolume = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
        .sIDVeiculo = String(STRING_IN86_IDVEICULO, 0)
        .sIndicadorMov = String(STRING_IN86_INDICADORMOV, 0)
        .sIndicadorMovCMP = String(STRING_IN86_INDICADORMOV, 0)
        .sInscEstSubsTRB = String(STRING_IN86_INSCR_ESTADUAL, 0)
        .sObservacao = String(STRING_IN86_OBSERVACAO, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sTipoFatura = String(STRING_IN86_TIPOFATURA, 0)
        
    End With
    
    'Se a filial Empresa passada como parametro for uma filial especifica =>
    'faz o filtro por essa filial
    If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Then
        
        sSelect = "SELECT ModeloArqICMS, Serie, NumNotaFiscal, " & _
                  "DataEmissao, Emitente, Origem, DataEntrada, ValorProdutos, " & _
                  "ValorDesconto, ValorFrete, ValorSeguro, ValorOutrasDespesas, IPIValor, " & _
                  "ICMSSubstValor, ValorTotal, Observacao, Fornecedor, FilialForn, Cliente, FilialCli " & _
                  "FROM IN86_NFiscais_Empresa WHERE Emitente<>? AND DataEmissao>=? AND DataEmissao<=? " & _
                  "AND FilialEmpresa=? ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal"
                  
        'Prepara a execucao do Comando SQL
        With tNFiscaisTerceirosIN86
            lErro = Comando_Executar(lComando, sSelect, .iModeloDoc, .sSerie, .lNumeroDoc, .dtDataEmissao, .iEmitente, .iOrigem, .dtDataEntrada, .dValorMercadoria, .dValorTotDesconto, .dValorFrete, .dValorSeguro, .dValorOutrasDesp, .dValorTotalIPI, .dValorTotalICMSSubsTRB, .dValorTotalNF, .sObservacao, .lFornecedor, .iFilialForn, .lCliente, .iFilialCliente, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97981
    
    Else
    'Se for a EMPRESA_TODA => Busca os registros para todas as filiais existentes
        
        sSelect = "SELECT ModeloArqICMS, Serie, NumNotaFiscal, " & _
                  "DataEmissao, Emitente, Origem, DataEntrada, ValorProdutos, " & _
                  "ValorDesconto, ValorFrete, ValorSeguro, ValorOutrasDespesas, IPIValor, " & _
                  "ICMSSubstValor, ValorTotal, Observacao, Fornecedor, FilialForn, Cliente, FilialCli " & _
                  "FROM IN86_NFiscais_Empresa WHERE Emitente<>? AND DataEmissao>=? AND DataEmissao<=? " & _
                  "ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal"
                  
        'Prepara a execucao do Comando SQL
        With tNFiscaisTerceirosIN86
            lErro = Comando_Executar(lComando, sSelect, .iModeloDoc, .sSerie, .lNumeroDoc, .dtDataEmissao, .iEmitente, .iOrigem, .dtDataEntrada, .dValorMercadoria, .dValorTotDesconto, .dValorFrete, .dValorSeguro, .dValorOutrasDesp, .dValorTotalIPI, .dValorTotalICMSSubsTRB, .dValorTotalNF, .sObservacao, .lFornecedor, .iFilialForn, .lCliente, .iFilialCliente, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97982
    
    End If
    
    'Busca o Primeiro Registro para o Comando SQL Acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97983
    
    'Se nao encontrou nenhum registro => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 97984
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os Dados Lidos do Type para o Obj
        Call Move_IN86NFiscaisEmpresa_Obj(tNFiscaisTerceirosIN86, objIN86NFiscaisTerceiros)
    
        'Prepara o Registro para ser inserido no Arquivo
        Call Monta_Registro_NFiscaisTerceiros(sRegistro, objIN86NFiscaisTerceiros)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere no Arquivo
        Print #1, sRegistro
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97986
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_NFiscaisTerceiros = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_NFiscaisTerceiros:

    Gera_IN86_Arquivo_NFiscaisTerceiros = gErr
    
    Select Case gErr
    
        Case 97980
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97981, 97982, 97983, 97986
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAISTERCEIROS", gErr)
        
        Case 97984
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAISTERCEIROS_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150882)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Sub Monta_Registro_NFiscaisTerceiros(sRegistro As String, ByVal objIN86NFiscaisTerceiros As ClassIN86NFiscaisEmpresa)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86NFiscaisEmpresa RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sCliente As String
Dim sFornecedor As String

    With objIN86NFiscaisTerceiros
        
        sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MODELODOC, " ", Format(.iModeloDoc, "00"))
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SERIE, " ", .sSerie)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_NUMERO_NF, "0", .lNumeroDoc)
        
        sRegistro = sRegistro & Format(.dtDataEmissao, "DDMMYYYY")
                
        'Se for um Cliente...
        If .iOrigem = EMITENTE_CLIENTE Then
        
            'Inclui o código do cliente no registro
            'sCliente = CodCliente + CodFilialCliente(Com zeros a esquerda)
            
            sCliente = CStr(.lCliente) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialCliente))

            sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sCliente)
            
        'Se a origem da nota é um fornecedor
        ElseIf .iOrigem = EMITENTE_FORNECEDOR Then
                
                'Inclui o código do fornecedor no registro
                'sFornecedor = CodFornecedor + CodFilialFornecedor(Com zeros a esquerda)
                
                sFornecedor = CStr(.lFornecedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialForn))
                
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sFornecedor)
            
        End If
        
        sRegistro = sRegistro & Format(.dtDataEntrada, "DDMMYYYY")
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorMercadoria), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotDesconto), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorFrete), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorSeguro), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorOutrasDesp), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotalIPI), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotalICMSSubsTRB), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotalNF), "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_INSCR_ESTADUAL, " ", .sInscEstSubsTRB)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_TIPOFATURA, " ", .sTipoFatura)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_OBSERVACAO, " ", .sObservacao)
        
    End With

End Sub

Public Function Gera_IN86_Arquivo_ItensNFiscaisTerceiros(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os itens das NFs emitidas por terceiros (entrada) a partir do nome do arquivo, data inicial e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim tItensNFiscaisTerceirosIN86 As typeItensNFiscaisEmpresaIN86
Dim objIN86ItensNFTerceiros As New ClassIN86ItensNFEmpresa
Dim sSelect As String
Dim lComando As Long
Dim sRegistro As String

On Error GoTo Erro_Gera_IN86_Arquivo_ItensNFiscaisTerceiros

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97987
    
    'Inicializa os Buffers
    With tItensNFiscaisTerceirosIN86
    
        .sDescProd = String(STRING_ITEMNF_DESCRICAO, 0)
        .sIndicadorMov = String(STRING_IN86_INDICADORMOV, 0)
        .sIndicadorMovCMP = String(STRING_IN86_INDICADORMOV, 0)
        .sIPICodigo = String(STRING_IN86_IPICODIGO, 0)
        .sNaturezaOp = String(STRING_IN86_NATUREZAOP, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        
    End With
    
    'Se a filial Empresa passada como parametro for uma filial especifica =>
    'faz o filtro por essa filial
    If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Then
        
        sSelect = "SELECT ModeloArqICMS, Serie, NumNotaFiscal, DataEmissao, Origem, " & _
                  "Fornecedor, FilialForn, Cliente, FilialCli, Item, Produto, Descricao, " & _
                  "NaturezaOP, IPICodigo, Quantidade, UnidadeMed, PrecoUnitario, ValorDesconto, " & _
                  "IPITipo, IPIAliquota, IPIBaseCalculo, IPIValor, OrigemMercadoria, TipoTribCST, " & _
                  "ICMSTipo, ICMSAliquota, ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, " & _
                  "TipoMovtoEstoque, ColunaNoLivro, ColunaNoLivroEntrada " & _
                  "FROM IN86_ItensNFiscais_Empresa WHERE Emitente<>? AND " & _
                  "DataEmissao>=? AND DataEmissao<=? AND FilialEmpresa=? " & _
                  "ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal, Item"
        
        'Prepara a execucao do Comando SQL
        With tItensNFiscaisTerceirosIN86
            lErro = Comando_Executar(lComando, sSelect, .iModeloDoc, .sSerie, .lNumeroNF, .dtDataEmissao, .iOrigem, .lFornecedor, .iFilialForn, .lCliente, .iFilialCli, .iItemNF, .sProduto, .sDescProd, .sNaturezaOp, .sIPICodigo, .dQuantidade, .sUM, .dPrecoUnitario, .dValorDesconto, .iIPITipo, .dIPIAliquota, .dIPIBaseCalculo, .dIPIValor, .iOrigemMerc, .iTipoTribCST, .iICMSTipo, .dICMSAliquota, .dICMSBase, .dICMSValor, .dICMSSubstBase, .dICMSSubstValor, .iTipoMovEstoque, .iColunaNoLivro, .iColunaNoLivroEntrada, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97988
    
    Else
    'Se for a EMPRESA_TODA => Busca os registros para todas as filiais existentes
        
        sSelect = "SELECT ModeloArqICMS, Serie, NumNotaFiscal, DataEmissao, Origem, " & _
                  "Fornecedor, FilialForn, Cliente, FilialCli, Item, Produto, Descricao, " & _
                  "NaturezaOP, IPICodigo, Quantidade, UnidadeMed, PrecoUnitario, ValorDesconto, " & _
                  "IPITipo, IPIAliquota, IPIBaseCalculo, IPIValor, OrigemMercadoria, TipoTribCST, " & _
                  "ICMSTipo, ICMSAliquota, ICMSBase, ICMSValor, ICMSSubstBase, ICMSSubstValor, " & _
                  "TipoMovtoEstoque, ColunaNoLivro, ColunaNoLivroEntrada " & _
                  "FROM IN86_ItensNFiscais_Empresa WHERE Emitente<>? AND " & _
                  "DataEmissao>=? AND DataEmissao<=? AND FilialEmpresa=? " & _
                  "ORDER BY FilialEmpresa, DataEmissao, NumNotaFiscal, Item"
        
        'Prepara a execucao do Comando SQL
        With tItensNFiscaisTerceirosIN86
            lErro = Comando_Executar(lComando, sSelect, .iModeloDoc, .sSerie, .lNumeroNF, .dtDataEmissao, .iOrigem, .lFornecedor, .iFilialForn, .lCliente, .iFilialCli, .iItemNF, .sProduto, .sDescProd, .sNaturezaOp, .sIPICodigo, .dQuantidade, .sUM, .dPrecoUnitario, .dValorDesconto, .iIPITipo, .dIPIAliquota, .dIPIBaseCalculo, .dIPIValor, .iOrigemMerc, .iTipoTribCST, .iICMSTipo, .dICMSAliquota, .dICMSBase, .dICMSValor, .dICMSSubstBase, .dICMSSubstValor, .iTipoMovEstoque, .iColunaNoLivro, .iColunaNoLivroEntrada, EMITENTE_EMPRESA, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 97989
    
    End If
    
    'Busca o Primeiro Registro para o Comando SQL Acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97990
    
    'Se nao encontrou nenhum registro => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 97991
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os Dados Lidos do Type para o Obj
        Call Move_IN86ItensNFiscaisEmpresa_Obj(tItensNFiscaisTerceirosIN86, objIN86ItensNFTerceiros)
    
        'Prepara o Registro para ser inserido no Arquivo
        Call Monta_Registro_ItensNFiscaisTerceiros(sRegistro, objIN86ItensNFTerceiros)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere no Arquivo
        Print #1, sRegistro
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97993
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_ItensNFiscaisTerceiros = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_ItensNFiscaisTerceiros:

    Gera_IN86_Arquivo_ItensNFiscaisTerceiros = gErr
    
    Select Case gErr
    
        Case 97987
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97988, 97989, 97990, 97993
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAISTERCEIROS", gErr)
        
        Case 97991
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNFISCAISTERCEIROS_INEXISTENTE", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150883)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Sub Monta_Registro_ItensNFiscaisTerceiros(sRegistro As String, objIN86ItensNFTerceiros As ClassIN86ItensNFEmpresa)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86NFiscaisTerceiros RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sS1 As String
Dim sS2 As String
Dim iVirgula As Integer
Dim sStrTEMP As String
Dim sCliente As String
Dim sFornecedor As String

    With objIN86ItensNFTerceiros
        
        sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_MODELODOC, " ", Format(.iModeloDoc, "00"))
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SERIE, " ", .sSerie)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_NUMERO_NF, "0", .lNumeroNF)
        
        sRegistro = sRegistro & Format(.dtDataEmissao, "DDMMYYYY")
        
        'Se for um Cliente...
        If .iOrigem = EMITENTE_CLIENTE Then
        
            'Inclui o código do cliente no registro
            'sCliente = CodCliente + CodFilialCliente(Com zeros a esquerda)
            
            sCliente = CStr(.lCliente) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialCli))

            sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sCliente)
            
        'Se a origem da nota é um fornecedor
        ElseIf .iOrigem = EMITENTE_FORNECEDOR Then
                
                'Inclui o código do fornecedor no registro
                'sFornecedor = CodFornecedor + CodFilialFornecedor(Com zeros a esquerda)
                
                sFornecedor = CStr(.lFornecedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilialForn))
                
                sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sFornecedor)
            
        End If
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_ITEMNF, "0", .iItemNF)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", .sProduto)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESCRICAO, " ", .sDescProd)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NATUREZAOP, " ", .sNaturezaOp)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NATUREZAOP, " ", .sNaturezaOp)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_IPICODIGO, " ", .sIPICodigo)
        
        'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQuantidade, "0.000"))
        
        'Encontra na Quantidade Retornada, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando a Quantidade Formatada (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UM, " ", .sUM)
        
        'Formata o PrecoUnitario com o tamanho da STRING_IN86_VALOR, sendo as 4 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dPrecoUnitario, "0.0000"))
        
        'Encontra no PrecoUnitario Retornado, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando o PrecoUnitario Formatado (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        'Formata o PrecoTotalItem  com o tamanho da STRING_IN86_VALOR, sendo as 2 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(((.dPrecoUnitario * .dQuantidade) - .dValorDesconto), "0.00"))
        
        'Encontra no PrecoTotalItem Retornado, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando o PrecoTotalItem Formatado (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & FormataCpoValor(.dValorDesconto, STRING_IN86_VALOR)
        
        'Tipo de Tributação do IPI
        sRegistro = sRegistro & .iColunaNoLivro
    
        sRegistro = sRegistro & FormataCpoValor(Format(.dIPIAliquota * 100, "0.00"), STRING_IN86_ALIQUOTA)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dIPIBaseCalculo, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dIPIValor, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SITUACAO_TRB, "0", CStr(.iOrigemMerc) & CStr(.iTipoTribCST))
        
        'Tipo de Tributacao do ICMS
        sRegistro = sRegistro & .iColunaNoLivroEntrada
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSAliquota * 100, "0.00"), STRING_IN86_ALIQUOTA)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSBase, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSValor, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSSubstBase, "0.00"), STRING_IN86_VALOR)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dICMSSubstValor, "0.00"), STRING_IN86_VALOR)
        
        If .iTipoMovEstoque <> 0 Then
            sRegistro = sRegistro & "S"
        Else
            sRegistro = sRegistro & "N"
        End If
                
    End With

End Sub

Public Function Gera_IN86_Arquivo_MovEstoque(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Movimentos de Estoque a partir do nome do arquivo, data inicial e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim tIN86MovEstoque As typeIN86MovEstoque
Dim objIN86MovEstoque As New ClassIN86MovEstoque
Dim sSelect As String
Dim lComando As Long
Dim sRegistro As String
Dim sNumDocOrigem As String
Dim sSerie As String

On Error GoTo Erro_Gera_IN86_Arquivo_MovEstoque

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101001
    
    'Inicializa os Buffers
    With tIN86MovEstoque
    
        .sProduto = String(STRING_PRODUTO, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sHistorico = String(STRING_IN86_HISTORICOMOVEST, 0)
        .sNumeroDoc = String(STRING_IN86_NUMERO_DOC, 0)
        .sSiglaDoc = String(STRING_IN86_SIGLA_DOC, 0)
        .sEntradaOuSaida = String(STRING_IN86_INDICADORMOV, 0)
        .sEntradaSaidaCMP = String(STRING_IN86_INDICADORMOV, 0)
        
    End With
    
    'Se a filial Empresa passada como parametro for uma filial especifica =>
    'faz o filtro por essa filial
    If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Then
        
        sSelect = "SELECT MovimentoEstoque.Quantidade, MovimentoEstoque.FilialEmpresa, MovimentoEstoque.Produto, MovimentoEstoque.TipoNumIntDocOrigem, MovimentoEstoque.NumIntDocOrigem, " & _
                  "TiposMovimentoEstoque.Sigla, MovimentoEstoque.Codigo, MovimentoEstoque.Data, MovimentoEstoque.SiglaUM, " & _
                  "TiposMovimentoEstoque.EntradaOuSaida, TiposMovimentoEstoque.EntradaSaidaCMP , MovimentoEstoque.Custo " & _
                  "FROM MovimentoEstoque, TiposMovimentoEstoque WHERE TiposMovimentoEstoque.Codigo = MovimentoEstoque.TipoMov " & _
                  "AND MovimentoEstoque.Data >= ? AND MovimentoEstoque.Data <= ? AND MovimentoEstoque.FilialEmpresa = ? AND TiposMovimentoEstoque.AtualizaSoLote = 0 " & _
                  "ORDER BY MovimentoEstoque.FilialEmpresa, MovimentoEstoque.Data, MovimentoEstoque.Hora"
        
        'Prepara a execucao do Comando SQL
        With tIN86MovEstoque
            lErro = Comando_Executar(lComando, sSelect, .dQuantidade, .iFilialEmpresa, .sProduto, .iTipoNumIntDocOrigem, .lNumIntDocOrigem, .sSiglaDoc, .lCodigo, .dtDataMov, .sSiglaUM, .sEntradaOuSaida, .sEntradaSaidaCMP, .dValorTotal, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 101002
    
    Else
    'Se for a EMPRESA_TODA => Busca os registros para todas as filiais existentes
        
        sSelect = "SELECT MovimentoEstoque.Quantidade, MovimentoEstoque.FilialEmpresa, MovimentoEstoque.Produto, MovimentoEstoque.TipoNumIntDocOrigem, MovimentoEstoque.NumIntDocOrigem, " & _
                  "TiposMovimentoEstoque.Sigla, MovimentoEstoque.Codigo, MovimentoEstoque.Data, MovimentoEstoque.SiglaUM, " & _
                  "TiposMovimentoEstoque.EntradaOuSaida, TiposMovimentoEstoque.EntradaSaidaCMP , MovimentoEstoque.Custo " & _
                  "FROM MovimentoEstoque, TiposMovimentoEstoque WHERE TiposMovimentoEstoque.Codigo = MovimentoEstoque.TipoMov " & _
                  "AND MovimentoEstoque.Data >= ? AND MovimentoEstoque.Data <= ? AND TiposMovimentoEstoque.AtualizaSoLote = 0 " & _
                  "ORDER BY MovimentoEstoque.FilialEmpresa, MovimentoEstoque.Data, MovimentoEstoque.Hora"
        
        'Prepara a execucao do Comando SQL
        With tIN86MovEstoque
            lErro = Comando_Executar(lComando, sSelect, .dQuantidade, .iFilialEmpresa, .sProduto, .iTipoNumIntDocOrigem, .lNumIntDocOrigem, .sSiglaDoc, .lCodigo, .dtDataMov, .sSiglaUM, .sEntradaOuSaida, .sEntradaSaidaCMP, .dValorTotal, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 101003
    
    End If
    
    'Busca o Primeiro Registro para o Comando SQL Acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101004
    
    'Se nao encontrou nenhum registro => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 101005
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Acha o NumDocOrigem e a Série do documento
        lErro = MovEstoque_Le_NumDocOrigem(objIN86MovEstoque.lNumIntDocOrigem, objIN86MovEstoque.iTipoNumIntDocOrigem, sNumDocOrigem, sSerie)
        If lErro <> SUCESSO Then gError 101023
        
        tIN86MovEstoque.sSerie = sSerie
        tIN86MovEstoque.sNumeroDoc = sNumDocOrigem
        
        'Transfere os Dados Lidos do Type para o Obj
        Call Move_IN86MovEstoque_Obj(tIN86MovEstoque, objIN86MovEstoque)
    
        'Prepara o Registro para ser inserido no Arquivo
        Call Monta_Registro_MovEstoque(sRegistro, objIN86MovEstoque)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere no Arquivo
        Print #1, sRegistro
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101006
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_MovEstoque = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_MovEstoque:

    Gera_IN86_Arquivo_MovEstoque = gErr
    
    Select Case gErr
    
        Case 101023
        
        Case 101001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 101002, 101003, 101004, 101006
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86MOVESTOQUE", gErr)
        
        Case 101005
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MOVESTOQUE_NAO_CADASTRADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150884)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Sub Move_IN86MovEstoque_Obj(tIN86MovEstoque As typeIN86MovEstoque, objIN86MovEstoque As ClassIN86MovEstoque)
'Transfere os dados do Obj para o Type

    With tIN86MovEstoque
    
        If .dQuantidade = 0 Then
            objIN86MovEstoque.dCustoUnitario = .dValorTotal
        Else
            objIN86MovEstoque.dCustoUnitario = (.dValorTotal / .dQuantidade)
        End If
        
        objIN86MovEstoque.dQuantidade = .dQuantidade
        objIN86MovEstoque.dtDataMov = .dtDataMov
        objIN86MovEstoque.dValorTotal = .dValorTotal
        objIN86MovEstoque.iFilialEmpresa = .iFilialEmpresa
        objIN86MovEstoque.iTipoNumIntDocOrigem = .iTipoNumIntDocOrigem
        objIN86MovEstoque.lCodigo = .lCodigo
        objIN86MovEstoque.lNumIntDocOrigem = .lNumIntDocOrigem
        objIN86MovEstoque.sEntradaOuSaida = .sEntradaOuSaida
        objIN86MovEstoque.sEntradaSaidaCMP = .sEntradaSaidaCMP
        
        'Como esse dado não existe no BD => atribui uma string "zerada"
        objIN86MovEstoque.sHistorico = ""
        
        objIN86MovEstoque.sNumeroDoc = .sNumeroDoc
        objIN86MovEstoque.sProduto = .sProduto
        objIN86MovEstoque.sSerie = .sSerie
        objIN86MovEstoque.sSiglaDoc = .sSiglaDoc
        objIN86MovEstoque.sSiglaUM = .sSiglaUM
        
    End With

End Sub

Private Sub Monta_Registro_MovEstoque(sRegistro As String, objIN86MovEstoque As ClassIN86MovEstoque)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86MovEstoque RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sStrTEMP As String
Dim sS1 As String
Dim sS2 As String
Dim iVirgula As Integer

    With objIN86MovEstoque
    
        sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", .sProduto)
        
        If .iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Or .iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_NFISCAL Then
            sRegistro = sRegistro & "F"
        Else
            sRegistro = sRegistro & "I"
        End If
         
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SIGLA_DOC, " ", .sSiglaDoc)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SERIE, " ", .sSerie)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NUMERO_DOC, " ", .sNumeroDoc)
        
        sRegistro = sRegistro & Format(.dtDataMov, "DDMMYYYY")
        
        'Apenas para constar, pois não está sendo lido do BD, já q essa informação é inexistente
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_HISTORICO, " ", .sHistorico)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UM, " ", .sSiglaUM)
        
        'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQuantidade, "0.000"))
        
        'Encontra na Quantidade Retornada, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando a Quantidade Formatada (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        If .sEntradaOuSaida = "N" Then
            sRegistro = sRegistro & .sEntradaSaidaCMP
        Else
            sRegistro = sRegistro & .sEntradaOuSaida
        End If
        
        'Formata o PrecoUnitario com o tamanho da STRING_IN86_VALOR, sendo as 4 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dCustoUnitario, "0.0000"))
        
        'Encontra no Valor Total Retornado, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando o Valor Total Formatado (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & FormataCpoValor(Format(Abs(.dValorTotal), "0.00"), STRING_IN86_VALOR)
        
    End With
        
End Sub

Private Function MovEstoque_Le_NumIntNF_ItemNF(lNumIntDocOrigem As Long, lNumIntNF As Long) As Long
'Le o numero interno da nf a partir do numintdoc passado como parametro
'lNumIntNF Recebe(INPUT)o NumIntDoc do ItemNF que será lido do BD
'lNumIntDocOrigem Retorna(OUTPUT) o NumIntNF do ItemNF que será lido

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MovEstoque_Le_NumIntNF_ItemNF

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101010
    
    'Prepara a Execução do Comando SQL
    lErro = Comando_Executar(lComando, "SELECT NumIntNF FROM ItensNFiscal WHERE NumIntDoc = ?", lNumIntNF, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 101011
    
    'Busca o Primeiro (e único) registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101012
    
    'Se não encontrou => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 101013

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_NumIntNF_ItemNF = SUCESSO
    
    Exit Function

Erro_MovEstoque_Le_NumIntNF_ItemNF:

    MovEstoque_Le_NumIntNF_ItemNF = gErr
    
    Select Case gErr
    
        Case 101010
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101011, 101012
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)
            
        Case 101013
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150885)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function MovEstoque_Le_NumDocOrigem(lNumIntDocOrigem As Long, iTipoNumIntDocOrigem As Integer, sNumDocOrigem As String, sSerie As String) As Long
'RECEBE o NumIntDocOrigem e o TipoNumIntDocOrigem e através deles busca
'(RETORNA) a Série e o Numero do Documento de Origem

Dim lErro As Long
Dim lNumIntNF As Long
Dim sCodInventario As String
Dim lNumNotaFiscal As Long

On Error GoTo Erro_MovEstoque_Le_NumDocOrigem

    'Limpa o conteúdo da string para evitar armazenamento de dados inválidos
    sSerie = ""
    
    'Faz uma seleção pelo tipo da Origem do Documento
    Select Case iTipoNumIntDocOrigem

        'se for um item de uma OP, MovEstoque ou se não tiver origem...
        Case 0, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
            
            'Apenas Guarda o código
            sNumDocOrigem = lNumIntDocOrigem
            
        'Se for NF ou ítem de NF...
        Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
        
            'Se for um item de NF...
            If iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Then
                                
                'Obtém o NumIntDoc da Nota que Originou o Movimento...
                lErro = MovEstoque_Le_NumIntNF_ItemNF(lNumIntDocOrigem, lNumIntNF)
                If lErro <> SUCESSO And lErro <> 101013 Then gError 101007
                
                'Se não encontrou => ERRO
                If lErro = 101013 Then gError 101014
                
                'Guarda o Numero Interno da NotaFiscal
                lNumIntDocOrigem = lNumIntNF
            
            End If
            
            'Obtém o Número da NF que Originou o Movimento...
            lErro = MovEstoque_Le_NumNotaFiscal_NF(lNumIntDocOrigem, lNumNotaFiscal, sSerie)
            If lErro <> SUCESSO And lErro <> 101018 Then gError 101008
            
            'Se não encontrou => ERRO
            If lErro = 101018 Then gError 101024
            
            sNumDocOrigem = lNumNotaFiscal
        
        Case MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO
            
            'Obtém o Código do Inventário que Originou o Movimento
            lErro = MovEstoque_Le_CodInventario_Inventario(lNumIntDocOrigem, sCodInventario)
            If lErro <> SUCESSO And lErro <> 101022 Then gError 101009
            
            'Se não encontrou => ERRO
            If lErro = 101022 Then gError 101025
            
            'Guarda o código
            sNumDocOrigem = sCodInventario
            
    End Select

    MovEstoque_Le_NumDocOrigem = SUCESSO
    
    Exit Function

Erro_MovEstoque_Le_NumDocOrigem:

    MovEstoque_Le_NumDocOrigem = gErr
    
    Select Case gErr

        Case 101007 To 101009
        
        Case 101014
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_INEXISTENTE", gErr, lNumIntDocOrigem)
        
        Case 101024
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, lNumIntDocOrigem)
            
        Case 101025
            Call Rotina_Erro(vbOKOnly, "ERRO_INVENTARIO_NAO_CADASTRADO1", gErr, lNumIntDocOrigem)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150886)
            
    End Select

End Function

Private Function MovEstoque_Le_NumNotaFiscal_NF(lNumIntDocOrigem As Long, lNumNotaFiscal As Long, sSerie As String) As Long
'Le o numero da nf a partir do numintdoc passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim sSerie1 As String

On Error GoTo Erro_MovEstoque_Le_NumNotaFiscal_NF

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101015
    
    sSerie1 = String(STRING_SERIE, 0)
    
    'Prepara a Execução do Comando SQL
    lErro = Comando_Executar(lComando, "SELECT NumNotaFiscal, Serie FROM NFiscal WHERE NumIntDoc = ?", lNumNotaFiscal, sSerie1, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 101016
    
    'Busca o Primeiro (e único) registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101017
    
    'Se não encontrou => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 101018
    
    sSerie = sSerie1

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_NumNotaFiscal_NF = SUCESSO
    
    Exit Function

Erro_MovEstoque_Le_NumNotaFiscal_NF:

    MovEstoque_Le_NumNotaFiscal_NF = gErr
    
    Select Case gErr
    
        Case 101015
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101016, 101017
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL4", gErr)
            
        Case 101018 'SEM DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150887)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function MovEstoque_Le_CodInventario_Inventario(lNumIntDocOrigem As Long, sCodigo As String) As Long
'Le o codigo do inventário a partir do numintdoc passado como parametro

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MovEstoque_Le_CodInventario_Inventario

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101019
    
    'Prepara a Execução do Comando SQL
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Inventario WHERE NumIntDoc = ?", sCodigo, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 101020
    
    'Busca o Primeiro (e único) registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101021
    
    'Se não encontrou => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 101022

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_CodInventario_Inventario = SUCESSO
    
    Exit Function

Erro_MovEstoque_Le_CodInventario_Inventario:

    MovEstoque_Le_CodInventario_Inventario = gErr
    
    Select Case gErr
    
        Case 101019
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101020, 101021
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIO", gErr)
            
        Case 101022 'SEM DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150888)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Public Function Gera_IN86_Arquivo_TiposMovEstoque(ByVal objIN86Modelos As ClassIN86Modelos, ByVal sNomeArquivo As String, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Tipos de Movimentos de Estoque a partir do nome do arquivo, data inicial e data final passados como parâmetros
'sNomeArquivo RECEBE(Input) o nome completo do Arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim tIN86TiposMovEstoque As typeIN86TiposMovEstoque
Dim objIN86TiposMovEstoque As New ClassTipoMovEst
Dim lComando As Long
Dim sRegistro As String

On Error GoTo Erro_Gera_IN86_Arquivo_TiposMovEstoque

    'Abertura do Comando
    lComando = Comando_Abrir
    If lErro <> SUCESSO Then gError 101026
    
    'Inicializa os buffers
    With tIN86TiposMovEstoque
    
        .sDescricao = String(STRING_IN86_DESC_TIPOSMOVEST, 0)
        .sSigla = String(STRING_IN86_SIGLA_DOC, 0)
    
        'Prepara a Execução do Comando SQL
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Sigla FROM TiposMovimentoEstoque ORDER BY Codigo", .iCodigo, .sDescricao, .sSigla)
    
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 101027
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101028
    
    'Se não encontrou => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 101029
    
    'Cria o arquivo
    Open sNomeArquivo For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os Dados Lidos do Type para o Obj
        Call Move_IN86TiposMovEstoque_Obj(tIN86TiposMovEstoque, objIN86TiposMovEstoque)
    
        'Prepara o Registro para ser inserido no Arquivo
        Call Monta_Registro_TiposMovEstoque(sRegistro, objIN86TiposMovEstoque)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere no Arquivo
        Print #1, sRegistro
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101030
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1
    
    Gera_IN86_Arquivo_TiposMovEstoque = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_TiposMovEstoque:

    Gera_IN86_Arquivo_TiposMovEstoque = gErr
    
    Select Case gErr
    
        Case 101026
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101027, 101028, 101030
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86TIPOSMOVEST", gErr)
            
        Case 101029
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86TIPOSMOVEST_VAZIA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150889)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Sub Move_IN86TiposMovEstoque_Obj(tIN86TiposMovEstoque As typeIN86TiposMovEstoque, objIN86TiposMovEstoque As ClassTipoMovEst)
'Transfere os dados do type para o obj

    With tIN86TiposMovEstoque

        objIN86TiposMovEstoque.iCodigo = .iCodigo
        objIN86TiposMovEstoque.sDescricao = .sDescricao
        objIN86TiposMovEstoque.sSigla = .sSigla

    End With

End Sub

Private Sub Monta_Registro_TiposMovEstoque(sRegistro As String, objIN86TiposMovEstoque As ClassTipoMovEst)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86TiposMovEstoque RECEBE (Input) os valores que serão utilizados para montar o registro

    With objIN86TiposMovEstoque

             sRegistro = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_COD_TIPOSMOVEST, "0", .iCodigo)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_SIGLA_DOC, " ", .sSigla)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_DESC_TIPOSMOVEST, " ", .sDescricao)
        
    End With

End Sub

Public Function Gera_IN86_Arquivo_RegInventario(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Registros de Inventário a partir do nome do arquivo, FilialEmpresa e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim objIN86RegInventario As New ClassRegInventario
Dim tIN86RegInventario As typeRegInventario
Dim lComando As Long
Dim sRegistro As String
Dim sSelecaoSQL As String

On Error GoTo Erro_Gera_IN86_Arquivo_RegInventario

    'Abertura do Comando
    lComando = Comando_Abrir
    If lErro <> SUCESSO Then gError 101032
    
    'Inicializa as strings que serão utilizadas na leitura
    tIN86RegInventario.sContaContabil = String(STRING_CONTA, 0)
    tIN86RegInventario.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tIN86RegInventario.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tIN86RegInventario.sModelo = String(STRING_PRODUTO_MODELO, 0)
    tIN86RegInventario.sObservacoes = String(STRING_OBSERVACAO, 0)
    tIN86RegInventario.sProduto = String(STRING_PRODUTO, 0)
    tIN86RegInventario.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    
    'Se for uma filial específica => faz o filtro por essa filialempresa passada como parametro
    If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Then
    
        'Monta a Selecao SQL com filtro para a filial específica
        sSelecaoSQL = "SELECT FilialEmpresa, Data, Produto, SiglaUMEstoque, QuantidadeUMEstoque, QtdeNossaEmTerc, QtdeDeTercConosco, ValorUnitario, CustoConsig3, CustoConsig, CustoDemo3, CustoDemo, CustoConserto3, CustoConserto, CustoOutras3, CustoOutras, CustoBenef3, CustoBenef FROM RegInventario WHERE Data = ? AND FilialEmpresa = ? ORDER BY Produto"
        
        With tIN86RegInventario
        
            'Procura no BD o Registro de inventário com a data e Filial Empresa passados em objRegInventario
            lErro = Comando_Executar(lComando, sSelecaoSQL, .iFilialEmpresa, .dtData, .sProduto, .sSiglaUMEstoque, .dQuantidadeUMEstoque, .dQtdeNossaEmTerc, .dQtdeDeTercConosco, .dValorUnitario, .dCustoConsig3, .dCustoConsig, .dCustoDemo3, .dCustoDemo, .dCustoConserto3, .dCustoConserto, .dCustoOutras3, .dCustoOutras, .dCustoBenef3, .dCustoBenef, objIN86Modelos.dtDataFim, objIN86Arquivos.iFilialEmpresa)
        
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 101033
    
    Else 'Se for empresa toda
    
        'Monta a Selecao SQL com filtro para a filial específica
        sSelecaoSQL = "SELECT FilialEmpresa, Data, Produto, SiglaUMEstoque, QuantidadeUMEstoque, QtdeNossaEmTerc, QtdeDeTercConosco, ValorUnitario, CustoConsig3, CustoConsig, CustoDemo3, CustoDemo, CustoConserto3, CustoConserto, CustoOutras3, CustoOutras, CustoBenef3, CustoBenef FROM RegInventario WHERE Data = ? ORDER BY Produto"
        
        With tIN86RegInventario
        
            'Procura no BD o Registro de inventário com a data e Filial Empresa passados em objRegInventario
            lErro = Comando_Executar(lComando, sSelecaoSQL, .iFilialEmpresa, .dtData, .sProduto, .sSiglaUMEstoque, .dQuantidadeUMEstoque, .dQtdeNossaEmTerc, .dQtdeDeTercConosco, .dValorUnitario, .dCustoConsig3, .dCustoConsig, .dCustoDemo3, .dCustoDemo, .dCustoConserto3, .dCustoConserto, .dCustoOutras3, .dCustoOutras, .dCustoBenef3, .dCustoBenef, objIN86Modelos.dtDataFim)
        
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 101034
    
    End If
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101037

    'Se não encontrou o Registro de inventário, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 101035
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tiver dados...
    While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os dados do type para o obj
        Call Move_IN86RegInventario_obj(tIN86RegInventario, objIN86RegInventario)
        
        With objIN86RegInventario
            
            'se a quantidade for positiva não nula ...
            If .dQuantidadeUMEstoque > 0 Then
                
                'faz o calculo do valor total
                 .dSaldoValorCusto = .dValorUnitario * .dQuantidadeUMEstoque
                            
                'Prepara o Registro para ser inserido no Arquivo
                Call Monta_Registro_RegInventario(sRegistro, objIN86RegInventario, SITUACAOEST_NOSSO_EM_PODER_CONTRIBUITE)
                
                'Guarda o registro na coleção que será retornada à função chamadora
                colRegistrosDump.Add sRegistro
                
                'Insere no Arquivo
                Print #1, sRegistro
                
            End If
                
            'se a quantidade nossa em terceiros for positiva não nula...
            If .dQtdeNossaEmTerc > 0 Then
            
                'faz o calculo do custo de produtos nossos em terceiros
                .dSaldoValorCusto = .dCustoConsig + .dCustoDemo + .dCustoConserto + .dCustoOutras + .dCustoBenef
                
                'Prepara o Registro para ser inserido no Arquivo
                Call Monta_Registro_RegInventario(sRegistro, objIN86RegInventario, SITUACAOEST_NOSSO_EM_PODER_TERCEIROS)
                
                'Guarda o registro na coleção que será retornada à função chamadora
                colRegistrosDump.Add sRegistro
                
                'Insere no Arquivo
                Print #1, sRegistro
            
            End If
            
            'se a quantidade de terceiros conosco for positiva não nula...
            If .dQtdeDeTercConosco > 0 Then
            
                'faz o calculo do custo de produtos de terceiros em nosso poder
                .dSaldoValorCusto = .dCustoConsig3 + .dCustoDemo3 + .dCustoConserto3 + .dCustoOutras3 + .dCustoBenef3
            
                'Prepara o Registro para ser inserido no Arquivo
                Call Monta_Registro_RegInventario(sRegistro, objIN86RegInventario, SITUACAOEST_TERCEIROS_EM_NOSSO_PODER)
                
                'Guarda o registro na coleção que será retornada à função chamadora
                colRegistrosDump.Add sRegistro
                
                'Insere no Arquivo
                Print #1, sRegistro
            
            End If
        
        End With
        
        'Busca o próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101036
        
    Wend

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1
    
    Gera_IN86_Arquivo_RegInventario = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_RegInventario:

    Gera_IN86_Arquivo_RegInventario = gErr
    
    Select Case gErr
    
        Case 101032
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101033, 101034, 101036, 101037
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGINVENTARIO", gErr)
            
        Case 101035
            Call Rotina_Erro(vbOKOnly, "ERRO_REGINVENTARIO_NAO_CADASTRADO1", gErr, objIN86Modelos.dtDataFim)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150890)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Sub Move_IN86RegInventario_obj(tRegInventario As typeRegInventario, objRegInventario As ClassRegInventario)
'Transfere os dados de tRegInventario para objRegInventario
'tRegInventario RECEBE(Input) os dados que serão passados para o obj
'objRegInventario RETORNA(Output) os dados

    'Passa dados lidos para o objRegInventario
    With objRegInventario
    
        .dCustoBenef = tRegInventario.dCustoBenef
        .dCustoBenef3 = tRegInventario.dCustoBenef3
        .dCustoConserto = tRegInventario.dCustoConserto
        .dCustoConserto3 = tRegInventario.dCustoConserto3
        .dCustoConsig = tRegInventario.dCustoConsig
        .dCustoConsig3 = tRegInventario.dCustoConsig3
        .dCustoDemo = tRegInventario.dCustoDemo
        .dCustoDemo3 = tRegInventario.dCustoDemo3
        .dCustoOutras = tRegInventario.dCustoOutras
        .dCustoOutras3 = tRegInventario.dCustoOutras3
        .dQtdeDeTercConosco = tRegInventario.dQtdeDeTercConosco
        .dQtdeNossaEmTerc = tRegInventario.dQtdeNossaEmTerc
        .dQuantidadeUMEstoque = tRegInventario.dQuantidadeUMEstoque
        .dtData = tRegInventario.dtData
        .dValorUnitario = tRegInventario.dValorUnitario
        .iFilialEmpresa = tRegInventario.iFilialEmpresa
        .sProduto = tRegInventario.sProduto
        .sSiglaUMEstoque = tRegInventario.sSiglaUMEstoque
        
    End With

End Sub

Private Sub Monta_Registro_RegInventario(sRegistro As String, objIN86RegInventario As ClassRegInventario, iSituacaoEST As Integer)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86RegInventario RECEBE (Input) os valores que serão utilizados para montar o registro
'e a situação do estoque (nosso em nosso poder, Nosso em poder terceiros, terceiros em nosso poder)
     
Dim sStrTEMP As String
Dim iVirgula As Integer
Dim sS1 As String
Dim sS2 As String
    
    With objIN86RegInventario
    
        sRegistro = Format(.dtData, "DDMMYYYY")
        
        'SITUACAOEST_NOSSO_EM_PODER_CONTRIBUITE ou SITUACAOEST_NOSSO_EM_PODER_TERCEIROS ou
        'SITUACAOEST_TERCEIROS_EM_NOSSO_PODER
        sRegistro = sRegistro & iSituacaoEST
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", .sProduto)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UM, " ", .sSiglaUMEstoque)
        
        Select Case iSituacaoEST
        
            Case SITUACAOEST_NOSSO_EM_PODER_CONTRIBUITE
                    
                'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
                sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQuantidadeUMEstoque, "0.000"))
            
            Case SITUACAOEST_NOSSO_EM_PODER_TERCEIROS
            
                'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
                sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQtdeNossaEmTerc, "0.000"))
            
            Case SITUACAOEST_TERCEIROS_EM_NOSSO_PODER
        
                'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
                sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQtdeDeTercConosco, "0.000"))
        
        End Select
        
        'Encontra na Quantidade Retornada, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando a Quantidade Formatada (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & FormataCpoValor(.dSaldoValorCusto, STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIALEMPRESA, "0", .iFilialEmpresa)
        
    End With

End Sub

Public Function Gera_IN86_Arquivo_Insumos(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Insumos a partir do nome do arquivo, FilialEmpresa e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim objIN86Insumos As New ClassIN86Insumos
Dim tIN86Insumos As typeIN86Insumos
Dim colIN86Insumos As New Collection
Dim lComando As Long
Dim sSelecaoSQL As String
Dim dtDataFinal As Date

On Error GoTo Erro_Gera_IN86_Arquivo_Insumos

    'Abertura do Comando
    lComando = Comando_Abrir
    If lErro <> SUCESSO Then gError 101032

    'Guarda a Selecao SQL
    sSelecaoSQL = "SELECT Kit.Data, ProdutoKit.ProdutoRaiz, ProdutoKit.Versao, ProdutoKit.Nivel, ProdutoKit.Seq, ProdutoKit.Produto, ProdutoKit.SeqPai, ProdutoKit.Composicao, ProdutoKit.Quantidade, ProdutoKit.UnidadeMed, Produtokit.PercentualPerda FROM Kit, ProdutoKit WHERE Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Kit.Versao = ProdutoKit.Versao ORDER BY ProdutoKit.ProdutoRaiz, Kit.Data, ProdutoKit.Versao, ProdutoKit.Nivel DESC, ProdutoKit.Seq"
    
    'Inicializa os buffers
    With tIN86Insumos
        
        .sProduto = String(STRING_PRODUTO, 0)
        .sProdutoPai = String(STRING_PRODUTO, 0)
        .sUnidadeMedInsumo = String(STRING_UM_SIGLA, 0)
        .sVersao = String(STRING_KIT_VERSAO, 0)
    
        'Prepara a execucao do Comando SQL
        lErro = Comando_Executar(lComando, sSelecaoSQL, .dtData, .sProdutoPai, .sVersao, .iNivel, .iSeq, .sProduto, .iSeqPai, .iComposicao, .dQuantidade, .sUnidadeMedInsumo, .dPercentualPerda)
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 101038
    
    'Busca o primeiro registro para o comando acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101039
    
    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 101040
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'joga para o obj os dados lidos no type
    Call Move_IN86Insumos_Obj(tIN86Insumos, objIN86Insumos)
    
    'Enquanto tem registro...
    Do While lErro <> AD_SQL_SEM_DADOS
            
            'Se a versao foi a mesma e o produto raiz é o mesmo => Kit do mesmo produto e versao
            If tIN86Insumos.sVersao = objIN86Insumos.sVersao And tIN86Insumos.sProdutoPai = objIN86Insumos.sProdutoPai Then
            
                Set objIN86Insumos = New ClassIN86Insumos
                
                'joga para o obj os dados lidos no type
                Call Move_IN86Insumos_Obj(tIN86Insumos, objIN86Insumos)
                
                'Guarda na colecao os dados lidos
                colIN86Insumos.Add objIN86Insumos
                
            Else 'Indica que a versao foi alterada, ou seja, Novo kit
            
                'Se for kit do mesmo produto...
                If tIN86Insumos.sProdutoPai = objIN86Insumos.sProdutoPai Then
                
                    'Gera a data final do kit anterior do mesmo produto
                    dtDataFinal = tIN86Insumos.dtData - 1
                    
                Else 'Novo kit de outro produto...
                
                    'Como essa é a última versão do Kit => não existe data final, pois ainda está em vigor
                    dtDataFinal = 0
                    
                End If
                    
                'processa a colecao, monta o registro e armazena no arquivo
                Call Processa_Colecao_IN86Insumos(objIN86Modelos, colIN86Insumos, dtDataFinal, colRegistrosDump)
                
                Set colIN86Insumos = New Collection
                Set objIN86Insumos = New ClassIN86Insumos
                
                'joga para o obj os dados lidos no type
                Call Move_IN86Insumos_Obj(tIN86Insumos, objIN86Insumos)
                
                'Guarda na colecao os dados lidos
                colIN86Insumos.Add objIN86Insumos
            
            End If
            
            'Busca o próximo registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101041
            
    Loop
    
    'Processa a última colecao armazenada
    Call Processa_Colecao_IN86Insumos(objIN86Modelos, colIN86Insumos, dtDataFinal, colRegistrosDump)
    
    'Fecha o Arquivo
    Close #1
    
    Call Comando_Fechar(lComando)
    
    Gera_IN86_Arquivo_Insumos = SUCESSO
    
    Exit Function

Erro_Gera_IN86_Arquivo_Insumos:

    Gera_IN86_Arquivo_Insumos = gErr
    
    Select Case gErr

        Case 101032
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 101038, 101039, 101041
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86_INSUMOS", gErr)
    
        Case 101040
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86_INSUMOS_NAO_CADASTRADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150891)
    
    End Select

    Call Comando_Fechar(lComando)

End Function

Private Sub Move_IN86Insumos_Obj(tIN86Insumos As typeIN86Insumos, objIN86Insumos As ClassIN86Insumos)
'Transfere os dados do type para o obj

    With tIN86Insumos
    
        objIN86Insumos.dPercentualPerda = .dPercentualPerda
        objIN86Insumos.dQuantidade = .dQuantidade
        objIN86Insumos.dtData = .dtData
        objIN86Insumos.iComposicao = .iComposicao
        objIN86Insumos.iNivel = .iNivel
        objIN86Insumos.iSeq = .iSeq
        objIN86Insumos.iSeqPai = .iSeqPai
        objIN86Insumos.sProduto = .sProduto
        objIN86Insumos.sProdutoPai = .sProdutoPai
        objIN86Insumos.sUnidadeMedInsumo = .sUnidadeMedInsumo
        objIN86Insumos.sVersao = .sVersao
        
    End With

End Sub

Private Sub Processa_Colecao_IN86Insumos(ByVal objIN86Modelos As ClassIN86Modelos, colIN86Insumos As Collection, dtDataFinal As Date, ByVal colRegistrosDump As Collection)
'Recebe a data final de vigencia do kit e uma colecao com o ProdutoRaiz e seus Insumos
'Monta o registro e insere no arquivo
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim objIN86Insumos As ClassIN86Insumos
Dim objIN86Insumos1 As ClassIN86Insumos
Dim sRegistro As String
Dim dQuantidadePai As Double

    'Para cada item da coleção
    For Each objIN86Insumos In colIN86Insumos
    
        'Se o nivel do item atual da colecao for 0 => sai do loop
        If objIN86Insumos.iNivel = 0 Then Exit For
            
        'Para cada item da coleção
        For Each objIN86Insumos1 In colIN86Insumos
            
            'Se o nível do produto atual é o nível do produto do loop anterior -1 _
            e o seq. do produto atual é o seq.pai do produto do loop anterior
            If objIN86Insumos1.iNivel = objIN86Insumos.iNivel - 1 And objIN86Insumos1.iSeq = objIN86Insumos.iSeqPai Then
            
                'Obtém o Produto Pai do insumo e a quantidade base do produto pai
                objIN86Insumos.sProdutoPai = objIN86Insumos1.sProduto
                dQuantidadePai = objIN86Insumos1.dQuantidade
                objIN86Insumos.sUnidadeMedPai = objIN86Insumos1.sUnidadeMedInsumo
                
                'Sai do loop
                Exit For
                
            End If
                
        Next
        
        'Se a composicao é variável...
        If objIN86Insumos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
            
            'Calcula a proporção necessária para produzir uma unidade do produto pai a partir do insumo
            'quantos do insumo para produzir 1 unidade do produto pai e
            objIN86Insumos.dQuantidade = (objIN86Insumos.dQuantidade / dQuantidadePai)
            
        End If
        
        'verifica se a formula está dentro do periodo
        If objIN86Insumos.dtData > objIN86Modelos.dtDataFim Or _
            (dtDataFinal <> 0 And dtDataFinal < objIN86Modelos.dtDataInicio) Then
            
            'está fora
        
        Else
        
            'Monta o registro
            Call Monta_Registro_Insumos(sRegistro, objIN86Insumos, dtDataFinal)
            
            'Guarda o registro na coleção que será retornada à função chamadora
            colRegistrosDump.Add sRegistro
            
            'Insere-o no arquivo
            Print #1, sRegistro
                
        End If
        
    Next

End Sub

Private Sub Monta_Registro_Insumos(sRegistro As String, objIN86Insumos As ClassIN86Insumos, dtDataFinal As Date)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86Insumos RECEBE (Input) os valores que serão utilizados para montar o registro

Dim iVirgula As Integer
Dim sS1 As String
Dim sS2 As String
Dim sStrTEMP As String
    
    With objIN86Insumos

        sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", .sProdutoPai)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UM, " ", .sUnidadeMedPai)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_PRODUTO, " ", .sProduto)
        
        'Formata a Quantidade com o tamanho da STRING_IN86_VALOR, sendo as 3 últimas casas decimais
        sStrTEMP = Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_VALOR + 1, "0", Format(.dQuantidade, "0.000"))
        
        'Encontra na Quantidade Retornada, a posicao da vírgula
        iVirgula = InStr(1, sStrTEMP, ",")
        
        'Quebra a String em duas partes (uma antes e outra depois da Vírgula)
        sS1 = Mid(sStrTEMP, 1, iVirgula - 1)
        sS2 = Mid(sStrTEMP, iVirgula + 1, Len(sStrTEMP))
    
        'Adiciona ao Registro Concatenando a Quantidade Formatada (já sem a vírgula)
        sRegistro = sRegistro & (sS1 & sS2)
        
        sRegistro = sRegistro & FormataCpoValor(Format(.dPercentualPerda * 100, "0.00"), STRING_IN86_PERCENTUALPERDA)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_UM, " ", .sUnidadeMedInsumo)
        
        sRegistro = sRegistro & Format(.dtData, "DDMMYYYY")
        
        'Se a data final é 0 (o kit ainda está em vigor) adiciona 0's ao registro
        If dtDataFinal = 0 Then
            sRegistro = sRegistro & FormataCpoValor(dtDataFinal, STRING_IN86_DATA_ATUALIZACAO)
        Else
            sRegistro = sRegistro & Format(dtDataFinal, "DDMMYYYY")
        End If

    End With

End Sub

Public Function Gera_IN86_Arquivo_CTBClientes(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Insumos a partir do nome do arquivo, FilialEmpresa e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim objIN86CTBClientes As New ClassIN86CTBClientes
Dim tIN86CTBClientes As typeIN86CTBClientes
Dim lComando As Long
Dim sSelecaoSQL As String
Dim sRegistro As String
Dim sContaAux As String

On Error GoTo Erro_Gera_IN86_Arquivo_CTBClientes

    'Abertura do Comando
    lComando = Comando_Abrir
    If lErro <> SUCESSO Then gError 101050
    
    'Inicializa os buffers
    With tIN86CTBClientes
        
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilAux = String(STRING_CONTA, 0)
        .sTipoDocumento = String(STRING_IN86_TIPO_DOC, 0)
        .sTipoOperacao = String(STRING_IN86_TIPO_OPERACAO, 0)
    
        'Se a filial Empresa passada como parametro for uma filial especifica =>
        'faz o filtro por essa filial
        If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Or gobjCTB.giContabCentralizada <> 0 Then
            
            'Guarda a Selecao SQL
            sSelecaoSQL = "SELECT FilialEmpresa, TipoOperacao, MnemonicoCTBValor.Valor, ContaContabil, Cliente, Filial, DataOperacao, ValorOperacao, TipoDocumento, NumTitulo, NumParcela, ValorOriginal, DataEmissao, DataVencimento FROM MnemonicoCTBValor, IN86CTBClientes WHERE DataOperacao >= ? AND DataOperacao <= ? AND MnemonicoCTBValor.Mnemonico = ? AND FilialEmpresa = ?"
            
            'Prepara a execucao do Comando SQL
            lErro = Comando_Executar(lComando, sSelecaoSQL, .iFilialEmpresa, .sTipoOperacao, .sContaContabilAux, .sContaContabil, .lCliente, .iFilial, .dtDataOperacao, .dValorOperacao, .sTipoDocumento, .lDocumento, .iNumParcela, .dValorTitulo, .dtDataEmissao, .dtDataVencimento, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, "CtaReceberCarteira", objIN86Arquivos.iFilialEmpresa)
            
        Else 'Senao busca os registros de todas as FiliaisEmpresas
        
            'Guarda a Selecao SQL
            sSelecaoSQL = "SELECT FilialEmpresa, TipoOperacao, MnemonicoCTBValor.Valor, ContaContabil, Cliente, Filial, DataOperacao, ValorOperacao, TipoDocumento, NumTitulo, NumParcela, ValorOriginal, DataEmissao, DataVencimento FROM MnemonicoCTBValor, IN86CTBClientes WHERE DataOperacao >= ? AND DataOperacao <= ? AND MnemonicoCTBValor.Mnemonico = ?"
            
            'Prepara a execucao do Comando SQL
            lErro = Comando_Executar(lComando, sSelecaoSQL, .iFilialEmpresa, .sTipoOperacao, .sContaContabilAux, .sContaContabil, .lCliente, .iFilial, .dtDataOperacao, .dValorOperacao, .sTipoDocumento, .lDocumento, .iNumParcela, .dValorTitulo, .dtDataEmissao, .dtDataVencimento, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, "CtaReceberCarteira")
    
        End If
        
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 101051
    
    'Busca o primeiro registro para o comando acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101052
    
    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 101053
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tem dados...
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os dados lidos no type para o obj
        lErro = Move_CTBClientes_Obj(tIN86CTBClientes, objIN86CTBClientes, sContaAux)
        If lErro <> SUCESSO Then gError 101063
        
        'Monta o registro a partir dos dados obtidos na busca acima
        Call Monta_Registro_CTBClientes(sRegistro, objIN86CTBClientes)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere os dados no arquivo
        Print #1, sRegistro
    
        'Busca o Próximo registro...
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101054
        
    Loop
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_CTBClientes = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_CTBClientes:

    Gera_IN86_Arquivo_CTBClientes = gErr

    Select Case gErr

        Case 101050
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101051, 101052, 101054
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86CTBCLIENTES", gErr)
            
        Case 101053
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86CTBCLIENTES_INEXISTENTE", gErr)
            
        Case 101063
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150892)

    End Select

End Function

Public Function Gera_IN86_Arquivo_CTBForn(ByVal objIN86Modelos As ClassIN86Modelos, ByVal objIN86Arquivos As ClassIN86Arquivos, ByVal colRegistrosDump As Collection) As Long
'Gera o Arquivo de IN86 para os Insumos a partir do nome do arquivo, FilialEmpresa e data final passados como parâmetros
'objIN86Modelos RECEBE(Input) os parâmetros dtDataInicial e dtDataFinal
'objIN86Arquivos RECEBE(Input) os parâmetros sNomeArquivo e iFilialEmpresa
'colRegistrosDUMP RETORNA(Output) uma coleção com os registros que foram inseridos no arquivo

Dim lErro As Long
Dim objIN86CTBForn As New ClassIN86CTBForn
Dim tIN86CTBForn As typeIN86CTBForn
Dim lComando As Long
Dim sSelecaoSQL As String
Dim sRegistro As String
Dim sContaAux As String

On Error GoTo Erro_Gera_IN86_Arquivo_CTBForn

    'Abertura do Comando
    lComando = Comando_Abrir
    If lErro <> SUCESSO Then gError 101055

    'Inicializa os buffers
    With tIN86CTBForn
        
        .sContaContabil = String(STRING_CONTA, 0)
        .sContaContabilAux = String(STRING_CONTA, 0)
        .sTipoDocumento = String(STRING_IN86_TIPO_DOC, 0)
        .sTipoOperacao = String(STRING_IN86_TIPO_OPERACAO, 0)
    
        'Se a filial Empresa passada como parametro for uma filial especifica =>
        'faz o filtro por essa filial
        If objIN86Arquivos.iFilialEmpresa <> EMPRESA_TODA Or gobjCTB.giContabCentralizada <> 0 Then
            
            'Guarda a Selecao SQL
            sSelecaoSQL = "SELECT FilialEmpresa, TipoOperacao, MnemonicoCTBValor.Valor, ContaContabil, Fornecedor, Filial, DataOperacao, ValorOperacao, TipoDocumento, NumTitulo, NumParcela, ValorOriginal, DataEmissao, DataVencimento FROM MnemonicoCTBValor, IN86CTBForn WHERE DataOperacao >= ? AND DataOperacao <= ? AND MnemonicoCTBValor.Mnemonico = ? AND FilialEmpresa = ?"
            
            'Prepara a execucao do Comando SQL
            lErro = Comando_Executar(lComando, sSelecaoSQL, .iFilialEmpresa, .sTipoOperacao, .sContaContabilAux, .sContaContabil, .lFornecedor, .iFilial, .dtDataOperacao, .dValorOperacao, .sTipoDocumento, .lDocumento, .iNumParcela, .dValorTitulo, .dtDataEmissao, .dtDataVencimento, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, "CtaFornecedores", objIN86Arquivos.iFilialEmpresa)
            
        Else 'Senao busca os registros de todas as FiliaisEmpresas
        
            'Guarda a Selecao SQL
            sSelecaoSQL = "SELECT FilialEmpresa, TipoOperacao, MnemonicoCTBValor.Valor, ContaContabil, Fornecedor, Filial, DataOperacao, ValorOperacao, TipoDocumento, NumTitulo, NumParcela, ValorOriginal, DataEmissao, DataVencimento FROM MnemonicoCTBValor, IN86CTBForn WHERE DataOperacao >= ? AND DataOperacao <= ? AND MnemonicoCTBValor.Mnemonico = ?"
            
            'Prepara a execucao do Comando SQL
            lErro = Comando_Executar(lComando, sSelecaoSQL, .iFilialEmpresa, .sTipoOperacao, .sContaContabilAux, .sContaContabil, .lFornecedor, .iFilial, .dtDataOperacao, .dValorOperacao, .sTipoDocumento, .lDocumento, .iNumParcela, .dValorTitulo, .dtDataEmissao, .dtDataVencimento, objIN86Modelos.dtDataInicio, objIN86Modelos.dtDataFim, "CtaFornecedores")
    
        End If
        
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 101056
    
    'Busca o primeiro registro para o comando acima
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101057
    
    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 101058
    
    'Cria o arquivo
    Open objIN86Arquivos.sNome For Output As #1
    
    'Enquanto tem dados...
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Transfere os dados lidos no type para o obj
        lErro = Move_CTBFornecedores_Obj(tIN86CTBForn, objIN86CTBForn, sContaAux)
        If lErro <> SUCESSO Then gError 101062
        
        'Monta o registro a partir dos dados obtidos na busca acima
        Call Monta_Registro_CTBForn(sRegistro, objIN86CTBForn)
        
        'Guarda o registro na coleção que será retornada à função chamadora
        colRegistrosDump.Add sRegistro
        
        'Insere os dados no arquivo
        Print #1, sRegistro
    
        'Busca o Próximo registro...
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101059
        
    Loop
        
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    'Fecha o Arquivo
    Close #1

    Gera_IN86_Arquivo_CTBForn = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_Arquivo_CTBForn:

    Gera_IN86_Arquivo_CTBForn = gErr

    Select Case gErr

        Case 101055
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 101056, 101057, 101059
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IN86CTBFORN", gErr)
            
        Case 101058
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86CTBFORN_INEXISTENTE", gErr)
        
        Case 101062
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150893)

    End Select

End Function

Private Function Move_CTBClientes_Obj(tIN86CTBClientes As typeIN86CTBClientes, objIN86CTBClientes As ClassIN86CTBClientes, sContaAux As String) As Long
'Transfere os dados do type para o obj

Dim iContaPreenchida As Integer
Dim lErro As Long
    
    With objIN86CTBClientes

        .dtDataEmissao = tIN86CTBClientes.dtDataEmissao
        .dtDataOperacao = tIN86CTBClientes.dtDataOperacao
        .dtDataVencimento = tIN86CTBClientes.dtDataVencimento
        .dValorOperacao = tIN86CTBClientes.dValorOperacao
        .dValorTitulo = tIN86CTBClientes.dValorTitulo
        .iFilial = tIN86CTBClientes.iFilial
        .lCliente = tIN86CTBClientes.lCliente
        .lDocumento = tIN86CTBClientes.lDocumento
        .sArquivamento = " " 'Não existe essa informação em nosso BD's
        .sHistorico = " " 'Não existe essa informação em nosso BD's
        .sTipoDocumento = tIN86CTBClientes.sTipoDocumento
        .sTipoOperacao = tIN86CTBClientes.sTipoOperacao
        .iNumParcela = tIN86CTBClientes.iNumParcela
        .iFilialEmpresa = tIN86CTBClientes.iFilialEmpresa
        
        'Se a conta contábil estiver zerada =>
        If Len(Trim(tIN86CTBClientes.sContaContabil)) = 0 Then
            
            'Se a Conta padrao ainda não foi lida =>
            If Len(Trim(sContaAux)) = 0 Then
            
                'Le a Conta padrao de Clientes
                lErro = CF("Conta_Formata", tIN86CTBClientes.sContaContabilAux, sContaAux, iContaPreenchida)
                If lErro <> SUCESSO Then gError 101061
                
            End If
            
            'Guarda o conteudo da conta
            .sContaContabil = sContaAux
            
        Else 'A conta existe => apenas guarda
        
            .sContaContabil = tIN86CTBClientes.sContaContabil
        
        End If
        
    End With

    Move_CTBClientes_Obj = SUCESSO
    
    Exit Function

Erro_Move_CTBClientes_Obj:

    Move_CTBClientes_Obj = gErr
    
    Select Case gErr
    
        Case 101061
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150894)
    
    End Select

End Function

Private Function Move_CTBFornecedores_Obj(tIN86CTBForn As typeIN86CTBForn, objIN86CTBForn As ClassIN86CTBForn, sContaAux As String) As Long
'Transfere os dados do type para o obj

Dim lErro As Long
Dim iContaPreenchida As Integer
    
On Error GoTo Erro_Move_CTBFornecedores_Obj

    With objIN86CTBForn

        .dtDataEmissao = tIN86CTBForn.dtDataEmissao
        .dtDataOperacao = tIN86CTBForn.dtDataOperacao
        .dtDataVencimento = tIN86CTBForn.dtDataVencimento
        .dValorOperacao = tIN86CTBForn.dValorOperacao
        .dValorTitulo = tIN86CTBForn.dValorTitulo
        .iFilial = tIN86CTBForn.iFilial
        .lFornecedor = tIN86CTBForn.lFornecedor
        .lDocumento = tIN86CTBForn.lDocumento
        .sArquivamento = " " 'Não existe essa informação em nosso BD's
        .sHistorico = " " 'Não existe essa informação em nosso BD's
        .sTipoDocumento = tIN86CTBForn.sTipoDocumento
        .sTipoOperacao = tIN86CTBForn.sTipoOperacao
        .iNumParcela = tIN86CTBForn.iNumParcela
        .iFilialEmpresa = tIN86CTBForn.iFilialEmpresa
        
        'Se a conta contábil estiver zerada =>
        If tIN86CTBForn.sContaContabil = "" Then
            
            'Se a Conta padrao ainda não foi lida =>
            If sContaAux = "" Then
            
                'Le a Conta padrao de Clientes
                lErro = CF("Conta_Formata", tIN86CTBForn.sContaContabilAux, sContaAux, iContaPreenchida)
                If lErro <> SUCESSO Then gError 101060
                
            End If
            
            'Guarda o conteudo da conta
            .sContaContabil = sContaAux
            
        Else 'A conta existe => apenas guarda
        
            .sContaContabil = tIN86CTBForn.sContaContabil
        
        End If
        
    End With
    
    Move_CTBFornecedores_Obj = SUCESSO
    
    Exit Function

Erro_Move_CTBFornecedores_Obj:

    Move_CTBFornecedores_Obj = gErr
    
    Select Case gErr
    
        Case 101060
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150895)
    
    End Select

End Function

Private Sub Monta_Registro_CTBClientes(sRegistro As String, objIN86CTBClientes As ClassIN86CTBClientes)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86CTBClientes RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sCliente As String
Dim sTitulo As String
    
    With objIN86CTBClientes
    
        sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(.sContaContabil))
        
        'sCliente = CodCliente + CodFilialCliente(Com zeros a esquerda)
        sCliente = CStr(.lCliente) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilial))
    
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sCliente)
        
        sRegistro = sRegistro & Format(.dtDataOperacao, "DDMMYYYY")
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_HISTORICOMOVEST, " ", .sHistorico)
        
        sRegistro = sRegistro & FormataCpoValor(.dValorOperacao, STRING_IN86_VALOR)
        
        sRegistro = sRegistro & .sTipoOperacao
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_TIPO_DOC, " ", .sTipoDocumento)
        
        'sTitulo = Numero do Titulo + NumParcela()com zeros a esquerda)
        sTitulo = CStr(.lDocumento) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_PARCELA, "0", .iNumParcela)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NUMERO_DOC, " ", sTitulo)
        
        sRegistro = sRegistro & FormataCpoValor(.dValorTitulo, STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Format(.dtDataEmissao, "DDMMYYYY")
        
        sRegistro = sRegistro & Format(.dtDataVencimento, "DDMMYYYY")
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ARQUIVAMENTO, " ", .sArquivamento)
        
    End With

End Sub

Private Sub Monta_Registro_CTBForn(sRegistro As String, objIN86CTBForn As ClassIN86CTBForn)
'Monta o registro que será inserido no arquivo, ou seja, aplica os formatos e concatena os dados
'de acordo com as regras para geração do IN86
'sRegistro RETORNA (Output) o registro que será inserido no arquivo
'objIN86CTBClientes RECEBE (Input) os valores que serão utilizados para montar o registro

Dim sFornecedor As String
Dim sTitulo As String
    
    With objIN86CTBForn
    
        sRegistro = Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_CONTA_CTB, " ", ContaCtb_Formata(.sContaContabil))
        
        'sFornecedor = CodFornecedor + CodFilialFornecedor(Com zeros a esquerda)
        sFornecedor = CStr(.lFornecedor) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_FILIAL, "0", (.iFilial))
    
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_COD_GENERICO, " ", sFornecedor)
        
        sRegistro = sRegistro & Format(.dtDataOperacao, "DDMMYYYY")
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_HISTORICOMOVEST, " ", .sHistorico)
        
        sRegistro = sRegistro & FormataCpoValor(.dValorOperacao, STRING_IN86_VALOR)
        
        sRegistro = sRegistro & .sTipoOperacao
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_TIPO_DOC, " ", .sTipoDocumento)
        
        'sTitulo = Numero do Titulo + NumParcela(com zeros a esquerda)
        sTitulo = CStr(.lDocumento) & Formata_Campo(ALINHAMENTO_ESQUERDA, STRING_IN86_PARCELA, "0", .iNumParcela)
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_NUMERO_DOC, " ", sTitulo)
        
        sRegistro = sRegistro & FormataCpoValor(.dValorTitulo, STRING_IN86_VALOR)
        
        sRegistro = sRegistro & Format(.dtDataEmissao, "DDMMYYYY")
        
        sRegistro = sRegistro & Format(.dtDataVencimento, "DDMMYYYY")
        
        sRegistro = sRegistro & Formata_Campo(ALINHAMENTO_DIREITA, STRING_IN86_ARQUIVAMENTO, " ", .sArquivamento)
        
    End With

End Sub

'Revisada e alterada por Luiz Nogueira em 28/01/04
'criada por Leo em 13/05/02
Public Function Gera_IN86_Rel_Dump(ByVal colRegistroDump As Collection, ByVal objIN86Arquivo As ClassIN86Arquivos) As Long
'Função que irá gerar o Relatório de Dump por Arquivo à partir da coleçcão de registros
'recebida em colRegistroDump (input) e os dados dos arquivos recebidos em objIN86Arquivo (input).

Dim lIndice As Long
Dim sRegistro As String
Dim iCount As Integer
Dim objFilEmp As New AdmFiliais
Dim lErro As Long

On Error GoTo Erro_Gera_IN86_Rel_Dump

    'Se a coleção estiver vazia, sai da função
    If colRegistroDump.Count = 0 Then Exit Function
    
    'Lê os dados da filial que está gerando os arquivos
    lErro = CF("FilialEmpresa_Le", objFilEmp)
    If lErro <> SUCESSO Then gError 127002
    
    Printer.FontName = "Times New Roman"

    'Altera a posição do papel para Horizontal(deitado)
    Printer.Orientation = vbPRORLandscape

    'Aramzena em iCount a qtd de registros a ser gerado(o nº máximo não poderá ultrapassar QUANT_REGISTROS_DUMP)
    If colRegistroDump.Count < QUANT_REGISTROS_DUMP Then
        iCount = colRegistroDump.Count
    Else
        iCount = QUANT_REGISTROS_DUMP
    End If

    'Altera a fonte para negrito.
    Printer.FontBold = True

    'Altera o tamanho da fonte para 14.
    Printer.FontSize = 14
    
    'Pula 1 linha, como margem superior
    Printer.Print Chr(13)
    
    'Imprime um cabeçalho identificador
    Printer.Print ARQUIVO_MARGEM_ESQUERDA & Replace(ARQUIVO_TITULO_RELDUMP1, "XX", iCount) & objIN86Arquivo.sNome

    'Altera o tamanho da fonte para 7
    Printer.FontSize = 7

    'Altera a fonte para sem negrito.
    Printer.FontBold = False

    'Pula uma Linha
    Printer.Print Chr(13)

    'Imprime os trinta primeiros registros (se houverem)
    For lIndice = 1 To colRegistroDump.Count

        sRegistro = colRegistroDump.Item(lIndice)
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & sRegistro

        If lIndice >= QUANT_REGISTROS_DUMP Then Exit For

    Next

    'Altera o tamanho da fonte para 10
    Printer.FontSize = 10

    Printer.Print Chr(13)

    'Cidade onde a empresa se localiza
    Printer.Print ARQUIVO_MARGEM_ESQUERDA & ARQUIVO_LOCAL_DATA & objFilEmp.objEndereco.sCidade & ", " & Format(gdtDataHoje, "dd ""de"" mmmm ""de"" yyyy") & "."

    Printer.Print Chr(13)
    Printer.Print ARQUIVO_MARGEM_ESQUERDA & ARQUIVO_RESPONSAVEL & "  " & ARQUIVO_LINHA_GRANDE
    Printer.Print Chr(13)
    Printer.Print ARQUIVO_MARGEM_ESQUERDA & ARQUIVO_LINHA_GRANDE

    'Nome empresa-filial
    Printer.Print ARQUIVO_MARGEM_ESQUERDA & gsNomeEmpresa & " " & SEPARADOR & " " & objFilEmp.sNome

    'Se existem mais de 30 registros
    If colRegistroDump.Count > QUANT_REGISTROS_DUMP Then

        'Nova página
        Printer.NewPage

        'Armazena a qtd de registros existentes após o registro de nº 30 e assumindo como o máximo, trinta registros
        If colRegistroDump.Count - QUANT_REGISTROS_DUMP > QUANT_REGISTROS_DUMP Then
            iCount = QUANT_REGISTROS_DUMP
        Else
            iCount = colRegistroDump.Count - QUANT_REGISTROS_DUMP
        End If

        'Altera a fonte para negrito
        Printer.FontBold = True

        'Altera a fonte para tamanho 14.
        Printer.FontSize = 14

        'Pula 1 linha, como margem superior
        Printer.Print Chr(13)
        
        'Imprime um cabeçalho identificador
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & Replace(ARQUIVO_TITULO_RELDUMP2, "XX", iCount) & objIN86Arquivo.sNome

        'Altera a fonte para tamanho 7
        Printer.FontSize = 7

        'Altera a fonte para sem negrito
        Printer.FontBold = False

        'Pula uma Linha
        Printer.Print Chr(13)

        'Imprime os trinta últimos registros (se houverem) à partir do último registro impresso mais um.
        For lIndice = colRegistroDump.Count - iCount + 1 To colRegistroDump.Count

            sRegistro = colRegistroDump.Item(lIndice)
            Printer.Print ARQUIVO_MARGEM_ESQUERDA & sRegistro

        Next

        'Altera o tamanho da fonte para 10
        Printer.FontSize = 10

        Printer.Print Chr(13)
        
        'Cidade onde a empresa se localiza
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & ARQUIVO_LOCAL_DATA & objFilEmp.objEndereco.sCidade & ", " & Format(gdtDataHoje, "dd ""de"" mmmm ""de"" yyyy") & "."

        Printer.Print Chr(13)
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & ARQUIVO_RESPONSAVEL & "  " & ARQUIVO_LINHA_GRANDE
        Printer.Print Chr(13)
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & ARQUIVO_LINHA_GRANDE

        'Nome empresa-filial
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & gsNomeEmpresa & " " & SEPARADOR & " " & objFilEmp.sNome

    End If

    'Finaliza o documento e executa a impressão.
    Printer.EndDoc
    
    Gera_IN86_Rel_Dump = SUCESSO
    
    Exit Function

Erro_Gera_IN86_Rel_Dump:

    Gera_IN86_Rel_Dump = gErr
    
    Select Case gErr
        
        Case 127002
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_GERACAO_DUMP", gErr, objIN86Arquivo.sNome, Error)
            Printer.KillDoc ' Cancela a impressão em arquivo.

    End Select

    Exit Function

End Function

'Revisada e alterada por Luiz Nogueira em 28/01/04
'Função criada por Leo em 14/05/02
Public Function Gera_IN86_ArqRel_Layout(ByVal objIN86TipoArquivo As ClassIN86TiposArquivos, ByVal sNomeArq As String) As Long
'Função que irá gerar/preencher e imprimir um arquivo de texto, cujo nome foi recebido como
'parâmetro em sNomeArq (input) com os dados de Layout recebidos em objIN86TipoArquivo.

Dim sPath As String
Dim sTexto As String

On Error GoTo Erro_Gera_IN86_ArqRel_Layout

    'buscar diretorio configurado
    sPath = String(128, 0)
    Call GetPrivateProfileString("Forprint", "DirTsks", "c:\forpw40\", sPath, 128, "ADM100.INI")

    sPath = StringZ(sPath) & objIN86TipoArquivo.sLayoutArquivo

    'Abre o Arquivo a ser incluido no Leiame como somente leitura.
    Open sPath For Input Access Read As #1

    'Cria/abre o arquivo que irá armazenar os Layouts.
    Open sNomeArq For Append As #2

    'Prepara a impressão para o papel na forma Horizontal(deitado)
    Printer.Orientation = vbPRORLandscape

    'Seleciona a fonte Times New Roman
    Printer.FontName = "Times New Roman"
    
    'Seta o tamanho 14 para o título
    Printer.FontSize = 14

    'coloca a fonte em negrito para o cabeçalho
    Printer.FontBold = True

    'Insere linhas em branco para servirem de margem superior
    Printer.Print Chr(13)
    
    'Lê a primeira linha do arquivo, correspondente ao título do mesmo
    Input #1, sTexto

    'Grava o título  do arquivo
    Print #2, sTexto
    Printer.Print ARQUIVO_MARGEM_ESQUERDA & sTexto
    
    'Seleciona a fonte Courier
    Printer.FontName = "Courier"
    
    'Seta o tamanho 11 para o restante do arquivo
    Printer.FontSize = 11

    'coloca a fonte normal para o restante do arquivo
    Printer.FontBold = False
    
    'Percorre as linhas do arquivo
    Do While Not (EOF(1))

        'Armazena em sTexto os dados do arquivo #1
        Input #1, sTexto

        'Em caso de arquivos de texto onde suas linhas estejam tabuladas, esssas linhas deverão estar entre Aspas.

        'escreve ao final do arquivo #2, os dados do arquivo #1.
        Print #2, sTexto

        'Insere a linha p/ ser impressa.
        Printer.Print ARQUIVO_MARGEM_ESQUERDA & sTexto

    Loop

    'pula uma linha no arquivo.
    Print #2, Chr(13)

    'Fecha o arquivo
    Close #1

    'Fecha o arquivo
    Close #2

    'encerra e executa a impressão
    Printer.EndDoc
    
    Gera_IN86_ArqRel_Layout = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_ArqRel_Layout:

    Gera_IN86_ArqRel_Layout = gErr
    
    Select Case gErr

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_GERACAO_LAYOUT", gErr, objIN86TipoArquivo.sDescricao, Error)

            Printer.Print ARQUIVO_ERRO_GERACAO_LAYOUT & objIN86TipoArquivo.sDescricao
            Printer.KillDoc ' Cancela a impressão em arquivo.
            Print #2, Chr(13) 'Pula uma linha
            Print #2, ARQUIVO_ERRO_GERACAO_LAYOUT & objIN86TipoArquivo.sDescricao
            Print #2, Chr(13) 'Pula uma linha
            Close #1 'fecha o arquivo
            Close #2 'fecha o arquivo

    End Select

    Exit Function

End Function

'Revisada e Alterada por Luiz Nogueira em 28/01/04
'Função criada por Leo em 14/05/02
Public Function Gera_IN86_ArqRel_Acompanhamento(ByVal objIN86Arquivo As ClassIN86Arquivos, ByVal objIN86TipoArquivo As ClassIN86TiposArquivos, ByVal sNomeArq As String, ByVal iMeioFisico As Integer, ByVal lQuantRegistros As Long) As Long
'Fução que irá gerar/preencher e imprimir um arquivo de texto, cujo nome foi recebido como
'parâmetro em sNomeArq (input) com os dados de Acompanhamento auxiliados por
'objIN86Arquivo, sDescricao e iMeioFisico recebidos por parâmetro (Input).

Dim sPath As String
Dim objFilEmp As New AdmFiliais
Dim lErro As Long
Dim objIN86MeioEntrega As New ClassIN86MeioEntrega
Dim iCodigo As Integer
Dim sCGCFormatado As String

On Error GoTo Erro_Gera_IN86_ArqRel_Acompanhamento

    'Lê os dados da filial que está gerando os arquivos
    lErro = CF("FilialEmpresa_Le", objFilEmp)
    If lErro <> SUCESSO Then gError 127003
    
    'Cria/abre o arquivo que irá armazenar o relatório de acompanhamento.
    Open sNomeArq For Append As #1

    Printer.ScaleMode = vbUser
    
    Printer.ScaleLeft = -2
    Printer.ScaleTop = -1
    
    'pula uma linha no arquivo
    Print #1, Chr(13)
    
    'Seleciona a fonte Times New Roman
    Printer.FontName = "Times New Roman"
    
    'Seta fonte tamanho 14 para o cabeçalho
    Printer.FontSize = 14
    
    'Seta fonte em negrito
    Printer.FontBold = True
    
    'pula 1 linha na impressão, para criar uma margem superior
    Printer.Print Chr(13)

    'Daqui para baixo, trecho destinado a montagem do arquivo e a impressão do mesmo.
    'a função Printer.Print envia o texto para a impressora.
    'a função Print #1 envia o texto para o arquivo.

    Printer.Print ARQUIVO_TITULO_RELACOMPANHAMENTO
    Print #1, ARQUIVO_TITULO_RELACOMPANHAMENTO
    
    'Seta fonte normal
    Printer.FontBold = False
    
    'Seta fonte tamanho 9 para o restante do documento
    Printer.FontSize = 9
    
    'Pula uma linha após o título do relatório
    Printer.Print Chr(13)
    
    'Formata o CGC
    Call Formata_CNPJ_CPF(sCGCFormatado, objFilEmp.sCgc)
    
    'CNPJ da filial
    Printer.Print ARQUIVO_CNPJ & sCGCFormatado
    Print #1, ARQUIVO_CNPJ & sCGCFormatado
    
    Printer.Print Chr(13)
    Print #1, Chr(13)

    'Nome da empresa-filial
    Printer.Print ARQUIVO_NOME_EMPRESA & gsNomeEmpresa & SEPARADOR & objFilEmp.sNome
    Print #1, ARQUIVO_NOME_EMPRESA; gsNomeEmpresa & SEPARADOR & objFilEmp.sNome

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_NOME_ARQ & objIN86Arquivo.sNome
    Print #1, ARQUIVO_NOME_ARQ & objIN86Arquivo.sNome

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_DATA_GERACAO & gdtDataAtual
    Print #1, ARQUIVO_DATA_GERACAO & gdtDataAtual

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_CONTEUDO_ARQ & objIN86TipoArquivo.sDescricao
    Print #1, ARQUIVO_CONTEUDO_ARQ & objIN86TipoArquivo.sDescricao
    
    Printer.Print Chr(13)
    Print #1, Chr(13)

    'carrega o codigo do objIN86meioEntrega
    objIN86MeioEntrega.iCodigo = iMeioFisico
    
    'Busca na tabela de meios de entrega
    lErro = CF("IN86MeioEntrega_Le", objIN86MeioEntrega)
    If lErro <> SUCESSO And lErro <> 103641 Then gError 103642
    
    'se o BD estiver vazio
    If lErro = 103641 Then gError 103643

    Printer.Print ARQUIVO_MEIO_FISICO & objIN86MeioEntrega.sDescricao
    Print #1, ARQUIVO_MEIO_FISICO & objIN86MeioEntrega.sDescricao

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_DUMP_PRIMEIROS & "1" & ARQUIVO_QTD_PAGINAS
    Print #1, ARQUIVO_DUMP_PRIMEIROS & "1" & ARQUIVO_QTD_PAGINAS

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_DUMP_ULTIMOS & "1" & ARQUIVO_QTD_PAGINAS
    Print #1, ARQUIVO_DUMP_ULTIMOS & "1" & ARQUIVO_QTD_PAGINAS

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_DESCRICAO_DETALHADA & ARQUIVO_LINHA_PEQUENA & ARQUIVO_QTD_PAGINAS
    Print #1, ARQUIVO_DESCRICAO_DETALHADA & ARQUIVO_LINHA_PEQUENA & ARQUIVO_QTD_PAGINAS

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_QTD_VOLUMES & ARQUIVO_LINHA_PEQUENA
    Print #1, ARQUIVO_QTD_VOLUMES & ARQUIVO_LINHA_PEQUENA

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_QTD_REGISTROS & lQuantRegistros
    Print #1, ARQUIVO_QTD_REGISTROS & lQuantRegistros

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_TAMANHO_ARQ_BYTES & FileLen(objIN86Arquivo.sNome)
    Print #1, ARQUIVO_TAMANHO_ARQ_BYTES & FileLen(objIN86Arquivo.sNome)

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_OUTROS & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_OUTROS & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_LOCAL_DATA & objFilEmp.objEndereco.sCidade & ", " & Format(gdtDataHoje, "dd ""de"" mmmm ""de"" yyyy") & "."
    Print #1, ARQUIVO_LOCAL_DATA & objFilEmp.objEndereco.sCidade & ", " & Format(gdtDataHoje, "dd ""de"" mmmm ""de"" yyyy") & "."

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_CONTRIBUINTE_RESP_PREPOSTO
    Print #1, ARQUIVO_CONTRIBUINTE_RESP_PREPOSTO

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_NOME & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_NOME & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_ASSINATURA & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_ASSINATURA & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_CPF & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_CPF & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_TELEFONE & ARQUIVO_LINHA_MEDIA & ARQUIVO_FAX & ARQUIVO_LINHA_MEDIA
    Print #1, ARQUIVO_TELEFONE & ARQUIVO_LINHA_MEDIA & ARQUIVO_FAX & ARQUIVO_LINHA_MEDIA

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_EMAIL & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_EMAIL & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_RESP_TECNICO
    Print #1, ARQUIVO_RESP_TECNICO

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_NOME & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_NOME & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_ASSINATURA & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_ASSINATURA & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_CPF & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_CPF & ARQUIVO_LINHA_GRANDE

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_TELEFONE & ARQUIVO_LINHA_MEDIA & ARQUIVO_FAX & ARQUIVO_LINHA_MEDIA
    Print #1, ARQUIVO_TELEFONE & ARQUIVO_LINHA_MEDIA & ARQUIVO_FAX & ARQUIVO_LINHA_MEDIA

    Printer.Print Chr(13)
    Print #1, Chr(13)

    Printer.Print ARQUIVO_EMAIL & ARQUIVO_LINHA_GRANDE
    Print #1, ARQUIVO_EMAIL & ARQUIVO_LINHA_GRANDE

    'pula uma linha no arquivo.
    Print #1, Chr(13)

    'Fecha o arquivo
    Close #1

    'encerra e executa a impressão
    Printer.EndDoc
    
    Gera_IN86_ArqRel_Acompanhamento = SUCESSO
    
    Exit Function
    
Erro_Gera_IN86_ArqRel_Acompanhamento:

    Gera_IN86_ArqRel_Acompanhamento = gErr
    
    Select Case gErr
    
        Case 103642, 127003
        
        Case 103643
            Call Rotina_Erro(vbOKOnly, "ERRO_IN86MEIOENTREGA_NAO_ENCONTRADO", gErr)
            
        Case Else

            Call Rotina_Erro(vbOKOnly, "ARQUIVO_ERRO_GERACAO_ACOMPANHAMENTO", gErr, objIN86Arquivo.sNome, Error)

            Printer.Print ARQUIVO_ERRO_GERACAO_ACOMPANHAMENTO & objIN86Arquivo.sNome
            Printer.KillDoc ' Cancela a impressão em arquivo.
            Print #1, Chr(13) 'Pula uma linha
            Print #1, ARQUIVO_ERRO_GERACAO_ACOMPANHAMENTO & objIN86Arquivo.sNome
            Print #1, Chr(13) 'Pula uma linha
            Close #1 'fecha o arquivo
            Close #2 'fecha o arquivo

    End Select

    Exit Function

End Function

'Inicio Daniel
Private Function FornecedorHistorico_Le_Todos(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal colFornecedorHistorico As Collection) As Long
'Le todos os historicos de fornecedores já existentes no banco de dados
'e devolve dentro da colecao

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tFornecedorHistorico As typeFornecedorHistorico
Dim objFornecedorHistorico As ClassFornecedorHistorico

On Error GoTo Erro_FornecedorHistorico_Le_Todos
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97790
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT DataAtualizacao, Codigo, Filial , CPF_CGC, InscricaoEstadual, InscricaoSuframa, InscricaoMunicipal, RazaoSocial, Endereco, Bairro, Cidade, UF, Pais, CEP FROM FornecedorHistorico F WHERE DataAtualizacao <= ? AND EXISTS (SELECT * FROM NFiscal WHERE Fornecedor = F.Codigo AND FilialForn = F.Filial AND DataEmissao >= ? AND DataEmissao <= ?) AND NOT EXISTS (SELECT * FROM FornecedorHistorico F2 WHERE F.Codigo = F2.Codigo AND F.Filial = F2.Filial AND F.DataAtualizacao = F2.DataAtualizacao AND F.NumIntDoc < F2.NumIntDoc) ORDER BY Codigo, Filial, DataAtualizacao"
    
    With tFornecedorHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_FORNECEDOR_RAZAO_SOC, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .lCodigo, .iFilialForn, .sCgc, .sInscricaoEstadual, .sInscricaoSuframa, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP, dtDataFinal, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 97791
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97792

    'Se não encontrou => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 97794
    
    'Enquanto tem fornecedor => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o obj
        Set objFornecedorHistorico = New ClassFornecedorHistorico
        
        'Transfere os dados para um obj para adicionar a colecao
        With objFornecedorHistorico
    
            .dtDataAtualizacao = tFornecedorHistorico.dtDataAtualizacao
            .iFilialForn = tFornecedorHistorico.iFilialForn
            .lCodigo = tFornecedorHistorico.lCodigo
            .sBairro = tFornecedorHistorico.sBairro
            .sCEP = tFornecedorHistorico.sCEP
            .sCgc = tFornecedorHistorico.sCgc
            .sCidade = tFornecedorHistorico.sCidade
            .sEndereco = tFornecedorHistorico.sEndereco
            .sInscricaoEstadual = tFornecedorHistorico.sInscricaoEstadual
            .sInscricaoMunicipal = tFornecedorHistorico.sInscricaoMunicipal
            .sInscricaoSuframa = tFornecedorHistorico.sInscricaoSuframa
            .sUF = tFornecedorHistorico.sUF
            .sRazaoSocial = tFornecedorHistorico.sRazaoSocial
            .sPais = tFornecedorHistorico.sPais
            
        End With
    
        'Adiciona a Colecao o ObjFornecedorHistorico
        colFornecedorHistorico.Add objFornecedorHistorico
                
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97793
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    FornecedorHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_FornecedorHistorico_Le_Todos:

    FornecedorHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97790
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97791, 97792, 97793
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORHISTORICO", gErr)
                    
        Case 97794
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148761)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function
'Fim Daniel

'Revisado - OK Luiz 19/02/02
Private Function ClienteHistorico_Le_Todos(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal colClienteHistorico As Collection) As Long
'Lê todos os registros da tabela ClienteHistorico e devolve-os na coleção
'colClienteHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tClienteHistorico As typeClienteHistorico
Dim objClienteHistorico As ClassClienteHistorico

On Error GoTo Erro_ClienteHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97806
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT DataAtualizacao, Codigo, Filial , CPF_CGC, InscricaoEstadual, InscricaoSuframa, InscricaoMunicipal, RazaoSocial, Endereco, Bairro, Cidade, UF, Pais, CEP FROM ClienteHistorico C WHERE DataAtualizacao <= ? AND EXISTS (SELECT * FROM NFiscal WHERE Cliente = C.Codigo AND FilialCli = C.Filial AND DataEmissao >= ? AND DataEmissao <= ?) AND NOT EXISTS (SELECT * FROM ClienteHistorico C2 WHERE C.Codigo = C2.Codigo AND C.Filial = C2.Filial AND C.DataAtualizacao = C2.DataAtualizacao AND C.NumIntDoc < C2.NumIntDoc) ORDER BY Codigo, Filial, DataAtualizacao"
    
    With tClienteHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .lCodigo, .iFilialCliente, .sCgc, .sInscricaoEstadual, .sInscricaoSuframa, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP, dtDataFinal, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 97807
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97808

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97805
    
    'Enquanto tem registro => adiciona na coleção
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o obj
        Set objClienteHistorico = New ClassClienteHistorico
        
        With objClienteHistorico
    
            .dtDataAtualizacao = tClienteHistorico.dtDataAtualizacao
            .iFilialCliente = tClienteHistorico.iFilialCliente
            .lCodigo = tClienteHistorico.lCodigo
            .sBairro = tClienteHistorico.sBairro
            .sCEP = tClienteHistorico.sCEP
            .sCgc = tClienteHistorico.sCgc
            .sCidade = tClienteHistorico.sCidade
            .sEndereco = tClienteHistorico.sEndereco
            .sInscricaoEstadual = tClienteHistorico.sInscricaoEstadual
            .sInscricaoMunicipal = tClienteHistorico.sInscricaoMunicipal
            .sInscricaoSuframa = tClienteHistorico.sInscricaoSuframa
            .sUF = tClienteHistorico.sUF
            .sRazaoSocial = tClienteHistorico.sRazaoSocial
            .sPais = tClienteHistorico.sPais
            
        End With
    
        'Adiciona na Colecao
        colClienteHistorico.Add objClienteHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97809
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    ClienteHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_ClienteHistorico_Le_Todos:

    ClienteHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97806
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97807, 97808, 97809
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTEHISTORICO", gErr)
            
        Case 97805 'Sem Dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149185)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

'Revisado - OK Luiz 19/02/02
Private Function ProdutoHistorico_Le_Todos(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal colProdutoHistorico As Collection) As Long
'Lê todos os registros da tabela ProdutoHistorico e devolve-os na coleção
'colProdutoOPHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tProdutoHistorico As typeProdutoHistorico
Dim objProdutoHistorico As ClassProdutoHistorico

On Error GoTo Erro_ProdutoHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97821
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT DataAtualizacao, CodProduto, DescProduto FROM ProdutoHistorico P WHERE DataAtualizacao <= ? AND NOT EXISTS (SELECT * FROM ProdutoHistorico P2 WHERE P.CodProduto = P2.CodProduto AND P.DataAtualizacao = P2.DataAtualizacao AND P.NumIntDoc < P2.NumIntDoc ) ORDER BY CodProduto, DataAtualizacao"
    
    With tProdutoHistorico
        
        'Inicializa as Strings do Type
        .sCodigoProd = String(STRING_PRODUTO, 0)
        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .sCodigoProd, .sDescProduto, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 97822
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97823

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97820
    
    'Enquanto tem Produto => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o obj
        Set objProdutoHistorico = New ClassProdutoHistorico
        
        With objProdutoHistorico
    
            .dtDataAtualizacao = tProdutoHistorico.dtDataAtualizacao
            .sCodProduto = tProdutoHistorico.sCodigoProd
            .sDescProduto = tProdutoHistorico.sDescProduto
            
        End With
    
        'Adiciona na Colecao
        colProdutoHistorico.Add objProdutoHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97824
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    ProdutoHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_ProdutoHistorico_Le_Todos:

    ProdutoHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97821
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97822, 97823, 97824
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOHISTORICO", gErr)
            
        Case 97820 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152154)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

'Revisado - OK Luiz 19/02/02
Private Function TranspHistorico_Le_Todos(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal colTranspHistorico As Collection) As Long
'Lê todos os registros da tabela TransportadoraHistorico e devolve-os na coleção
'colTranspHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tTranspHistorico As typeTranspHistorico
Dim objTranspHistorico As ClassTranspHistorico

On Error GoTo Erro_TranspHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97795
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT DataAtualizacao, CodTransp, CGC, InscricaoEstadual, InscricaoMunicipal, RazaoSocial, Endereco, Bairro, Cidade, UF, Pais, CEP FROM TransportadoraHistorico T WHERE DataAtualizacao <= ? AND EXISTS (SELECT * FROM NFiscal WHERE CodTransportadora = T.CodTransp AND DataEmissao >= ? AND DataEmissao <= ?) AND NOT EXISTS (SELECT * FROM TransportadoraHistorico T2 WHERE T.CodTransp = T2.CodTransp AND T.DataAtualizacao = T2.DataAtualizacao AND T.NumIntDoc < T2.NumIntDoc) ORDER BY CodTransp, DataAtualizacao"
    
    With tTranspHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_TRANSPORTADORA_NOME, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .iCodTransp, .sCgc, .sInscricaoEstadual, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP, dtDataFinal, dtDataInicial, dtDataFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 97796
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97797

    If lErro = AD_SQL_SEM_DADOS Then gError 97798
    
    'Enquanto tem Transportadora => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Inicializa o obj
        Set objTranspHistorico = New ClassTranspHistorico
        
        With objTranspHistorico
    
            .dtDataAtualizacao = tTranspHistorico.dtDataAtualizacao
            .iCodTransp = tTranspHistorico.iCodTransp
            .sBairro = tTranspHistorico.sBairro
            .sCEP = tTranspHistorico.sCEP
            .sCgc = tTranspHistorico.sCgc
            .sCidade = tTranspHistorico.sCidade
            .sEndereco = tTranspHistorico.sEndereco
            .sInscricaoEstadual = tTranspHistorico.sInscricaoEstadual
            .sInscricaoMunicipal = tTranspHistorico.sInscricaoMunicipal
            .sUF = tTranspHistorico.sUF
            .sRazaoSocial = tTranspHistorico.sRazaoSocial
            .sPais = tTranspHistorico.sPais
            
        End With
    
        'Adiciona na colecao
        colTranspHistorico.Add objTranspHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97799
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    TranspHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_TranspHistorico_Le_Todos:

    TranspHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97795
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97796, 97797, 97799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPHISTORICO", gErr)
            
        Case 97798 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149186)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

