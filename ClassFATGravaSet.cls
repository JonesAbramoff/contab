VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Importa_Dados_Excel(ByVal iFilialEmpresa As Integer, ByVal sArquivo As String, ByVal objMsg As Object, ByVal objProgresso As Object, ByVal objTela As Object) As Long
        
Dim lErro As Long
Dim iIndice As Integer
Dim sDiretorio As String
Dim lRetorno As Long
Dim objCliente As ClassCliente
Dim objPastaTrabalho As Object 'Excel.Workbook
Dim objPlanilhaExcel As Object 'Excel.Worksheet
Dim objImportacaoPV As New ClassImportacaoPV
Dim objPVTela As ClassPedidoDeVenda
Dim iPosNew As Integer
Dim iPos As Integer
Dim sDir As String
Dim sArq As String
Dim lTransacao As Long
Dim sDiretorioImportTemp As String

On Error GoTo Erro_Importa_Dados_Excel


    iPosNew = 1
    iPos = 1
   
    Do While iPosNew <> 0

        iPosNew = InStr(iPos, sArquivo, Chr(0))
   
        If iPosNew = 0 Then

            sArq = sDir & Mid(sArquivo, iPos)

        Else
        
            If iPos = 1 Then
                sDir = Mid(sArquivo, iPos, iPosNew - iPos) & "\"
            End If
        
            iPos = iPosNew + 1
            iPosNew = InStr(iPos, sArquivo, Chr(0))
        
            sArq = sDir & Mid(sArquivo, iPos, iPosNew - iPos)
            iPos = iPosNew + 1

        End If


        'Abertura de transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 197638
    
        'Obtém o diretório onde estão os arquivos
        sDiretorio = String(512, 0)
        lRetorno = GetPrivateProfileString("Geral", "dirArqImportTemp", "c:\", sDiretorio, 512, "ADM100.INI")
        sDiretorio = Left(sDiretorio, lRetorno)
        
        sDiretorioImportTemp = sDiretorio
        
        FileCopy sArq, sDiretorioImportTemp & Dir(sArq)
        
        'Abre o excel
        lErro = CF("Excel_Abrir")
        If lErro <> 0 Then gError 197640
        
        'Desabilita as mensagens do Excel
        GL_objExcelSist.DisplayAlerts = False
        
        objMsg.AddItem "Abrindo a planilha " & sArq
        
        Set objPastaTrabalho = GL_objExcelSist.Workbooks.Open(sArq)
        
        DoEvents
        
        If objTela.giStop = 1 Then gError 197641
        
        objMsg.AddItem "Planilha aberta, iniciando importação"
        
        'Seleciona a planilha ativa na pasta de trabalho criada
        Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
    
        Set objCliente = New ClassCliente
    
        lErro = Importa_Dados_Cliente_Excel(objPlanilhaExcel, objCliente)
        If lErro <> SUCESSO Then gError 197642
    
        Set objPVTela = New ClassPedidoDeVenda
    
        lErro = Importa_Dados_PV_Excel(iFilialEmpresa, objPlanilhaExcel, objCliente, objPVTela)
        If lErro <> SUCESSO Then gError 197643
                    
        objImportacaoPV.iFilialEmpresa = iFilialEmpresa
        
        objImportacaoPV.sArquivo = Dir(sArq)
        
        objImportacaoPV.lCodigoPV = objPVTela.lCodigo
        
        objImportacaoPV.dtData = gdtDataAtual
        
        objImportacaoPV.dtHora = Time
        
        objImportacaoPV.sUsuario = gsUsuario
        
        lErro = CF("ImportacaoPV_Exclui", objImportacaoPV)
        If lErro <> SUCESSO Then gError 197644
        
        lErro = CF("ImportacaoPV_Grava", objImportacaoPV)
        If lErro <> SUCESSO Then gError 197645
    
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 197646
    
        objMsg.AddItem "Arquivo " & sArq & " importado com sucesso."
    
        DoEvents
        
        If objTela.giStop = 1 Then gError 197647
        
        Call CF("Excel_Fechar")
    
    Loop
    
    Importa_Dados_Excel = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Excel:

    Importa_Dados_Excel = gErr
     
    Select Case gErr
    
        Case 197638
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 197640
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_EXCEL", gErr)
    
        Case 197641
        
        Case 197642 To 197645, 197647
        
        Case 197646
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197648)
     
    End Select
    
    Call Transacao_Rollback

    'Fecha o Excel
    Call CF("Excel_Fechar")
    
    Exit Function

End Function

Private Function Importa_Dados_Cliente_Excel(ByVal objPlanilhaExcel As Object, ByVal objCliente As ClassCliente) As Long
        
Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim lCodigo As Long
Dim objEnd As New ClassEndereco
Dim objCidades As New ClassCidades
Dim colEndereco As New Collection
Dim iCodigoCidade As Integer

On Error GoTo Erro_Importa_Dados_Cliente_Excel

    objCliente.lCodigo = StrParaLong(objPlanilhaExcel.Cells(1, 1).Value)
        
    If objCliente.lCodigo = 0 Then
        
        'Gera código automático do próximo cliente
        lErro = CF("Config_ObterAutomatico_EmTrans", "CPRConfig", "NUM_PROX_CLIENTE", "Clientes", "Codigo", lCodigo)
        If lErro <> SUCESSO Then gError 197649

        objCliente.lCodigo = lCodigo
            
        objCliente.sRazaoSocial = Left(objPlanilhaExcel.Cells(1, 2).Value, STRING_CLIENTE_RAZAO_SOCIAL)
        If Len(Trim(objCliente.sRazaoSocial)) = 0 Then gError 197650
        
        objCliente.sNomeReduzido = Left(objPlanilhaExcel.Cells(1, 3).Value, STRING_CLIENTE_NOME_REDUZIDO)
        If Len(Trim(objCliente.sNomeReduzido)) = 0 Then gError 197651
        
        objCliente.iTipo = StrParaInt(objPlanilhaExcel.Cells(1, 4).Value)
        
        objEnd.sEndereco = Left(objPlanilhaExcel.Cells(1, 5).Value, STRING_ENDERECO)
        objEnd.sBairro = Left(objPlanilhaExcel.Cells(1, 6).Value, STRING_BAIRRO)
        objEnd.sCidade = Left(objPlanilhaExcel.Cells(1, 7).Value, STRING_CIDADE)
        objEnd.sSiglaEstado = Left(objPlanilhaExcel.Cells(1, 8).Value, STRING_ESTADO)
        objEnd.sCEP = Left(objPlanilhaExcel.Cells(1, 9).Value, STRING_CEP)
        objEnd.sTelefone1 = Left(objPlanilhaExcel.Cells(1, 10).Value, STRING_TELEFONE)
        objEnd.sFax = Left(objPlanilhaExcel.Cells(1, 11).Value, STRING_FAX)
        objEnd.sEmail = Left(objPlanilhaExcel.Cells(1, 12).Value, STRING_EMAIL)
        objEnd.sContato = Left(objPlanilhaExcel.Cells(1, 13).Value, STRING_CONTATO)
        objEnd.iCodigoPais = 1

        objCliente.iAtivo = CLIENTE_ATIVO

        objCliente.iTipoFrete = TIPO_CIF
    
        'verifica preenchimento de campos de acordo com configuracao
        lErro = CF("Cliente_Verifica_Preenchimento", objCliente)
        If lErro <> SUCESSO Then gError 197652

        objFilialCliente.lCodCliente = objCliente.lCodigo

        objFilialCliente.iTipoFrete = objCliente.iTipoFrete
    
        objFilialCliente.iAtivo = objCliente.iAtivo

        'verifica preenchimento de campos de acordo com configuracao
        lErro = CF("FilialCliente_Verifica_Preenchimento", objFilialCliente)
        If lErro <> SUCESSO Then gError 197653
    
        objCidades.sDescricao = objEnd.sCidade
    
        lErro = CF("Cidade_Le_Nome", objCidades)
        
        If lErro <> SUCESSO Then
        
            lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigoCidade)
            If lErro <> SUCESSO Then gError 197654
                    
            objCidades.iCodigo = iCodigoCidade
            objCidades.sDescricao = objEnd.sCidade
            
            lErro = CF("Cidade_Grava_EmTrans", objCidades)
            If lErro <> SUCESSO Then gError 197655
        
        End If
    
        colEndereco.Add objEnd
        
        Set objEnd = New ClassEndereco
        
        objEnd.iCodigoPais = 1
        objEnd.sSiglaEstado = Left(objPlanilhaExcel.Cells(1, 8).Value, STRING_ESTADO)
    
        colEndereco.Add objEnd
    
        Set objEnd = New ClassEndereco
    
        objEnd.iCodigoPais = 1
        objEnd.sSiglaEstado = Left(objPlanilhaExcel.Cells(1, 8).Value, STRING_ESTADO)
    
        colEndereco.Add objEnd
    
        'Grava o Cliente no BD
        lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
        If lErro <> SUCESSO Then gError 197656
    
    Else
    
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 197657
        
        If lErro <> SUCESSO Then gError 197658
    
    End If

    Importa_Dados_Cliente_Excel = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_Cliente_Excel:

    Importa_Dados_Cliente_Excel = gErr
     
    Select Case gErr
    
        Case 197649
    
        Case 197650
            Call Rotina_Erro(vbOKOnly, "ERRO_RAZ_SOC_NAO_PREENCHIDA", gErr)

        Case 197651
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", gErr)
            
        Case 197652 To 197657

        Case 197658
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197659)
     
    End Select
 
    Exit Function

End Function

Private Function Importa_Dados_PV_Excel(ByVal iFilialEmpresa As Integer, ByVal objPlanilhaExcel As Object, ByVal objCliente As ClassCliente, ByVal objPVTela As ClassPedidoDeVenda) As Long
        
Dim lErro As Long
Dim lSeq As Long
Dim lTransacao As Long
Dim lTotalRegistros As Long
Dim objPV As New ClassPedidoDeVenda
Dim sTipoReg As String
Dim sRegistro As String
Dim colIntegracaoPV As New Collection
Dim objTelaPedido As Object
Dim objTelaPedidoOCX As Object
Dim bLoad As Boolean
Dim iCancel As Integer
Dim objItemPV As ClassItemPedido
Dim iCodVendedor As Integer
Dim iLinha As Integer
Dim iItem As Integer
Dim colReserva As Collection

On Error GoTo Erro_Importa_Dados_PV_Excel

    Set objTelaPedidoOCX = CreateObject("TelasFAT.PedidoVendaOCX")
    Set objTelaPedido = objTelaPedidoOCX.objCTtela

    bLoad = False
    Call objTelaPedidoOCX.Form_Load
    bLoad = True
    Call objTelaPedido.Form_Activate

    objPV.iFilialEmpresa = iFilialEmpresa
    objPV.lCliente = objCliente.lCodigo
    objPV.iFilial = FILIAL_MATRIZ
    objPV.iFilialEntrega = FILIAL_MATRIZ
    objPV.iCondicaoPagto = StrParaLong(objPlanilhaExcel.Cells(2, 1).Value)
    objPV.dtDataEmissao = StrParaDate(objPlanilhaExcel.Cells(2, 2).Value)
    objPV.dtDataReferencia = StrParaDate(objPlanilhaExcel.Cells(2, 3).Value)
    iCodVendedor = StrParaInt(objPlanilhaExcel.Cells(2, 4).Value)
    
    If objPV.dtDataReferencia = DATA_NULA Then objPV.dtDataReferencia = objPV.dtDataEmissao
    
    iLinha = 3
    iItem = 0
    
    Do While Len(objPlanilhaExcel.Cells(iLinha, 1).Value) <> 0
    
        Set objItemPV = New ClassItemPedido
    
        iItem = iItem + 1
        objItemPV.sProduto = objPlanilhaExcel.Cells(iLinha, 1).Value
        objItemPV.dQuantidade = StrParaDbl(objPlanilhaExcel.Cells(iLinha, 2).Value)
        objItemPV.dPrecoUnitario = StrParaDbl(objPlanilhaExcel.Cells(iLinha, 3).Value)
        objItemPV.dValorDesconto = (objItemPV.dQuantidade * objItemPV.dPrecoUnitario) * (StrParaDbl(objPlanilhaExcel.Cells(iLinha, 4).Value) / 100)
    
        iLinha = iLinha + 1
        
        objPV.colItensPedido.Add objPV.iFilialEmpresa, objPV.lCodigo, iItem, objItemPV.sProduto, _
        objItemPV.dQuantidade, objItemPV.dPrecoUnitario, 0, objItemPV.dValorDesconto, DATA_NULA, _
        "", 0, 0, 0, colReserva, "", "", "", "", 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    
    Loop
    
    lErro = Importa_Dados_PV_Preenche_Excel(objPV, objPVTela, objTelaPedido, iCodVendedor)
    If lErro <> SUCESSO Then gError 197660
    
    lErro = CF("PedidoVenda_Grava_EmTrans", objPVTela)
    If lErro <> SUCESSO Then gError 197661
    
    Call objTelaPedidoOCX.Form_Unload(iCancel)
    bLoad = False

    Set objTelaPedidoOCX = Nothing
    Set objTelaPedido = Nothing
    
    Importa_Dados_PV_Excel = SUCESSO
     
    Exit Function
    
Erro_Importa_Dados_PV_Excel:

    Importa_Dados_PV_Excel = gErr
     
    Select Case gErr
    
        Case 197660, 197661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197662)
     
    End Select

    If bLoad Then Call objTelaPedidoOCX.Form_Unload(iCancel)
    
    Set objTelaPedidoOCX = Nothing
    Set objTelaPedido = Nothing
  
    Exit Function

End Function

Public Function Importa_Dados_PV_Preenche_Excel(ByVal objPV As ClassPedidoDeVenda, ByVal objPVTela As ClassPedidoDeVenda, ByVal objTelaPedido As Object, ByVal iCodVendedor As Integer) As Long

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim lCodigo As Long
Dim lCliente As Long
Dim iFilial As Integer
Dim iCondPagto As Integer
Dim iVendedor As Integer
Dim sProduto As String
Dim iTransp As Integer
Dim bAchouCliente As Boolean
Dim bAchouCondPagto As Boolean
Dim bAchouVendedor As Boolean
Dim bAchouProduto As Boolean
Dim bAchouTransp As Boolean
Dim bAchouAlm As Boolean
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objMensagem As New ClassMensagem
Dim objTipoCliente As New ClassTipoCliente
Dim objTransportadora As New ClassTransportadora
Dim objProduto As ClassProduto
Dim iSeq As Integer
Dim iProdutoPreenchido As Integer
Dim objEndereco As New ClassEndereco
Dim iAlmoxarifado As Integer
Dim bNovoPedido As Boolean
Dim sProdutoRegistro As String
Dim objReservaItem As ClassReservaItem
Dim objComissao As ClassComissaoPedVendas
Dim sNaturezaOPAnt As String
Dim iIndice As Integer
Dim dValorPV As Double
Dim objOV As New ClassOrcamentoVenda

On Error GoTo Erro_Importa_Dados_PV_Preenche_Excel

    If objPV.colItensPedido.Count = 0 Then Exit Function
    
    'Obtém o código
    lErro = CF("Config_ObterAutomatico_EmTrans", "FatConfig", "NUM_PROX_CODIGO_PEDVENDA", "PedVenTodos", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 197681
    
    objPV.lCodigo = lCodigo
    
    objPV.dtDataEntrega = DATA_NULA
    
    objFilialCliente.lCodCliente = objPV.lCliente
    objFilialCliente.iCodFilial = objPV.iFilial
    
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 197663
    
    If lErro <> SUCESSO Then gError 197664
    
    objCliente.lCodigo = objPV.lCliente
    
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 197665
    
    If lErro <> SUCESSO Then gError 197666
    
    objTipoCliente.iCodigo = objCliente.iTipo
    
    lErro = CF("TipoCliente_Le", objTipoCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then gError 197667
    
    If lErro <> SUCESSO Then gError 197668
    
    objEndereco.lCodigo = objFilialCliente.lEndereco

    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 197669
    
    If lErro <> SUCESSO Then gError 197670
    
    dValorPV = 0
    
    For Each objItemPV In objPV.colItensPedido
        dValorPV = dValorPV + objItemPV.dPrecoUnitario * objItemPV.dQuantidade - objItemPV.dValorDesconto
    Next
    
    objPV.dValorTotal = dValorPV
    
    objPV.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO

    objPV.iFilialEmpresaFaturamento = giFilialEmpresa
    objPV.iFreteRespons = 1
    
    iVendedor = iCodVendedor
    
    iSeq = 0
    
    For Each objItemPV In objPV.colItensPedido
        
        iSeq = iSeq + 1
        
        Set objProduto = New ClassProduto
        
        objItemPV.iItem = iSeq
        
        sProdutoRegistro = objItemPV.sProduto
        
        lErro = CF("Produto_Formata", sProdutoRegistro, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 197671
        
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 197672
        
        If lErro = 28030 Then gError 197673
            
        objItemPV.sProduto = objProduto.sCodigo
        objItemPV.sDescricao = objProduto.sDescricao
        objItemPV.sProdutoDescricao = objProduto.sDescricao
        
        objItemPV.sUnidadeMed = objProduto.sSiglaUMVenda
        objItemPV.dPrecoTotal = (objItemPV.dPrecoUnitario * objItemPV.dQuantidade) - objItemPV.dValorDesconto

    Next
        
    lErro = objTelaPedido.Traz_Pedido_Tela(objPV, objOV)
    If lErro <> SUCESSO Then gError 197674
    
    lErro = objTelaPedido.Trata_FilialCliente
    If lErro <> SUCESSO Then gError 197675

    For iSeq = 1 To objPV.colItensPedido.Count
        Call CF2(objTelaPedido, "PV_PrecoTotal_Calcula", iSeq)
    Next
    
    lErro = objTelaPedido.ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 197676
    
    objTelaPedido.giNaturezaOpAlterada = MARCADO
    sNaturezaOPAnt = objPV.sNaturezaOp
    
    Call objTelaPedido.TributacaoRecalcular_Click
    
    objTelaPedido.NaturezaOp.Text = sNaturezaOPAnt
    Call objTelaPedido.NaturezaOp_Validate(bSGECancelDummy)
    
    objTelaPedido.cobrancaautomatica.Value = 1

    Call objTelaPedido.Cobranca_Automatica
    
    If iVendedor <> 0 Then
        objFilialCliente.iVendedor = iVendedor
    End If
    
    lErro = objTelaPedido.objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
    If lErro <> SUCESSO Then gError 197677
    
    Call objTelaPedido.objTabComissoes.Comissoes_Calcula
    
    Call objTelaPedido.Pesos_Recalcula
    
    Call objTelaPedido.BotaoRefazAlocacao_Click
                
    lErro = objTelaPedido.Move_Pedido_Memoria(objPVTela)
    If lErro <> SUCESSO Then gError 197678
    
'    If bNovoPedido Then
'
'        If iVendedor <> 0 Then
'            For Each objComissao In objPVTela.colComissoes
'                objComissao.iCodVendedor = iVendedor
'            Next
'        End If
'
'        For iIndice = objPVTela.colComissoes.Count To 1 Step -1
'            Set objComissao = objPVTela.colComissoes.Item(iIndice)
'            If objComissao.dValor < 0.01 Then objComissao.dValor = 0.01
''            If objComissao.dValor <= 0 Then
''                objPVTela.colComissoes.Remove iIndice
''            End If
'        Next
'
'        lErro = objTelaPedido.objTabComissoes.Carrega_Tab_Comissoes(objPVTela)
'        If lErro <> SUCESSO Then gError 99999
'
'        iSeq = 0
'        For Each objIntegracaoPV In colIntegracaoPV
'            iSeq = iSeq + 1
'            Set objItemPV = objPVTela.colItensPedido.Item(iSeq)
'
'            bAchouAlm = False
'            If objIntegracaoPV.iCodAlmoxarifado = 0 Then
'
'                If objIntegracaoPV.iCodAlmoxarifadoExterno <> 0 Then
'
'                    lErro = Comando_Executar(lComando9, "SELECT Codigo FROM Almoxarifado WHERE CodExterno = ? ", iAlmoxarifado, objIntegracaoPV.iCodAlmoxarifadoExterno)
'                    If lErro <> AD_SQL_SUCESSO Then
'                        bErroNoRegistro = True
'                        Call Log_Adiciona("Erro na leitura do almoxarifado com código externo " & CStr(objIntegracaoPV.iCodAlmoxarifadoExterno) & ".", objLogBase)
'                    End If
'
'                    lErro = Comando_BuscarProximo(lComando9)
'                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
'                        bErroNoRegistro = True
'                        Call Log_Adiciona("Erro na leitura do almoxarifado com código externo " & CStr(objIntegracaoPV.iCodAlmoxarifadoExterno) & ".", objLogBase)
'                    End If
'
'                    If lErro = AD_SQL_SUCESSO Then
'                        bAchouAlm = True
'                        Call Log_Adiciona("O almoxarifado foi localizado pelo código externo. " & "Almoxarifado: " & CStr(iAlmoxarifado) & ".", objLogBase, True)
'                    Else
'                        Call Log_Adiciona("O almoxarifado não foi localizado pelo código externo.", objLogBase, True)
'                    End If
'
'                End If
'
'                If Len(Trim(objIntegracaoPV.sNomeAlmoxarifadoExterno)) <> 0 And Not bAchouAlm Then
'
'                    lErro = Comando_Executar(lComando10, "SELECT Codigo FROM Almoxarifado WHERE NomeReduzido = ? ", iAlmoxarifado, objIntegracaoPV.sNomeAlmoxarifadoExterno)
'                    If lErro <> AD_SQL_SUCESSO Then
'                        bErroNoRegistro = True
'                        Call Log_Adiciona("Erro na leitura do almoxarifado com nome reduzido " & CStr(objIntegracaoPV.sNomeAlmoxarifadoExterno) & ".", objLogBase)
'                    End If
'
'                    lErro = Comando_BuscarProximo(lComando10)
'                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then
'                        bErroNoRegistro = True
'                        Call Log_Adiciona("Erro na leitura do almoxarifado com nome reduzido " & CStr(objIntegracaoPV.sNomeAlmoxarifadoExterno) & ".", objLogBase)
'                    End If
'
'                    If lErro = AD_SQL_SUCESSO Then
'                        bAchouAlm = True
'                        Call Log_Adiciona("O Almoxarifado foi localizado pelo nome reduzido. " & "Almoxarifado: " & CStr(iAlmoxarifado) & ".", objLogBase, True)
'                    Else
'                        Call Log_Adiciona("O Almoxarifado não foi localizado pelo nome reduzido.", objLogBase, True)
'                    End If
'
'                End If
'
'            Else
'                iAlmoxarifado = objIntegracaoPV.iCodAlmoxarifado
'            End If
'
'            If iAlmoxarifado <> 0 Then
'                For Each objReservaItem In objItemPV.ColReserva
'                    objReservaItem.iAlmoxarifado = iAlmoxarifado
'                Next
'            End If
'
'        Next
        
        lErro = objTelaPedido.Gravar_Registro(True)
        If lErro <> SUCESSO Then gError 197679
        
 '   End If
    
'    'If Not bNovoPedido Then
'    If Not bCancelamento Then
'
'        'VERIFICAR AS QUANTIDADE CANCELADAS PARA VER SE O PEDIDO FOI CANCELADO
'        'SEGUNDO O SASCHA É O ÚNICO TIPO DE ALTERAÇÃO POSSÍVEL
'        iSeq = 0
'        For Each objIntegracaoPV In colIntegracaoPV
'            iSeq = iSeq + 1
'            Set objItemPV = objPVTela.colItensPedido.Item(iSeq)
'            objItemPV.dQuantCancelada = objIntegracaoPV.dQuantidadeCancelada
'        Next
'
'    End If
    'End If
    
'    'Verifica se todos os itens estão cancelados
'    bCancelamento = True
'    For Each objItemPV In objPV.colItensPedido
'        If objItemPV.dQuantidade - objItemPV.dQuantCancelada > QTDE_ESTOQUE_DELTA Then
'            bCancelamento = False
'            Exit For
'        End If
'    Next
'
'    If bCancelamento Then
'        objPV.dValorDesconto = 0
'    End If
    
    Importa_Dados_PV_Preenche_Excel = SUCESSO
    
    Exit Function

Erro_Importa_Dados_PV_Preenche_Excel:

    Importa_Dados_PV_Preenche_Excel = gErr

    Select Case gErr
    
        
        Case 197663, 197665, 197667, 197669, 197671, 197672, 197674 To 197679, 197681

        Case 197664
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA2", gErr, objPV.iFilial, objPV.lCliente)
    
        Case 197666
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.lCodigo)

        Case 197668
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.iTipo)

        Case 197670
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO1", gErr, objPV.iFilial, objPV.lCliente)

        Case 197673
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197680)

    End Select
       
    Exit Function

End Function

Function ImportacaoPV_Exclui(objImportacaoPV As ClassImportacaoPV) As Long
'Exclui o registro na tabela ImportacaoPV se existir e o documento associado

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2)  As Long
Dim lCodigoPV As Long
Dim dtData As Date
Dim dHora As Double
Dim iTransacaoCTB As Integer
Dim colLancamentos As New Collection
Dim sSelect As String
Dim objPV As New ClassPedidoDeVenda

On Error GoTo Erro_ImportacaoPV_Exclui
    
    'Abre o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197629
    Next
    
    'Guarda a expressão SQL
    sSelect = "SELECT CodigoPV FROM ImportacaoPV WHERE FilialEmpresa = ? AND Arquivo = ?"
            
    'Grava historico na tabela de CclHistorico
    lErro = Comando_ExecutarPos(alComando(1), sSelect, 0, lCodigoPV, objImportacaoPV.iFilialEmpresa, objImportacaoPV.sArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 197630
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197631
        
    If lErro = AD_SQL_SUCESSO Then
        
        iTransacaoCTB = 75
    
        objPV.lCodigo = lCodigoPV
        objPV.iFilialEmpresa = objImportacaoPV.iFilialEmpresa
    
        'exclui os lancamentos contabeis vinculados a ImportacaoCTB que esta sendo excluida
        lErro = CF("PedidoVenda_Exclui_Trans", objPV)
        If lErro <> SUCESSO Then gError 197632
        
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ImportacaoPV", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 197633
        
    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ImportacaoPV_Exclui = SUCESSO
    
    Exit Function

Erro_ImportacaoPV_Exclui:

    ImportacaoPV_Exclui = gErr
    
    Select Case gErr
    
        Case 197629
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 197630, 197631
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTACAOPV", gErr)
        
        Case 197632
        
        Case 197633
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_IMPORTACAOPV", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197634)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function ImportacaoPV_Grava(objImportacaoPV As ClassImportacaoPV) As Long
'Insere o registro na tabela ImportacaoPV

Dim lErro As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_ImportacaoPV_Grava
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 197635
    
    'Guarda a expressão SQL
    sInsert = "INSERT INTO ImportacaoPV (FilialEmpresa, Arquivo, CodigoPV, Data, Hora, Usuario) VALUES (?,?,?,?,?,?)"
            
    'Grava historico na tabela de CclHistorico
    lErro = Comando_Executar(lComando, sInsert, objImportacaoPV.iFilialEmpresa, objImportacaoPV.sArquivo, objImportacaoPV.lCodigoPV, objImportacaoPV.dtData, CDbl(objImportacaoPV.dtHora), objImportacaoPV.sUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 197636
        
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ImportacaoPV_Grava = SUCESSO
    
    Exit Function

Erro_ImportacaoPV_Grava:

    ImportacaoPV_Grava = gErr
    
    Select Case gErr
    
        Case 197635
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 197636
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IMPORTACAOPV", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197637)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

