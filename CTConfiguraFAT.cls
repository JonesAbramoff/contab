VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguraFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

Dim m_Caption As String
Event Unload()

'VARIAVEIS GLOBAIS A TELA
Dim iAlterado As Integer
Dim iFrameAtual As Integer

'GridDescontos
Dim objGridDescontos As AdmGrid
Dim iGrid_TipoDesconto_Col As Integer
Dim iGrid_Dias_Col As Integer
Dim iGrid_Percentual_Col As Integer


'--- inicio dos properties get dos controles  da tela
Public Property Get TestaDescontoMaxTabPreco() As Object
     Set TestaDescontoMaxTabPreco = objUserControl.Controls("TestaDescontoMaxTabPreco")
End Property

Public Property Get CheckBloquearPrecoBaixoPV() As Object
     Set CheckBloquearPrecoBaixoPV = objUserControl.Controls("CheckBloquearPrecoBaixoPV")
End Property

Public Property Get TotTribTipo() As Object
     Set TotTribTipo = objUserControl.Controls("TotTribTipo")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

'Incluido por Leo em 21/02/02
Public Property Get DiasBloqueio() As Object
     Set DiasBloqueio = objUserControl.Controls("DiasBloqueio")
End Property
'Incluido por Leo em 15/03/02
Public Property Get UsaComissoesRegras() As Object
     Set UsaComissoesRegras = objUserControl.Controls("UsaComissoesRegras")
End Property


Public Property Get ProxNumNFiscal() As Object
     Set ProxNumNFiscal = objUserControl.Controls("ProxNumNFiscal")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get ListaConfigura() As Object
     Set ListaConfigura = objUserControl.Controls("ListaConfigura")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get ReservaManPV() As Object
     Set ReservaManPV = objUserControl.Controls("ReservaManPV")
End Property

Public Property Get ReservaAutoPV() As Object
     Set ReservaAutoPV = objUserControl.Controls("ReservaAutoPV")
End Property

Public Property Get CheckeditaComissoesPV() As Object
     Set CheckeditaComissoesPV = objUserControl.Controls("CheckeditaComissoesPV")
End Property

Public Property Get FilialFaturamento() As Object
     Set FilialFaturamento = objUserControl.Controls("FilialFaturamento")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get CheckeditaComissoesNF() As Object
     Set CheckeditaComissoesNF = objUserControl.Controls("CheckeditaComissoesNF")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get AlocacaoAutoNF() As Object
     Set AlocacaoAutoNF = objUserControl.Controls("AlocacaoAutoNF")
End Property

Public Property Get AlocacaoManNF() As Object
     Set AlocacaoManNF = objUserControl.Controls("AlocacaoManNF")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get TipoDesconto() As Object
     Set TipoDesconto = objUserControl.Controls("TipoDesconto")
End Property

Public Property Get Dias() As Object
     Set Dias = objUserControl.Controls("Dias")
End Property

Public Property Get PercentualDesc() As Object
     Set PercentualDesc = objUserControl.Controls("PercentualDesc")
End Property

Public Property Get GridDescontos() As Object
     Set GridDescontos = objUserControl.Controls("GridDescontos")
End Property

Public Property Get CheckHabilitaFifoNF() As Object
     Set CheckHabilitaFifoNF = objUserControl.Controls("CheckHabilitaFifoNF")
End Property

Public Property Get CheckValidaEmbalagemPV() As Object
     Set CheckValidaEmbalagemPV = objUserControl.Controls("CheckValidaEmbalagemPV")
End Property

Public Property Get RecebForaFaixa() As Object
     Set RecebForaFaixa = objUserControl.Controls("RecebForaFaixa")
End Property

Public Property Get PercentMaisReceb() As Object
     Set PercentMaisReceb = objUserControl.Controls("PercentMaisReceb")
End Property

Public Property Get PercentMenosReceb() As Object
     Set PercentMenosReceb = objUserControl.Controls("PercentMenosReceb")
End Property

Public Property Get NaoTemFaixaReceb() As Object
     Set NaoTemFaixaReceb = objUserControl.Controls("NaoTemFaixaReceb")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get FrameC() As Object
     Set FrameC = objUserControl.Controls("FrameC")
End Property

Public Property Get UsaNFe() As Object
     Set UsaNFe = objUserControl.Controls("UsaNFe")
End Property

Public Property Get GravaNFe() As Object
     Set GravaNFe = objUserControl.Controls("GravaNFe")
End Property

Public Property Get TaxaFinanceira() As Object
     Set TaxaFinanceira = objUserControl.Controls("TaxaFinanceira")
End Property

Public Property Get DespesasDiarias() As Object
     Set DespesasDiarias = objUserControl.Controls("DespesasDiarias")
End Property

Public Property Get VerificaLimCred() As Object
     Set VerificaLimCred = objUserControl.Controls("VerificaLimCred")
End Property

Public Property Get UsaNFSE() As Object
     Set UsaNFSE = objUserControl.Controls("UsaNFSE")
End Property

Public Property Get GravaNFSE() As Object
     Set GravaNFSE = objUserControl.Controls("GravaNFSE")
End Property

Public Property Get versaoNFE() As Object
     Set versaoNFE = objUserControl.Controls("versaoNFE")
End Property

'--- fim dos properties get dos controles da tela

Public Sub AlocacaoAutoNF_GotFocus()
    
    'Se AlocacaoAutoNF não estava com o foco
    If AlocacaoAutoNF.Value = False Then
        
        'Registro Alterado
        iAlterado = REGISTRO_ALTERADO
        
    End If

End Sub

Public Sub AlocacaoManNF_GotFocus()
    
    'Se AlocacaoManNF nãoa estava com o foco
    If AlocacaoManNF.Value = False Then
        
        'Registro Alterado
        iAlterado = REGISTRO_ALTERADO
        
    End If

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Inicializa_Grid_Descontos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Descontos
    
    Set objGridInt.objForm = Me
    
    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Tipo Desconto")
    objGridInt.colColuna.Add ("Dias")
    objGridInt.colColuna.Add ("Percentual")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (TipoDesconto.Name)
    objGridInt.colCampo.Add (Dias.Name)
    objGridInt.colCampo.Add (PercentualDesc.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridDescontos

    'Colunas do Grid
    iGrid_TipoDesconto_Col = 1
    iGrid_Dias_Col = 2
    iGrid_Percentual_Col = 3

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_DESCONTOS + 1
    objGridInt.iLinhasExistentes = NUM_MAXIMO_DESCONTOS
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridDescontos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Descontos = SUCESSO

    Exit Function

End Function

Sub Rotina_Grid_Enable(iLinha As Integer, objControle As Object, iCaminho As Integer)
   
Dim iTipo As Integer

    'Pesquisa a controle da coluna em questão
    Select Case objControle.Name
        
        Case PercentualDesc.Name, Dias.Name
        
            'Se tipo de desconto estiver preenchido
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                
                'Habilita controle
                objControle.Enabled = True

            Else
                
                'Desabilita controle
                objControle.Enabled = False
            
            End If
        
    End Select
    
    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155650)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
           
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    CheckeditaComissoesNF.Value = vbChecked
    CheckeditaComissoesPV.Value = vbChecked
    CheckValidaEmbalagemPV.Value = vbUnchecked
    CheckBloquearPrecoBaixoPV = vbUnchecked
    TestaDescontoMaxTabPreco = vbUnchecked
    
    'Carrega os Tipos de Desconto
    lErro = Carrega_TipoDesconto()
    If lErro <> SUCESSO Then gError 56455
    
    Set objGridDescontos = New AdmGrid
    
    'Incializa o Grid de Descontos
    lErro = Inicializa_Grid_Descontos(objGridDescontos)
    If lErro <> SUCESSO Then gError 56456
    
    'Carrega a Combo de Filiais
    lErro = Carrega_FilialFaturamento()
    If lErro <> SUCESSO Then gError 56474
    
'    If giTipoVersao = VERSAO_LIGHT Then
'        FilialFaturamento.ListIndex = 0
'    End If
    
    'Carrega a Combo de VersaoNFE
    lErro = Carrega_versaoNFE()
    If lErro <> SUCESSO Then gError 207404
    
    'Traz os dados para tela
    lErro = Traz_ConfiguraFAT_Tela()
    If lErro <> SUCESSO Then gError 56457
    
    iAlterado = 0
    iFrameAtual = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 56455, 56456, 56457, 56474, 207404
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155651)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Grava as Alterações
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 56458
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 56458

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155652)

    End Select

    Exit Sub
    
End Sub

Private Function Carrega_TipoDesconto() As Long
'Carrega os Tipos de Desconto

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", STRING_TIPOSDEDESCONTO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 56459

    For Each objCodDescricao In colCodigoDescricao

        'se for desconto do tipo percentual
        If objCodDescricao.iCodigo = Percentual Or objCodDescricao.iCodigo = PERC_ANT_DIA Or objCodDescricao.iCodigo = PERC_ANT_DIA_UTIL Then
        
            'Adiciona o ítem na List da Combo TipoDesconto
            TipoDesconto.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
            TipoDesconto.ItemData(TipoDesconto.NewIndex) = objCodDescricao.iCodigo
            
        End If
        
    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = Err

    Select Case Err

        Case 56459

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155653)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialFaturamento() As Long
'Carrega FilialFaturamento com as Filiais Empresas

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCodigoNome As AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialFaturamento
    
    FilialFaturamento.Text = ""
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_ALMOXARIFADO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 56475

    For Each objCodigoNome In colCodigoDescricao

        'coloca na combo
        FilialFaturamento.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialFaturamento.ItemData(FilialFaturamento.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_FilialFaturamento = SUCESSO

    Exit Function

Erro_Carrega_FilialFaturamento:

    Carrega_FilialFaturamento = Err

    Select Case Err

        'Erro já tratado
        Case 56475

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155654)

    End Select

    Exit Function

End Function

Private Function Traz_ConfiguraFAT_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As ClassSerie
Dim objMatConfig As New ClassMATConfig
Dim lCodFifo As Long
Dim dTaxaFin As Double

On Error GoTo Erro_Traz_ConfiguraFAT_Tela

    'Preenche a List de Contabilização
    ListaConfigura.Selected(FATCONFIG_AGLUTINA_LANCAM_POR_DIA) = (gobjFAT.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA)
    
'    If giTipoVersao <> VERSAO_LIGHT Then
    
        ListaConfigura.Selected(FATCONFIG_GERA_LOTE_AUTOMATICO) = (gobjFAT.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO)
        
'    End If
    
    'Preenche a check Editar Comissões nos Pedidos
    If gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_EDITA_COMISSAO Then
        CheckeditaComissoesPV.Value = vbChecked
    Else
        CheckeditaComissoesPV.Value = vbUnchecked
    End If
    
    'Preenche a check Editar Comissões nos Pedidos
    CheckValidaEmbalagemPV.Value = gobjFAT.iPedidoValidaEmbalagem
    
    CheckBloquearPrecoBaixoPV.Value = gobjCRFAT.iBloquearPVPrecoBaixo
    TestaDescontoMaxTabPreco.Value = gobjCRFAT.iValidaDescMaxTabPreco
    
    'Se a quantidade de dias para o bloqueio de atraso for diferente de Zero.
    If gobjCRFAT.iDiasBloqueioAtraso <> BLOQUEIO_ATRASO_NAO_BLOQUEAR Then 'Trecho por leo em 21/02/02
        
        'Preenche o campo de dias p/ bloqueio por atraso de pagamento.
        DiasBloqueio.Text = gobjCRFAT.iDiasBloqueioAtraso
    
    End If
    
    'Incluido por Leo em 15/03/02
    'Preenche a CheckBox UsaComissoesRegras com o valor guardado em gobjCRFAT.iUsaComissoesRegras(Marcado ou Desmarcado)
    UsaComissoesRegras.Value = gobjCRFAT.iUsaComissoesRegras
    
    UsaNFe.Value = gobjCRFAT.iUsaNFe
    
    GravaNFe.Value = gobjCRFAT.iGravaNFe
    
    UsaNFSE.Value = gobjCRFAT.iUsaNFSE
    
    GravaNFSE.Value = gobjCRFAT.iGravaNFSE
    
    
    If gobjCRFAT.iCreditoVerificaLimite = MARCADO Then
        VerificaLimCred.Value = vbChecked
    Else
        VerificaLimCred.Value = vbUnchecked
    End If
    
    'Preenche a Check grupo Reserva de Produtos
    If gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
        ReservaAutoPV.Value = True
    Else
        ReservaManPV.Value = True
    End If
    
    'Preenche a Check Editar Comissões nas Notas Fiscais
    If gobjFAT.iNFiscalEditaComissao = NFISCAL_EDITA_COMISSAO Then
        CheckeditaComissoesNF.Value = vbChecked
    Else
        CheckeditaComissoesNF.Value = vbUnchecked
    End If
    
    'William
    objMatConfig.iFilialEmpresa = EMPRESA_TODA
    objMatConfig.sCodigo = HABILITA_FIFO_NFISCAIS

    lErro = CF("MATConfig_Le", objMatConfig)
    If lErro <> SUCESSO And lErro <> 89118 Then gError 87447
    
    If lErro = 89118 Then gError 87449

    lCodFifo = StrParaLong(objMatConfig.sConteudo)

    If lCodFifo = HABILITA_FIFO_NF Then
        CheckHabilitaFifoNF.Value = vbChecked
    Else
        CheckHabilitaFifoNF.Value = vbUnchecked
    End If

    'Preenche a Check grupo Alocacao de Produtos
    If gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
        AlocacaoAutoNF.Value = True
    Else
        AlocacaoManNF.Value = True
    End If
    
    DespesasDiarias.Caption = Format(gobjCRFAT.dTaxaFinanceira, "##0.00##%")
        
    'Calcula a taxa Mensal a partir da taxa diária utilizando juros compostos
    Call Calcula_TaxaFinMensal(dTaxaFin)
    
    TaxaFinanceira.Text = Format(dTaxaFin, "Fixed")
    
    'Preenche o Grid de Descontos
    For iIndice = 0 To TipoDesconto.ListCount - 1
        
        If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo1 Then
            GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
            GridDescontos.TextMatrix(1, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias1
            GridDescontos.TextMatrix(1, iGrid_Percentual_Col) = Format(gobjCRFAT.dDescontoPerc1, "Percent")
        End If
        
        If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo2 Then
            GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
            GridDescontos.TextMatrix(2, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias2
            GridDescontos.TextMatrix(2, iGrid_Percentual_Col) = Format(gobjCRFAT.dDescontoPerc2, "Percent")
        End If
        
        If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo3 Then
            GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
            GridDescontos.TextMatrix(3, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias3
            GridDescontos.TextMatrix(3, iGrid_Percentual_Col) = Format(gobjCRFAT.dDescontoPerc3, "Percent")
        End If
        
    Next
    
    If gobjFAT.iFilialFaturamento <> 0 Then
    
        'Preenche a Filial Faturamento
        FilialFaturamento.Text = gobjFAT.iFilialFaturamento
        Call FilialFaturamento_Validate(bSGECancelDummy)
    
    End If
    
    'Faixa de Recebimento
    If gobjFAT.iTemFaixaReceb = 0 Then
        NaoTemFaixaReceb.Value = vbUnchecked
    Else
        NaoTemFaixaReceb.Value = vbChecked
    End If
    
    PercentMenosReceb.Text = (gobjFAT.dPercentMenosReceb) * 100
    PercentMaisReceb.Text = (gobjFAT.dPercentMaisReceb) * 100

    RecebForaFaixa(gobjFAT.iRecebForaFaixa).Value = True
    
    ListaConfigura.Selected(FATCONFIG_INVENTARIOCODBARRAAUTO) = (gobjFAT.iPedVendaCodBarra = 1)
    
    For iIndice = 0 To versaoNFE.ListCount - 1
        If versaoNFE.ItemData(iIndice) = gobjCRFAT.iVersaoNFE Then
            versaoNFE.ListIndex = iIndice
            Exit For
       End If
    Next
    
    Call Combo_Seleciona_ItemData(TotTribTipo, gobjCRFAT.iTotTribTipo)
    
    iAlterado = 0
    
    Traz_ConfiguraFAT_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_ConfiguraFAT_Tela:

    Traz_ConfiguraFAT_Tela = gErr
    
    Select Case gErr
        
        Case 62135
        
        Case 62136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, SERIE_VERSAO_LIGHT)
        
        Case 87447
        
        Case 87449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FIFO_INEXISTENTE", gErr, objMatConfig.sCodigo, objMatConfig.iFilialEmpresa)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155655)
            
    End Select
    
    Exit Function
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer, iMaiorLinhaPreenchida As Integer
Dim objSerie As ClassSerie
Dim objMatConfig As New ClassMATConfig
Dim lCodFifo As Long
Dim sDespesa As String

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    iMaiorLinhaPreenchida = 0
    
    'Verifica qual foi a última linha preenchida no Grid de Descontos
    For iIndice = 1 To objGridDescontos.iLinhasExistentes
     
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) <> 0 _
            Or Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Dias_Col))) <> 0 _
            Or Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Percentual_Col))) <> 0 Then iMaiorLinhaPreenchida = iIndice
    
    Next
            
    'Verifica se foi deixado de Preencher algum campo até a última linha preenchida
    For iIndice = 1 To iMaiorLinhaPreenchida
     
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) = 0 Then Error 52987
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Dias_Col))) = 0 Then Error 52988
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Percentual_Col))) = 0 Then Error 52989
    
    Next
    
    'se a linha 3 está preenchida
    If Len(Trim(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))) <> 0 Then
    
        'a qtde de dias e o percentual tem que ser estritamente decrescentes no grid para que os descontos fiquem ordenados
        If PercentParaDbl(GridDescontos.TextMatrix(3, iGrid_Percentual_Col)) >= PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_Percentual_Col)) Then Error 52984
        If StrParaInt(GridDescontos.TextMatrix(3, iGrid_Dias_Col)) >= StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col)) Then Error 52985
        
    End If
    
    'se a linha 2 está preenchida
    If Len(Trim(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))) <> 0 Then
    
        'a qtde de dias e o percentual tem que ser estritamente decrescentes no grid para que os descontos fiquem ordenados
        If PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_Percentual_Col)) >= PercentParaDbl(GridDescontos.TextMatrix(1, iGrid_Percentual_Col)) Then Error 52986
        If StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col)) >= StrParaInt(GridDescontos.TextMatrix(1, iGrid_Dias_Col)) Then Error 56704
        
    End If
    
    'Move os itens da Lista de Contablização para o gobjFat
    If ListaConfigura.Selected(FATCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
        gobjFAT.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
    Else
        gobjFAT.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
    End If
    
'    If giTipoVersao <> VERSAO_LIGHT Then
    
        If ListaConfigura.Selected(FATCONFIG_GERA_LOTE_AUTOMATICO) = True Then
            gobjFAT.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
        Else
            gobjFAT.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
    
'    End If
    
    'Passa a FilialFaturamento pra gobjFat
    gobjFAT.iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
    
    'Passa a QTD de dias p/ bloqueio por de falta de pagamento pra gobjFat
    gobjCRFAT.iDiasBloqueioAtraso = StrParaInt(DiasBloqueio.Text) 'Por Leo em 21/02/02
    
    'Incluido por Leo em 15/03/02
    'Carrega gobjCRFAT.iUsaComissoesRegras com o valor retirado da tela (marcado ou desmarcado)
    gobjCRFAT.iUsaComissoesRegras = UsaComissoesRegras.Value
    
    gobjCRFAT.iUsaNFe = UsaNFe.Value
    
    gobjCRFAT.iGravaNFe = GravaNFe.Value
    
    gobjCRFAT.iUsaNFSE = UsaNFSE.Value
    
    gobjCRFAT.iGravaNFSE = GravaNFSE.Value
    
    'Passa valores de Permite Editar Comissões Nos Pedidos p/ gobjFAT
    If CheckeditaComissoesPV.Value = vbChecked Then
        gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_EDITA_COMISSAO
    Else
        gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_NAO_EDITA_COMISSAO
    End If

    If VerificaLimCred.Value = vbChecked Then
        gobjCRFAT.iCreditoVerificaLimite = MARCADO
    Else
        gobjCRFAT.iCreditoVerificaLimite = DESMARCADO
    End If
    gobjFAT.iPedidoValidaEmbalagem = CheckValidaEmbalagemPV.Value
    
    If CheckBloquearPrecoBaixoPV.Value = vbChecked Then
        gobjCRFAT.iBloquearPVPrecoBaixo = MARCADO
    Else
        gobjCRFAT.iBloquearPVPrecoBaixo = DESMARCADO
    End If
    
    If TestaDescontoMaxTabPreco.Value = vbChecked Then
        gobjCRFAT.iValidaDescMaxTabPreco = MARCADO
    Else
        gobjCRFAT.iValidaDescMaxTabPreco = DESMARCADO
    End If

    objMatConfig.iFilialEmpresa = EMPRESA_TODA
    objMatConfig.sCodigo = HABILITA_FIFO_NFISCAIS

    'William
    If CheckHabilitaFifoNF.Value = vbChecked Then
        objMatConfig.sConteudo = CStr(HABILITA_FIFO_NF)
    Else
        objMatConfig.sConteudo = CStr(NAO_HABILITA_FIFO_NF)
    End If

    'Passa valores de Reserva de Produtos p/ gobjFAT
    If ReservaAutoPV.Value = True Then
        gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA
    Else
        gobjFAT.iPedidoReservaAutomatica = PEDVENDA_NAO_RESERVA_AUTOMATICA
    End If

    'Passa valores de Permite Editar Comissões nas Notas Fiscais p/ gobjFAT
    If CheckeditaComissoesNF.Value = vbChecked Then
        gobjFAT.iNFiscalEditaComissao = NFISCAL_EDITA_COMISSAO
    Else
        gobjFAT.iNFiscalEditaComissao = NFISCAL_NAO_EDITA_COMISSAO
    End If

    'Passa valores de Reserva de Produtos p/ gobjFAT
    If AlocacaoAutoNF.Value = True Then
        gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA
    Else
        gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_NAO_ALOCA_AUTOMATICA
    End If
    
    'Verifica se Despesa Financeira foi preenchida
    If Len(Trim(DespesasDiarias.Caption)) <> 0 Then
        'Retira o "%" da Despesa
        sDespesa = Mid(DespesasDiarias.Caption, 1, Len(Trim(DespesasDiarias.Caption)) - 1)
        gobjCRFAT.dTaxaFinanceira = CDbl(sDespesa) / 100
    Else
        gobjCRFAT.dTaxaFinanceira = 0
    End If

    'Verifica se linha do Grid Descontos esta preenchida
    If Len(Trim(GridDescontos.TextMatrix(1, iGrid_Percentual_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(1, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col))) <> 0 Then
        
        'Passa os itens da linha 1 do Grid de Descontos para gobjCRFAT
        gobjCRFAT.dDescontoPerc1 = PercentParaDbl(GridDescontos.TextMatrix(1, iGrid_Percentual_Col))
        gobjCRFAT.iDescontoDias1 = GridDescontos.TextMatrix(1, iGrid_Dias_Col)
        gobjCRFAT.iDescontoCodigo1 = Codigo_Extrai(GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col))
        
        If Len(Trim(GridDescontos.TextMatrix(2, iGrid_Percentual_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(2, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))) <> 0 Then
        
            'Passa os itens da linha 2 do Grid de Descontos para gobjCRFAT
            gobjCRFAT.dDescontoPerc2 = PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_Percentual_Col))
            gobjCRFAT.iDescontoDias2 = GridDescontos.TextMatrix(2, iGrid_Dias_Col)
            gobjCRFAT.iDescontoCodigo2 = Codigo_Extrai(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))
        
            If Len(Trim(GridDescontos.TextMatrix(3, iGrid_Percentual_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(3, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))) <> 0 Then
                
                'Passa os itens da linha 3 do Grid de Descontos para gobjCRFAT
                gobjCRFAT.dDescontoPerc3 = PercentParaDbl(GridDescontos.TextMatrix(3, iGrid_Percentual_Col))
                gobjCRFAT.iDescontoDias3 = GridDescontos.TextMatrix(3, iGrid_Dias_Col)
                gobjCRFAT.iDescontoCodigo3 = Codigo_Extrai(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))
            
            Else 'Se linha 3 nao estiver preenchida
                gobjCRFAT.dDescontoPerc3 = 0
                gobjCRFAT.iDescontoDias3 = 0
                gobjCRFAT.iDescontoCodigo3 = 0
            
            End If
    
        Else    'Se linha 2 nao estiver preenchida
            gobjCRFAT.dDescontoPerc2 = 0
            gobjCRFAT.iDescontoDias2 = 0
            gobjCRFAT.iDescontoCodigo2 = 0
            gobjCRFAT.dDescontoPerc3 = 0
            gobjCRFAT.iDescontoDias3 = 0
            gobjCRFAT.iDescontoCodigo3 = 0
            
        End If
    
    Else    'Se linha 1 nao estiver preenchida
        gobjCRFAT.dDescontoPerc1 = 0
        gobjCRFAT.iDescontoDias1 = 0
        gobjCRFAT.iDescontoCodigo1 = 0
        gobjCRFAT.dDescontoPerc2 = 0
        gobjCRFAT.iDescontoDias2 = 0
        gobjCRFAT.iDescontoCodigo2 = 0
        gobjCRFAT.dDescontoPerc3 = 0
        gobjCRFAT.iDescontoDias3 = 0
        gobjCRFAT.iDescontoCodigo3 = 0
               
    End If
    
    gobjCRFAT.iVersaoNFE = versaoNFE.ItemData(versaoNFE.ListIndex)
    
    gobjCRFAT.iTotTribTipo = TotTribTipo.ItemData(TotTribTipo.ListIndex)
    
    'Faixa de Recebimento
    If NaoTemFaixaReceb.Value = vbChecked Then
        gobjFAT.iTemFaixaReceb = MARCADO
    ElseIf NaoTemFaixaReceb.Value = vbUnchecked Then
        gobjFAT.iTemFaixaReceb = DESMARCADO
    End If
    
    If Len(Trim(PercentMenosReceb.Text)) <> 0 Then
        gobjFAT.dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text) / 100
    Else
        gobjFAT.dPercentMenosReceb = 0
    End If

    If Len(Trim(PercentMaisReceb.Text)) <> 0 Then
        gobjFAT.dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text) / 100
    Else
        gobjFAT.dPercentMaisReceb = 0
    End If

    If RecebForaFaixa(0).Value Then
        gobjFAT.iRecebForaFaixa = REJEITA_RECEBIMENTO
    End If

    If RecebForaFaixa(1).Value Then
        gobjFAT.iRecebForaFaixa = ACEITA_RECEBIMENTO
    End If
    
    
    If ListaConfigura.Selected(FATCONFIG_INVENTARIOCODBARRAAUTO) = True Then
        gobjFAT.iPedVendaCodBarra = 1
    Else
        gobjFAT.iPedVendaCodBarra = 0
    End If
    
    'William
    lErro = CF("MATConfig_Grava", objMatConfig)
    If lErro <> SUCESSO Then gError 87448
    
    'Grava as Configurações em FatConfig
    lErro = gobjFAT.Gravar()
    If lErro <> SUCESSO Then gError 56460
    
    'Grava as Configurações em CRFATConfig
    lErro = gobjCRFAT.Gravar()
    If lErro <> SUCESSO Then gError 56657
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 52984, 52985, 52986, 56704
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_NAO_ORDEM_DECRESCENTE", gErr)
        
        Case 52987
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_TIPODESCONTO_NAO_PRENCHIDO", gErr, iIndice)
        
        Case 52988
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_DIAS_NAO_PRENCHIDO", gErr, iIndice)
        
        Case 52989
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_PERCENTUAL_NAO_PRENCHIDO", gErr, iIndice)
        
        Case 56460, 56657, 62134 'Tratados nas rotinas chamadas
                   
        Case 62133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNUMNFISCAL_NAO_PREENCHIDA", gErr)
        
        Case 87448
            'Tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155656)
    
    End Select

    Exit Function
    
End Function

Public Sub Dias_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Dias_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDescontos)

End Sub

Public Sub Dias_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDescontos)

End Sub

Public Sub Dias_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDescontos.objControle = Dias
    lErro = Grid_Campo_Libera_Foco(objGridDescontos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FilialFaturamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialFaturamento_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub
    
Public Sub FilialFaturamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialFaturamento_Validate

    'Verifica se FilialFaturamento esta preenchido
    If Len(Trim(FilialFaturamento.Text)) = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFaturamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 56476

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then Error 56477

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 56478

    Exit Sub

Erro_FilialFaturamento_Validate:

    Cancel = True


    Select Case Err

        Case 56476

        Case 56477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, iCodigo)

        Case 56478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, FilialFaturamento.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155657)

    End Select

    Exit Sub

End Sub

Public Sub ListaConfigura_ItemCheck(Item As Integer)
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub PercentualDesc_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PercentualDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDescontos)

End Sub

Public Sub PercentualDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDescontos)

End Sub

Public Sub PercentualDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDescontos.objControle = TipoDesconto
    lErro = Grid_Campo_Libera_Foco(objGridDescontos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ReservaAutoPV_GotFocus()

    If ReservaAutoPV.Value = False Then
        
        iAlterado = REGISTRO_ALTERADO
        
    End If

End Sub

Public Sub ReservaManPV_GotFocus()

    If ReservaManPV.Value = False Then
        
        iAlterado = REGISTRO_ALTERADO
        
    End If

End Sub

Public Sub TipoDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDescontos)

End Sub

Public Sub TipoDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDescontos)

End Sub

Public Sub TipoDesconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDescontos.objControle = TipoDesconto
    lErro = Grid_Campo_Libera_Foco(objGridDescontos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_GridDescontos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDescontos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_TipoDesconto_Col
            'Faz a crítica do tipo de desconto
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then Error 56461

        Case iGrid_Dias_Col
            'Faz a crítica de Dias
            lErro = Saida_Celula_Dias(objGridInt)
            If lErro <> SUCESSO Then Error 56462

        Case iGrid_Percentual_Col
            'Faz a crítica do Percentual do desconto
            lErro = Saida_Celula_Percentual(objGridInt)
            If lErro <> SUCESSO Then Error 56463

    End Select

    Saida_Celula_GridDescontos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDescontos:

    Saida_Celula_GridDescontos = Err

    Select Case Err

        Case 56461, 56462, 56463

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155658)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_TipoDesconto

    Set objGridInt.objControle = TipoDesconto

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoDesconto.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoDesconto.ListIndex = -1 Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoDesconto, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then Error 56465
        
            'Não foi encontrado
            If lErro = 25085 Then Error 56466
            
            If lErro = 25086 Then Error 56467

        End If
            
        'Acrescenta uma linha no Grid se for o caso
        If GridDescontos.Row = objGridInt.iLinhasExistentes And objGridInt.iLinhasExistentes < 3 Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    Else
        
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Percentual_Col) = ""
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Dias_Col) = ""
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56468

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = Err

    Select Case Err

        Case 56465
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", Err, TipoDesconto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56468
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155659)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = PercentualDesc

    If Len(Trim(PercentualDesc.ClipText)) > 0 Then

        'Verifica se o Percentual foi preenchido
        lErro = Porcentagem_Critica(PercentualDesc.Text)
        If lErro <> SUCESSO Then Error 56469

        'Formata o Percentual
        PercentualDesc.Text = Format(PercentualDesc.Text, "Fixed")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56470

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err

    Select Case Err

        Case 56469, 56470
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155660)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        lErro = Saida_Celula_GridDescontos(objGridInt)
        If lErro <> SUCESSO Then Error 56471
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 56472

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 56471
        
        Case 56472
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155661)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Dias(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Dias

    Set objGridInt.objControle = Dias

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56473

    Saida_Celula_Dias = SUCESSO

    Exit Function

Erro_Saida_Celula_Dias:

    Saida_Celula_Dias = Err

    Select Case Err

        Case 56473
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155662)

    End Select

    Exit Function

End Function

Public Sub GridDescontos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDescontos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDescontos, iAlterado)
    End If

End Sub

Public Sub GridDescontos_EnterCell()

    Call Grid_Entrada_Celula(objGridDescontos, iAlterado)

End Sub

Public Sub GridDescontos_GotFocus()

    Call Grid_Recebe_Foco(objGridDescontos)

End Sub

Public Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDescontos)
     
End Sub

Public Sub GridDescontos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDescontos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDescontos, iAlterado)
    End If

End Sub
Public Sub GridDescontos_LeaveCell()

    Call Saida_Celula(objGridDescontos)

End Sub

Public Sub GridDescontos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridDescontos)
    
End Sub

Public Sub GridDescontos_RowColChange()

    Call Grid_RowColChange(objGridDescontos)

End Sub

Public Sub GridDescontos_Scroll()

    Call Grid_Scroll(objGridDescontos)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridDescontos = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIGURA_FAT
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Faturamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraFAT"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub



Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******





Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Sub ProxNumNFiscal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ProxNumNFiscal_GotFocus()
        
    Call MaskEdBox_TrataGotFocus(ProxNumNFiscal, iAlterado)

End Sub

Public Sub ProxNumNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProxNumNFiscal_Validate

    'Verifica se foi preenchido o campo ProxNumNFiscal
    If Len(Trim(ProxNumNFiscal.Text)) = 0 Then Exit Sub

    lErro = Long_Critica(ProxNumNFiscal.Text)
    If lErro <> SUCESSO Then Error 22178

    Exit Sub

Erro_ProxNumNFiscal_Validate:

    Cancel = True


    Select Case Err

        Case 22178

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155663)

    End Select

    Exit Sub
    
End Sub

'Incluida por leo em 21/02/02
Public Sub DiasBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Incluida por Leo me 15/03/02
Public Sub UsaComissoesRegras_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub NaoTemFaixaReceb_Click()

Dim lErro As Long

On Error GoTo Erro_NaoTemFaixaReceb_Click

    'Verifica valor na checkbox
    If NaoTemFaixaReceb.Value = False Then

        'Habilita os controles
        PercentMaisReceb.Enabled = True
        PercentMenosReceb.Enabled = True
        RecebForaFaixa(0).Enabled = True
        RecebForaFaixa(1).Enabled = True

    Else

        'Desabilita os controles
        PercentMaisReceb.Enabled = False
        PercentMenosReceb.Enabled = False
        RecebForaFaixa(0).Enabled = False
        RecebForaFaixa(1).Enabled = False

    End If
        
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_NaoTemFaixaReceb_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154686)

    End Select

    Exit Sub

End Sub


Public Sub PercentMaisReceb_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMaisReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMaisReceb, iAlterado)
    
End Sub

Public Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMaisReceb As Double

On Error GoTo Erro_PercentMaisReceb_Validate

    'Verifica se PercentMaisRecebe foi preenchido
    If Len(Trim(PercentMaisReceb.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMaisReceb.Text)
        If lErro <> SUCESSO Then gError 49340

        dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text)

        'Coloca o valor no formato fixed da tela
        PercentMaisReceb.Text = Format(dPercentMaisReceb, "Fixed")

    End If

    Exit Sub

Erro_PercentMaisReceb_Validate:

    Cancel = True
    
    Select Case gErr

        Case 49340

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154688)

    End Select

    Exit Sub

End Sub

Public Sub PercentMenosReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMenosReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMenosReceb, iAlterado)
    
End Sub

Public Sub PercentMenosReceb_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMenosReceb As Double

On Error GoTo Erro_PercentMenosReceb_Validate

    'Verifica se PercentMenosRecebe foi preenchido
    If Len(Trim(PercentMenosReceb.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMenosReceb.Text)
        If lErro <> SUCESSO Then gError 49342

        dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text)

        'Verifica se porcentagem igual a 100%
        If dPercentMenosReceb = 100 Then gError 49343

        'Coloca o valor no formato fixed da tela
        PercentMenosReceb.Text = Format(dPercentMenosReceb, "Fixed")

    End If

    Exit Sub

Erro_PercentMenosReceb_Validate:

    Cancel = True

    Select Case gErr

        Case 49342

        Case 49343
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154690)

    End Select

    Exit Sub

End Sub

Public Sub RecebForaFaixa_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    FrameC(TabStrip1.SelectedItem.Index).Visible = True

    'Torna Frame atual invisivel
    FrameC(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154692)

    End Select

    Exit Sub

End Sub

Public Sub GravaNFe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UsaNFe_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GravaNFSE_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UsaNFSE_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TaxaFinanceira_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

'1 - Marcio - 08/2000 - incluiu tratamento p/ campo TaxaFinanceira
Public Sub TaxaFinanceira_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dJuros As Double

On Error GoTo Erro_TaxaFinanceira_Validate
    
    'Verifica se percentual foi digitado
    If Len(Trim(TaxaFinanceira.ClipText)) <> 0 Then

        'Critica a porcentagem
        lErro = Porcentagem_Critica(TaxaFinanceira.Text)
        If lErro <> SUCESSO Then gError 500000

        'Calcula o juros Diarios (juros compostos)
        Call Calcula_TaxaFinanceira(CDbl(TaxaFinanceira.Text), dJuros)
        
        'Põe o percentual formatado na tela
        TaxaFinanceira.Text = Format(TaxaFinanceira.Text, "Fixed")
        
        'Formata com 4 casas está correto
        DespesasDiarias.Caption = Format(dJuros / 100, "##0.00##%")
        
    Else
        
        DespesasDiarias.Caption = ""
    
    End If
         
    Exit Sub

Erro_TaxaFinanceira_Validate:

    Cancel = True

    Select Case gErr

        Case 500000

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Sub Calcula_TaxaFinanceira(dTaxaFin As Double, dJuros As Double)
'Calcula taxa financeira diária a partir dos número de dias e da taxa financeira mensal passada
        
    'Calcula os juros compostos ao dia da taxa financeira (ao mês) passada
    'TaxaFin ao dia = (1 + TaxaFin ao mês) elevado a (1 / 30 dias)-1
    dJuros = (((1 + (dTaxaFin / 100)) ^ (1 / 30)) - 1) * 100

End Sub

Sub Calcula_TaxaFinMensal(dTaxaFin As Double)
    
    'Calcula a taxa Mensal a partir da taxa diária utilizando juros compostos
    dTaxaFin = (((1 + gobjCRFAT.dTaxaFinanceira) ^ 30) - 1) * 100

End Sub

Public Sub VerificaLimCred_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_versaoNFE() As Long
'Carrega versaoNFE

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCodigoNome As AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_versaoNFE
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("Cod_Nomes_Le", "versaoNFE", "Codigo", "Descricao", STRING_VERSAONFE_DESCRICAO, colCodigoDescricao, "Ativo=1")
    If lErro <> SUCESSO Then gError 207402

    For Each objCodigoNome In colCodigoDescricao

        'coloca na combo
        versaoNFE.AddItem objCodigoNome.sNome
        versaoNFE.ItemData(versaoNFE.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_versaoNFE = SUCESSO

    Exit Function

Erro_Carrega_versaoNFE:

    Carrega_versaoNFE = gErr

    Select Case gErr

        'Erro já tratado
        Case 207402

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207403)

    End Select

    Exit Function

End Function

Public Sub TotTribTipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoConfigOutras_Click()
    Call Chama_Tela("ConfigOutros", "FAT")
End Sub

