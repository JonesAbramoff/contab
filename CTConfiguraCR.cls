VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguraCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'obs.: Não transferir.
Private Const CRCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Private Const CRCONFIG_GERA_LOTE_AUTOMATICO = 1

'Constantes que verificam se os Valores dos campos foram alterados
Dim iAlterado As Integer
Private iPercMultaAlterado As Integer
Private iPercJurosAlterado As Integer

'GridDescontos
Dim objGridDescontos As AdmGrid
Dim iGrid_TipoDesconto_Col As Integer
Dim iGrid_Dias_Col As Integer
Dim iGrid_PercentualDesc_Col As Integer

'--- inicio dos properties get dos controles da tela

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get OptionCobranca() As Object
     Set OptionCobranca = objUserControl.Controls("OptionCobranca")
End Property

Public Property Get ComboFilialCobr() As Object
     Set ComboFilialCobr = objUserControl.Controls("ComboFilialCobr")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
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

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get JurosMensais() As Object
     Set JurosMensais = objUserControl.Controls("JurosMensais")
End Property

Public Property Get JurosDiarios() As Object
     Set JurosDiarios = objUserControl.Controls("JurosDiarios")
End Property

Public Property Get PercMulta() As Object
     Set PercMulta = objUserControl.Controls("PercMulta")
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

Public Property Get NumFatura() As Object
     Set NumFatura = objUserControl.Controls("NumFatura")
End Property

Public Property Get NomeDiretorioBoleto() As Object
     Set NomeDiretorioBoleto = objUserControl.Controls("NomeDiretorioBoleto")
End Property

'--- fim dos properties get dos controles da tela

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
    iGrid_PercentualDesc_Col = 3

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_DESCONTOS + 1
    objGridInt.iLinhasExistentes = NUM_MAXIMO_DESCONTOS
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridDescontos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Descontos = SUCESSO

    Exit Function

End Function

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155625)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
           
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Carrega os Tipos de Desconto
    lErro = Carrega_TipoDesconto()
    If lErro <> SUCESSO Then Error 56431
        
    'Inicializa o GRID
    Set objGridDescontos = New AdmGrid
    
    lErro = Inicializa_Grid_Descontos(objGridDescontos)
    If lErro <> SUCESSO Then Error 56430
    
    lErro = Carrega_FiliaisEmpresa()
    If lErro <> SUCESSO Then Error 59256
    
    'Traz os dados para tela
    lErro = Traz_ConfiguraCR_Tela()
    If lErro <> SUCESSO Then Error 56433
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
 
        Case 56430, 56431, 56433, 59256 'Tratados nas Rotinas chamadas
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155626)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 41880
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 41880 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155627)

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
    If lErro <> SUCESSO Then Error 56432

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

        Case 56432 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155628)

    End Select

    Exit Function

End Function
    
Private Function Traz_ConfiguraCR_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sConteudo As String

On Error GoTo Erro_Traz_ConfiguraCR_Tela

    'ListaConfigura
    ListaConfigura.Selected(CRCONFIG_AGLUTINA_LANCAM_POR_DIA) = (gobjCR.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA)
    
'    If giTipoVersao <> VERSAO_LIGHT Then
    
        ListaConfigura.Selected(CRCONFIG_GERA_LOTE_AUTOMATICO) = (gobjCR.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO)
    
'    End If
    
    If gobjCR.iFilialCentralizadora = 0 Then
        OptionCobranca(1).Value = True
    Else
        OptionCobranca(2).Value = True
                
        For iIndice = 0 To ComboFilialCobr.ListCount - 1
        
            If ComboFilialCobr.ItemData(iIndice) = gobjCR.iFilialCentralizadora Then
            
                ComboFilialCobr.ListIndex = iIndice
                Exit For
                
            End If
            
        Next
        
    End If
    
    ' Frame Para Atrasos de Pagamento
    PercMulta.Text = CStr(gobjCRFAT.dPercMulta * 100)
    JurosDiarios.Caption = Format(gobjCRFAT.dPercJurosDiario, "##0.00##%")
    JurosMensais.Text = Round((gobjCRFAT.dPercJurosDiario * 3000), 2)
    NomeDiretorioBoleto.Text = gobjCRFAT.sDirBoletoGer
    
    'Tipo Desconto
    For iIndice = 0 To TipoDesconto.ListCount - 1
        
        If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo1 Then
            GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
            GridDescontos.TextMatrix(1, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias1
            GridDescontos.TextMatrix(1, iGrid_PercentualDesc_Col) = Format(gobjCRFAT.dDescontoPerc1, "Percent")
        End If
        
        If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo2 Then
            GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
            GridDescontos.TextMatrix(2, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias2
            GridDescontos.TextMatrix(2, iGrid_PercentualDesc_Col) = Format(gobjCRFAT.dDescontoPerc2, "Percent")
        End If
        
        If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo3 Then
            GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
            GridDescontos.TextMatrix(3, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias3
            GridDescontos.TextMatrix(3, iGrid_PercentualDesc_Col) = Format(gobjCRFAT.dDescontoPerc3, "Percent")
        End If
        
    Next

'    If giTipoVersao = VERSAO_FULL Then
    
        'Obtem o Próximo Número da Fatura
        lErro = CF("CRFatConfig_Le", CRFATCFG_FATURA_NUM_PROX, EMPRESA_TODA, sConteudo)
        If lErro <> SUCESSO Then Error 61460
        
        'Coloca na Tela
        NumFatura.PromptInclude = False
        NumFatura.Text = sConteudo
        NumFatura.PromptInclude = True
    
'    End If
    
    iAlterado = 0
    iPercMultaAlterado = 0
    iPercJurosAlterado = 0
    
    Traz_ConfiguraCR_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_ConfiguraCR_Tela:

    Traz_ConfiguraCR_Tela = Err
    
    Select Case Err
        
        Case 61460
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155629)
            
    End Select
    
    Exit Function
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer, iMaiorLinhaPreenchida As Integer
Dim sJuros As String

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    iMaiorLinhaPreenchida = 0
    
    For iIndice = 1 To objGridDescontos.iLinhasExistentes
     
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) <> 0 _
            Or Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Dias_Col))) <> 0 _
            Or Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_PercentualDesc_Col))) <> 0 Then iMaiorLinhaPreenchida = iIndice
    
    Next
    
    'ou os tres campos de uma linha estao preenchidos ou nenhum deles está.
    'Nao pode haver "buraco" (ex.: preencher linha 2 sem preencher a um)
    For iIndice = 1 To iMaiorLinhaPreenchida
     
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) = 0 Then Error 56711
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Dias_Col))) = 0 Then Error 56712
        If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_PercentualDesc_Col))) = 0 Then Error 56713
    
    Next
    
    'se a linha 3 está preenchida
    If Len(Trim(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))) <> 0 Then
    
        'a qtde de dias e o percentual tem que ser estritamente decrescentes no grid para que os descontos fiquem ordenados
        If PercentParaDbl(GridDescontos.TextMatrix(3, iGrid_PercentualDesc_Col)) >= PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_PercentualDesc_Col)) Then Error 56707
        If StrParaInt(GridDescontos.TextMatrix(3, iGrid_Dias_Col)) >= StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col)) Then Error 56708
        
    End If
    
    'se a linha 2 está preenchida
    If Len(Trim(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))) <> 0 Then
    
        'a qtde de dias e o percentual tem que ser estritamente decrescentes no grid para que os descontos fiquem ordenados
        If PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_PercentualDesc_Col)) >= PercentParaDbl(GridDescontos.TextMatrix(1, iGrid_PercentualDesc_Col)) Then Error 56709
        If StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col)) >= StrParaInt(GridDescontos.TextMatrix(1, iGrid_Dias_Col)) Then Error 56710
        
    End If
    
    If ListaConfigura.Selected(CRCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
        gobjCR.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
    Else
        gobjCR.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
    End If
    
'    If giTipoVersao <> VERSAO_LIGHT Then
    
        If ListaConfigura.Selected(CRCONFIG_GERA_LOTE_AUTOMATICO) = True Then
            gobjCR.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
        Else
            gobjCR.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
    
'    End If
    
    'Verifica se PercJuros foi preenchido
    If Len(Trim(JurosDiarios.Caption)) <> 0 Then
        'Retira o "%" do Juros
        sJuros = Mid(JurosDiarios.Caption, 1, Len(Trim(JurosDiarios.Caption)) - 1)
        gobjCRFAT.dPercJurosDiario = CDbl(sJuros) / 100
    Else
        gobjCRFAT.dPercJurosDiario = 0
    End If
    
    'Verifica se PercMulta foi preenchido
    If Len(Trim(PercMulta.Text)) <> 0 Then
        gobjCRFAT.dPercMulta = CDbl(PercMulta.Text) / 100
    Else
        gobjCRFAT.dPercMulta = 0
    End If
    
    gobjCRFAT.sDirBoletoGer = NomeDiretorioBoleto.Text
    
    'Verifica se linha do Grid Descontos esta preenchida
    If Len(Trim(GridDescontos.TextMatrix(1, iGrid_PercentualDesc_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(1, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col))) <> 0 Then
        gobjCRFAT.dDescontoPerc1 = PercentParaDbl(GridDescontos.TextMatrix(1, iGrid_PercentualDesc_Col))
        gobjCRFAT.iDescontoDias1 = GridDescontos.TextMatrix(1, iGrid_Dias_Col)
        gobjCRFAT.iDescontoCodigo1 = Codigo_Extrai(GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col))
        
        If Len(Trim(GridDescontos.TextMatrix(2, iGrid_PercentualDesc_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(2, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))) <> 0 Then
            gobjCRFAT.dDescontoPerc2 = PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_PercentualDesc_Col))
            gobjCRFAT.iDescontoDias2 = GridDescontos.TextMatrix(2, iGrid_Dias_Col)
            gobjCRFAT.iDescontoCodigo2 = Codigo_Extrai(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))
        
            If Len(Trim(GridDescontos.TextMatrix(3, iGrid_PercentualDesc_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(3, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))) <> 0 Then
                gobjCRFAT.dDescontoPerc3 = PercentParaDbl(GridDescontos.TextMatrix(3, iGrid_PercentualDesc_Col))
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
    
    If OptionCobranca(2).Value = True And ComboFilialCobr.ListIndex = -1 Then Error 59257
    
    If ComboFilialCobr.ListIndex <> -1 Then
        gobjCR.iFilialCentralizadora = ComboFilialCobr.ItemData(ComboFilialCobr.ListIndex)
    Else
        gobjCR.iFilialCentralizadora = 0
    End If
                
'    If giTipoVersao = VERSAO_FULL Then
        'Grava em CRFATConfig o NumProx Fatura
        lErro = CF("CRFATConfig_Grava", CRFATCFG_FATURA_NUM_PROX, EMPRESA_TODA, NumFatura.Text)
        If lErro <> SUCESSO Then Error 61461
'    End If
    
    lErro = gobjCR.Gravar()
    If lErro <> SUCESSO Then Error 41881
    
    lErro = gobjCRFAT.Gravar()
    If lErro <> SUCESSO Then Error 56655
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 41881, 56655, 61461 'Tratados nos métodos chamados
            
        Case 56707, 56708, 56709, 56710
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_NAO_ORDEM_DECRESCENTE", Err)
        
        Case 56711
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_TIPODESCONTO_NAO_PRENCHIDO", Err, iIndice)
        
        Case 56712
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_DIAS_NAO_PRENCHIDO", Err, iIndice)
        
        Case 56713
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_PERCENTUAL_NAO_PRENCHIDO", Err, iIndice)
        
        Case 59257
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_CENTR_COBR_NAO_SELEC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155630)
    
    End Select

    Exit Function
    
End Function

Public Sub Dias_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumFatura_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumFatura, iAlterado)

End Sub

Public Sub NumFatura_Change()
    
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

Public Sub ListaConfigura_ItemCheck(Item As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub OptionCobranca_Click(Index As Integer)
    
    If Index = 1 Then
        ComboFilialCobr.ListIndex = -1
        ComboFilialCobr.Enabled = False
    Else
        ComboFilialCobr.Enabled = True
    End If

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

Public Sub JurosMensais_Change()

    iPercJurosAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub JurosMensais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dJuros As Double

On Error GoTo Erro_JurosMensais_Validate

    If iPercJurosAlterado = 0 Then Exit Sub
    
    'Verifica se percentual foi digitado
    If Len(Trim(JurosMensais.ClipText)) <> 0 Then

        'Critica a porcentagem
        lErro = Porcentagem_Critica(JurosMensais.Text)
        If lErro <> SUCESSO Then Error 56418

        'Calcula o juros Diarios
        dJuros = CDbl(JurosMensais.Text)
        dJuros = dJuros / 30
        
        'Põe o percentual formatado na tela
        JurosMensais.Text = Format(JurosMensais.Text, "Fixed")
        
        'Formata com 4 casas está correto
        JurosDiarios.Caption = Format(dJuros / 100, "##0.00##%")
        
    Else
        
        JurosDiarios.Caption = ""
    
    End If
    
    iPercJurosAlterado = 0
     
    Exit Sub

Erro_JurosMensais_Validate:

    Cancel = True


    Select Case Err

        Case 56418

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155631)

    End Select

    Exit Sub

End Sub

Public Sub PercMulta_Change()

    iPercMultaAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercMulta_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercMulta_Validate

    If iPercMultaAlterado = 0 Then Exit Sub
    
    'Verifica se percentual foi digitado
    If Len(Trim(PercMulta.ClipText)) <> 0 Then

        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercMulta.Text)
        If lErro <> SUCESSO Then Error 56417
        
        'Põe o Percentual formatado na tela
        PercMulta.Text = Format(PercMulta.Text, "Fixed")
    
    End If
    
    iPercMultaAlterado = 0
     
    Exit Sub

Erro_PercMulta_Validate:

    Cancel = True


    Select Case Err

        Case 56417

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155632)

    End Select

    Exit Sub

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
            If lErro <> SUCESSO Then Error 56419

        Case iGrid_Dias_Col
            'Faz a crítica de Dias
            lErro = Saida_Celula_Dias(objGridInt)
            If lErro <> SUCESSO Then Error 56420

        Case iGrid_PercentualDesc_Col
            'Faz a crítica do Percentual do desconto
            lErro = Saida_Celula_Percentual(objGridInt)
            If lErro <> SUCESSO Then Error 56421

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56422

    Saida_Celula_GridDescontos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDescontos:

    Saida_Celula_GridDescontos = Err

    Select Case Err

        Case 56419, 56420, 56421 'Tratados nas Rotinas chamadas

        Case 56422
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155633)

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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then Error 56425
        
            'Não foi encontrado
            If lErro = 25085 Then Error 56426
            
            If lErro = 25086 Then Error 56427

        End If
        
    Else
        
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_PercentualDesc_Col) = ""
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Dias_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56428

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = Err

    Select Case Err

        Case 56425
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", Err, TipoDesconto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56428 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155634)

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
        If lErro <> SUCESSO Then Error 56423

        'Formata o Percentual
        PercentualDesc.Text = Format(PercentualDesc.Text, "Fixed")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56424

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err

    Select Case Err

        Case 56423, 56424
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155635)

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
        If lErro <> SUCESSO Then Error 56434
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 56435

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 56434, 56435 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155636)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Dias(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Dias

    Set objGridInt.objControle = Dias

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56429

    Saida_Celula_Dias = SUCESSO

    Exit Function

Erro_Saida_Celula_Dias:

    Saida_Celula_Dias = Err

    Select Case Err

        Case 56429
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155637)

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

    Parent.HelpContextID = IDH_CONFIGURA_CR
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Contas a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraCR"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
    
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


Function Carrega_FiliaisEmpresa() As Long
'preenche a combo com os nomes das filiais da empresa e coloca no itemdata o codigo da filial

Dim lErro As Long, colFiliais As New Collection, objFilial As AdmFiliais

On Error GoTo Erro_Carrega_FiliaisEmpresa

    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then Error 59255
    
    For Each objFilial In colFiliais
    
        If objFilial.iCodFilial <> EMPRESA_TODA Then
        
            ComboFilialCobr.AddItem objFilial.sNome
            ComboFilialCobr.ItemData(ComboFilialCobr.NewIndex) = objFilial.iCodFilial
            
        End If
    
    Next
    
    Carrega_FiliaisEmpresa = SUCESSO
     
    Exit Function
    
Erro_Carrega_FiliaisEmpresa:

    Carrega_FiliaisEmpresa = Err
     
    Select Case Err
          
        Case 59255
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155638)
     
    End Select
     
    Exit Function

End Function

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

Public Property Get hWnd() As Long
    hWnd = objUserControl.hWnd
End Property

Public Sub BotaoProcurarBoleto_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurarBoleto_Click

    szTitle = "Localização física dos arquivos gerados"
    With tBrowseInfo
        .hWndOwner = hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorioBoleto.Text = sBuffer
        Call NomeDiretorioBoleto_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurarBoleto_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Public Sub NomeDiretorioBoleto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NomeDiretorioBoleto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorioBoleto_Validate

    If Len(Trim(NomeDiretorioBoleto.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorioBoleto.Text, 1) <> "\" And right(NomeDiretorioBoleto.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorioBoleto.Text, "/")
        If iPos = 0 Then
            NomeDiretorioBoleto.Text = NomeDiretorioBoleto.Text & "\"
        Else
            NomeDiretorioBoleto.Text = NomeDiretorioBoleto.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorioBoleto.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_NomeDiretorioBoleto_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorioBoleto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192328)

    End Select

    Exit Sub

End Sub
