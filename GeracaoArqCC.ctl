VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl GeracaoArqCC 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   ScaleHeight     =   4320
   ScaleWidth      =   3930
   Begin VB.ListBox Msgs 
      Height          =   840
      ItemData        =   "GeracaoArqCC.ctx":0000
      Left            =   225
      List            =   "GeracaoArqCC.ctx":0002
      TabIndex        =   8
      Top             =   3315
      Width           =   3420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Leitura de Arquivo"
      Height          =   825
      Left            =   240
      TabIndex        =   3
      Top             =   1860
      Width           =   3405
      Begin VB.CommandButton BotaoLerBack 
         Caption         =   "BackOffice"
         Height          =   345
         Left            =   1935
         TabIndex        =   5
         Top             =   360
         Width           =   1155
      End
      Begin VB.CommandButton BotaoLer 
         Caption         =   "Caixa"
         Height          =   345
         Left            =   495
         TabIndex        =   4
         Top             =   375
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gravação de Arquivo"
      Height          =   1200
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3405
      Begin VB.CheckBox RecalcularTribProd 
         Caption         =   "Recalcular Tributação dos Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   9
         Top             =   855
         Width           =   2865
      End
      Begin VB.CommandButton BotaoGravarBack 
         Caption         =   "BackOffice"
         Height          =   345
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   1155
      End
      Begin VB.CommandButton BotaoGerar 
         Caption         =   "Caixa"
         Height          =   345
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   315
      Left            =   2880
      Picture         =   "GeracaoArqCC.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   780
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   225
      TabIndex        =   7
      Top             =   2820
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "GeracaoArqCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declarações Globais
Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoGravarBack_Click()
    
    Call Chama_Tela("OperacaoArqCCBack")
    
End Sub

Private Sub BotaoLerBack_Click()
'le as informações do backoffice e grava no caixa central

Dim lErro As Long
Dim objBarraProgresso As Object
Dim sNomeArqParam As String
Dim objObject As Object
Dim lIntervaloTrans As Long

On Error GoTo Erro_BotaoLerBack_Click
    
     If Len(Trim(gobjLoja.sFTPURL)) > 0 Then
            
        lIntervaloTrans = gobjLoja.lIntervaloTrans
            
        'colocou 0 para so enviar para o servidor FPT e sair fora
        gobjLoja.lIntervaloTrans = 0
            
        'Prepara para chamar rotina batch
        lErro = Sistema_Preparar_Batch(sNomeArqParam)
        If lErro <> SUCESSO Then gError 133616
            
        gobjLoja.sNomeArqParam = sNomeArqParam
            
        Set objObject = gobjLoja
            
        Msgs.AddItem RECEPCAO_INICIADA
            
        lErro = CF("Rotina_FTP_Recepcao_CC", objObject, 2)
        If lErro <> SUCESSO And lErro <> 133628 Then gError 133615
            
        If lErro <> SUCESSO Then gError 133633
            
        Msgs.AddItem RECEPCAO_CONCLUIDA
            
        gobjLoja.lIntervaloTrans = lIntervaloTrans
        
    End If
    
    Set objBarraProgresso = BarraProgresso
    
    lErro = CF("Rotina_Carga_Back_Caixa_Central", objBarraProgresso)
    If lErro <> SUCESSO Then gError 133614
    
    Call Rotina_Aviso(vbOK, "AVISO_ARQUIVO_CARREGADO", glEmpresa & "_" & giFilialEmpresa & "_" & NOME_ARQUIVOBACK)
    
    Exit Sub
    
Erro_BotaoLerBack_Click:
    
    Select Case gErr
    
        Case 133614 To 133616
    
        Case 133633
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_CARREGOU_ROTINA_RECEPCAO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160713)
        
    End Select
    
    Call Transacao_Rollback
    
    Exit Sub
    
End Sub

Private Sub Caixa_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLer_Click()
    
Dim sNomeArq As String
Dim lErro As Long
Dim objArq As New AdmCodigoNome
Dim objBarraProgresso As Object
Dim objLojaConfig As New ClassLojaConfig
Dim sFile As String
Dim sDir As String
Dim sDir1 As String

On Error GoTo Erro_BotaoLer_Click

    Call Chama_Tela_Modal("ExibirArquivos", objArq)
    
    If giRetornoTela = vbOK Then
    
        sNomeArq = objArq.sNome
        
        sFile = Dir(sNomeArq)
        
        Set objBarraProgresso = BarraProgresso
        
        lErro = CF("Rotina_Carga_ECF_Caixa_Central", sNomeArq, objBarraProgresso)
        If lErro <> SUCESSO Then gError 112725
    
        objLojaConfig.iFilialEmpresa = EMPRESA_TODA
        objLojaConfig.sCodigo = DIRETORIO_TELA_EXIBIRARQUIVOS
        
        lErro = CF("LojaConfig_Le1", objLojaConfig)
        If lErro <> SUCESSO And lErro <> 126361 Then gError 214160
    
        'se nao encontrou o registro q armazena o ultimo diretorio acessado para esta tela
        If lErro = 126361 Then objLojaConfig.sConteudo = "."
    
        sDir = objLojaConfig.sConteudo & IIf(Right(objLojaConfig.sConteudo, 1) <> "\", "\", "")
    
        sDir1 = Dir(sDir & "back\" & sFile)

        If Len(sDir1) > 0 Then
            Kill sDir & "back\" & sFile
        End If

        Name sNomeArq As sDir & "back\" & sFile
    
    
    
    End If
    
    Exit Sub
        
Erro_BotaoLer_Click:
    
   Select Case gErr

        Case 112725, 214160
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160714)

        End Select
        
    Exit Sub
    
End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim objBarraProgresso As Object
Dim objMsgs As Object


On Error GoTo Erro_BotaoGerar_Click
    
    Set objBarraProgresso = BarraProgresso
    Set objMsgs = Msgs
    
    lErro = CF("GeracaoArqCC_Grava", objBarraProgresso, objMsgs, RecalcularTribProd.Value = vbChecked)
    If lErro <> SUCESSO Then gError 126481
    
    Msgs.Clear
    
    Exit Sub
    
Erro_BotaoGerar_Click:
    
   Select Case gErr

        Case 126481
            Msgs.Clear
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160715)

    End Select
        
    Exit Sub
    
End Sub

Public Sub Trata_Parametros()
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE Then
        BotaoLerBack.Enabled = False
        BotaoGravarBack.Enabled = False
    End If
    
    BarraProgresso.Min = 0
    BarraProgresso.Max = 100
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160716)

    End Select

    Exit Sub

End Sub

'Private Sub Le_Arquivo_Vendas(iIndice As Integer, asReg() As String)
'
'Dim iPOS As Integer
'Dim iIndice2 As Integer
'Dim iPosInicio As Integer
'Dim iPosFim As Integer
'Dim iPosMeio As Integer
'Dim iPosColInicio As Integer
'Dim iPosColFim As Integer
'Dim iPosStrFim As String
'Dim sNomeArq As String
'Dim objCheque As ClassChequePre
'Dim objCarne As ClassCarne
'Dim objMovcx As ClassMovimentoCaixa
'Dim objCarneParc As ClassCarneParcelas
'Dim objItens As ClassItemCupomFiscal
'Dim objCupomFiscal As ClassCupomFiscal
'Dim objOrcamentoLoja As New ClassOrcamentoLoja
'Dim objTroca As ClassTroca
'Dim sTipo As String
'Dim lErro As Long
'Dim lNumIntDoc As Long
'Dim iTipo As Integer
'Dim colMovCaixa As New Collection
'Dim objLog As New ClassLog
'Dim iItem As Integer
'Dim iTipoOrc As Integer
'Dim iPosorc As Integer
'Dim lTransacao As Long
'
'On Error GoTo Erro_Le_Arquivo_Vendas
'
'    'abre a transacao
'    lTransacao = Transacao_Abrir
'    If lTransacao = 0 Then gError 99909
'
'    'Primeira Posição
'    iPosInicio = 1
'
'    'Procura o Primeiro Control para saber onde começa a string
'    iPosInicio = InStr(iPosInicio, asReg(iIndice), Chr(vbKeyControl)) + 1
'
'    'iPosicao1 Guarda a posição do Segundo Control(referente ao tipo)
'    iPosStrFim = InStr(iPosInicio + 1, asReg(iIndice), Chr(vbKeyControl))
'
'    'Recolhe o tipo
'    iTipo = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosStrFim - iPosInicio))
'
'    iPosInicio = iPosStrFim + 1
'
'    'Procura o Primeiro Escape dentro da String
'    iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'    'iPosicao1 Guarda a posição do terceiro Control(referente ao objcupom)
'    iPosColFim = InStr(iPosInicio, asReg(iIndice), Chr(vbKeyControl))
'
'    'iPosicao1 Guarda a posição do inicio da colparcelas
'    iPosColInicio = InStr(iPosInicio, asReg(iIndice), Chr(vbKeyShift))
'    If iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'    'Pega última posição e guarda
'    iPosFim = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEnd)))
'
'    iIndice2 = 0
'    Set objCarne = New ClassCarne
'
'    Do While iPosMeio <> 0
'
'        iIndice2 = iIndice2 + 1
'        'Recolhe o objCarne
'        Select Case iIndice2
'
'            Case 1: objCarne.dtDataReferencia = StrParaDate(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 2: objCarne.iFilialEmpresa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 3: objCarne.iStatus = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 4: objCarne.lCupomFiscal = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 5: objCarne.lNumIntDoc = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 6: objCarne.lCliente = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 7: objCarne.sAutorizacao = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'            Case 8: objCarne.sCodBarrasCarne = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'            Case 9: Exit Do
'
'        End Select
'
'        'Atualiza as Posições
'        iPosInicio = iPosMeio + 1
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'    Loop
'
'    If objCarne.dtDataReferencia <> 0 Then
'        'Função que Gera Inteiro Automático
'        lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_CARNE", lNumIntDoc, 1, giFilialEmpresa)
'        If lErro <> SUCESSO Then gError 99780
'
'        objCarne.lNumIntDoc = lNumIntDoc
'    End If
'
'    Do While iPosInicio < iPosColFim
'
'        iPosInicio = iPosColInicio + 1
'
'        'Procura o Primeiro Escape dentro da String
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        'Procura um shift dentro da String
'        iPosColInicio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyShift)))
'        If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'        iIndice2 = 0
'        Set objCarneParc = New ClassCarneParcelas
'
'        Do While iPosMeio <> 0
'
'            iIndice2 = iIndice2 + 1
'            'Recolhe as parcelas
'            Select Case iIndice2
'
'                Case 1: objCarneParc.dtDataVencimento = StrParaDate(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 2: objCarneParc.dValor = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 3: objCarneParc.iFilialEmpresa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 4: objCarneParc.iParcela = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 5: objCarneParc.iStatus = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 6: objCarneParc.lNumIntCarne = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 7: objCarneParc.lNumIntDoc = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 8: Exit Do
'
'            End Select
'
'            'Atualiza as Posições
'            iPosInicio = iPosMeio + 1
'            iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'            If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'        Loop
'
'        objCarneParc.lNumIntCarne = objCarne.lNumIntDoc
'
'        If objCarneParc.lNumIntCarne <> 0 Then
'            'Função que Gera Inteiro Automático
'            lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_CARNEPARC", lNumIntDoc, 1, giFilialEmpresa)
'            If lErro <> SUCESSO Then gError 99781
'
'            objCarneParc.lNumIntDoc = lNumIntDoc
'        End If
'
'        objCarne.colParcelas.Add objCarneParc
'    Loop
'
'    iPosInicio = iPosColFim + 1
'
'    'Procura o Primeiro Escape dentro da String
'    iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'    'iPosicao1 Guarda a posição do quarto Control(referente ao colcheques)
'    iPosColFim = InStr(iPosInicio + 1, asReg(iIndice), Chr(vbKeyControl))
'
'    'iPosicao1 Guarda a posição do inicio da colitens
'    iPosColInicio = InStr(iPosInicio + 1, asReg(iIndice), Chr(vbKeyShift))
'    If iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'    iIndice2 = 0
'
'    Set objCupomFiscal = New ClassCupomFiscal
'
'    Do While iPosMeio <> 0
'
'        iIndice2 = iIndice2 + 1
'        'Recolhe o objCupom
'        Select Case iIndice2
'
'            Case 1: objCupomFiscal.dtDataEmissao = StrParaDate(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 2: objCupomFiscal.dHoraEmissao = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 3: objCupomFiscal.dValorAcrescimo = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 4: objCupomFiscal.dValorDesconto = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 5: objCupomFiscal.dValorProdutos = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 6: objCupomFiscal.dValorTotal = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 7: objCupomFiscal.dValorTroco = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 8: objCupomFiscal.iECF = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 9: objCupomFiscal.iFilialEmpresa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 10: objCupomFiscal.iStatus = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 11: objCupomFiscal.iTabelaPreco = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'''''            Case 13: objCupomFiscal.lVolumeQuant = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 12: objCupomFiscal.lGerenteCancel = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 13: objCupomFiscal.lNumero = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 14: objCupomFiscal.lNumIntDoc = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 15: objCupomFiscal.lNumOrcamento = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 16: objCupomFiscal.iVendedor = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'            Case 17: objCupomFiscal.sCPFCGC = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'            Case 18: objCupomFiscal.sMotivoCancel = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'            Case 19: objCupomFiscal.sNaturezaOp = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'''''            Case 21: objCupomFiscal.sObservacao = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'''''            Case 20: objCupomFiscal.lDuracao = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'            Case 20: Exit Do
'
'        End Select
'
'        'Atualiza as Posições
'        iPosInicio = iPosMeio + 1
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'    Loop
'
'    If iTipo = OPTION_DAV Or iTipo = OPTION_PREVENDA Then
'        For iIndice2 = 1 To UBound(asReg)
'
'            'Procura o Primeiro Control para saber o tipo do registro
'            iPosorc = InStr(1, asReg(iIndice2), Chr(vbKeyControl))
'
'            iTipoOrc = StrParaInt(Mid(asReg(iIndice2), 1, iPosorc - 1))
'
'            If iTipoOrc = TIPOREGISTROECF_EXCLUSAOORCAMENTO Then
'                If objCupomFiscal.lNumOrcamento = StrParaLong(Mid(asReg(iIndice2), iPosorc + 1, Len(asReg(iIndice2)) - (iPosorc + 1))) Then Exit Sub
'            End If
'        Next
'        'Função que Gera Inteiro Automático
'        lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_ORCAMENTO", lNumIntDoc, 1, giFilialEmpresa)
'        If lErro <> SUCESSO Then gError 99783
'    Else
'        'Função que Gera Inteiro Automático
'        lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_CUPOMFISCAL", lNumIntDoc, 1, giFilialEmpresa)
'        If lErro <> SUCESSO Then gError 99782
'    End If
'
'    objCupomFiscal.lNumIntDoc = lNumIntDoc
'
'    iItem = 1
'    Do While iPosInicio < iPosColFim
'
'        iPosInicio = iPosColInicio + 1
'
'        'Procura o Primeiro Escape dentro da String
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        'Procura um shift dentro da String
'        iPosColInicio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyShift)))
'        If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'        iIndice2 = 0
'        Set objItens = New ClassItemCupomFiscal
'
'        Do While iPosMeio <> 0
'
'            iIndice2 = iIndice2 + 1
'            'Recolhe os itens
'            Select Case iIndice2
'
'                Case 1: objItens.dAliquotaICMS = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 2: objItens.dPercDesc = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 3: objItens.dPrecoUnitario = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 4: objItens.dQuantidade = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 5: objItens.dValorDesconto = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 6: objItens.icancel = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 7: objItens.iStatus = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 8: objItens.lNumIntCupom = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 9: objItens.lNumIntDoc = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 10: objItens.sProduto = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 11: objItens.sUnidadeMed = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 12: Exit Do
'            End Select
'
'            'Atualiza as Posições
'            iPosInicio = iPosMeio + 1
'            iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'            If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'        Loop
'        'Verifica se esse item foi cancelado
'        If objItens.icancel = ITEM_NORMAL Then
'            If iTipo = OPTION_CF Then
'                'Função que Gera Inteiro Automático
'                lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_ITEM_CUPOMFISCAL", lNumIntDoc, 1, giFilialEmpresa)
'                If lErro <> SUCESSO Then gError 99784
'            Else
'                'Função que Gera Inteiro Automático
'                lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_ITEM_ORCAMENTO", lNumIntDoc, 1, giFilialEmpresa)
'                If lErro <> SUCESSO Then gError 99785
'            End If
'
'            objItens.lNumIntDoc = lNumIntDoc
'            objItens.lNumIntCupom = objCupomFiscal.lNumIntDoc
'            objItens.iItem = iItem
'            iItem = iItem + 1
'
'            objCupomFiscal.colItens.Add objItens
'        End If
'    Loop
'
'
'    iPosInicio = iPosColFim + 2
'
'    'iPosicao1 Guarda a posição do sexto Control(referente ao coltroca)
'    iPosColFim = InStr(iPosInicio + 1, asReg(iIndice), Chr(vbKeySeparator))
'
'    Do While iPosInicio < iPosColFim
'
'        'Procura o Primeiro Escape dentro da String
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        'Procura um control dentro da String
'        iPosColInicio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyControl)))
'        If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'        iIndice2 = 0
'        Set objMovcx = New ClassMovimentoCaixa
'
'        Do While iPosMeio <> 0
'
'            iIndice2 = iIndice2 + 1
'            'Recolhe os Movimentos de caixa
'            Select Case iIndice2
'
'                Case 1: objMovcx.dHora = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 2: objMovcx.dtDataMovimento = StrParaDate(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 3: objMovcx.dValor = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 4: objMovcx.iAdmMeioPagto = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 5: objMovcx.iCaixa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 6: objMovcx.iCodConta = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 7: objMovcx.iCodOperador = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 8: objMovcx.iFilialEmpresa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 9: objMovcx.iGerente = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 10: objMovcx.iParcelamento = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 11: objMovcx.iTipo = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 12: objMovcx.iTipoCartao = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 13: objMovcx.lCupomFiscal = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 14: objMovcx.lMovtoEstorno = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 15: objMovcx.lMovtoTransf = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 16: objMovcx.lNumero = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 17: objMovcx.lNumMovto = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 18: objMovcx.lNumRefInterna = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 19: objMovcx.lSequencial = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 20: objMovcx.lSequencialConta = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 21: objMovcx.sFavorecido = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 22: objMovcx.sHistorico = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 23: Exit Do
'
'            End Select
'
'            'Atualiza as Posições
'            iPosInicio = iPosMeio + 1
'            iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'            If iPosMeio > iPosColInicio Then iPosMeio = iPosColInicio
'
'        Loop
'
'        iPosInicio = iPosColInicio + 1
'
'        If objMovcx.iTipo = MOVIMENTOCAIXA_RECEB_CARNE Then
'            For Each objCarneParc In objCarne.colParcelas
'                If objMovcx.lNumMovto = objCarneParc.iParcela Then
'                    objMovcx.lNumero = objCarne.lNumIntDoc
'                    objMovcx.lNumRefInterna = objCarneParc.lNumIntDoc
'                End If
'            Next
'        End If
'
'        objMovcx.lCupomFiscal = objCupomFiscal.lNumIntDoc
'
'        Call CF("MovimentosCaixa_Inserir", objMovcx)
'
'        colMovCaixa.Add objMovcx
'    Loop
'
'    iPosInicio = iPosColFim + 1
'
'    'iPosicao1 Guarda a posição do quinto Control(referente ao colmovcaixa)
'    iPosColFim = InStr(iPosInicio, asReg(iIndice), Chr(vbKeySeparator))
'    If iPosColFim = 0 Then iPosColFim = iPosFim
'
'    iPosInicio = iPosInicio + 1
'
'    Do While iPosInicio < iPosColFim
'
'        'Procura o Primeiro Escape dentro da String
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        'Procura um control dentro da String
'        iPosColInicio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyControl)))
'        If iPosColInicio = 0 Or iPosColInicio > iPosColFim Then iPosColInicio = iPosColFim
'
'        iIndice2 = 0
'        Set objCheque = New ClassChequePre
'
'        Do While iPosMeio <> 0
'
'            iIndice2 = iIndice2 + 1
'            'Recolhe os Cheques
'            Select Case iIndice2
'
'                Case 1: objCheque.dtDataDeposito = StrParaDate(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 2: objCheque.dValor = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 3: objCheque.iAprovado = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 4: objCheque.iBanco = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 5: objCheque.iChequeSel = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 6: objCheque.iECF = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 7: objCheque.iFilial = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 8: objCheque.iFilialEmpresa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 9: objCheque.iFilialEmpresaLoja = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 10: objCheque.iNaoEspecificado = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 11: objCheque.lCliente = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 12: objCheque.lCupomFiscal = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 13: objCheque.lNumBordero = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 14: objCheque.lNumBorderoLoja = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 15: objCheque.lNumero = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 16: objCheque.lNumIntCheque = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 17: objCheque.lNumMovtoCaixa = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 18: objCheque.lSequencial = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 19: objCheque.lSequencialBack = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 20: objCheque.lSequencialLoja = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 21: objCheque.sAgencia = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 22: objCheque.sContaCorrente = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 23: objCheque.sCPFCGC = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 24: Exit Do
'
'            End Select
'
'            'Atualiza as Posições
'            iPosInicio = iPosMeio + 1
'            iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'            If iPosMeio > iPosColInicio Or iPosMeio = 0 Then iPosMeio = iPosColInicio
'
'        Loop
'
'        iPosInicio = iPosColInicio + 1
'
'        For Each objMovcx In colMovCaixa
'            If objCheque.lNumIntCheque = objMovcx.lNumRefInterna And objMovcx.iTipo = MOVIMENTOCAIXA_RECEB_CHEQUE Then objCheque.lNumMovtoCaixa = objMovcx.lNumMovto
'        Next
'
'        'Log de Inclusão
'        objLog.iOperacao = INCLUSAO_CHEQUE
'
'        'Chama a rotina que gera o sequencial
'        lErro = CF("Config_ObterNumInt", "LojaConfig", "COD_PROX_CHEQUE_LOJA", lNumIntDoc, 1, giFilialEmpresa)
'        If lErro <> SUCESSO Then gError 99227
'
'        objCheque.lSequencial = lNumIntDoc
'
'        lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_CHEQUE_PRE", lNumIntDoc, 1, giFilialEmpresa)
'        If lErro <> SUCESSO Then gError 99781
'
'        objCheque.lNumIntCheque = lNumIntDoc
'
'        'Função que Carrega o objLog para a Gravação no Banco de Dados
'        lErro = CF("Mover_Dados_Cheque_Log", objCheque, objLog)
'        If lErro <> SUCESSO Then gError 99927
'
'        'Função de Gravação de Log
'        lErro = CF("Log_Grava", objLog)
'        If lErro <> SUCESSO Then gError 99928
'
'        Call CF("Cheque_Insere", objCheque)
'
'    Loop
'
'    iPosInicio = iPosColFim + 2
'
'    Do While iPosInicio < iPosFim
'
'        'Procura o Primeiro Escape dentro da String
'        iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'        'Procura um control dentro da String
'        iPosColInicio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyControl)))
'        If iPosColInicio = 0 Then iPosColInicio = iPosFim
'
'        iIndice2 = 0
'        Set objTroca = New ClassTroca
'
'        Do While iPosMeio <> 0
'
'            iIndice2 = iIndice2 + 1
'            'Recolhe as trocas
'            Select Case iIndice2
'
'                Case 1: objTroca.dQuantidade = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 2: objTroca.dValor = StrParaDbl(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 3: objTroca.iFilialEmpresa = StrParaInt(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 4: objTroca.lNumIntDoc = StrParaLong(Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio))
'                Case 5: objTroca.sProduto = Mid(asReg(iIndice), iPosInicio, iPosMeio - iPosInicio)
'                Case 6: Exit Do
'
'            End Select
'
'            'Atualiza as Posições
'            iPosInicio = iPosMeio + 1
'            iPosMeio = (InStr(iPosInicio, asReg(iIndice), Chr(vbKeyEscape)))
'
'            If iPosMeio > iPosColInicio Or iPosMeio = 0 Then iPosMeio = iPosColInicio
'
'        Loop
'
'        iPosInicio = iPosColInicio + 1
'
'        For Each objMovcx In colMovCaixa
'            If objTroca.lNumIntDoc = objMovcx.lNumRefInterna And objMovcx.iTipo = MOVIMENTOCAIXA_RECEB_TROCA Then objTroca.lNumMovtoCaixa = objMovcx.lNumMovto
'        Next
'
'        'Função que Gera Inteiro Automático
'        lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_TROCA", lNumIntDoc, 1, giFilialEmpresa)
'        If lErro <> SUCESSO Then gError 99789
'
'        objTroca.lNumIntDoc = lNumIntDoc
'        Call CF("Troca_Grava", objTroca)
'
'    Loop
'
'    'De acordo com o tipo grava como orçamento ou como cupom fiscal
'    If iTipo = OPTION_CF Then
'        Call CF("CupomFiscal_Grava", objCupomFiscal)
'    Else
'        Call Transfere_Cupom(objCupomFiscal, objOrcamentoLoja)
'        Call CF("OrcamentoLoja_Grava", objOrcamentoLoja)
'    End If
'
'    objCarne.lCupomFiscal = objCupomFiscal.lNumIntDoc
'
'    lErro = CF("Carne_Grava", objCarne)
'    If lErro <> SUCESSO Then gError 99910
'
'    lErro = Transacao_Commit()
'    If lErro <> SUCESSO Then gError 99911
'
'    Exit Sub
'
'Erro_Le_Arquivo_Vendas:
'
'    Select Case gErr
'
'        Case 99780 To 99789, 99910, 99927, 99928
'
'        Case 99909
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 99911
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160717)
'
'    End Select
'
'    Call Transacao_Rollback
'
'    Exit Sub
'
'End Sub

'Private Sub Transfere_Cupom(objCupomFiscal As ClassCupomFiscal, objOrcamentoLoja As ClassOrcamentoLoja)
'
'Dim objItemOrcamento As New ClassItemOrcamentoLoja
'Dim objItemCupom As New ClassItemCupomFiscal
'
'    objOrcamentoLoja.dtDataEmissao = objCupomFiscal.dtDataEmissao
'    objOrcamentoLoja.dHoraEmissao = objCupomFiscal.dHoraEmissao
'    objOrcamentoLoja.dValorAcrescimo = objCupomFiscal.dValorAcrescimo
'    objOrcamentoLoja.dValorDesconto = objCupomFiscal.dValorDesconto
'    objOrcamentoLoja.dValorProdutos = objCupomFiscal.dValorProdutos
'    objOrcamentoLoja.dValorTotal = objCupomFiscal.dValorTotal
'    objOrcamentoLoja.dValorTroco = objCupomFiscal.dValorTroco
'    objOrcamentoLoja.iECF = objCupomFiscal.iECF
'    objOrcamentoLoja.iFilialEmpresa = objCupomFiscal.iFilialEmpresa
'    objOrcamentoLoja.iStatus = objCupomFiscal.iStatus
'    objOrcamentoLoja.iTabelaPreco = objCupomFiscal.iTabelaPreco
'''''    objOrcamentoLoja.lVolumeQuant = objCupomFiscal.lVolumeQuant
'    objOrcamentoLoja.lGerenteCancel = objCupomFiscal.lGerenteCancel
'    objOrcamentoLoja.lNumero = objCupomFiscal.lNumero
'    objOrcamentoLoja.lNumIntDoc = objCupomFiscal.lNumIntDoc
'    objOrcamentoLoja.lNumOrcamento = objCupomFiscal.lNumOrcamento
'    objOrcamentoLoja.iVendedor = objCupomFiscal.iVendedor
'    objOrcamentoLoja.sCPFCGC = objCupomFiscal.sCPFCGC
'    objOrcamentoLoja.sMotivoCancel = objCupomFiscal.sMotivoCancel
'    objOrcamentoLoja.sNaturezaOp = objCupomFiscal.sNaturezaOp
'''''    objOrcamentoLoja.sObservacao = objCupomFiscal.sObservacao
'''''    objOrcamentoLoja.lDuracao = objCupomFiscal.lDuracao
'
'    For Each objItemCupom In objCupomFiscal.colItens
'
'        objItemOrcamento.dAliquotaICMS = objItemCupom.dAliquotaICMS
'        objItemOrcamento.dPercDesc = objItemCupom.dPercDesc
'        objItemOrcamento.dPrecoUnitario = objItemCupom.dPrecoUnitario
'        objItemOrcamento.dQuantidade = objItemCupom.dQuantidade
'        objItemOrcamento.dValorDesconto = objItemCupom.dValorDesconto
'        objItemOrcamento.icancel = objItemCupom.icancel
'        objItemOrcamento.iStatus = objItemCupom.iStatus
'        objItemOrcamento.lNumIntOrcamento = objItemCupom.lNumIntCupom
'        objItemOrcamento.lNumIntDoc = objItemCupom.lNumIntDoc
'        objItemOrcamento.sProduto = objItemCupom.sProduto
'        objItemOrcamento.sUnidadeMed = objItemCupom.sUnidadeMed
'        objItemOrcamento.iItem = objItemCupom.iItem
'
'        objOrcamentoLoja.colItens.Add objItemOrcamento
'
'    Next
'
'End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Geração Arquivos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GeracaoArqCC"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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
