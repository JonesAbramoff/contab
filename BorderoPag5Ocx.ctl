VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BorderoPag5Ocx 
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   LockControls    =   -1  'True
   ScaleHeight     =   2685
   ScaleWidth      =   3555
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   622
      Picture         =   "BorderoPag5Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2025
      Width           =   1035
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   525
      Left            =   1897
      Picture         =   "BorderoPag5Ocx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2025
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1575
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Geração de Arquivo CNAB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3330
   End
   Begin VB.Label TotalTitulos 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2025
      TabIndex        =   6
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label TitulosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2025
      TabIndex        =   5
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Títulos Processados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Total de Títulos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   3
      Top             =   630
      Width           =   1455
   End
End
Attribute VB_Name = "BorderoPag5Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Falta revisao Jones

Option Explicit


''''
''''Dim giCancelaBatch As Integer
''''Dim giExecutando As Integer ' 0: nao está executando, 1: em andamento
''''
''''Dim gobjBorderoPagEmissao As ClassBorderoPagEmissao
''''
''''Function Trata_Parametros(objobjBorderoPagEmissao As ClassBorderoPagEmissao) As Long
''''
''''Dim lErro As Long
''''
''''On Error GoTo Erro_Trata_Parametros
''''
''''    giCancelaBatch = 0
''''    giExecutando = ESTADO_PARADO
''''
''''    If (objBorderoPagEmissao Is Nothing) Then
''''        Error #####
''''    Else
''''        Set gobjBorderoPagEmissao = objBorderoPagEmissao
''''    End If
''''
''''    Set gobjBorderoPagEmissao.objEvolucao = Me
''''
''''    'Passa para a tela os dados dos Títulos selecionados
''''    TotalTitulos.Caption = CStr(gobjBorderoPagEmissao.iQtdeParcelasSelecionadas)
''''    TitulosProcessados.Caption = "0"
''''
''''    ProgressBar1.Min = 0
''''    ProgressBar1.Max = 100
''''
''''    Trata_Parametros = SUCESSO
''''
''''    Exit Function
''''
''''Erro_Trata_Parametros:
''''
''''    Trata_Parametros = Err
''''
''''    Select Case Err
''''
''''        Case #####
''''            giCancelaBatch = CANCELA_BATCH
''''
''''        Case Else
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143819)
''''
''''    End Select
''''
''''    Exit Function
''''
''''End Function
''''
''''Private Sub BotaoCancela_Click()
''''
''''    If giExecutando = ESTADO_ANDAMENTO Then
''''        giCancelaBatch = CANCELA_BATCH
''''        BotaoCancela.Enabled = False
''''        Exit Sub
''''    End If
''''
''''    'Fecha a tela
''''    Unload Me
''''
''''End Sub
''''
''''Private Sub BotaoOK_Click()
''''
''''Dim lErro As Long
''''
''''On Error GoTo Erro_BotaoOk_Click
''''
''''    BotaoOK.Enabled = False
''''
''''    BotaoCancela.Enabled = True
''''
''''    If giCancelaBatch <> CANCELA_BATCH Then
''''
''''        giExecutando = ESTADO_ANDAMENTO
''''        lErro = ... AtualizarBD(gobjBorderoPagEmissao)
''''        giExecutando = ESTADO_PARADO
''''
''''        BotaoCancela.Enabled = False
''''
''''        If lErro <> SUCESSO And lErro <> ##### Then Error #####
''''
''''        If lErro = ##### Then Error ##### 'interrompeu
''''
''''    End If
''''
''''    Exit Sub
''''
''''Erro_BotaoOk_Click:
''''
''''    Select Case Err
''''
''''        Case #####
''''            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
''''            Unload Me
''''
''''        Case #####
''''
''''        Case Else
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143820)
''''
''''    End Select
''''
''''    Exit Sub
''''
''''End Sub
''''
''''Public Function Mostra_Evolucao(iCancela As Integer, iNumProc As Integer) As Long
''''
''''Dim lErro As Long
''''Dim iEventos As Integer
''''Dim iProcessados As Integer
''''Dim iTotal As Integer
''''
''''On Error GoTo Erro_Mostra_Evolucao
''''
''''    iEventos = DoEvents()
''''
''''    If giCancelaBatch = CANCELA_BATCH Then
''''
''''        iCancela = CANCELA_BATCH
''''        giExecutando = ESTADO_PARADO
''''
''''    Else
''''        'atualiza dados da tela ( registros atualizados e a barra )
''''        iProcessados = CInt(TitulosProcessados.Caption)
''''        iTotal = CInt(TotalTitulos.Caption)
''''
''''        iProcessados = iProcessados + iNumProc
''''        TitulosProcessados.Caption = CStr(iProcessados)
''''
''''        ProgressBar1.Value = CInt((iProcessados / iTotal) * 100)
''''
''''        giExecutando = ESTADO_ANDAMENTO
''''
''''    End If
''''
''''    Mostra_Evolucao = SUCESSO
''''
''''    Exit Function
''''
''''Erro_Mostra_Evolucao:
''''
''''    Mostra_Evolucao = Err
''''
''''    Select Case Err
''''
''''        Case Else
''''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143821)
''''
''''    End Select
''''
''''    giCancelaBatch = CANCELA_BATCH
''''
''''    Exit Function
''''
''''End Function
''''
''''Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
''''
''''    If giExecutando = ESTADO_ANDAMENTO Then
''''        If giCancelaBatch <> CANCELA_BATCH Then giCancelaBatch = CANCELA_BATCH
''''        Cancel = 1
''''    End If
''''
''''End Sub
''''
''''Public Sub Form_Unload(Cancel As Integer)
''''
''''    Set gobjBorderoPagEmissao = Nothing
''''
''''End Sub


''Private Declare Function CNAB_PagRem_Abrir Lib "ADCNAB.DLL" (lCNABPagRem As Long, ByVal sNomeArq As String, ByVal iCodigoBanco As Integer, ByVal iNumRemessa As Integer, vDataEmissao As Variant, ByVal iTipoCobranca As Integer, ByVal iLiqTitOutrosBcos As Integer) As Long
''Private Declare Function CNAB_PagRem_Fechar Lib "ADCNAB.DLL" (ByVal lCNABPagRem As Long) As Long
''Private Declare Function CNAB_PagRem_DefCtaEmpresa Lib "ADCNAB.DLL" (ByVal lCNABPagRem As Long, ByVal sAgencia As String, ByVal sConta As String, ByVal sDVConta As String) As Long
''Private Declare Function CNAB_PagRem_IncluirReg Lib "ADCNAB.DLL" (ByVal lCNABPagRem As Long, ByVal dValorBaixado As Double, ByVal sSiglaDocumento As String, vDataVencimento As Variant, vDataEmissao As Variant, ByVal sPagtoId As String, ByVal lNumTitulo As Long, ByVal sNossoNumero As String, _
''            ByVal sEndereco As String, ByVal sCidade As String, ByVal sSiglaEstado As String, ByVal sCEP As String, ByVal sCGC As String, ByVal sRazaoSocial As String, ByVal iBanco As Integer, ByVal sAgencia As String, ByVal sContaCorrente As String) As Integer
''
''Dim gobjBorderoPagEmissao As ClassBorderoPagEmissao
''
''Dim ERRO_LEITURA_PAGTO_BORDERO 'Erro na leitura de pagamento efetuado por borderô
''
''Function Trata_Parametros(objBorderoPagEmissao As ClassBorderoPagEmissao) As Long
'''Traz os dados das Parcelas a pagar para a Tela
''
''    Set gobjBorderoPagEmissao = objBorderoPagEmissao
''
''    Exit Function
''
''End Function
''
'''@@@@ TRANSFERIR P/CPRSELECT
''
''Function BorderoPag_CriarArqCNAB1(objBorderoPagto As ClassBorderoPagto, lCNABPagRem As Long) As Long
'''passar dados gerais da empresa e da cta corrente
''Dim lErro As Long
''Dim objContaCorrenteInt As New ClassContasCorrentesInternas
''On Error GoTo Erro_BorderoPag_CriarArqCNAB1
''
''    'passar dados da cta corrente da empresa
''    lErro = CF("ContaCorrenteInt_Le",objBorderoPagto.iCodConta, objContaCorrenteInt)
''    If lErro <> SUCESSO Then Error 7745
''
''    'abre o arquivo CNAB
''    lErro = CNAB_PagRem_Abrir(lCNABPagRem, objBorderoPagto.sNomeArq, objContaCorrenteInt.iCodBanco, objBorderoPagto.iNumArqRemessa, objBorderoPagto.dtDataEmissao, objBorderoPagto.iTipoDeCobranca, objBorderoPagto.iTitOutroBanco)
''    If lErro <> SUCESSO Then Error 7746
''
''    'passar dados da cta corrente pagadora
''    lErro = CNAB_PagRem_DefCtaEmpresa(lCNABPagRem, objContaCorrenteInt.sAgencia, objContaCorrenteInt.sNumConta, objContaCorrenteInt.sDVAgConta)
''    If lErro <> SUCESSO Then Error 7747
''
''    BorderoPag_CriarArqCNAB1 = SUCESSO
''
''    Exit Function
''
''Erro_BorderoPag_CriarArqCNAB1:
''
''    BorderoPag_CriarArqCNAB1 = Err
''
''    Select Case Err
''
''        Case 7745, 7746, 7747
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143822)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function BorderoPag_CriarArqCNAB(objBorderoPagto As ClassBorderoPagto) As Long
'''cria o arquivo de remessa para pagtos atraves de banco
''Dim lErro As Long, lCNABPagRem As Long
''On Error GoTo Erro_BorderoPag_CriarArqCNAB
''
''    'abrir o arquivo CNAB
''    lErro = BorderoPag_CriarArqCNAB1(objBorderoPagto, lCNABPagRem)
''    If lErro <> SUCESSO Then Error 7748
''
''    'incluir dados dos pagtos de titulos nao baixados
''    lErro = BorderoPag_CriarArqCNAB2(objBorderoPagto, lCNABPagRem)
''    If lErro <> SUCESSO Then Error 7749
''
'''''    'incluir dados dos pagtos de titulos baixados
'''''    lErro = BorderoPag_CriarArqCNAB3(objBorderoPagto, lCNABPagRem)
'''''    If lErro <> SUCESSO Then Error 7750
'''''
'''''    'fecha o arquivo CNAB
'''''    lErro = BorderoPag_CriarArqCNAB4(objBorderoPagto, lCNABPagRem)
'''''    If lErro <> SUCESSO Then Error 7751
'''''
'''''    'atualizar o nome do arquivo associado ao bordero
'''''    lErro = CF("BorderoPagto_Gravar",objBorderoPagto)
'''''    If lErro <> SUCESSO Then Error 7752
''
''    BorderoPag_CriarArqCNAB = SUCESSO
''
''    Exit Function
''
''Erro_BorderoPag_CriarArqCNAB:
''
''    BorderoPag_CriarArqCNAB = Err
''
''    Select Case Err
''
''        Case 7748 To 7752
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143823)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function BorderoPag_CriarArqCNAB2(objBorderoPagto As ClassBorderoPagto, lCNABPagRem As Long) As Long
'''inclui registros correspodentes a titulos nao baixados
''Dim lErro As Long, lComando As Long
''Dim iSeqBaixaParc As Integer, dValorBaixado As Double, sSiglaDocumento As String, dtDataEmissao As Date, lNumTitulo As Long, dtDataVencimento As Date, lNumIntParcela As Long, sNossoNumero As String, sRazaoSocial As String
''Dim sCGC As String, iBanco As Integer, sAgencia As String, sContaCorrente As String, sEndereco As String, sCidade As String, sSiglaEstado As String, sCEP As String
''Dim sPagtoId As String
''On Error GoTo Erro_BorderoPag_CriarArqCNAB2
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 7744
''
''    'inicializa os buffers com zeros
''    Call BorderoPag_CriarArqCNAB5(sSiglaDocumento, sNossoNumero, sRazaoSocial, sCGC, sAgencia, sContaCorrente, sEndereco, sCidade, sSiglaEstado, sCEP)
''
''    lErro = Comando_Executar(lComando, _
''        "SELECT BaixasParcPag.Sequencial, BaixasParcPag.ValorBaixado, " & _
''        "TitulosPag.SiglaDocumento, TitulosPag.DataEmissao, TitulosPag.NumTitulo, ParcelasPag.DataVencimento, ParcelasPag.NumIntDoc, ParcelasPag.NossoNumero, Fornecedores.RazaoSocial, FiliaisFornecedores.Endereco, FiliaisFornecedores.CGC, FiliaisFornecedores.Banco, FiliaisFornecedores.Agencia, FiliaisFornecedores.ContaCorrente, Enderecos.Endereco, Enderecos.Cidade, Enderecos.SiglaEstado, Enderecos.CEP FROM BorderosPagto, MovimentosContaCorrente, BaixasPag, BaixasParcPag, ParcelasPag, TitulosPag, Fornecedores, FiliaisFornecedores, Enderecos WHERE MovimentosContaCorrente.Tipo = ? AND MovimentosContaCorrente.NumRefInterna = ? AND MovimentosContaCorrente.NumMovto = BaixasPag.NumMovCta AND BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo = TitulosPag.NumIntDoc AND TitulosPag.Fornecedor = Fornecedores.Codigo AND TitulosPag.Fornecedor = FiliaisFornecedores.CodFornecedor " & _
''        "AND TitulosPag.Filial = FiliaisFornecedores.CodFilial AND FiliaisFornecedores.Endereco = Enderecos.Codigo", _
''        iSeqBaixaParc, dValorBaixado, sSiglaDocumento, dtDataEmissao, lNumTitulo, dtDataVencimento, lNumIntParcela, sNossoNumero, sRazaoSocial, _
''        sEndereco, sCGC, iBanco, sAgencia, sContaCorrente, sEndereco, sCidade, sSiglaEstado, sCEP, MOVCCI_PAGTO_TITULO_POR_BORDERO, objBorderoPagto.lNumIntBordero)
''    If lErro <> AD_SQL_SUCESSO Then Error 7753
''
''    lErro = Comando_BuscarProximo(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7754
''
''    Do While lErro <> AD_SQL_SEM_DADOS
''
''        sPagtoId = CStr(lNumIntParcela) & "-" & CStr(iSeqBaixaParc)
''
''        'incluir registro no arquivo
''        lErro = CNAB_PagRem_IncluirReg(lCNABPagRem, dValorBaixado, sSiglaDocumento, dtDataVencimento, dtDataEmissao, sPagtoId, lNumTitulo, sNossoNumero, _
''            sEndereco, sCidade, sSiglaEstado, sCEP, sCGC, sRazaoSocial, iBanco, sAgencia, sContaCorrente)
''        If lErro <> SUCESSO Then Error 7756
''
''        lErro = Comando_BuscarProximo(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7755
''
''    Loop
''
''    lErro = Comando_Fechar(lComando)
''
''    BorderoPag_CriarArqCNAB2 = SUCESSO
''
''    Exit Function
''
''Erro_BorderoPag_CriarArqCNAB2:
''
''    BorderoPag_CriarArqCNAB2 = Err
''
''    Select Case Err
''
''        Case 7756
''
''        Case 7753, 7754, 7755
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PAGTO_BORDERO", Err)
''
''        Case 7744
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143824)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
''Sub BorderoPag_CriarArqCNAB5(sSiglaDocumento As String, sNossoNumero As String, sRazaoSocial As String, sCGC As String, sAgencia As String, sContaCorrente As String, sEndereco As String, sCidade As String, sSiglaEstado As String, sCEP As String)
''
''    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
''    sNossoNumero = String(STRING_NOSSO_NUMERO, 0)
''    sRazaoSocial = String(STRING_FORNECEDOR_RAZAO_SOC, 0)
''    sCGC = String(STRING_CGC, 0)
''    sAgencia = String(STRING_AGENCIA, 0)
''    sContaCorrente = String(STRING_CONTA_CORRENTE, 0)
''    sEndereco = String(STRING_ENDERECO, 0)
''    sCidade = String(STRING_CIDADE, 0)
''    sSiglaEstado = String(STRING_ESTADO, 0)
''    sCEP = String(STRING_CEP, 0)
''
''End Sub
''
''Public Sub Form_Load()
''Dim lErro As Long, objBorderoPagto As New ClassBorderoPagto
''On Error GoTo Erro_Form_Load
''
''    objBorderoPagto.dtDataEmissao
''    objBorderoPagto.dtDataEnvio
''    objBorderoPagto.iCodConta
''    objBorderoPagto.iExcluido
''    objBorderoPagto.iNumArqRemessa
''    objBorderoPagto.iTipoDeCobranca
''    objBorderoPagto.iTitOutroBanco
''    objBorderoPagto.lNumero
''    objBorderoPagto.lNumIntBordero
''    objBorderoPagto.sNomeArq
''
''    lErro = CF("BorderoPag_CriarArqCNAB",objBorderoPagto)
''    If lErro <> SUCESSO Then Error 7782
''
''    Exit Sub
''
''Erro_Form_Load:
''
''    lErro_Chama_Tela = Err
''
''    Select Case Err
''
''        Case 7782
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143825)
''
''    End Select
''
''    Exit Sub
''
''
''End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub TitulosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TitulosProcessados, Source, X, Y)
End Sub

Private Sub TitulosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TitulosProcessados, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub


Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property
