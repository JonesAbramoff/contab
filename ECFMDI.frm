VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm ECF 
   BackColor       =   &H8000000C&
   Caption         =   "ECF"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   17760
   Icon            =   "ECFMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   1470
      Top             =   1290
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   735
      Top             =   1215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECFMDI.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECFMDI.frx":13AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECFMDI.frx":28CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECFMDI.frx":33AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECFMDI.frx":4572
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ECFMDI.frx":5FFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17760
      _ExtentX        =   31327
      _ExtentY        =   1799
      ButtonWidth     =   2381
      ButtonHeight    =   1640
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Venda M"
            Object.ToolTipText     =   "Venda de muitos produtos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Venda P"
            Object.ToolTipText     =   "Venda de poucos produtos (Teclado Configurado)"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pesquisar Preço"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Carnê"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Suspender"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMenuFiscal 
      Caption         =   "&Menu Fiscal"
      Begin VB.Menu mnuMF 
         Caption         =   "LX"
         Index           =   1
      End
      Begin VB.Menu mnuMF 
         Caption         =   "LMF"
         Index           =   2
         Begin VB.Menu mnuMFLMF 
            Caption         =   "completa"
            Index           =   1
            Begin VB.Menu mnuMFLMFC 
               Caption         =   "por intervalo de data"
               Index           =   1
            End
            Begin VB.Menu mnuMFLMFC 
               Caption         =   "por intervalo de redução Z"
               Index           =   2
            End
         End
         Begin VB.Menu mnuMFLMF 
            Caption         =   "simplificada"
            Index           =   3
            Begin VB.Menu mnuMFLMFS 
               Caption         =   "por intervalo de data"
               Index           =   1
            End
            Begin VB.Menu mnuMFLMFS 
               Caption         =   "por intervalo de redução Z"
               Index           =   2
            End
         End
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Arq. MF"
         Index           =   4
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Arq. MFD"
         Index           =   5
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Identificação do PAF-ECF"
         Index           =   11
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Vendas do Periodo"
         Index           =   12
         Begin VB.Menu mnuMFVP 
            Caption         =   "SINTEGRA"
            Index           =   1
         End
         Begin VB.Menu mnuMFVP 
            Caption         =   "SPED"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Tab. Índice Técnico Produção"
         Index           =   13
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Parametros de Configuração"
         Index           =   14
      End
      Begin VB.Menu mnuMF 
         Caption         =   "Registros do PAF-ECF"
         Index           =   15
         Begin VB.Menu mnuMFEstoque 
            Caption         =   "Estoque Total"
            Index           =   1
         End
         Begin VB.Menu mnuMFEstoque 
            Caption         =   "Estoque Parcial"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuVenda 
      Caption         =   "&Venda"
      Begin VB.Menu mnuECFVenda 
         Caption         =   "Venda &M"
         Index           =   3
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuECFVenda 
         Caption         =   "Venda &P (Teclado)"
         Index           =   4
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuECFVenda 
         Caption         =   "Nota Fiscal Mod.d2"
         Index           =   16
      End
      Begin VB.Menu mnuECFVenda 
         Caption         =   "NFe - Nota Fiscal Eletronica"
         Index           =   17
      End
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Index           =   1
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "&Abertura / Fechamento"
         Index           =   1
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Sangria/Suprimento &Dinheiro"
         Index           =   3
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Sangria &Cheque  "
         Index           =   4
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Sangria &Boleto"
         Index           =   5
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Sangria &Ticket"
         Index           =   6
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Sangria &Outros Meios de Pagamento"
         Index           =   7
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "T&ransferência de Caixa"
         Index           =   9
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Cancelamento"
         Index           =   10
      End
      Begin VB.Menu mnuECFMovimentos 
         Caption         =   "Reimpressão"
         Index           =   11
      End
   End
   Begin VB.Menu mnuFuncional 
      Caption         =   "&Funções"
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "Consultar &Preço"
         Index           =   8
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "Operação &Arquivo"
         Index           =   10
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "Imprimir &Carnê"
         Index           =   11
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "Atualizar Ta&belas"
         Index           =   12
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "&Pré-Autorização"
         Index           =   13
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "&Fininvest"
         Index           =   14
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "Consu&ltar Produtos"
         Index           =   15
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mnuECFFuncional 
         Caption         =   "NFCe - Enviar Xml Pendente"
         Index           =   18
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configurações"
      Begin VB.Menu mnuECFConfig 
         Caption         =   "C&onfigurações Gerais"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuECFConfigSAT 
         Caption         =   "Configurações SAT"
      End
      Begin VB.Menu mnuECFConfigNFe 
         Caption         =   "Configurações NFe/NFCe"
      End
      Begin VB.Menu mnuECFConfigBkp 
         Caption         =   "Backup"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
      Begin VB.Menu mnuECFSair 
         Caption         =   "Sai&r"
         Shortcut        =   +{F6}
      End
   End
End
Attribute VB_Name = "ECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constantes de Indices para o Menu:
Const MENU_ECF_MOV_ABERTURA_FECHAMENTO = 1
Const MENU_ECF_MOV_MOVIMENTO_DINHEIRO = 3
Const MENU_ECF_MOV_MOVIMENTO_CHEQUE = 4
Const MENU_ECF_MOV_MOVIMENTO_BOLETO = 5
Const MENU_ECF_MOV_MOVIMENTO_TICKET = 6
Const MENU_ECF_MOV_MOVIMENTO_OUTROS = 7
Const MENU_ECF_MOV_TRANSFCAIXA = 9

Const MENU_ECF_MOV_VENDAM = 3
Const MENU_ECF_MOV_VENDAP = 4
Const MENU_ECF_MOV_PRECO = 5
Const MENU_ECF_MOV_CONSULTAR_PRECO = 8
Const MENU_ECF_MOV_OPERACAOARQ = 10
Const MENU_ECF_MOV_IMPRIMIR_CARNE = 11
Const MENU_ECF_MOV_ATUALIZAR_DADOS = 12
Const MENU_ECF_MOV_PRE_AUTORIZACAO = 13
Const MENU_ECF_MOV_FININVEST = 14
Const MENU_ECF_MOV_CONSULTAR_PRODUTOS = 15
Const MENU_ECF_MOV_NFD2 = 16
Const MENU_ECF_MOV_NFE = 17
Const MENU_ECF_MOV_NFCE_XML_OFFLINE = 18
Const MENU_ECF_MOV_CANCELAR_VENDA = 20

Const MENU_MF_LX = 1
Const MENU_MF_ARQMF = 4
Const MENU_MF_ARQMFD = 5
Const MENU_MF_TABPROD = 6
Const MENU_MF_ESTPROD = 7
Const MENU_MF_MVPORECF = 8
Const MENU_MF_MEIOPAGTO = 9
Const MENU_MF_IDENTPAFECF = 11
Const MENU_MF_INDICETECNICO = 13
Const MENU_MF_PARAMCONFIG = 14

Const MENU_MFLMFC_PORDATA = 1
Const MENU_MFLMFC_PORREDZ = 2
Const MENU_MFLMFS_PORDATA = 1
Const MENU_MFLMFS_PORREDZ = 2

Const MENU_MFESPELHO_PORDATA = 1
Const MENU_MFESPELHO_PORCOO = 2

Const MENU_MFARQ_PORDATA = 1
Const MENU_MFARQ_PORCOO = 2


Const MENU_MFMOVECF_PORDATA = 1
Const MENU_MFMOVECF_PORECF = 2

Const MENU_MFDAV_RELGERENCIAL = 1
Const MENU_MFDAV_ARQUIVO = 2

Const MENU_MFVP_SINTEGRA = 1
Const MENU_MFVP_SPED = 2


Const MENU_MF_REGPAFECF_TOTAL = 1
Const MENU_MF_REGPAFECF_PARCIAL = 2

Dim iCountBkp As Integer

Private Sub MDIForm_Load()
    
Dim i As Integer
Dim lErro As Long
Dim objAux As Object
Dim sPerfil As String
Dim sRetorno As String
Dim lTamanho As Long
Dim iTelaVendaMaximizada As Integer

On Error GoTo Erro_MDIForm_Load

    lTamanho = 150
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "TelaVendaMaximizada", "1", sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    iTelaVendaMaximizada = StrParaInt(sRetorno)
    
    Me.left = 0
    Me.top = 0
    
    If iTelaVendaMaximizada <> 0 Then Me.WindowState = 2
    
    Set GL_objMDIForm = Me
    
    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Image = i
    Next
    
    Set gobjCheckboxChecked = LoadPicture("checkboxchecked.bmp")
    Set gobjCheckboxUnchecked = LoadPicture("checkboxunchecked.bmp")
    Set gobjOptionButtonChecked = LoadPicture("optionbuttonchecked.bmp")
    Set gobjOptionButtonUnChecked = LoadPicture("optionbuttonunchecked.bmp")
    Set gobjButton = LoadPicture("botao.bmp")
            
    'Incluído por Cyntia
'    gdtDataHoje = Date

    Set GL_objKeepAlive = New AdmKeepAlive
       
    Set objAux = CreateObject("AdmLib.Adm")
    If objAux Is Nothing Then gError 109538
    GL_objKeepAlive.Add objAux
    
'    lErro = Inicializa_Rotinas()
'    If lErro <> SUCESSO Then gError 99898
        
'    Set objAux = gcolRotinasECF
'    GL_objKeepAlive.Add objAux
    
    Set objAux = CreateObject("GlobaisAdm.AdmAdm")
    If objAux Is Nothing Then gError 109537
    GL_objKeepAlive.Add objAux
    
    Set objAux = CreateObject("RotinasECF.ClassECFSelect")
    If objAux Is Nothing Then gError 99907
    GL_objKeepAlive.Add objAux
    
    Set objAux = CreateObject("RotinasECF.ClassECFGrava")
    If objAux Is Nothing Then gError 99908
    GL_objKeepAlive.Add objAux
    
    Set objAux = CreateObject("TelasAdm.ClassTelasAdm")
    If objAux Is Nothing Then Error 109539
    GL_objKeepAlive.Add objAux
        
    GL_objKeepAlive.Add gobjSATInfo
    GL_objKeepAlive.Add gobjNFeInfo
    
    Set objAux = Me
        
    Call CF_ECF("Customiza_Menu_Principal", objAux)
        
    'Call Afrac_UF_ObtemPerfil(sPerfil)
    'If sPerfil = "W" Or sPerfil = "X" Then mnuECFVenda.Item(16).Visible = False
    mnuECFVenda.Item(16).Visible = False
    
    If AFRAC_ImpressoraCFe(giCodModeloECF) Then mnuMenuFiscal.Visible = False
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_MDIForm_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 99907, 99908, 99898, 109537, 109538, 109539
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159193)

    End Select
    
    Exit Sub
        
End Sub

'Function Inicializa_Rotinas() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Inicializa_Rotinas
'
'    lErro = Leitura_Arquivo_ECF
'    If lErro <> SUCESSO Then gError 99899
'
'    Exit Function
'
'Erro_Inicializa_Rotinas:
'
'Inicializa_Rotinas = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159194)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function Leitura_Arquivo_ECF() As Long
'
'Dim sNomeRotina As String
'Dim sLocal As String
'Dim sRegistro As String
'Dim iPos As Integer
'
'On Error GoTo Erro_Leitura_Arquivo_ECF
'
'    Set gcolRotinasECF = New Collection
'
'    'Abre o arquivo de retorno
'    Open NOME_ARQUIVO_ROTINAS For Input As #1
'
'    'Até chegar ao fim do arquivo
'    Do While Not EOF(1)
'
'        'Busca o próximo registro do arquivo
'        Line Input #1, sRegistro
'
'        'Procura o sinal para separar o nome da rotina do seu local
'        iPos = InStr(1, sRegistro, SINAL_IGUAL)
'
'        sNomeRotina = Left(sRegistro, iPos - 1)
'        sLocal = Mid(sRegistro, iPos + 1, Len(sRegistro) - (iPos - 1))
'
'        gcolRotinasECF.Add sLocal, sNomeRotina
'
'    Loop
'
'    Close #1
'
'    Exit Function
'
'Erro_Leitura_Arquivo_ECF:
'
'Leitura_Arquivo_ECF = gErr
'
'    Close 1
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159195)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If (Sistema_QueryUnload() = False) Then Cancel = True

    If Not Cancel Then Call ECF_Grava_Log("Fechamento do Sistema.")

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_MDIForm_Unload

''    If gl_ST_ComandoSeta <> 0 Then
''
''        lErro = Comando_Fechar(gl_ST_ComandoSeta)
''        gl_ST_ComandoSeta = 0
''
''    End If
''
''    Set objAdmSeta = Nothing
''
''    'Edicao Tela
''    If mnuEdicao.Checked = True Then Call mnuEdicao_Click
''
''    If (Not (gobjEstInicial Is Nothing)) Then
''
''        Unload gobjEstInicial
''        Set gobjEstInicial = Nothing
''
''    End If
    
    'Call Usuario_Altera_SituacaoLogin(gsUsuario, USUARIO_NAO_LOGADO)

    'Se a Sessão Estiver Fechada então gera Erro
    If giStatusSessao = SESSAO_ABERTA Then
        
        'Função que Executa a Suspenção da Sessão
        lErro = CF_ECF("Sessao_Executa_Suspensao")
        If lErro <> SUCESSO Then gError 133528

    End If

    'Função que fecha as conexoes de bd
    lErro = CF_ECF("FechaBDs_PAFECF")
    If lErro <> SUCESSO Then gError 210420
    
    lErro = Sistema_Fechar_ECF()
    
    
    Exit Sub
    
Erro_MDIForm_Unload:

    Select Case gErr
    
        Case 133528, 210420
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159196)

    End Select
    
    Exit Sub

End Sub

Private Sub mnuECFConfig_Click()

On Error GoTo Erro_mnuECFConfig_Click

    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105884
    
    Call Chama_TelaECF("CaixaECFConfig")
    
    Exit Sub
    
Erro_mnuECFConfig_Click:

    Select Case gErr
    
        Case 105884
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159197)

    End Select
    
    Exit Sub
    
End Sub

Private Sub mnuECFConfigBkp_Click()

'Dim frmBackupConfig As New BackupConfigECF

On Error GoTo Erro_mnuECFConfig_Click

    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105884
    
'    Load frmBackupConfig
'
'    Call frmBackupConfig.Trata_Parametros
'
'    frmBackupConfig.Show

    Call Chama_TelaECF("BackupConfigECF")
    
    Exit Sub
    
Erro_mnuECFConfig_Click:

    Select Case gErr
    
        Case 105884
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159197)

    End Select
    
    Exit Sub
    
End Sub

Private Sub mnuECFConfigSAT_Click()

On Error GoTo Erro_mnuECFConfig_Click

    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105884
    
    Call Chama_TelaECF("SATConfig")
    
    Exit Sub
    
Erro_mnuECFConfig_Click:

    Select Case gErr
    
        Case 105884
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159197)

    End Select
    
    Exit Sub

End Sub

Private Sub mnuECFConfigNFe_Click()

On Error GoTo Erro_mnuECFConfig_Click

    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105884
    
    Call Chama_TelaECF("NFeConfig")
    
    Exit Sub
    
Erro_mnuECFConfig_Click:

    Select Case gErr
    
        Case 105884
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159197)

    End Select
    
    Exit Sub

End Sub

Private Sub mnuECFFuncional_Click(Index As Integer)
    
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim objEventoProduto As New AdmEvento
    
On Error GoTo Erro_mnuECFFuncional_Click
    
    Select Case Index
        
        Case MENU_ECF_MOV_CONSULTAR_PRECO
            Call Chama_TelaECF("Preco")
                
        Case MENU_ECF_MOV_OPERACAOARQ
            If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105882
            Call Chama_TelaECF("OperacaoArq")
            
        Case MENU_ECF_MOV_IMPRIMIR_CARNE
            If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105883
            Call Chama_TelaECF("ImpressaoCarne")
    
        Case MENU_ECF_MOV_ATUALIZAR_DADOS
            
            Me.MousePointer = vbHourglass
            
            lErro = CF_ECF("Carrega_Arquivo_FonteDados", 1)
            If lErro <> SUCESSO Then gError 133727
        
            Me.MousePointer = vbDefault
        
            Call Rotina_AvisoECF(vbOKOnly, AVISO_TABELAS_ATUALIZADAS)
    
        Case MENU_ECF_MOV_PRE_AUTORIZACAO
            Call Chama_TelaECF("PreAutorizacao")
            
        Case MENU_ECF_MOV_FININVEST
            Call Chama_TelaECF("Fininvest")
            
            
        Case MENU_ECF_MOV_CONSULTAR_PRODUTOS
            'Chama tela de ProdutosLista
            Call Chama_TelaECF_Modal("ProdutosLista", colSelecao, objProduto, objEventoProduto)
            
        Case MENU_ECF_MOV_NFCE_XML_OFFLINE
            Call Chama_TelaECF("NFCeOffline")
            
    End Select
    
    Exit Sub
    
Erro_mnuECFFuncional_Click:

    Select Case gErr
    
        Case 105882, 105883
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case 133727
            Me.MousePointer = vbDefault

        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159198)

    End Select
    
    Exit Sub

End Sub

Private Sub mnuECFMovimentos_Click(Index As Integer)
    
    
On Error GoTo Erro_mnuECFMovimentos_Click

    If Index = MENU_ECF_MOV_ABERTURA_FECHAMENTO Then
        Call Chama_TelaECF("OperacoesECF")
    Else

        If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105881

        'Se a Sessão Estiver Fechada então gera Error
        If giStatusSessao = SESSAO_ENCERRADA Then gError 105698
    
        'Se Sessão estiver Suspensa
        If giStatusSessao = SESSAO_SUSPENSA Then gError 105699
    
        Select Case Index
            
            Case MENU_ECF_MOV_MOVIMENTO_DINHEIRO
                Call Chama_TelaECF("MovimentoDinheiro")
                
            Case MENU_ECF_MOV_MOVIMENTO_CHEQUE
                Call Chama_TelaECF("MovimentoCheque")
                
            Case MENU_ECF_MOV_MOVIMENTO_BOLETO
                Call Chama_TelaECF("MovimentoBoleto")
            
            Case MENU_ECF_MOV_MOVIMENTO_TICKET
                Call Chama_TelaECF("MovimentoTicket")
                
            Case MENU_ECF_MOV_MOVIMENTO_OUTROS
                Call Chama_TelaECF("MovimentoOutros")
            
            Case MENU_ECF_MOV_TRANSFCAIXA
                Call Chama_TelaECF("TransfCaixa")
            
            Case 10
                Call Chama_TelaECF("CancelaCupom")
            
            Case 11
                Call Chama_TelaECF("ReimprimeCupom")
            
        End Select

    End If

    Exit Sub
    
Erro_mnuECFMovimentos_Click:

    Select Case gErr
    
        Case 105698
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_ABERTA_INEXISTENTE, gErr, giCodCaixa)

        Case 105699
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_SUSPENSA, gErr, giCodCaixa)
        
        Case 105881
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159199)

    End Select
    
    Exit Sub

End Sub

Private Sub mnuECFSair_Click()
Dim lErro As Long

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then lErro = AFRAC_FechaPortaSerial()
    
    End

End Sub

Private Sub mnuECFVenda_Click(Index As Integer)

Dim lErro As Long, iRZPendente As Integer, bECFComProblema As Boolean
Dim iAux As Integer

On Error GoTo Erro_mnuECFVenda_Click

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Verifica se já foi executa a redução z para a data de hoje
        If gdtUltimaReducao = Date Then gError 111324

    End If

    Select Case Index
    
        Case MENU_ECF_MOV_VENDAM
            Call Chama_TelaECF_Modal("VendaM")

        Case MENU_ECF_MOV_VENDAP
            Call Chama_TelaECF_Modal("VendaP")

        Case MENU_ECF_MOV_NFD2
            lErro = AFRAC_ECFComProblema(bECFComProblema)
            If lErro <> SUCESSO Then
                lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "ECF com problema")
                If lErro <> SUCESSO Then gError 214009
            End If
            
            If Not bECFComProblema Then
                lErro = AFRAC_RZPendente(iRZPendente)
                If lErro <> SUCESSO Then
                    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Reducao Z Pendente")
                    If lErro <> SUCESSO Then gError 214009
                End If
            End If
            
            If iRZPendente = 0 Or bECFComProblema Then
        
                Call Chama_TelaECF_Modal("NFD2")
                
            Else
                gError 214336
            End If

        Case MENU_ECF_MOV_NFE
            giCodModeloECF = IMPRESSORA_NFE
            Call Chama_TelaECF_Modal("VendaM")
            giCodModeloECF = giCodModeloECFConfig
            
    End Select

    Exit Sub
    
Erro_mnuECFVenda_Click:

    Select Case gErr

        Case 111324
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
    
        Case 214336
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFD2_DESABILITADA, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159200)

    End Select

    Exit Sub

End Sub

Private Sub mnuMF_Click(Index As Integer)

Dim objTela As Object
Dim lErro As Long

On Error GoTo Erro_mnuMF_Click
    
    If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
    '    lErro = CF_ECF("Requisito_XXII")
    '    If lErro <> SUCESSO Then gError 210887
    
        Select Case Index
        
            Case MENU_MF_LX
                Call CF_ECF("Executa_LeituraX")
    
            Case MENU_MF_ARQMF
                Call CF_ECF("ArqMF_Executa")
    
            Case MENU_MF_ARQMFD
                Call Chama_TelaECF("ArqPorData")
    
            Case MENU_MF_TABPROD
                Call CF_ECF("TabProd_Grava")
            
            Case MENU_MF_MVPORECF
                Call Chama_TelaECF("MVECF")
            
            Case MENU_MF_MEIOPAGTO
                Call Chama_TelaECF("MeioPagamentoPorData")
            
            Case MENU_MF_IDENTPAFECF
            
                Set objTela = Me
                
                Call CF_ECF("Requisito_XLIII", objTela)
    
            Case MENU_MF_INDICETECNICO
                Call Rotina_AvisoECF(vbOKOnly, AVISO_NAO_INDICETECNICO)
    
            Case MENU_MF_PARAMCONFIG
                
                Set objTela = Me
                
                Call CF_ECF("Rel_Parametros_Configuracao", objTela)
    
    
        End Select
    
    End If
    
    Exit Sub

Erro_mnuMF_Click:

    Select Case gErr

        Case 210887

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 210888)

    End Select

    Exit Sub

End Sub

Private Sub mnuMFArqMFD_Click(Index As Integer)

    Select Case Index
    
        Case MENU_MFARQ_PORDATA
                Call Chama_TelaECF("ArqPorData")
        
        Case MENU_MFARQ_PORCOO
                Call Chama_TelaECF("ArqPorCOO")
    
    End Select

End Sub

Private Sub mnuMFDAV_Click(Index As Integer)

    Select Case Index
    
        Case MENU_MFDAV_RELGERENCIAL
                Call Chama_TelaECF("DAVEmitidosRelGer")
        
        Case MENU_MFDAV_ARQUIVO
                Call Chama_TelaECF("DAVEmitidosArquivo")

    End Select

End Sub

Private Sub mnuMFEstoque_Click(Index As Integer)

    Select Case Index

        Case MENU_MF_REGPAFECF_TOTAL
            Call Chama_TelaECF("RegPAFECFTotal")

        Case MENU_MF_REGPAFECF_PARCIAL
            Call Chama_TelaECF("RegPAFECFParcial")

    End Select

End Sub

Private Sub mnuMFVP_Click(Index As Integer)

    Select Case Index
    
        Case MENU_MFVP_SINTEGRA
                Call Chama_TelaECF("VendaPeriodoSintegra")
        
        Case MENU_MFVP_SPED
                Call Chama_TelaECF("VendaPeriodoSPED")

    End Select

End Sub

Private Sub mnuMFEspelho_Click(Index As Integer)

    Select Case Index
    
        Case MENU_MFESPELHO_PORDATA
                Call Chama_TelaECF("EspelhoPorData1")
        
        Case MENU_MFESPELHO_PORCOO
                Call Chama_TelaECF("EspelhoPorCOO")
    
    End Select

End Sub

Private Sub mnuMFLMFC_Click(Index As Integer)

    Select Case Index
    
        Case MENU_MFLMFC_PORDATA
                Call Chama_TelaECF("LMFCPorData")
        
        Case MENU_MFLMFC_PORREDZ
                Call Chama_TelaECF("LMFCPorReducaoZ")
    
    End Select


End Sub

Private Sub mnuMFLMFS_Click(Index As Integer)

    Select Case Index
    
        Case MENU_MFLMFS_PORDATA
                Call Chama_TelaECF("LMFSPorData")
        
        Case MENU_MFLMFS_PORREDZ
                Call Chama_TelaECF("LMFSPorReducaoZ")
    
    End Select


End Sub

'Private Sub mnuMFMovECF_Click(Index As Integer)
'
'Dim lErro As Long
'
'On Error GoTo Erro_mnuMFMovECF_Click
'
'    Select Case Index
'
'        Case MENU_MFMOVECF_PORDATA
'                Call Chama_TelaECF("MvECFPorData")
'
'        Case MENU_MFMOVECF_PORECF
'                Call Chama_TelaECF("MVECFPorECF")
'
'
'    End Select
'
'
'    Exit Sub
'
'Erro_mnuMFMovECF_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 204595)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub mnuMovimentos_Click(Index As Integer)

End Sub

Private Sub Timer2_Timer()
    iCountBkp = iCountBkp + 1
    'Só executa o teste para ver se tem ou não que fazer o backup se:
    '1 - Já passou 5 minutos do último teste
    '2 - Se não tem nenhum transação aberta que possa ser prejudicada pela demora do backup
    '3 - Se o timer está ativo por conta de um backup habilitado
    '4 - Se não iniciou o processo (pode estar no meio de uma execução)
    If iCountBkp >= 5 And glTransacaoPAFECF = 0 And Timer2.Interval > 0 And giExeBkp = DESMARCADO Then
        iCountBkp = 0
        giExeBkp = MARCADO
        Call CF_ECF("Backup_Executa")
        giExeBkp = DESMARCADO
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
Dim lErro As Long

On Error GoTo Erro_Toolbar1_ButtonClick
    
    If Button.Index = 6 Then
        mnuECFSair_Click
    End If

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Verifica se já foi executa a redução z para a data de hoje
'        If gdtUltimaReducao = gdtDataHoje Then gError 111325

    End If

    If Button.Index = 1 Then
        Call Chama_TelaECF_Modal("VendaM")
    
    ElseIf Button.Index = 2 Then
        Call Chama_TelaECF_Modal("VendaP")
    
    ElseIf Button.Index = 3 Then
        Call Chama_TelaECF("Preco")
    
    ElseIf Button.Index = 4 Then
        If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105885
        Call Chama_TelaECF("ImpressaoCarne")
    
    ElseIf Button.Index = 5 Then
        Call Suspender_Sessao
    End If
    
    Exit Sub
    
Erro_Toolbar1_ButtonClick:

    Select Case gErr

        Case 105885
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
    
        Case 111325
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159201)

    End Select

    Exit Sub

End Sub

Private Sub Suspender_Sessao()

'Função que Suspende a Sessão

Dim objOperador As New ClassOperador
Dim iCogGerente As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoSessaoSuspende_Click

    'Se a Sessão Estiver Fechada então gera Erro
    If giStatusSessao = SESSAO_ENCERRADA Then gError 107605

    'Se Sessão estiver Suspensa
    If giStatusSessao = SESSAO_SUSPENSA Then gError 107606

    'Função que Executa a Suspenção da Sessão
    lErro = CF_ECF("Sessao_Executa_Suspensao")
    If lErro <> SUCESSO Then gError 107607

    'funcao que executa o termino da suspensao se a senha for digitada.
    lErro = CF_ECF("Sessao_Executa_Termino_Susp")
    If lErro <> SUCESSO Then gError 117542


    Exit Sub

Erro_BotaoSessaoSuspende_Click:

    Select Case gErr

        Case 107605
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_ABERTA_INEXISTENTE, gErr, giCodCaixa)

        Case 107606
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_SUSPENSA, gErr, giCodCaixa)

        Case 107607
            'Erros Tratados Dentro da Função Chamadora

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 159202)

    End Select

    Exit Sub

End Sub


Public Function objParent() As Object

    Set objParent = Me
    
End Function

Public Sub Refresh()

    Me.Show
    
End Sub

Private Sub Teste()
    Me.Show
End Sub

