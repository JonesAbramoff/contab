VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl TransfCartCobrOcx 
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6945
   Begin VB.Frame Frame1 
      Caption         =   "Situação Atual"
      Height          =   720
      Left            =   150
      TabIndex        =   12
      Top             =   1920
      Width           =   6630
      Begin VB.Label Carteira 
         Height          =   195
         Left            =   4005
         TabIndex        =   17
         Top             =   315
         Width           =   2400
      End
      Begin VB.Label Cobrador 
         Height          =   195
         Left            =   1170
         TabIndex        =   18
         Top             =   315
         Width           =   1785
      End
      Begin VB.Label label1007 
         Caption         =   "Carteira:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cobrador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   20
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   1740
      Left            =   165
      TabIndex        =   11
      Top             =   75
      Width           =   6630
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4425
         TabIndex        =   1
         Top             =   300
         Width           =   1920
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         ItemData        =   "TransfCartCobrOcx.ctx":0000
         Left            =   1005
         List            =   "TransfCartCobrOcx.ctx":0002
         TabIndex        =   2
         Top             =   825
         Width           =   2520
      End
      Begin VB.CommandButton BotaoTrazer 
         Caption         =   "Trazer Parcela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4395
         TabIndex        =   5
         Top             =   1305
         Width           =   1755
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   75
         Left            =   1530
         TabIndex        =   13
         Top             =   510
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   132
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   4425
         TabIndex        =   3
         Top             =   825
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   300
         Left            =   3615
         TabIndex        =   4
         Top             =   1305
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   990
         TabIndex        =   0
         Top             =   315
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   2115
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1305
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1020
         TabIndex        =   34
         Top             =   1305
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label16 
         Caption         =   "Emissão:"
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
         Left            =   210
         TabIndex        =   35
         Top             =   1350
         Width           =   750
      End
      Begin VB.Label ClienteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   375
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3840
         TabIndex        =   22
         Top             =   360
         Width           =   525
      End
      Begin VB.Label LabelTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   495
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   885
         Width           =   450
      End
      Begin VB.Label LabelParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2850
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label NumeroLabel 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3645
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   885
         Width           =   720
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   2355
      Picture         =   "TransfCartCobrOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4155
      Width           =   1020
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   525
      Left            =   3615
      Picture         =   "TransfCartCobrOcx.ctx":015E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4155
      Width           =   1020
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   3375
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataContabil 
      Height          =   300
      Left            =   2220
      TabIndex        =   8
      Top             =   3660
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame3 
      Caption         =   "Após a Transferência"
      Height          =   750
      Index           =   0
      Left            =   135
      TabIndex        =   15
      Top             =   2745
      Width           =   6630
      Begin VB.Label Label2 
         Caption         =   "Em Carteira"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   26
         Top             =   345
         Width           =   1290
      End
      Begin VB.Label Label3 
         Caption         =   "Própria Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   27
         Top             =   345
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Carteira:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3165
         TabIndex        =   28
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Cobrador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   345
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Após a Transferência"
      Height          =   735
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   6630
      Begin VB.ComboBox ComboCobrador 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   285
         Width           =   1920
      End
      Begin VB.ComboBox ComboCartCobrador 
         Height          =   315
         Left            =   4005
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   285
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Cobrador:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   210
         TabIndex        =   30
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label10 
         Caption         =   "Carteira:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3225
         TabIndex        =   31
         Top             =   345
         Width           =   735
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Data da Transferência:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   180
      TabIndex        =   32
      Top             =   3705
      Width           =   2040
   End
End
Attribute VB_Name = "TransfCartCobrOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private glNumIntParc As Long '# interno da parcela que está sendo alterada

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1
Private WithEvents objEventoTipoDoc As AdmEvento
Attribute objEventoTipoDoc.VB_VarHelpID = -1

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTransfCartCobr As ClassTransfCartCobr
Private gobjTituloReceber As New ClassTituloReceber
Private gobjParcelaReceber As New ClassParcelaReceber

'para evitar acessos desnecessarios durante o calculo de mnemonicos (contabilizacao)
Private giCartCobrInfoOK As Integer 'indica que as contas abaixo já foram lidas
Private gsContaCartCobrOrigem As String, gsContaCartCobrDestino As String, gsContaCobrDesconto As String 'contas contábeis associadas às carteiras de cobranca origem e destino da transferencia e cta de desconto de duplicata
Private giCobrancaDescontada As Integer

Private Sub BotaoCancela_Click()
    
    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 48920
    
    lErro = Limpa_Tela_Transf()
    If lErro <> SUCESSO Then Error 49526
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err

        Case 48920, 49526

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175382)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_Transf() As Long

    Call Limpa_Tela(Me)
    Filial.Clear
    
    ComboCobrador.ListIndex = 0
    ComboCartCobrador.ListIndex = 0
    Cobrador.Caption = ""
    Carteira.Caption = ""
    
    Call DateParaMasked(DataContabil, gdtDataAtual)
    
End Function

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objCobrador As New ClassCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim sCobrador As String
Dim sCarteira As String

On Error GoTo Erro_BotaoTrazer_Click

    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.Text)) = 0 Then Error 48921

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Error 48922

    'Verifica preenchimento do Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Error 48923

    'Verifica preenchimento de NumTítulo
    If Len(Trim(Numero.Text)) = 0 Then Error 48924

    'Verifica preenchimento da Parcela
    If Len(Trim(Parcela.ClipText)) = 0 Then Error 48925

    objCliente.sNomeReduzido = Cliente.Text

    'Lê Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 48926

    'Não encontrou o Cliente - - - > Erro
    If lErro = 12348 Then Error 48927

   'Preenche objTituloReceber
    gobjTituloReceber.iFilialEmpresa = giFilialEmpresa
    gobjTituloReceber.lCliente = objCliente.lCodigo
    gobjTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    gobjTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    gobjTituloReceber.lNumTitulo = CLng(Numero.Text)
    gobjTituloReceber.dtDataEmissao = StrParaDate(DataEmissao.Text)

    'Pesquisa no BD o Título Receber
    lErro = CF("TituloReceber_Le_NumeroInterno", gobjTituloReceber)
    If lErro <> SUCESSO And lErro <> 28574 Then Error 48928

    'Não encontrou o Título ==> Erro
    If lErro = 28574 Then Error 48929

    'Preenche objParcelaReceber
    gobjParcelaReceber.lNumIntTitulo = gobjTituloReceber.lNumIntDoc
    gobjParcelaReceber.iNumParcela = CInt(Parcela.Text)

    'Pesquisa no BD a Parcela
    lErro = CF("ParcelaReceber_Le_NumeroInterno", gobjParcelaReceber)
    If lErro <> SUCESSO And lErro <> 28590 Then Error 48930

    'Não encontrou a Parcela ==> Erro
    If lErro = 28590 Then Error 48931

    If gobjParcelaReceber.iCobrador = COBRADOR_PROPRIA_EMPRESA And gobjParcelaReceber.iCarteiraCobranca = CARTEIRA_CHEQUEPRE Then Error 41596
    
    'Preenche objCobrador
    objCobrador.iCodigo = gobjParcelaReceber.iCobrador

    'Lê Cobrador
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 48932

    'If objCobrador.iCobrancaEletronica = COBRANCA_ELETRONICA Then Error 48933

    'Preenche objCarteiraCobranca
    objCarteiraCobranca.iCodigo = gobjParcelaReceber.iCarteiraCobranca

    'Lê CarteiraCobranca
    lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
    If lErro <> SUCESSO And lErro <> 23413 Then Error 48934

    'Preenche cobrador e a crteira respctiva daquele cobrador

    'Preenche Cobrador da tela
    sCobrador = CStr(objCobrador.iCodigo) + SEPARADOR + objCobrador.sNomeReduzido
    Cobrador.Caption = sCobrador

    'Preenche Carteira na Tela
    sCarteira = CStr(objCarteiraCobranca.iCodigo) + SEPARADOR + objCarteiraCobranca.sDescricao
    Carteira.Caption = sCarteira
    
    'preencho o global numero interno da parcela
    glNumIntParc = gobjParcelaReceber.lNumIntDoc

    If objCarteiraCobranca.iCodigo = CARTEIRA_CARTEIRA And objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA Then

        Frame3(1).Visible = True
        Frame3(0).Visible = False

    Else

        Frame3(0).Visible = True
        Frame3(1).Visible = False

    End If
    
    Exit Sub

Erro_BotaoTrazer_Click:

    Select Case Err

        Case 48921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
            Cliente.SetFocus
            
        Case 48922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            Filial.SetFocus
            
        Case 48923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", Err)
            Tipo.SetFocus
            
        Case 48924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", Err, Error$)
            Numero.SetFocus
            
        Case 48925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_PREENCHIDA", Err, Error$)
            Parcela.SetFocus
            
        Case 48926, 48928, 48930, 48932, 48934
        
        Case 48927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)
            Cliente.SetFocus
            
        Case 48929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO2", Err, gobjTituloReceber.iFilialEmpresa, gobjTituloReceber.lCliente, gobjTituloReceber.iFilial, gobjTituloReceber.sSiglaDocumento, gobjTituloReceber.lNumTitulo)

        Case 48931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_NAO_CADASTRADA", Err, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)
        
        Case 48933
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_COBR_MANUAL", Err)
        
        Case 41596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_TRANSF_CHEQUEPRE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175383)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    glNumIntParc = 0
    
    Call Cliente_Preenche
    
End Sub

Private Sub DataContabil_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataContabil)

End Sub

Private Sub Filial_Change()

    glNumIntParc = 0
    
End Sub

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    If Filial.ListIndex = -1 Then Exit Sub
    
    glNumIntParc = 0
    
    'Verifica se Cliente e Filial estão preenchidas
    If Len(Trim(Cliente.Text)) > 0 And Len(Trim(Filial.Text)) > 0 Then

        Call Filial_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_Filial_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175384)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sCliente As String
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

        'Verifica se a filial foi preenchida
        If Len(Trim(Filial.Text)) = 0 Then Exit Sub

        'Verifica se é uma filial selecionada
        If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

        'Tenta selecionar na combo
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 48935

        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then

            'Verifica se o Cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then Error 48936

            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo

            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then Error 48937

            If lErro = 17660 Then Error 48938

            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        End If

        'Não encontrou a STRING
        If lErro = 6731 Then Error 48939

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 48935, 48937

        Case 48936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 48938
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 48939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175385)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    'Inicializa os Eventos da Tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoParcela = New AdmEvento
    Set objEventoTipoDoc = New AdmEvento

    'Carrega os Tipos de Documento
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then Error 48940

    'Carrega as combos de carteira e cobrador
    lErro = Carrega_Cobrador()
    If lErro <> SUCESSO Then Error 48941
    
    Call DateParaMasked(DataContabil, gdtDataAtual)
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 48940, 48941
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175386)

    End Select

    Exit Sub

End Sub

Private Sub ComboCartCobrador_Click()

Dim objCobrador As New ClassCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_ComboCartCobrador_Click
    
    If ComboCartCobrador.ListIndex = -1 Then Exit Sub
    
    'Se a Carteira foi preenchida
    If Len(Trim(ComboCartCobrador.Text)) <> 0 Then

        'Se é a Carteira selecionada na Combo
        If ComboCartCobrador.Text = ComboCartCobrador.List(ComboCartCobrador.ListIndex) Then Exit Sub

        'Verifica se a Carteira existe na Combo. Se existir, seleciona
        lErro = Combo_Seleciona(ComboCartCobrador, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 Then Error 48942

        'Se a Carteira(CODIGO) não existe na Combo
        If lErro = 6730 Then

            'Verifica se o Cobrador foi digitado
            If Len(Trim(ComboCobrador.Text)) = 0 Then Error 48943

            'Passa os Códigos da Carteira para o Obj
            objCarteiraCobranca.iCodigo = CInt(ComboCartCobrador.Text)

            'Pesquisa se existe Carteira com o código em questão
            lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
            If lErro <> SUCESSO And lErro <> 23413 Then Error 48944

            If lErro = 23413 Then Error 48945

        End If

    End If

    Exit Sub

Erro_ComboCartCobrador_Click:

    Select Case Err

        Case 48942, 48944
            ComboCartCobrador.SetFocus

        Case 48943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_INFORMADO", Err)
            ComboCobrador.SetFocus

        Case 48945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)
            ComboCartCobrador.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175387)

    End Select

    Exit Sub

End Sub

Private Sub ComboCobrador_Click()

Dim iCodCobrador As Integer
Dim objCobrador As New ClassCobrador
Dim lErro As Long
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim sListBoxItem As String
Dim colCarteirasCobrador As New Collection

On Error GoTo Erro_ComboCobrador_Click
    
    If ComboCobrador.ListIndex = -1 Then Exit Sub
    
    'Limpa a Combo de Carteiras
    ComboCartCobrador.Clear

    'Se Cobrador está preenchido
    If Len(Trim(ComboCobrador.Text)) <> 0 Then

        'Extrai o código do Cobrador
        iCodCobrador = Codigo_Extrai(ComboCobrador.Text)

        'Passa o Código do Cobrador que está na tela para o Obj
        objCobrador.iCodigo = iCodCobrador

        'Lê os dados do Cobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 48946

        'Se o Cobrador não estiver cadastrado
        If lErro = 19294 Then Error 48947

        'Le as carteiras associadas ao Cobrador
        lErro = CF("Cobrador_Le_Carteiras", objCobrador, colCarteirasCobrador)
        If lErro <> SUCESSO And lErro <> 23500 Then Error 48948

        If lErro = SUCESSO Then

            'Preencher a Combo
            For Each objCarteiraCobrador In colCarteirasCobrador

                objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca

                If iCodCobrador <> COBRADOR_PROPRIA_EMPRESA Or (objCarteiraCobranca.iCodigo <> CARTEIRA_CARTEIRA And objCarteiraCobranca.iCodigo <> CARTEIRA_CHEQUEPRE) Then
                
                    lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
                    If lErro <> SUCESSO And lErro <> 23413 Then Error 48949
    
                    'Carteira não está cadastrado
                    If lErro = 23413 Then Error 48950
    
                    'Concatena Código e a Descricao da carteira
                    sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
                    sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
    
                    ComboCartCobrador.AddItem sListBoxItem
                    ComboCartCobrador.ItemData(ComboCartCobrador.NewIndex) = objCarteiraCobranca.iCodigo
            
                End If
                
            Next
            
            ComboCartCobrador.ListIndex = 0
    
        End If

    
    End If

    Exit Sub

Erro_ComboCobrador_Click:

    Select Case Err

        Case 48946, 48948, 48949
            ComboCobrador.SetFocus

        Case 48947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, ComboCobrador.Text)
            ComboCobrador.SetFocus

        Case 48950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175388)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Cobrador() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Carrega_Cobrador

    'Lê cada código e Nome Reduzido da tabela cobradores
    lErro = CF("Cod_Nomes_Le", "Cobradores", "Codigo", "NomeReduzido", STRING_COBRADOR_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 48951

    For Each objCodigoDescricao In colCodigoDescricao

        ComboCobrador.AddItem objCodigoDescricao.iCodigo & SEPARADOR & objCodigoDescricao.sNome
        ComboCobrador.ItemData(ComboCobrador.NewIndex) = objCodigoDescricao.iCodigo

    Next
    
    If ComboCobrador.ListCount > 0 Then ComboCobrador.ListIndex = 0
    
    Call ComboCobrador_Click
    
    Carrega_Cobrador = SUCESSO

    Exit Function

Erro_Carrega_Cobrador:

    Carrega_Cobrador = Err

    Select Case Err
        
        Case 48951
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175389)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoDocumento()

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Lê os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then Error 48952

    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = Err

    Select Case Err

        Case 48952

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175390)

    End Select

    Exit Function

End Function

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then Error 48953

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then Error 48954

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
         Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    Else
        If Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear

        End If
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err

        Case 48953
            
        Case 48954

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175391)

    End Select

    Exit Sub

End Sub

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Numero_Change()
    
    glNumIntParc = 0
    
End Sub

Private Sub Numero_GotFocus()
    
Dim lNumAux As Long
    
    lNumAux = glNumIntParc
    Call MaskEdBox_TrataGotFocus(Numero)
    glNumIntParc = lNumAux

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(Cancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloReceber = obj1

    lErro = Traz_TitReceber_Tela(objTituloReceber)
    If lErro <> SUCESSO Then gError 87215

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 87215
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175392)

    End Select

    Exit Sub

End Sub

Private Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objParcelaReceber = obj1

    If Not (objParcelaReceber Is Nothing) Then
        Parcela.PromptInclude = False
        Parcela.Text = CStr(objParcelaReceber.iNumParcela)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bSGECancelDummy)
    End If

    Me.Show

    Call BotaoTrazer_Click
    
    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175393)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDoc_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDocumento.sSigla

    Call Tipo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175394)

     End Select

     Exit Sub

End Sub

Private Sub LabelParcela_Click()
'lista as parcelas do titulo selecionado

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objParcelaReceber As ClassParcelaReceber
Dim colSelecao As New Collection

On Error GoTo Erro_LabelParcela_Click

    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Cliente.ClipText)) = 0 Then Error 48955
    If Len(Trim(Filial.Text)) = 0 Then Error 48956
    If Len(Trim(Tipo.Text)) = 0 Then Error 48957
    If Len(Trim(Numero.ClipText)) = 0 Then Error 48958

    objCliente.sNomeReduzido = Cliente.Text
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 48959

    'Se não achou o Cliente --> erro
    If lErro <> SUCESSO Then Error 48960

    colSelecao.Add objCliente.lCodigo
    colSelecao.Add Codigo_Extrai(Filial.Text)
    colSelecao.Add SCodigo_Extrai(Tipo.Text)
    colSelecao.Add StrParaLong(Numero.Text)
    
    'Chama a tela
    Call Chama_Tela("ParcelasRecLista", colSelecao, objParcelaReceber, objEventoParcela)

    Exit Sub

Erro_LabelParcela_Click:

    Select Case Err

        Case 48955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 48956
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 48957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 48958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 48959

        Case 48960
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175395)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipo_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = SCodigo_Extrai(Tipo.Text)

    'Chama a tela TipoDocTituloRecLista
    Call Chama_Tela("TipoDocTituloRecLista", colSelecao, objTipoDocumento, objEventoTipoDoc)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Número está preenchido
    If Len(Trim(Numero.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(Numero.ClipText)
    If lErro <> SUCESSO Then Error 48961

    Exit Sub

Erro_Numero_Validate:

    Cancel = True


    Select Case Err

        Case 48961

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175396)

    End Select

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloReceber As New ClassTituloReceber
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim lErro As Long
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_NumeroLabel_Click

    If Len(Trim(Cliente.Text)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 87208
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 87209
    
    End If
    
    'Guarda o código no objTituloReceber
    objTituloReceber.lCliente = objCliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)

    'Verifica se os obj(s) estão preenchidos antes de serem incluídos na coleção
    If objTituloReceber.lCliente <> 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        colSelecao.Add (objTituloReceber.lCliente)
    End If

    If objTituloReceber.iFilial <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        colSelecao.Add (objTituloReceber.iFilial)
    End If

    If Len(Trim(objTituloReceber.sSiglaDocumento)) <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND SiglaDocumento = ?"
        Else
            iPreenchido = 1
            sSelecao = "SiglaDocumento = ?"
        End If
        colSelecao.Add (objTituloReceber.sSiglaDocumento)
    End If

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TituloReceberLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 87208

        Case 87209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175397)

    End Select

    Exit Sub

End Sub

Private Sub Parcela_Change()

    glNumIntParc = 0
    
End Sub

Private Sub Parcela_GotFocus()
    
Dim lNumAux As Long
    
    lNumAux = glNumIntParc
    Call MaskEdBox_TrataGotFocus(Parcela)
    glNumIntParc = lNumAux

End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    'Verifica se está preenchido
    If Len(Trim(Parcela.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Valor_Positivo_Critica(Parcela.ClipText)
    If lErro <> SUCESSO Then Error 48967

    Exit Sub

Erro_Parcela_Validate:

    Cancel = True


    Select Case Err

        Case 48967

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175398)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()

    glNumIntParc = 0
    
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se o Tipo foi selecionado
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Tenta localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 48968

    'Se não encontrar -> Erro
    If lErro = 60483 Then Error 48969

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case Err

        Case 48968

        Case 48969
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175399)

    End Select

    Exit Sub

End Sub

Private Sub DataContabil_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'Verifica se a DataContabil foi digitada
    If Len(Trim(DataContabil.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataContabil.Text)
    If lErro <> SUCESSO Then Error 48970

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True


    Select Case Err

        Case 48970

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175400)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    'Diminui a DataContabil em um dia
    lErro = Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
    If lErro Then Error 48971

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 48971

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175401)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    'Aumenta a DataContabil em um dia
    lErro = Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
    If lErro Then Error 48972

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 48972

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175402)

    End Select

    Exit Sub

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoParcela = Nothing
    Set objEventoTipoDoc = Nothing

    Set gobjContabAutomatica = Nothing
    Set gobjTransfCartCobr = Nothing
    Set gobjTituloReceber = Nothing
    Set gobjParcelaReceber = Nothing
    
End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTransfCartCobr As New ClassTransfCartCobr

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.Text)) = 0 Then Error 48973

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Error 48974

    'Verifica preenchimento do Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Error 48975

    'Verifica preenchimento de NumTítulo
    If Len(Trim(Numero.Text)) = 0 Then Error 48976

    'Verifica preenchimento da Parcela
    If Len(Trim(Parcela.ClipText)) = 0 Then Error 48977
    
    'Verifica o preenchimento da data
    If Len(Trim(DataContabil.ClipText)) = 0 Then Error 48990
    
    If glNumIntParc = 0 Then Error 48978
        
    'se esta tudo preenchido entao testa e quarda no objeto
    lErro = MoveTela_Memoria(objTransfCartCobr)
    If lErro <> SUCESSO Then Error 48979
        
    objTransfCartCobr.objTelaAtualizacao = Me
    lErro = CF("TransfCartCobr_Grava", objTransfCartCobr)
    If lErro <> SUCESSO Then Error 48980
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = Err

        Select Case Err
            
            Case 48973
                lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
                Cliente.SetFocus
            
            Case 48974
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
                Filial.SetFocus
                
            Case 48975
                lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", Err)
                Tipo.SetFocus
                
            Case 48976
                lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", Err, Error$)
                Numero.SetFocus
                
            Case 48977
                lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_PREENCHIDA", Err, Error$)
                Parcela.SetFocus
            
            Case 48978
                lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLAR_BOTAO_TRAZER", Err)
                            
            Case 48979, 48980
            
            Case 48990
                lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
                DataContabil.SetFocus
                
            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175403)

        End Select

    Exit Function

End Function

Function MoveTela_Memoria(objTransfCartCobr As ClassTransfCartCobr) As Long

Dim lErro As Long

On Error GoTo Erro_MoveTela_Memoria
    
    objTransfCartCobr.dtData = CDate(DataContabil.Text)
    objTransfCartCobr.dtDataRegistro = gdtDataAtual
    objTransfCartCobr.lNumIntParc = glNumIntParc
    
    If Frame3(1).Visible = False Then
        objTransfCartCobr.iCarteiraCobranca = CARTEIRA_CARTEIRA
        objTransfCartCobr.iCobrador = COBRADOR_PROPRIA_EMPRESA
    Else
        objTransfCartCobr.iCarteiraCobranca = ComboCartCobrador.ItemData(ComboCartCobrador.ListIndex)
        objTransfCartCobr.iCobrador = ComboCobrador.ItemData(ComboCobrador.ListIndex)
    End If
    
    MoveTela_Memoria = SUCESSO
    
    Exit Function
    
Erro_MoveTela_Memoria:

    MoveTela_Memoria = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175404)
    
    End Select
    
    Exit Function
    
End Function

Private Function CarteiraCobrador_ObtemInfoContab() As Long
'funcao auxiliar a calcula_mnemonico para a obtencao de informacoes
'associadas à carteira de cobranca do cobrador diferente da propria empresa para onde foi ou estava a parcela

Dim lErro As Long, objCarteiraCobrador As New ClassCarteiraCobrador, objMnemonico As New ClassMnemonicoCTBValor
Dim sContaTela As String, sContaCarteira As String, sCampoGlobal As String, sContaCobr As String

On Error GoTo Erro_CarteiraCobrador_ObtemInfoContab

    'obtem conta contabil da carteira "em carteira" do cobrador "própria empresa"
    objMnemonico.sMnemonico = "CtaReceberCarteira"
    lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
    If lErro <> SUCESSO And lErro <> 39690 Then Error 56805
    If lErro <> SUCESSO Then Error 56806
    
    sContaCarteira = objMnemonico.sValor
    
    'se nao estava "em carteira"
    If gobjParcelaReceber.iCobrador <> COBRADOR_PROPRIA_EMPRESA Or gobjParcelaReceber.iCarteiraCobranca <> CARTEIRA_CARTEIRA Then
    
        objCarteiraCobrador.iCobrador = gobjParcelaReceber.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = gobjParcelaReceber.iCarteiraCobranca
        
    Else
    
        objCarteiraCobrador.iCobrador = gobjTransfCartCobr.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = gobjTransfCartCobr.iCarteiraCobranca
    
    End If
    
    If objCarteiraCobrador.iCobrador <> COBRADOR_PROPRIA_EMPRESA Then
    
        lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 32234
        If lErro <> SUCESSO Then Error 56810
        
        If objCarteiraCobrador.sContaDuplDescontadas <> "" Then
        
            lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaDuplDescontadas, gsContaCobrDesconto)
            If lErro <> SUCESSO Then Error 59308
            
        Else
        
            gsContaCobrDesconto = ""
            
        End If
        
        If objCarteiraCobrador.iCodCarteiraCobranca <> CARTEIRA_DESCONTADA Then
            giCobrancaDescontada = 0
        Else
            giCobrancaDescontada = 1
        End If
        
        sContaCobr = objCarteiraCobrador.sContaContabil
        If sContaCobr <> "" Then
        
            lErro = Mascara_RetornaContaTela(sContaCobr, sContaTela)
            If lErro <> SUCESSO Then Error 32235
        
            sContaCobr = sContaTela
            
        End If
        
    Else
    
        Select Case objCarteiraCobrador.iCodCarteiraCobranca
        
            Case CARTEIRA_JURIDICO
                sCampoGlobal = "CtaJuridico"
            
            Case Else
                Error 56807
                
        End Select
        
        objMnemonico.sMnemonico = sCampoGlobal
        lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
        If lErro <> SUCESSO And lErro <> 39690 Then Error 56808
        If lErro <> SUCESSO Then Error 56809
        
        sContaCobr = objMnemonico.sValor
    
    End If
    
    'se nao estava "em carteira"
    If gobjParcelaReceber.iCobrador <> COBRADOR_PROPRIA_EMPRESA Or gobjParcelaReceber.iCarteiraCobranca <> CARTEIRA_CARTEIRA Then
    
        gsContaCartCobrOrigem = sContaCobr
        gsContaCartCobrDestino = sContaCarteira
        
    Else
    
        gsContaCartCobrOrigem = sContaCarteira
        gsContaCartCobrDestino = sContaCobr
    
    End If
    
    giCartCobrInfoOK = 1
    
    CarteiraCobrador_ObtemInfoContab = SUCESSO
     
    Exit Function
    
Erro_CarteiraCobrador_ObtemInfoContab:

    CarteiraCobrador_ObtemInfoContab = Err
     
    Select Case Err
          
        Case 32234, 32235, 32236, 56805, 56808, 59308
        
        Case 56807, 56810
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56806, 56809
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175405)
     
    End Select
     
    Exit Function

End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case "Cobrador_Codigo"
            objMnemonicoValor.colValor.Add gobjTransfCartCobr.iCobrador
            
        Case "CartCobrOrigem_Conta"
        
            If giCartCobrInfoOK = 0 Then
            
                lErro = CarteiraCobrador_ObtemInfoContab
                If lErro <> SUCESSO Then Error 56535
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaCartCobrOrigem
            
        Case "CartCobrDest_Conta"
        
            If giCartCobrInfoOK = 0 Then
            
                lErro = CarteiraCobrador_ObtemInfoContab
                If lErro <> SUCESSO Then Error 56536
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaCartCobrDestino
            
        Case "CartCobr_CtaDesconto"
        
            If giCartCobrInfoOK = 0 Then
            
                lErro = CarteiraCobrador_ObtemInfoContab
                If lErro <> SUCESSO Then Error 59307
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaCobrDesconto
            
        Case "Cobranca_Descontada"
        
            If giCartCobrInfoOK = 0 Then
            
                lErro = CarteiraCobrador_ObtemInfoContab
                If lErro <> SUCESSO Then Error 59309
                
            End If
            
            objMnemonicoValor.colValor.Add giCobrancaDescontada
            
        Case "Valor_Cobrado"

            objMnemonicoValor.colValor.Add gobjTransfCartCobr.dSaldo

        Case "Cliente_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.lCliente
        
        Case "FilialCli_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.iFilial
        
        Case "Titulo_Numero"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.lNumTitulo
        
        Case "Titulo_Filial"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.iFilialEmpresa
        
        Case "Parcela_Numero"
        
            objMnemonicoValor.colValor.Add gobjParcelaReceber.iNumParcela
                                                        
        Case Else

            Error 56537

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 56535, 56536, 59307

        Case 56537
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175406)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada na atualizacao da transferencia no bd e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iFilialEmpresaLote As Integer, lNumIntDoc As Long

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjTransfCartCobr = vParams(0)
    
    giCartCobrInfoOK = 0
    
    'teste de filiais com autonomia contabil
    iFilialEmpresaLote = IIf(giContabCentralizada, giFilialEmpresa, gobjTituloReceber.iFilialEmpresa)
    
    'obtem numero de doc
    lErro = objContabAutomatica.Obter_Doc(lDoc, iFilialEmpresaLote)
    If lErro <> SUCESSO Then Error 32232

    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, "TransfCartCobr", gobjTransfCartCobr.lNumIntDoc, gobjTituloReceber.lCliente, gobjTituloReceber.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, iFilialEmpresaLote)
    If lErro <> SUCESSO Then Error 32233
            
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32232, 32233
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175407)
     
    End Select
     
    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TRANSFERENCIA_CARTEIRA_COBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Transferência de Carteira de Cobrança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TransfCartCobr"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Parcela Then
            Call LabelParcela_Click
        End If
    
    End If
    
End Sub

Private Sub Carteira_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Carteira, Source, X, Y)
End Sub

Private Sub Carteira_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Carteira, Button, Shift, X, Y)
End Sub

Private Sub Cobrador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cobrador, Source, X, Y)
End Sub

Private Sub Cobrador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cobrador, Button, Shift, X, Y)
End Sub

Private Sub label1007_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(label1007, Source, X, Y)
End Sub

Private Sub label1007_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label1007, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub

Private Sub LabelParcela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelParcela, Source, X, Y)
End Sub

Private Sub LabelParcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelParcela, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub


Function Traz_TitReceber_Tela(objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TitReceber_Tela

    'Lê o Título à Receber
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 87213

    'Não encontrou o Título à Receber --> erro
    If lErro = 26061 Then gError 87214
    
    'Coloca o Cliente na Tela
    Cliente.Text = objTituloReceber.lCliente
    Call Cliente_Validate(bSGECancelDummy)

    'Coloca a Filial na Tela
    Filial.Text = objTituloReceber.iFilial
    Call Filial_Validate(bSGECancelDummy)
    
    'Coloca o Tipo na tela
    Tipo.Text = objTituloReceber.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)

   If objTituloReceber.lNumTitulo = 0 Then
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
    Else
        Numero.PromptInclude = False
        Numero.Text = CStr(objTituloReceber.lNumTitulo)
        Numero.PromptInclude = True
    End If

    Call Numero_Validate(bSGECancelDummy)

    Call DateParaMasked(DataEmissao, objTituloReceber.dtDataEmissao)
    Call DataEmissao_Validate(bSGECancelDummy)
    
    Me.Show

Traz_TitReceber_Tela = SUCESSO

    Exit Function

Erro_Traz_TitReceber_Tela:

    Traz_TitReceber_Tela = gErr

    Select Case gErr

        Case 87213
        
        Case 87214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175408)

    End Select

    Exit Function

End Function

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134024

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134024

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175409)

    End Select
    
    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
Dim lNumAux As Long
Dim iAlterado As Integer
    
    lNumAux = glNumIntParc
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    glNumIntParc = lNumAux
    
End Sub

Public Sub DataEmissao_Change()

    'Se Número foi alterado zera glNUmIntParc
    glNumIntParc = 0

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate
    
    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 26140
        
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case Err

        Case 26140
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175191)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 26141

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 26141

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175195)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 26142

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 26142

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175196)

    End Select

    Exit Sub

End Sub

