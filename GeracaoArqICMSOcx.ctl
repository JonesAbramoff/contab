VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoArqICMSOcx 
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   ScaleHeight     =   4845
   ScaleWidth      =   5190
   Begin VB.Frame Frame2 
      Caption         =   "Endereço da Empresa"
      Height          =   1485
      Left            =   150
      TabIndex        =   13
      Top             =   3240
      Width           =   4905
      Begin VB.TextBox Endereco 
         Height          =   285
         Left            =   1800
         MaxLength       =   34
         TabIndex        =   6
         Top             =   300
         Width           =   2895
      End
      Begin VB.TextBox Complemento 
         Height          =   285
         Left            =   1800
         MaxLength       =   22
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   690
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   885
         TabIndex        =   17
         Top             =   345
         Width           =   885
      End
      Begin VB.Label Label6 
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
         Left            =   1050
         TabIndex        =   18
         Top             =   705
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
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
         Left            =   570
         TabIndex        =   19
         Top             =   1125
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa"
      Height          =   1545
      Left            =   150
      TabIndex        =   16
      Top             =   1560
      Width           =   4905
      Begin VB.TextBox Contato 
         Height          =   285
         Left            =   1770
         MaxLength       =   28
         TabIndex        =   4
         Top             =   660
         Width           =   2895
      End
      Begin VB.TextBox TelContato 
         Height          =   285
         Left            =   1770
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1050
         Width           =   1725
      End
      Begin VB.TextBox NomeEmpresa 
         Height          =   285
         Left            =   1770
         MaxLength       =   35
         TabIndex        =   3
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Telef. de Contato:"
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
         Left            =   180
         TabIndex        =   20
         Top             =   1065
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Left            =   1005
         TabIndex        =   21
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Empresa:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   315
         Width           =   1605
      End
   End
   Begin VB.TextBox NomeArquivo 
      Height          =   285
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   2805
      ScaleHeight     =   495
      ScaleWidth      =   2190
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2250
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1635
         Picture         =   "GeracaoArqICMSOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   330
         Left            =   120
         Picture         =   "GeracaoArqICMSOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   930
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1125
         Picture         =   "GeracaoArqICMSOcx.ctx":0910
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDownDataInicial 
      Height          =   300
      Left            =   1965
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   300
      Left            =   795
      TabIndex        =   0
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataFinal 
      Height          =   300
      Left            =   1965
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   300
      Left            =   810
      TabIndex        =   1
      Top             =   600
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Arquivo:"
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
      Left            =   330
      TabIndex        =   23
      Top             =   1230
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "De:"
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
      Left            =   390
      TabIndex        =   24
      Top             =   150
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Até:"
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
      Left            =   345
      TabIndex        =   25
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "GeracaoArqICMSOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'
''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
'Dim iAlterado As Integer
'
'Private Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
'Private Sub BotaoLimpar_Click()
'
'    Call Limpa_Tela(Me)
'
'End Sub
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'Dim objFilial As New AdmFiliais
'
'On Error GoTo Erro_Form_Load
'
'    Call Inicializa_Datas
'
'    objFilial.iCodFilial = giFilialEmpresa
'
'    lErro = CF("FilialEmpresa_Le", objFilial)
'    If lErro <> SUCESSO Then Error 53075
'
'    'default p/telefone e contato a partir de objFilial.objEndereco.sContato e objFilial.objEndereco.sTelefone1
'    Contato.Text = objFilial.objEndereco.sContato
'    TelContato.Text = objFilial.objEndereco.sTelefone1
'    NomeEmpresa.Text = gsNomeEmpresa
'    NomeArquivo.Text = "ArquivoICMS.txt"
'    Endereco.Text = objFilial.objEndereco.sEndereco
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = Err
'
'    Select Case Err
'
'        Case 53075
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160749)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoSeguir_Click()
'
'Dim lErro As String
'Dim dtDataFinal As Date
'Dim dtDataInicial As Date
'Dim sNomeArqParam As String
'
'On Error GoTo Erro_BotaoSeguir_Click
'
'    'verificar se os campos obrigatorios estao preenchidos
'    If Len(Trim(DataInicial.ClipText)) = 0 Then Error 53080
'    If Len(Trim(DataFinal.ClipText)) = 0 Then Error 53081
'    If Len(Trim(NomeArquivo.Text)) = 0 Then Error 61487
'    If Len(Trim(NomeEmpresa.Text)) = 0 Then Error 61490
'    If Len(Trim(Contato.Text)) = 0 Then Error 53082
'    If Len(Trim(TelContato.Text)) = 0 Then Error 53083
'    If Len(Trim(Endereco.Text)) = 0 Then Error 60365
'    If Len(Trim(Numero.Text)) = 0 Then Error 60366
'    If Len(Trim(Complemento.Text)) = 0 Then Error 60367
'
'    'validar se a data inicial é menor que a final
'    dtDataFinal = MaskedParaDate(DataFinal)
'    dtDataInicial = MaskedParaDate(DataInicial)
'    If dtDataInicial > dtDataFinal Then Error 53084
'
'    'chamar tela que irá efetuar a criacao do arquivo utilizando a classe ClassGeracaoArqICMS
'
'    lErro = Sistema_Preparar_Batch(sNomeArqParam)
'    If lErro <> SUCESSO Then Error 64000
'
'    Call CriarArquivo(dtDataInicial, dtDataFinal, Contato.Text, TelContato.Text, Endereco.Text, CLng(Numero.Text), Complemento.Text, NomeArquivo.Text, NomeEmpresa.Text, sNomeArqParam)
'
'    Unload Me
'
'    Exit Sub
'
'Erro_BotaoSeguir_Click:
'
'    lErro_Chama_Tela = Err
'
'    Select Case Err
'
'        Case 53080
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", Err)
'
'        Case 53081
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", Err)
'
'        Case 53082
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_PREENCHIDO", Err)
'
'        Case 53083
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELCONTATO_NAO_PREENCHIDO", Err)
'
'        Case 53084
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case 60365
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_PREENCHIDO", Err)
'
'        Case 60366
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", Err)
'
'        Case 60367
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPLEMENTO_NAO_PREENCHIDO", Err)
'
'        Case 61487
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", Err)
'
'        Case 61490
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_PREENCHIDA", Err)
'
'        Case 64000
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160750)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Contato_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataFinal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dtUltimoDiaMes As Date
'
'On Error GoTo Erro_DataFinal_Validate
'
'    'verifica se a data está preenchida
'    If Len(Trim(DataFinal.ClipText)) > 0 Then
'
'        'verifica se a data final é válida
'        lErro = Data_Critica(DataFinal.Text)
'        If lErro <> SUCESSO Then Error 58850
'
'        If Month(CDate(DataFinal.Text)) = 12 Then
'            dtUltimoDiaMes = CDate("01/" & Month(CDate(DataFinal.Text) + 1) & "/" & (1 + Year(CDate(DataFinal.Text)))) - 1
'        Else
'            dtUltimoDiaMes = CDate("01/" & Month(CDate(DataFinal.Text) + 1) & "/" & Year(CDate(DataFinal.Text))) - 1
'        End If
'        If CDate(DataFinal.Text) <> dtUltimoDiaMes Then Error 58851
'
'        If Len(Trim(DataInicial.ClipText)) > 0 Then
'            If DataInicial.Text > DataFinal.Text Then Error 58852
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_DataFinal_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 58850
'
'        Case 58851
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_DO_MES", Err, dtUltimoDiaMes)
'
'        Case 58852
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160751)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataInicial_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataInicial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dtPrimeiroDiaMes As Date
'
'On Error GoTo Erro_DataInicial_Validate
'
'    'verifica se a data está preenchida
'    If Len(Trim(DataInicial.ClipText)) > 0 Then
'
'        'verifica se a data final é válida
'        lErro = Data_Critica(DataInicial.Text)
'        If lErro <> SUCESSO Then Error 58850
'
'        dtPrimeiroDiaMes = CDate("01/" & Month(CDate(DataInicial.Text)) & "/" & Year(CDate(DataInicial.Text)))
'
'        If CDate(DataInicial.Text) <> dtPrimeiroDiaMes Then Error 58851
'
'        If Len(Trim(DataFinal.ClipText)) > 0 Then
'            If DataInicial.Text > DataFinal.Text Then Error 58852
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_DataInicial_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 58850
'
'        Case 58851
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_DO_MES", Err, dtPrimeiroDiaMes)
'
'        Case 58852
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160752)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub NomeArquivo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub NomeEmpresa_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Numero_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(Numero)
'
'End Sub
'
'Private Sub TelContato_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub UpDownDataFinal_DownClick()
''diminui a data final
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownDataFinal_DownClick
'
'    DataFinal.SetFocus
'
'    If Len(DataFinal.ClipText) > 0 Then
'
'        sData = DataFinal.Text
'
'        lErro = Data_Diminui(sData)
'        If lErro <> SUCESSO Then Error 53076
'
'        DataFinal.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownDataFinal_DownClick:
'
'    Select Case Err
'
'        Case 53076
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160753)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownDataFinal_UpClick()
''aumenta a data final
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownDataFinal_UpClick
'
'    DataFinal.SetFocus
'
'    If Len(DataFinal.ClipText) > 0 Then
'
'        sData = DataFinal.Text
'
'        lErro = Data_Aumenta(sData)
'        If lErro <> SUCESSO Then Error 53077
'
'        DataFinal.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownDataFinal_UpClick:
'
'    Select Case Err
'
'        Case 53077
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160754)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownDataInicial_DownClick()
''diminui a data inicial
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownDataInicial_DownClick
'
'    DataInicial.SetFocus
'
'    If Len(DataInicial.ClipText) > 0 Then
'
'        sData = DataInicial.Text
'
'        lErro = Data_Diminui(sData)
'        If lErro <> SUCESSO Then Error 53078
'
'        DataInicial.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownDataInicial_DownClick:
'
'    Select Case Err
'
'        Case 53078
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160755)
'
'    End Select
'
'    Exit Sub
'
'
'End Sub
'
'Private Sub UpDownDataInicial_UpClick()
''aumenta a data inicial
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownDataInicial_UpClick
'
'    DataInicial.SetFocus
'
'    If Len(DataInicial.ClipText) > 0 Then
'
'        sData = DataInicial.Text
'
'        lErro = Data_Aumenta(sData)
'        If lErro <> SUCESSO Then Error 53079
'
'        DataInicial.Text = sData
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownDataInicial_UpClick:
'
'    Select Case Err
'
'        Case 53079
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160756)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Inicializa_Datas()
''inicializa as datas inicial e final e coloca nos respectivos campos
'
'    Dim dtDataInicial As Date
'    Dim dtDataFinal As Date
'    Dim iMesAtual As Integer
'    Dim iAnoAtual As Integer
'
'    'coloca o mes e o ano correntes nas variaveis iMes e iAno
'    iMesAtual = Month(gdtDataAtual)
'    iAnoAtual = Year(gdtDataAtual)
'
'    'obter data inicial
'    If iMesAtual < 4 Then
'
'        dtDataInicial = CDate("01/" & CStr(iMesAtual + 9) & "/" & CStr(iAnoAtual - 1))
'
'    Else
'
'        dtDataInicial = CDate("01/" & CStr(iMesAtual - 3) & "/" & CStr(iAnoAtual))
'
'    End If
'
'    'obter data final
'    dtDataFinal = CDate("01/" & CStr(iMesAtual) & "/" & CStr(iAnoAtual)) - 1
'
'    Call DateParaMasked(DataInicial, dtDataInicial)
'    Call DateParaMasked(DataFinal, dtDataFinal)
'
'End Sub
'
'
''**** inicio do trecho a ser copiado *****
'Public Function Form_Load_Ocx() As Object
'
'    Parent.HelpContextID = IDH_GERACAO_ARQICMS
'    Set Form_Load_Ocx = Me
'    Caption = "Geração de Arquivo para ICMS"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "GeracaoArqICMS"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'   ' Parent.UnloadDoFilho
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''***** fim do trecho a ser copiado ******
'
'Function Trata_Parametros() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Trata_Parametros
'
'
'    Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160757)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label6, Source, X, Y)
'End Sub
'
'Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label7, Source, X, Y)
'End Sub
'
'Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label5, Source, X, Y)
'End Sub
'
'Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label9, Source, X, Y)
'End Sub
'
'Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label8, Source, X, Y)
'End Sub
'
'Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'
