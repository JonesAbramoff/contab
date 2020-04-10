VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpMovFinDetOcx 
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   6150
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3825
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMovFinDetOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpMovFinDetOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpMovFinDetOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpMovFinDetOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4110
      Picture         =   "RelOpMovFinDetOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMovFinDetOcx.ctx":0A96
      Left            =   855
      List            =   "RelOpMovFinDetOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2850
   End
   Begin VB.ComboBox CodContaCorrente 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   885
      Width           =   2115
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2730
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1477
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   1477
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   2730
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2100
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   315
      Left            =   1605
      TabIndex        =   3
      Top             =   2100
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      Caption         =   "Opção:"
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
      Left            =   165
      TabIndex        =   15
      Top             =   270
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
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
      TabIndex        =   14
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial:"
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
      Left            =   465
      TabIndex        =   13
      Top             =   1537
      Width           =   1050
   End
   Begin VB.Label LabelCtaCorrente 
      AutoSize        =   -1  'True
      Caption         =   "Conta Corrente:"
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
      Left            =   165
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   12
      Top             =   915
      Width           =   1350
   End
End
Attribute VB_Name = "RelOpMovFinDetOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Function Critica_Datas_RelOpMovFin() As Long
'Faz a crítica da data inicial e da data final

Dim lErro As Long

On Error GoTo Erro_Critica_Datas_RelOpMovFin

    'data inicial não pode ser vazia
    If Len(DataInicial.ClipText) = 0 Then Error 23232

    'data final não pode ser vazia
    If Len(DataFinal.ClipText) = 0 Then Error 23233

    'data inicial não pode ser maior que a data final
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 23234

    Critica_Datas_RelOpMovFin = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpMovFin:

    Critica_Datas_RelOpMovFin = Err

    Select Case Err

        Case 23232
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err, Error$)
            DataInicial.SetFocus

        Case 23233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err, Error$)
            DataFinal.SetFocus

        Case 23234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170034)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 23235

    'pega cta corrente e preenche na tela
    lErro = objRelOpcoes.ObterParametro("TCTACORRENT", sParam)
    If lErro <> SUCESSO Then Error 23236
    CodContaCorrente.Text = sParam
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 23238

    DataInicial.PromptInclude = False
    DataInicial.Text = sParam
    DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 23274

    DataFinal.PromptInclude = False
    DataFinal.Text = sParam
    DataFinal.PromptInclude = True

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23235, 23236, 23238, 23274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170035)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpMovFin
    If lErro <> SUCESSO Then Error 23239

    'Verificar se Cta Corrente foi informada
    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 23272
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23240
    
    lErro = objRelOpcoes.IncluirParametro("TCTACORRENT", CodContaCorrente.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23273
    
    lErro = objRelOpcoes.IncluirParametro("NCODCONTA", CStr(Codigo_Extrai(CodContaCorrente.Text)))
    If lErro <> AD_BOOL_TRUE Then Error 32005
    
    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23241

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23242

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23239, 23240, 23241, 23242
        
        Case 23272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err, Error$)
            CodContaCorrente.SetFocus

        Case 23273, 32005
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170036)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29905
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche combo com as opções de relatório
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 23253

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 23253
        
        Case 29905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170037)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 23243

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPMOVFIN")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 23244

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela
        
        'Iniciar data final com data corrente
        DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 23243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 23244

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170038)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23245

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23245

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170039)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 23246

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23247

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 23248

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 57757
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 23247, 23248, 57757

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170040)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    CodContaCorrente.Text = ""
    
    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    
    'Iniciar data final com data corrente
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")

End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim iCodigo As Integer

On Error GoTo Erro_CodContaCorrente_Validate

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If CodContaCorrente.Text = CodContaCorrente.List(CodContaCorrente.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 23268

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objContaCorrenteInt.iCodigo = iCodigo
        lErro = CF("ContaCorrenteInt_Le",objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 23269

        'Não encontrou a Conta Corrente no BD
        If lErro = 11807 Then Error 23270

        'Encontrou a Conta Corrente no BD, coloca no Text da Combo
        CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 23271

    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 23268, 23269

        Case 23270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrenteInt.iCodigo)
            
        Case 23271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, CodContaCorrente.Text)
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170041)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 23251

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 23251

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170042)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 23252

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 23252

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170043)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim colCodigoDescricao As New AdmCollCodigoNome
Dim lErro As Long, iIndice As Integer
Dim objCodigoDescricao As AdmlCodigoNome

On Error GoTo Erro_OpcoesRel_Form_Load
    
    'Inicia data final com data corrente
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")

    'Lê as contas correntes  com codigo e o nome reduzido existentes no BD e carrega na ComboBox
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then Error 23266
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 23266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170044)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23254

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 23254
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170045)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23255

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 23255
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170046)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 23256

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 23256
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170047)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 23257

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 23257
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170048)

    End Select

    Exit Sub

End Sub

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente

    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed",colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 23267

    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        CodContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo
        
    Next

    'Seleciona uma das Contas
    CodContaCorrente.Text = CodContaCorrente.List(0)

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 23267

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170049)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_MOVFIN_DET
    Set Form_Load_Ocx = Me
    Caption = "Movimentação Financeira Detalhada"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMovFinDet"
    
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

Public Sub Unload(objme As Object)
    
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



Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaCorrente, Source, X, Y)
End Sub

Private Sub LabelCtaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaCorrente, Button, Shift, X, Y)
End Sub

