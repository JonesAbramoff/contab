VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpDiarioAuxOcx 
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   ScaleHeight     =   3810
   ScaleWidth      =   6915
   Begin VB.TextBox Diario 
      Height          =   285
      Left            =   5760
      TabIndex        =   13
      Top             =   3285
      Width           =   510
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprime"
      Height          =   1335
      Left            =   300
      TabIndex        =   9
      Top             =   2370
      Width           =   2940
      Begin VB.OptionButton OpcaoLivro 
         Caption         =   "Somente o Livro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   540
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton OpcaoTermos 
         Caption         =   "Somente os Termos"
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
         Left            =   540
         TabIndex        =   11
         Top             =   585
         Width           =   2010
      End
      Begin VB.OptionButton OpcaoAmbos 
         Caption         =   "Ambos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   10
         Top             =   900
         Width           =   1905
      End
   End
   Begin VB.TextBox QtdeFolhas 
      Height          =   285
      Left            =   5775
      TabIndex        =   8
      Top             =   2775
      Width           =   510
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
      Height          =   600
      Left            =   4200
      Picture         =   "RelOpDiarioAuxOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox PrimeiraFolha 
      Height          =   285
      Left            =   5790
      TabIndex        =   6
      Top             =   2250
      Width           =   510
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDiarioAuxOcx.ctx":0102
      Left            =   1335
      List            =   "RelOpDiarioAuxOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2325
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDiarioAuxOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDiarioAuxOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDiarioAuxOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDiarioAuxOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2490
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataInicial 
      Height          =   285
      Left            =   1335
      TabIndex        =   15
      Top             =   1095
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   2490
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1830
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataFinal 
      Height          =   285
      Left            =   1335
      TabIndex        =   17
      Top             =   1845
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Número do Diário:"
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
      Left            =   4185
      TabIndex        =   23
      Top             =   3330
      Width           =   1545
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade de  Folhas:"
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
      Left            =   3735
      TabIndex        =   22
      Top             =   2805
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Número da Primeira Folha:"
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
      Left            =   3480
      TabIndex        =   21
      Top             =   2280
      Width           =   2250
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   660
      TabIndex        =   20
      Top             =   405
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
      Left            =   330
      TabIndex        =   19
      Top             =   1875
      Width           =   945
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   18
      Top             =   1125
      Width           =   1050
   End
End
Attribute VB_Name = "RelOpDiarioAuxOcx"
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

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 65364
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 65365

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 65365
        
        Case 65364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168271)

    End Select

    Exit Function

End Function

Function Critica_Datas_RelOpDiario(iExercicio As Integer) As Long
'as datas ( inicial e final ) não podem ser vazias
'a data inicial não pode ser maior que a data final
'as datas devem estar no mesmo exercício

Dim lErro As Long
Dim objPer_I As New ClassPeriodo, objPer_F As New ClassPeriodo

On Error GoTo Erro_Critica_Datas_RelOpDiario

    'data inicial não pode ser vazia
    If Len(DataInicial.ClipText) = 0 Then Error 65366

    'data final não pode ser vazia
    If Len(DataFinal.ClipText) = 0 Then Error 65368

    'data inicial não pode ser maior que a data final
    If CDate(DataInicial.Text) > CDate(DataFinal.Text) Then Error 65370

    'pega o período da Data Inicial
    lErro = CF("Periodo_Le", DataInicial.Text, objPer_I)
    If lErro <> SUCESSO Then Error 65367

    'pega o período da Data Final
    lErro = CF("Periodo_Le", DataFinal.Text, objPer_F)
    If lErro <> SUCESSO Then Error 65369

    'Data Inicial e Final devem estar num mesmo exercício
    If objPer_I.iExercicio <> objPer_F.iExercicio Then Error 65371

    iExercicio = objPer_I.iExercicio
    
    Critica_Datas_RelOpDiario = SUCESSO

    Exit Function

Erro_Critica_Datas_RelOpDiario:

    Critica_Datas_RelOpDiario = Err

    Select Case Err

        Case 65366, 65367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
            DataInicial.SetFocus

        Case 65368, 65369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", Err)
            DataFinal.SetFocus

        Case 65370
           lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case 65371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_COM_EXERCICIOS_DIFERENTES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168272)

    End Select
    
    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iExercicio As Integer

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Datas_RelOpDiario(iExercicio)
    If lErro <> SUCESSO Then Error 65372

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 65373

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(iExercicio))
    If lErro <> AD_BOOL_TRUE Then Error 65374

    lErro = objRelOpcoes.IncluirParametro("NPAGRELINI", PrimeiraFolha.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65375

    lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65376

    lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65377

    lErro = objRelOpcoes.IncluirParametro("NNUMDIARIO", Diario.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65377

    lErro = objRelOpcoes.IncluirParametro("NQTDFOLHAS", QtdeFolhas.Text)
    If lErro <> AD_BOOL_TRUE Then Error 65377
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 65378

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 65372 To 65378

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168273)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arqquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 65379

    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 65380

    DataInicial.PromptInclude = False
    DataInicial.Text = sParam
    DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 65381

    DataFinal.PromptInclude = False
    DataFinal.Text = sParam
    DataFinal.PromptInclude = True

    'pega primeira folha e exibe
    lErro = objRelOpcoes.ObterParametro("NPAGRELINI", sParam)
    If lErro <> SUCESSO Then Error 65382

    PrimeiraFolha.Text = sParam

    'pega número diário e exibe
    lErro = objRelOpcoes.ObterParametro("NNUMDIARIO", sParam)
    If lErro <> SUCESSO Then Error 65382

    Diario.Text = sParam
    
    'pega número diário e exibe
    lErro = objRelOpcoes.ObterParametro("NQTDFOLHAS", sParam)
    If lErro <> SUCESSO Then Error 65382

    QtdeFolhas.Text = sParam

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 65379 To 65382

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168274)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se não existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 65383

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 65384

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 65385

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 65383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 65384, 65385

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168275)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long, objTermo As New AdmRelatorio
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 65386

    If OpcaoLivro.Value = True Then
    
        Call gobjRelatorio.Executar_Prossegue2(Me)

    End If
    
    If OpcaoTermos.Value = True Then
    
        'Executa o relatório
        Call objTermo.ExecutarDireto("Termo de Abertura", "", 1, "termoab", "DINIC", DataInicial.Text, "DFIM", DataFinal.Text, "NNUMDIARIO", Diario.Text, "NQTDFOLHAS", QtdeFolhas.Text, "NPAGRELINI", PrimeiraFolha.Text)
    
        Set objTermo = Nothing
    
        'Executa o relatório
        Call objTermo.ExecutarDireto("Termo de Encerramento", "", 1, "termoenc", "DINIC", DataInicial.Text, "DFIM", DataFinal.Text, "NNUMDIARIO", Diario.Text, "NQTDFOLHAS", QtdeFolhas.Text, "NPAGRELINI", PrimeiraFolha.Text)
    
    End If
    
    If OpcaoAmbos.Value = True Then
    
        'Executa o relatório
        Call objTermo.ExecutarDireto("Termo de Abertura", "", 1, "termoab", "DINIC", DataInicial.Text, "DFIM", DataFinal.Text, "NNUMDIARIO", Diario.Text, "NQTDFOLHAS", QtdeFolhas.Text, "NPAGRELINI", PrimeiraFolha.Text)
    
        Set objTermo = Nothing
    
        Call gobjRelatorio.Executar_Prossegue2(Me)
    
        'Executa o relatório
        Call objTermo.ExecutarDireto("Termo de Encerramento", "", 1, "termoenc", "DINIC", DataInicial.Text, "DFIM", DataFinal.Text, "NNUMDIARIO", Diario.Text, "NQTDFOLHAS", QtdeFolhas.Text, "NPAGRELINI", PrimeiraFolha.Text)
    
    End If

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 65386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168276)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 65387

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 65388

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 65389

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 65390
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 65387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 65388

        Case 65389, 65390
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168277)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 65391
    
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 65391
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168278)

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
        If lErro <> SUCESSO Then Error 65392

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 65392

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168279)

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
        If lErro <> SUCESSO Then Error 65393

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 65393

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168280)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168281)

    End Select

    Unload Me

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 65394

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 65394
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168282)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 65395

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 65395
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168283)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 65396

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 65396
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168284)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 65397

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 65397
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168285)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    Select Case giFilialEmpresa
        Case EMPRESA_TODA
            If giContabGerencial <> 0 Then
                If sExpressao <> "" Then sExpressao = sExpressao & " E "
                sExpressao = sExpressao & "FilialEmpresaLcto < " & Forprint_ConvInt(Abs(giFilialAuxiliar))
            End If
        
        Case Abs(giFilialAuxiliar)
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "FilialEmpresaLcto > " & Forprint_ConvInt(Abs(giFilialAuxiliar))
        
        Case Else
            If sExpressao <> "" Then sExpressao = sExpressao & " E "
            sExpressao = sExpressao & "FilialEmpresaLcto = " & Forprint_ConvInt(giFilialEmpresa)
    End Select
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168286)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DIARIO
    Set Form_Load_Ocx = Me
    Caption = "Diário Auxiliar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDiarioAux"
    
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub


