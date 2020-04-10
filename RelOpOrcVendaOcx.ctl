VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpOrcVendaOcx 
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   KeyPreview      =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7155
   Begin VB.CheckBox ExibeValoresOrcamento 
      Caption         =   "Exibir os valores do orçamento"
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
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.Frame FrameCliente 
      Caption         =   "Clientes"
      Height          =   1380
      Left            =   285
      TabIndex        =   23
      Top             =   1687
      Width           =   4305
      Begin MSMask.MaskEdBox ClienteDe 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteAte 
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteDe 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   413
         Width           =   315
      End
      Begin VB.Label LabelClienteAte 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   345
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   893
         Width           =   360
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
      Height          =   600
      Left            =   5145
      Picture         =   "RelOpOrcVendaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   900
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4665
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOrcVendaOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpOrcVendaOcx.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOrcVendaOcx.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOrcVendaOcx.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOrcVendaOcx.ctx":0A96
      Left            =   1035
      List            =   "RelOpOrcVendaOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2670
   End
   Begin VB.Frame FrameOrcamento 
      Caption         =   "Orçamento"
      Height          =   675
      Left            =   285
      TabIndex        =   18
      Top             =   840
      Width           =   4305
      Begin MSMask.MaskEdBox OrcamentoDe 
         Height          =   300
         Left            =   765
         TabIndex        =   1
         Top             =   270
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OrcamentoAte 
         Height          =   300
         Left            =   2865
         TabIndex        =   2
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelOrcamentoAte 
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
         Height          =   195
         Left            =   2415
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   308
         Width           =   360
      End
      Begin VB.Label LabelOrcamentoDe 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   810
      Left            =   285
      TabIndex        =   13
      Top             =   3240
      Width           =   4305
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   660
         TabIndex        =   5
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   315
         Left            =   3780
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   2835
         TabIndex        =   6
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   345
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.Label LabelOpcao 
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
      Left            =   345
      TabIndex        =   21
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOrcVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'eventos dos browsers
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoOrcamento As AdmEvento
Attribute objEventoOrcamento.VB_VarHelpID = -1

'objetos dos relatorios
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'controle de browser
Dim giClienteInicial  As Integer
Dim giOrcamentoInicial As Integer

Public Sub Form_Load()

On Error GoTo Erro_Form_Load
    
    'instancia os objs dos browsers
    Set objEventoCliente = New AdmEvento
    Set objEventoOrcamento = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170443)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
'espera o objrelatorio e o obj relopcoes como parametro

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'se não foi passado o parametro, erro
    If Not (gobjRelatorio Is Nothing) Then gError 119878
    
    'instancia o obj global c/ o obj recebido
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 119879

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 119879
        
        Case 119878
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170444)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 119880

    'preenche o gobj c/ a opção de relatorio que está na tela
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 119881

    'grava o nome da opção
    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'grava a opção de relatorio
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 119882

    'testa p/ ver se foi gravado direito
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 119883
    
    'limpa a tela
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 119880
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 119881 To 119883

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170445)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'exclui a opção de relatorio que foi selecionada

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 119884

    'pergunta se deseja excluir o relatorio
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    'se sim
    If vbMsgRes = vbYes Then

        'chama a rotina que exclui a opção de rel.
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 119885

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click
            
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 119884
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 119885

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170446)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'limpa a tela
 
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa o rel
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 119886
            
    'limpa a combo e seta o foco
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
        
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 119886
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170447)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoExecutar_Click()
'executa o relatorio

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'preenche o relatorio c/ as opcoes da tela
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 119887
    
    'executa o relatorio
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 119887

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170448)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lCliente_I As Long
Dim lCliente_F As Long

On Error GoTo Erro_PreencherRelOp

    'formata os parametros da tela e verifica se estão corretos
    lErro = Formata_E_Critica_Parametros(lCliente_I, lCliente_F)
    If lErro <> SUCESSO Then gError 119888
    
    'limpa as opções
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 119889
   
    'coloca o orçamento
    lErro = objRelOpcoes.IncluirParametro("NORCINIC", OrcamentoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 119890
    
    lErro = objRelOpcoes.IncluirParametro("NORCFIM", OrcamentoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 119891
    
    'coloca a opção do cliente (oq está na tela e o cód.)
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 119892
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", CStr(lCliente_I))
    If lErro <> AD_BOOL_TRUE Then gError 119893
    
    '???
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", ClienteAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 119894
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", CStr(lCliente_F))
    If lErro <> AD_BOOL_TRUE Then gError 119895
    
    'adiciona a data
    If DataDe.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINI", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 119896

    If DataAte.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 119897
   
    'inclui o tratamento de verificação p/ exibir orçamentos
    lErro = objRelOpcoes.IncluirParametro("NEXIBIRORC", ExibeValoresOrcamento.Value)
    If lErro <> AD_BOOL_TRUE Then gError 119898
    
    'monta a expressão que vai ser usada no gerador de rel.
    lErro = Monta_Expressao_Selecao(objRelOpcoes, lCliente_I, lCliente_F)
    If lErro <> SUCESSO Then gError 119899
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr


    Select Case gErr

        Case 119888 To 119899
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170449)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    'carrega as opções
    lErro = objRelOpcoes.Carregar
    If lErro Then gError 119900
            
    'pega parâmetro orcamento Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NORCINIC", sParam)
    If lErro Then gError 119901
    
    OrcamentoDe.Text = sParam
    
    'pega parâmetro orcamento Final e exibe
    lErro = objRelOpcoes.ObterParametro("NORCFIM", sParam)
    If lErro Then gError 119902
    
    OrcamentoAte.Text = sParam
            
    'pega Cliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro Then gError 119903
    
    'se o cód. for <> de 0
    If sParam <> 0 Then
    
        ClienteDe.Text = sParam
        Call ClienteDe_Validate(bSGECancelDummy)
    
    End If
    
    'pega  Cliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro Then gError 119904
    
    If sParam <> 0 Then
    
        ClienteAte.Text = sParam
        Call ClienteAte_Validate(bSGECancelDummy)
                
    End If
                
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINI", sParam)
    If lErro <> SUCESSO Then gError 119905

    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then gError 119906

    Call DateParaMasked(DataAte, CDate(sParam))
                
    'pega a opção de exibir os orçamentos
    lErro = objRelOpcoes.ObterParametro("NEXIBIRORC", sParam)
    If lErro <> SUCESSO Then gError 119907

    ExibeValoresOrcamento.Value = sParam
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 119900 To 119907
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170450)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(lCliente_I As Long, lCliente_F As Long) As Long
'formata os parametros que estão na tela e verifica se estão corretos

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

       
    'orcamento inicial não pode ser maior que o orcamento final
    If Trim(OrcamentoDe.Text) <> "" And Trim(OrcamentoAte.Text) <> "" Then
    
         If StrParaLong(OrcamentoDe.Text) > StrParaLong(OrcamentoAte.Text) Then gError 119908
         
    End If
    
    'o cód. do cliente não pode ser maior do que o final
    If ClienteDe.Text <> "" Then
        lCliente_I = LCodigo_Extrai(ClienteDe.Text)
    Else
        lCliente_I = 0
    End If
    
    If ClienteAte.Text <> "" Then
        lCliente_F = LCodigo_Extrai(ClienteAte.Text)
    Else
        lCliente_F = 0
    End If
            
    If lCliente_I <> 0 And lCliente_F <> 0 Then
        
        If lCliente_I > lCliente_F Then gError 119909
        
    End If
    
    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 119910
    
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
            
        Case 119908
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOINICIAL_MAIOR_ORCAMENTOFINAL", gErr)
        
        Case 119909
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEINICIAL_MAIOR_CLIENTEFINAL", gErr)
        
        Case 119910
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
               
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170451)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, lCliente_I As Long, lCliente_F As Long) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    'monta a expressão do orçamento inicial e final
    If Trim(OrcamentoDe.Text) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Orcamento >= " & Forprint_ConvLong(StrParaLong(OrcamentoDe.Text))
        
    End If

    If Trim(OrcamentoAte.Text) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Orcamento <= " & Forprint_ConvLong(StrParaLong(OrcamentoAte.Text))

    End If
   
   'monta a expressão do cliente inicial e final
   If lCliente_I <> 0 Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvLong(lCliente_I)

   End If

   If lCliente_F <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(lCliente_F)

    End If
    
    'monta a expressaõ da data inicial e final
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(StrParaDate(DataDe.Text))

    End If
    
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If
   
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170452)

    End Select

    Exit Function

End Function

Private Sub ClienteDe_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    giClienteInicial = 1

    'se estiver preenchido
    If Len(Trim(ClienteDe.Text)) <> 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objCliente, 0)
        If lErro <> SUCESSO Then gError 119911

    End If
    
    Exit Sub

Erro_ClienteDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 119911

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170453)

    End Select

    Exit Sub

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    giClienteInicial = 1

    'se estiver preenchido
    If Len(Trim(ClienteAte.Text)) <> 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objCliente, 0)
        If lErro <> SUCESSO Then gError 119912

    End If
    
    Exit Sub

Erro_ClienteAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 119912

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170454)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()
'chama a lista de clientes cadastrados

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 0
    
    'se estiver preenchido
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelClienteDe_Click()
'chama a lista de clientes cadastrados

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    giClienteInicial = 1

    'se estiver preenchido
    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)
'coloca o cliente selecionado na tela

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente de ou ate
    If giClienteInicial = 1 Then
        ClienteDe.Text = CStr(objCliente.lCodigo)
        Call ClienteDe_Validate(bSGECancelDummy)
    Else
        ClienteAte.Text = CStr(objCliente.lCodigo)
        Call ClienteAte_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)
    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'valida a data

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'se estiver prenchida
    If Len(DataAte.ClipText) > 0 Then
        
        'verifica se é valida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 119913

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 119913

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170455)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataDe)
End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'valida a data

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'se estiver preenchida
    If Len(DataDe.ClipText) > 0 Then

        'verifica se e valida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 119914

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 119914

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170456)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOrcamento_evSelecao(obj1 As Object)
'traz o orcamento de acordo c/ oq foi selecionado no browser

Dim objOrcamento As ClassOrcamentoVenda

    Set objOrcamento = obj1
    
    'Preenche campo orcamento
    If giOrcamentoInicial = 1 Then
        OrcamentoDe.Text = CStr(objOrcamento.lCodigo)
        Call OrcamentoDe_Validate(bSGECancelDummy)
    Else
        OrcamentoAte.Text = CStr(objOrcamento.lCodigo)
        Call OrcamentoAte_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Private Sub OrcamentoDe_Validate(Cancel As Boolean)
'verifica se o orçamento é valido

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
        
On Error GoTo Erro_OrcamentoDe_Validate
        
    giOrcamentoInicial = 1

    'se estiver preenchido
    If Len(Trim(OrcamentoDe.Text)) > 0 Then
        
        'verifica se foi preenchido corretamante
        lErro = Long_Critica(OrcamentoDe.Text)
        If lErro <> SUCESSO Then gError 119915
    
        'preenche o cód. e a filial correspondente
        objOrcamentoVenda.lCodigo = StrParaLong(OrcamentoDe.Text)
        objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
       
        'verifica se está cadastrado no BD
        lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101232 Then gError 119916
            
        'não está cadastrado
        If lErro = 101232 Then gError 119917
        
    End If

    Exit Sub

Erro_OrcamentoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 119915, 119916
        
        Case 119917
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, OrcamentoDe.Text, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170457)

    End Select

    Exit Sub

End Sub

Private Sub OrcamentoAte_Validate(Cancel As Boolean)
'verifica se o orçamento é valido

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
        
On Error GoTo Erro_OrcamentoAte_Validate
        
    giOrcamentoInicial = 0

    'se estiver preenchido
    If Len(Trim(OrcamentoAte.Text)) > 0 Then
        
        'verifica se foi preenchido corretamente
        lErro = Long_Critica(OrcamentoAte.Text)
        If lErro <> SUCESSO Then gError 119918
    
        'preenche o cód. e a filial correspondente
        objOrcamentoVenda.lCodigo = StrParaLong(OrcamentoAte.Text)
        objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
       
        'procura no bd o orçamento a partir do cód. passado e da filial
        lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101232 Then gError 119919
            
        'não está cadastrado
        If lErro = 101232 Then gError 119920
        
    End If

    Exit Sub

Erro_OrcamentoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 119918, 119919

        Case 119920
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, OrcamentoAte.Text, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170458)

    End Select

    Exit Sub

End Sub

Private Sub OrcamentoDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(OrcamentoDe)
End Sub

Private Sub OrcamentoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(OrcamentoAte)
End Sub

Private Sub LabelOrcamentoDe_Click()
'chama o browser de orcamentos cadastrados

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_LabelOrcamentoDe_Click

    giOrcamentoInicial = 1
        
    'se estiver preenchido
    If Len(Trim(OrcamentoDe.Text)) > 0 Then
    
        'carrega o cód do orcamento
        objOrcamentoVenda.lCodigo = StrParaLong(OrcamentoDe.Text)

    End If

    'chama a tela de orçamentos
    Call Chama_Tela("OrcamentoVendaLista", colSelecao, objOrcamentoVenda, objEventoOrcamento)
    
    Exit Sub

Erro_LabelOrcamentoDe_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170459)

    End Select

    Exit Sub

End Sub

Private Sub LabelOrcamentoAte_Click()
'chama o browser de orcamentos cadastrados

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_LabelOrcamentoAte_Click

    giOrcamentoInicial = 0

    'se estiver preenchido
    If Len(Trim(OrcamentoAte.Text)) > 0 Then
    
        'carrega o cód do orcamento
        objOrcamentoVenda.lCodigo = StrParaLong(OrcamentoAte.Text)

    End If

    'chama a tela de orçamentos
    Call Chama_Tela("OrcamentoVendaLista", colSelecao, objOrcamentoVenda, objEventoOrcamento)
    
    Exit Sub

Erro_LabelOrcamentoAte_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170460)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()
'diminui a data

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    'diminui a data em 1 dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 119921

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 119921
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170461)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    'aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 119922

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 119922
            DataDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170462)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()
'diminui a data

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    'diminui a data em 1 dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 119923

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 119923
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170463)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    'aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 119924

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 119924
            DataAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170464)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        ElseIf Me.ActiveControl Is OrcamentoDe Then
            Call LabelOrcamentoDe_Click
        ElseIf Me.ActiveControl Is OrcamentoAte Then
            Call LabelOrcamentoAte_Click
        End If
    
    End If

End Sub

Public Sub Form_Unload(Cancel As Integer)
'libera os obj do relatorio e os do browsers
    
    Set objEventoCliente = Nothing
    Set objEventoOrcamento = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Relação de Orçamentos de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOrcVenda"
    
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

Private Sub LabelOrcamentoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrcamentoDe, Source, X, Y)
End Sub

Private Sub LabelOrcamentoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrcamentoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelOrcamentoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrcamentoAte, Source, X, Y)
End Sub

Private Sub LabelOrcamentoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrcamentoAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataDe, Source, X, Y)
End Sub

Private Sub LabelDataDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDataAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataAte, Source, X, Y)
End Sub

Private Sub LabelDataAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelOpcao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpcao, Source, X, Y)
End Sub

Private Sub LabelOpcao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpcao, Button, Shift, X, Y)
End Sub
