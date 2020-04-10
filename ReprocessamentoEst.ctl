VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ReprocessamentoEst 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   1425
   ScaleMode       =   0  'User
   ScaleWidth      =   7400
   Begin VB.CheckBox ApenasSaldoTerc 
      Caption         =   "Apenas ajustar saldos por terceiros"
      Height          =   270
      Left            =   135
      TabIndex        =   28
      Top             =   1020
      Width           =   6975
   End
   Begin VB.CommandButton BotaoSuporte 
      Caption         =   "Exibir opções para suporte"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   3255
   End
   Begin VB.Frame FrameSuporte 
      Caption         =   "Opções para suporte"
      Height          =   6465
      Left            =   75
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CheckBox IgnoraHora 
         Caption         =   "Ignorar a hora dos movimentos e reprocessar primeiro as entradas"
         Height          =   315
         Left            =   345
         TabIndex        =   27
         Top             =   285
         Value           =   1  'Checked
         Width           =   6015
      End
      Begin VB.TextBox FilEmpGrupo 
         Height          =   285
         Left            =   930
         TabIndex        =   23
         Top             =   4050
         Width           =   1260
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "ALERTA !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1215
         Left            =   105
         TabIndex        =   19
         Top             =   5100
         Width           =   6855
         Begin VB.Label Label2 
            Caption         =   $"ReprocessamentoEst.ctx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   6495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   6855
         Begin MSMask.MaskEdBox DataFim 
            Height          =   300
            Left            =   1320
            TabIndex        =   13
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoCodigo 
            Height          =   315
            Left            =   1320
            TabIndex        =   14
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataFim 
            Height          =   300
            Left            =   2400
            TabIndex        =   15
            Top             =   960
            Width           =   240
            _ExtentX        =   370
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelDataFim 
            Caption         =   "Data Final:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label LabelProduto 
            Caption         =   "Produto:"
            Height          =   255
            Left            =   240
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label LabelProdutoDescricao 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            TabIndex        =   16
            Top             =   300
            Width           =   4215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opções para depuração"
         Height          =   1710
         Left            =   120
         TabIndex        =   8
         Top             =   2205
         Width           =   6855
         Begin VB.CommandButton BotaoTeste 
            Caption         =   "Executar Testes"
            Height          =   315
            Left            =   4575
            TabIndex        =   26
            Top             =   1335
            Width           =   2190
         End
         Begin VB.CheckBox AcertaEstProd 
            Caption         =   "Acerta EstoqueProduto QuantDispNossa"
            Height          =   225
            Left            =   240
            TabIndex        =   22
            Top             =   1455
            Width           =   5865
         End
         Begin VB.CheckBox ReprocTestaInt 
            Caption         =   "Testar integridade"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   375
            Width           =   2295
         End
         Begin VB.CheckBox LogReproc 
            Caption         =   "Fazer log passo a passo durante o reprocessamento"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   4935
         End
         Begin VB.CheckBox PulaFaseDesfaz 
            Caption         =   "Zerar saldos e pular a fase Desfaz"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   4935
         End
      End
      Begin VB.Label Label3 
         Caption         =   $"ReprocessamentoEst.ctx":00B1
         Height          =   1095
         Left            =   2520
         TabIndex        =   25
         Top             =   3915
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Filiais:"
         Height          =   255
         Left            =   300
         TabIndex        =   24
         Top             =   4065
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3630
      ScaleHeight     =   735
      ScaleWidth      =   3345
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   3405
      Begin VB.CommandButton BotaoLimpar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2160
         Picture         =   "ReprocessamentoEst.ctx":016F
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton BotaoReprocessar 
         Height          =   630
         Left            =   120
         Picture         =   "ReprocessamentoEst.ctx":06A1
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   1995
      End
      Begin VB.CommandButton BotaoFechar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2760
         Picture         =   "ReprocessamentoEst.ctx":3563
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   495
      End
   End
   Begin MSMask.MaskEdBox DataInicio 
      Height          =   300
      Left            =   2100
      TabIndex        =   0
      Top             =   615
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataInicio 
      Height          =   300
      Left            =   3165
      TabIndex        =   1
      Top             =   600
      Width           =   240
      _ExtentX        =   370
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label LabelDataInicio 
      AutoSize        =   -1  'True
      Caption         =   "Iniciar o Reprocessamento em:"
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   1920
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ReprocessamentoEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim dtDataPrimeiroMovEstoque As Date
Dim dtDataMinimaReprocessamento As Date
Dim dtDataInicioUltReproc As Date

'Evento que ocorre ao clicar no label Produto
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Variáveis de controle de alterações na tela
Dim iProdutoCodigoAlterado As Integer
Dim iAlterado As Integer

Private Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load
    
    Set objEventoProduto = New AdmEvento
    
    Parent.Height = 1825
    'Parent.Width = 7400
    UserControl.Height = 1825
    'UserControl.Width = 7400
    FrameSuporte.Visible = False

    'Inicializa a máscara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoCodigo)
    If lErro <> SUCESSO Then gError 90569
    
    'Obtém a data mínima a partir da qual o BD precisa ser reprocessado
    'Ou seja, nenhuma data posterior a essa será aceita como data de início do reprocessamento
    lErro = DataInicio_Reprocessamento_Obtem()
    If lErro <> SUCESSO Then gError 79958

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 79958, 90569
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173738)

    End Select
    
    Exit Sub
    
End Sub

Private Function DataInicio_Reprocessamento_Obtem() As Long

Dim lErro As Long
Dim objMATConfig As New ClassMATConfig

On Error GoTo Erro_DataInicio_Reprocessamento_Obtem
    
    'Guarda em objMATConfig os parâmetros necessários para se obter a data default para início do reprocessamento
    objMATConfig.iFilialEmpresa = giFilialEmpresa
    objMATConfig.sCodigo = DATA_REPROCESSAMENTO
    
    lErro = CF("MATConfig_Le", objMATConfig)
    If lErro <> SUCESSO And lErro <> 89118 Then gError 79956
    
    'Início Daniel em 15/04/2002
    'Se encontrou alguma data anterior => formata e guarda
    If lErro = AD_SQL_SUCESSO Then
        
        'Guarda em uma variável global à tela a data encontrada
        dtDataMinimaReprocessamento = StrParaDate(objMATConfig.sConteudo)
    
        DataInicio.PromptInclude = False
        DataInicio.Text = Format(dtDataMinimaReprocessamento, "dd/mm/yy")
        DataInicio.PromptInclude = True
        
    Else
    
        DataInicio.PromptInclude = False
        DataInicio.Text = ""
        DataInicio.PromptInclude = True
        
        'Guarda em uma variável global à tela a data encontrada
        dtDataMinimaReprocessamento = DATA_NULA
    
    End If
    'Fim Daniel em 15/04/2002
    
    DataInicio_Reprocessamento_Obtem = SUCESSO
    
    Exit Function

Erro_DataInicio_Reprocessamento_Obtem:

    DataInicio_Reprocessamento_Obtem = gErr
    
    Select Case gErr
        
        Case 79956
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173739)
            
    End Select
    
    Exit Function
    
End Function

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub AcertaEstProd_Click()

Dim lErro As Long

On Error GoTo Erro_AcertaEstProd_Click

    'Se estiver marcada a opção para executar o teste de integridade
    If AcertaEstProd.Value = vbChecked Then
    
        'Altera a caption do LabelDataInicio
        LabelDataInicio.Caption = "Iniciar o acerto do EstoqueProduto total (independe de data):"
        
        'Desmarca a opção de geração de lock e a deixa desabilitada
        LogReproc.Value = False
        LogReproc.Enabled = False
        
        'Desmarca a opção de pular a fase desfaz e a deixa desabilitada
        PulaFaseDesfaz.Value = False
        PulaFaseDesfaz.Enabled = False

    'Senão
    Else
    
        'Altera a caption do LabelDataInicio
        LabelDataInicio.Caption = "Iniciar o Reprocessamento em:"

        'Habilita a opção de geração de lock
        LogReproc.Enabled = True
        
        'Habilita a opção de pular a fase desfaz
        PulaFaseDesfaz.Enabled = True
        
    End If
    
    Exit Sub
    
Erro_AcertaEstProd_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173740)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub ApenasSaldoTerc_Click()
    If ApenasSaldoTerc.Value = vbChecked Then
        BotaoSuporte.Visible = False
        Parent.Height = 1825
        'Parent.Width = 7400
        UserControl.Height = 1825
        'UserControl.Width = 7400
        FrameSuporte.Visible = False
    Else
        BotaoSuporte.Visible = True
    End If
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Reprocessamento
    
End Sub

Private Sub BotaoSuporte_Click()

    If FrameSuporte.Visible = False Then
    
        Parent.Height = 8100
        'Parent.Width = 7400
        UserControl.Height = 8100
        'UserControl.Width = 7400
        FrameSuporte.Visible = True
    
    Else
    
        Parent.Height = 1825
        'Parent.Width = 7400
        UserControl.Height = 1825
        'UserControl.Width = 7400
        FrameSuporte.Visible = False
    
    End If
    
End Sub

Private Sub BotaoTeste_Click()
    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    Call CF("Estoque_TestaIntegridade", giFilialEmpresa, True, True, True)
    GL_objMDIForm.MousePointer = vbDefault
End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate

    If Len(DataInicio.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(DataInicio.Text)
    If lErro <> SUCESSO Then gError 79957

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True
    
    Select Case gErr

        Case 79957
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173741)

    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sSelecao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_LabelProduto_Click

    sSelecao = "ControleEstoque <> ? AND FilialEmpresa=?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    colSelecao.Add giFilialEmpresa
    
    'Verifica se Produto está preenchido
    If Len(Trim(ProdutoCodigo.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", ProdutoCodigo.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 90574
        
        objProduto.sCodigo = sProdutoFormatado

    End If
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 90574

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173742)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoCodigo_Change()
    
    iProdutoCodigoAlterado = 1
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ProdutoCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoCodigo_Validate
   
    'função que preenche a tela e trata o produto no caso de ser ou nao inventariado
    lErro = Trata_Produto()
    If lErro <> SUCESSO Then gError 90568
    
    Exit Sub

Erro_ProdutoCodigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 90568

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173743)

    End Select

    Exit Sub

End Sub

Private Sub PulaFaseDesfaz_Click()

Dim lErro As Long
Dim objItemMovEstoque As New ClassItemMovEstoque

On Error GoTo Erro_PulaFaseDesfaz_Click

    'Se for para pular a fase desfaz
    'É preciso verificar qual a primeira data de movimento de estoque
    'e essa data deverá ser a data de início do reprocessamento,
    'pois como as tabelas de saldos serão zeradas para todos os meses
    'o refaz também deverá ser feito desde o início
    'A data final também será desabilitada para que o reprocessamento seja completo
    If PulaFaseDesfaz.Value = vbChecked Then
    
        'Desabilita a Data Final
        DataFim.Enabled = False
        LabelDataFim.Enabled = False
        UpDownDataFim.Enabled = False
    
        'Se ainda não foi lida a primeira data de movimento de estoque
        If dtDataPrimeiroMovEstoque = 0 Then
        
            'Lê a data do primeiro movimento de estoque
            lErro = CF("MovimentoEstoque_Le_Primeira_Data", giFilialEmpresa, objItemMovEstoque)
            If lErro <> SUCESSO And lErro <> 90708 Then gError 90708
        
            'Guarda a data em uma variável global à tela
            dtDataPrimeiroMovEstoque = objItemMovEstoque.dtData

        End If
        
        'Exibe a data lida como data para início do reprocessamento
        DataInicio.PromptInclude = False
        DataInicio.Text = Format(dtDataPrimeiroMovEstoque, "dd/mm/yy")
        DataInicio.PromptInclude = True
        
    'Senão
    'Exibe no campo data início a data mínima para início do reprocessamento
    'Habilita a data final
    Else
    
        If dtDataMinimaReprocessamento <> DATA_NULA Then
            DataInicio.PromptInclude = False
            DataInicio.Text = Format(dtDataMinimaReprocessamento, "dd/mm/yy")
            DataInicio.PromptInclude = True
        Else
            DataInicio.PromptInclude = False
            DataInicio.Text = ""
            DataInicio.PromptInclude = True
        End If
        
        'Habilita a Data Final
        DataFim.Enabled = True
        LabelDataFim.Enabled = True
        UpDownDataFim.Enabled = True
    
    End If
    
    Exit Sub

Erro_PulaFaseDesfaz_Click:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173744)
            
    End Select
    
    Exit Sub
        
End Sub

Private Sub ReprocTestaInt_Click()
    
Dim lErro As Long

On Error GoTo Erro_ReprocTestaInt_Click

    'Se estiver marcada a opção para executar o teste de integridade
    If ReprocTestaInt.Value = vbChecked Then
    
        'Altera a caption do LabelDataInicio
        LabelDataInicio.Caption = "Iniciar o teste de integridade em:"
        
        'Desmarca a opção de geração de lock e a deixa desabilitada
        LogReproc.Value = False
        LogReproc.Enabled = False
        
        'Desmarca a opção de pular a fase desfaz e a deixa desabilitada
        PulaFaseDesfaz.Value = False
        PulaFaseDesfaz.Enabled = False

        AcertaEstProd.Value = False
        AcertaEstProd.Enabled = False
        'Lê no BD a data em que se iniciou o último reprocessamento feito
        'E seta a data inicial para o teste de integridade igual à lida
        lErro = DataInicio_UltimoReprocessamento_Obtem()
        If lErro <> SUCESSO Then gError 90586
        
    'Senão
    Else
    
        'Altera a caption do LabelDataInicio
        LabelDataInicio.Caption = "Iniciar o Reprocessamento em:"

        'Habilita a opção de geração de lock
        LogReproc.Enabled = True
        
        'Habilita a opção de pular a fase desfaz
        PulaFaseDesfaz.Enabled = True
        
        AcertaEstProd.Enabled = True
        
        'Obtém a data mínima a partir da qual o BD precisa ser reprocessado
        'Ou seja, nenhuma data posterior a essa será aceita como data de início do reprocessamento
        lErro = DataInicio_Reprocessamento_Obtem()
        If lErro <> SUCESSO Then gError 90587
        
    End If
    
    Exit Sub
    
Erro_ReprocTestaInt_Click:

    Select Case gErr
    
        Case 90586, 90587
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173745)
    
    End Select
    
    Exit Sub
    
    
End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_DownClick

    If Len(DataInicio.ClipText) = 0 Then Exit Sub
    
    lErro = Data_Up_Down_Click(DataInicio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 79959

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 79959
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173746)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataInicio_UpClick

    If Len(DataInicio.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataInicio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 79960
    
    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 79960

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173747)

    End Select

    Exit Sub

End Sub

Private Sub BotaoReprocessar_Click()

Dim lErro As Long
Dim sNomeArqParam As String
Dim objReprocessamentoEST As New ClassReprocessamentoEST

On Error GoTo Erro_BotaoReprocessar_Click
    
    lErro = Reprocessamento_Critica()
    If lErro <> SUCESSO Then gError 90602
    
    Call Move_Tela_Memoria(objReprocessamentoEST)

    '*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar comentado
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 79966
    '***
    
    lErro = CF("Rotina_Reproc_MovEstoque", sNomeArqParam, objReprocessamentoEST)
    If lErro <> SUCESSO Then gError 79967
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoReprocessar_Click:

    Select Case gErr
        
        Case 79966, 79967, 90602
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173748)
            
    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = 0 'Definir IDH
    Set Form_Load_Ocx = Me
    Caption = "Reprocessamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Reprocessamento"
    
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
    'm_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

'Trecho referente ao modo de edição

Private Sub LabelDataInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataInicio, Source, X, Y)
End Sub

Private Sub LabelDataInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataInicio, Button, Shift, X, Y)
End Sub
Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Picture1, Source, X, Y)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Picture1, Button, Shift, X, Y)
End Sub

'*** Fernando, favor subir essa função para o RotinasMAT ***

'Function MATConfig_Le(objMATConfig As ClassMATConfig) As Long
''Le um registro de MATConfig
'
'Dim lErro As Long
'Dim lComando As Long
'Dim sConteudo As String
'
'
'On Error GoTo Erro_MATConfig_Le
'
'    'Inicia o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 79952
'
'    'Inicializa a string
'    sConteudo = String(STRING_CONTEUDO, 0)
'
'    'Executa a leitura no BD
'    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo= ? And FilialEmpresa= ?", sConteudo, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 79953
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79954
'
'    'Se não encontrou o registro => erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 79955
'
'    objMATConfig.sConteudo = sConteudo
'
'    Call Comando_Fechar(lComando)
'
'    MATConfig_Le = SUCESSO
'
'    Exit Function
'
'Erro_MATConfig_Le:
'
'    MATConfig_Le = gErr
'
'    Select Case gErr
'
'        Case 79953, 79954
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG", gErr, objMATConfig.sCodigo)
'
'        Case 79955 'sem dados
'
'        Case 79952
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173749)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Private Sub DataFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFim_Validate

    If Len(DataFim.ClipText) = 0 Then Exit Sub

    lErro = Data_Critica(DataFim.Text)
    If lErro <> SUCESSO Then gError 90567

    Exit Sub

Erro_DataFim_Validate:

    Cancel = True
    
    Select Case gErr

        Case 90567
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173750)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 90570

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 90571

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoCodigo, LabelProdutoDescricao)
    If lErro <> SUCESSO Then gError 90572
    
'    For iIndice = 0 To 1
'        ProdutoLabel(iIndice).Caption = ProdutoCodigo.Text
'    Next

    Call ProdutoCodigo_Validate(bSGECancelDummy)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 90573

    Me.Show
    
    iAlterado = 0

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 90570, 90572, 90573

        Case 90571
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173751)

    End Select

    Exit Sub

End Sub

Function Trata_Produto() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim colAlmoxarifados As New Collection
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Produto

    If iProdutoCodigoAlterado = 1 Then
    
        'Verifica preenchimento de Produto
        If Len(Trim(ProdutoCodigo.ClipText)) > 0 Then

            sProduto = ProdutoCodigo.Text

'            Call Limpa_Reprocessamento
            
            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido, True)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 90575
            
            'O Produto não está cadastrado
            If lErro = 25041 Then gError 90576

            'O Produto é Gerencial
            If lErro = 25043 Then gError 90577
            
            'O Produto não é Inventariado
            If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 90578
            
            'Preenche ProdutoDescricao com Descrição do Produto
            LabelProdutoDescricao.Caption = objProduto.sDescricao
            
            ProdutoCodigo.PromptInclude = False
            ProdutoCodigo.Text = objProduto.sCodigo
            ProdutoCodigo.PromptInclude = True

'        'Se Produto não está preenchido
'        Else
'
'            'Limpa a tela
'            Call Limpa_Reprocessamento
            
        End If

        iProdutoCodigoAlterado = 0
        
        iAlterado = 0

    End If

    Trata_Produto = SUCESSO

    Exit Function
    
Erro_Trata_Produto:
    
    Trata_Produto = gErr
    
    Select Case gErr

        Case 90575, 90577
            ProdutoCodigo.SetFocus
            
        Case 90576
            'Não encontrou Produto no BD
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            ProdutoCodigo.SetFocus

        Case 90578
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
            ProdutoCodigo.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173752)

    End Select

    Exit Function

End Function

Private Function Limpa_Reprocessamento()

    If dtDataMinimaReprocessamento <> DATA_NULA Then
    
        DataInicio.Text = Format(dtDataMinimaReprocessamento, "dd/mm/yy")
    
    Else
    
        DataInicio.PromptInclude = False
        DataInicio.Text = ""
        DataInicio.PromptInclude = True
    
    End If
    
    DataFim.PromptInclude = False
    DataFim.Text = ""
    DataFim.PromptInclude = True
    
    ProdutoCodigo.PromptInclude = False
    ProdutoCodigo.Text = ""
    ProdutoCodigo.PromptInclude = True
    
    LabelProdutoDescricao.Caption = ""
    IgnoraHora.Value = vbUnchecked
    ReprocTestaInt.Value = vbUnchecked
    LogReproc.Value = vbUnchecked
    LogReproc.Enabled = True
    PulaFaseDesfaz.Value = vbUnchecked
    PulaFaseDesfaz.Enabled = True
    AcertaEstProd.Value = vbUnchecked
    AcertaEstProd.Enabled = True
    
    'Altera a caption do LabelDataInicio
    LabelDataInicio.Caption = "Iniciar o Reprocessamento em:"


End Function

Private Sub UpDownDataFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_DownClick

    If Len(DataFim.ClipText) = 0 Then Exit Sub
    
    lErro = Data_Up_Down_Click(DataFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 90580

    Exit Sub

Erro_UpDownDataFim_DownClick:

    Select Case gErr

        Case 90580
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173753)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataFim_UpClick

    If Len(DataFim.ClipText) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 90581
    
    Exit Sub

Erro_UpDownDataFim_UpClick:

    Select Case gErr

        Case 90581

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173754)

    End Select

    Exit Sub

End Sub

Private Function DataInicio_UltimoReprocessamento_Obtem() As Long

Dim lErro As Long
Dim objMATConfig As New ClassMATConfig

On Error GoTo Erro_DataInicio_UltimoReprocessamento_Obtem
    
    'Guarda em objMATConfig os parâmetros necessários para se obter a data default para início do reprocessamento
    objMATConfig.iFilialEmpresa = giFilialEmpresa
    objMATConfig.sCodigo = DATAINICIO_ULTIMO_REPROC
    
    lErro = CF("MATConfig_Le", objMATConfig)
    If lErro <> SUCESSO And lErro <> 89118 Then gError 90585
    
    'Guarda em uma variável global à tela a data encontrada
    dtDataInicioUltReproc = StrParaDate(objMATConfig.sConteudo)
    
    'Se a data encontrada é diferente de data nula => exibe a data na tela como default
    If dtDataInicioUltReproc <> DATA_NULA Then DataInicio.Text = Format(dtDataInicioUltReproc, "dd/mm/yy")
    
    DataInicio_UltimoReprocessamento_Obtem = SUCESSO
    
    Exit Function

Erro_DataInicio_UltimoReprocessamento_Obtem:

    DataInicio_UltimoReprocessamento_Obtem = gErr
    
    Select Case gErr
        
        Case 90585
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173755)
            
    End Select
    
    Exit Function
    
End Function

Function Reprocessamento_Critica() As Long

On Error GoTo Erro_Reprocessamento_Critica

    'Se a data de início do reprocessamento não foi informada =>erro
    If Len(Trim(DataInicio.ClipText)) = 0 Then gError 79964
    
    'Se vai ser executado o teste de integridade
    If ReprocTestaInt.Value = vbChecked Then
        
        'Se a data de início do último reprocessamento executado é nula => erro
        'pois não pode ser executado teste de integridade sem que tenha sido feito pelo menos um reprocessamento
        If dtDataInicioUltReproc = DATA_NULA Then gError 90594
        
        'Se a data de sugestão para início do reprocessamento for diferente de data nula => erro
        'Isso significa que o BD deve ser reprocessado e portanto não faz sentido testar a integridade antes de executar o reprocessamento
        If dtDataMinimaReprocessamento <> DATA_NULA Then gError 90601
        
        'Se a data de início do último reprocessamento executado é menor do que a data de início do teste de integridade => erro
        If dtDataInicioUltReproc < StrParaDate(DataInicio.Text) Then gError 90595
    
    'Senão
    Else
    
        'Se encontrou uma data mínima para início do reprocessamento e essa data informada é anterior à data informada para início do reprocessamento =>erro
        If (dtDataMinimaReprocessamento <> DATA_NULA) And (StrParaDate(DataInicio.Text) > dtDataMinimaReprocessamento) Then gError 79965
    
    End If
    
    'Se a fase desfaz não vai ser executada
    If PulaFaseDesfaz.Value = vbChecked Then
    
        'Verifica se a data de início do reprocessamento é a data do primeiro movestoque do sistema
        If DataInicio.Text > dtDataPrimeiroMovEstoque Then gError 90709
    
    End If
        
    'Se a data final foi preenchida e for menor que a data inicial => erro
    If Len(Trim(DataFim.ClipText)) > 0 Then
        
        If DataFim.Text < DataInicio.Text Then gError 90579
    
    End If
    
    Reprocessamento_Critica = SUCESSO
    
    Exit Function
    
Erro_Reprocessamento_Critica:

    Reprocessamento_Critica = gErr
    
    Select Case gErr
    
        Case 90709
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_REPROC_DATA_PRIMEIRO_MOVESTOQUE", gErr, dtDataPrimeiroMovEstoque)
        
        Case 79964
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAOPREENCHIDA", gErr)
        
        Case 79965
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_REPROC_MENOR_DATA_INICIO", gErr, DataInicio.Text, dtDataMinimaReprocessamento)
        
        Case 90579
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_MENOR_DATA_INICIO", gErr, DataFim.Text, DataInicio.Text)
            
        Case 90594
            Call Rotina_Erro(vbOKOnly, "ERRO_REPROC_NAO_EXECUTADO", gErr)

        Case 90595
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_TESTEINT_MAIOR_DATA_ULTREPROC", gErr, DataInicio.Text, dtDataInicioUltReproc)
        
        Case 90601
            Call Rotina_Erro(vbOKOnly, "ERRO_REPROC_NAO_EXECUTADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173756)
        
    End Select
    
    Exit Function
    
End Function
    
Private Sub Move_Tela_Memoria(objReprocessamentoEST As ClassReprocessamentoEST)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objReprocessamentoEST.iFilialEmpresa = giFilialEmpresa
    objReprocessamentoEST.dtDataInicio = DataInicio.Text
    
    'se estiver preenchido
    If Len(Trim(ProdutoCodigo.ClipText)) > 0 Then

        'Formata o código do produto
        lErro = CF("Produto_Formata", ProdutoCodigo.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objReprocessamentoEST.sProdutoCodigo = sProdutoFormatado
        
    Else
        
        objReprocessamentoEST.sProdutoCodigo = ""
    
    End If
    
    'Se a data final foi preenchida
    If Len(Trim(DataFim.ClipText)) > 0 Then
        'Guarda a data no obj
        objReprocessamentoEST.dtDataFim = DataFim.Text
    'Senão
    Else
        'Guarda data nula no obj
        objReprocessamentoEST.dtDataFim = DATA_NULA
    End If
    
    'Se foi marcada a flag que indica para rodar o teste de integridade
    If ReprocTestaInt.Value = vbChecked Then
        'Guarda no obj a informação de que deverá ser executado o teste de integridade
        'Nesse caso o reprocessamento não é feito
        objReprocessamentoEST.iReprocTestaInt = REPROCESSAMENTO_TESTA_INTEGRIDADE
    Else
        'Guarda no obj a informação de que deverá ser executado reprocessamento normalmente
        objReprocessamentoEST.iReprocTestaInt = REPROCESSAMENTO_NORMAL
    End If
    
    'Se foi marcada a flag para pular fase desfaz
    If PulaFaseDesfaz.Value = vbChecked Then
        'Guarda no obj a informação de que a fase desfaz não deverá ser executada
        objReprocessamentoEST.iPulaFaseDesfaz = REPROCESSAMENTO_PULA_DESFAZ
    Else
        'Guarda no obj a informação de que não deverá pular a fase desfaz
        objReprocessamentoEST.iLogReproc = REPROCESSAMENTO_NORMAL
    End If
    
    'Se foi marcada a flag para efetuar log enquanto reprocessa
    If LogReproc.Value = vbChecked Then
        'Guarda no obj a informação de que deverá ser gerado log
        objReprocessamentoEST.iLogReproc = REPROCESSAMENTO_GERA_LOG
    Else
        'Guarda no obj a informação de que não deverá ser gerado log
        objReprocessamentoEST.iLogReproc = REPROCESSAMENTO_NAO_GERA_LOG
    End If
        
    
    'Se foi marcada para acertar o estoqueproduto
    If AcertaEstProd.Value = vbChecked Then
        objReprocessamentoEST.iAcertaEstProd = REPROCESSAMENTO_ACERTA_ESTPROD
    Else
        objReprocessamentoEST.iAcertaEstProd = REPROCESSAMENTO_NAO_ACERTA_ESTPROD
    End If
    
    
    'Se foi marcada a flag para reprocessar primeiro as entradas
    If IgnoraHora.Value = vbChecked Then
        'Indica que primeiro serão reprocessados os movimentos de entrada e depois os movimentos de saída
        objReprocessamentoEST.iOrdemReproc = REPROCESSAMENTO_ORDENA_ENTRADAS
    Else
        'Indica que os movimentos serão reprocessados pela ordem horária em que foram gerados
        objReprocessamentoEST.iOrdemReproc = REPROCESSAMENTO_ORDENA_HORA
    End If
    
    If ApenasSaldoTerc.Value = vbChecked Then
        objReprocessamentoEST.iApenasSaldoTerc = MARCADO
    Else
        objReprocessamentoEST.iApenasSaldoTerc = DESMARCADO
    End If
    
    objReprocessamentoEST.sFilialEmpGrupo = Trim(FilEmpGrupo.Text)
    
    Exit Sub

Erro_Move_Tela_Memoria:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173757)
            
    End Select
    
    Exit Sub
    
End Sub

Private Function Estoque_AtualizaItemMov2(objItemMovEst As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objEstoqueProduto As ClassEstoqueProduto) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Estoque_AtualizaItemMov2
    
    objEstoqueProduto.sProduto = objItemMovEst.sProduto
    objEstoqueProduto.iAlmoxarifado = objItemMovEst.iAlmoxarifado
    
    'se a movimentação é referente a um conserto
    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConserto = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConserto3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se é uma movimentação referente a um conserto
    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConserto = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConserto3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a uma demonstração
    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantDemo = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantDemo3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a uma demonstração
    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantDemo = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantDemo3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a material em consignação
    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConsig = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConsig3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a material em consignação
    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConsig = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConsig3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material
    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantOutras = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantOutras3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material
    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantOutras = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantOutras3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a material em beneficiamento
    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantBenef = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantBenef3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se é uma movimentação referente a um conserto
    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantBenef = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantBenef3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    Estoque_AtualizaItemMov2 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov2:

    Estoque_AtualizaItemMov2 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173758)
        
    End Select
        
    Exit Function
    
End Function
    
Private Function Estoque_AtualizaItemMov3(objItemMovEst As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objEstoqueProduto As ClassEstoqueProduto, objSldDiaEst As ClassSldDiaEst) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Estoque_AtualizaItemMov3
    
    'se a movimentação é referente a outras movimentações de material indisponivel
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
        objEstoqueProduto.dQuantInd = objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a outras movimentações de material indisponivel
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
        objEstoqueProduto.dQuantInd = -objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a nosso material disponivel
    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
        objEstoqueProduto.dQuantDispNossa = objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a nosso material disponivel
    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
        objEstoqueProduto.dQuantDispNossa = -objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a material defeituoso
    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
        objEstoqueProduto.dQuantDefeituosa = objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a material defeituoso
    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
        objEstoqueProduto.dQuantDefeituosa = -objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a material recebido e indisponível
    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
        objEstoqueProduto.dQuantRecIndl = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dQuantEntRecIndl = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorEntRecIndl = objItemMovEst.dCusto
    End If
    
    'se a movimentação é referente a material recebido e indisponível
    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
        objEstoqueProduto.dQuantRecIndl = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dQuantSaiRecIndl = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorSaiRecIndl = objItemMovEst.dCusto
    End If
    
    'se a movimentação é referente a material em ordem de producao
    If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
        objEstoqueProduto.dQuantOP = objItemMovEst.dQuantidadeOPEst
    End If
    
    'se a movimentação é referente a material em ordem de producao
    If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
        objEstoqueProduto.dQuantOP = -objItemMovEst.dQuantidadeOPEst
    End If
    
    objSldDiaEst.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldDiaEst.dtData = objItemMovEst.dtData
    objSldDiaEst.sProduto = objItemMovEst.sProduto
    
    'se a movimentação se refere a uma entrada de material no estoque
    If (objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.iCodigo = MOV_EST_AJUSTE_CUSTO_STD_NOSSO) And objTipoMovEstoque.iCodigo <> MOV_EST_MAT_NOSSO_PARA_BENEF_ENTRADA And objTipoMovEstoque.sEntradaSaidaCMP = objTipoMovEstoque.sEntradaOuSaida Then
'        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'            objSldDiaEst.dQuantEntrada = -objItemMovEst.dQuantidadeEst
'            objSldDiaEst.dValorEntrada = -objItemMovEst.dCusto
'        Else
            objSldDiaEst.dQuantEntrada = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorEntrada = objItemMovEst.dCusto
'        End If
    'se a movimentação se refere a uma saida de material do estoque
    ElseIf objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA And objTipoMovEstoque.iCodigo <> MOV_EST_MAT_NOSSO_PARA_BENEF_SAIDA And objTipoMovEstoque.sEntradaSaidaCMP = objTipoMovEstoque.sEntradaOuSaida Then
'        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'            objSldDiaEst.dQuantSaida = -objItemMovEst.dQuantidadeEst
'            objSldDiaEst.dValorSaida = -objItemMovEst.dCusto
'        Else
            objSldDiaEst.dQuantSaida = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorSaida = objItemMovEst.dCusto
'        End If
    End If
    
    'se a movimentação se refere a consumo de material
    If objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO Then
        objSldDiaEst.dQuantCons = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorCons = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a consumo de material
    If objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO Then
        objSldDiaEst.dQuantCons = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorCons = -objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material
    If objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_ADICIONAVENDA Then
        objSldDiaEst.dQuantVend = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVend = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material
    If objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_SUBTRAIVENDA Then
        objSldDiaEst.dQuantVend = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVend = -objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material em consignação de terceiros
    If objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_ADICIONAVENDACONSIG3 Then
        objSldDiaEst.dQuantVendConsig3 = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVendConsig3 = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material em consignação de terceiros
    If objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_SUBTRAIVENDACONSIG3 Then
        objSldDiaEst.dQuantVendConsig3 = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVendConsig3 = -objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a compra de material
    If objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_ADICIONACOMPRA Then
        objSldDiaEst.dQuantComp = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorComp = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a consumo de material
    If objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_SUBTRAICOMPRA Then
        objSldDiaEst.dQuantComp = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorComp = -objItemMovEst.dCusto
    End If
    
    Estoque_AtualizaItemMov3 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov3:

    Estoque_AtualizaItemMov3 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173759)
        
    End Select
        
    Exit Function
    
End Function
    
Private Function Estoque_AtualizaItemMov4(objItemMovEst As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objEstoqueProduto As ClassEstoqueProduto, objSldDiaEst As ClassSldDiaEst) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Estoque_AtualizaItemMov4
    
    'se a movimentação se refere a uma entrada de material no estoque
    If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.iCodigo = MOV_EST_AJUSTE_CUSTO_STD_NOSSO Then
    
        'se é para atualizar os saldos que impactam no calculo do custo
        If objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            objSldDiaEst.dQuantEntCusto = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorEntCusto = objItemMovEst.dCusto
        ElseIf objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI Then
            objSldDiaEst.dQuantEntCusto = -objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorEntCusto = -objItemMovEst.dCusto
        End If
        
    'se a movimentação se refere a uma saida de material do estoque
    ElseIf objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA Then
    
        If objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            objSldDiaEst.dQuantSaiCusto = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorSaiCusto = objItemMovEst.dCusto
        ElseIf objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI Then
            objSldDiaEst.dQuantSaiCusto = -objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorSaiCusto = -objItemMovEst.dCusto
        End If
        
    End If
  
    'se é um produto de terceiros
    If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
        'se a movimentação é referente a material consignado
        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConsig3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConsig3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConsig3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConsig3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig3 = objItemMovEst.dCusto
'            End If
               
        ElseIf objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConsig3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConsig3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConsig3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConsig3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig3 = -objItemMovEst.dCusto
'            End If
        End If
        
        'se a movimentação é referente a uma demonstração
        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiDemo3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiDemo3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntDemo3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntDemo3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo3 = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntDemo3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntDemo3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiDemo3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiDemo3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo3 = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a um conserto
        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConserto3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConserto3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConserto3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConserto3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto3 = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConserto3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConserto3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConserto3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConserto3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto3 = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a outras movimentações de material
        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiOutros3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiOutros3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntOutros3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntOutros3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras3 = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntOutros3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntOutros3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiOutros3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiOutros3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras3 = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a material em beneficiamento
        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiBenef3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiBenef3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntBenef3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntBenef3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef3 = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntBenef3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntBenef3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiBenef3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiBenef3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef3 = -objItemMovEst.dCusto
'            End If
        End If
     
    End If
           
    'se é um produto nosso
    If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
        'se a movimentação é referente a material consignado
        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConsig = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConsig = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConsig = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConsig = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConsig = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConsig = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConsig = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConsig = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig = -objItemMovEst.dCusto
'            End If
        End If
        
        'se a movimentação é referente a uma demonstração
        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiDemo = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiDemo = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntDemo = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntDemo = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntDemo = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntDemo = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiDemo = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiDemo = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a um conserto
        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConserto = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConserto = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConserto = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConserto = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConserto = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConserto = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConserto = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConserto = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a outras movimentações de material
        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiOutros = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiOutros = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntOutros = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntOutros = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntOutros = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntOutros = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiOutros = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiOutros = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a material em beneficiamento
        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiBenef = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiBenef = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntBenef = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntBenef = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntBenef = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntBenef = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiBenef = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiBenef = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef = -objItemMovEst.dCusto
'            End If
        End If
     
    End If
           
    Estoque_AtualizaItemMov4 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov4:

    Estoque_AtualizaItemMov4 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173760)
        
    End Select
        
    Exit Function
    
End Function

