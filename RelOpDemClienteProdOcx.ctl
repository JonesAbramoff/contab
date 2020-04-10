VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpDemClienteProd 
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   ScaleHeight     =   4020
   ScaleWidth      =   8175
   Begin VB.Frame Frame3 
      Caption         =   "Mês / Ano"
      Height          =   825
      Left            =   225
      TabIndex        =   20
      Top             =   720
      Width           =   5505
      Begin VB.ComboBox Ano 
         Height          =   315
         ItemData        =   "RelOpDemClienteProdOcx.ctx":0000
         Left            =   3270
         List            =   "RelOpDemClienteProdOcx.ctx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   315
         Width           =   1095
      End
      Begin VB.ComboBox Mes 
         Height          =   315
         ItemData        =   "RelOpDemClienteProdOcx.ctx":006C
         Left            =   630
         List            =   "RelOpDemClienteProdOcx.ctx":0097
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   330
         Width           =   1050
      End
      Begin VB.Label LabelAno 
         Caption         =   "Ano:"
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
         Height          =   240
         Left            =   2820
         TabIndex        =   24
         Top             =   360
         Width           =   420
      End
      Begin VB.Label labelMes 
         Caption         =   "Mês:"
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
         Height          =   285
         Left            =   75
         TabIndex        =   23
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5865
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemClienteProdOcx.ctx":0100
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemClienteProdOcx.ctx":025A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemClienteProdOcx.ctx":03E4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemClienteProdOcx.ctx":0916
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   210
      TabIndex        =   7
      Top             =   2565
      Width           =   5505
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   615
         TabIndex        =   8
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   315
         Left            =   615
         TabIndex        =   9
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label LabelProdutoAte 
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
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   840
         Width           =   360
      End
      Begin VB.Label LabelProdutoDe 
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   405
         Width           =   315
      End
      Begin VB.Label DescProdDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2175
         TabIndex        =   11
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label DescProdAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2175
         TabIndex        =   10
         Top             =   840
         Width           =   3000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   900
      Left            =   210
      TabIndex        =   2
      Top             =   1635
      Width           =   5505
      Begin MSMask.MaskEdBox ClienteDe 
         Height          =   300
         Left            =   630
         TabIndex        =   3
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteAte 
         Height          =   300
         Left            =   3255
         TabIndex        =   4
         Top             =   360
         Width           =   1890
         _ExtentX        =   3334
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   405
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
         Left            =   2835
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   5
         Top             =   420
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDemClienteProdOcx.ctx":0A94
      Left            =   1950
      List            =   "RelOpDemClienteProdOcx.ctx":0A96
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   2916
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
      Left            =   6120
      Picture         =   "RelOpDemClienteProdOcx.ctx":0A98
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   855
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   1260
      TabIndex        =   19
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpDemClienteProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Inserido por Wagner
'#####################
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
'#####################

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis do Browser
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'***** CARREGAMENTO DA TELA - INÍCIO *****
Public Sub Form_Load()

Dim lErro As Long
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_Form_Load
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 125707

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 125708

    'Carrega o ano e o mês
    Call Carrega_Mes_Ano(sMes, sAno)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 125707, 125708

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179457)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 125709
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 125710
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 125709
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 125710
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179458)

    End Select

    Exit Function

End Function
'***** CARREGAMENTO DA TELA - FIM *****

'***** EVENTO VALIDATE DOS CONTROLES - INÍCIO *****
Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoDe_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoDe, DescProdDe)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 125711
    
    If lErro <> SUCESSO Then gError 125712

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125711

        Case 125712
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179459)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoAte_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoAte, DescProdAte)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 125713
    
    If lErro <> SUCESSO Then gError 125714

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125713

        Case 125714
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179460)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteDe_Validate

    'se está Preenchido
    If Len(Trim(ClienteDe.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteDe, objCliente, 0)
        If lErro <> SUCESSO Then gError 125715

    End If

    Exit Sub

Erro_ClienteDe_Validate:

    Cancel = True

    Select Case gErr

        Case 125715

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179461)

    End Select

End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteAte_Validate

    'Se está Preenchido
    If Len(Trim(ClienteAte.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(ClienteAte, objCliente, 0)
        If lErro <> SUCESSO Then gError 125716

    End If

    Exit Sub

Erro_ClienteAte_Validate:

    Cancel = True

    Select Case gErr

        Case 125716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179462)

    End Select

End Sub
'***** EVENTO VALIDATE DOS CONTROLES - FIM *****

'***** EVENTO CLICK DOS CONTROLES - INÍCIO *****
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoAte.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 125717

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 125717

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179463)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoDe.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 125718

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 125718

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179464)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    If Len(Trim(ClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteAte.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)

End Sub

Private Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteDe.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 125719

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 125720

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = LimpaRelatorioDemClienteProd()
        If lErro <> SUCESSO Then gError 125721
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 125719
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 125720, 125721

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179465)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    'Preenche o Relatório
    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 125722

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 125722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179466)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 125723

    'Preenche o Relatório com os dados da tela
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 125724

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 125725

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 125723
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 125724, 125725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179467)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar

    'Limpa a tela
    lErro = LimpaRelatorioDemClienteProd()
    If lErro <> SUCESSO Then gError 125726
    
    Exit Sub
    
Erro_BotaoLimpar:

    Select Case gErr

        Case 125726
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179468)

    End Select

    Exit Sub

End Sub
'***** EVENTO CLICK DOS CONTROLES - FIM *****

'***** FUNÇÕES DE APOIO À TELA *****
Private Function LimpaRelatorioDemClienteProd() As Long
'Limpa a tela
    
Dim lErro As Long
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_LimpaRelatorioDemClienteProd
    
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 125727
    
    'limpa a combo opções
    ComboOpcoes.Text = ""
    
    'limpa a descrição do produto
    DescProdDe.Caption = ""
    DescProdAte.Caption = ""
    
    Call Carrega_Mes_Ano(sMes, sAno)
    
    LimpaRelatorioDemClienteProd = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioDemClienteProd:
    
    LimpaRelatorioDemClienteProd = gErr
    
    Select Case gErr

        Case 125727
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179469)

    End Select

    Exit Function

End Function

Private Sub Carrega_Mes_Ano(Optional sMes As String, Optional sAno As String)
'Função responsável pelo carregamento das combos ano e mês que não são editaveis

Dim iMes As Integer
Dim iAno As Integer
Dim iIndice As Integer
Dim iMax As Integer

    'Se a função for chamada de Define_Padrao
    If Len(Trim(sMes)) = 0 Then
        iMes = Month(Date)
    'Se a função for chamada de PreencheParametros na tela
    Else
        iMes = CInt(sMes)
    End If
    
    'Se a função for chamada de Define_Padrao
    If Len(Trim(sAno)) = 0 Then
        iAno = Year(Date)
    'Se a função for chamada de PreencheParametros na tela
    Else
        iAno = CInt(sAno)
    End If
    
    Mes.ListIndex = iMes - 1
    
    iMax = Ano.ListCount
       
    For iIndice = 0 To iMax
        If Ano.List(iIndice) = CStr(iAno) Then
            Ano.ListIndex = iIndice
            Exit For
        End If
    Next

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long, lNumIntRel As Long
Dim sCliente_De As String
Dim sCliente_Ate As String
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sCliente_De, sCliente_Ate, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 125728
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 125729
        
    'Preenche o produto Inicial
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 125730

    'Preenche o produto Final
    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 125731
    
    'Inclui o cliente inicial
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sCliente_De)
    If lErro <> AD_BOOL_TRUE Then gError 125732
    
    'Inclui o cliente final
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sCliente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 125733
            
    'Inclui o Mês
    lErro = objRelOpcoes.IncluirParametro("NMES", CStr(Mes.ItemData(Mes.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then gError 125734

    'Inclui o Ano
    lErro = objRelOpcoes.IncluirParametro("NANO", Ano.Text)
    If lErro <> AD_BOOL_TRUE Then gError 125735
    
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCliente_De, sCliente_Ate, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 125736
    
    If bExecutar Then
    
        'Alterado por Wagner
        lErro = RelDemoMensalVenda_Prepara(lNumIntRel, StrParaLong(sCliente_De), StrParaLong(sCliente_Ate), Mes.ItemData(Mes.ListIndex), StrParaInt(Ano.Text), sProd_I, sProd_F)
        If lErro <> SUCESSO Then gError 125736
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 125788
    
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 125728 To 125736
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179470)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCliente_De As String, sCliente_Ate As String, sProd_I As String, sProd_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 125737

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 125738

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 125739

    End If
    
    'Verifica se o Cliente inicial foi preenchido
    If ClienteDe.Text <> "" Then
        sCliente_De = CStr(LCodigo_Extrai(ClienteDe.Text))
    Else
        sCliente_De = ""
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If ClienteAte.Text <> "" Then
        sCliente_Ate = CStr(LCodigo_Extrai(ClienteAte.Text))
    Else
        sCliente_Ate = ""
    End If
            
    'Verifica se o Cliente Inicial é menor que o final, se não for --> ERRO
    If sCliente_De <> "" And sCliente_Ate <> "" Then
        
        If CInt(sCliente_De) > CInt(sCliente_Ate) Then gError 125740
        
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
        
        Case 125737
            ProdutoDe.SetFocus

        Case 125738
            ProdutoAte.SetFocus

        Case 125739
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus
        
        Case 125740
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            ClienteDe.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179471)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCliente_De As String, sCliente_Ate As String, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
    
'    'Verifica se o Produto foi preenchido
'    If sProd_I <> "" Then sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)
'
'    If sProd_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)
'
'    End If
'
'    'Verifica se o Cliente Inicial foi preenchido
'    If sCliente_De <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente >= " & Forprint_ConvInt(CInt(sCliente_De))
'
'    End If
'
'    'Verifica se o Cliente Final foi preenchido
'    If sCliente_Ate <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvInt(CInt(sCliente_Ate))
'
'    End If
'
'    'Verifica se o Mes está preenchido
'    If Trim(Mes.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Mês = " & Forprint_ConvTexto(Mes.Text)
'
'    End If
'
'    'Verifica se o ano está preenchido
'    If Trim(Ano.Text) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Ano = " & Forprint_ConvTexto(Ano.Text)
'
'    End If
'
'    If sExpressao <> "" Then
'
'        objRelOpcoes.sSelecao = sExpressao
'
'    End If

    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179472)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long
Dim sParam As String
Dim sMes As String
Dim sAno As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 125741
    
    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 125742

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, DescProdDe)
    If lErro <> SUCESSO Then gError 125743

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 125744

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, DescProdAte)
    If lErro <> SUCESSO Then gError 125745
    
    'Preenche Cliente inicial
    lErro = objRelOpcoes.ObterParametro("NCLIENTEINIC", sParam)
    If lErro <> SUCESSO Then gError 125746
    
    ClienteDe.Text = sParam
    Call ClienteDe_Validate(bSGECancelDummy)
    
    'Prenche Cliente final
    lErro = objRelOpcoes.ObterParametro("NCLIENTEFIM", sParam)
    If lErro <> SUCESSO Then gError 125747
    
    ClienteAte.Text = sParam
    Call ClienteAte_Validate(bSGECancelDummy)
    
    'pega o mês
    lErro = objRelOpcoes.ObterParametro("NMES", sParam)
    If lErro <> SUCESSO Then gError 125748
        
    'Aribui mês
    sMes = sParam
            
    'pega o ano
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    
    'Atribui o ano
    sAno = sParam
    
    'Com valores atribuídos de mês e ano, carrega as combos
    Call Carrega_Mes_Ano(sMes, sAno)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 125741 To 125748
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179473)

    End Select

    Exit Function

End Function

'***** FUNÇÕES DE APOIO À TELA - FIM
Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    ClienteDe.Text = CStr(objCliente.lCodigo)
    Call ClienteDe_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    'Preenche campo Cliente
    ClienteAte.Text = CStr(objCliente.lCodigo)
    Call ClienteAte_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 125749

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 125750
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, DescProdAte)
    If lErro <> SUCESSO Then gError 125751

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 125749, 125751

        Case 125750
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179474)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 125752

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 125753

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, DescProdDe)
    If lErro <> SUCESSO Then gError 125754

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 125752, 125754

        Case 125753
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179475)

    End Select

    Exit Sub

End Sub
'***** FUNÇÕES DO BROWSER - FIM *****

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Demonstrativo Mensal de Materiais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDemClienteProd"
    
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

Function RelDemoMensalVenda_Prepara(lNumIntRel As Long, ByVal lClienteDe As Long, ByVal lClienteAte As Long, ByVal iMes As Integer, ByVal iAno As Integer, ByVal sProdDe As String, ByVal sProdAte As String) As Long

Dim lErro As Long, iIndice As Integer, iComplementar As Integer
Dim lTransacao As Long, alComando(1 To 2) As Long
Dim dtDataInicio As Date, dtDataFim As Date, iFaturamento As Integer
Dim sProduto As String, sUnidadeMed As String
Dim lCliente As Long, dQuantidade As Double, dPrecoUnitarioMoeda As Double
Dim lClienteAnt As Long, sProdutoAnt As String, dQuantidadeAcum As Double, dValorAcum As Double
Dim sSQL As String


On Error GoTo Erro_RelDemoMensalVenda_Prepara

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 128180

    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 128181
    
    dtDataInicio = CDate("01/" & iMes & "/" & iAno)
    'Alterado por Wagner
    dtDataFim = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & iMes & "/" & iAno)))
    
    'Obtêm o NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_RENTABILIDADECLI", lNumIntRel)
    If lErro <> SUCESSO Then gError 128182
    
    sProduto = String(STRING_PRODUTO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
    'Alterado por Wagner
    '##########################################
    Call RelDemoMensalVendaSQL_Prepara(lClienteDe, lClienteAte, sProdDe, sProdAte, sSQL)
    
    lErro = RelDemoMensalVendaInt_Prepara(alComando(1), lClienteDe, lClienteAte, dtDataInicio, dtDataFim, sProdDe, sProdAte, DOCINFO_NFISCP, DOCINFO_NFISFCP, lCliente, sProduto, sUnidadeMed, dQuantidade, dPrecoUnitarioMoeda, iComplementar, iFaturamento, sSQL)
'    lErro = Comando_Executar(alComando(1), "SELECT Cliente, Produto, UnidadeMed, Quantidade, PrecoUnitarioMoeda, Complementar, Faturamento FROM ItensNFiscal, NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Faturamento IN (1,2) AND ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND PrecoUnitarioMoeda <> 0 AND DataEmissao BETWEEN ? and ? AND NFiscal.Status <> 7 AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) ORDER BY Cliente, Produto, UnidadeMed", _
'        lCliente, sProduto, sUnidadeMed, dQuantidade, dPrecoUnitarioMoeda, iComplementar, iFaturamento, dtDataInicio, dtDataFim, DOCINFO_NFISCP, DOCINFO_NFISFCP)
    If lErro <> AD_SQL_SUCESSO Then gError 124254
    '#########################################
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124255
    
    If lErro = AD_SQL_SUCESSO Then
        
        lClienteAnt = lCliente
        sProdutoAnt = sProduto
                
        Do While lErro = AD_SQL_SUCESSO
        
            If (lClienteAnt <> lCliente Or UCase(sProdutoAnt) <> UCase(sProduto)) Then
            
                lErro = Comando_Executar(alComando(2), "INSERT INTO RelDemoMensalVenda (NumIntRel, Cliente, Produto, Quantidade, Valor, PrecoMedio) VALUES (?,?,?,?,?,?)", _
                    lNumIntRel, lClienteAnt, sProdutoAnt, dQuantidadeAcum, dValorAcum, IIf(dQuantidadeAcum <> 0, dValorAcum / dQuantidadeAcum, 0))
                If lErro <> AD_SQL_SUCESSO Then gError 124256
                
                lClienteAnt = lCliente
                sProdutoAnt = sProduto
            
                dValorAcum = 0
                dQuantidadeAcum = 0
                
            End If
            
            'venda
            If iFaturamento = 1 Then
            
                dValorAcum = dValorAcum + (dQuantidade * dPrecoUnitarioMoeda)
            
                If iComplementar = 0 Then
                    dQuantidadeAcum = dQuantidadeAcum + dQuantidade
                End If
                
            Else 'devolucao de venda
                
                dValorAcum = dValorAcum - (dQuantidade * dPrecoUnitarioMoeda)
            
                If iComplementar = 0 Then
                    dQuantidadeAcum = dQuantidadeAcum - dQuantidade
                End If
            
            End If
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124257
        
        Loop
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelDemoMensalVenda (NumIntRel, Cliente, Produto, Quantidade, Valor, PrecoMedio) VALUES (?,?,?,?,?,?)", _
            lNumIntRel, lCliente, sProduto, dQuantidadeAcum, dValorAcum, IIf(dQuantidadeAcum <> 0, dValorAcum / dQuantidadeAcum, 0))
        If lErro <> AD_SQL_SUCESSO Then gError 124258

    End If
    
    'Fecha a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128198

    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelDemoMensalVenda_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelDemoMensalVenda_Prepara:

    RelDemoMensalVenda_Prepara = gErr
     
    Select Case gErr
          
        Case 128180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 128181
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 128198
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 124254 To 124258
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_RELDEMOMENSAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179476)
     
    End Select
     
    Call Transacao_Rollback
    
    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function


Private Sub RelDemoMensalVendaSQL_Prepara(ByVal vlClienteDe As Variant, ByVal vlClienteAte As Variant, ByVal vsProdDe As Variant, ByVal vsProdAte As Variant, sSQL As String)
'monta o comando SQL para obtencao das fretes dinamicamente e retorna.
Dim sSelect As String, sWhere As String, sFrom As String, sOrderBy As String

On Error GoTo Erro_RelDemoMensalVendaSQL_Prepara

    sSelect = "SELECT Cliente, " & _
                    "Produto, " & _
                    "UnidadeMed, " & _
                    "Quantidade, " & _
                    "PrecoUnitarioMoeda, " & _
                    "Complementar, " & _
                    "Faturamento "

    sFrom = "FROM  ItensNFiscal, " & _
                    "NFiscal, " & _
                    "TiposDocInfo "
                     
    sWhere = "WHERE  NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND " & _
                    "TiposDocInfo.Faturamento IN (1,2) AND " & _
                    "ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND " & _
                    "PrecoUnitarioMoeda <> 0 AND " & _
                    "DataEmissao BETWEEN ? and ? AND " & _
                    "NFiscal.Status <> 7 AND " & _
                    "(TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) "
     
     sOrderBy = "ORDER BY Cliente, " & _
                        "Produto, " & _
                        "UnidadeMed "
                         
   
    If vlClienteDe <> 0 Then
        sWhere = sWhere & "AND NFiscal.Cliente >= ? "
    End If
    
    If vlClienteAte <> 0 Then
        sWhere = sWhere & "AND NFiscal.Cliente <= ? "
    End If
    
    If Len(Trim(vsProdDe)) > 0 Then
        sWhere = sWhere & "AND ItensNFiscal.Produto >= ? "
    End If
    
    If Len(Trim(vsProdAte)) > 0 Then
        sWhere = sWhere & "AND ItensNFiscal.Produto <= ? "
    End If
    
    sSQL = sSelect & sFrom & sWhere & sOrderBy

    Exit Sub

Erro_RelDemoMensalVendaSQL_Prepara:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179477)

    End Select

    Exit Sub

End Sub

Private Function RelDemoMensalVendaInt_Prepara(ByVal lComando As Long, ByVal vlClienteDe As Variant, ByVal vlClienteAte As Variant, ByVal vdtDataInicio As Variant, ByVal vdtDataFim As Variant, ByVal vsProdDe As Variant, ByVal vsProdAte As Variant, ByVal viNumDoc1 As Variant, ByVal viNumDoc2 As Variant, vlCliente As Variant, vsProduto As Variant, vsUnidadeMed As Variant, vdQuantidade As Variant, vdPrecoUnitarioMoeda As Variant, viComplementar As Variant, viFaturamento As Variant, ByVal sSQL As String) As Long

Dim lErro As Long

On Error GoTo Erro_RelDemoMensalVendaInt_Prepara

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129171

    lErro = Comando_BindVarInt(lComando, vlCliente)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129172

    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129173
    
    lErro = Comando_BindVarInt(lComando, vsUnidadeMed)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129174

    lErro = Comando_BindVarInt(lComando, vdQuantidade)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129175

    lErro = Comando_BindVarInt(lComando, vdPrecoUnitarioMoeda)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129176
    
    lErro = Comando_BindVarInt(lComando, viComplementar)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129177
    
    lErro = Comando_BindVarInt(lComando, viFaturamento)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129178

    lErro = Comando_BindVarInt(lComando, vdtDataInicio)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129179

    lErro = Comando_BindVarInt(lComando, vdtDataFim)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129180
    
    lErro = Comando_BindVarInt(lComando, viNumDoc1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129181

    lErro = Comando_BindVarInt(lComando, viNumDoc2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129182
  
    If vlClienteDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129183
    End If
    
    If vlClienteAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlClienteAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129184
    End If
   
    If Len(Trim(vsProdDe)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsProdDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129185
    End If
    
    If Len(Trim(vsProdAte)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsProdAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129186
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129187
    
    RelDemoMensalVendaInt_Prepara = SUCESSO

    Exit Function

Erro_RelDemoMensalVendaInt_Prepara:

    RelDemoMensalVendaInt_Prepara = gErr

    Select Case gErr
    
        Case 129171 To 129187

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179478)

    End Select

    Exit Function

End Function


