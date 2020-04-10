VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpFlashVendas 
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   ScaleHeight     =   3945
   ScaleWidth      =   7785
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   735
      Left            =   233
      TabIndex        =   22
      Top             =   1680
      Width           =   5175
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   260
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   260
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCaixaDe 
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
         Height          =   195
         Left            =   960
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   315
         Width           =   315
      End
      Begin VB.Label LabelCaixaAte 
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
         Left            =   2760
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   320
         Width           =   360
      End
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   5160
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   4
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
         Left            =   510
         TabIndex        =   5
         Top             =   825
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
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   870
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
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   390
         Width           =   315
      End
      Begin VB.Label ProdutoDescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   19
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label ProdutoDescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   18
         Top             =   825
         Width           =   2970
      End
   End
   Begin VB.Frame FramePeriodoReferencia 
      Caption         =   "Período de Referência"
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   5175
      Begin VB.CommandButton BotaoRefresh 
         Height          =   345
         Left            =   4200
         Picture         =   "RelOpFlashVendas.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Atualiza Hora"
         Top             =   270
         Width           =   420
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   450
         TabIndex        =   16
         Top             =   345
         Width           =   480
      End
      Begin VB.Label LabelHora 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Left            =   2595
         TabIndex        =   15
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Hora 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   285
         Width           =   1095
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
      Left            =   5745
      Picture         =   "RelOpFlashVendas.ctx":0632
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5452
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpFlashVendas.ctx":0734
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpFlashVendas.ctx":08B2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpFlashVendas.ctx":0DE4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpFlashVendas.ctx":0F6E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpFlashVendas.ctx":10C8
      Left            =   1080
      List            =   "RelOpFlashVendas.ctx":10CA
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2670
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
      Left            =   360
      TabIndex        =   11
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpFlashVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''################### Só para teste ###############
'
'Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
'Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
'Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
'
''Declaração de um Obj Glogal ..
'Dim gobjUltRelFlashVendas As New ClassFlashVendas
'
'##################################################


Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

' Browse 's Relacionados a Tela Flash de Vendas
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoCaixaDe As AdmEvento
Attribute objEventoCaixaDe.VB_VarHelpID = -1
Private WithEvents objEventoCaixaAte As AdmEvento
Attribute objEventoCaixaAte.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Relatório Flash de Vendas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpFlashVendas"

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
    'Parent.UnloadDoFilho

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

'*********************************************************

 ' Sergio Ricardo Pacheco da Vitoria
 ' Inicio dia 13/12/2002
    'Supervisor:    Shirley
'*********************************************************

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCaixaDe = New AdmEvento
    Set objEventoCaixaAte = New AdmEvento

    'Formata o produto com o formato do Banco de Dados
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 113196

    'Formata o produto com o formato do Banco de Dados
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 113197

    'Carrega a data atual para a Tela
    Data.Caption = gdtDataHoje

    'Carrega a Hora atual para a Tela
    HORA.Caption = Time

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 113196, 113197

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169151)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 113198

    ComboOpcoes.Text = ""
    ProdutoDescricaoDe.Caption = ""
    ProdutoDescricaoAte.Caption = ""
    HORA.Caption = Time
    Data.Caption = Date
    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 113198

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169152)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 113199

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 113200

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 113200

        Case 113199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169153)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 113201

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 113202

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 113203

    'se a opção de relatório foi gravada em RelatorioOpcoes então adcionar a opção de relatório na comboopções
    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 113201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 113202, 113203

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169154)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim dHoraIni As Double
Dim dHoraFim As Double

On Error GoTo Erro_PreenchgerrelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 113204

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 113205

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 113206

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 113207

    lErro = objRelOpcoes.IncluirParametro("NCAIXADE", CaixaDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113208

    lErro = objRelOpcoes.IncluirParametro("NCAIXAATE", CaixaAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113209

    lErro = objRelOpcoes.IncluirParametro("DDATA", Data.Caption)
    If lErro <> AD_BOOL_TRUE Then gError 113218

    'Transformar a Hora Inicial em Double e Hora final( Hora Inicial + Uma Hora ) e m Double
    dHoraFim = CDbl(CDate(HORA.Caption))

    dHoraIni = CDbl(CDate(Format(CDbl(CDate(HORA.Caption)) - CDbl(CDate("01:00:00")), "hh:mm:ss")))

    lErro = objRelOpcoes.IncluirParametro("NHORAFIM", Forprint_ConvDouble(dHoraFim))
    If lErro <> AD_BOOL_TRUE Then gError 113228

    lErro = objRelOpcoes.IncluirParametro("NHORAINI", Forprint_ConvDouble(dHoraIni))
    If lErro <> AD_BOOL_TRUE Then gError 113227

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, StrParaInt(CaixaDe.Text), StrParaInt(CaixaAte.Text))
    If lErro <> SUCESSO Then gError 113210

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 113204 To 113210, 113218, 113227, 113228

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169155)

    End Select

    Exit Function

End Function

Private Sub CaixaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CaixaDe)

End Sub

Private Sub CaixaDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_DataDe_Validate

    'Verifica se existe alguma CaixaDe
    If Len(Trim(CaixaDe.Text)) <> 0 Then

        objCaixa.iCodigo = StrParaInt(CaixaDe.Text)
        objCaixa.iFilialEmpresa = giFilialEmpresa

        'Lê a Caixa
        lErro = CF("Caixas_Le", objCaixa)
        If lErro <> SUCESSO And lErro <> 79405 Then gError 113211

        If lErro = 79405 Then gError 113290
        
    End If

    Exit Sub

Erro_DataDe_Validate:

Cancel = True

    Select Case gErr

        Case 113211

        Case 113290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169156)

    End Select

    Exit Sub

End Sub

Private Sub CaixaAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CaixaAte)

End Sub


Private Sub CaixaAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_CaixaAte_Validate

    'Verifica se existe alguma CaixaDe
    If Len(Trim(CaixaAte.Text)) <> 0 Then

        objCaixa.iFilialEmpresa = giFilialEmpresa
        objCaixa.iCodigo = StrParaInt(CaixaAte.Text)

        'Lê a Caixa
        lErro = CF("Caixas_Le", objCaixa)
        If lErro <> SUCESSO And lErro <> 79405 Then gError 113212
    
        If lErro = 79405 Then gError 113289


    End If

    Exit Sub

Erro_CaixaAte_Validate:

Cancel = True

    Select Case gErr

        Case 113212

        Case 113289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169157)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoDe)

End Sub

Private Sub ProdutoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoAte)

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoDe_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 113229

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 113230

        ProdutoDescricaoDe.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 113231

        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 113232

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 113233

'        'Se não controla estoque => Erro
'        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 113234


    Else

        ProdutoDescricaoDe.Caption = ""

    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    ProdutoDescricaoDe.Caption = ""

    Select Case gErr

        Case 113229, 113230

        Case 113231
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoDe.Text)

        Case 113232
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoDe.Text)

        Case 113233
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoDe.Text)

        Case 113234
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169158)

    End Select

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoAte_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 113235

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 113236

        ProdutoDescricaoAte.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 113237

        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 113238

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 113239

'        'Se não controla estoque => Erro
'        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 113240


    Else

        ProdutoDescricaoAte.Caption = ""

    End If

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    ProdutoDescricaoAte.Caption = ""

    Select Case gErr

        Case 113235, 113236

        Case 113237
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoAte.Text)

        Case 113238
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoAte.Text)

        Case 113239
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoAte.Text)

        Case 113240
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169159)

    End Select

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 113213

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 113214

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 113215

    End If

    'Caixa final não pode ser maior que a caixa inicial
    If Trim(CaixaDe.ClipText) <> "" And Trim(CaixaAte.ClipText) <> "" Then

         If StrParaInt(CaixaDe.Text) > StrParaInt(CaixaAte.Text) Then gError 113216

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 113213
            ProdutoDe.SetFocus

        Case 113214
            ProdutoAte.SetFocus

        Case 113215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus

      Case 113216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXADE_MAIOR", gErr)
            CaixaDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169160)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, iCaixaDe As Integer, iCaixaAte As Integer) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If

    If iCaixaDe <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa >= " & Forprint_ConvInt(iCaixaDe)

    End If

    If iCaixaAte <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa <= " & Forprint_ConvInt(iCaixaAte)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169161)

    End Select

    Exit Function

End Function

Private Sub BotaoRefresh_Click()
'Botão que atualiza a Hora

Dim lErro As Long

On Error GoTo Erro_BotaoRefresh_Click

    If Len(Trim(HORA.Caption)) <> 0 Then

        HORA.Caption = Time

    End If

    Exit Sub

Erro_BotaoRefresh_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169162)

    End Select

Exit Sub

End Sub



'##### Inicio tratameto Browse's  ###########

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
        If lErro <> SUCESSO Then gError 113217

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 113217

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169163)

    End Select

    Exit Sub

End Sub

Private Sub LabelCaixaDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_LabelCaixaDe_Click

    'Verifica se o Caixa foi preenchido
    If Len(CaixaDe.ClipText) <> 0 Then

        objCaixa.iCodigo = StrParaInt(CaixaDe.Text)

    End If

    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixaDe)

    Exit Sub

Erro_LabelCaixaDe_Click:

    Select Case gErr

        Case 113254

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169164)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCaixaDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixaDe_evSelecao

    Set objCaixa = obj1

    'Lê a Caixa
    lErro = CF("Caixas_Le", objCaixa)
    If lErro <> SUCESSO And lErro <> 79405 Then gError 113256

    'Se não achou o Caixa --> erro
    If lErro = 79405 Then gError 113257

    CaixaDe.PromptInclude = False
    CaixaDe.Text = objCaixa.iCodigo
    CaixaDe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCaixaDe_evSelecao:

    Select Case gErr

        Case 113256

        Case 113257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169165)

    End Select

    Exit Sub

End Sub

Private Sub LabelCaixaAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCaixa As New ClassCaixa

On Error GoTo Erro_LabelCaixaAte_Click

    'Verifica se o Caixa foi preenchido
    If Len(CaixaAte.ClipText) <> 0 Then

        objCaixa.iCodigo = StrParaInt(CaixaAte.Text)

    End If

    Call Chama_Tela("CaixaLista", colSelecao, objCaixa, objEventoCaixaAte)

    Exit Sub

Erro_LabelCaixaAte_Click:

    Select Case gErr

        Case 113255

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169166)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCaixaAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCaixa As ClassCaixa

On Error GoTo Erro_objEventoCaixaAte_evSelecao

    Set objCaixa = obj1

    'Lê a Caixa
    lErro = CF("Caixas_Le", objCaixa)
    If lErro <> SUCESSO And lErro <> 79405 Then gError 113258

    'Se não achou o Caixa --> erro
    If lErro = 79405 Then gError 113259

    CaixaAte.PromptInclude = False
    CaixaAte.Text = objCaixa.iCodigo
    CaixaAte.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCaixaAte_evSelecao:

    Select Case gErr

        Case 113258

        Case 113259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169167)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 113220

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 113220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169168)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 113221

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 113222

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 113223

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 113221, 113223

        Case 113222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169169)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 113224

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 113225

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 113226

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 113224, 113226

        Case 113225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169170)

    End Select

    Exit Sub

End Sub

'##### Fim tratameto Browse's  ###########

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim bFalse As Boolean

On Error GoTo Erro_PreencherParametrosNaTela

    'Função que lê no Banco de dados o Codigo do Relatorio e Traz a Coleção de parâmetro carregados
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 113241

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 113242

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 113243

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 113244

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 113245

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NCAIXADE", sParam)
    If lErro <> SUCESSO Then gError 113246
    CaixaDe.PromptInclude = False
    CaixaDe.Text = sParam
    CaixaDe.PromptInclude = True
    Call CaixaDe_Validate(bFalse)

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NCAIXAATE", sParam)
    If lErro <> SUCESSO Then gError 113247
    CaixaAte.PromptInclude = False
    CaixaAte.Text = sParam
    CaixaAte.PromptInclude = True
    Call CaixaAte_Validate(bFalse)

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then gError 113248
    Data.Caption = sParam


    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NHORAFIM", sParam)
    If lErro <> SUCESSO Then gError 113248
    HORA.Caption = Format(sParam, "hh:mm:ss")

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 113241 To 113248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169171)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 113249

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FLASH_VENDAS")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 113250

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 113249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 113250

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169172)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

'So para Teste
Dim sProduto As String
Dim iFilialEmpresa  As Integer
Dim dtData As Date
Dim dHoraIni As Double
Dim dHoraFim As Double
Dim iCaixaIni As Integer
Dim iCaixaFim As Integer
Dim lNumVendas As Long
Dim dQuantVendida As Double

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 113252

'    'Provisório
'    sProduto = "0000001"
'    iFilialEmpresa = giFilialEmpresa
'    dtData = "19/12/02"
'    'Transformar a Hora Inicial em Double e Hora final( Hora Inicial + Uma Hora ) e m Double
    
    dHoraFim = CDbl(CDate(HORA.Caption))
    dHoraIni = CDbl(CDate(Format(CDbl(CDate(HORA.Caption)) - CDbl(CDate("01:00:00")), "hh:mm:ss")))
    iCaixaIni = StrParaInt(CaixaDe.Text)
    iCaixaFim = StrParaInt(CaixaAte.Text)
    
    
'    'Chamada Provisória para teste da Tela,  Funções em AdrelVb2
'    lErro = Obtem_QuantVend_IntervHorasAux(dQuantVendida, sProduto, iFilialEmpresa, dtData, dHoraIni, dHoraFim, iCaixaIni, iCaixaFim)

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 113252

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169173)

    End Select

    Exit Sub

End Sub


''Rotinas só para teste....
'
'Function Obtem_NumVendas_IntervHorasAux(lNumVendas As Long, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
''Função que Retorna o Numero de Vendas para uma determinada data e Horario
'
'Dim lErro As Long
'Dim objRelFlashVendas As ClassFlashVendas
'
'On Error GoTo Erro_Obtem_NumVendas_IntervHorasAux
'
'    'Verifica se os Valores guardados no Obj são diferentes aos passados por parâmetro
'    If sProduto <> gobjUltRelFlashVendas.sProduto Or iFilialEmpresa <> gobjUltRelFlashVendas.iFilialEmpresa Or dtData <> gobjUltRelFlashVendas.dtData Or dHoraIni <> gobjUltRelFlashVendas.dHoraIni Or dHoraFim <> gobjUltRelFlashVendas.dHoraFim Or iCaixaIni <> gobjUltRelFlashVendas.iCaixaIni Or iCaixaFim <> gobjUltRelFlashVendas.iCaixaFim Then
'
'        Set objRelFlashVendas = New ClassFlashVendas
'
'        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
'        objRelFlashVendas.dHoraFim = dHoraFim
'        objRelFlashVendas.dHoraIni = dHoraIni
'        objRelFlashVendas.dtData = dtData
'        objRelFlashVendas.iCaixaFim = iCaixaFim
'        objRelFlashVendas.iCaixaIni = iCaixaIni
'        objRelFlashVendas.iFilialEmpresa = iFilialEmpresa
'        objRelFlashVendas.iCaixaIni = iCaixaIni
'        objRelFlashVendas.sProduto = sProduto
'
'        'Função que lê as Estatisticas (Numero de vendas ) que serão utilizadas pelo Relatório
'        lErro = RelFlashVendas_Le_Estatisticas_Hora(objRelFlashVendas)
'        If lErro <> SUCESSO Then gError 113261
'
'        'Aponta o Obj Global para o Obj Local a Função
'        Set gobjUltRelFlashVendas = objRelFlashVendas
'
'    End If
'
'    'Guarda o Numero de Vandas para a Data e o Intervalo de Hora Especificado
'    lNumVendas = gobjUltRelFlashVendas.lNumVendas
'
'    Obtem_NumVendas_IntervHorasAux = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_NumVendas_IntervHorasAux:
'
'    Obtem_NumVendas_IntervHorasAux = gErr
'
'        Select Case gErr
'
'        Case 113261
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169174)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function Obtem_QuantVend_IntervHorasAux(dQuantVendida As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
''Função que Retorna a Quantidade vendida para uma determinada data e Horario
'
'Dim lErro As Long
'Dim objRelFlashVendas As ClassFlashVendas
'
'On Error GoTo Erro_Obtem_QuantVend_IntervHorasAux
'
'    'Verifica se os Valores guardados no Obj são diferentes aos passados por parâmetro
'    If sProduto <> gobjUltRelFlashVendas.sProduto Or iFilialEmpresa <> gobjUltRelFlashVendas.iFilialEmpresa Or dtData <> gobjUltRelFlashVendas.dtData Or dHoraIni <> gobjUltRelFlashVendas.dHoraIni Or dHoraFim <> gobjUltRelFlashVendas.dHoraFim Or iCaixaIni <> gobjUltRelFlashVendas.iCaixaIni Or iCaixaFim <> gobjUltRelFlashVendas.iCaixaFim Then
'
'        Set objRelFlashVendas = New ClassFlashVendas
'
'        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
'        objRelFlashVendas.dHoraFim = dHoraFim
'        objRelFlashVendas.dHoraIni = dHoraIni
'        objRelFlashVendas.dtData = dtData
'        objRelFlashVendas.iCaixaFim = iCaixaFim
'        objRelFlashVendas.iCaixaIni = iCaixaIni
'        objRelFlashVendas.iFilialEmpresa = iFilialEmpresa
'        objRelFlashVendas.iCaixaIni = iCaixaIni
'        objRelFlashVendas.sProduto = sProduto
'
'        'Função que lê as Estatisticas (Numero de vendas ) que serão utilizadas pelo Relatório
'        lErro = RelFlashVendas_Le_Estatisticas_Hora(objRelFlashVendas)
'        If lErro <> SUCESSO Then gError 113285
'
'        'Aponta o Obj Global para o Obj Local a Função
'        Set gobjUltRelFlashVendas = objRelFlashVendas
'
'    End If
'
'    'Guarda a Quantidade de vendas
'    dQuantVendida = gobjUltRelFlashVendas.dQuantVendida
'
'    Obtem_QuantVend_IntervHorasAux = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_QuantVend_IntervHorasAux:
'
'    Obtem_QuantVend_IntervHorasAux = gErr
'
'        Select Case gErr
'
'        Case 113285
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169175)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Obtem_VlrVendas_IntervHorasAux(dValorVendido As Double, ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtData As Date, ByVal dHoraIni As Double, ByVal dHoraFim As Double, ByVal iCaixaIni As Integer, ByVal iCaixaFim As Integer) As Long
''Função que Retorna o Valor vendido  para uma determinada data e Horario
'
'Dim lErro As Long
'Dim objRelFlashVendas As ClassFlashVendas
'
'On Error GoTo Erro_Obtem_VlrVendas_IntervHorasAux
'
'    'Verifica se os Valores guardados no Obj são diferentes aos passados por parâmetro
'    If sProduto <> gobjUltRelFlashVendas.sProduto Or iFilialEmpresa <> gobjUltRelFlashVendas.iFilialEmpresa Or dtData <> gobjUltRelFlashVendas.dtData Or dHoraIni <> gobjUltRelFlashVendas.dHoraIni Or dHoraFim <> gobjUltRelFlashVendas.dHoraFim Or iCaixaIni <> gobjUltRelFlashVendas.iCaixaIni Or iCaixaFim <> gobjUltRelFlashVendas.iCaixaFim Then
'
'        Set objRelFlashVendas = New ClassFlashVendas
'
'        'Atribui ao Obj Os valores passados pelo Gerador de Relatório
'        objRelFlashVendas.dHoraFim = dHoraFim
'        objRelFlashVendas.dHoraIni = dHoraIni
'        objRelFlashVendas.dtData = dtData
'        objRelFlashVendas.iCaixaFim = iCaixaFim
'        objRelFlashVendas.iCaixaIni = iCaixaIni
'        objRelFlashVendas.iFilialEmpresa = iFilialEmpresa
'        objRelFlashVendas.iCaixaIni = iCaixaIni
'        objRelFlashVendas.sProduto = sProduto
'
'        'Função que lê as Estatisticas (Numero de vendas ) que serão utilizadas pelo Relatório
'        lErro = RelFlashVendas_Le_Estatisticas_Hora(objRelFlashVendas)
'        If lErro <> SUCESSO Then gError 113286
'
'        'Aponta o Obj Global para o Obj Local a Função
'        Set gobjUltRelFlashVendas = objRelFlashVendas
'
'    End If
'
'    'Guarda o Valor vendido
'    dValorVendido = gobjUltRelFlashVendas.dValorVendido
'
'    Obtem_VlrVendas_IntervHorasAux = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_VlrVendas_IntervHorasAux:
'
'    Obtem_VlrVendas_IntervHorasAux = gErr
'
'        Select Case gErr
'
'        Case 113286
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169176)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Function RelFlashVendas_Le_Estatisticas_Hora(ByVal objRelFlashVendas As ClassFlashVendas) As Long
''Função que monta o Select que será executado posteriormente
'
'Dim lErro As Long
'Dim sSelect As String
'Dim tRelFlashVendas As TypeRelFlashVendasVar
'Dim lComando As Long
'
'On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora
'
'    'abre o comando
'    lComando = Comando_Abrir
'    If lComando = 0 Then gError 113262
'
'    'Função que Monta o select
'    lErro = RelFlashVendas_Le_Estatisticas_Hora1(objRelFlashVendas, sSelect)
'    If lErro <> SUCESSO Then gError 113263
'
'    'Função que Prepara a Parte Fixa do Select
'    lErro = RelFlashVendas_Le_Estatisticas_Hora2(lComando, sSelect, tRelFlashVendas)
'    If lErro <> SUCESSO Then gError 113264
'
'    'Prepara os parâmetros que variam de acordo com a seleção do usuário
'    lErro = RelFlashVendas_Le_Estatisticas_Hora3(lComando, tRelFlashVendas, objRelFlashVendas, sSelect)
'    If lErro <> SUCESSO Then gError 113265
'
'    'Função que Processa os Registros Retormados Pelo Select
'    lErro = RelFlashVendas_Le_Estatisticas_Hora4(lComando, tRelFlashVendas, objRelFlashVendas, sSelect)
'    If lErro <> SUCESSO Then gError 113266
'
'    'fecha o comando
'    Call Comando_Fechar(lComando)
'
'    RelFlashVendas_Le_Estatisticas_Hora = SUCESSO
'
'    Exit Function
'
'Erro_RelFlashVendas_Le_Estatisticas_Hora:
'
'    RelFlashVendas_Le_Estatisticas_Hora = gErr
'
'    Select Case gErr
'
'        Case 113262
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 113263 To 113266
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169177)
'
'    End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Function RelFlashVendas_Le_Estatisticas_Hora1(ByVal objRelFlashVendas As ClassFlashVendas, sSelect As String) As Long
''Função que Guarda na String sSelect o comando que será utilizada para ler os registro em SldDiaFat
'
'Dim lErro As Long
'Dim sFrom As String
'Dim sWhere As String
'
'On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora1
'
'    'select fixo
'    sSelect = "SELECT  ItensCupomFiscal.UnidadeMed , COUNT(ItensCupomFiscal.UnidadeMed)AS NumUnidadeMed , SUM(ItensCupomFiscal.Quantidade)AS NumQuantidade , SUM(ItensCupomFiscal.Quantidade * ItensCupomFiscal.PrecoUnitario) AS Valor "
'
'    'From Fixo
'    sFrom = " FROM CupomFiscal , ItensCupomFiscal"
'
'    'Cláusula Where
'    sWhere = " WHERE  CupomFiscal.NumIntDoc = ItensCupomFiscal.NumIntCupom AND ItensCupomFiscal.Tipo = 1 AND CupomFiscal.Tipo = 1 AND ItensCupomFiscal.Produto = ? AND CupomFiscal.HoraEmissao >= ? AND CupomFiscal.HoraEmissao <= ? AND CupomFiscal.DataEmissao = ? "
'
'    'Verifica se o Filtro utilizado vai ser por filial ou não
'    If objRelFlashVendas.iFilialEmpresa <> EMPRESA_TODA Then
'
'        sWhere = sWhere & " AND CupomFiscal.FilialEmpresa = ItensCupomFiscal.FilialEmpresa AND ItensCupomFiscal.FilialEmpresa = ? "
'
'    End If
'
'
'
'    'Verifica se o Filtro utilizado vai ser por Caixa DE ou Não
'    If objRelFlashVendas.iCaixaIni <> CODIGO_NAO_PREENCHIDO Then
'
'        sWhere = sWhere & " AND CupomFiscal.Caixa >= ? "
'
'    End If
'
'
'    'Verifica se o Filtro utilizado vai ser por Caixa ATE  ou Não
'    If objRelFlashVendas.iCaixaFim <> CODIGO_NAO_PREENCHIDO Then
'
'        sWhere = sWhere & " AND CupomFiscal.Caixa <= ? "
'
'    End If
'
'    'Adciona o group By
'    sWhere = sWhere & "Group By (ItensCupomFiscal.Produto),( ItensCupomFiscal.UnidadeMed)"
'
'    sSelect = sSelect & sFrom & sWhere
'
'    RelFlashVendas_Le_Estatisticas_Hora1 = SUCESSO
'
'    Exit Function
'
'Erro_RelFlashVendas_Le_Estatisticas_Hora1:
'
'    RelFlashVendas_Le_Estatisticas_Hora1 = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169178)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function RelFlashVendas_Le_Estatisticas_Hora2(ByVal lComando As Long, sSelect As String, tRelFlashVendas As TypeRelFlashVendasVar) As Long
''Função que Bind o as variáveis que serão recebidas pelo select
'
'Dim lErro As Long
'
'On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora2
'
'    With tRelFlashVendas
'
'        .vsUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
'
'        lErro = Comando_PrepararInt(lComando, sSelect)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113267
'
'        .vsUMVenda = CStr(.vsUMVenda)
'        lErro = Comando_BindVarInt(lComando, .vsUMVenda)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113268
'
'        .vlNumVendas = CLng(.vlNumVendas)
'        lErro = Comando_BindVarInt(lComando, .vlNumVendas)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113282
'
'
'        .vdQuantVendas = CDbl(.vdQuantVendas)
'        lErro = Comando_BindVarInt(lComando, .vdQuantVendas)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113269
'
'        .vdValorVendido = CDbl(.vdValorVendido)
'        lErro = Comando_BindVarInt(lComando, .vdValorVendido)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113270
'
'
'    End With
'
'    RelFlashVendas_Le_Estatisticas_Hora2 = SUCESSO
'
'    Exit Function
'
'Erro_RelFlashVendas_Le_Estatisticas_Hora2:
'
'    RelFlashVendas_Le_Estatisticas_Hora2 = gErr
'
'    Select Case gErr
'
'        Case 113267 To 113270, 113282
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169179)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function RelFlashVendas_Le_Estatisticas_Hora3(ByVal lComando As Long, tRelFlashVendas As TypeRelFlashVendasVar, ByVal objRelFlashVendas As ClassFlashVendas, sSelect As String)
''Função que Bind os filtros passados pelo usuário
'
'Dim lErro As Long
'
'On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora3
'
'
'    tRelFlashVendas.vsProduto = CStr(objRelFlashVendas.sProduto)
'    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vsProduto)
'    If (lErro <> AD_SQL_SUCESSO) Then gError 113271
'
'    tRelFlashVendas.vdHoraIni = CDbl(objRelFlashVendas.dHoraIni)
'    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vdHoraIni)
'    If (lErro <> AD_SQL_SUCESSO) Then gError 113272
'
'    tRelFlashVendas.vdHoraFim = CDbl(objRelFlashVendas.dHoraFim)
'    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vdHoraFim)
'    If (lErro <> AD_SQL_SUCESSO) Then gError 113273
'
'    tRelFlashVendas.vdtData = CDate(objRelFlashVendas.dtData)
'    lErro = Comando_BindVarInt(lComando, tRelFlashVendas.vdtData)
'    If (lErro <> AD_SQL_SUCESSO) Then gError 113274
'
'
'    'Verifica se o filtro é por filial empresa
'    If objRelFlashVendas.iFilialEmpresa <> EMPRESA_TODA Then
'
'        tRelFlashVendas.viFilialEmpresa = CInt(giFilialEmpresa)
'        lErro = Comando_BindVarInt(lComando, tRelFlashVendas.viFilialEmpresa)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113275
'
'
'    End If
'
'    'Verifica se o filtro também será por Caixa
'    If objRelFlashVendas.iCaixaIni <> CODIGO_NAO_PREENCHIDO Then
'
'        tRelFlashVendas.viCaixaIni = CInt(objRelFlashVendas.iCaixaIni)
'        lErro = Comando_BindVarInt(lComando, tRelFlashVendas.viCaixaIni)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113276
'
'    End If
'
'    'Verifica se o filtro também será por Caixa
'    If objRelFlashVendas.iCaixaFim <> CODIGO_NAO_PREENCHIDO Then
'
'        tRelFlashVendas.viCaixaFim = CInt(objRelFlashVendas.iCaixaFim)
'        lErro = Comando_BindVarInt(lComando, tRelFlashVendas.viCaixaFim)
'        If (lErro <> AD_SQL_SUCESSO) Then gError 113277
'
'    End If
'
'    lErro = Comando_ExecutarInt(lComando)
'    If (lErro <> AD_SQL_SUCESSO) Then gError 113278
'
'    RelFlashVendas_Le_Estatisticas_Hora3 = SUCESSO
'
'    Exit Function
'
'Erro_RelFlashVendas_Le_Estatisticas_Hora3:
'
'    RelFlashVendas_Le_Estatisticas_Hora3 = gErr
'
'    Select Case gErr
'
'        Case 113271 To 113278
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169180)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function RelFlashVendas_Le_Estatisticas_Hora4(ByVal lComando As Long, tRelFlashVendas As TypeRelFlashVendasVar, ByVal objRelFlashVendas As ClassFlashVendas, sSelect As String) As Long
''Busca no Banco de dados os Calculos Referentes a cada Produto vendido entre um determinado período de tempo
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim dFatorConv As Double
'
'On Error GoTo Erro_RelFlashVendas_Le_Estatisticas_Hora4
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113279
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 113280
'
'    'Atribui o Codigo do Produto em Questão para a Leitura no Banco de Dados
'    objProduto.sCodigo = objRelFlashVendas.sProduto
'
'    'Se Encontrou então Verifica qual é a unidade de venda do produto em questão
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO Then gError 113281
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        With objRelFlashVendas
'
'            'acumular o Numero de Vendas
'            .lNumVendas = .lNumVendas + tRelFlashVendas.vlNumVendas
'
'            'Verifica se a unidade de venda do produto é diferente da unidade na Tabela de itens de Cupom
'            If objProduto.sSiglaUMVenda <> tRelFlashVendas.vsUMVenda Then
'                'Usar a Função que converte para unidade de venda
'                lErro = CF("UM_CONVERSAO_REL", objProduto.iClasseUM, tRelFlashVendas.vsUMVenda, objProduto.sSiglaUMVenda, dFatorConv)
'                If lErro <> SUCESSO Then gError 113283
'
'                tRelFlashVendas.vdQuantVendas = tRelFlashVendas.vdQuantVendas * dFatorConv
'
'            End If
'
'            .sProduto = tRelFlashVendas.vsProduto
'            .dQuantVendida = .dQuantVendida + tRelFlashVendas.vdQuantVendas
'            .dValorVendido = .dValorVendido + tRelFlashVendas.vdValorVendido
'
'        End With
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113284
'
'    Loop
'
'    RelFlashVendas_Le_Estatisticas_Hora4 = SUCESSO
'
'    Exit Function
'
'Erro_RelFlashVendas_Le_Estatisticas_Hora4:
'
'    RelFlashVendas_Le_Estatisticas_Hora4 = gErr
'
'    Select Case gErr
'
'        Case 113279, 113284
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSelect)
'
'        Case 113280, 113283
'            'Só desvia o Código sem Msg
'
'        Case 113280
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr, sSelect)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169181)
'
'    End Select
'
'    Exit Function
'
'End Function
