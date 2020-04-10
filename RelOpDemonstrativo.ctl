VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpDemonstrativoOcx 
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   4440
   ScaleWidth      =   6000
   Begin VB.CheckBox Recalculo 
      Caption         =   "Recalcular o relatório para o mês em questão"
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
      Left            =   210
      TabIndex        =   7
      Top             =   4080
      Value           =   1  'Checked
      Width           =   4275
   End
   Begin VB.TextBox Codigo 
      Height          =   315
      Left            =   1650
      MaxLength       =   10
      TabIndex        =   1
      Top             =   960
      Width           =   1950
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   735
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   5625
      Begin MSMask.MaskEdBox MesFinal 
         Height          =   315
         Left            =   2430
         TabIndex        =   4
         Top             =   270
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AnoFinal 
         Height          =   315
         Left            =   2940
         TabIndex        =   5
         Top             =   270
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelReqProdDataFinal 
         AutoSize        =   -1  'True
         Caption         =   "Mês/Ano:"
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
         Left            =   1530
         TabIndex        =   22
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2820
         TabIndex        =   21
         Top             =   270
         Width           =   90
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
      Height          =   570
      Left            =   4140
      Picture         =   "RelOpDemonstrativo.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemonstrativo.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemonstrativo.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemonstrativo.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemonstrativo.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CheckBox Consolidado 
      Caption         =   "Incluir Relatório Consolidado"
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
      Left            =   210
      TabIndex        =   6
      Top             =   3780
      Width           =   2835
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      Left            =   870
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   315
      Width           =   1950
   End
   Begin VB.Frame Frame1 
      Caption         =   "Produto"
      Height          =   1245
      Index           =   1
      Left            =   210
      TabIndex        =   13
      Top             =   1530
      Width           =   5625
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   300
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoAte 
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   300
         Width           =   3420
      End
      Begin VB.Label ProdutoLabelDe 
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.Label DescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Top             =   720
         Width           =   3420
      End
      Begin VB.Label ProdutoLabelAte 
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   780
         Width           =   360
      End
   End
   Begin VB.Label LabelCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Previsão Venda:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   23
      Top             =   1020
      Width           =   1410
   End
   Begin VB.Label Opcao 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   390
      Width           =   630
   End
End
Attribute VB_Name = "RelOpDemonstrativoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

'??? transferir as ctes abaixo p/.bas onde estiverem as definicoes referentes a tabela de almoxarifados
Private Const ALMOX_TIPO_DISPONIVEL = 1
Private Const ALMOX_TIPO_NAOCONFORME = 2
Private Const ALMOX_TIPO_EMRECUP = 3
Private Const ALMOX_TIPO_FORAVALIDADE = 4
Private Const ALMOX_TIPO_ELABORACAO = 5
Private Const ALMOX_TIPO_SOBRAS = 6

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Relatório Demonstrativo"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpDemonstrativo"

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

'*********************************************************
'
'  Sergio Ricardo Pacheco da Vitoria
'  Inicio dia 29/10/2002 14:05
'  Supervisor : Leonardo
'*********************************************************

Private Sub Form_Load()


Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    Set objEventoCodigo = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 111376

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 111377

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 111376, 111377

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168121)

    End Select

End Sub



Public Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 111378

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 111379

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 111378
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 111379

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168122)

    End Select

    Exit Function

End Function

Private Sub ProdutoLabelDe_Click()


Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabelDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoDe.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 111370

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_ProdutoLabelDe_Click:

    Select Case gErr

        Case 111370

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168123)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 111371

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 111372

    'Função que Verifica no Banco de Dados se o Produto existe, se existir Traz a Descrição
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, DescricaoDe)
    If lErro <> SUCESSO Then gError 111373

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 111371, 111373

        Case 111372
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168124)

    End Select

    Exit Sub
End Sub

Private Sub ProdutoLabelAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabelAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoAte.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 111374

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_ProdutoLabelAte_Click:

    Select Case gErr

        Case 111374

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168125)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 111426

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 111375

    'Função que Verifica no Banco de Dados se o Produto existe, se existir Traz a Descrição
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, DescricaoAte)
    If lErro <> SUCESSO Then gError 111376

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 111376, 111426

        Case 111375
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168126)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoDe_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 103260

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 103261

        DescricaoDe.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 103266

        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 103262

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 103263

        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 103264

        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 103265

    Else

        DescricaoDe.Caption = ""

    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    DescricaoDe.Caption = ""

    Select Case gErr

        Case 103261, 103260

        Case 103266
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoDe.Text)

        Case 103262
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoDe.Text)

        Case 103263
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoDe.Text)

        Case 103264
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoDe.Text)

        Case 103265
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168127)

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
    If lErro <> SUCESSO Then gError 103267

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 103268

        DescricaoAte.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 103269

        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 103270

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 103271

        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 103272

        'Se nao for um produto produzido => Erro
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then gError 103273

    Else

        DescricaoAte.Caption = ""

    End If

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    DescricaoAte.Caption = ""

    Select Case gErr

        Case 103268, 103267

        Case 103269
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoAte.Text)

        Case 103270
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoAte.Text)

        Case 103271
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoAte.Text)

        Case 103272
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoAte.Text)

        Case 103273
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL", gErr, ProdutoAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168128)

    End Select

End Sub

Private Sub LabelReqProdDataFinal_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(LabelReqProdDataFinal, Source, X, Y)

End Sub

Private Sub LabelReqProdDataFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Controle_MouseDown(LabelReqProdDataFinal, Button, Shift, X, Y)

End Sub

Private Sub Opcao_DragDrop(Source As Control, X As Single, Y As Single)

   Call Controle_DragDrop(Opcao, Source, X, Y)

End Sub

Private Sub Opcao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call Controle_MouseDown(Opcao, Button, Shift, X, Y)

End Sub

Private Sub ProdutoLabelDe_DragDrop(Source As Control, X As Single, Y As Single)

   Call Controle_DragDrop(ProdutoLabelDe, Source, X, Y)

End Sub

Private Sub ProdutoLabelDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call Controle_MouseDown(ProdutoLabelDe, Button, Shift, X, Y)

End Sub

Private Sub ProdutoLabelAte_DragDrop(Source As Control, X As Single, Y As Single)

   Call Controle_DragDrop(ProdutoLabelAte, Source, X, Y)

End Sub

Private Sub ProdutoLabelAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Call Controle_MouseDown(ProdutoLabelAte, Button, Shift, X, Y)

End Sub

Private Sub ProdutoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoDe)

End Sub

Private Sub ProdutoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProdutoAte)

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub BotaoLimpar_Click()
'Faz a Limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 111387

    ComboOpcoes.Text = ""
    DescricaoDe.Caption = ""
    DescricaoAte.Caption = ""
    Consolidado.Value = DESMARCADO
    Recalculo.Value = MARCADO

    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 111387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168129)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
'Sai da Tela

    Unload Me

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 111388

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REL_PREV_DEMONSTRATIVO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 111389

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem (ComboOpcoes.ListIndex)

        'limpa as opções da tela
         Call BotaoLimpar_Click

        ComboOpcoes.Text = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 111388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 111389

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168130)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If Len(Trim(ComboOpcoes.Text)) = 0 Then gError 111390

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 111391

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 111392

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 111393

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 111390
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 111391, 111392, 111393

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168131)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoCodigo = Nothing

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    'Formta o Produto
    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 111396

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 111397

    'Inclui o Parâmetro na col de parâmetros
    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 111399

    'Inclui o Parâmetro na col de parâmetros
    lErro = objRelOpcoes.IncluirParametro("TCODPREVISAO", Codigo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 111493

    'Inclui o Parâmetro na col de parâmetros
    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 111400

    'Verifica se a Check esta marcada
    If Consolidado.Value = MARCADO Then

        'Inclui o Parâmetro na col de parâmetros
        lErro = objRelOpcoes.IncluirParametro("NRELCONSOLIDADO", MARCADO)
        If lErro <> AD_BOOL_TRUE Then gError 111403

    Else

        'Inclui o Parâmetro na col de parâmetros
        lErro = objRelOpcoes.IncluirParametro("NRELCONSOLIDADO", DESMARCADO)
        If lErro <> AD_BOOL_TRUE Then gError 111404

    End If

    'Verifica se a Check de recalculo está  marcada
    If Recalculo.Value = MARCADO Then

        'Inclui o Parâmetro na col de parâmetros
        lErro = objRelOpcoes.IncluirParametro("NRELRECALCULO", MARCADO)
        If lErro <> AD_BOOL_TRUE Then gError 111485

    Else

        'Inclui o Parâmetro na col de parâmetros
        lErro = objRelOpcoes.IncluirParametro("NRELRECALCULO", DESMARCADO)
        If lErro <> AD_BOOL_TRUE Then gError 111486

    End If

    'Passa o Ano ser incluido na Col de Parâmetros
    lErro = objRelOpcoes.IncluirParametro("NANO", StrParaInt(AnoFinal.Text))
    If lErro <> AD_BOOL_TRUE Then gError 111405

    lErro = objRelOpcoes.IncluirParametro("NMES", StrParaInt(MesFinal.Text))
    If lErro <> AD_BOOL_TRUE Then gError 111406

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 111494

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 111396 To 111406, 111485, 111486, 111493, 111494

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168132)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "PRODUTO >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PRODUTO <= " & Forprint_ConvTexto(sProd_F)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168133)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 111407

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 111408

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 111409

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 111407
            ProdutoDe.SetFocus

        Case 111408
            ProdutoAte.SetFocus

        Case 111409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168134)

    End Select

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    'Função que lê no Banco de dados o Codigo do Relatorio e Traz a Coleção de parâmetro carregados
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 111411

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 111413

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, DescricaoDe)
    If lErro <> SUCESSO Then gError 111414

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 111415

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, DescricaoAte)
    If lErro <> SUCESSO Then gError 111416

    'Verifica se o relatório é Consolidado
    lErro = objRelOpcoes.ObterParametro("NRELCONSOLIDADO", sParam)
    If lErro <> SUCESSO Then gError 111418

    Consolidado.Value = StrParaInt(sParam)

    'Verifica se é para recriar o arquivo base do relatorio
    lErro = objRelOpcoes.ObterParametro("NRELRECALCULO", sParam)
    If lErro <> SUCESSO Then gError 111487

    Recalculo.Value = StrParaInt(sParam)

    'Pega o ano em questão relacionado ao relório
    lErro = objRelOpcoes.ObterParametro("NANO", sParam)
    If lErro <> SUCESSO Then gError 111419
    AnoFinal.PromptInclude = False
    AnoFinal.Text = sParam
    AnoFinal.PromptInclude = True

    'pega a DataFinal e exibe e valida
    lErro = objRelOpcoes.ObterParametro("NMES", sParam)
    If lErro <> SUCESSO Then gError 111420
    MesFinal.PromptInclude = False
    MesFinal.Text = sParam
    MesFinal.PromptInclude = True

    'preenche a previsao de vendas
    lErro = objRelOpcoes.ObterParametro("TCODPREVISAO", sParam)
    If lErro <> SUCESSO Then gError 111495
    Codigo.Text = sParam

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 111411, 111413 To 111416, 111418, 111419, 111420, 111487, 111495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168135)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim sProdutoFormatadoDe As String
Dim sProdutoFormatadoATE As String
Dim iProdutoPreenchidoDe As Integer
Dim iProdutoPreenchidoATE As Integer
Dim iRecalculo As Integer
Dim sCodPrevisao As String
Dim iResultado As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objConsRecal As New AdmRelatorio, sSelecaoAux As String

On Error GoTo Erro_BotaoExecutar_Click

    'Se o Mes não estiver preenchido, Erro.
    If Len(Trim(MesFinal.Text)) = 0 Then gError 111422

    'Se o Ano não estiver preenchido, Erro.
    If Len(Trim(AnoFinal.Text)) = 0 Then gError 111421

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 111424

    'guarda selecao p/relatorio resumido
    sSelecaoAux = gobjRelOpcoes.sSelecao

    'Verifica se esta marcado o recalculo
    If Recalculo.Value = MARCADO Then
        iRecalculo = MARCADO
    Else
        iRecalculo = DESMARCADO
    End If

   'Pega o código da Previsão
    sCodPrevisao = Codigo.Text

    If Recalculo.Value = MARCADO Then

        vbMsgRes = vbYes

        'Se mês-ano atual é diferente do mês-ano do relatório
        If StrParaInt(MesFinal.Text) <> Month(gdtDataHoje) Or StrParaInt(AnoFinal.Text) <> Year(gdtDataHoje) Then

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_REL_PREV_DEMONSTRATIVO", StrParaInt(MesFinal.Text))

        End If

        If vbMsgRes = vbYes Then

            GL_objMDIForm.MousePointer = vbHourglass

            'Função que lê as Unidades de venda e Estoque e Classe de Unidade para os produtos compreendidos do ProdutoDE ao ProdutoATE para o mes e o ano em questão
            lErro = RelDemonstrativo_CriarArqRel(StrParaInt(AnoFinal.Text), StrParaInt(MesFinal.Text), iRecalculo, sCodPrevisao, giFilialEmpresa)
            If lErro <> SUCESSO Then gError 111410

        End If

    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    If Consolidado.Value = MARCADO Then

        'formata o Produto Inicial
        lErro = CF("Produto_Formata", ProdutoDe.Text, sProdutoFormatadoDe, iProdutoPreenchidoDe)
        If lErro <> SUCESSO Then gError 111407

        If iProdutoPreenchidoDe <> PRODUTO_PREENCHIDO Then sProdutoFormatadoDe = ""

        'formata o Produto Final
        lErro = CF("Produto_Formata", ProdutoAte.Text, sProdutoFormatadoATE, iProdutoPreenchidoATE)
        If lErro <> SUCESSO Then gError 111408

        If iProdutoPreenchidoATE <> PRODUTO_PREENCHIDO Then sProdutoFormatadoATE = ""

        'Executa o relatório consolidado
        Call objConsRecal.ExecutarDireto("Demonstrativo de Produção Resumido", sSelecaoAux, 1, "RELDEMRI", "TPRODINIC", sProdutoFormatadoDe, "TPRODFIM", sProdutoFormatadoATE, "NANO", AnoFinal.Text, "NMES", MesFinal.Text)

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExecutar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 111407, 111408, 111410

'        Case 111425
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
'            Codigo.SetFocus

        Case 111421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", gErr)
            AnoFinal.SetFocus

        Case 111422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", gErr)
            MesFinal.SetFocus

        Case 111424

        Case 111467

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168136)

    End Select

    Exit Sub

End Sub

Private Sub MesFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(MesFinal)

End Sub

Private Sub AnoFinal_GotFocus()

    Call MaskEdBox_TrataGotFocus(AnoFinal)

End Sub

Private Sub MesFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesFinal_Validate

    'Se o mes foi preenchido
    If Len(Trim(MesFinal.ClipText)) > 0 Then

        'Tem que estar entre 1 e 12 (Jan - Dez)
        If StrParaInt(MesFinal.Text) < 1 Or StrParaInt(MesFinal.Text) > 12 Then gError 111445

    End If

    Exit Sub

Erro_MesFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 111445
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168137)

    End Select

End Sub

Private Sub AnoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AnoFinal_Validate

    'Se o ano estiver preenchido ...
    If Len(Trim(AnoFinal.Text)) > 0 Then

        'Se o ano nao tiver 4 dígitos => Erro
        If Len(Trim(AnoFinal.Text)) <> 4 Then gError 111446

    End If

    Exit Sub

Erro_AnoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 111446
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO_CORRETAMENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168138)

    End Select

End Sub


'****** Inicio das Funções de Leitura faturamento do produto ************
'**
'** Sergio Ricardo dia 31/10/2002 Hora 19:40 ******

Function RelDemonstrativo_CriarArqRel(ByVal iAno As Integer, ByVal iMes As Integer, ByVal iRecalculo As Integer, ByVal sCodPrevisao As String, ByVal iFilialEmpresa As Integer) As Long
'Dispara os cálculos necessários para o Relatório Demonstrativo
'Inicia o processo de gravação em RelDemonstrativo

Dim lErro As Long
Dim lComando As Long
Dim sCodProd As String
Dim iClassUM As Integer
Dim sSiglaUMEstoque As String
Dim sSiglaUMVenda As String
Dim objRelDemonstrativo As ClassRelDemonstrativo
Dim colRelDemonstrativo As Collection, bObtevePrevVenda As Boolean
Dim sDescricao As String, dtDataInicial As Date

On Error GoTo Erro_RelDemonstrativo_CriarArqRel

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 111432

    '** Inicilaza as Constantes
    sCodProd = String(STRING_PRODUTO, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    sSiglaUMVenda = String(STRING_UM_SIGLA, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    'seleciona os Produtos
    lErro = Comando_Executar(lComando, "SELECT Descricao, Codigo , SiglaUMVenda , SiglaUMEstoque, ClasseUM FROM Produtos, ProdutosFilial WHERE ProdutosFilial.FilialEmpresa = ? AND Produtos.Codigo = ProdutosFilial.Produto AND Gerencial = 0 AND Compras = 0 Order by Codigo", sDescricao, sCodProd, sSiglaUMVenda, sSiglaUMEstoque, iClassUM, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 111433

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111434

    Set colRelDemonstrativo = New Collection

    'Verifica enquanto for sucesso
    Do While lErro = SUCESSO

        Set objRelDemonstrativo = New ClassRelDemonstrativo

        objRelDemonstrativo.iFilialEmpresa = iFilialEmpresa
        objRelDemonstrativo.sProduto = sCodProd
        objRelDemonstrativo.sUMEstoque = sSiglaUMEstoque
        objRelDemonstrativo.sUMVenda = sSiglaUMVenda
        objRelDemonstrativo.iClasseUM = iClassUM
        objRelDemonstrativo.sCodigoPrevVenda = sCodPrevisao
        objRelDemonstrativo.sDescricao = sDescricao
        objRelDemonstrativo.iMes = iMes
        objRelDemonstrativo.iAno = iAno

        'obtem qtd a ser faturada referente aos pedidos de venda em aberto
        lErro = RelDemonstrativo_ObterQtdPedVenda(objRelDemonstrativo)
        If lErro <> SUCESSO Then gError 106753
        
        'Obtém a data inicial do produto no estoque e as quantidades disponíveis
        lErro = RelDemonstrativo_ObterQTD(objRelDemonstrativo, dtDataInicial)
        If lErro <> SUCESSO Then gError 106531

        'Função Calcula a Media das quantidade faturadas do Produto em Questão passada por parâmetro em um período de 12 meses(passados)
        lErro = RelDemonstrativo_ObterVendaMedia(objRelDemonstrativo, dtDataInicial)
        If lErro <> SUCESSO And lErro <> 111439 Then gError 111449

        'Função que verifica os Acúmlos da Quantide de produção de um Determinado Produto em um período do mês em questão
        lErro = RelDemonstrativo_ObterDadosProd(objRelDemonstrativo)
        If lErro <> SUCESSO Then gError 111450

        'Função que retorna a qtde em pedidos do Produto para entrega no mes
        lErro = RelDemonstrativo_ObterProd_Programacao_Vendas(objRelDemonstrativo)
        If lErro <> SUCESSO Then gError 111468

        'Função que retorna a Quantidade em K's de tudo que foi faturado no mes em questão
        lErro = RelDemonstrativo_ObterQuant_Faturada(objRelDemonstrativo)
        If lErro <> SUCESSO Then gError 111456

        bObtevePrevVenda = False

        'Verifica a previsão de consumo para o mes e ano passado por parâmetro
        lErro = RelDemonstrativo_PrevisaoConsumo(objRelDemonstrativo, bObtevePrevVenda)
        If lErro <> SUCESSO Then gError 111488

        If bObtevePrevVenda = False Then

            'Verifica a quantidade de produtos faturados para de um determinado produto para o mes em questão
            lErro = RelDemonstrativo_PrevisaoVendas(objRelDemonstrativo)
            If lErro <> SUCESSO Then gError 111469

        End If

        colRelDemonstrativo.Add objRelDemonstrativo

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 111435

    Loop

    'fecha comando
    Call Comando_Fechar(lComando)

    'Função que grava os calculos de relatório para um determinado Produto
    lErro = RelDemonstrativo_Grava(colRelDemonstrativo, StrParaInt(AnoFinal.Text), StrParaInt(MesFinal.Text), objRelDemonstrativo.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 111467

    RelDemonstrativo_CriarArqRel = SUCESSO

    Exit Function

Erro_RelDemonstrativo_CriarArqRel:

    RelDemonstrativo_CriarArqRel = gErr

    Select Case gErr

        Case 111432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111433 To 111435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)

        Case 111449 To 111450, 111456, 111468, 111469, 106531, 106753

        Case 111467

        Case 111488

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168139)

    End Select

    'fecha comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDemonstrativo_PrevisaoConsumo(ByVal objRelDemonstrativo As ClassRelDemonstrativo, bObtevePrevVenda As Boolean) As Long
'Obtém o que será necessário para a produção de um determinado produto(só produtos que entram na composção de outros)

Dim lErro As Long
Dim lComando As Long
Dim dQuantPrevInsumo As Double, dQuantPrevVenda As Double

On Error GoTo Erro_RelDemonstrativo_PrevisaoConsumo

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 111489

    'Verifica Atraves da Chave se o produto entra na composição de outro atraves do campo selecionado na tabela QuantPrevInsumo
    lErro = Comando_Executar(lComando, "SELECT QuantPrevInsumo, QuantPrevVenda FROM PrevVendaPrevConsumo WHERE FilialEmpresa = ? AND CodigoPrevVenda = ? AND Produto = ? AND Ano = ? AND Mes = ? ", dQuantPrevInsumo, dQuantPrevVenda, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.sCodigoPrevVenda, objRelDemonstrativo.sProduto, objRelDemonstrativo.iAno, objRelDemonstrativo.iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 111490

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111491

    If lErro = AD_SQL_SEM_DADOS Then

        objRelDemonstrativo.dConsumoInterno = 0
        bObtevePrevVenda = False

    Else

        objRelDemonstrativo.dConsumoInterno = Round(dQuantPrevInsumo - dQuantPrevVenda, 4)
        objRelDemonstrativo.dQuantidadePrevVenda = dQuantPrevVenda
        bObtevePrevVenda = True

    End If

    'fecha comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_PrevisaoConsumo = SUCESSO

    Exit Function

Erro_RelDemonstrativo_PrevisaoConsumo:

    RelDemonstrativo_PrevisaoConsumo = gErr

    Select Case gErr

        Case 111489
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111490 To 111491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDA_PREVCONSUMO", gErr)

        Case 111492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168140)

    End Select

    'fecha comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDemonstrativo_Grava(ByVal colRelDemonstrativo As Collection, ByVal iAno As Integer, ByVal iMes As Integer, ByVal iFilialEmpresa As Integer) As Long
'Função de Gravação para o Relatório em questão

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim objRelDemonstrativo As ClassRelDemonstrativo

On Error GoTo Erro_RelDemonstrativo_Grava

     'abre os comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 111472

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 111473

    'Função que Exclui todos os registros da Tabela RelDemosntrativo para o mes e ano passado por parâmetro
    lErro = RelDemonstrativo_Exclui_EmTrans(iAno, iMes, iFilialEmpresa)
    If lErro <> SUCESSO Then gError 111475

    For Each objRelDemonstrativo In colRelDemonstrativo

        'insere na tabela
        With objRelDemonstrativo
            lErro = Comando_Executar(lComando, _
                "INSERT INTO RelDemonstrativo(Produto,Mes,Ano,FilialEmpresa,Descricao,UMEstoque,ConsumoInterno, VendaMedia , " & _
                "ProducaoAcumulada, ProducaoDiaria , EstoqueAtual , QuantidadeFaturada , ProgramacaoVendas,QuantidadePrevVenda, QtdPedVenda , Data , CodigoPrevVenda, EstoqueNaoConforme, EstoqueEmRecup, EstoqueForaValidade, EstoqueElaboracao) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .sProduto, iMes, iAno, .iFilialEmpresa, .sDescricao, .sUMEstoque, .dConsumoInterno, .dVendaMedia, _
                .dProducaoAcumulada, .dProdAnterior, .dEstoqueAtual, .dQuantidadeFaturada, .dProgramacaoVendas, .dQuantidadePrevVenda, .dQtdPedVenda, gdtDataHoje, .sCodigoPrevVenda, .dEstoqueNaoConforme, .dEstoqueEmRecup, .dEstoqueForaValidade, .dEstoqueElaboracao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 111476

    Next

    'confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 111474

    'fecha o comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_Grava = SUCESSO

    Exit Function

Erro_RelDemonstrativo_Grava:

    RelDemonstrativo_Grava = gErr

    Select Case gErr

        Case 111472
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111473
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 111474
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 111475

        Case 111476
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELDEMONSTRATIVO", gErr, objRelDemonstrativo.sProduto, iMes, iAno)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168141)

    End Select

    'fecha comando
    Call Comando_Fechar(lComando)

    'cancela a transacao
    Call Transacao_Rollback

    Exit Function

End Function

Function RelDemonstrativo_Exclui_EmTrans(ByVal iAno As Integer, ByVal iMes As Integer, ByVal iFilialEmpresa As Integer) As Long
'Exclui todos os registros em RelDemonstrativo cujo Ano=iAno, Mes=iMes e FilialEmpresa=iFilialEmpresa

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dConsumoInterno As Double

On Error GoTo Erro_RelDemonstrativo_Exclui_EmTrans

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 111477

    Next

    'seleciona os parâmtros do relatórios com os parâmetros passados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ConsumoInterno FROM RelDemonstrativo WHERE Mes = ? AND Ano = ? AND FilialEmpresa = ?", 0, dConsumoInterno, iMes, iAno, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 111478

    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111479

    Do While lErro = SUCESSO

'        'locka o registro
'        lErro = Comando_LockExclusive(alComando(0))
'        If lErro <> AD_SQL_SUCESSO Then gError 111481

        'deleta os Calculos do relatório para o mes em questão
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RelDemonstrativo", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 111482

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111483

    Loop

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelDemonstrativo_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_RelDemonstrativo_Exclui_EmTrans:

    RelDemonstrativo_Exclui_EmTrans = gErr

    Select Case gErr

        Case 111477
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111478, 111479, 111483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDEMONSTRATIVO", gErr)

        Case 111481
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELDEMONSTRATIVO", gErr)

        Case 111482
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDEMONSTRATIVO", gErr, iMes, iAno)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168142)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RelDemonstrativo_ObterVendaMedia(ByVal objRelDemonstrativo As ClassRelDemonstrativo, ByVal dtDataInicial As Date) As Long
'Função que obtem o valor medio faturado de um determinado Produto num período de até 12 meses anteriores ao do demonstrativo,
'comecando a partir do mes de dtDataInicial

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iCont As Integer
Dim tSldMesFat As typeSldMesFat, tSldMesFatAux As typeSldMesFat
Dim iIndice As Integer
Dim iContAux As Integer
Dim dtDataAux As Date
Dim iMesIni As Integer, dQuantidadeFaturada As Double

On Error GoTo Erro_RelDemonstrativo_ObterVendaMedia

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 103773

    Next

    objRelDemonstrativo.dVendaMedia = 0

    If objRelDemonstrativo.iMes = 1 Then

        With objRelDemonstrativo

            dtDataAux = StrParaDate("01/01/" & .iAno)

        End With

        If dtDataInicial < dtDataAux Then

            With tSldMesFat

                'seleciona a quantidade faturada de um produto com o ano e filialempresa (passados)
                lErro = Comando_Executar(alComando(1), "SELECT  QuantFaturada1, QuantFaturada2, QuantFaturada3, QuantFaturada4, QuantFaturada5, QuantFaturada6, QuantFaturada7, QuantFaturada8, QuantFaturada9, QuantFaturada10, QuantFaturada11, QuantFaturada12 FROM SldMesFat WHERE Produto = ? AND Ano = ? AND FilialEmpresa = ? ", _
                .adQuantFaturada(1), .adQuantFaturada(2), .adQuantFaturada(3), .adQuantFaturada(4), .adQuantFaturada(5), .adQuantFaturada(6), .adQuantFaturada(7), .adQuantFaturada(8), .adQuantFaturada(9), .adQuantFaturada(10), .adQuantFaturada(11), .adQuantFaturada(12), objRelDemonstrativo.sProduto, objRelDemonstrativo.iAno - 1, objRelDemonstrativo.iFilialEmpresa)
                If lErro <> AD_SQL_SUCESSO Then gError 111496

            End With

            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111497
            If lErro = AD_SQL_SUCESSO Then

                'Se o produto foi cadastrado em ano anterior ao anterior ao do relatorio
                If Year(dtDataInicial) < (objRelDemonstrativo.iAno - 1) Then

                    iMesIni = 1

                Else

                    iMesIni = Month(dtDataInicial)

                End If

                For iCont = iMesIni To 12

                    'Armazena a Quantidade do produto em estoque
                    dQuantidadeFaturada = dQuantidadeFaturada + tSldMesFat.adQuantFaturada(iCont)
                    iContAux = iContAux + 1

                Next

            End If

        End If

    'Se o produto foi cadastrado no mesmo ano do relatório
    ElseIf objRelDemonstrativo.iAno = Year(dtDataInicial) Then

        'Se o produto foi cadastrado anteriormente ao mes do relatório
        If objRelDemonstrativo.iMes > Month(dtDataInicial) Then

            iMesIni = Month(dtDataInicial)

            With tSldMesFat

                'seleciona a quantidade faturada de um produto com o ano e filialempresa (passados)
                lErro = Comando_Executar(alComando(1), "SELECT  QuantFaturada1, QuantFaturada2, QuantFaturada3, QuantFaturada4, QuantFaturada5, QuantFaturada6, QuantFaturada7, QuantFaturada8, QuantFaturada9, QuantFaturada10, QuantFaturada11, QuantFaturada12 FROM SldMesFat WHERE Produto = ? AND Ano = ? AND FilialEmpresa = ? ", _
                .adQuantFaturada(1), .adQuantFaturada(2), .adQuantFaturada(3), .adQuantFaturada(4), .adQuantFaturada(5), .adQuantFaturada(6), .adQuantFaturada(7), .adQuantFaturada(8), .adQuantFaturada(9), .adQuantFaturada(10), .adQuantFaturada(11), .adQuantFaturada(12), objRelDemonstrativo.sProduto, objRelDemonstrativo.iAno, objRelDemonstrativo.iFilialEmpresa)
                If lErro <> AD_SQL_SUCESSO Then gError 111496

            End With

            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111497
            If lErro = AD_SQL_SUCESSO Then

                For iCont = iMesIni To objRelDemonstrativo.iMes - 1

                    'Armazena a Quantidade do produto em estoque
                    dQuantidadeFaturada = dQuantidadeFaturada + tSldMesFat.adQuantFaturada(iCont)
                    iContAux = iContAux + 1

                Next

            End If

        End If

    Else 'produto cadastrado em ano anterior ao do relatorio

        With tSldMesFat

            'seleciona a quantidade faturada de um produto com o ano anterior + filialempresa (passados)
            lErro = Comando_Executar(alComando(1), "SELECT  QuantFaturada1 , QuantFaturada2 , QuantFaturada3 , QuantFaturada4 , QuantFaturada5 ,QuantFaturada6 , QuantFaturada7 , QuantFaturada8 , QuantFaturada9 , QuantFaturada10 , QuantFaturada11 , QuantFaturada12 FROM SldMesFat WHERE Produto = ? AND Ano = ? AND FilialEmpresa = ? ", _
            .adQuantFaturada(1), .adQuantFaturada(2), .adQuantFaturada(3), .adQuantFaturada(4), .adQuantFaturada(5), .adQuantFaturada(6), .adQuantFaturada(7), .adQuantFaturada(8), .adQuantFaturada(9), .adQuantFaturada(10), .adQuantFaturada(11), .adQuantFaturada(12), objRelDemonstrativo.sProduto, objRelDemonstrativo.iAno - 1, objRelDemonstrativo.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 111498

        End With

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111499
        If lErro = AD_SQL_SUCESSO Then

            With objRelDemonstrativo

                dtDataAux = StrParaDate("01/" & .iMes & "/" & .iAno)
                dtDataAux = DateAdd("yyyy", -1, dtDataAux)

            End With

            If dtDataAux < dtDataInicial Then
                iCont = Month(dtDataInicial)
            Else
                iCont = Month(dtDataAux)
            End If

            Do While iCont <= 12

                dQuantidadeFaturada = dQuantidadeFaturada + tSldMesFat.adQuantFaturada(iCont)
                iCont = iCont + 1
                iContAux = iContAux + 1

            Loop

        End If

        With tSldMesFatAux

            'seleciona a quantidade faturada de um produto com o ano + filialempresa (passados)
            lErro = Comando_Executar(alComando(2), "SELECT  QuantFaturada1 , QuantFaturada2 , QuantFaturada3 , QuantFaturada4 , QuantFaturada5 ,QuantFaturada6 , QuantFaturada7 , QuantFaturada8 , QuantFaturada9 , QuantFaturada10 , QuantFaturada11 , QuantFaturada12 FROM SldMesFat WHERE Produto = ? AND Ano = ? AND FilialEmpresa = ? ", _
            .adQuantFaturada(1), .adQuantFaturada(2), .adQuantFaturada(3), .adQuantFaturada(4), .adQuantFaturada(5), .adQuantFaturada(6), .adQuantFaturada(7), .adQuantFaturada(8), .adQuantFaturada(9), .adQuantFaturada(10), .adQuantFaturada(11), .adQuantFaturada(12), objRelDemonstrativo.sProduto, objRelDemonstrativo.iAno, objRelDemonstrativo.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 113001

        End With

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113002
        If lErro = AD_SQL_SUCESSO Then

            iCont = objRelDemonstrativo.iMes - 1

            Do While iCont >= 1

                dQuantidadeFaturada = dQuantidadeFaturada + tSldMesFatAux.adQuantFaturada(iCont)
                iCont = iCont - 1
                iContAux = iContAux + 1

            Loop

        End If

    End If

    'Calcula a média das quantidades faturadas para o produto
    If iContAux <> 0 Then
        objRelDemonstrativo.dVendaMedia = dQuantidadeFaturada / iContAux
    Else
        objRelDemonstrativo.dVendaMedia = 0
    End If

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        Call Comando_Fechar(alComando(iIndice))

    Next

    RelDemonstrativo_ObterVendaMedia = SUCESSO

    Exit Function

Erro_RelDemonstrativo_ObterVendaMedia:

    RelDemonstrativo_ObterVendaMedia = gErr

    Select Case gErr

        Case 111373
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111439

        Case 111496 To 111497
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr, objRelDemonstrativo.iAno, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.sProduto)

        Case 113001 To 113002
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr, objRelDemonstrativo.iAno, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168143)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        Call Comando_Fechar(alComando(iIndice))

    Next

    Exit Function

End Function

Function RelDemonstrativo_ObterDadosProd(ByVal objRelDemonstrativo As ClassRelDemonstrativo) As Long
'Acumula o total produzido do produto no mês em questão e no dia anteriror

Dim lErro As Long
Dim lComando As Long
Dim sCodProd As String
Dim sSiglaUM As String
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim dtDataMovto As Date
Dim dFator As Double
Dim dQuantidade As Double
Dim iTipoMovEst As Integer
Dim sProduto As String, dtDiaAnterior As Date
On Error GoTo Erro_RelDemonstrativo_ObterDadosProd

    dtDiaAnterior = DateAdd("d", -1, gdtDataAtual)
    
    'abre o Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 111440

    'Inicilaza a string
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    sProduto = String(20, 0)
    'Monta a data
    dtDataDe = StrParaDate(1 & "/" & objRelDemonstrativo.iMes & "/" & objRelDemonstrativo.iAno)
    dtDataAte = DateAdd("m", 1, dtDataDe)

    'Seleciona os movimentos de entrada de material produzido do produto, descartando os movimentos de estorno
    lErro = Comando_Executar(lComando, "SELECT Produto, Data, SiglaUm, Quantidade , TipoMov FROM MovimentoEstoque WHERE Data >= ? AND Data < ? AND Produto = ? AND FilialEmpresa = ? AND TipoMov IN(?,?) AND NumIntDocEst = 0", sProduto, dtDataMovto, sSiglaUM, dQuantidade, iTipoMovEst, dtDataDe, dtDataAte, objRelDemonstrativo.sProduto, objRelDemonstrativo.iFilialEmpresa, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3)
    If lErro <> AD_SQL_SUCESSO Then gError 111441

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111442

    Do While lErro = AD_SQL_SUCESSO

        'Obtém o fator de conversão entre a Unidade de medida atual para unidade de Estoque
        lErro = CF("UM_Conversao_Trans", objRelDemonstrativo.iClasseUM, sSiglaUM, objRelDemonstrativo.sUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 111443

        'Acumula a Quantidade de produção do produto para o mês em questão
        objRelDemonstrativo.dProducaoAcumulada = objRelDemonstrativo.dProducaoAcumulada + (dQuantidade * dFator)

        'Verifica se a data de Movimento de Estoque de produção é do dia anterior se for acumula
        If dtDataMovto = dtDiaAnterior Then

            objRelDemonstrativo.dProdAnterior = objRelDemonstrativo.dProdAnterior + (dQuantidade * dFator)

        End If

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 111444

    Loop

    'fecha Comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_ObterDadosProd = SUCESSO

    Exit Function

Erro_RelDemonstrativo_ObterDadosProd:

    RelDemonstrativo_ObterDadosProd = gErr

    Select Case gErr

        Case 111440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111441, 111442, 111444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 111443

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168144)

    End Select

    'fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDemonstrativo_ObterProd_Programacao_Vendas(ByVal objRelDemonstrativo As ClassRelDemonstrativo) As Long
'Calcula a quantiadade programada de venda para um Produto em um período.

Dim lErro  As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim dQuantidade As Double
Dim sSQL As String
Dim iClasseUM As Integer
Dim sUnidadeMed As String
Dim dtDataEntrega As Date
Dim dQuantidadePedVenda As Double
Dim sUnidadeMedPedVenda As String
Dim dtDataEmissaoPedVenda As Date
Dim iClasseUMPedVenda As Integer
Dim lComando As Long
Dim dFator As Double

On Error GoTo Erro_RelDemonstrativo_ObterProd_Programacao_Vendas

    'abre o Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 111451

    sUnidadeMed = String(STRING_UM_SIGLA, 0)

    dtDataDe = CDate(1 & "/" & objRelDemonstrativo.iMes & "/" & objRelDemonstrativo.iAno)
    dtDataAte = DateAdd("m", 1, dtDataDe)

    'Seleciona todos os ItensDePedidoDeVendas que sejam programação, que não estejam
    'atendidos, onde a data de entrega ou data do pedido esteja contida no período do relatório
    sSQL = "SELECT Quantidade - QuantFaturada - QuantCancelada, ClasseUM, UnidadeMed, ItensPedidoDeVenda.DataEntrega AS Data FROM ItensPedidoDeVenda , PedidosDeVenda " & _
    "WHERE ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND PedidosDeVenda.FilialEmpresa = ? AND " & _
    "PedidosDeVenda.Programacao = ? AND ItensPedidoDeVenda.DataEntrega >= ? AND ItensPedidoDeVenda.DataEntrega < ? AND ItensPedidoDeVenda.Produto = ? " & _
    "AND ItensPedidoDeVenda.Status <> ?" & _
    " UNION SELECT Quantidade - QuantFaturada - QuantCancelada, ClasseUM, UnidadeMed, PedidosDeVenda.DataEmissao AS Data FROM ItensPedidoDeVenda , PedidosDeVenda " & _
    "WHERE ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND PedidosDeVenda.FilialEmpresa = ? AND " & _
    "PedidosDeVenda.Programacao = ? AND ItensPedidoDeVenda.DataEntrega = ? AND PedidosDeVenda.DataEmissao >= ? AND PedidosDeVenda.DataEmissao < ? AND ItensPedidoDeVenda.Produto = ? " & _
    "AND ItensPedidoDeVenda.Status <> ?"

    lErro = Comando_Executar(lComando, sSQL, dQuantidade, iClasseUM, sUnidadeMed, dtDataEntrega, objRelDemonstrativo.iFilialEmpresa, MARCADO, dtDataDe, dtDataAte, objRelDemonstrativo.sProduto, STATUS_BAIXADO, objRelDemonstrativo.iFilialEmpresa, MARCADO, DATA_NULA, dtDataDe, dtDataAte, objRelDemonstrativo.sProduto, STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 111452

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111453

    Do While lErro = SUCESSO

        'Obtém o fator de conversão entre a Unidade de medida atual e unidade de Estoque
        lErro = CF("UM_Conversao_Trans", iClasseUM, sUnidadeMed, objRelDemonstrativo.sUMEstoque, dFator)
         If lErro <> SUCESSO Then gError 111454

        'Acumula a Previsão de venda do produto para o mes em questão
        objRelDemonstrativo.dProgramacaoVendas = objRelDemonstrativo.dProgramacaoVendas + (dQuantidade * dFator)

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 111455

    Loop

    'fecha Comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_ObterProd_Programacao_Vendas = SUCESSO

    Exit Function

Erro_RelDemonstrativo_ObterProd_Programacao_Vendas:

    RelDemonstrativo_ObterProd_Programacao_Vendas = gErr

    Select Case gErr

        Case 111451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111452, 111453, 111455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)

        Case 111454

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168145)

    End Select

    'fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDemonstrativo_ObterQuant_Faturada(ByVal objRelDemonstrativo As ClassRelDemonstrativo) As Long
'Obtém a quantidade faturada do produto no mês do relatório (verificando em SldDiaFat)

Dim lErro As Long
Dim lComando As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim dFator As Double
Dim dQuantidadeFat As Double
Dim dQuantidadeDevolvida As Double

On Error GoTo Erro_RelDemonstrativo_ObterQuant_Faturada

    'abre o Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 111457

    'Monta a data
    dtDataDe = CDate(1 & "/" & objRelDemonstrativo.iMes & "/" & objRelDemonstrativo.iAno)
    dtDataAte = DateAdd("m", 1, dtDataDe)

    'Data , Quantidade , Sigla da Unidade para o Produto em Questão
    lErro = Comando_Executar(lComando, "SELECT QuantFaturada, QuantDevolvida FROM SldDiaFat WHERE Data >= ? AND Data < ? AND Produto = ? AND FilialEmpresa = ? ", dQuantidadeFat, dQuantidadeDevolvida, dtDataDe, dtDataAte, objRelDemonstrativo.sProduto, objRelDemonstrativo.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 111458

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111460

    Do While lErro = SUCESSO

        objRelDemonstrativo.dQuantidadeFaturada = objRelDemonstrativo.dQuantidadeFaturada + dQuantidadeFat - dQuantidadeDevolvida

        'busca o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 111462

    Loop

    'Função que modifica a Unidade de medida para unidade de Estoque
    lErro = CF("UM_Conversao_Trans", objRelDemonstrativo.iClasseUM, objRelDemonstrativo.sUMVenda, objRelDemonstrativo.sUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 111461

    objRelDemonstrativo.dQuantidadeFaturada = objRelDemonstrativo.dQuantidadeFaturada * dFator

    'fecha Comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_ObterQuant_Faturada = SUCESSO

    Exit Function

Erro_RelDemonstrativo_ObterQuant_Faturada:

    RelDemonstrativo_ObterQuant_Faturada = gErr

    Select Case gErr

        Case 111457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111458, 111460, 111462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFAT", gErr, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.sProduto, dtDataDe)

        Case 111461

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168146)

    End Select

    'fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDemonstrativo_PrevisaoVendas(ByVal objRelDemonstrativo As ClassRelDemonstrativo) As Long
'Função que fornece a Previsão de Venda para um determinado Produto

Dim lErro As Long
Dim lComando As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim dFator As Double
Dim dQuantidadePrev As Double

On Error GoTo Erro_RelDemonstrativo_PrevisaoVendas

    'abre o Comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 111463

    'Data , Quantidade , Sigla da Unidade para o Produto em Questão
    lErro = Comando_Executar(lComando, "SELECT SUM(Quantidade" & CStr(objRelDemonstrativo.iMes) & ") FROM PrevVendaMensal WHERE Codigo = ? AND FilialEmpresa = ? AND Ano = ? AND Produto = ? GROUP BY Produto", dQuantidadePrev, objRelDemonstrativo.sCodigoPrevVenda, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.iAno, objRelDemonstrativo.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 111464

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 111465

    'Função que modifica a Unidade de medida para unidade de Estoque
    lErro = CF("UM_Conversao_Trans", objRelDemonstrativo.iClasseUM, objRelDemonstrativo.sUMVenda, objRelDemonstrativo.sUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 111466

    objRelDemonstrativo.dQuantidadePrevVenda = dQuantidadePrev * dFator

    'fecha Comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_PrevisaoVendas = SUCESSO

    Exit Function

Erro_RelDemonstrativo_PrevisaoVendas:

    RelDemonstrativo_PrevisaoVendas = gErr

    Select Case gErr

        Case 111463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 111464, 111465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDA", gErr, objRelDemonstrativo.sCodigoPrevVenda)

        Case 111466

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168147)

    End Select

    'fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objPrevVenda As New ClassPrevVendaMensal
Dim colSelecao As Collection

    If Len(Trim(Codigo.Text)) > 0 Then objPrevVenda.sCodigo = CStr(Codigo.Text)

    'Chama a Tela que Lista as PrevVendas
    Call Chama_Tela("PrevVMensalCodLista", colSelecao, objPrevVenda, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPrevVenda As ClassPrevVendaMensal

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPrevVenda = obj1

    Codigo.Text = objPrevVenda.sCodigo
    Call Codigo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168148)

    End Select

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPrevVenda As New ClassPrevVenda

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then

        'Verifica se a previsão existe
        objPrevVenda.sCodigo = Codigo.Text
        lErro = CF("PrevVenda_Le2", objPrevVenda)
        If lErro <> SUCESSO And lErro <> 108663 Then gError 111470

        'Se não encontrou => Erro
        If lErro <> SUCESSO Then gError 111471

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 111470

        Case 111471
            Call Rotina_Erro(vbOKOnly, "ERRO_PREVVENDA_NAO_CADASTRADA", gErr, Codigo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168149)

    End Select

End Sub

Public Function RelDemonstrativo_ObterQTD(ByVal objRelDemonstrativo As ClassRelDemonstrativo, dtDataInicial As Date) As Long
'Obtém a quantidade atual total (proprio, reservado ou nao, e de 3os na empresa) de um produto em uma filial empresa
'com separacao por tipo de almoxarifado e a data de cadastramento de estoque do produto na filial.dtDataInicial

Dim lErro As Long
Dim lComando As Long
Dim sProduto As String
Dim dQuantidade As Double, iTipoAlmox As Integer
Dim sComando_SQL As String, dtDataInicialAlmox As Date

On Error GoTo Erro_RelDemonstrativo_ObterQTD

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103274

    dtDataInicial = DATA_NULA
    With objRelDemonstrativo
        .dEstoqueNaoConforme = 0
        .dEstoqueEmRecup = 0
        .dEstoqueForaValidade = 0
        .dEstoqueElaboracao = 0
    End With
    sComando_SQL = "SELECT Almoxarifado.Tipo, DataInicial, QuantDispNossa+QuantReservada+QuantConsig3+QuantBenef3+QuantOutras3+QuantConserto3+QuantDemo3 FROM EstoqueProduto, Almoxarifado WHERE Almoxarifado.FilialEmpresa = ? AND EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo"

    lErro = Comando_Executar(lComando, sComando_SQL, iTipoAlmox, dtDataInicialAlmox, dQuantidade, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 103275

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103276

    Do While lErro <> AD_SQL_SEM_DADOS

        Select Case iTipoAlmox

            Case ALMOX_TIPO_NAOCONFORME
                objRelDemonstrativo.dEstoqueNaoConforme = objRelDemonstrativo.dEstoqueNaoConforme + dQuantidade

            Case ALMOX_TIPO_EMRECUP
                objRelDemonstrativo.dEstoqueEmRecup = objRelDemonstrativo.dEstoqueEmRecup + dQuantidade

            Case ALMOX_TIPO_FORAVALIDADE
                objRelDemonstrativo.dEstoqueForaValidade = objRelDemonstrativo.dEstoqueForaValidade + dQuantidade

            '??? ATENCAO, usando sobras em vez de elaboracao
            Case ALMOX_TIPO_SOBRAS
                objRelDemonstrativo.dEstoqueElaboracao = objRelDemonstrativo.dEstoqueElaboracao + dQuantidade

            Case ALMOX_TIPO_DISPONIVEL
                objRelDemonstrativo.dEstoqueAtual = objRelDemonstrativo.dEstoqueAtual + dQuantidade
            
            Case Else
                '???
                
        End Select

        'obtem a menor data inicial dos almoxarifados
        If dtDataInicial = DATA_NULA Or dtDataInicial > dtDataInicialAlmox Then
            dtDataInicial = dtDataInicialAlmox
        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103276

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    RelDemonstrativo_ObterQTD = SUCESSO

    Exit Function

Erro_RelDemonstrativo_ObterQTD:

    RelDemonstrativo_ObterQTD = gErr

    Select Case gErr

        Case 103274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103275, 103276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr)

        Case 103277 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168150)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function RelDemonstrativo_ObterQtdPedVenda(ByVal objRelDemonstrativo As ClassRelDemonstrativo) As Long
'obtem qtde a ser faturada referente aos pedidos de venda. Deve ignorar pedidos marcados como "programacao"

Dim lErro As Long, dFator As Double
Dim lComando As Long
Dim dQuantidade As Double, sUMItemPV As String

On Error GoTo Erro_RelDemonstrativo_ObterQtdPedVenda

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106754

    sUMItemPV = String(STRING_UM_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT UnidadeMed, SUM(Quantidade-QuantCancelada-QuantFaturada) FROM ItensPedidoDeVenda, PedidosDeVenda WHERE ItensPedidoDeVenda.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND ItensPedidoDeVenda.CodPedido = PedidosDeVenda.Codigo AND PedidosDeVenda.Programacao = 0 AND ItensPedidoDeVenda.Status <> ? AND ItensPedidoDeVenda.FilialEmpresa = ? AND Produto = ? GROUP BY UnidadeMed", sUMItemPV, dQuantidade, STATUS_BAIXADO, objRelDemonstrativo.iFilialEmpresa, objRelDemonstrativo.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106755
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106756
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Obtém o fator de conversão entre a Unidade de medida atual para unidade de Estoque
        lErro = CF("UM_Conversao_Trans", objRelDemonstrativo.iClasseUM, sUMItemPV, objRelDemonstrativo.sUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 106757
    
        objRelDemonstrativo.dQtdPedVenda = objRelDemonstrativo.dQtdPedVenda + (dQuantidade * dFator)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106758
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    RelDemonstrativo_ObterQtdPedVenda = SUCESSO
     
    Exit Function
    
Erro_RelDemonstrativo_ObterQtdPedVenda:

    RelDemonstrativo_ObterQtdPedVenda = gErr
     
    Select Case gErr
          
        Case 106757
        
        Case 106755, 106756, 106758
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_QTDPEDVENDA", gErr)
        
        Case 106754
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168151)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
