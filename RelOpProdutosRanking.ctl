VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpProdutosRanking 
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   ScaleHeight     =   6660
   ScaleWidth      =   7770
   Begin VB.Frame FrameCategoria 
      Caption         =   "Categoria"
      Height          =   795
      Left            =   240
      TabIndex        =   33
      Top             =   5640
      Width           =   5160
      Begin VB.ComboBox CategoriaProduto 
         Height          =   315
         Left            =   1740
         TabIndex        =   9
         Top             =   300
         Width           =   2820
      End
      Begin VB.Label LabelCategoria 
         Caption         =   "Categoria:"
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
         Height          =   210
         Left            =   750
         TabIndex        =   34
         Top             =   345
         Width           =   930
      End
   End
   Begin VB.Frame FrameTiposProduto 
      Caption         =   "Tipos de Produto"
      Height          =   1290
      Left            =   240
      TabIndex        =   28
      Top             =   4200
      Width           =   5160
      Begin MSMask.MaskEdBox TipoDe 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TipoAte 
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   825
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label TipoAteDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1800
         TabIndex        =   32
         Top             =   825
         Width           =   2970
      End
      Begin VB.Label TipoDeDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1800
         TabIndex        =   31
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label LabelTipoDe 
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
         Height          =   255
         Left            =   465
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   390
         Width           =   360
      End
      Begin VB.Label LabelTipoAte 
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
         Height          =   255
         Left            =   435
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   855
         Width           =   435
      End
   End
   Begin VB.Frame FrameClassificar 
      Caption         =   "Classificar por"
      Height          =   690
      Left            =   240
      TabIndex        =   27
      Top             =   1920
      Width           =   5160
      Begin VB.OptionButton Valor 
         Caption         =   "Valor"
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
         Left            =   3240
         TabIndex        =   4
         Top             =   320
         Width           =   855
      End
      Begin VB.OptionButton Quantidade 
         Caption         =   "Quantidade"
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
         Left            =   480
         TabIndex        =   3
         Top             =   320
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Width           =   5160
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   390
         Width           =   315
      End
      Begin VB.Label ProdutoDescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   24
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label ProdutoDescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   23
         Top             =   825
         Width           =   2970
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpProdutosRanking.ctx":0000
      Left            =   1200
      List            =   "RelOpProdutosRanking.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2670
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   240
      TabIndex        =   16
      Top             =   975
      Width           =   5160
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   2130
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   1170
         TabIndex        =   1
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   4125
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   3165
         TabIndex        =   2
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         Left            =   780
         TabIndex        =   20
         Top             =   300
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
         Left            =   2745
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5452
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   240
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpProdutosRanking.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpProdutosRanking.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpProdutosRanking.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpProdutosRanking.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   12
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
      Height          =   600
      Left            =   5745
      Picture         =   "RelOpProdutosRanking.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1605
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
      Left            =   480
      TabIndex        =   21
      Top             =   390
      Width           =   615
   End
End
Attribute VB_Name = "RelOpProdutosRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Erros para Incluir no Banco de Dados
'ERRO_INSERCAO_PRODUTOS_RANKING = Erro ao Tentar Incluir um registro na tabela Produtos Ranking.

'################### Só para teste ###############

'Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
'Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
'Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'##################################################

Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Dim lNumIntRel As Long

' Browse 's Relacionados a Tela de RelProdutoRanking
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoTipoDe As AdmEvento
Attribute objEventoTipoDe.VB_VarHelpID = -1
Private WithEvents objEventoTipoAte As AdmEvento
Attribute objEventoTipoAte.VB_VarHelpID = -1

Private mvardQuantPercParticip As Double
Private mvardVlrPercParticip As Double

'Type TypeProdutosRankingVar
'
'    vdtDataDe As Variant
'    vdtDataAte As Variant
'    viClassificacao As Variant
'    vsProdutoDe As Variant
'    vsProdutoAte As Variant
'    viTipoProdutoDe As Variant
'    viTipoProdutoAte As Variant
'    vsCategoria As Variant
'    viNumIntRel As Variant
'    vlRanking As Variant
'    vsProduto As Variant
'    vsNomeReduzido As Variant
'    vsItemCategoria As Variant
'    vdQuantidade As Variant
'    vdValor As Variant
'    viFilialEmpresa As Variant
'
'End Type


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Relatório de Ranking de Produtos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpProdutosRanking"

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
 ' Inicio dia 14/11/2002
'Supervisor:    Shirley
'*********************************************************

Private Sub Form_Load()

Dim lErro As Long
Dim colCategoriaProduto As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Form_Load

    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    Set objEventoTipoDe = New AdmEvento
    Set objEventoTipoAte = New AdmEvento


    'Formata o produto com o formato do Banco de Dados
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 113021

    'Formata o produto com o formato do Banco de Dados
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 113022

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 113029

    If lErro = 22542 Then gError 113030

    'preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    Quantidade.Value = MARCADO

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 113021, 113022, 113029, 113030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171871)

    End Select

    Exit Sub

End Sub

Private Sub TipoDe_Validate(Cancel As Boolean)
'Verifica se o valor para o tipo de produto é válido se for Traz a Descrição do Tipo de Produto

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoDe_Validate

    If Len(Trim(TipoDe.Text)) <> 0 Then

        'Critica o valor
        lErro = Inteiro_Critica(StrParaInt(TipoDe.Text))
        If lErro <> SUCESSO Then gError 113023

        objTipoProduto.iTipo = StrParaInt(TipoDe.Text)

        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 113024

        'Se não encontrar --> gerro
        If lErro = 22531 Then gError 113025

        'Preenche o Campo relacionado a Descrição
        TipoDeDescricao.Caption = objTipoProduto.sDescricao

    Else

        TipoDeDescricao.Caption = ""
    
    End If
    
    Exit Sub

Erro_TipoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113023, 113024

        Case 113025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171872)

    End Select

    Exit Sub

End Sub

Private Sub TipoAte_Validate(Cancel As Boolean)
' Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoAte_Validate

    If Len(Trim(TipoAte.Text)) <> 0 Then

        'Critica o valor
        lErro = Inteiro_Critica(TipoAte.Text)
        If lErro <> SUCESSO Then gError 113026

        objTipoProduto.iTipo = StrParaInt(TipoAte.Text)

        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 113027

        'Se não encontrar --> gerro
        If lErro = 22531 Then gError 113028

        'Preenche o Campo relacionado a Descrição
        TipoAteDescricao.Caption = objTipoProduto.sDescricao

    Else

        TipoAteDescricao.Caption = ""

    End If
    
    Exit Sub

Erro_TipoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 113026, 113027

        Case 113028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171873)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)
'Função que valida a categoria selecionada ou digitada pelo usuario

Dim lErro As Long

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(Trim(CategoriaProduto.Text)) > 0 Then

        lErro = Combo_Item_Igual(CategoriaProduto)
        If lErro <> SUCESSO Then gError 113030

    End If

    Exit Sub

Erro_CategoriaProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 113030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171874)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 113031

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171875)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe)

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 113032

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 113032
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171876)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 113140

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 113140
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171877)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 113033

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 113033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171878)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte)

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 113034

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 113034
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171879)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 113035

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 113035
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171880)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 113036

    ComboOpcoes.Text = ""
    ProdutoDescricaoDe.Caption = ""
    ProdutoDescricaoAte.Caption = ""
    CategoriaProduto.Text = ""
    TipoDeDescricao.Caption = ""
    TipoAteDescricao.Caption = ""
    Quantidade.Value = MARCADO

    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 113036

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171881)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 113037

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 113038

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 113038

        Case 113037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171882)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 113039

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 113040

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 113041

    'se a opção de relatório foi gravada em RelatorioOpcoes então adcionar a opção de relatório na comboopções
    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 113039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 113040, 113041

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171883)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String

On Error GoTo Erro_PreenchgerrelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 113044

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 113045

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 113046

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 113047

    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODDE", TipoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113052

    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODATE", TipoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113053

    If Quantidade.Value = MARCADO Then

        lErro = objRelOpcoes.IncluirParametro("NQUANTIDADE", CStr(MARCADO))
        If lErro <> AD_BOOL_TRUE Then gError 113054

        lErro = objRelOpcoes.IncluirParametro("NVALOR", CStr(DESMARCADO))
        If lErro <> AD_BOOL_TRUE Then gError 113055


    Else

        lErro = objRelOpcoes.IncluirParametro("NQUANTIDADE", CStr(DESMARCADO))
        If lErro <> AD_BOOL_TRUE Then gError 113056

        lErro = objRelOpcoes.IncluirParametro("NVALOR", CStr(MARCADO))
        If lErro <> AD_BOOL_TRUE Then gError 113057

    End If

    lErro = objRelOpcoes.IncluirParametro("TCATEGORIA", CategoriaProduto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113058
    
    If Len(Trim(DataDe.ClipText)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113108
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 113108
    End If
    
    
    If Len(Trim(DataAte.ClipText)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113109
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 113109
    
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
    If lErro <> AD_BOOL_TRUE Then gError 115021
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), StrParaInt(TipoDe.Text), StrParaInt(TipoAte.Text), CategoriaProduto.Text)
    If lErro <> SUCESSO Then gError 113048

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 113044 To 113048, 113052 To 113058, 113108 To 113109

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171884)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO Then gError 113049

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 113050

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 113051

    End If

    'data inicial não pode ser maior que a data final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then

         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 113052

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 113049
            ProdutoDe.SetFocus

        Case 113050
            ProdutoAte.SetFocus

        Case 113051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus

      Case 113052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171885)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, dtdataDe As Date, dtdataAte As Date, iTipoDe As Integer, iTipoAte As Integer, sCategoriaProduto As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

'    If sProd_I <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = "Produto >= " & Forprint_ConvTexto(sProd_I)
'
'    End If
'
'    If sProd_F <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)
'
'    End If
'
'
'    If dtdataDe <> DATA_NULA Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = "Data >= " & Forprint_ConvData(dtdataDe)
'
'    End If
'
'    If dtdataAte <> DATA_NULA Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(dtdataAte)
'
'    End If
'
'
'    If iTipoDe <> 0 Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "iTipo <= " & Forprint_ConvInt(iTipoDe)
'
'    End If
'
'    If iTipoAte <> 0 Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Tipo <= " & Forprint_ConvInt(iTipoAte)
'
'    End If


    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171886)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO Then gError 103053

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 103053

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171887)

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
        If lErro <> SUCESSO Then gError 113054

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 113054

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171888)

    End Select

    Exit Sub

End Sub

Private Sub LabelTipoDe_Click()

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabelTipoDe_Click

    If Len(Trim(TipoDe.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = StrParaInt(TipoDe.Text)


    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoDe)

    Exit Sub

Erro_LabelTipoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171889)

    End Select

    Exit Sub

End Sub

Private Sub LabeltipoAte_Click()

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto
Dim colSelecao As Collection

On Error GoTo Erro_LabeltipoAte_Click

    If Len(Trim(TipoAte.Text)) <> 0 Then

        Set objTipoProduto = New ClassTipoDeProduto
        objTipoProduto.iTipo = StrParaInt(TipoAte.Text)

    End If

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoAte)

    Exit Sub

Erro_LabeltipoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171890)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoDe_evSelecao

    Set objTipoProduto = obj1

    TipoDe.Text = objTipoProduto.iTipo
    TipoDeDescricao.Caption = objTipoProduto.sDescricao
    
    Me.Show

    Exit Sub

Erro_objEventoTipoDe_evSelecao:

    Select Case gErr

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171891)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoAte_evSelecao

    Set objTipoProduto = obj1

    TipoAte.Text = objTipoProduto.iTipo
    TipoAteDescricao.Caption = objTipoProduto.sDescricao
    
    Me.Show

    Exit Sub

Erro_objEventoTipoAte_evSelecao:

    Select Case gErr

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171892)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 113055

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 113056

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 113057

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 113055, 113057

        Case 113056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171893)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 113058

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 113059

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 113060

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 113058, 113060

        Case 113059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171894)

    End Select

    Exit Sub

End Sub

'##### Fim tratameto Browse's  ###########

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoTipoDe = Nothing
    Set objEventoTipoAte = Nothing

End Sub

'##### Botão executar #########

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objProdutoRankingTela As New ClassProdutosRankingTela

On Error GoTo Erro_BotaoExecutar_Click

    'Função que move os dados que estão na tela para a memoria
    lErro = Move_Tela_Memoria(objProdutoRankingTela)
    If lErro <> SUCESSO Then gError 113072

    'Preenche a Tela de Produtos Ranking com os Produtos Lidos no Banco de Dados
    lErro = CF("ProdutosRanking_Preenche", lNumIntRel, objProdutoRankingTela)
    If lErro <> SUCESSO Then gError 113065
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 113063

    'Verifica qual é o relatorio q vai ser Impresso se é por percentual(sobre a qdt total) percentual(sobre a valor total) de um determinado produto
    If Quantidade.Value = True Then

        gobjRelatorio.sNomeTsk = "PRRKQT"

    Else
        
        gobjRelatorio.sNomeTsk = "PRRKVL"

    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 113063, 113065, 113072

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171895)

    End Select

    Exit Sub

End Sub



Function Move_Tela_Memoria(objProdutoRankingTela As ClassProdutosRankingTela) As Long
'Função que move os dados que que serão usados como Filtro no Select Dinâmico

Dim lErro As Long
Dim sProd_I As String
Dim sProd_F As String
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Verificação se o relatorio vai ser a nivel de empresa todo ou a nivel d filial
    If giFilialEmpresa = EMPRESA_TODA Then

        objProdutoRankingTela.iFilialEmpresa = EMPRESA_TODA

    Else

        objProdutoRankingTela.iFilialEmpresa = giFilialEmpresa

    End If

    'Verifica se os Filttros por data estão preenchido se estiverem passa para a memória
    If Len(Trim(DataDe.ClipText)) <> 0 Then

        objProdutoRankingTela.dtdataDe = StrParaDate(DataDe.Text)

    Else

        objProdutoRankingTela.dtdataDe = DATA_NULA

    End If

    'Verifica se os Filttros por data estão preenchido se estiverem passa para a memória
    If Len(Trim(DataAte.ClipText)) <> 0 Then

        objProdutoRankingTela.dtdataAte = StrParaDate(DataAte.Text)

    Else

        objProdutoRankingTela.dtdataAte = DATA_NULA

    End If

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoDe.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 113073

    objProdutoRankingTela.sProdutoDe = sProd_I

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 113074

    objProdutoRankingTela.sProdutoAte = sProd_F

    objProdutoRankingTela.iTipoProdutoDe = StrParaInt(TipoDe.Text)
    objProdutoRankingTela.iTipoProdutoAte = StrParaInt(TipoAte.Text)

    objProdutoRankingTela.sCategoria = CategoriaProduto.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 113073, 113074
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171896)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim bFalse As Boolean

On Error GoTo Erro_PreencherParametrosNaTela

    'Função que lê no Banco de dados o Codigo do Relatorio e Traz a Coleção de parâmetro carregados
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 113096

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 113097

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 113098

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 113099

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 1130100

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODDE", sParam)
    If lErro <> SUCESSO Then gError 113101
    TipoDe.Text = sParam
    Call TipoDe_Validate(bFalse)

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODATE", sParam)
    If lErro <> SUCESSO Then gError 113102
    TipoAte.Text = sParam
    Call TipoAte_Validate(bFalse)

    
    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NQUANTIDADE", sParam)
    If lErro <> SUCESSO Then gError 113103
    
    If Len(Trim(sParam)) > 0 Then
        Quantidade.Value = MARCADO
        Valor.Value = DESMARCADO
    Else
        Quantidade.Value = DESMARCADO
        Valor.Value = MARCADO
    End If

    'Verifica se é para recriar o arquivo base do relatorio
    lErro = objRelOpcoes.ObterParametro("TCATEGORIA", sParam)
    If lErro <> SUCESSO Then gError 113104

    For iIndice = 0 To CategoriaProduto.ListCount - 1
        If Trim(sParam) = Trim(CategoriaProduto.List(iIndice)) Then
            CategoriaProduto.ListIndex = iIndice
            Exit For
        End If
    Next

    'DdataDe
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 113105
    Call DateParaMasked(DataDe, CDate(sParam))
    
    'pega a DataFinal e exibe e valida
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 113106
    Call DateParaMasked(DataAte, CDate(sParam))
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 113096 To 113106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171897)

    End Select

    Exit Function

End Function

Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoDe_Validate

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 113110

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 113111

        ProdutoDescricaoDe.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 113112

        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 113113

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 113114

        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 113115

        
    Else

        ProdutoDescricaoDe.Caption = ""

    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    ProdutoDescricaoDe.Caption = ""

    Select Case gErr

        Case 113110, 113111

        Case 113112
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoDe.Text)

        Case 113113
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoDe.Text)

        Case 113114
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoDe.Text)

        Case 113115
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171898)

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
    If lErro <> SUCESSO Then gError 113116

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 113117

        ProdutoDescricaoAte.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 113118

        'se for gerencial => Erro
        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 113119

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 113120

        'Se não controla estoque => Erro
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 113121

        
    Else

        ProdutoDescricaoAte.Caption = ""

    End If

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    ProdutoDescricaoAte.Caption = ""

    Select Case gErr

        Case 113116, 113117

        Case 113118
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoAte.Text)

        Case 113119
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoAte.Text)

        Case 113120
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoAte.Text)

        Case 113121
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171899)

    End Select

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 113123

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELRODUTOS_RANKING")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 113124

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 113123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 113124

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171900)

    End Select

    Exit Sub

End Sub
