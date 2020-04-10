VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpMapaVendas 
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   7770
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      Height          =   735
      Left            =   240
      TabIndex        =   37
      Top             =   1680
      Width           =   5175
      Begin MSMask.MaskEdBox CaixaDe 
         Height          =   315
         Left            =   1185
         TabIndex        =   4
         Top             =   255
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CaixaAte 
         Height          =   315
         Left            =   3150
         TabIndex        =   5
         Top             =   255
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
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
         Left            =   2715
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   315
         Width           =   360
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
         Left            =   795
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   38
         Top             =   315
         Width           =   315
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
      Picture         =   "RelOpMapaVendas.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   945
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5452
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpMapaVendas.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpMapaVendas.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpMapaVendas.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMapaVendas.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   735
      Left            =   240
      TabIndex        =   29
      Top             =   840
      Width           =   5175
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   2130
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
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
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   4125
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   292
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   3165
         TabIndex        =   3
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
         TabIndex        =   33
         Top             =   345
         Width           =   360
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
         TabIndex        =   32
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMapaVendas.ctx":0A96
      Left            =   1080
      List            =   "RelOpMapaVendas.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   270
      Width           =   2670
   End
   Begin VB.Frame FrameProdutos 
      Caption         =   "Produtos"
      Height          =   1290
      Left            =   240
      TabIndex        =   24
      Top             =   2520
      Width           =   5175
      Begin MSMask.MaskEdBox ProdutoDe 
         Height          =   315
         Left            =   510
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   825
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label ProdutoDescricaoAte 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   28
         Top             =   825
         Width           =   2970
      End
      Begin VB.Label ProdutoDescricaoDe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   27
         Top             =   360
         Width           =   2970
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
         TabIndex        =   26
         Top             =   390
         Width           =   315
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
         TabIndex        =   25
         Top             =   870
         Width           =   360
      End
   End
   Begin VB.Frame FrameTiposProduto 
      Caption         =   "Tipos de Produto"
      Height          =   1290
      Left            =   240
      TabIndex        =   19
      Top             =   3915
      Width           =   5175
      Begin MSMask.MaskEdBox TipoDe 
         Height          =   315
         Left            =   840
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   825
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
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
         TabIndex        =   23
         Top             =   855
         Width           =   435
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
         TabIndex        =   22
         Top             =   390
         Width           =   360
      End
      Begin VB.Label TipoDeDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label TipoAteDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Top             =   825
         Width           =   2970
      End
   End
   Begin VB.Frame FrameCategoria 
      Caption         =   "Categoria"
      Height          =   1740
      Left            =   240
      TabIndex        =   0
      Top             =   5310
      Width           =   5175
      Begin VB.ComboBox ItemCatAte 
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   1200
         Width           =   1900
      End
      Begin VB.ComboBox ItemCatDe 
         Height          =   315
         Left            =   480
         TabIndex        =   12
         Top             =   1200
         Width           =   1900
      End
      Begin VB.CheckBox TodasCategorias 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   435
         TabIndex        =   10
         Top             =   320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.ComboBox CategoriaProduto 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   720
         Width           =   2820
      End
      Begin VB.Label LabelItemCatDe 
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
         Left            =   135
         TabIndex        =   41
         Top             =   1260
         Width           =   315
      End
      Begin VB.Label LabelItemCatAte 
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
         Left            =   2730
         TabIndex        =   40
         Top             =   1260
         Width           =   360
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
         Left            =   360
         TabIndex        =   18
         Top             =   765
         Width           =   930
      End
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
      TabIndex        =   36
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMapaVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

' Browse 's Relacionados a Tela de RelProdutoRanking
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1
Private WithEvents objEventoTipoDe As AdmEvento
Attribute objEventoTipoDe.VB_VarHelpID = -1
Private WithEvents objEventoTipoAte As AdmEvento
Attribute objEventoTipoAte.VB_VarHelpID = -1
Private WithEvents objEventoCaixaDe As AdmEvento
Attribute objEventoCaixaDe.VB_VarHelpID = -1
Private WithEvents objEventoCaixaAte As AdmEvento
Attribute objEventoCaixaAte.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_SALDO_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "Relatório Mapa Vendas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpMapaVendas"

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
 ' Inicio dia 14/01/2003
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

    Set objEventoCaixaDe = New AdmEvento
    Set objEventoCaixaAte = New AdmEvento

    'Formata o produto com o formato do Banco de Dados
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoDe)
    If lErro <> SUCESSO Then gError 113341

    'Formata o produto com o formato do Banco de Dados
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAte)
    If lErro <> SUCESSO Then gError 113342

    TodasCategorias.Value = MARCADO
    ItemCatDe.Enabled = False
    ItemCatAte.Enabled = False
    
    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 113343

    If lErro = 22542 Then gError 113344

    'preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 113341, 113342, 113343, 113344

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169914)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    If Len(DataDe.ClipText) > 0 Then

        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError 113345

    End If

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113345

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169915)

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
    If lErro <> SUCESSO Then gError 113346

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 113346
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169916)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 113347

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 113347
            DataDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169917)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    If Len(DataAte.ClipText) > 0 Then

        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError 113348

    End If

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 113348

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169918)

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
    If lErro <> SUCESSO Then gError 113349

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 113349
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169919)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 113350

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 113350
            DataAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169920)

    End Select

    Exit Sub

End Sub

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
        If lErro <> SUCESSO And lErro <> 79405 Then gError 113351

        If lErro = 79405 Then gError 113352
        
    End If

    Exit Sub

Erro_DataDe_Validate:

Cancel = True

    Select Case gErr

        Case 113351

        Case 113352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169921)

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
        If lErro <> SUCESSO And lErro <> 79405 Then gError 113353
    
        If lErro = 79405 Then gError 113354


    End If

    Exit Sub

Erro_CaixaAte_Validate:

Cancel = True

    Select Case gErr

        Case 113353

        Case 113354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169922)

    End Select

    Exit Sub

End Sub

'### Inicio dos Tratamentos de Browser's

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

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169923)

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
    If lErro <> SUCESSO And lErro <> 79405 Then gError 113355

    'Se não achou o Caixa --> erro
    If lErro = 79405 Then gError 113356

    CaixaDe.PromptInclude = False
    CaixaDe.Text = objCaixa.iCodigo
    CaixaDe.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCaixaDe_evSelecao:

    Select Case gErr

        Case 113355

        Case 113356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169924)

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

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169925)

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
    If lErro <> SUCESSO And lErro <> 79405 Then gError 113357

    'Se não achou o Caixa --> erro
    If lErro = 79405 Then gError 113358

    CaixaAte.PromptInclude = False
    CaixaAte.Text = objCaixa.iCodigo
    CaixaAte.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCaixaAte_evSelecao:

    Select Case gErr

        Case 113357

        Case 113358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXA_NAO_CADASTRADO", gErr, objCaixa.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169926)

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
        If lErro <> SUCESSO Then gError 113377

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 113377

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169927)

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
        If lErro <> SUCESSO Then gError 113378

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 113378

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169928)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 113379

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 113380

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 113381

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 113379, 113381

        Case 113380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169929)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 113382

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 113383

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 113384

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 113382, 113384

        Case 113383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169930)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169931)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169932)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169933)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169934)

    End Select

    Exit Sub

End Sub


'### Fim dos Tratamentos de Browser's

Public Sub TodasCategorias_Click()

Dim lErro As Long

On Error GoTo Erro_TodasCategorias_Click

    If TodasCategorias.Value = MARCADO Then
        
        CategoriaProduto.Text = ""
        ItemCatDe.Clear
        ItemCatAte.Clear
        ItemCatDe.Enabled = False
        ItemCatAte.Enabled = False
        
    End If
    
    Exit Sub
    
Erro_TodasCategorias_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169935)

    End Select

    Exit Sub
    
End Sub


Private Sub ProdutoDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoDe_Validate

    If Len(Trim(ProdutoDe.ClipText)) = 0 Then Exit Sub

    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoDe.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 113359

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 113360
        
        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 113361

        ProdutoDescricaoDe.Caption = objProduto.sDescricao

'        'se for gerencial => Erro
'        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 113362

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 113363

'        'Se não controla estoque => Erro
'        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 113364

        
    Else

        ProdutoDescricaoDe.Caption = ""

    End If

    Exit Sub

Erro_ProdutoDe_Validate:

    Cancel = True

    ProdutoDescricaoDe.Caption = ""

    Select Case gErr

        Case 113359, 113360

        Case 113361
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoDe.Text)

'        Case 113362
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoDe.Text)

        Case 113363
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoDe.Text)

        Case 113364
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoDe.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169936)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdFormatado As String
Dim iProdPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoAte_Validate

    If Len(Trim(ProdutoAte.ClipText)) = 0 Then Exit Sub
    
    sProdFormatado = String(STRING_PRODUTO, 0)

    lErro = CF("Produto_Formata", ProdutoAte.Text, sProdFormatado, iProdPreenchido)
    If lErro <> SUCESSO Then gError 113365

    If iProdPreenchido = PRODUTO_PREENCHIDO Then

        objProduto.sCodigo = sProdFormatado

        'verifica se a Produto existe
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 113366

        ProdutoDescricaoAte.Caption = objProduto.sDescricao

        'Se nao Encontrou => Erro
        If lErro = 28030 Then gError 113367

'        'se for gerencial => Erro
'        If objProduto.iGerencial = PRODUTO_GERENCIAL Then gError 113368

        'Se não for ativo => Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 113369

'        'Se não controla estoque => Erro
'        If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 113370

        
    Else

        ProdutoDescricaoAte.Caption = ""

    End If

    Exit Sub

Erro_ProdutoAte_Validate:

    Cancel = True

    ProdutoDescricaoAte.Caption = ""

    Select Case gErr

        Case 113365, 113366

        Case 113367
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ProdutoAte.Text)

'        Case 113368
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, ProdutoAte.Text)

        Case 113369
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, ProdutoAte.Text)

        Case 113370
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_COM_ESTOQUE", gErr, ProdutoAte.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169937)

    End Select
    
    Exit Sub

End Sub

Private Sub ItemCatDe_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ItemCatDe_Click

    If Len(Trim(ItemCatDe.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ItemCatDe)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = CategoriaProduto.Text
            objCategoriaProdutoItem.sItem = ItemCatDe.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError 113432

            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then gError 113433
            
        End If

    End If

    Exit Sub

Erro_ItemCatDe_Click:

    Select Case gErr

        Case 113432
            ItemCatDe.SetFocus

        Case 113433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ItemCatDe.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169938)

    End Select

    Exit Sub

End Sub

Private Sub ItemCatAte_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_ItemCatAte_Click

    If Len(Trim(ItemCatAte.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ItemCatAte)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = CategoriaProduto.Text
            objCategoriaProdutoItem.sItem = ItemCatAte.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError 113434
            
            'Item da Categoria não está cadastrado
            If lErro <> SUCESSO Then gError 113435
        End If

    End If

    Exit Sub

Erro_ItemCatAte_Click:

    Select Case gErr

        Case 113434
            ItemCatAte.SetFocus

        Case 113435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
            ItemCatAte.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 169939)

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
        If lErro <> SUCESSO Then gError 113371

        objTipoProduto.iTipo = StrParaInt(TipoDe.Text)

        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 113372

        'Se não encontrar --> gerro
        If lErro = 22531 Then gError 113373

        'Preenche o Campo relacionado a Descrição
        TipoDeDescricao.Caption = objTipoProduto.sDescricao

    Else

        TipoDeDescricao.Caption = ""
    
    End If
    
    Exit Sub

Erro_TipoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 113371, 113372

        Case 113373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169940)

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
        If lErro <> SUCESSO Then gError 113374

        objTipoProduto.iTipo = StrParaInt(TipoAte.Text)

        'Lê o tipo
        lErro = CF("TipoDeProduto_Le", objTipoProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then gError 113375

        'Se não encontrar --> gerro
        If lErro = 22531 Then gError 113376

        'Preenche o Campo relacionado a Descrição
        TipoAteDescricao.Caption = objTipoProduto.sDescricao

    Else

        TipoAteDescricao.Caption = ""

    End If
    
    Exit Sub

Erro_TipoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 113374, 113375

        Case 113376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", gErr, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169941)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)
'Função que valida a categoria selecionada ou digitada pelo usuario

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(Trim(CategoriaProduto.Text)) > 0 Then

        ItemCatDe.Clear
        ItemCatAte.Clear
        
        'Preenche o objeto com a Categoria
        objCategoriaProduto.sCategoria = CategoriaProduto.Text

        'Lê Categoria De Produto no BD
        lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 22540 Then gError 113385

        If lErro <> SUCESSO Then gError 113389 'Categoria não está cadastrada

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then gError 113390

        'Preenche Item de Categoria de e Ate
        For Each objCategoriaProdutoItem In colCategoria

            ItemCatDe.AddItem objCategoriaProdutoItem.sItem
            ItemCatAte.AddItem objCategoriaProdutoItem.sItem

        Next

        TodasCategorias.Value = DESMARCADO
        ItemCatDe.Enabled = True
        ItemCatAte.Enabled = True

    
    Else
    
        TodasCategorias.Value = MARCADO
        ItemCatDe.Enabled = False
        ItemCatAte.Enabled = False


    End If

    Exit Sub

Erro_CategoriaProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 113385
        
        Case 113389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", gErr)

        Case 113390

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169942)

    End Select

    Exit Sub

End Sub


Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 113386

    ComboOpcoes.Text = ""
    ProdutoDescricaoDe.Caption = ""
    ProdutoDescricaoAte.Caption = ""
    CategoriaProduto.Text = ""
    TipoDeDescricao.Caption = ""
    TipoAteDescricao.Caption = ""
    TodasCategorias.Value = MARCADO
    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 113386

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169943)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 113387

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 113388

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 113388

        Case 113387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169944)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 113393

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 113394

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 113395

    'se a opção de relatório foi gravada em RelatorioOpcoes então adcionar a opção de relatório na comboopções
    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 113393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 113394, 113395

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169945)

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
    If lErro <> SUCESSO Then gError 113396

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 113397

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then gError 113398

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then gError 113399

    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODDE", TipoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113400

    lErro = objRelOpcoes.IncluirParametro("NTIPOPRODATE", TipoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113401

    
    If Len(Trim(DataDe.ClipText)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATADE", DataDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113402
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATADE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 113402
    End If
    
    If Len(Trim(DataAte.ClipText)) <> 0 Then
    
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", DataAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113403
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAATE", CStr(DATA_NULA))
        If lErro <> AD_BOOL_TRUE Then gError 113403
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NCAIXADE", CaixaDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113404

    lErro = objRelOpcoes.IncluirParametro("NCAIXAATE", CaixaAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 113405

    lErro = objRelOpcoes.IncluirParametro("NTODASCAT", CStr(TodasCategorias.Value))
    If lErro <> AD_BOOL_TRUE Then gError 113428


    If TodasCategorias.Value = DESMARCADO Then

        lErro = objRelOpcoes.IncluirParametro("TCATEGORIA", CategoriaProduto.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113406
    
        lErro = objRelOpcoes.IncluirParametro("TITEMDE", ItemCatDe.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113407
    
        lErro = objRelOpcoes.IncluirParametro("TITEMATE", ItemCatAte.Text)
        If lErro <> AD_BOOL_TRUE Then gError 113408
    
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sProd_I, sProd_F, StrParaDate(DataDe.Text), StrParaDate(DataAte.Text), StrParaInt(TipoDe.Text), StrParaInt(TipoAte.Text), StrParaInt(CaixaDe.Text), StrParaInt(CaixaAte.Text), CategoriaProduto.Text, ItemCatDe.Text, ItemCatAte.Text)
    If lErro <> SUCESSO Then gError 113409

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreenchgerrelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 113396 To 113409, 113428

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169946)

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
    If lErro <> SUCESSO Then gError 113410

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoAte.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 113411

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 113412

    End If

    'data inicial não pode ser maior que a data final
    If Len(Trim(DataDe.ClipText)) <> 0 And Len(Trim(DataAte.ClipText)) <> 0 Then

         If StrParaDate(DataDe.Text) > StrParaDate(DataAte.Text) Then gError 113413

    End If

    'a Caixa inicial não pode ser maior que a Caixa final
    If Len(Trim(CaixaDe.Text)) <> 0 And Len(Trim(CaixaAte.Text)) <> 0 Then

         If StrParaInt(CaixaDe.Text) > StrParaInt(CaixaAte.Text) Then gError 113414

    End If

    'O Tipo de Produto inicial não pode ser maior que o Tipo de Produto
    If Len(Trim(TipoDe.Text)) <> 0 And Len(Trim(TipoAte.Text)) <> 0 Then

         If StrParaInt(TipoDe.Text) > StrParaInt(TipoAte.Text) Then gError 113415

    End If

    'Verifica es valores do campo relacionado a itens de categoria estão preenchiodos de estiverem então verifica
    'se o Item de categoria Final é maior que o Inicial
    If Len(Trim(ItemCatDe.Text)) <> 0 And Len(Trim(ItemCatAte.Text)) <> 0 Then
    
        If ItemCatDe.Text > ItemCatAte.Text Then gError 113439
         
    Else
        
        If Len(Trim(ItemCatDe.Text)) = 0 And Len(Trim(ItemCatAte.Text)) = 0 And TodasCategorias.Value = DESMARCADO Then gError 113440
         
    End If
            
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 113410
            ProdutoDe.SetFocus

        Case 113411
            ProdutoAte.SetFocus

        Case 113412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
            ProdutoDe.SetFocus

        Case 113413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus

        Case 113414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAIXADE_MAIOR", gErr)
            CaixaDe.SetFocus

        Case 113415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_INICIAL_MAIOR", gErr)
            TipoDe.SetFocus

        Case 113439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", gErr)
            ItemCatDe.SetFocus
            
        Case 113440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO", gErr)
            ItemCatDe.SetFocus
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169947)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sProd_I As String, sProd_F As String, dtdataDe As Date, dtdataAte As Date, iTipoDe As Integer, iTipoAte As Integer, iCaixaDe As Integer, iCaixaAte As Integer, sCategoriaProduto As String, sItemCatDe As String, sItemCatAte As String) As Long
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


    If dtdataDe <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(dtdataDe)

    End If

    If dtdataAte <> DATA_NULA Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(dtdataAte)

    End If


    If iTipoDe <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Tipo >= " & Forprint_ConvInt(iTipoDe)

    End If

    If iTipoAte <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Tipo <= " & Forprint_ConvInt(iTipoAte)

    End If


    If iCaixaDe <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa >= " & Forprint_ConvInt(iCaixaDe)

    End If

    If iCaixaAte <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Caixa <= " & Forprint_ConvInt(iCaixaAte)

    End If

    If sCategoriaProduto <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Categoria = " & Forprint_ConvTexto(sCategoriaProduto)

    End If
    
    If sItemCatDe <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Item >= " & Forprint_ConvTexto(sItemCatDe)

    End If

    If sItemCatAte <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Item <= " & Forprint_ConvTexto(sItemCatAte)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169948)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    Set objEventoTipoDe = Nothing
    Set objEventoTipoAte = Nothing

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 113416

    'Verifica qual é o relatorio q vai ser Impresso se é por percentual(sobre a qdt total) percentual(sobre a valor total) de um determinado produto
    If TodasCategorias.Value = MARCADO Then

        gobjRelatorio.sNomeTsk = "MPVENDAS"

    Else
        gobjRelatorio.sNomeTsk = "MPVENCAT"

    End If

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 113416

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169949)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO Then gError 113417

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro <> SUCESSO Then gError 113418

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoDe, ProdutoDescricaoDe)
    If lErro <> SUCESSO Then gError 113419

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro <> SUCESSO Then gError 113420

    'Função que Traz do Bd a Descrição do Produto
    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoAte, ProdutoDescricaoAte)
    If lErro <> SUCESSO Then gError 113421

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODDE", sParam)
    If lErro <> SUCESSO Then gError 113422
    TipoDe.Text = sParam
    Call TipoDe_Validate(bFalse)

    'Pega o Parâmetro Inicial do Tipo de produto
    lErro = objRelOpcoes.ObterParametro("NTIPOPRODATE", sParam)
    If lErro <> SUCESSO Then gError 113423
    TipoAte.Text = sParam
    Call TipoAte_Validate(bFalse)

    'Pega o Parâmetro  Data Inicial
    lErro = objRelOpcoes.ObterParametro("DDATADE", sParam)
    If lErro <> SUCESSO Then gError 113424
    Call DateParaMasked(DataDe, CDate(sParam))

    'Pega o Parâmetro data Final
    lErro = objRelOpcoes.ObterParametro("DDATAATE", sParam)
    If lErro <> SUCESSO Then gError 113425
    Call DateParaMasked(DataAte, CDate(sParam))
    
    'Pega o Parâmetro Caixa Inicial
    lErro = objRelOpcoes.ObterParametro("NCAIXADE", sParam)
    If lErro <> SUCESSO Then gError 113426
    CaixaDe.PromptInclude = False
    CaixaDe.Text = sParam
    Call CaixaDe_Validate(bFalse)
    CaixaDe.PromptInclude = True
    
    'Pega o Parâmetro data Final
    lErro = objRelOpcoes.ObterParametro("NCAIXAATE", sParam)
    If lErro <> SUCESSO Then gError 113427
    CaixaAte.PromptInclude = False
    CaixaAte.Text = sParam
    Call CaixaAte_Validate(bFalse)
    CaixaAte.PromptInclude = True
    
    'Pega o Parâmetro data Final
    lErro = objRelOpcoes.ObterParametro("NTODASCAT", sParam)
    If lErro <> SUCESSO Then gError 113429
    TodasCategorias.Value = sParam
    Call TodasCategorias_Click

    If TodasCategorias.Value = DESMARCADO Then

        lErro = objRelOpcoes.ObterParametro("TCATEGORIA", sParam)
        If lErro <> SUCESSO Then gError 113430
        CategoriaProduto.Text = sParam
        Call CategoriaProduto_Validate(bFalse)

        lErro = objRelOpcoes.ObterParametro("TITEMDE", sParam)
        If lErro <> SUCESSO Then gError 113431
        ItemCatDe.Text = sParam
        Call ItemCatDe_Click
        
        lErro = objRelOpcoes.ObterParametro("TITEMATE", sParam)
        If lErro <> SUCESSO Then gError 113436
        ItemCatAte.Text = sParam
        Call ItemCatAte_Click
        
    End If
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 113417 To 113427
        
        Case 113429 To 113431
        
        Case 113436

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169950)

    End Select

    Exit Function

End Function


Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 113437

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELMAPAPRODUTOS")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 113438

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Call BotaoLimpar_Click

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 113437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 113438

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169951)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is CaixaDe Then
            Call LabelCaixaDe_Click

        ElseIf Me.ActiveControl Is CaixaAte Then
            Call LabelCaixaAte_Click
            
        ElseIf Me.ActiveControl Is ProdutoDe Then
            Call LabelProdutoDe_Click
            
        ElseIf Me.ActiveControl Is ProdutoAte Then
            Call LabelProdutoAte_Click

        ElseIf Me.ActiveControl Is TipoDe Then
            Call LabelTipoDe_Click
            
        ElseIf Me.ActiveControl Is TipoAte Then
            Call LabeltipoAte_Click

        End If

    End If

End Sub

