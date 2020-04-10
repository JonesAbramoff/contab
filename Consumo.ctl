VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Consumo 
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   LockControls    =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   8370
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6015
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Consumo.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Consumo.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Consumo.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "Consumo.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Ano 
      Height          =   315
      Left            =   1125
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1215
      Width           =   1215
   End
   Begin VB.TextBox Mes 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   645
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2085
      Width           =   1305
   End
   Begin VB.CommandButton BotaoRecalculo 
      Caption         =   "Recálculo do Consumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3570
      TabIndex        =   5
      Top             =   2730
      Width           =   1350
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   270
      Left            =   1980
      TabIndex        =   3
      Top             =   2100
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
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
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Top             =   255
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridConsumo 
      Height          =   3285
      Left            =   585
      TabIndex        =   4
      Top             =   1740
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   5794
      _Version        =   393216
      Rows            =   13
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSComctlLib.TreeView TvwProduto 
      Height          =   3900
      Left            =   5250
      TabIndex        =   6
      Top             =   1080
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   6879
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   660
      TabIndex        =   17
      Top             =   1245
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   135
      TabIndex        =   16
      Top             =   810
      Width           =   930
   End
   Begin VB.Label Descricao 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1125
      TabIndex        =   15
      Top             =   750
      Width           =   3030
   End
   Begin VB.Label ProdutoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
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
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   300
      Width           =   735
   End
   Begin VB.Label LabelProduto 
      AutoSize        =   -1  'True
      Caption         =   "Produtos"
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
      Left            =   5220
      TabIndex        =   13
      Top             =   840
      Width           =   765
   End
   Begin VB.Label LblUMEstoque 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3345
      TabIndex        =   12
      Top             =   255
      Width           =   825
   End
End
Attribute VB_Name = "Consumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()


'permite ao usuario informar consumos mensais historicos,
'inclusive alterando valores calculados pelo sistema
'estes valores ficam armazenados na tabela SldMesEst

'copiar parte da carga inicial e do tratamento de outras telas
    'carga da treeview de produtos, tratamento da edit de produto, etc

'criar trata_parametros recebendo objProduto
    'se estiver preenchido mostrar o consumo do ano (atual)

'permitir usar a seta p/navegar de um produto p/outro, mantendo o ano

'qdo trocar o ano ou o produto preencher o grid com os consumos mensais do (produto,ano) que estiver selecionado

'manter flags p/controlar se usuario editou alguma coisa
    'se alterou e for sair perguntar se deseja salvar
    
'só permitir alterar o consumo de meses anteriores ao da implantacao do produto

Public Sub Form_Load()

    'se entrar como EMPRESA_TODA nao poderá alterar, apenas consultar

    'carregar a combo de anos a partir de 1996 (c/4 digitos)
    'selecionar o ano atual
    
    'carregar os meses p/o grid (deve ter isto em alguma tela da contabilidade, ver c/ffl)
    
End Sub

Private Sub BotaoRecalculo_Click()
    'dispara a tela de ConsumoRecalculo p/o (produto,ano) selecionados
End Sub

Private Sub BotaoLimpar_Click()
    'coloca zero nos consumos do ano corrente
End Sub

Private Sub BotaoGravar_Click()
    'chamar ProdutoConsumo_Grava()
End Sub

Private Sub BotaoExcluir_Click()
    '??? nao sei se devemos permitir
End Sub

Function ProdutoConsumo_Le(objProduto As ClassProduto, iAno As Integer, colQtdes As Collection) As Long
'le os consumos do p/o (produto,ano) selecionados
'colQtdes conterá as qtdes consumidas para os doze meses, onde o indice é o mes


End Function

Function ProdutoConsumo_Grava_Ano(objProduto As ClassProduto, iAno As Integer, colQtdes As Collection) As Long
'grava os consumos do p/o (produto,ano) selecionados
'colQtdes contem as qtdes consumidas para os doze meses, onde o indice é o mes

    'giFilialEmpresa tem que ser <> EMPRESA_TODA

    'obter a data de implantacao do produto na filial corrente:
        'é a menor "data inicial" em estoque produto p/os almoxarifados desta filial
    
    'se o registro do ano do sldmesest ainda nao existir,
        'cria -lo
    'senao
        'basta alterar os meses os meses até o limite do mes da implantacao do produto na filial

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONSUMO
    Set Form_Load_Ocx = Me
    Caption = "Consumo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Consumo"
    
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
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProduto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
End Sub

Private Sub LblUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMEstoque, Source, X, Y)
End Sub

Private Sub LblUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMEstoque, Button, Shift, X, Y)
End Sub

