VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TabelaPrecoItemOcx 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7695
   Begin VB.TextBox Observacao 
      Height          =   900
      Left            =   1380
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4275
      Width           =   6030
   End
   Begin VB.CommandButton BotaoKitVenda 
      Caption         =   "Kits de Venda"
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
      Left            =   1350
      TabIndex        =   30
      Top             =   5355
      Width           =   1620
   End
   Begin VB.CommandButton BotaoFormacaoPreco 
      Caption         =   "Formação de Preço"
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
      Left            =   5610
      TabIndex        =   28
      Top             =   2865
      Width           =   1815
   End
   Begin VB.CommandButton BotaoAtualizaPreco 
      Caption         =   "Atualizar Preço"
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
      Left            =   3675
      TabIndex        =   27
      Top             =   2865
      Width           =   1695
   End
   Begin VB.CheckBox TabelaDefault 
      Caption         =   "Tabela Padrão"
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
      Left            =   1380
      TabIndex        =   26
      Top             =   1290
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tabela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   210
      TabIndex        =   25
      Top             =   60
      Width           =   4860
      Begin VB.CommandButton BotaoImportarItens 
         Caption         =   "Importar"
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
         Left            =   3705
         TabIndex        =   29
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton BotaoEditarTabela 
         Caption         =   "Editar"
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
         Left            =   1275
         TabIndex        =   1
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton BotaoExcluirTabela 
         Caption         =   "Excluir"
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
         Left            =   2490
         TabIndex        =   2
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton BotaoCriarTabela 
         Caption         =   "Criar"
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
         Left            =   60
         TabIndex        =   0
         Top             =   195
         Width           =   1080
      End
   End
   Begin VB.ComboBox DataVigencia 
      Height          =   315
      Left            =   5745
      TabIndex        =   7
      Top             =   3825
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5430
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "TabelaPrecoItemOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TabelaPrecoItemOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TabelaPrecoItemOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TabelaPrecoItemOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Tabela 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   1380
      TabIndex        =   5
      Top             =   2895
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   1665
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PercDesconto 
      Height          =   315
      Left            =   1380
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label PrecoComDesconto 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5745
      TabIndex        =   34
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor com Desconto:"
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
      Left            =   3840
      TabIndex        =   33
      Top             =   3420
      Width           =   1785
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "% Desconto:"
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
      TabIndex        =   32
      Top             =   3390
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observação:"
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
      Left            =   225
      TabIndex        =   31
      Top             =   4305
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data de Vigência:"
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
      Left            =   4095
      TabIndex        =   24
      Top             =   3885
      Width           =   1545
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
      Height          =   165
      Left            =   585
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   23
      Top             =   1710
      Width           =   735
   End
   Begin VB.Label DescricaoProduto 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   1380
      TabIndex        =   22
      Top             =   2130
      Width           =   6060
   End
   Begin VB.Label DescricaoTabela 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2220
      TabIndex        =   21
      Top             =   840
      Width           =   5355
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Valor Filial:"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2925
      Width           =   960
   End
   Begin VB.Label Label9 
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
      Height          =   165
      Left            =   390
      TabIndex        =   19
      Top             =   2175
      Width           =   930
   End
   Begin VB.Label TabelaLabel 
      AutoSize        =   -1  'True
      Caption         =   "Tabela:"
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
      Height          =   210
      Left            =   660
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   870
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor Empresa:"
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
      Height          =   195
      Left            =   30
      TabIndex        =   17
      Top             =   3900
      Width           =   1290
   End
   Begin VB.Label ValorEmpresa 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1380
      TabIndex        =   16
      Top             =   3855
      Width           =   1695
   End
   Begin VB.Label UnidadeMedida 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6000
      TabIndex        =   15
      Top             =   1650
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Unidade:"
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
      Left            =   5115
      TabIndex        =   14
      Top             =   1695
      Width           =   780
   End
End
Attribute VB_Name = "TabelaPrecoItemOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTTabelaPrecoItem
Attribute objCT.VB_VarHelpID = -1

Private Sub Observacao_Change()
    Call objCT.Observacao_Change
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTTabelaPrecoItem
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoAtualizaPreco_Click()
     Call objCT.BotaoAtualizaPreco_Click
End Sub

Private Sub BotaoCriarTabela_Click()
     Call objCT.BotaoCriarTabela_Click
End Sub

Private Sub BotaoEditarTabela_Click()
     Call objCT.BotaoEditarTabela_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoExcluirTabela_Click()
     Call objCT.BotaoExcluirTabela_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoFormacaoPreco_Click()
     Call objCT.BotaoFormacaoPreco_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoImportarItens_Click()
     Call objCT.BotaoImportarItens_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub DataVigencia_Click()
     Call objCT.DataVigencia_Click
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub ProdutoLabel_Click()
     Call objCT.ProdutoLabel_Click
End Sub

Private Sub Tabela_Change()
     Call objCT.Tabela_Change
End Sub

Private Sub Tabela_Click()
     Call objCT.Tabela_Click
End Sub

Private Sub TabelaLabel_Click()
     Call objCT.TabelaLabel_Click
End Sub

Private Sub Valor_Change()
     Call objCT.Valor_Change
End Sub

Private Sub Valor_GotFocus()
     Call objCT.Valor_GotFocus
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
     Call objCT.Valor_Validate(Cancel)
End Sub

Function Trata_Parametros(Optional objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
     Trata_Parametros = objCT.Trata_Parametros(objTabelaPrecoItem)
End Function

Private Sub DataVigencia_Change()
     Call objCT.DataVigencia_Change
End Sub

Private Sub DataVigencia_Validate(Cancel As Boolean)
     Call objCT.DataVigencia_Validate(Cancel)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub
Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub
Private Sub DescricaoProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoProduto, Source, X, Y)
End Sub
Private Sub DescricaoProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoProduto, Button, Shift, X, Y)
End Sub
Private Sub DescricaoTabela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoTabela, Source, X, Y)
End Sub
Private Sub DescricaoTabela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoTabela, Button, Shift, X, Y)
End Sub
Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub
Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub
Private Sub TabelaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TabelaLabel, Source, X, Y)
End Sub
Private Sub TabelaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TabelaLabel, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub ValorEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorEmpresa, Source, X, Y)
End Sub
Private Sub ValorEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorEmpresa, Button, Shift, X, Y)
End Sub
Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub
Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objUserControl = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

'#####################################
'Inserido por Wagner 18/05/2006
Private Sub BotaoKitVenda_Click()
    Call objCT.BotaoKitVenda_Click
End Sub
'#####################################

Private Sub PercDesconto_Change()
    Call objCT.PercDesconto_Change
End Sub

Private Sub PercDesconto_Validate(Cancel As Boolean)
    Call objCT.PercDesconto_Validate(Cancel)
End Sub
