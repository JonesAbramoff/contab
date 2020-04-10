VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AnaliseVendaOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   ScaleHeight     =   6000
   ScaleWidth      =   13995
   Begin VB.CommandButton BotaoAnalisar 
      Caption         =   "Análise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12765
      TabIndex        =   29
      Top             =   5565
      Width           =   1110
   End
   Begin VB.ComboBox Tabelapreco 
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   180
      Width           =   3900
   End
   Begin VB.Frame FrameComparacao 
      Caption         =   "Comparação"
      Height          =   435
      Left            =   6660
      TabIndex        =   26
      Top             =   90
      Visible         =   0   'False
      Width           =   5730
      Begin VB.OptionButton OptPrecoBase 
         Caption         =   "Preço Base"
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
         Left            =   3135
         TabIndex        =   28
         Top             =   165
         Width           =   1575
      End
      Begin VB.OptionButton OptPrecoUnitario 
         Caption         =   "Preço Unitário"
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
         Left            =   1170
         TabIndex        =   27
         Top             =   165
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   13200
      ScaleHeight     =   450
      ScaleWidth      =   645
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   705
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   120
         Picture         =   "AnaliseVendaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Itens"
      Height          =   4890
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   13890
      Begin MSMask.MaskEdBox PrecoTabela 
         Height          =   225
         Left            =   7620
         TabIndex        =   21
         Top             =   1485
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   397
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "AnaliseVendaOcx.ctx":017E
         Left            =   6675
         List            =   "AnaliseVendaOcx.ctx":0180
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1485
         Width           =   600
      End
      Begin MSMask.MaskEdBox DifPrecoTabela 
         Height          =   225
         Left            =   8610
         TabIndex        =   19
         Top             =   2865
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   4155
         TabIndex        =   17
         Top             =   1830
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PrecoTotalTabela 
         Height          =   225
         Left            =   5715
         TabIndex        =   18
         Top             =   3150
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.TextBox DescricaoProduto 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4005
         MaxLength       =   250
         TabIndex        =   11
         Top             =   900
         Width           =   2400
      End
      Begin MSMask.MaskEdBox PrecoTotalB 
         Height          =   225
         Left            =   2820
         TabIndex        =   10
         Top             =   2985
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   225
         Left            =   900
         TabIndex        =   12
         Top             =   2130
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercentDesc 
         Height          =   225
         Left            =   4800
         TabIndex        =   13
         Top             =   2325
         Width           =   700
         _ExtentX        =   1244
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PrecoUnitario 
         Height          =   225
         Left            =   6330
         TabIndex        =   14
         Top             =   2055
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   397
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   3465
         TabIndex        =   15
         Top             =   2190
         Width           =   700
         _ExtentX        =   1244
         _ExtentY        =   397
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
      Begin MSMask.MaskEdBox PrecoTotal 
         Height          =   225
         Left            =   8850
         TabIndex        =   16
         Top             =   2385
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   1455
         Left            =   30
         TabIndex        =   1
         Top             =   240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   21
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.Label LabelValorTotalNoDoc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1530
      TabIndex        =   25
      Top             =   5550
      Width           =   1485
   End
   Begin VB.Label LabelValorTotalPelaTabela 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4725
      TabIndex        =   24
      Top             =   5535
      Width           =   1485
   End
   Begin VB.Label Label5 
      Caption         =   "Total na Venda:"
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
      Left            =   90
      TabIndex        =   23
      Top             =   5580
      Width           =   1530
   End
   Begin VB.Label Label4 
      Caption         =   "Total pela tabela:"
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
      Left            =   3090
      TabIndex        =   22
      Top             =   5580
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "Diferença em (%):"
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
      Left            =   9495
      TabIndex        =   7
      Top             =   5610
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Diferença em $:"
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
      Left            =   6300
      TabIndex        =   6
      Top             =   5595
      Width           =   1635
   End
   Begin VB.Label LabelDifTotalPerc 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   11130
      TabIndex        =   5
      Top             =   5565
      Width           =   1485
   End
   Begin VB.Label LabelDifTotalValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7950
      TabIndex        =   4
      Top             =   5565
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Tabela para comparação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   300
      TabIndex        =   3
      Top             =   225
      Width           =   2280
   End
End
Attribute VB_Name = "AnaliseVendaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjAnaliseVenda As ClassAnaliseVendaInfo

Dim giTabPrecoAnalise As Integer
Dim iAlterado As Integer
Dim gbPrecoBaseDif As Boolean
Dim giPrecoUnitario As Integer

'Grid Itens
Dim objGridItens As AdmGrid
Dim iGrid_ItemProduto_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_PercDesc_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_PrecoTotal_Col As Integer
Dim iGrid_PrecoTotalB_Col As Integer
Dim iGrid_PrecoTabela_Col As Integer
Dim iGrid_PrecoTotalTabela_Col As Integer
Dim iGrid_DifPrecoTabela_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Análise de Venda"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "AnaliseVenda"

End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Private Sub BotaoAnalisar_Click()
    Call CF("AnaliseVenda_ValidaDescMax", gobjAnaliseVenda)
End Sub

Private Sub OptPrecoBase_Click()
    If giPrecoUnitario <> OptPrecoUnitario.Value Then
        Call Carrega_Grid_Itens(gobjAnaliseVenda)
        Call Trata_TabelaPreco
        giPrecoUnitario = OptPrecoUnitario.Value
    End If
End Sub

Private Sub OptPrecoUnitario_Click()
    If giPrecoUnitario <> OptPrecoUnitario.Value Then
        Call Carrega_Grid_Itens(gobjAnaliseVenda)
        Call Trata_TabelaPreco
        giPrecoUnitario = OptPrecoUnitario.Value
    End If
End Sub

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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing
    Set gobjAnaliseVenda = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201516)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call Carrega_TabelaPreco

    Set objGridItens = New AdmGrid
    
    If gobjCRFAT.iValidaDescMaxTabPreco = 0 Then BotaoAnalisar.Visible = False
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201517)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(ByVal objAnaliseVenda As ClassAnaliseVendaInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gobjAnaliseVenda = objAnaliseVenda
    
    gbPrecoBaseDif = False
    giPrecoUnitario = OptPrecoUnitario.Value

    Call Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Traz_Analise_Tela(gobjAnaliseVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 201518)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function Traz_Analise_Tela(ByVal objAnaliseVenda As ClassAnaliseVendaInfo) As Long
    
Dim iTabelaPreco As Integer

    Call Carrega_Grid_Itens(objAnaliseVenda)

    iTabelaPreco = objAnaliseVenda.iTabelaPreco
    If giTabPrecoAnalise <> 0 Then iTabelaPreco = giTabPrecoAnalise
    
    'Se a tabela de preços estiver preenchida coloca na tela
    If iTabelaPreco > 0 Then
        TabelaPreco.Text = iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        TabelaPreco.Text = ""
    End If

End Function

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Tipo = 2")
    If lErro <> SUCESSO Then gError 84012

    For Each objCodDescricao In colCodigoDescricao

        giTabPrecoAnalise = objCodDescricao.iCodigo
        
        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next

    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Tipo = 0")
    If lErro <> SUCESSO Then gError 84012

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next
    
    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 84012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157722)

    End Select

    Exit Function

End Function

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 84243 '26806

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 84243

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157727)

    End Select

    Exit Sub

End Sub

Public Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Erro_GridItens_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 84145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157746)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If


End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub PercentDesc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoTotal_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Public Sub PrecoTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PrecoTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TabelaPreco_Click()

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    If TabelaPreco.ListIndex = -1 Then Exit Sub

    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 84013

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 84013

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157768)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84014

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 84015 '26539

        If lErro <> SUCESSO Then gError 84016 '26540 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 84017 '30527

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 84018 '26541

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 84014, 84015, 84017

    Case 84016  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        End If

    Case 84018
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157769)

    End Select

    Exit Sub

End Sub

Public Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Trata_TabelaPreco

    For iLinha = 1 To objGridItens.iLinhasExistentes

        lErro = Trata_TabelaPreco_Item(iLinha)
        If lErro <> SUCESSO Then gError 84019

    Next

    'Calcula o Valor Total da Nota
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84020

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 84019 'tratado na rotina chamada

        Case 84020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157774)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer, dQtde As Double, dPrecoTotal As Double

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 84021 '39147

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 84022 '39148

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 84023 '39149

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 84024 '39150
            
            'Coloca no Grid
            GridItens.TextMatrix(iLinha, iGrid_PrecoTabela_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            dQtde = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
            GridItens.TextMatrix(iLinha, iGrid_PrecoTotalTabela_Col) = Format((dPrecoUnitario * dQtde), "Standard")
            GridItens.TextMatrix(iLinha, iGrid_DifPrecoTabela_Col) = Format(IIf(dPrecoUnitario <> 0, dPrecoTotal - (dPrecoUnitario * dQtde), 0), "Standard")
            
            'Calcula o Preco Total do Ítem
            Call PrecoTotal_Calcula(iLinha)
            
         End If

    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 84021, 84022, 84023, 84024, ERRO_SEM_MENSAGEM 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157775)

    End Select

    Exit Function

End Function

Function ValorTotal_Calcula() As Long
'

Dim lErro As Long, iLinha As Integer
Dim dTotalDoc As Double, dTotalPelaTabela As Double
Dim dValorTabela As Double

On Error GoTo Erro_ValorTotal_Calcula

    dTotalDoc = 0
    dTotalPelaTabela = 0
    
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        dValorTabela = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotalTabela_Col))
        
        If dValorTabela <> 0 Then
        
            dTotalPelaTabela = dTotalPelaTabela + dValorTabela
            dTotalDoc = dTotalDoc + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
            
        End If
    
    Next
    
    If dTotalPelaTabela <> 0 Then
    
        LabelValorTotalNoDoc.Caption = Format(dTotalDoc, "Standard")
        LabelValorTotalPelaTabela.Caption = Format(dTotalPelaTabela, "Standard")
        LabelDifTotalPerc.Caption = Format(100 * (dTotalDoc - dTotalPelaTabela) / dTotalPelaTabela, "Standard")
        LabelDifTotalValor.Caption = Format((dTotalDoc - dTotalPelaTabela), "Standard")
    
    Else
    
        LabelValorTotalNoDoc.Caption = ""
        LabelValorTotalPelaTabela.Caption = ""
        LabelDifTotalPerc.Caption = ""
        LabelDifTotalValor.Caption = ""
    
    End If

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 101102, 101103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157776)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 84133 '26065

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 84134 '26068

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 84132, 84133, 84134

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157785)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Preço Unitário
        Case iGrid_PrecoUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 84139 '26596

        'Se for a de Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 84140 '26599
            
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 84140

        Case iGrid_PrecoTotal_Col
            lErro = Saida_Celula_PrecoTotal(objGridInt)
            If lErro <> SUCESSO Then gError 141390
    
    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 84135, 84136, 84137, 84138, 84139, 84140, 84141, 129216, 129965, 129966, 141410, 141385, 141389, 141390

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157787)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 84147 '31389

    Select Case objControl.Name

        Case PrecoUnitario.Name, PercentDesc.Name, Desconto.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

            
        Case PrecoTotal.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Else
            objControl.Enabled = False
        
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 84148, 84150, 84147, 141412

        Case 84149
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157789)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitario

    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 84170  '26684

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
    
    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84171 '26685

    If Not bPrecoUnitarioIgual Then
    
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84172 '51042

    End If

   Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 84170, 84171, 84172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157794)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim sValorPercAnterior As String

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    If Len(PercentDesc.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(PercentDesc.Text) 'Alterado por Wagner
        If lErro <> SUCESSO Then gError 84329 '26694

        dPercentDesc = CDbl(PercentDesc.Text)

        If Format(dPercentDesc, "#0.#0\%") <> GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 84330 '26695

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
        dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)
    If sValorPercAnterior = "" Then sValorPercAnterior = "0,00%"

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84331 '26696
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then

        'iDescontoAlterado = REGISTRO_ALTERADO
        
        'Recalcula o preço total
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84333 '51044
        
    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 84329, 84331, 84333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84330
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84332

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157795)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto
    'Verifica se o o desconto foi digitado
    If Len(Trim(Desconto.ClipText)) > 0 Then
        
        'Critica o valor digitado
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 42219

        dDesconto = CDbl(Desconto.Text)
        
    End If
        
    If dDesconto <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) Then
        'iDescontoAlterado = REGISTRO_ALTERADO
        dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotalB_Col))
        'Se o Preço total e positivo
        If dPrecoTotal > 0 Then
            'Verifica se o Valor do desconto é superior ao Preço Total
            If dDesconto >= dPrecoTotal Then gError 42220
            
            'Recalcula o percentual de desconto
            dPercentDesc = dDesconto / dPrecoTotal

            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42221

    Call PrecoTotal_Calcula(GridItens.Row)

    Call ValorTotal_Calcula
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 42219, 42221
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42220
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157211)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_Itens(ByVal objAnaliseVenda As ClassAnaliseVendaInfo) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objGridItens1 As Object
Dim objItem As ClassItemAnaliseVenda

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For Each objItem In objAnaliseVenda.colItens

        iIndice = iIndice + 1
        
        lErro = Mascara_RetornaProdutoEnxuto(objItem.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 84406

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItem.sDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItem.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItem.dQuantidade)
        
        If OptPrecoUnitario.Value Then
            GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItem.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            
            'Calcula o percentual de desconto
            If objItem.dPrecoTotal + objItem.dValorDesconto > 0 Then
                dPercDesc = objItem.dValorDesconto / (objItem.dPrecoTotal + objItem.dValorDesconto)
            Else
                dPercDesc = 0
            End If
            
            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
            GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItem.dValorDesconto, "Standard")
            GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItem.dPrecoTotal, "Standard")
            GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItem.dPrecoTotal + objItem.dValorDesconto, "Standard")
        
        Else
            GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItem.dPrecoBase, gobjFAT.sFormatoPrecoUnitario)
            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(0, "Percent")
            GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(0, "Standard")
            GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItem.dPrecoBase * objItem.dQuantidade, "Standard")
            GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItem.dPrecoBase * objItem.dQuantidade, "Standard")
        End If
        
        If Abs(objItem.dPrecoUnitario - objItem.dPrecoBase) > DELTA_VALORMONETARIO Then
            gbPrecoBaseDif = True
        End If
        
    Next
    
    If gbPrecoBaseDif Then
        FrameComparacao.Visible = True
    End If

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItens)

    Carrega_Grid_Itens = SUCESSO
    
    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 84406
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItem.sProduto)

        Case 141404 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157815)

    End Select

    Exit Function

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = Me
    Set objUserControl = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Qtde")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desc.")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Unit. Tabela")
    objGridInt.colColuna.Add ("Total Tabela")
    objGridInt.colColuna.Add ("Dif. Total")

    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (PrecoTabela.Name)
    objGridInt.colCampo.Add (PrecoTotalTabela.Name)
    objGridInt.colCampo.Add (DifPrecoTabela.Name)
    
    'Colunas do Grid
    iIncremento = 0
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1 + iIncremento
    iGrid_DescProduto_Col = 2 + iIncremento
    iGrid_UnidadeMed_Col = 3 + iIncremento
    iGrid_Quantidade_Col = 4 + iIncremento
    iGrid_PrecoUnitario_Col = 5 + iIncremento
    iGrid_PercDesc_Col = 6 + iIncremento
    iGrid_Desconto_Col = 7 + iIncremento
    iGrid_PrecoTotal_Col = 8 + iIncremento
    iGrid_PrecoTotalB_Col = 9 + iIncremento
    iGrid_PrecoTabela_Col = 10 + iIncremento
    iGrid_PrecoTotalTabela_Col = 11 + iIncremento
    iGrid_DifPrecoTabela_Col = 12 + iIncremento
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.iLinhasVisiveis = 12
    objGridInt.objGrid.Rows = gobjAnaliseVenda.colItens.Count + 1
    If objGridInt.objGrid.Rows < (objGridInt.iLinhasVisiveis + 1) Then objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoIncluirNoMeioGrid = GRID_PROIBIDO_INCLUIR_NOMEIO
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function Gravar_Registro() As Long
    Gravar_Registro = SUCESSO
End Function

Public Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim dPercAcresFin As Double, objTabelaPreco As New ClassTabelaPreco
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objCotacao1 As New ClassCotacaoMoeda
Dim objCotacaoAnterior1 As New ClassCotacaoMoeda
Dim objCotacao2 As New ClassCotacaoMoeda
Dim objCotacaoAnterior2 As New ClassCotacaoMoeda
Dim dCotacao1 As Double, dCotacao2 As Double
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 84029 '26638
    If lErro = 28030 Then gError 84030 '26639
    
    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 84031 '26640

    dPrecoUnitario = objTabelaPrecoItem.dPreco * dFator

    If objTabelaPrecoItem.iCodTabela <> 0 Then
    
        objTabelaPreco.iCodigo = objTabelaPrecoItem.iCodTabela
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 84029
        If lErro = SUCESSO And objTabelaPreco.iMoeda <> gobjAnaliseVenda.iMoeda Then
        'Se possuem moedas diferentes ou se trocou a moeda precisa calcular/recalcular a cotação
            If objTabelaPreco.iMoeda <> MOEDA_REAL Then
            
                objCotacao1.dtData = gobjAnaliseVenda.dtDataEmissao
                objCotacao1.iMoeda = objTabelaPreco.iMoeda
                objCotacaoAnterior1.iMoeda = objTabelaPreco.iMoeda
            
                'Chama função de leitura
                lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao1, objCotacaoAnterior1)
                If lErro <> SUCESSO Then gError 84029
                
            Else
                objCotacao1.dValor = 1
            End If
            
            If gobjAnaliseVenda.iMoeda <> MOEDA_REAL Then
            
                objCotacao2.dtData = gobjAnaliseVenda.dtDataEmissao
                objCotacao2.iMoeda = gobjAnaliseVenda.iMoeda
                objCotacaoAnterior2.iMoeda = gobjAnaliseVenda.iMoeda
            
                'Chama função de leitura
                lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao2, objCotacaoAnterior2)
                If lErro <> SUCESSO Then gError 84029
                
            Else
                objCotacao2.dValor = 1
            End If
            
            If objCotacao1.dValor <> 0 Then
                dCotacao1 = StrParaDbl(Format(objCotacao1.dValor, "#.0000"))
            Else
                vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao1.iMoeda, Format(gobjAnaliseVenda.dtDataEmissao, "dd/mm/yyyy"), Format(objCotacaoAnterior1.dValor, "#.0000"), Format(objCotacaoAnterior1.dtData, "dd/mm/yyyy"))
                If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
                dCotacao1 = StrParaDbl(Format(objCotacaoAnterior1.dValor, "#.0000"))
            End If
            
            If objCotacao2.dValor <> 0 Then
                dCotacao2 = StrParaDbl(Format(objCotacao2.dValor, "#.0000"))
            Else
                vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao2.iMoeda, Format(gobjAnaliseVenda.dtDataEmissao, "dd/mm/yyyy"), Format(objCotacaoAnterior2.dValor, "#.0000"), Format(objCotacaoAnterior2.dtData, "dd/mm/yyyy"))
                If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
                dCotacao2 = StrParaDbl(Format(objCotacaoAnterior2.dValor, "#.0000"))
            End If
            
            If dCotacao1 = 0 Then gError 211631
            If dCotacao2 = 0 Then gError 211632
            
            'Se nao existe cotacao para a data informada => Mostra a última.
            dPrecoUnitario = dPrecoUnitario * dCotacao1 / dCotacao2
        
        End If
    
    End If
    
    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 84029, 84031

        Case 84030
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case 211631
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao1.iMoeda)

        Case 211632
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao2.iMoeda)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157777)

    End Select

    Exit Function

End Function

Private Sub PrecoTotal_Calcula(ByVal iLinha As Integer)
'???
Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim dValorDescontoGlobal As Double, dValorDescontoItens As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_PrecoTotal_Calcula

    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
    
    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dPrecoUnitario = 0 Or dQuantidade = 0 Then
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(0, "Standard")
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(0, "Standard")
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(0, "Standard")
    Else
        dPrecoTotal = dPrecoUnitario * dQuantidade
        dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))

        dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, dPercentDesc, dDesconto, dPrecoTotalReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, 0, 0, dPrecoTotalReal)

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotalReal, "Standard")

    End If
    
    GridItens.TextMatrix(iLinha, iGrid_DifPrecoTabela_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) - StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotalTabela_Col)), "Standard")
    
    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr
    
        Case 56883, 184285
        
        Case 132029
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157199)
            
    End Select
            
    Exit Sub

End Sub

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = Arredonda_Moeda(dValorUnitario * dQuantidade)

    'Se o Percentual Desconto estiver preenchido
    If dPercentDesc > 0 Then

        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = dDesconto / dValorTotal
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = Arredonda_Moeda(dPercentDesc * dValorTotal)

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Private Function Saida_Celula_PrecoTotal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Total que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoAnt As Double, dPrecoAtual As Double
Dim dQuantidade As Double, dPrecoUnit As Double
Dim dPrecoBrutoCalc As Double, dDesconto As Double
Dim bPrecoTotalIgual As Boolean, dPrecoTotal As Double, dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PrecoTotal

    bPrecoTotalIgual = False

    Set objGridInt.objControle = PrecoTotal

    If Len(Trim(PrecoTotal.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoTotal.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    dPrecoAnt = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    dPrecoUnit = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
    dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
    dPrecoAtual = StrParaDbl(PrecoTotal.Text)
    
    dPrecoBrutoCalc = Arredonda_Moeda(dQuantidade * dPrecoUnit)

    'Comparação com Preço Unitário anterior
    If Abs(dPrecoAnt - dPrecoAtual) < DELTA_VALORMONETARIO Then
        bPrecoTotalIgual = True
    Else
        bPrecoTotalIgual = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Not bPrecoTotalIgual Then
        
        If dPrecoBrutoCalc < dPrecoAtual Then
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = Format(dPrecoAtual / dQuantidade, gobjFAT.sFormatoPrecoUnitario)
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = ""
            
        Else
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dPrecoBrutoCalc - dPrecoAtual, "STANDARD")
            
            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotalB_Col))
            dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
            
            'Se o Preço total é positivo
            If dPrecoTotal > 0 Then
                'Recalcula o percentual de desconto
                dPercentDesc = dDesconto / dPrecoTotal
                GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
            End If
        End If
        
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

   Saida_Celula_PrecoTotal = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoTotal:

    Saida_Celula_PrecoTotal = gErr


    Select Case gErr

        Case 84170, 84171, 84172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157794)

    End Select

    Exit Function

End Function

