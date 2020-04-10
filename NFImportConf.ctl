VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl NFImportConf 
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10095
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Compra"
      Height          =   2055
      Left            =   7005
      TabIndex        =   34
      Top             =   0
      Width           =   3030
      Begin VB.ComboBox FilialCompra 
         Height          =   315
         Left            =   690
         TabIndex        =   1
         Top             =   195
         Width           =   2295
      End
      Begin VB.ListBox PedidosCompra 
         Height          =   1410
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   510
         Width           =   2910
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Index           =   33
         Left            =   210
         TabIndex        =   35
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4080
      Picture         =   "NFImportConf.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6345
      Width           =   885
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5070
      Picture         =   "NFImportConf.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6345
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   2055
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   6870
      Begin VB.TextBox TipoTribDesc 
         BackColor       =   &H8000000F&
         Height          =   280
         Left            =   1650
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   1065
         Width           =   5115
      End
      Begin VB.TextBox CFOPIntDesc 
         BackColor       =   &H8000000F&
         Height          =   280
         Left            =   2400
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   765
         Width           =   4365
      End
      Begin VB.TextBox CFOPDesc 
         BackColor       =   &H8000000F&
         Height          =   280
         Left            =   2400
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   465
         Width           =   4365
      End
      Begin VB.ComboBox TipoNFiscal 
         Height          =   315
         ItemData        =   "NFImportConf.ctx":025C
         Left            =   1155
         List            =   "NFImportConf.ctx":025E
         TabIndex        =   0
         Top             =   1380
         Width           =   5640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trib.:"
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
         Index           =   6
         Left            =   195
         TabIndex        =   50
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label TipoTrib 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1155
         TabIndex        =   49
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label LabelCli 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   435
         TabIndex        =   38
         Top             =   1755
         Width           =   660
      End
      Begin VB.Label CFOPInt 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1155
         TabIndex        =   37
         Top             =   765
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFOP Int.:"
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
         Index           =   4
         Left            =   195
         TabIndex        =   36
         Top             =   825
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Index           =   1
         Left            =   645
         TabIndex        =   33
         Top             =   1425
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CFOP Ext.:"
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
         Index           =   3
         Left            =   165
         TabIndex        =   27
         Top             =   510
         Width           =   945
      End
      Begin VB.Label CFOP 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1155
         TabIndex        =   26
         Top             =   465
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Index           =   2
         Left            =   4800
         TabIndex        =   25
         Top             =   210
         Width           =   510
      End
      Begin VB.Label Valor 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5370
         TabIndex        =   24
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Núm.NF:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   210
         Width           =   735
      End
      Begin VB.Label NumNF 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1155
         TabIndex        =   22
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Index           =   0
         Left            =   2535
         TabIndex        =   21
         Top             =   210
         Width           =   750
      End
      Begin VB.Label DataEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3345
         TabIndex        =   20
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label LabelForn 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   60
         TabIndex        =   42
         Top             =   1755
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Filial 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5370
         TabIndex        =   41
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
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
         Index           =   5
         Left            =   4875
         TabIndex        =   40
         Top             =   1755
         Width           =   465
      End
      Begin VB.Label CliForn 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1155
         TabIndex        =   39
         Top             =   1710
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Itens da Nota Fiscal"
      Height          =   4290
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   2040
      Width           =   9975
      Begin VB.TextBox DescTipoTrib 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4110
         MaxLength       =   50
         TabIndex        =   47
         Top             =   2580
         Width           =   3510
      End
      Begin VB.TextBox DescCFOP 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4155
         MaxLength       =   50
         TabIndex        =   46
         Top             =   390
         Width           =   3510
      End
      Begin VB.TextBox Detalhe 
         BackColor       =   &H8000000F&
         Height          =   555
         Left            =   3945
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   3675
         Width           =   5925
      End
      Begin VB.CommandButton BotaoCFOP 
         Caption         =   "CFOP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1380
         TabIndex        =   5
         Top             =   3675
         Width           =   1245
      End
      Begin VB.CommandButton BotaoTipoTrib 
         Caption         =   "Tipos de Tributação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2655
         TabIndex        =   6
         Top             =   3675
         Width           =   1245
      End
      Begin VB.TextBox CFOPXml 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1230
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2370
         Width           =   1140
      End
      Begin MSMask.MaskEdBox TipoTributacao 
         Height          =   225
         Left            =   3675
         TabIndex        =   31
         Top             =   2235
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NatOpInterna 
         Height          =   225
         Left            =   2910
         TabIndex        =   30
         Top             =   2235
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.ComboBox UnidadeMed 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4020
         TabIndex        =   29
         Text            =   "UnidadeMed"
         Top             =   780
         Width           =   660
      End
      Begin MSMask.MaskEdBox ValorTotal 
         Height          =   225
         Left            =   5190
         TabIndex        =   28
         Top             =   3225
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton BotaoProdutos 
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
         Height          =   540
         Left            =   90
         TabIndex        =   4
         Top             =   3675
         Width           =   1245
      End
      Begin VB.TextBox EANXml 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6885
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2370
         Width           =   1845
      End
      Begin VB.TextBox EANProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6855
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1860
         Width           =   1845
      End
      Begin VB.TextBox DescProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2250
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1365
         Width           =   3510
      End
      Begin MSMask.MaskEdBox ProdutoXml 
         Height          =   225
         Left            =   495
         TabIndex        =   17
         Top             =   1350
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox UnidadeMedXml 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1635
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2670
         Width           =   945
      End
      Begin MSMask.MaskEdBox ValorUnitario 
         Height          =   225
         Left            =   7185
         TabIndex        =   13
         Top             =   975
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   "#,##0.00####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   225
         Left            =   8910
         TabIndex        =   12
         Top             =   990
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin VB.TextBox DescricaoItem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2655
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1875
         Width           =   3510
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3390
         Left            =   60
         TabIndex        =   3
         Top             =   255
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5980
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "NFImportConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridItens As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_ProdutoXml_Col As Integer
Dim iGrid_DescrItem_Col As Integer
Dim iGrid_DescrProd_Col As Integer
Dim iGrid_CFOPXml_Col As Integer
Dim iGrid_CFOP_Col As Integer
Dim iGrid_TipoTrib_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_UnidadeMedXml_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_ValorUnitario_Col As Integer
Dim iGrid_ValorTotal_Col As Integer
Dim iGrid_EANProd_Col As Integer
Dim iGrid_EANXml_Col As Integer
Dim iGrid_DescCFOP_Col As Integer
Dim iGrid_DescTipoTrib_Col As Integer

Private gobjNFiscal As ClassNFiscal
Private gcolPedidoCompra As Collection
Private gcolPedidoCompraSel As Collection
Private iFilialCompraAnterior As Integer

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCFOP As AdmEvento
Attribute objEventoCFOP.VB_VarHelpID = -1
Private WithEvents objEventoTipoTrib As AdmEvento
Attribute objEventoTipoTrib.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Confirmação de dados importados"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFImportConf"

End Function

Public Sub Show()
'???? comentei para nao dar erro nesta tela pq é modal
'    Parent.Show
 '   Parent.SetFocus
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

Private Sub BotaoCancela_Click()
    Unload Me
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

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objGridItens = Nothing

    Set gobjNFiscal = Nothing
    Set gcolPedidoCompra = Nothing
    Set gcolPedidoCompraSel = Nothing
    
    Set objEventoProduto = Nothing
    Set objEventoCFOP = Nothing
    Set objEventoTipoTrib = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211650)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objFiliais As AdmFiliais
Dim colFiliais As New Collection
Dim objConfiguraCOM As New ClassConfiguraCOM

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento
    Set objEventoCFOP = New AdmEvento
    Set objEventoTipoTrib = New AdmEvento
    Set gcolPedidoCompra = New Collection
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Carrega a combo de FiliaisCompra com as Filiais Empresa
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objFiliais In colFiliais
        If objFiliais.iCodFilial <> 0 Then
            FilialCompra.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialCompra.ItemData(FilialCompra.NewIndex) = objFiliais.iCodFilial
        End If
    Next
    
    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If

    Call FilialCompra_Validate(bSGECancelDummy)
    
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211651)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objNFiscal As ClassNFiscal, Optional colPedCom As Collection = Nothing) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjNFiscal = objNFiscal
    Set gcolPedidoCompraSel = colPedCom
    
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Trata_CliForn
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel
    
    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211652)

    End Select

    Exit Function

End Function

Private Function Traz_NFiscal_Tela(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto
Dim objCli As New ClassCliente
Dim objFilCli As New ClassFilialCliente
Dim objForn As New ClassFornecedor
Dim objFilForn As New ClassFilialFornecedor
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo, sNomeTela As String
Dim objNaturezaOpInt As New ClassNaturezaOp
Dim objNaturezaOpExt As New ClassNaturezaOp
Dim objTipoDeTributacao As ClassTipoDeTributacaoMovto
Dim objNaturezaOp As ClassNaturezaOp
Dim objTipoTrib As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Traz_NFiscal_Tela

    objNaturezaOpInt.sCodigo = objNFiscal.sNaturezaOp
    lErro = CF("NaturezaOperacao_Le", objNaturezaOpInt)
    If lErro <> SUCESSO And lErro <> 17958 Then gError ERRO_SEM_MENSAGEM

    objNaturezaOpExt.sCodigo = objNFiscal.objTributacao.sNaturezaOpInterna
    lErro = CF("NaturezaOperacao_Le", objNaturezaOpExt)
    If lErro <> SUCESSO And lErro <> 17958 Then gError ERRO_SEM_MENSAGEM

    objTipoTrib.iTipo = objNFiscal.objTributacao.iTipoTributacao

    lErro = CF("TipoTributacao_Le", objTipoTrib)
    If lErro <> SUCESSO And lErro <> 27259 Then gError ERRO_SEM_MENSAGEM

    DataEmissao.Caption = Format(objNFiscal.dtDataEmissao, "dd/mm/yyyy")
    NumNF.Caption = CStr(objNFiscal.lNumNotaFiscal)
    CFOP.Caption = objNFiscal.sNaturezaOp
    CFOPInt.Caption = objNFiscal.objTributacao.sNaturezaOpInterna
    Valor.Caption = Format(objNFiscal.dValorTotal, "STANDARD")
    CFOPDesc.Text = objNaturezaOpInt.sDescricao
    CFOPIntDesc.Text = objNaturezaOpExt.sDescricao
    
    TipoTrib.Caption = CStr(objNFiscal.objTributacao.iTipoTributacao)
    TipoTribDesc.Text = objTipoTrib.sDescricao
    
    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    For Each objTipoDocInfo In colTipoDocInfo
    
        If objNFiscal.iTipoNFiscal = objTipoDocInfo.iCodigo Then
        
            sNomeTela = objTipoDocInfo.sNomeTelaNFiscal
            Exit For
            
        End If
    
    Next
    
    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
    
        If sNomeTela = objTipoDocInfo.sNomeTelaNFiscal And objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA Then
        
            TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
            TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
            'se for o tipo padrao, seleciona-o
            
            If objNFiscal.iTipoNFiscal = 0 Then
                If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
            Else
                If objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
            End If
        
        End If
        
    Next
    
    If TipoNFiscal.ListIndex = -1 And TipoNFiscal.ListCount <> 0 Then
        TipoNFiscal.ListIndex = 0
    End If
    
    iIndice = 0

    'Para cada ítem
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        If objItemNF.sProduto <> "" Then
            'Formata o Produto
            lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Else
            sProdutoEnxuto = ""
        End If

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        objProduto.sCodigo = objItemNF.sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
       
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescrItem_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
                
        GridItens.TextMatrix(iIndice, iGrid_CFOP_Col) = objItemNF.objTributacaoItemNF.sNaturezaOp
        GridItens.TextMatrix(iIndice, iGrid_CFOPXml_Col) = objItemNF.sCFOPXml
        GridItens.TextMatrix(iIndice, iGrid_TipoTrib_Col) = CStr(objItemNF.objTributacaoItemNF.iTipoTributacao)
                
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col) = Format(objItemNF.dValorTotal, "Standard")

        GridItens.TextMatrix(iIndice, iGrid_ProdutoXml_Col) = objItemNF.sProdutoXml
        GridItens.TextMatrix(iIndice, iGrid_DescrProd_Col) = objProduto.sDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMedXml_Col) = objItemNF.sUnidadeMedXml
        GridItens.TextMatrix(iIndice, iGrid_EANXml_Col) = objItemNF.sEANXml
        GridItens.TextMatrix(iIndice, iGrid_EANProd_Col) = objProduto.sCodigoBarras
        
        Set objTipoDeTributacao = New ClassTipoDeTributacaoMovto
        objTipoDeTributacao.iTipo = objItemNF.objTributacaoItemNF.iTipoTributacao
        lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
        If lErro <> SUCESSO And lErro <> 27259 Then gError ERRO_SEM_MENSAGEM
        GridItens.TextMatrix(iIndice, iGrid_DescTipoTrib_Col) = objTipoDeTributacao.sDescricao
        
        Set objNaturezaOp = New ClassNaturezaOp
        objNaturezaOp.sCodigo = objItemNF.objTributacaoItemNF.sNaturezaOp
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError ERRO_SEM_MENSAGEM
        GridItens.TextMatrix(iIndice, iGrid_DescCFOP_Col) = objNaturezaOp.sDescricao

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:
   
    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211653)

    End Select

    Exit Function
    
End Function

Private Function Move_Tela_Memoria()

Dim lErro As Long, iIndice As Integer
Dim sProdutoFormatado As String, iProdutoPreenchido As Integer
Dim objItemNF As ClassItemNF
Dim objItemNFAux As ClassItemNF
Dim objPC As ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim bAchou As Boolean, iAux1 As Integer
Dim vbMsg As VbMsgBoxResult
Dim dQuantidade As Double, dValorTotal As Double, dValorDesconto As Double, dValorFreteItem As Double, dValorDescontoItem As Double
Dim dValorOutrasDespesasItem As Double, dValorSeguroItem As Double, dICMSBase As Double, dICMSValor As Double, dICMSSubstBase As Double
Dim dICMSSTCobrAntBase As Double, dICMSSTCobrAntValor As Double, dICMSpercBaseOperacaoPropria As Double, dICMSvBCSTRet As Double
Dim dICMSvICMSSTRet As Double, dICMSvBCSTDest As Double, dICMSvICMSSTDest As Double, dICMSvCredSN As Double
Dim dQtdTrib As Double, dIPIBase As Double, dIPIValor As Double, dIPIUnidadePadraoQtde As Double
Dim dPISBase As Double, dPISValor As Double, dPISQtde As Double, dPISSTBase As Double
Dim dPISSTValor As Double, dPISSTQtde As Double, dCOFINSBase As Double, dCOFINSValor As Double
Dim dCOFINSQtde As Double, dCOFINSSTBase As Double, dCOFINSSTValor As Double, dCOFINSSTQtde As Double
Dim dIIValor As Double, dIIBase As Double, dIIDespAduaneira As Double, dIIIOF As Double
Dim dISSBase As Double, dISSValor As Double, dICMSSubstValor As Double

On Error GoTo Erro_Move_Tela_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objItemNF = gobjNFiscal.colItens(iIndice)

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If iProdutoPreenchido = PRODUTO_VAZIO Then gError 211654
        
        objItemNF.sProduto = sProdutoFormatado
        objItemNF.objTributacaoItemNF.sProduto = sProdutoFormatado
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 201451
        
        objItemNF.sUnidadeMed = Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))
        
        If Abs(objItemNF.dQuantidade - objItemNF.objTributacao.dQtdTrib) < QTDE_ESTOQUE_DELTA Then objItemNF.objTributacao.sUMTrib = objItemNF.sUnidadeMed

        objItemNF.objTributacaoItemNF.sNaturezaOp = GridItens.TextMatrix(iIndice, iGrid_CFOP_Col)
        objItemNF.objTributacaoItemNF.iTipoTributacao = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_TipoTrib_Col))

        If iIndice = 1 Then
        
            gobjNFiscal.objTributacaoNF.sNaturezaOpInterna = objItemNF.objTributacaoItemNF.sNaturezaOp
            gobjNFiscal.objTributacaoNF.iTipoTributacao = objItemNF.objTributacaoItemNF.iTipoTributacao
        
        End If

    Next
    
    If Not (gcolPedidoCompraSel Is Nothing) Then
        
        For iIndice = gcolPedidoCompraSel.Count To 1 Step -1
            gcolPedidoCompraSel.Remove (iIndice)
        Next
               
        'Procura produtos repetidos, se encontrar avisa que terá de agrupar
        'se os preços forem diferentes dá erro
        iAux1 = 0
        For Each objItemNF In gobjNFiscal.ColItensNF
            iAux1 = iAux1 + 1
            dQuantidade = 0
            dValorTotal = 0
            dValorDesconto = 0
            dValorFreteItem = 0
            dValorDescontoItem = 0
            dValorOutrasDespesasItem = 0
            dValorSeguroItem = 0
            dICMSBase = 0
            dICMSValor = 0
            dICMSSubstBase = 0
            dICMSSTCobrAntBase = 0
            dICMSSTCobrAntValor = 0
            dICMSpercBaseOperacaoPropria = 0
            dICMSvBCSTRet = 0
            dICMSvICMSSTRet = 0
            dICMSvBCSTDest = 0
            dICMSvICMSSTDest = 0
            dICMSvCredSN = 0
            dQtdTrib = 0
            dIPIBase = 0
            dIPIValor = 0
            dIPIUnidadePadraoQtde = 0
            dPISBase = 0
            dPISValor = 0
            dPISQtde = 0
            dPISSTBase = 0
            dPISSTValor = 0
            dPISSTQtde = 0
            dCOFINSBase = 0
            dCOFINSValor = 0
            dCOFINSQtde = 0
            dCOFINSSTBase = 0
            dCOFINSSTValor = 0
            dCOFINSSTQtde = 0
            dIIValor = 0
            dIIBase = 0
            dIIDespAduaneira = 0
            dIIIOF = 0
            dISSBase = 0
            dISSValor = 0
            
            For Each objItemNFAux In gobjNFiscal.ColItensNF
                If objItemNF.sProduto = objItemNFAux.sProduto Then
                    If Abs(objItemNF.dPrecoUnitario - objItemNFAux.dPrecoUnitario) > DELTA_VALORMONETARIO Then gError 211739
                    If objItemNF.sUnidadeMed <> objItemNFAux.sUnidadeMed Then gError 211740
                    dQuantidade = dQuantidade + objItemNFAux.dQuantidade
                    dValorTotal = dValorTotal + objItemNFAux.dValorTotalXml
                    dValorDesconto = dValorDesconto + objItemNFAux.dValorDesconto
                    dValorFreteItem = dValorFreteItem + objItemNFAux.objTributacao.dValorFreteItem
                    dValorDescontoItem = dValorDescontoItem + objItemNFAux.objTributacao.dValorDescontoItem
                    dValorOutrasDespesasItem = dValorOutrasDespesasItem + objItemNFAux.objTributacao.dValorOutrasDespesasItem
                    dValorSeguroItem = dValorSeguroItem + objItemNFAux.objTributacao.dValorSeguroItem
                    dICMSBase = dICMSBase + objItemNFAux.objTributacao.dICMSBase
                    dICMSValor = dICMSValor + objItemNFAux.objTributacao.dICMSValor
                    dICMSSubstBase = dICMSSubstBase + objItemNFAux.objTributacao.dICMSSubstBase
                    dICMSSubstValor = dICMSSubstValor + objItemNFAux.objTributacao.dICMSSubstValor
                    dICMSSTCobrAntBase = dICMSSTCobrAntBase + objItemNFAux.objTributacao.dICMSSTCobrAntBase
                    dICMSSTCobrAntValor = dICMSSTCobrAntValor + objItemNFAux.objTributacao.dICMSSTCobrAntValor
                    dICMSpercBaseOperacaoPropria = dICMSpercBaseOperacaoPropria + objItemNFAux.objTributacao.dICMSpercBaseOperacaoPropria
                    dICMSvBCSTRet = dICMSvBCSTRet + objItemNFAux.objTributacao.dICMSvBCSTRet
                    dICMSvICMSSTRet = dICMSvICMSSTRet + objItemNFAux.objTributacao.dICMSvICMSSTRet
                    dICMSvBCSTDest = dICMSvBCSTDest + objItemNFAux.objTributacao.dICMSvBCSTDest
                    dICMSvICMSSTDest = dICMSvICMSSTDest + objItemNFAux.objTributacao.dICMSvICMSSTDest
                    dICMSvCredSN = dICMSvCredSN + objItemNFAux.objTributacao.dICMSvCredSN
                    dQtdTrib = dQtdTrib + objItemNFAux.objTributacao.dQtdTrib
                    dIPIBase = dIPIBase + objItemNFAux.objTributacao.dIPIBaseCalculo
                    dIPIValor = dIPIValor + objItemNFAux.objTributacao.dIPIValor
                    dIPIUnidadePadraoQtde = dIPIUnidadePadraoQtde + objItemNFAux.objTributacao.dIPIUnidadePadraoQtde
                    dPISBase = dPISBase + objItemNFAux.objTributacao.dPISBase
                    dPISValor = dPISValor + objItemNFAux.objTributacao.dPISValor
                    dPISQtde = dPISQtde + objItemNFAux.objTributacao.dPISQtde
                    dPISSTBase = dPISSTBase + objItemNFAux.objTributacao.dPISSTBase
                    dPISSTValor = dPISSTValor + objItemNFAux.objTributacao.dPISSTValor
                    dPISSTQtde = dPISSTQtde + objItemNFAux.objTributacao.dPISSTQtde
                    dCOFINSBase = dCOFINSBase + objItemNFAux.objTributacao.dCOFINSBase
                    dCOFINSValor = dCOFINSValor + objItemNFAux.objTributacao.dCOFINSValor
                    dCOFINSQtde = dCOFINSQtde + objItemNFAux.objTributacao.dCOFINSQtde
                    dCOFINSSTBase = dCOFINSSTBase + objItemNFAux.objTributacao.dCOFINSSTBase
                    dCOFINSSTValor = dCOFINSSTValor + objItemNFAux.objTributacao.dCOFINSSTValor
                    dCOFINSSTQtde = dCOFINSSTQtde + objItemNFAux.objTributacao.dCOFINSSTQtde
                    dIIValor = dIIValor + objItemNFAux.objTributacao.dIIValor
                    dIIBase = dIIBase + objItemNFAux.objTributacao.dIIBase
                    dIIDespAduaneira = dIIDespAduaneira + objItemNFAux.objTributacao.dIIDespAduaneira
                    dIIIOF = dIIIOF + objItemNFAux.objTributacao.dIIIOF
                    dISSBase = dISSBase + objItemNFAux.objTributacao.dISSBase
                    dISSValor = dISSValor + objItemNFAux.objTributacao.dISSValor
                End If
            Next
            'Se não bateu é porque os itens se repetem
            If Abs(objItemNF.dQuantidade - dQuantidade) > QTDE_ESTOQUE_DELTA Then
                
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_NF_IMPORT_PROD_REPETIDO_PC", objItemNF.sProduto)
                If vbMsg = vbNo Then gError ERRO_SEM_MENSAGEM
            
                For iIndice = gobjNFiscal.ColItensNF.Count To 1 Step -1
                    Set objItemNFAux = gobjNFiscal.ColItensNF.Item(iIndice)
                    If iIndice <> iAux1 Then
                        If objItemNF.sProduto = objItemNFAux.sProduto Then
                            gobjNFiscal.ColItensNF.Remove (iIndice)
                        End If
                    End If
                Next
                
                objItemNF.dQuantidade = dQuantidade
                objItemNF.dValorTotalXml = dValorTotal
                objItemNF.dValorDesconto = dValorDescontoItem
                
'                If (dValorTotal - (Arredonda_Moeda(dQuantidade * objItemNF.dPrecoUnitario) - dValorDescontoItem)) > DELTA_VALORMONETARIO Then
'                    objItemNF.dPrecoUnitario = (dValorTotal + dValorDescontoItem) / dQuantidade
'                End If
'                objItemNF.dValorDesconto = Abs(Arredonda_Moeda(objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - dValorTotal) 'dValorDesconto
'                If Abs(dValorDescontoItem - objItemNF.dValorDesconto) > DELTA_VALORMONETARIO Then
'                    gobjNFiscal.dValorDesconto = gobjNFiscal.dValorDesconto + (dValorDescontoItem - objItemNF.dValorDesconto)
'                End If
                objItemNF.objTributacaoItemNF.dDescontoGrid = objItemNF.dValorDesconto
                objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotalXml
                objItemNF.objTributacaoItemNF.dValorFreteItem = dValorFreteItem
                objItemNF.objTributacaoItemNF.dValorDescontoItem = dValorDescontoItem
                objItemNF.objTributacaoItemNF.dValorOutrasDespesasItem = dValorOutrasDespesasItem
                objItemNF.objTributacaoItemNF.dValorSeguroItem = dValorSeguroItem
                objItemNF.objTributacaoItemNF.dICMSBase = dICMSBase
                objItemNF.objTributacaoItemNF.dICMSValor = dICMSValor
                objItemNF.objTributacaoItemNF.dICMSSubstBase = dICMSSubstBase
                objItemNF.objTributacaoItemNF.dICMSSubstValor = dICMSSubstValor
                objItemNF.objTributacaoItemNF.dICMSSTCobrAntBase = dICMSSTCobrAntBase
                objItemNF.objTributacaoItemNF.dICMSSTCobrAntValor = dICMSSTCobrAntValor
                objItemNF.objTributacaoItemNF.dICMSpercBaseOperacaoPropria = dICMSpercBaseOperacaoPropria
                objItemNF.objTributacaoItemNF.dICMSvBCSTRet = dICMSvBCSTRet
                objItemNF.objTributacaoItemNF.dICMSvICMSSTRet = dICMSvICMSSTRet
                objItemNF.objTributacaoItemNF.dICMSvBCSTDest = dICMSvBCSTDest
                objItemNF.objTributacaoItemNF.dICMSvICMSSTDest = dICMSvICMSSTDest
                objItemNF.objTributacaoItemNF.dICMSvCredSN = dICMSvCredSN
                objItemNF.objTributacaoItemNF.dQtdTrib = dQtdTrib
                objItemNF.objTributacaoItemNF.dIPIBaseCalculo = dIPIBase
                objItemNF.objTributacaoItemNF.dIPIValor = dIPIValor
                objItemNF.objTributacaoItemNF.dIPIUnidadePadraoQtde = dIPIUnidadePadraoQtde
                objItemNF.objTributacaoItemNF.dPISBase = dPISBase
                objItemNF.objTributacaoItemNF.dPISValor = dPISValor
                objItemNF.objTributacaoItemNF.dPISQtde = dPISQtde
                objItemNF.objTributacaoItemNF.dPISSTBase = dPISSTBase
                objItemNF.objTributacaoItemNF.dPISSTValor = dPISSTValor
                objItemNF.objTributacaoItemNF.dPISSTQtde = dPISSTQtde
                objItemNF.objTributacaoItemNF.dCOFINSBase = dCOFINSBase
                objItemNF.objTributacaoItemNF.dCOFINSValor = dCOFINSValor
                objItemNF.objTributacaoItemNF.dCOFINSQtde = dCOFINSQtde
                objItemNF.objTributacaoItemNF.dCOFINSSTBase = dCOFINSSTBase
                objItemNF.objTributacaoItemNF.dCOFINSSTValor = dCOFINSSTValor
                objItemNF.objTributacaoItemNF.dCOFINSSTQtde = dCOFINSSTQtde
                objItemNF.objTributacaoItemNF.dIIValor = dIIValor
                objItemNF.objTributacaoItemNF.dIIBase = dIIBase
                objItemNF.objTributacaoItemNF.dIIDespAduaneira = dIIDespAduaneira
                objItemNF.objTributacaoItemNF.dIIIOF = dIIIOF
                objItemNF.objTributacaoItemNF.dISSBase = dISSBase
                objItemNF.objTributacaoItemNF.dISSValor = dISSValor
                
            End If
            
            If Abs(dValorTotal - (Arredonda_Moeda(dQuantidade * objItemNF.dPrecoUnitario) - dValorDescontoItem)) > DELTA_VALORMONETARIO2 Then
                objItemNF.dPrecoUnitario = (dValorTotal + dValorDescontoItem) / dQuantidade
            End If
        Next
        'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz
        
        'Para cada Pedido de Compras da lista
        For iIndice = 0 To PedidosCompra.ListCount - 1
            'Se o Pedido estiver selecionado
            If PedidosCompra.Selected(iIndice) = True Then
                gcolPedidoCompraSel.Add gcolPedidoCompra.Item(iIndice + 1)
            End If
        Next
        
        If gcolPedidoCompraSel.Count = 0 Then gError 211708 ' Para tela via pedido de compra o PC tem que estar marcado
        
        For Each objPC In gcolPedidoCompraSel
            lErro = CF("ItensPCTodos_Le_Codigo", objPC)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Next
        
        For Each objItemNF In gobjNFiscal.colItens
            bAchou = False
            For Each objPC In gcolPedidoCompraSel
                For Each objItemPC In objPC.colItens
                    If objItemPC.sProduto = objItemNF.sProduto Then
                        bAchou = True
                        Exit For
                    End If
                Next
            Next
            If Not bAchou Then gError 211709 'Não pode ter produto na NF que não consta em um pedido de compra marcado
        Next
        
    End If

    gobjNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    gobjNFiscal.iFilialPedido = Codigo_Extrai(FilialCompra.Text)
    
    lErro = CF("NFImport_AtualizaConv", gobjNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:
   
    Move_Tela_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201451
             Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA_LINHA", gErr, iIndice)

        Case 211654
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_LINHA", gErr, iIndice)

        Case 211708
             Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_SELECIONADO", gErr)
        
        Case 211709
             Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_ITEM_INEXISTENTE", gErr, objItemNF.sProduto)

        Case 211739, 211740 'ERRO_NF_PC_PROD_IGUAIS_PRECO_DIF
             Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_ITEM_INEXISTENTE", gErr, objItemNF.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211655)

    End Select

    Exit Function
    
End Function

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Produto XML")
    objGrid.colColuna.Add ("Descrição do Item")
    objGrid.colColuna.Add ("Descrição do Produto")
    objGrid.colColuna.Add ("CFOP XML")
    objGrid.colColuna.Add ("CFOP")
    objGrid.colColuna.Add ("Desc.CFOP")
    objGrid.colColuna.Add ("Tipo Trib")
    objGrid.colColuna.Add ("Desc.Tipo Trib")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("U.M. XML")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Valor Unitário")
    objGrid.colColuna.Add ("Valor Total")
    objGrid.colColuna.Add ("EAN do Produto")
    objGrid.colColuna.Add ("EAN no XML")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (ProdutoXml.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (DescProd.Name)
    objGrid.colCampo.Add (CFOPXml.Name)
    objGrid.colCampo.Add (NatOpInterna.Name)
    objGrid.colCampo.Add (DescCFOP.Name)
    objGrid.colCampo.Add (TipoTributacao.Name)
    objGrid.colCampo.Add (DescTipoTrib.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (UnidadeMedXml.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (ValorUnitario.Name)
    objGrid.colCampo.Add (ValorTotal.Name)
    objGrid.colCampo.Add (EANProd.Name)
    objGrid.colCampo.Add (EANXml.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_ProdutoXml_Col = 2
    iGrid_DescrItem_Col = 3
    iGrid_DescrProd_Col = 4
    iGrid_CFOPXml_Col = 5
    iGrid_CFOP_Col = 6
    iGrid_DescCFOP_Col = 7
    iGrid_TipoTrib_Col = 8
    iGrid_DescTipoTrib_Col = 9
    iGrid_UnidadeMed_Col = 10
    iGrid_UnidadeMedXml_Col = 11
    iGrid_Quantidade_Col = 12
    iGrid_ValorUnitario_Col = 13
    iGrid_ValorTotal_Col = 14
    iGrid_EANProd_Col = 15
    iGrid_EANXml_Col = 16

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
    Call Exibe_CampoDet_Grid(objGridItens, GridItens.Col, Detalhe)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UnidadeMed_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UnidadeMed_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'GridItensNF
        If objGridInt.objGrid.Name = GridItens.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Produto_Col

                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                Case iGrid_TipoTrib_Col
                    lErro = Saida_Celula_TipoTributacao(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_CFOP_Col
                    lErro = Saida_Celula_CFOP(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End Select
                    
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 211656

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211656
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211657)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name
        
        Case Produto.Name
            objControl.Enabled = True
            
        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                If lErro = 28030 Then gError 211658 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError ERRO_SEM_MENSAGEM
                End If
            End If
            
        Case NatOpInterna.Name
            objControl.Enabled = True
        
        Case TipoTributacao.Name
            objControl.Enabled = True
        
        Case Else
            objControl.Enabled = False

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 211658
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211659)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then
        
        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25042 And lErro <> 25043 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 25041 Then gError 211669
        If lErro = 25043 Then gError 25043
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
        
        'Executa a saida de célula
        lErro = Produto_Saida_Celula(objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 25043
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, sProduto)

        Case 211669
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211660)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211661)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Public Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    'Verifica se o Produto está preenchido
'    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 211662

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        If Not (Me.ActiveControl Is Produto) Then
    
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            'para forçar preencher descrica, um, ean,...
            GridItens.TextMatrix(GridItens.Row, iGrid_DescrProd_Col) = ""
            
            lErro = Produto_Saida_Celula(objProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
    'End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoProduto_evSelecao:

    Select Case gErr
    
        Case 211662
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211663)
            
    End Select

    Exit Sub
    
End Sub

Function Produto_Saida_Celula(Optional ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String, sProdAnt As String
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUMXml As String, bAchou As Boolean, sProdXml As String, sUM As String

On Error GoTo Erro_Produto_Saida_Celula

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    If lErro <> SUCESSO Then gError 211663
    
    If objProduto.iGerencial = GERENCIAL Then gError 211664 'Alterado por Wagner
    
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdAnt, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If sProdAnt <> objProduto.sCodigo Or GridItens.TextMatrix(GridItens.Row, iGrid_DescrProd_Col) <> objProduto.sDescricao Then

        'Descricao
        GridItens.TextMatrix(GridItens.Row, iGrid_DescrProd_Col) = objProduto.sDescricao
        'EAN
        GridItens.TextMatrix(GridItens.Row, iGrid_EANProd_Col) = objProduto.sCodigoBarras

        sUMXml = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMedXml_Col)
        
        bAchou = False
        
        sProdXml = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoXml_Col)
        lErro = CF("ImportNFeXml_ObterSiglaUM", gobjNFiscal.iFilialEmpresa, gobjNFiscal.sCGCEmitente, sProdXml, sUMXml, sUM)
        If lErro <> SUCESSO And lErro <> ERRO_SEM_MENSAGEM Then gError ERRO_SEM_MENSAGEM
        If lErro = SUCESSO Then
        
            bAchou = True
            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = sUM
            
        End If
        
        If bAchou = False Then
        
            objClasseUM.iClasse = objProduto.iClasseUM
            'Lâ as Unidades de Medidas da Classe do produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            For Each objUM In colSiglas
                If UCase(objUM.sSigla) = UCase(sUMXml) Then
                    bAchou = True
                    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objUM.sSigla
                    Exit For
                End If
            Next
            
            If Not bAchou Then GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
    
        End If
        
    End If
    
    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 211663
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If
            
        Case 211664
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211665)

    End Select

    Exit Function

End Function

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 211666

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoCompraLista
    Call Chama_Tela_Modal("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 211666
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211667)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is NatOpInterna Then
            Call BotaoCFOP_Click
        ElseIf Me.ActiveControl Is TipoTributacao Then
            Call BotaoTipoTrib_Click
        End If
    End If

End Sub

Private Sub BotaoOK_Click()
Dim lErro As Long
    lErro = Move_Tela_Memoria
    If lErro = SUCESSO Then
        'Nao mexer no obj da tela
        giRetornoTela = vbOK
    
        Unload Me
    End If
    
End Sub

Private Sub TipoTributacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub TipoTributacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TipoTributacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NatOpInterna_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NatOpInterna_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub NatOpInterna_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub NatOpInterna_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = NatOpInterna
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CFOP(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long, sNatOp As String, objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_Saida_Celula_CFOP

    Set objGridInt.objControle = NatOpInterna

    sNatOp = Trim(NatOpInterna.Text)

    If sNatOp <> "" Then
    
        objNaturezaOp.sCodigo = sNatOp
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 56879
        
        'se for entrada garantir que tenha codigo < 500
        If objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then gError 56992

        If Len(sNatOp) <> 4 Then gError 32284
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCFOP_Col) = objNaturezaOp.sDescricao
    
    End If
    
    If GridItens.Row = 1 Then
    
        CFOPInt.Caption = sNatOp
        CFOPIntDesc.Text = objNaturezaOp.sDescricao
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_CFOP = SUCESSO

    Exit Function

Erro_Saida_Celula_CFOP:

    Saida_Celula_CFOP = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56879
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, sNatOp)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 56992
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", gErr, sNatOp)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211661)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_TipoTributacao(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Saida_Celula_TipoTributacao

    Set objGridInt.objControle = TipoTributacao

    If Len(Trim(TipoTributacao.Text)) <> 0 Then

        objTipoDeTributacao.iTipo = CInt(Trim(TipoTributacao.Text))

        lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
        If lErro <> SUCESSO And lErro <> 27259 Then gError ERRO_SEM_MENSAGEM

        If lErro = 27259 Then gError 27794

        If objTipoDeTributacao.iEntrada = 0 Then gError 59378

        GridItens.TextMatrix(GridItens.Row, iGrid_DescTipoTrib_Col) = objTipoDeTributacao.sDescricao

    End If
    
    If GridItens.Row = 1 Then
    
        TipoTrib.Caption = Trim(TipoTributacao.Text)
        TipoTribDesc.Text = objTipoDeTributacao.sDescricao
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_TipoTributacao = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoTributacao:

    Saida_Celula_TipoTributacao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 59378
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIB_INCOMPAT_ENTRADA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211661)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Function Atualiza_ListaPedidos() As Long
'Atualiza a Lista de Pedidos de Compra com os códigos desses pedidos

Dim lErro As Long
Dim iFilial As Integer
Dim iFilialCompra As Integer
Dim objPedidoCompras As ClassPedidoCompras
Dim objFornecedor As New ClassFornecedor
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_ListaPedidos

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
    
    If UCase(objTipoDocInfo.sNomeTelaNFiscal) = UCase("NFiscalFatEntradaCom") Or UCase(objTipoDocInfo.sNomeTelaNFiscal) = UCase("NFiscalEntradaCom") Then
        
        'Se Fornecedor, Filial e Filial de Compra estão preenchidos
        If Len(Trim(TipoNFiscal.Text)) > 0 And gobjNFiscal.lFornecedor <> 0 And gobjNFiscal.iFilialForn <> 0 And Len(Trim(FilialCompra.Text)) > 0 Then
    
            'Lê Pedidos de Compra com o Fornecedor, Filial e FilialCompra da tela
            lErro = CF("PedidosCompraEnv_Le_Recebimento", gcolPedidoCompra, gobjNFiscal.lFornecedor, gobjNFiscal.iFilialForn, Codigo_Extrai(FilialCompra.Text))
            If lErro <> SUCESSO And lErro <> 65845 Then gError ERRO_SEM_MENSAGEM
            If lErro = 65845 Then gError 65785
        
            'Verifica para cada Pedido lido se existe pelo menos um itemPC com quantidade a receber
            lErro = CF("ItensPedidosCompra_VerificaQuantidade", gcolPedidoCompra)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Para cada Pedido Lido
            For iIndice = gcolPedidoCompra.Count To 1 Step -1
            
                Set objPedidoCompras = gcolPedidoCompra(iIndice)
            
                'Se for uma nota de beneficiamento
                If gobjNFiscal.iTipoNFiscal = DOCINFO_NFEEBFCOM Or gobjNFiscal.iTipoNFiscal = DOCINFO_NFIEBFCOM Then
                    'Remove os Pedidos c\ destino <> Fornecedor
                    If objPedidoCompras.iTipoDestino <> TIPO_DESTINO_FORNECEDOR Then gcolPedidoCompra.Remove (iIndice)
                Else
                
                    If objPedidoCompras.iTipoDestino = TIPO_DESTINO_FORNECEDOR And objFornecedor.lCodigo = objPedidoCompras.lFornCliDestino And iFilial = objPedidoCompras.iFilialDestino Then
                    
                        'if iTipoNFAlterado <> DOCINFO.... then gcolPedidoCompra.Remove (iIndice)
                        
                    Else
                    
                        'Remove os PCs com destino <> empresa ou outra filialempresa
                        lErro = CF("Remove_Pedido_Compra", objPedidoCompras, gcolPedidoCompra, iIndice)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    End If
                    
                End If
            
            Next
    
            'Preenche a ListBox de pedidos com os pedidos lidos do BD
            For Each objPedidoCompras In gcolPedidoCompra
                PedidosCompra.AddItem objPedidoCompras.lCodigo
            Next
    
        End If
        
    Else
        PedidosCompra.Clear
        'Para cada Pedido Lido
        For iIndice = gcolPedidoCompra.Count To 1 Step -1
            gcolPedidoCompra.Remove iIndice
        Next
    End If

    Atualiza_ListaPedidos = SUCESSO

    Exit Function

Erro_Atualiza_ListaPedidos:

    Atualiza_ListaPedidos = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 65783
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gobjNFiscal.lFornecedor)
        
        Case 65785
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDOCOMPRAS", gErr, CliForn.Caption, gobjNFiscal.iFilialForn, Codigo_Extrai(FilialCompra.Text))
        
        Case 92134
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSCOMPRA_JA_RECEBIDOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156720)

    End Select

    Exit Function

End Function

Private Sub TipoNFiscal_Click()
    If gobjNFiscal.iTipoNFiscal <> Codigo_Extrai(TipoNFiscal.Text) Then
        gobjNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
        Call Trata_CliForn
    End If
End Sub

Public Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Validate

    'Verifica se o tipo está preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionado
    If TipoNFiscal.List(TipoNFiscal.ListIndex) = TipoNFiscal.Text Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42505
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then Error 42506

    'tenta ler a natureza de operacao
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 55550

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 55551

    If gobjNFiscal.iTipoNFiscal <> Codigo_Extrai(TipoNFiscal.Text) Then
        gobjNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
        Call Trata_CliForn
    End If

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case Err

        Case 42505, 55550

        Case 42506
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, TipoNFiscal.Text)

        Case 55551
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156895)

    End Select

    Exit Sub

End Sub

Public Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCompra_Validate
    
    'Verifica se é uma FilialEmpresa selecionada
    If FilialCompra.Text = FilialCompra.List(FilialCompra.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 66226

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 66227

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 66228

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 66229

    If Len(Trim(TipoNFiscal.Text)) > 0 And iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text) Then
        lErro = Atualiza_ListaPedidos
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)
        
    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 66226, 66227
            
        Case 66228
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialCompra.Text)

        Case 66229
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialCompra.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157053)

    End Select

    Exit Sub

End Sub

Private Function Trata_CliForn() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objCli As New ClassCliente
Dim objFilCli As New ClassFilialCliente
Dim objForn As New ClassFornecedor
Dim objFilForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_CliForn

    gobjNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    
    gobjNFiscal.lCliente = 0
    gobjNFiscal.lFornecedor = 0
    
    lErro = CF("NFImport_AtualizaConv", gobjNFiscal, False)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If gobjNFiscal.lCliente <> 0 Then
    
        LabelCli.Visible = True
        LabelForn.Visible = False
    
        objCli.lCodigo = gobjNFiscal.lCliente
        objFilCli.lCodCliente = gobjNFiscal.lCliente
        objFilCli.iCodFilial = gobjNFiscal.iFilialCli

        lErro = CF("Cliente_Le", objCli)
        If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM

        lErro = CF("FilialCliente_Le", objFilCli)
        If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM

        CliForn.Caption = CStr(objCli.lCodigo) & SEPARADOR & objCli.sNomeReduzido
        Filial.Caption = CStr(objFilCli.iCodFilial) & SEPARADOR & objFilCli.sNome

    Else
    
        LabelCli.Visible = False
        LabelForn.Visible = True

        objForn.lCodigo = gobjNFiscal.lFornecedor
        objFilForn.lCodFornecedor = gobjNFiscal.lFornecedor
        objFilForn.iCodFilial = gobjNFiscal.iFilialForn
    
        lErro = CF("Fornecedor_Le", objForn)
        If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM

        lErro = CF("FilialFornecedor_Le", objFilForn)
        If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
        
        CliForn.Caption = CStr(objForn.lCodigo) & SEPARADOR & objForn.sNomeReduzido
        Filial.Caption = CStr(objFilForn.iCodFilial) & SEPARADOR & objFilForn.sNome
    
    End If
    
    lErro = Atualiza_ListaPedidos
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Trata_CliForn = SUCESSO
    
    Exit Function

Erro_Trata_CliForn:

    Trata_CliForn = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157053)

    End Select

    Exit Function

End Function

Sub BotaoCFOP_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date
Dim sCFOP As String

On Error GoTo Erro_BotaoCFOP_Click

    If Me.ActiveControl Is NatOpInterna Then
    
        sCFOP = NatOpInterna.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 211666

        sCFOP = GridItens.TextMatrix(GridItens.Row, iGrid_CFOP_Col)
        
    End If

    dtDataRef = StrParaDate(DataEmissao.Caption)

    'Se NaturezaOP estiver preenchida coloca no Obj
    If Len(Trim(sCFOP)) > 0 Then objNaturezaOp.sCodigo = sCFOP

    colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
    colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
    
    Call Chama_Tela_Modal("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoCFOP, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))
    
    Exit Sub

Erro_BotaoCFOP_Click:

    Select Case gErr
    
        Case 211666
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211667)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoCFOP_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNaturezaOp As ClassNaturezaOp

On Error GoTo Erro_objEventoCFOP_evSelecao
    
    Set objNaturezaOp = obj1

    NatOpInterna.PromptInclude = False
    NatOpInterna.Text = objNaturezaOp.sCodigo
    NatOpInterna.PromptInclude = True
    
    If Not (Me.ActiveControl Is NatOpInterna) Then

        GridItens.TextMatrix(GridItens.Row, iGrid_CFOP_Col) = NatOpInterna.Text
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescCFOP_Col) = objNaturezaOp.sDescricao
               
        If GridItens.Row = 1 Then
        
            CFOPInt.Caption = NatOpInterna.Text
            CFOPIntDesc.Text = objNaturezaOp.sDescricao
        
        End If
    
    End If
    Call Exibe_CampoDet_Grid(objGridItens, GridItens.Col, Detalhe)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCFOP_evSelecao:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211663)
            
    End Select

    Exit Sub

End Sub

Sub BotaoTipoTrib_Click()

Dim colSelecao As New Collection
Dim iTipoTrib As Integer
Dim objTipoTrib As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_BotaoTipoTrib_Click

    If Me.ActiveControl Is TipoTributacao Then
    
        iTipoTrib = StrParaInt(TipoTributacao.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 211666

        iTipoTrib = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_TipoTrib_Col))
        
    End If

    colSelecao.Add "1"
    colSelecao.Add "1"
    
    Call Chama_Tela_Modal("TiposDeTribMovtoLista", colSelecao, objTipoTrib, objEventoTipoTrib)
    
    Exit Sub

Erro_BotaoTipoTrib_Click:

    Select Case gErr
    
        Case 211666
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211667)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoTipoTrib_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoTrib As ClassTipoDeTributacaoMovto

On Error GoTo Erro_objEventoTipoTrib_evSelecao
    
    Set objTipoTrib = obj1

    TipoTributacao.PromptInclude = False
    TipoTributacao.Text = objTipoTrib.iTipo
    TipoTributacao.PromptInclude = True
    
    If Not (Me.ActiveControl Is NatOpInterna) Then

        GridItens.TextMatrix(GridItens.Row, iGrid_TipoTrib_Col) = TipoTributacao.Text
        
        GridItens.TextMatrix(GridItens.Row, iGrid_DescTipoTrib_Col) = objTipoTrib.sDescricao
               
        If GridItens.Row = 1 Then
        
            TipoTrib.Caption = Trim(TipoTributacao.Text)
            TipoTribDesc.Text = objTipoTrib.sDescricao
        
        End If
    
    End If
    Call Exibe_CampoDet_Grid(objGridItens, GridItens.Col, Detalhe)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipoTrib_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211663)
            
    End Select

    Exit Sub

End Sub

Private Sub Exibe_CampoDet_Grid(ByVal objGridInt As AdmGrid, ByVal iColunaExibir As Integer, ByVal objControle As Object)

Dim iLinha As Integer

On Error GoTo Erro_Exibe_CampoDet_Grid

    iLinha = objGridInt.objGrid.Row
    
    If iLinha > 0 And iLinha <= objGridInt.iLinhasExistentes Then
        objControle.Text = objGridInt.objGrid.TextMatrix(iLinha, iColunaExibir)
    Else
        objControle.Text = ""
    End If

    Exit Sub

Erro_Exibe_CampoDet_Grid:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208641)

    End Select

    Exit Sub
    
End Sub
