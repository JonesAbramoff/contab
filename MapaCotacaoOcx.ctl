VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl MapaCotacaoOcx 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   11325
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4665
      Index           =   2
      Left            =   108
      TabIndex        =   28
      Top             =   810
      Visible         =   0   'False
      Width           =   11085
      Begin VB.CommandButton BotaoGeraImprime 
         Caption         =   "Gerar e Imprimir o Mapa"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2124
         TabIndex        =   16
         Top             =   4035
         Width           =   2508
      End
      Begin VB.CommandButton BotaoGerarMapa 
         Caption         =   "Gerar Mapa"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4740
         TabIndex        =   17
         Top             =   4035
         Width           =   1572
      End
      Begin VB.CommandButton BotaoPedCotacao 
         Caption         =   "Pedido de Cotação..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   24
         TabIndex        =   15
         Top             =   4035
         Width           =   2004
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   312
         Left            =   1476
         Picture         =   "MapaCotacaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Numeração Automática"
         Top             =   72
         Width           =   300
      End
      Begin VB.Frame Frame4 
         Caption         =   " Produtos x Pedidos de Cotação "
         Height          =   3405
         Left            =   24
         TabIndex        =   29
         Top             =   480
         Width           =   10965
         Begin VB.TextBox FilialForn 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   252
            Left            =   252
            TabIndex        =   50
            Top             =   1728
            Width           =   2088
         End
         Begin VB.TextBox Fornecedor 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   252
            Left            =   2484
            TabIndex        =   49
            Top             =   1296
            Width           =   2088
         End
         Begin VB.TextBox UM 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   252
            Left            =   1368
            TabIndex        =   48
            Top             =   1296
            Width           =   1056
         End
         Begin VB.CommandButton BotaoMarcarTodosGrid 
            Caption         =   "Marcar Todos"
            Height          =   540
            Left            =   1044
            Picture         =   "MapaCotacaoOcx.ctx":00EA
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2808
            Width           =   1932
         End
         Begin VB.CommandButton BotaoDesmarcarTodosGrid 
            Caption         =   "Desmarcar Todos"
            Height          =   540
            Left            =   3204
            Picture         =   "MapaCotacaoOcx.ctx":1104
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2808
            Width           =   1932
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   252
            Left            =   3456
            TabIndex        =   43
            Top             =   864
            Width           =   2088
            _ExtentX        =   3704
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Quant 
            Height          =   252
            Left            =   252
            TabIndex        =   42
            Top             =   1296
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.TextBox PedidoCot 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   252
            Left            =   1224
            TabIndex        =   26
            Top             =   864
            Width           =   1056
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   252
            Left            =   2376
            TabIndex        =   25
            Top             =   864
            Width           =   996
            _ExtentX        =   1773
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.CheckBox Seleciona 
            Height          =   240
            Left            =   288
            TabIndex        =   12
            Top             =   870
            Width           =   852
         End
         Begin MSFlexGridLib.MSFlexGrid GridMapaCotacao 
            Height          =   2370
            Left            =   150
            TabIndex        =   11
            Top             =   225
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   4180
            _Version        =   393216
         End
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   312
         Left            =   756
         TabIndex        =   9
         Top             =   72
         Width           =   708
         _ExtentX        =   1270
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataMapaCotacao 
         Height          =   312
         Left            =   4020
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   72
         Width           =   228
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataMapaCotacao 
         Height          =   312
         Left            =   2832
         TabIndex        =   10
         Top             =   72
         Width           =   1176
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label TaxaFinanceira 
         BorderStyle     =   1  'Fixed Single
         Height          =   312
         Left            =   5292
         TabIndex        =   47
         Top             =   72
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Taxa:"
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
         Height          =   192
         Left            =   4776
         TabIndex        =   46
         Top             =   132
         Width           =   480
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   2340
         TabIndex        =   45
         Top             =   132
         Width           =   456
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
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
         Height          =   192
         Left            =   60
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Top             =   132
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4680
      Index           =   1
      Left            =   72
      TabIndex        =   4
      Top             =   765
      Width           =   11145
      Begin VB.Frame Frame2 
         Caption         =   " Pedidos de Cotação "
         Height          =   2100
         Left            =   360
         TabIndex        =   27
         Top             =   144
         Width           =   5676
         Begin VB.Frame Frame5 
            Caption         =   " Data "
            Height          =   1536
            Left            =   2772
            TabIndex        =   34
            Top             =   324
            Width           =   2604
            Begin MSComCtl2.UpDown UpDownDataDe 
               Height          =   312
               Left            =   1980
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   384
               Width           =   228
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   312
               Left            =   792
               TabIndex        =   2
               Top             =   384
               Width           =   1176
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataAte 
               Height          =   312
               Left            =   1980
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   996
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataAte 
               Height          =   312
               Left            =   792
               TabIndex        =   3
               Top             =   996
               Width           =   1176
               _ExtentX        =   2064
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label3 
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
               Height          =   192
               Left            =   432
               TabIndex        =   38
               Top             =   1056
               Width           =   360
            End
            Begin VB.Label Label1 
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
               Height          =   192
               Left            =   480
               TabIndex        =   37
               Top             =   444
               Width           =   312
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   " Número "
            Height          =   1536
            Left            =   288
            TabIndex        =   39
            Top             =   324
            Width           =   2208
            Begin MSMask.MaskEdBox CodPCDe 
               Height          =   300
               Left            =   828
               TabIndex        =   0
               Top             =   384
               Width           =   888
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CodPCAte 
               Height          =   300
               Left            =   828
               TabIndex        =   1
               Top             =   1008
               Width           =   888
               _ExtentX        =   1561
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin VB.Label LabelCodPCAte 
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
               Height          =   192
               Left            =   444
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   41
               Top             =   1068
               Width           =   360
            End
            Begin VB.Label LabelCodPCDe 
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
               Height          =   192
               Left            =   492
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   40
               Top             =   444
               Width           =   312
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Produtos "
         Height          =   2160
         Left            =   360
         TabIndex        =   22
         Top             =   2304
         Width           =   5676
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            Height          =   612
            Left            =   3744
            Picture         =   "MapaCotacaoOcx.ctx":22E6
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   705
            Width           =   1752
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            Height          =   600
            Left            =   3744
            Picture         =   "MapaCotacaoOcx.ctx":3300
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1410
            Width           =   1740
         End
         Begin VB.Frame Frame10 
            Caption         =   " Itens de Categoria "
            Height          =   1404
            Left            =   144
            TabIndex        =   33
            Top             =   612
            Width           =   3456
            Begin VB.ListBox ItensCategoria 
               Height          =   510
               Left            =   144
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   6
               Top             =   288
               Width           =   3156
            End
         End
         Begin VB.ComboBox Categoria 
            Height          =   288
            Left            =   1068
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   252
            Width           =   2532
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
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
            Height          =   192
            Left            =   144
            TabIndex        =   32
            Top             =   300
            Width           =   876
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   9135
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "MapaCotacaoOcx.ctx":44E2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   120
         Picture         =   "MapaCotacaoOcx.ctx":466C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "MapaCotacaoOcx.ctx":476E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "MapaCotacaoOcx.ctx":48EC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5085
      Left            =   30
      TabIndex        =   23
      Top             =   420
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   8969
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produtos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MapaCotacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFramePrincipalAlterado As Integer

Dim objGridMapaCotacao As AdmGrid
Dim iGrid_Seleciona_Col As Integer
Dim iGrid_PedidoCot_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Quant_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_FilialForn_Col As Integer

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoCodigoMapaCotacao As AdmEvento
Attribute objEventoCodigoMapaCotacao.VB_VarHelpID = -1

'Variável global a tela que guarda o conteudo do tab de selecao e dos itens de MapaCotacao
Dim gobjGeracaoMapaCotacao As ClassGeracaoMapaCotacao

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Mapa de Cotação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "MapaCotacao"

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

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 114602

    gobjGeracaoMapaCotacao.objMapaCotacao.lCodigo = StrParaLong(Codigo.Text)
    
    'Verifica se o mapa existe
    lErro = CF("MapaCotacao_Le", gobjGeracaoMapaCotacao)
    If lErro <> SUCESSO And lErro <> 114583 Then gError 114603
    
    'Se nao encontrou => Erro
    If lErro = 114583 Then gError 114605
    
    'Confirma a exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_MAPA_COTACAO", Codigo.Text)

    'Se a resposta for negativa
    If vbMsgRes = vbYes Then

        'Exclui o Pedido de Compra
        lErro = CF("MapaCotacao_Exclui", gobjGeracaoMapaCotacao.objMapaCotacao)
        If lErro <> SUCESSO Then gError 114604
    
        'Limpa a tela
        Call Limpa_Tela_MapaCotacao
    
        iAlterado = 0
        iFramePrincipalAlterado = 0

    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 114602
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 114603, 114604
            
        Case 114605
            Call Rotina_Erro(vbOKOnly, "ERRO_MAPACOTACAO_INEXISTENTE", gErr, Codigo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162546)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objGeracaoMapaCotacao As New ClassGeracaoMapaCotacao
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimir_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 86260

    objGeracaoMapaCotacao.objMapaCotacao.lCodigo = StrParaDbl(Codigo.Text)
    objGeracaoMapaCotacao.objMapaCotacao.iFilialEmpresa = giFilialEmpresa
    
    'lErro = CF("MapaCotacao_Le", objGeracaoMapaCotacao)
    If lErro <> SUCESSO And lErro <> 114583 Then gError 86261
    If lErro = 114583 Then gError 86262

    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Mapa Cotação", "MapaCota.Codigo = @NMAPA E MapaCota.FilialEmpresa = @NFILIALMAPA", 0, "MapaCota", "NMAPA", objGeracaoMapaCotacao.objMapaCotacao.lCodigo, "NFILIALMAPA", objGeracaoMapaCotacao.objMapaCotacao.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 86263

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 86260
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 86261, 86263

        Case 86262
            Call Rotina_Erro(vbOKOnly, "ERRO_MAPACOTACAO_NAO_CADASTRADO", gErr, Codigo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162547)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todas CheckBox do Item de Categoria

Dim iIndice As Integer

    'Percorre todas os itens de categoria
    For iIndice = 0 To ItensCategoria.ListCount - 1

        'Marca na tela a linha em questão
        ItensCategoria.Selected(iIndice) = MARCADO

    Next
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todas CheckBox do Item de Categoria

Dim iIndice As Integer

    'Percorre todas os itens de categoria
    For iIndice = 0 To ItensCategoria.ListCount - 1

        'Marca na tela a linha em questão
        ItensCategoria.Selected(iIndice) = DESMARCADO

    Next
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItensCategoria_Click()
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItensCategoria_ItemCheck(Item As Integer)
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

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

Private Sub LabelCodPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_LabelCodPCDe_Click

    If Len(Trim(CodPCDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedidoCotacao.lCodigo = StrParaLong(CodPCDe.Text)
    End If

    'Chama Tela PedidoCotacaoTodosLista
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, objPedidoCotacao, objEventoCodPCDe)

   Exit Sub

Erro_LabelCodPCDe_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162548)

    End Select

End Sub

Private Sub LabelCodPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCotacao As New ClassPedidoCotacao

On Error GoTo Erro_LabelCodPCAte_Click

    If Len(Trim(CodPCAte.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCotacao.lCodigo = StrParaLong(CodPCAte.Text)
    End If

    'Chama Tela PedidoCotacaoTodosLista
    Call Chama_Tela("PedidoCotacaoTodosLista", colSelecao, objPedCotacao, objEventoCodPCAte)

   Exit Sub

Erro_LabelCodPCAte_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162549)

    End Select

End Sub

Private Sub CodPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCDe, iAlterado)

End Sub

Private Sub CodPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCAte, iAlterado)

End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataMapaCotacao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataMapaCotacao, iAlterado)

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 114521

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 114521
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162550)

    End Select

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 114522

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 114522
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162551)

    End Select

End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 114523

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 114523
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162552)

    End Select

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 114524

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 114524
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162553)

    End Select

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 114525

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 114525
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162554)

    End Select

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 114526

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 114526
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162555)

    End Select

End Sub

Private Sub Categoria_Click()
'Preenche os itens da categoria selecionada

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Categoria_Click

    'Preenche o Obj
    objCategoriaProduto.sCategoria = Categoria.List(Categoria.ListIndex)

    'Le as categorias do Produto
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO And lErro <> 22541 Then gError 114527

    ItensCategoria.Clear

    For Each objCategoriaProdutoItem In colItensCategoria
        ItensCategoria.AddItem (objCategoriaProdutoItem.sItem)
    Next
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Categoria_Click:

    Select Case gErr

         Case 114527

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162556)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Seta o primeiro frame
    iFrameAtual = 1

    'Inicializa os objs que irao tratar os eventos de browser
    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
    Set objEventoCodigoMapaCotacao = New AdmEvento

    Set gobjGeracaoMapaCotacao = New ClassGeracaoMapaCotacao

    'Inicializa o obj correspondente ao grid
    Set objGridMapaCotacao = New AdmGrid

    'Inicializa o grid
    lErro = Inicializa_Grid_Cotacao(objGridMapaCotacao)
    If lErro <> SUCESSO Then gError 114535

    'Carrega as categorias existentes em BD
    lErro = Carrega_Categorias()
    If lErro <> SUCESSO Then gError 114528

    'Exibe a taxa financeira
    TaxaFinanceira.Caption = Format(gobjCOM.dTaxaFinanceiraEmpresa, "PERCENT")
    
    iFramePrincipalAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 114528, 114535

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162557)

    End Select

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
    Set objGridMapaCotacao = Nothing
    Set objEventoCodigoMapaCotacao = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Function Trata_Parametros(Optional objMapaCotacao As ClassMapaCotacao) As Long
'Trata os parametros

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum mapa de cotacao foi passado por parametro
    If Not (objMapaCotacao Is Nothing) Then

        Set gobjGeracaoMapaCotacao.objMapaCotacao = objMapaCotacao
        
        'Lê o Mapa de Cotacao
        lErro = CF("MapaCotacao_Le", gobjGeracaoMapaCotacao)
        If lErro <> SUCESSO And lErro <> 114583 Then gError 114575

        'Se existe => Traz para a tela ...
        If lErro = SUCESSO Then

            lErro = Traz_MapaCotacao_Tela(gobjGeracaoMapaCotacao)
            If lErro <> SUCESSO Then gError 114576

        End If

    End If

    iAlterado = 0
    iFramePrincipalAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 114575, 114576

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162558)

    End Select

    iAlterado = 0
    iFramePrincipalAlterado = 0

End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código do proximo mapa de cotacao
    lErro = MapaCotacao_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 114528

    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True

    Call Codigo_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 114528

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162559)

    End Select

End Sub

Function MapaCotacao_Automatico(lCodigo As Long) As Long
'Gera o próximo mapa de cotacao

Dim lErro As Long

On Error GoTo Erro_MapaCotacao_Automatico

    lErro = CF("Config_ObterAutomatico", "ComprasConfig", "NUM_PROX_MAPACOTACAO", "MapaCotacao", "NumIntDoc", lCodigo)
    If lErro <> SUCESSO Then gError 114529

    MapaCotacao_Automatico = SUCESSO

    Exit Function

Erro_MapaCotacao_Automatico:

    MapaCotacao_Automatico = gErr

    Select Case gErr

        Case 114529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162560)

    End Select

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 114530

    Call Limpa_Tela_MapaCotacao

    iAlterado = 0
    iFramePrincipalAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 114530

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162561)

    End Select

End Sub

Private Sub Limpa_Tela_MapaCotacao()
'Limpa a tela

    Call Limpa_Tela(Me)

    Categoria.ListIndex = -1
    ItensCategoria.Clear

    Call Grid_Limpa(objGridMapaCotacao)

    Set gobjGeracaoMapaCotacao = New ClassGeracaoMapaCotacao

    BotaoGerarMapa.Enabled = False
    BotaoGeraImprime.Enabled = False

    iAlterado = 0
    iFramePrincipalAlterado = 0
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then

        Call BotaoProxNum_Click

    ElseIf KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is CodPCDe Then
            Call LabelCodPCDe_Click
        ElseIf Me.ActiveControl Is CodPCAte Then
            Call LabelCodPCAte_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If

    End If

End Sub

Private Function Carrega_Categorias() As Long

Dim lErro As Long
Dim objCategoria As New ClassCategoriaProduto
Dim colCategorias As New Collection

On Error GoTo Erro_Carrega_Categorias

    'Le a categoria
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 114531

    Categoria.AddItem ("")

    'Carrega as combos de Categorias
    For Each objCategoria In colCategorias

        Categoria.AddItem objCategoria.sCategoria

    Next

    Carrega_Categorias = SUCESSO

    Exit Function

Erro_Carrega_Categorias:

    Carrega_Categorias = gErr

    Select Case gErr

        Case 114531

        Case 114532
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_CADASTRADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162562)

    End Select

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True

    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    'Se o frame selecionado foi o de Produtos e houve alteracao do Tab de Selecao
    If TabStrip1.SelectedItem.Index = 2 Then

        If iFramePrincipalAlterado = REGISTRO_ALTERADO Then

            Call Grid_Limpa(objGridMapaCotacao)
                
            Set gobjGeracaoMapaCotacao = New ClassGeracaoMapaCotacao
            
            'Recolhe os dados do Tab de Selecao
            lErro = Move_Tela_Memoria(gobjGeracaoMapaCotacao)
            If lErro <> SUCESSO Then gError 114640
            
            'Verifica se o mapa existe
            lErro = CF("MapaCotacao_Le", gobjGeracaoMapaCotacao)
            If lErro <> SUCESSO And lErro <> 114583 Then gError 114641
            
            'Se nao encontrou
            If lErro = 114583 Then
            
                'Faz a leitura dos dd´s
                lErro = CF("MapaCotacao_ObterPedidos", gobjGeracaoMapaCotacao)
                If lErro <> SUCESSO Then gError 114642
            
            End If
    
            'Traz para a tela os Produtos com as características determinadas no Tab Selecao
            lErro = Traz_MapaCotacao_Tela(gobjGeracaoMapaCotacao)
            If lErro <> SUCESSO Then gError 114643
            
        End If

    End If

    iFramePrincipalAlterado = 0

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 114568, 114569, 114567, 114640 To 114643

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162563)

    End Select
    
End Sub


Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub categoria_change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodPCAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodPCDe_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataDe_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItensCategoria_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigoMapaCotacao_evSelecao(obj1 As Object)

Dim lErro As Long

On Error GoTo Erro_objEventoCodigoMapaCotacao_evSelecao

    Set gobjGeracaoMapaCotacao.objMapaCotacao = obj1

    Codigo.PromptInclude = False
    Codigo.Text = CStr(gobjGeracaoMapaCotacao.objMapaCotacao.lCodigo)
    Codigo.PromptInclude = True
    
    'Lê o Mapa de Cotacao
    lErro = CF("MapaCotacao_Le", gobjGeracaoMapaCotacao)
    If lErro <> SUCESSO And lErro <> 114583 Then gError 114631
    
    'Se nao encontrou => Erro
    If lErro = 114583 Then gError 114632
    
    'Traz o mapa de cotacao para a tela
    lErro = Traz_MapaCotacao_Tela(gobjGeracaoMapaCotacao)
    If lErro <> SUCESSO Then gError 114633
    
    Me.Show
    
    iFramePrincipalAlterado = 0
    iAlterado = 0
    
    Call ComandoSeta_Liberar(Me.Name)
    
    Exit Sub
    
Erro_objEventoCodigoMapaCotacao_evSelecao:

    Select Case gErr

        Case 114631, 114633
        
        Case 114632
            Call Rotina_Erro(vbOKOnly, "ERRO_MAPACOTACAO_INEXISTENTE", gErr, Codigo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162564)

    End Select

End Sub

Private Sub objEventoCodPCAte_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao

    Set objPedCotacao = obj1

    CodPCAte.Text = CStr(objPedCotacao.lCodigo)

    Call ComandoSeta_Liberar(Me.Name)
    
    Me.Show
    
End Sub

Private Sub objEventoCodPCDe_evSelecao(obj1 As Object)

Dim objPedCotacao As New ClassPedidoCotacao

    Set objPedCotacao = obj1

    CodPCDe.Text = CStr(objPedCotacao.lCodigo)

    Call ComandoSeta_Liberar(Me.Name)
    
    Me.Show

End Sub

Private Function Inicializa_Grid_Cotacao(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cotacao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Seleciona")
    objGridInt.colColuna.Add ("Ped. Cotação")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial Fornecedor")

    'campos de edição do grid
    objGridInt.colCampo.Add (Seleciona.Name)
    objGridInt.colCampo.Add (PedidoCot.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (Quant.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (FilialForn.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Seleciona_Col = 1
    iGrid_PedidoCot_Col = 2
    iGrid_Produto_Col = 3
    iGrid_Descricao_Col = 4
    iGrid_Quant_Col = 5
    iGrid_UM_Col = 6
    iGrid_Fornecedor_Col = 7
    iGrid_FilialForn_Col = 8

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridMapaCotacao

    'Linhas do grid
    objGridInt.objGrid.Rows = 150 + 1

    'Não permite incluir e excluir linhas do grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    GridMapaCotacao.ColWidth(0) = 400

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cotacao = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cotacao:

    Inicializa_Grid_Cotacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162565)

    End Select

End Function

Private Sub UpDownDataMapaCotacao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataMapaCotacao_DownClick

    'Diminui um dia em DataMapaCotacao
    lErro = Data_Up_Down_Click(DataMapaCotacao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 114538

    Exit Sub

Erro_UpDownDataMapaCotacao_DownClick:

    Select Case gErr

        Case 114538
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162566)

    End Select

End Sub

Private Sub UpDownDataMapaCotacao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataMapaCotacao_UpClick

    'Diminui um dia em DataMapaCotacao
    lErro = Data_Up_Down_Click(DataMapaCotacao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 114537

    Exit Sub

Erro_UpDownDataMapaCotacao_UpClick:

    Select Case gErr

        Case 114537
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 162567)

    End Select

End Sub

Private Sub DataMapaCotacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataMapaCotacao_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataMapaCotacao.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataMapaCotacao.Text)
    If lErro <> SUCESSO Then gError 114536

    Exit Sub

Erro_DataMapaCotacao_Validate:

    Cancel = True

    Select Case gErr

        Case 114536
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162568)

    End Select

End Sub

Function Move_Grid_Memoria(ColItensMapaCotacao As Collection) As Long
'Move os itens selecionados do grid para a colecao

Dim iIndice As Integer
Dim objMapaCotacaoItemCotacao As ClassMapaCotacaoItemCotacao

On Error GoTo Erro_Move_Grid_Memoria

    'Para cada linha do grid
    For iIndice = 1 To objGridMapaCotacao.iLinhasExistentes

        'Se ela estiver selecionada
        If GridMapaCotacao.TextMatrix(iIndice, iGrid_Seleciona_Col) = MARCADO Then

            'Instancia um novo obj
            Set objMapaCotacaoItemCotacao = New ClassMapaCotacaoItemCotacao

            'Armazena os dados necessários
            objMapaCotacaoItemCotacao.iFilialEmpresa = giFilialEmpresa
            objMapaCotacaoItemCotacao.lCodMapaCotacao = StrParaLong(Codigo.Text)
            objMapaCotacaoItemCotacao.lNumIntItemCotacao = gobjGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao.Item(iIndice).lNumIntItemCotacao

            'Adiciona a Colecao de Itens de Mapa de Cotacao
            ColItensMapaCotacao.Add objMapaCotacaoItemCotacao

        End If

    Next

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162569)

    End Select

End Function

Function Move_Tela_Memoria(objGeracaoMapaCotacao As ClassGeracaoMapaCotacao) As Long
'Transfere os dados da tela para o objGeracaoMapaCotacao

Dim lErro As Long
Dim bAchou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Faz a validacao dos campos ...

    If Len(Trim(CodPCDe.ClipText)) > 0 And Len(Trim(CodPCAte.ClipText)) > 0 Then
        If StrParaLong(CodPCDe.Text) > StrParaLong(CodPCAte.Text) Then gError 114545
    End If

    If Len(Trim(DataDe.ClipText)) > 0 And Len(Trim(DataAte.ClipText)) > 0 Then
        If MaskedParaDate(DataDe) > MaskedParaDate(DataAte) Then gError 114546
    End If

    If Len(Trim(Categoria.Text)) > 0 Then

        For iIndice = 0 To ItensCategoria.ListCount - 1
            If ItensCategoria.Selected(iIndice) = True Then
                bAchou = True
                Exit For
            End If
        Next

        If Not bAchou Then gError 114547

    End If

    If Len(Trim(CodPCDe.ClipText)) > 0 Then objGeracaoMapaCotacao.lCodigoDe = StrParaLong(CodPCDe.Text)
    If Len(Trim(CodPCAte.ClipText)) > 0 Then objGeracaoMapaCotacao.lCodigoAte = StrParaLong(CodPCAte.Text)

    If Len(Trim(DataDe.ClipText)) > 0 Then
        objGeracaoMapaCotacao.dtDataDe = MaskedParaDate(DataDe)
    Else
         objGeracaoMapaCotacao.dtDataDe = DATA_NULA
    End If

    If Len(Trim(DataAte.ClipText)) > 0 Then
        objGeracaoMapaCotacao.dtDataAte = MaskedParaDate(DataAte)
    Else
         objGeracaoMapaCotacao.dtDataAte = DATA_NULA
    End If

    If Len(Trim(Categoria.Text)) > 0 Then

        objGeracaoMapaCotacao.sCategoria = Categoria.Text

        For iIndice = 0 To ItensCategoria.ListCount - 1
            If ItensCategoria.Selected(iIndice) = True Then objGeracaoMapaCotacao.colItensCategoria.Add ItensCategoria.List(iIndice)
        Next

    End If

    If Len(Trim(Codigo.ClipText)) > 0 Then objGeracaoMapaCotacao.objMapaCotacao.lCodigo = StrParaLong(Codigo.Text)

    'Guarda a data da gravacao
    If Len(Trim(DataMapaCotacao.ClipText)) > 0 Then
        objGeracaoMapaCotacao.objMapaCotacao.dtData = MaskedParaDate(DataMapaCotacao)
    Else
        objGeracaoMapaCotacao.objMapaCotacao.dtData = gdtDataHoje
    End If

    'Guarda a taxa informada
    objGeracaoMapaCotacao.objMapaCotacao.dTaxaFinanceira = PercentParaDbl(TaxaFinanceira.Caption)

    objGeracaoMapaCotacao.objMapaCotacao.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 114539

        Case 114545
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)

        Case 114546
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case 114547
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_PRODUTO_ITEM_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162570)

    End Select

End Function

Private Sub BotaoGerarMapa_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerarMapa_Click

    'Grava registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 49451

    'Limpa a tela
    Call Limpa_Tela_MapaCotacao

    iAlterado = 0

    Exit Sub

Erro_BotaoGerarMapa_Click:

    Select Case gErr

        Case 49451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162571)

    End Select

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim bAchou  As Boolean
Dim iIndice As Integer
Dim objMapaCotacao As New ClassMapaCotacao
Dim objGeracaoMapaCotacao As New ClassGeracaoMapaCotacao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Codigo foi preenchido
    If (Len(Trim(Codigo.Text))) = 0 Then gError 114540

    'Verifica se a data foi preenchida
    If Len(Trim(DataMapaCotacao.ClipText)) = 0 Then gError 114591

    'Verifica se existem linhas selecionadas no grid de itens de mapa de cotacao
    For iIndice = 1 To objGridMapaCotacao.iLinhasExistentes
        If GridMapaCotacao.TextMatrix(iIndice, iGrid_Seleciona_Col) = MARCADO Then
            bAchou = True
            Exit For
        End If
    Next

    'Se nao encontrou => Erro
    If bAchou = False Then gError 114541

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objGeracaoMapaCotacao)
    If lErro <> SUCESSO Then gError 114542
    
    'Move os dados do Grid para a Colecao de itens de Mapa de Cotacao
    lErro = Move_Grid_Memoria(objGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao)
    If lErro <> SUCESSO Then gError 114539

    'Grava o Mapa de Cotacao
    lErro = CF("MapaCotacao_Grava", objGeracaoMapaCotacao.objMapaCotacao)
    If lErro <> SUCESSO Then gError 114543

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 114540
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 114541
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_MP", gErr)

        Case 114543, 114544

        Case 114591
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162572)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

End Function

Private Sub BotaoMarcarTodosGrid_Click()

Dim iLinha As Integer

    For iLinha = 1 To objGridMapaCotacao.iLinhasExistentes
        GridMapaCotacao.TextMatrix(iLinha, iGrid_Seleciona_Col) = MARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridMapaCotacao)

End Sub

Private Sub BotaoDesmarcarTodosGrid_Click()

Dim iLinha As Integer

    For iLinha = 1 To objGridMapaCotacao.iLinhasExistentes
        GridMapaCotacao.TextMatrix(iLinha, iGrid_Seleciona_Col) = DESMARCADO
    Next

    Call Grid_Refresh_Checkbox(objGridMapaCotacao)

End Sub

Function Traz_MapaCotacao_Tela(ByVal objGeracaoMapaCotacao As ClassGeracaoMapaCotacao) As Long
'Coloca na tela os dados de objMapaCotacao

Dim lErro As Long
Dim iIndice As Integer
Dim objItemMapaCotacao As New ClassMapaCotacaoItemCotacao
Dim sProdutoFormatado As String

On Error GoTo Erro_Traz_MapaCotacao_Tela

    Call Grid_Limpa(objGridMapaCotacao)

    If objGeracaoMapaCotacao.objMapaCotacao.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = objGeracaoMapaCotacao.objMapaCotacao.lCodigo
        Codigo.PromptInclude = True
    End If

    If objGeracaoMapaCotacao.objMapaCotacao.dtData = DATA_NULA Then
        Call DateParaMasked(DataMapaCotacao, gdtDataHoje)
    Else
        Call DateParaMasked(DataMapaCotacao, objGeracaoMapaCotacao.objMapaCotacao.dtData)
    End If
    
    TaxaFinanceira.Caption = Format(objGeracaoMapaCotacao.objMapaCotacao.dTaxaFinanceira, "PERCENT")

    If objGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao.Count >= objGridMapaCotacao.objGrid.Rows Then
        Call Refaz_Grid(objGridMapaCotacao, objGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao.Count + 1)
    End If
    
    For Each objItemMapaCotacao In objGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao

        iIndice = iIndice + 1

        Call Mascara_MascararProduto(objItemMapaCotacao.sProduto, sProdutoFormatado)
        
        GridMapaCotacao.TextMatrix(iIndice, iGrid_Seleciona_Col) = MARCADO
        GridMapaCotacao.TextMatrix(iIndice, iGrid_PedidoCot_Col) = objItemMapaCotacao.lPedidoCotacao
        GridMapaCotacao.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoFormatado
        GridMapaCotacao.TextMatrix(iIndice, iGrid_Descricao_Col) = objItemMapaCotacao.sDescricao
        GridMapaCotacao.TextMatrix(iIndice, iGrid_Quant_Col) = Format(objItemMapaCotacao.dQuantidade, "STANDARD")
        GridMapaCotacao.TextMatrix(iIndice, iGrid_UM_Col) = objItemMapaCotacao.sUM
        GridMapaCotacao.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objItemMapaCotacao.sFornecedor
        GridMapaCotacao.TextMatrix(iIndice, iGrid_FilialForn_Col) = objItemMapaCotacao.iFilialForn & SEPARADOR & objItemMapaCotacao.sNomeFilialForn

    Next

    objGridMapaCotacao.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridMapaCotacao)
    
    'Seleciona o tab 2
    iFrameAtual = 2
    TabStrip1.Tabs.Item(1).Selected = False
    TabStrip1.Tabs.Item(2).Selected = True
    
    Traz_MapaCotacao_Tela = SUCESSO

    Exit Function

Erro_Traz_MapaCotacao_Tela:

    Traz_MapaCotacao_Tela = gErr
    
    Select Case gErr

        Case 114567

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162573)

    End Select
    
End Function

Private Sub GridMapaCotacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMapaCotacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMapaCotacao, iAlterado)
    End If

End Sub

Private Sub GridMapaCotacao_GotFocus()
    Call Grid_Recebe_Foco(objGridMapaCotacao)
End Sub

Private Sub GridMapaCotacao_EnterCell()
    Call Grid_Entrada_Celula(objGridMapaCotacao, iAlterado)
End Sub

Private Sub GridMapaCotacao_LeaveCell()
    Call Saida_Celula(objGridMapaCotacao)
End Sub

Private Sub GridMapaCotacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMapaCotacao)
End Sub

Private Sub GridMapaCotacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMapaCotacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMapaCotacao, iAlterado)
    End If

End Sub

Private Sub GridMapaCotacao_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridMapaCotacao)
End Sub

Private Sub GridMapaCotacao_RowColChange()
    Call Grid_RowColChange(objGridMapaCotacao)
End Sub

Private Sub GridMapaCotacao_Scroll()
    Call Grid_Scroll(objGridMapaCotacao)
End Sub

Private Sub Seleciona_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Seleciona_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMapaCotacao)

End Sub

Private Sub Seleciona_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMapaCotacao)

End Sub

Private Sub Seleciona_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMapaCotacao.objControle = Seleciona
    lErro = Grid_Campo_Libera_Foco(objGridMapaCotacao)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 114573

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 114573
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162574)

    End Select

End Function

Private Sub BotaoPedCotacao_Click()

Dim objPedidoCotacao As New ClassPedidoCotacao
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedCotacao_Click

    'Se nenhuma linha foi selecionada, sai da rotina
    If GridMapaCotacao.Row = 0 Then gError 114574

    objPedidoCotacao.lCodigo = StrParaLong(GridMapaCotacao.TextMatrix(GridMapaCotacao.Row, iGrid_PedidoCot_Col))
    objPedidoCotacao.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("PedidoCotacao", objPedidoCotacao)

    Exit Sub

Erro_BotaoPedCotacao_Click:

    Select Case gErr

        Case 114574
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162575)

    End Select

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objMapaCotacao As New ClassMapaCotacao
Dim objGeracaoMapaCotacao As New ClassGeracaoMapaCotacao

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 114578
    
    objGeracaoMapaCotacao.objMapaCotacao.lCodigo = StrParaLong(Codigo.Text)

    'Faz a leitura do Mapa
    lErro = CF("MapaCotacao_Le", objGeracaoMapaCotacao)
    If lErro <> SUCESSO And lErro <> 114583 Then gError 114590
    
    'Verifica se o codigo preenchido já está sendo usado ... Se estiver => Erro
    If lErro = SUCESSO Then

        BotaoGerarMapa.Enabled = False
        BotaoGeraImprime.Enabled = False

    'Senao, habilita o botao de gerar
    Else

'''        Comentado a pedido da Shirley.
'''        Set gobjGeracaoMapaCotacao = New ClassGeracaoMapaCotacao
'''
'''        'Move o tab de selecao pra memoria
'''        lErro = Move_Tela_Memoria(gobjGeracaoMapaCotacao)
'''        If lErro <> SUCESSO Then gError 114644
'''
'''        'Faz a leitura dos dd´s
'''        lErro = CF("MapaCotacao_ObterPedidos", gobjGeracaoMapaCotacao)
'''        If lErro <> SUCESSO Then gError 114645
'''
'''        'Traz para a tela os Produtos com as características determinadas no Tab Selecao
'''        lErro = Traz_MapaCotacao_Tela(gobjGeracaoMapaCotacao)
'''        If lErro <> SUCESSO Then gError 114646
        
        BotaoGerarMapa.Enabled = True
        BotaoGeraImprime.Enabled = True

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 114578, 114590, 114644 To 114646

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162576)

    End Select

End Sub

Private Sub BotaoGeraImprime_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGeraImprime_Click

    'Grava registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 114620
    
    Call BotaoImprimir_Click

    'Limpa a tela
    Call Limpa_Tela_MapaCotacao

    iAlterado = 0

    Exit Sub

Erro_BotaoGeraImprime_Click:

    Select Case gErr

        Case 114620

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162577)

    End Select

End Sub

Private Sub ImprimeMapa()

''''Dim lErro As Long
''''Dim objPedidoCompra As New ClassPedidoCompras
''''Dim objRelatorio As New AdmRelatorio
''''
''''On Error GoTo Erro_BotaoImprimir_Click
''''
''''    If Len(Trim(Codigo.Text)) = 0 Then gError 86122
''''
''''    lErro = Move_Tela_Memoria(gobjGeracaoMapaCotacao)
''''    If lErro <> SUCESSO Then gError XXX
''''
''''    lErro = CF("MapaCotacao_Le", gobjGeracaoMapaCotacao)
''''    If lErro <> SUCESSO And lErro <> XXX Then gError 76036
''''
''''    If lErro = XXX Then gError 76037
''''
''''    'Executa o relatório
''''    '??? lErro = objRelatorio.ExecutarDireto("Mapa Cotação", "PEDCOMTO.NumIntDoc = @NPEDCOM", 1, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc)
''''    If lErro <> SUCESSO Then gError XXX
''''
''''    Exit Sub
''''
''''Erro_BotaoImprimir_Click:
''''
''''    Select Case gErr
''''
''''        Case 56096, 76036, 76038, 76055
''''
''''        Case 76037
''''            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
''''
''''        Case 76049
''''            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
''''
''''        Case 86122
''''            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162578)
''''
''''    End Select

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    If Len(Trim(Codigo.Text)) > 0 Then gobjGeracaoMapaCotacao.objMapaCotacao.lCodigo = StrParaLong(Codigo.Text)
    
'???    colSelecao.Add giFilialEmpresa
    
    'Chama Tela PedidoCotacaoTodosLista
    Call Chama_Tela("MapaCotacaoLista", colSelecao, gobjGeracaoMapaCotacao.objMapaCotacao, objEventoCodigoMapaCotacao)

   Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 162579)

    End Select

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "MapaCotacao"
    
    'Move todos os dados Presentes na Tela
    lErro = Move_Tela_Memoria(gobjGeracaoMapaCotacao)
    If lErro <> SUCESSO Then gError 114651
    
    'Move todos os dados Presentes no Grid
    lErro = Move_Grid_Memoria(gobjGeracaoMapaCotacao.objMapaCotacao.ColItensMapaCotacao)
    If lErro <> SUCESSO Then gError 114652

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", gobjGeracaoMapaCotacao.objMapaCotacao.lCodigo, 0, "Codigo"
        
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 114651, 114652

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162580)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objRequisicaoCompras As New ClassRequisicaoCompras

On Error GoTo Erro_Tela_Preenche

    Set gobjGeracaoMapaCotacao = New ClassGeracaoMapaCotacao
    
    'Passa os dados da coleção para objRequisicaoCompras
    gobjGeracaoMapaCotacao.objMapaCotacao.lCodigo = colCampoValor.Item("Codigo").vValor
    
    If gobjGeracaoMapaCotacao.objMapaCotacao.lCodigo <> 0 Then
        
        'Le o mapa de cotacao
        lErro = CF("MapaCotacao_Le", gobjGeracaoMapaCotacao)
        If lErro <> SUCESSO And lErro <> 114583 Then gError 114653
        
        'Se encontrou => Erro
        If lErro = SUCESSO Then
            
            'Traz os dados da Requisição para a tela
            lErro = Traz_MapaCotacao_Tela(gobjGeracaoMapaCotacao)
            If lErro <> SUCESSO Then gError 114654
            
        End If

    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 114653, 114654

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162581)

    End Select

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub
