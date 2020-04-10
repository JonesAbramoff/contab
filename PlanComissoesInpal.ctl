VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PlanComissoesInpal 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   10380
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   2
      Left            =   480
      TabIndex        =   29
      Top             =   960
      Width           =   9615
      Begin VB.CheckBox Replicacao 
         Caption         =   "Habilita replicação de linha (F6)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   22
         Top             =   3990
         Width           =   3165
      End
      Begin VB.Frame FrameRegras 
         Caption         =   "Regras"
         Height          =   4545
         Left            =   120
         TabIndex        =   30
         Top             =   30
         Width           =   9255
         Begin MSMask.MaskEdBox TabelaB 
            Height          =   330
            Left            =   6405
            TabIndex        =   20
            Top             =   2205
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   582
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "0%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TabelaA 
            Height          =   330
            Left            =   5460
            TabIndex        =   19
            Top             =   2205
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   582
            _Version        =   393216
            BorderStyle     =   0
            Format          =   "0%"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Cliente 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3600
            TabIndex        =   15
            Top             =   1560
            Width           =   3255
         End
         Begin VB.TextBox RegiaoVenda 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   360
            TabIndex        =   14
            Top             =   1560
            Width           =   3255
         End
         Begin VB.ComboBox FilialCliente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6990
            TabIndex        =   16
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton BotaoConsultaCampo 
            Height          =   495
            Left            =   7680
            Picture         =   "PlanComissoesInpal.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Chama tela de consulta correspondente ao campo selecionado no grid"
            Top             =   3885
            Width           =   1335
         End
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            ItemData        =   "PlanComissoesInpal.ctx":04BA
            Left            =   1440
            List            =   "PlanComissoesInpal.ctx":04CA
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2160
            Width           =   1935
         End
         Begin VB.ComboBox ItemCatProduto 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3360
            TabIndex        =   18
            Top             =   2160
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid GridRegras 
            Height          =   3375
            Left            =   150
            TabIndex        =   21
            Top             =   270
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5953
            _Version        =   393216
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   1
      Left            =   510
      TabIndex        =   31
      Top             =   960
      Width           =   9615
      Begin VB.Frame FrameComissao 
         Caption         =   "Comissão"
         Height          =   2910
         Left            =   240
         TabIndex        =   32
         Top             =   1650
         Width           =   9255
         Begin VB.Frame FrameAjudaCustos 
            Caption         =   "Ajuda de custos"
            Height          =   1215
            Left            =   4560
            TabIndex        =   40
            Top             =   465
            Width           =   4335
            Begin VB.OptionButton AjudaCustosFixa 
               Caption         =   "Valor Fixo"
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
               Left            =   2430
               TabIndex        =   6
               Top             =   360
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton AjudaCustosMinima 
               Caption         =   "Valor Mínimo"
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
               Left            =   2430
               TabIndex        =   7
               Top             =   840
               Width           =   1575
            End
            Begin MSMask.MaskEdBox ValorAjudaCustos 
               Height          =   285
               Left            =   840
               TabIndex        =   5
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label LabelValorAjudaCustos 
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
               Left            =   240
               TabIndex        =   41
               Top             =   630
               Width           =   510
            End
         End
         Begin VB.Frame FrameIncidencia 
            Caption         =   "Incide sobre"
            Height          =   855
            Left            =   360
            TabIndex        =   37
            Top             =   1815
            Width           =   8535
            Begin VB.CheckBox IncideSobreIPI 
               Caption         =   "IPI"
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
               Left            =   7680
               TabIndex        =   27
               Top             =   360
               Width           =   600
            End
            Begin VB.CheckBox IncideSobreOutrasDesp 
               Caption         =   "Outras Desp."
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
               Left            =   5688
               TabIndex        =   13
               Top             =   360
               Width           =   1455
            End
            Begin VB.CheckBox IncideSobreSeguro 
               Caption         =   "Seguro"
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
               Left            =   4161
               TabIndex        =   12
               Top             =   360
               Width           =   990
            End
            Begin VB.CheckBox IncideSobreFrete 
               Caption         =   "Frete"
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
               Left            =   2844
               TabIndex        =   11
               Top             =   360
               Width           =   780
            End
            Begin VB.CheckBox IncideSobreVenda 
               Caption         =   "Venda"
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
               Height          =   255
               Left            =   1437
               TabIndex        =   10
               Top             =   360
               Value           =   1  'Checked
               Width           =   870
            End
            Begin VB.CheckBox IncideSobreTudo 
               Caption         =   "Tudo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   780
            End
         End
         Begin VB.Frame FramePercentuais 
            Caption         =   "Percentuais"
            Height          =   1215
            Left            =   360
            TabIndex        =   33
            Top             =   465
            Width           =   3975
            Begin MSMask.MaskEdBox PercentualEmissao 
               Height          =   315
               Left            =   1200
               TabIndex        =   4
               Top             =   480
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   7
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label PercentualBaixa 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3000
               TabIndex        =   36
               Top             =   480
               Width           =   765
            End
            Begin VB.Label LabelPercentualBaixa 
               AutoSize        =   -1  'True
               Caption         =   "% Baixa:"
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
               Left            =   2160
               TabIndex        =   35
               Top             =   540
               Width           =   735
            End
            Begin VB.Label LabelPercentualEmissao 
               AutoSize        =   -1  'True
               Caption         =   "% Emissão:"
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
               Left            =   120
               TabIndex        =   34
               Top             =   570
               Width           =   960
            End
         End
      End
      Begin VB.Frame FrameIdentificacao 
         Caption         =   "Identificação"
         Height          =   1425
         Left            =   210
         TabIndex        =   38
         Top             =   60
         Width           =   9255
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2115
            Picture         =   "PlanComissoesInpal.ctx":04E8
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Numeração Automática"
            Top             =   330
            Width           =   300
         End
         Begin VB.CheckBox Tecnico 
            Caption         =   "Técnico"
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
            Left            =   6960
            TabIndex        =   3
            Top             =   990
            Width           =   1095
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   330
            Left            =   1260
            TabIndex        =   2
            Top             =   945
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   330
            Left            =   1260
            TabIndex        =   1
            Top             =   315
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigo 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Height          =   225
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   44
            Top             =   420
            Width           =   750
         End
         Begin VB.Label LabelVendedorNome 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2430
            TabIndex        =   42
            Top             =   960
            Width           =   4215
         End
         Begin VB.Label LabelVendedor 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
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
            Height          =   225
            Left            =   315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   1050
            Width           =   960
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7920
      ScaleHeight     =   450
      ScaleWidth      =   2130
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1620
         Picture         =   "PlanComissoesInpal.ctx":05D2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1110
         Picture         =   "PlanComissoesInpal.ctx":0750
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   615
         Picture         =   "PlanComissoesInpal.ctx":0C82
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   105
         Picture         =   "PlanComissoesInpal.ctx":0E0C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Regras"
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
Attribute VB_Name = "PlanComissoesInpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Iniciado por Tulio em 25/02
'Supervisionado por Luiz Gustavo

'******************************************
'Variáveis Globais
'******************************************

Dim iAlterado As Integer
Dim iVendedorAlterado As Integer
Dim iFrameAtual As Integer

Public objGridRegras As AdmGrid

Dim iGrid_RegiaoVenda_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_FilialCliente_Col As Integer
Dim iGrid_CategoriaProduto_Col As Integer
Dim iGrid_ItemCatProduto_Col As Integer
Dim iGrid_TabelaA_Col As Integer
Dim iGrid_TabelaB_Col As Integer

'******************************************
'Constantes
'******************************************

Private Const TAB_Principal = 1
Private Const TAB_Regras = 2

'Constante feita somente para melhorar a legibilidade do código.
Private Const PERCENTUAL_ZERO = "0,00%"

'Inicialização do grid
Private Const LINHAS_VISIVEIS_GRIDREGRAS = 8
Private Const LARGURA_PRIMEIRA_COLUNA_GRIDREGRAS = 400

'Eventos browser
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoRegiao As AdmEvento
Attribute objEventoRegiao.VB_VarHelpID = -1
Private WithEvents objEventoComissoesInpalPlan As AdmEvento
Attribute objEventoComissoesInpalPlan.VB_VarHelpID = -1

'Indica a rotina que está chamando a função Trata_Combo_FilialCliente / Trata_Combo_ItemCatProduto
Const FUNCAO_CATEGORIAPRODUTO_KEYPRESS = -1
Const FUNCAO_CLIENTE_KEYPRESS = -1
Const FUNCAO_ROTINA_GRID_ENABLE = 2

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Planilha de Comissões"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "PlanComissoesInpal"

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

Private Sub AjudaCustosFixa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AjudaCustosMinima_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IncideSobreFrete_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If IncideSobreFrete.Value = vbUnchecked And IncideSobreTudo = vbChecked Then
        IncideSobreTudo.Value = vbUnchecked
        IncideSobreSeguro.Value = vbChecked
        IncideSobreIPI.Value = vbChecked
        IncideSobreOutrasDesp.Value = vbChecked
    End If

    If IncideSobreFrete.Value = vbChecked And IncideSobreIPI.Value = vbChecked And IncideSobreOutrasDesp = vbChecked And IncideSobreSeguro.Value = vbChecked Then IncideSobreTudo.Value = vbChecked

End Sub

Private Sub IncideSobreIPI_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If IncideSobreIPI.Value = vbUnchecked And IncideSobreTudo = vbChecked Then
        IncideSobreTudo.Value = vbUnchecked
        IncideSobreFrete.Value = vbChecked
        IncideSobreSeguro.Value = vbChecked
        IncideSobreOutrasDesp.Value = vbChecked
    End If

    If IncideSobreFrete.Value = vbChecked And IncideSobreIPI.Value = vbChecked And IncideSobreOutrasDesp = vbChecked And IncideSobreSeguro.Value = vbChecked Then IncideSobreTudo.Value = vbChecked

End Sub

Private Sub IncideSobreOutrasDesp_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If IncideSobreOutrasDesp.Value = vbUnchecked And IncideSobreTudo = vbChecked Then
        IncideSobreTudo.Value = vbUnchecked
        IncideSobreFrete.Value = vbChecked
        IncideSobreIPI.Value = vbChecked
        IncideSobreSeguro.Value = vbChecked
    End If

    If IncideSobreFrete.Value = vbChecked And IncideSobreIPI.Value = vbChecked And IncideSobreOutrasDesp = vbChecked And IncideSobreSeguro.Value = vbChecked Then IncideSobreTudo.Value = vbChecked

End Sub

Private Sub IncideSobreSeguro_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If IncideSobreSeguro.Value = vbUnchecked And IncideSobreTudo = vbChecked Then
        IncideSobreTudo.Value = vbUnchecked
        IncideSobreFrete.Value = vbChecked
        IncideSobreIPI.Value = vbChecked
        IncideSobreOutrasDesp.Value = vbChecked
    End If

    If IncideSobreFrete.Value = vbChecked And IncideSobreIPI.Value = vbChecked And IncideSobreOutrasDesp = vbChecked And IncideSobreSeguro.Value = vbChecked Then IncideSobreTudo.Value = vbChecked

End Sub

Private Sub IncideSobreTudo_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If IncideSobreTudo.Value = vbChecked Then
    
        IncideSobreFrete = vbChecked
        IncideSobreIPI = vbChecked
        IncideSobreOutrasDesp = vbChecked
        IncideSobreSeguro = vbChecked
    
    Else
        
        IncideSobreFrete = vbUnchecked
        IncideSobreIPI = vbUnchecked
        IncideSobreOutrasDesp = vbUnchecked
        IncideSobreSeguro = vbUnchecked

    End If

End Sub

Private Sub IncideSobreVenda_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If IncideSobreVenda.Value = vbUnchecked Then IncideSobreTudo.Value = vbUnchecked
    

End Sub

Private Sub LabelCodigo_Click()

Dim objComissoesInpalPlan As New ClassComisInpalPlan
Dim colSelecao As Collection

    'coloca no obj o conteudo do campo codigo caso seja preenchido...
    If Len(Trim(Codigo.Text)) > 0 Then objComissoesInpalPlan.lCodigo = StrParaLong(Codigo.Text)
        
    'Chama Tela de browser
    Call Chama_Tela("PlanComissoesLista", colSelecao, objComissoesInpalPlan, objEventoComissoesInpalPlan)

End Sub

Private Sub LabelPercentualBaixa_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(LabelPercentualBaixa, Source, X, Y)

End Sub

Private Sub LabelPercentualBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Controle_MouseDown(LabelPercentualBaixa, Button, Shift, X, Y)

End Sub

Private Sub LabelPercentualEmissao_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(LabelPercentualEmissao, Source, X, Y)

End Sub

Private Sub LabelPercentualEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Controle_MouseDown(LabelPercentualEmissao, Button, Shift, X, Y)

End Sub

Private Sub LabelValorAjudaCustos_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(LabelValorAjudaCustos, Source, X, Y)

End Sub

Private Sub LabelValorAjudaCustos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Controle_MouseDown(LabelValorAjudaCustos, Button, Shift, X, Y)

End Sub

Private Sub LabelVendedor_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(LabelVendedor, Source, X, Y)

End Sub

Private Sub LabelVendedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call Controle_MouseDown(LabelVendedor, Button, Shift, X, Y)

End Sub

Private Sub LabelVendedorNome_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(LabelVendedorNome, Source, X, Y)

End Sub

Private Sub LabelVendedorNome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Controle_MouseDown(LabelVendedorNome, Button, Shift, X, Y)

End Sub

Private Sub objEventoComissoesInpalPlan_evSelecao(obj1 As Object)

Dim objComissoesInpalPlan As ClassComisInpalPlan
Dim lErro As Long

On Error GoTo Erro_objEventoComissoesInpalPlan_evSelecao

    Set objComissoesInpalPlan = obj1

    'le as regras...
    lErro = CF("ComissoesInpalRegras_Le_CodPlanilha", objComissoesInpalPlan.lCodigo, objComissoesInpalPlan.colComissoesInpalRegras)
    If lErro <> SUCESSO And lErro <> 98785 Then gError 98965
    
    'se nao achou nenhuma regra relacionada com a planilha em questao => erro
    If lErro <> SUCESSO Then gError 98966
    
    'limpa a tela..
    Call Limpa_Tela_PlanComissoesInpal
    
    'Move os dados para a tela
    lErro = Traz_PlanComissoes_Tela(objComissoesInpalPlan)
    If lErro <> SUCESSO Then gError 98963
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoComissoesInpalPlan_evSelecao:

    Select Case gErr

        Case 98963, 98965
           
        Case 98966
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANILHA_SEM_REGRAS", gErr, objComissoesInpalPlan.lCodigo)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub PercentualBaixa_DragDrop(Source As Control, X As Single, Y As Single)

    Call Controle_DragDrop(PercentualBaixa, Source, X, Y)

End Sub

Private Sub PercentualBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Controle_MouseDown(PercentualBaixa, Button, Shift, X, Y)

End Sub

Private Sub PercentualEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentualEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentualEmissao_Validate

    'se o percentual de emissao estiver preenchido
    If Len(Trim(PercentualEmissao.Text)) > 0 Then
    
        'verifica se eh um valor valido
        lErro = Porcentagem_Critica(PercentualEmissao.Text)
        If lErro <> SUCESSO Then gError 98811
           
        'coloca o percentual da emissao no formato adequado
        PercentualEmissao.Text = Format(StrParaDbl(PercentualEmissao.ClipText), "standard")
        
        'calcula o percentual na baixa
        'que eh = a 100 menos o percentual na emissao e ja coloca no formato adequado
        PercentualBaixa.Caption = Format(1 - StrParaDbl(PercentualEmissao.ClipText) / 100, "percent")
    
    Else
        PercentualBaixa.Caption = STRING_VAZIO
    
    End If

    Exit Sub

Erro_PercentualEmissao_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 98811
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    Exit Sub

End Sub

'******************************************
'4 eventos do controle do Grid: RegiaoVenda
'******************************************

Private Sub RegiaoVenda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub RegiaoVenda_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub RegiaoVenda_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub RegiaoVenda_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = RegiaoVenda
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: RegiaoVenda
'******************************************

'******************************************
'4 eventos do controle do Grid: Cliente
'******************************************

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)

    'Se pressionou a tecla ESC, o controle ativo é Cliente
    If KeyAscii = vbKeyEscape Then
        Call Trata_Combo_FilialCliente(FUNCAO_CLIENTE_KEYPRESS)
    End If
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = Cliente
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: Cliente
'******************************************

'******************************************
'4 eventos do controle do Grid: FilialCliente
'******************************************

Private Sub FilialCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialCliente_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub FilialCliente_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = FilialCliente
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: FilialCliente
'******************************************

'******************************************
'4 eventos do controle do Grid: CategoriaProduto
'******************************************

Private Sub CategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CategoriaProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub CategoriaProduto_KeyPress(KeyAscii As Integer)

Dim lErro As Long

On Error GoTo Erro_CategoriaProduto_KeyPress

    'Se pressionou a tecla ESC, o controle ativo é CategoriaProduto
    If KeyAscii = vbKeyEscape Then
        lErro = Trata_Combo_ItemCatProduto(FUNCAO_CATEGORIAPRODUTO_KEYPRESS)
        If lErro <> SUCESSO Then gError 102999
    End If

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)
    
    Exit Sub

Erro_CategoriaProduto_KeyPress:

    Select Case gErr
    
        Case 102999
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = CategoriaProduto
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: CategoriaProduto
'******************************************

'******************************************
'4 eventos do controle do Grid: ItemCatProduto
'******************************************

Private Sub ItemCatProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCatProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub ItemCatProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub ItemCatProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = ItemCatProduto
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: ItemCatProduto
'******************************************

'******************************************
'4 eventos do controle do Grid: TabelaA
'******************************************

Private Sub TabelaA_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabelaA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub TabelaA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub TabelaA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = TabelaA
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: TabelaA
'******************************************


'******************************************
'4 eventos do controle do Grid: TabelaB
'******************************************

Private Sub TabelaB_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabelaB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRegras)

End Sub

Private Sub TabelaB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRegras)

End Sub

Private Sub TabelaB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRegras.objControle = TabelaB
    lErro = Grid_Campo_Libera_Foco(objGridRegras)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'******************************************
'Fim 4 eventos do controle do Grid: TabelaB
'******************************************

Private Sub Tecnico_Click()

    iAlterado = REGISTRO_ALTERADO

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
'''    m_Caption = New_Caption
End Property

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
                                
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'***** fim do trecho a ser copiado ******

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Function Trata_Parametros(Optional objPlanComissoes As ClassComisInpalPlan) As Long
'Trata os parametros passados para a tela..
'Criada em 25/02 por Tulio

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objPlanComissoes estiver preenchido...
    If Not (objPlanComissoes Is Nothing) Then

        'Le os dados da planilha de comissoes
        lErro = CF("ComissoesInpalPlan_Le_Completa", objPlanComissoes)
        If lErro <> SUCESSO And lErro <> 98792 And lErro <> 98794 Then gError 98760
        
        'Se encontrou
        If lErro = SUCESSO Then
            
            'Traz a planilha de comissoes para a tela
            lErro = Traz_PlanComissoes_Tela(objPlanComissoes)
            If lErro <> SUCESSO Then gError 98761
            
        'Se nao encontrou planilha
        ElseIf lErro = 98792 Then
            
            'limpar a tela
            Call Limpa_Tela_PlanComissoesInpal
            
            'Colocar o codigo na tela
            Vendedor.Text = objPlanComissoes.iVendedor
        
            'Chama o validate do vendedor
            Call Vendedor_Validate(bSGECancelDummy)
        
        Else
            'nao achou regras associadas a planilha
            gError 98796
        
        End If
        
    End If
    
    iAlterado = 0
    iVendedorAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr

        Case 98760, 98761

        Case 98796
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANILHA_SEM_REGRAS", gErr, objPlanComissoes.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    iAlterado = 0
    iVendedorAlterado = 0
    
    Exit Function

End Function

Private Function Traz_PlanComissoes_Tela(objPlanComissoes As ClassComisInpalPlan) As Long
'Recebe um objPlanComissoes carregado e
'coloca as informacoes na tela.
'Criada em 25/02 por Tulio

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Traz_PlanComissoes_Tela

    'coloca o codigo
    Codigo.Text = objPlanComissoes.lCodigo
    
    'Coloca a chave em objVendedor
    objVendedor.iCodigo = objPlanComissoes.iVendedor
    
    'Lê os dados do vendedor para obter o nome reduzido
    lErro = CF("Vendedor_Le", objVendedor)
    If lErro <> SUCESSO And lErro <> 12582 Then gError 98763
    
    'Se não é sucesso => significa que não encontrou o vendedor
    If lErro <> SUCESSO Then gError 98770
    
    'traz o codigo do vendedor
    Vendedor.Text = objPlanComissoes.iVendedor
    
    'traz o nome reduzido do vendedor
    LabelVendedorNome.Caption = objVendedor.sNomeReduzido
    
    'verifica se vendedor eh tecnico
    Tecnico.Value = objPlanComissoes.iTecnico
    
    'Traz o valor da ajuda de custo
    ValorAjudaCustos.Text = Format(objPlanComissoes.dAjudaCusto, "standard")
    
    'Traz o tipo da ajuda de custos
    If objPlanComissoes.iTipoAjudaCusto = AJUDACUSTO_MINIMA Then
        AjudaCustosMinima.Value = True
        AjudaCustosFixa.Value = False
    Else
        AjudaCustosMinima.Value = False
        AjudaCustosFixa.Value = True
    End If
    
    'traz o percentual da emissao
    PercentualEmissao.Text = objPlanComissoes.dPercComissaoEmissao * 100
    
    'traz o percentual da baixa
    PercentualBaixa.Caption = Format(objPlanComissoes.dPercComissaoBaixa, "percent")
    
    'traz as checkboxes restantes...
    IncideSobreFrete = objPlanComissoes.iComissaoFrete
    IncideSobreIPI = objPlanComissoes.iComissaoIPI
    IncideSobreSeguro = objPlanComissoes.iComissaoSeguro
    IncideSobreTudo = objPlanComissoes.iComissaoSobreTotal
    IncideSobreOutrasDesp = objPlanComissoes.iComissaoDesp
        
    'Traz as informacoes do grid para a tela
    lErro = Traz_GridRegras_Tela(objPlanComissoes.colComissoesInpalRegras)
    If lErro <> SUCESSO Then gError 98764
    
    Traz_PlanComissoes_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_PlanComissoes_Tela:

    Traz_PlanComissoes_Tela = gErr
    
    Select Case gErr
    
        Case 98763, 98764
        
        Case 98770
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
        
    End Select
        
    Exit Function

End Function

Private Function Traz_GridRegras_Tela(colRegras As Collection) As Long
'Traz para tela as regras associadas a Planilha de Comissoes
'colRegras RECEBE (Input) os dados que serão exibidos

Dim lErro As Long
Dim objRegra As ClassComisInpalRegras
Dim objRegiaoVenda As New ClassRegiaoVenda
Dim objcliente As New ClassCliente
Dim iLinha As Integer, iIndex As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Traz_GridRegras_Tela

    For Each objRegra In colRegras
    
        'incrementa o numero de linhas
        iLinha = iLinha + 1
            
        'Se o Codigo da Regiao de Venda estiver preenchido
        If objRegra.iRegiaoVenda <> CODIGO_NAO_PREENCHIDO Then
            
            'Coloca a chave no objRegiaoVenda
            objRegiaoVenda.iCodigo = objRegra.iRegiaoVenda
            
            'Le a regiao de venda
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 98765
    
            'se nao achou a regiao de venda, erro... inconsistencia no objregra
            If lErro <> SUCESSO Then gError 98774
    
            'Coloca na tela o codigo concatenado com a descricao da regiao
            GridRegras.TextMatrix(iLinha, iGrid_RegiaoVenda_Col) = objRegiaoVenda.iCodigo & SEPARADOR & objRegiaoVenda.sDescricao
            
        End If
        
        'Se o Codigo do Cliente estiver preenchido
        If objRegra.lCliente <> CODIGO_NAO_PREENCHIDO Then
        
            'Coloca a chave no objCliente
            objcliente.lCodigo = objRegra.lCliente
            
            'Le o cliente
            lErro = CF("Cliente_Le", objcliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 98766
    
            'se nao achou, erro...
            If lErro <> SUCESSO Then gError 98775
    
            'Coloca na tela o codigo concatenado com o nomereduzido do cliente
            GridRegras.TextMatrix(iLinha, iGrid_Cliente_Col) = objcliente.lCodigo & SEPARADOR & objcliente.sNomeReduzido
            
            'Se o codigo da filial estiver preenchido
            If objRegra.iFilialCliente <> CODIGO_NAO_PREENCHIDO Then
            
                'coloca a chave em objfilialcliente
                objFilialCliente.lCodCliente = objRegra.lCliente
                objFilialCliente.iCodFilial = objRegra.iFilialCliente
                
                'Le a filial no bd afim de obter o nome da mesma
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 98791
                
                'Seta textmatrix igual a filial do BD => exibe o codigo da filial e o nome
                GridRegras.TextMatrix(iLinha, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
                
            End If
        
        End If

        'Se categoria estiver preenchida
        If Len(Trim(objRegra.sCategoriaProduto)) > 0 Then
                        
            'Traz a categoria
            GridRegras.TextMatrix(iLinha, iGrid_CategoriaProduto_Col) = objRegra.sCategoriaProduto
            
            'Coloca a chave em objCategoriaProdutoItem
            objCategoriaProdutoItem.sCategoria = objRegra.sCategoriaProduto
            objCategoriaProdutoItem.sItem = objRegra.sItemCatProduto
                        
            'Le o item da categoria
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then gError 98772
                        
            'se item nao existe no bd, erro --> nao pode..
            If lErro <> SUCESSO Then gError 98773
            
            'Exibe o item na tela
            GridRegras.TextMatrix(iLinha, iGrid_ItemCatProduto_Col) = objCategoriaProdutoItem.sItem & SEPARADOR & objCategoriaProdutoItem.sDescricao
    
        End If
        
        'Exibe o percentual da tabela A
        GridRegras.TextMatrix(iLinha, iGrid_TabelaA_Col) = Format(objRegra.dPercTabelaA, "percent")
       
        'Se percentual da tabela B estiver preenchido, exibe na tela
        If objRegra.dPercTabelaB <> PERCENTUAL_NAO_PREENCHIDO Then
            GridRegras.TextMatrix(iLinha, iGrid_TabelaB_Col) = Format(objRegra.dPercTabelaB, "percent")
        End If
    
    Next

    objGridRegras.iLinhasExistentes = iLinha

    Traz_GridRegras_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_GridRegras_Tela:

    Traz_GridRegras_Tela = gErr
    
    Select Case gErr
        
        Case 98765, 98766, 98772, 98778, 98791
        
        Case 98773
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CATEGORIA_NAO_CADASTRADO", gErr, objCategoriaProdutoItem.sItem, objCategoriaProdutoItem.sCategoria)
        
        Case 98774
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", gErr, objRegiaoVenda.iCodigo)
        
        Case 98775
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objcliente.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
        
    End Select

    Exit Function

End Function

Private Sub ValorAjudaCustos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorAjudaCustos_Validate(Cancel As Boolean)
  
Dim lErro As Long

On Error GoTo Erro_ValorAjudaCustos_Validate

    'se ValorAjudaCustos estiver preenchido
    If Len(Trim(ValorAjudaCustos.ClipText)) > 0 Then
        
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorAjudaCustos.ClipText)
        If lErro <> SUCESSO Then gError 98958
        
        'coloca no formato adequado
        ValorAjudaCustos.Text = Format(ValorAjudaCustos.Text, "standard")
    
    End If
    
    Exit Sub

Erro_ValorAjudaCustos_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 98958
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  

End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedorAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Vendedor_GotFocus()

    Call MaskEdBox_TrataGotFocus(Vendedor, iAlterado)

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    'Se o codigo do vendedor nao foi alterado, sai
    If iVendedorAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Se o codigo do vendedor esta preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        'passa a chave para objvendedor
        objVendedor.iCodigo = StrParaInt(Vendedor.Text)
        
        'Le dados do vendedor no BD
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 98779
        
        'se nao achou vendedor
        If lErro <> SUCESSO Then gError 98803
        
        'Traz o vendedor pra tela
        lErro = Traz_Vendedor_Tela(objVendedor)
        If lErro <> SUCESSO Then gError 98770
    
    End If
    
    Exit Sub

Erro_Vendedor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 98770, 98779
        
        Case 98803
               
            'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
            If Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR1", objVendedor.iCodigo) = vbYes Then
                'Chama tela de Vendedores
                lErro = Chama_Tela("Vendedores", objVendedor)
            End If
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
            
    End Select

    Exit Sub
    
End Sub

Private Function Traz_Vendedor_Tela(objVendedor As ClassVendedor) As Long
'Traz os dados do vendedor para a tela

On Error GoTo Erro_Traz_Vendedor_Tela

    'traz o codigo, em alguns casos pode ser ambiguo..
    Vendedor.Text = objVendedor.iCodigo
    
    'traz o nome reduzido
    LabelVendedorNome.Caption = objVendedor.sNomeReduzido
    
    'traz os percentuais de emissao e baixa
    PercentualEmissao.Text = CStr(objVendedor.dPercComissaoEmissao * 100)
    PercentualBaixa.Caption = Format(objVendedor.dPercComissaoBaixa, "percent")
    
    'marca as checks de "incide sobre"
    IncideSobreVenda.Value = vbChecked 'sempre verdadeiro
    IncideSobreIPI = objVendedor.iComissaoIPI
    IncideSobreFrete = objVendedor.iComissaoFrete
    IncideSobreTudo = objVendedor.iComissaoSobreTotal
    IncideSobreSeguro = objVendedor.iComissaoSeguro
    IncideSobreOutrasDesp = objVendedor.iComissaoICM
    
    iVendedorAlterado = 0
    
    Traz_Vendedor_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Vendedor_Tela:

    Traz_Vendedor_Tela = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro("vbokonly", "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    iVendedorAlterado = 0
    
    Exit Function

End Function

Private Sub Form_Load()

Dim lErro As Long
Dim iIndiceFrame As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_Principal
    
    'Exibe o tab principal
    Frame1(TAB_Principal).Visible = True
    
    'Oculta o tab regras
    Frame1(TAB_Regras).Visible = False
    
    'Inicializa os Eventos
    Set objEventoCliente = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoRegiao = New AdmEvento
    Set objEventoComissoesInpalPlan = New AdmEvento
    
    Set objGridRegras = New AdmGrid
    
    'Executa inicializacao do Grid
    lErro = Inicializa_GridRegras(objGridRegras)
    If lErro <> SUCESSO Then gError 98799

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 98799, 98780
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridRegras(objGridInt As AdmGrid) As Long
'Inicializa o grid da tela

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridRegras

    'Tela em questão
    Set objGridInt.objForm = Me

    'Titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Região de Venda")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Categoria de Produto")
    objGridInt.colColuna.Add ("Item da Categoria")
    objGridInt.colColuna.Add ("Tabela A")
    objGridInt.colColuna.Add ("Tabela B")

    'campos de edição do grid
    objGridInt.colCampo.Add (RegiaoVenda.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (FilialCliente.Name)
    objGridInt.colCampo.Add (CategoriaProduto.Name)
    objGridInt.colCampo.Add (ItemCatProduto.Name)
    objGridInt.colCampo.Add (TabelaA.Name)
    objGridInt.colCampo.Add (TabelaB.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_RegiaoVenda_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_FilialCliente_Col = 3
    iGrid_CategoriaProduto_Col = 4
    iGrid_ItemCatProduto_Col = 5
    iGrid_TabelaA_Col = 6
    iGrid_TabelaB_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRegras

    'Numero Maximo de Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_REGRAS_COMISSOES_INPAL

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = LINHAS_VISIVEIS_GRIDREGRAS

    'Largura da primeira coluna
    GridRegras.ColWidth(0) = LARGURA_PRIMEIRA_COLUNA_GRIDREGRAS

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
    'seta execucao da rotina grid enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
       
    Inicializa_GridRegras = SUCESSO

    Exit Function

Erro_Inicializa_GridRegras:

    Inicializa_GridRegras = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sCodFilial As String
Dim sCodItem As String
Dim colCodigoNome As New AdmColCodigoNome
Dim objcliente As New ClassCliente
Dim iIndex As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Seleciona o controle atual
    Select Case objControl.Name

        'Se for FilialCliente
        Case FilialCliente.Name
        
            If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_Cliente_Col))) = 0 Then
            
                'Desabilita a coluna filialcliente
                FilialCliente.Enabled = False

            Else

                'Habilita a coluna filialcliente
                FilialCliente.Enabled = True
                
                Call Trata_Combo_FilialCliente(iLocalChamada)
                
            End If
            
        'Se for item cat produto
        Case ItemCatProduto.Name

            'Se categoria nao estiver preenchida
            If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_CategoriaProduto_Col))) = 0 Then
                  'desabilita coluna item de categoria de produto
                  ItemCatProduto.Enabled = False
            Else
                  'habilita coluna item de categoria de produto
                  ItemCatProduto.Enabled = True

                'Carrega a combo de cliente e valida a opção selecionada
                lErro = Trata_Combo_ItemCatProduto(iLocalChamada)
                If lErro <> SUCESSO Then gError 102998
                
            End If
            
        'Se for Tabela B
        Case TabelaB.Name
        
            'Se a Tabela A estiver preenchida
            If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_TabelaA_Col))) > 0 Then
                'Habilita coluna tabela B
                TabelaB.Enabled = True
            Else
                'Desabilita coluna tabela B
                TabelaB.Enabled = False
            End If
                        
    End Select
    
    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 98782, 98783, 102998
    
        Case 98804
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objcliente.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objComissoesInpalPlan As New ClassComisInpalPlan

On Error GoTo Erro_Tela_Extrai

    'Guarda na variavel o nome da tabela q sera lida...
    sTabela = "ComissoesInpalPlan"

    'Copia para a memoria (buffer) os dados da tela
    lErro = Move_Tela_Memoria(objComissoesInpalPlan)
    If lErro <> SUCESSO Then gError 98784

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objComissoesInpalPlan.lCodigo, 0, "Codigo"
    colCampoValor.Add "Vendedor", objComissoesInpalPlan.iVendedor, 0, "Vendedor"
    colCampoValor.Add "Tecnico", objComissoesInpalPlan.iTecnico, 0, "Tecnico"
    colCampoValor.Add "PercComissaoEmissao", objComissoesInpalPlan.dPercComissaoEmissao, 0, "PercComissaoEmissao"
    colCampoValor.Add "PercComissaoBaixa", objComissoesInpalPlan.dPercComissaoBaixa, 0, "PercComissaoBaixa"
    colCampoValor.Add "ComissaoSobreTotal", objComissoesInpalPlan.iComissaoSobreTotal, 0, "ComissaoSobreTotal"
    colCampoValor.Add "ComissaoFrete", objComissoesInpalPlan.iComissaoFrete, 0, "ComissaoFrete"
    colCampoValor.Add "ComissaoDesp", objComissoesInpalPlan.iComissaoDesp, 0, "ComissaoDesp"
    colCampoValor.Add "ComissaoIPI", objComissoesInpalPlan.iComissaoIPI, 0, "ComissaoIPI"
    colCampoValor.Add "ComissaoSeguro", objComissoesInpalPlan.iComissaoSeguro, 0, "ComissaoSeguro"
    colCampoValor.Add "AjudaCusto", objComissoesInpalPlan.dAjudaCusto, 0, "AjudaCusto"
    colCampoValor.Add "TipoAjudaCusto", objComissoesInpalPlan.iTipoAjudaCusto, 0, "TipoAjudaCusto"
            
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 98784

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objComissoesInpalPlan As New ClassComisInpalPlan
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'se codigo do vendedor estiver preenchido...
    If colCampoValor.Item("Vendedor").vValor > 0 Then
    
        'carrega o obj com as informacoes de colcampovalor
        objComissoesInpalPlan.lCodigo = colCampoValor.Item("Codigo").vValor
        objComissoesInpalPlan.dAjudaCusto = colCampoValor.Item("AjudaCusto").vValor
        objComissoesInpalPlan.dPercComissaoBaixa = colCampoValor.Item("PercComissaoBaixa").vValor
        objComissoesInpalPlan.dPercComissaoEmissao = colCampoValor.Item("PercComissaoEmissao").vValor
        objComissoesInpalPlan.iComissaoDesp = colCampoValor.Item("ComissaoDesp").vValor
        objComissoesInpalPlan.iComissaoFrete = colCampoValor.Item("ComissaoFrete").vValor
        objComissoesInpalPlan.iComissaoIPI = colCampoValor.Item("ComissaoIPI").vValor
        objComissoesInpalPlan.iComissaoSeguro = colCampoValor.Item("ComissaoSeguro").vValor
        objComissoesInpalPlan.iComissaoSobreTotal = colCampoValor.Item("ComissaoSobreTotal").vValor
        objComissoesInpalPlan.iTecnico = colCampoValor.Item("Tecnico").vValor
        objComissoesInpalPlan.iTipoAjudaCusto = colCampoValor.Item("TipoAjudaCusto").vValor
        objComissoesInpalPlan.iVendedor = colCampoValor.Item("Vendedor").vValor
                
        'le as regras para preencher o grid posteriormente
        lErro = CF("ComissoesInpalRegras_Le_CodPlanilha", objComissoesInpalPlan.lCodigo, objComissoesInpalPlan.colComissoesInpalRegras)
        If lErro <> SUCESSO And lErro <> 98785 Then gError 98786
    
        'se nao achou regra(s) relacionada(s) a planilha em questao...
        If lErro <> SUCESSO Then gError 98806
        
        'limpa a tela
        Call Limpa_Tela_PlanComissoesInpal
        
        'traz para a tela a planilha
        lErro = Traz_PlanComissoes_Tela(objComissoesInpalPlan)
        If lErro <> SUCESSO Then gError 98787
    
    End If
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 98786, 98787

        Case 98806
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANILHA_SEM_REGRAS", gErr, objComissoesInpalPlan.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Carrega em objComissoesInpalPlan os dados da tela
'objComissoesInpalPlan RETORNA(OUPUT) os dados da tela

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'bota o codigo no obj
    objComissoesInpalPlan.lCodigo = StrParaLong(Codigo.Text)
    
    'guarda o codigo do vendedor no obj
    objComissoesInpalPlan.iVendedor = StrParaInt(Vendedor.Text)
    
    'guarda a informacao de tecnico
    objComissoesInpalPlan.iTecnico = Tecnico.Value

    'guarda os percentuais de emissao e baixa
    objComissoesInpalPlan.dPercComissaoEmissao = PercentParaDbl(PercentualEmissao.FormattedText)
    objComissoesInpalPlan.dPercComissaoBaixa = PercentParaDbl(PercentualBaixa.Caption)
    
    'guarda o valor da ajuda de custos
    objComissoesInpalPlan.dAjudaCusto = StrParaDbl(ValorAjudaCustos.Text)
    
    'se o valor da ajuda de custo for maior que 0
    If objComissoesInpalPlan.dAjudaCusto > 0 Then
        
        'guarda o tipo de ajuda de custos
        If AjudaCustosFixa.Value = True Then
            objComissoesInpalPlan.iTipoAjudaCusto = AJUDACUSTO_FIXA
        ElseIf AjudaCustosMinima.Value = True Then
            objComissoesInpalPlan.iTipoAjudaCusto = AJUDACUSTO_MINIMA
        End If
        
    'senao, o tipo de ajuda deve ser, obrigatoriamente, fixa
    Else
        
        objComissoesInpalPlan.iTipoAjudaCusto = AJUDACUSTO_FIXA
    
    End If

    'verifica as checks que estao marcadas e grava as informacoes no obj
    objComissoesInpalPlan.iComissaoDesp = IncideSobreOutrasDesp.Value
    objComissoesInpalPlan.iComissaoFrete = IncideSobreFrete.Value
    objComissoesInpalPlan.iComissaoIPI = IncideSobreIPI.Value
    objComissoesInpalPlan.iComissaoSeguro = IncideSobreSeguro.Value
    objComissoesInpalPlan.iComissaoSobreTotal = IncideSobreTudo.Value
    
    'carrega no obj o conteudo do grid de regras
    lErro = Move_GridRegras_Memoria(objComissoesInpalPlan)
    If lErro <> SUCESSO Then gError 98789

    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 98789
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Function

End Function

Private Function Move_GridRegras_Memoria(objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Move os dados do grid para a memoria (colecao dentro do obj)
'objComissoesInpalPlan RETORNA(OUTPUT) as informacoes do grid

Dim iLinha As Integer
Dim objComissoesInpalRegras As ClassComisInpalRegras

On Error GoTo Erro_Move_GridRegras_Memoria

    'para cada linha do grid
    For iLinha = 1 To objGridRegras.iLinhasExistentes
    
        'Instancia uma nova area de memoria a ser apontada pelo obj
        Set objComissoesInpalRegras = New ClassComisInpalRegras
    
        'se a regiao estiver preenchida
        If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_RegiaoVenda_Col))) > 0 Then
            
            'extrai o codigo da regiao e coloca no obj
            objComissoesInpalRegras.iRegiaoVenda = Codigo_Extrai(GridRegras.TextMatrix(iLinha, iGrid_RegiaoVenda_Col))
    
        End If
        
        'se o cliente estiver preenchido
        If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_Cliente_Col))) > 0 Then
        
            'extrai o codigo do cliente e coloca no obj
            objComissoesInpalRegras.lCliente = LCodigo_Extrai(GridRegras.TextMatrix(iLinha, iGrid_Cliente_Col))
        
            'se a filial do cliente estiver preenchida
            If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_FilialCliente_Col))) > 0 Then
            
                'extrai o codigo da filial e coloca no obj
                objComissoesInpalRegras.iFilialCliente = Codigo_Extrai(GridRegras.TextMatrix(iLinha, iGrid_FilialCliente_Col))
        
            End If
        
        End If
        
        'se a categoria do produto estiver preenchida
        If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_CategoriaProduto_Col))) > 0 Then
        
            'guarda a mesma no obj
            objComissoesInpalRegras.sCategoriaProduto = GridRegras.TextMatrix(iLinha, iGrid_CategoriaProduto_Col)
    
            'se o item da categoria estiver preenchido
            If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_ItemCatProduto_Col))) > 0 Then
            
                'guarda o item no obj
                objComissoesInpalRegras.sItemCatProduto = SCodigo_Extrai(GridRegras.TextMatrix(iLinha, iGrid_ItemCatProduto_Col))
    
            End If
    
        End If
    
        'guarda os percentuais de emissao e baixa no obj
        objComissoesInpalRegras.dPercTabelaA = PercentParaDbl(GridRegras.TextMatrix(iLinha, iGrid_TabelaA_Col))
        objComissoesInpalRegras.dPercTabelaB = PercentParaDbl(GridRegras.TextMatrix(iLinha, iGrid_TabelaB_Col))
    
        'adiciona o obj na colecao
        objComissoesInpalPlan.colComissoesInpalRegras.Add objComissoesInpalRegras
    
    Next
    
    Move_GridRegras_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_GridRegras_Memoria:
    
    Move_GridRegras_Memoria = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Function

End Function

Private Sub Limpa_Tela_PlanComissoesInpal()

    'limpa uma parte da tela
    Call Limpa_Tela(Me)
    
    'limpa o nome reduzido do vendedor
    LabelVendedorNome.Caption = STRING_VAZIO
    
    'desmarca as checks
    Call Desmarca_Checks_Tela
    
    'limpa o percentual de baixa
    PercentualBaixa.Caption = STRING_VAZIO
    
    'seleciona o tipo de ajuda de custos fixo (default da tela)
    AjudaCustosFixa.Value = True
    AjudaCustosMinima.Value = False
    
    'chama a grid limpa (limpa o grid de regras)
    Call Grid_Limpa(objGridRegras)

End Sub

Private Sub Desmarca_Checks_Tela()
'Desmarca as checkboxes da tela

    'desmarca a check de incidencia sobre frete
    IncideSobreFrete.Value = vbUnchecked
    
    'desmarca a check de incidencia sobre ipi
    IncideSobreIPI.Value = vbUnchecked
    
    'desmarca a check de incidencia sobre outrasdesp
    IncideSobreOutrasDesp.Value = vbUnchecked
    
    'desmarca a check de incidencia sobre seguro
    IncideSobreSeguro.Value = vbUnchecked
    
    'desmarca a check de incidencia sobre tudo
    IncideSobreTudo.Value = vbUnchecked
    
    'desmarca check de replicacao
    Replicacao.Value = vbUnchecked
    
    'desmarca check tecnico
    Tecnico.Value = vbUnchecked
    
    'obs, a check de incidencia sobre venda nunca eh
    'desmarcada
    
End Sub


'*****************************************
'9 eventos do grid que devem ser tratados
'
'
'*****************************************

Public Sub GridRegras_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRegras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegras, iAlterado)
    End If

End Sub

Public Sub GridRegras_GotFocus()
    Call Grid_Recebe_Foco(objGridRegras)
End Sub

Public Sub GridRegras_EnterCell()
    
    'se o modo de replicacao de linha esta ativado e se linha corrente estiver vazia
    If Replicacao.Value = vbChecked And Verifica_Linha_Em_Branco = True Then
    
        'se linha do grid nao for a primeira...
        If GridRegras.Row - GridRegras.FixedRows > 0 Then
        
            'repete os dados da linha anterior na corrente linha
            GridRegras.TextMatrix(GridRegras.Row, iGrid_CategoriaProduto_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_CategoriaProduto_Col)
            GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_Cliente_Col)
            GridRegras.TextMatrix(GridRegras.Row, iGrid_FilialCliente_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_FilialCliente_Col)
            GridRegras.TextMatrix(GridRegras.Row, iGrid_ItemCatProduto_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_ItemCatProduto_Col)
            GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_RegiaoVenda_Col)
            GridRegras.TextMatrix(GridRegras.Row, iGrid_TabelaA_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_TabelaA_Col)
            GridRegras.TextMatrix(GridRegras.Row, iGrid_TabelaB_Col) = GridRegras.TextMatrix(GridRegras.Row - 1, iGrid_TabelaB_Col)
        
            'adiciona a linha se for o caso...
            Call Adiciona_Linha_Seguinte
        
        End If

    End If
    
    Call Grid_Entrada_Celula(objGridRegras, iAlterado)
        
End Sub

Public Sub GridRegras_LeaveCell()
    Call Saida_Celula(objGridRegras)
End Sub

Public Sub GridRegras_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRegras)
    
End Sub

Public Sub GridRegras_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRegras, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRegras, iAlterado)
    End If

End Sub

Public Sub GridRegras_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridRegras)

End Sub

Public Sub GridRegras_RowColChange()
    
    Call Grid_RowColChange(objGridRegras)
    
End Sub

Public Sub GridRegras_Scroll()
    
    Call Grid_Scroll(objGridRegras)

End Sub

'************************************************
'fim dos 9 eventos do grid que devem ser tratados
'
'
'************************************************

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'inicializa a saida
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
        
        'Verifica qual celula do grid esta deixando
        'de ser a corrente para chamar a funcao de
        'saida celula adequada...
        Select Case objGridInt.objGrid.Col

            'se for a celula de Regiao de Venda
            Case iGrid_RegiaoVenda_Col
        
                lErro = Saida_Celula_RegiaoVenda(objGridInt)
                If lErro <> SUCESSO Then gError 98817
            
            'se for a celula de Cliente
            Case iGrid_Cliente_Col
                
                lErro = Saida_Celula_Cliente(objGridInt)
                If lErro <> SUCESSO Then gError 98814
                
            'se for a celula de FilialCliente
            Case iGrid_FilialCliente_Col
                
                lErro = Saida_Celula_FilialCliente(objGridInt)
                If lErro <> SUCESSO Then gError 98815
                
            'se for a celula de Categoria de Produto
            Case iGrid_CategoriaProduto_Col
                
                lErro = Saida_Celula_CategoriaProduto(objGridInt)
                If lErro <> SUCESSO Then gError 98813
            
            'se for a celula de ItemCatProduto
            Case iGrid_ItemCatProduto_Col
                
                lErro = Saida_Celula_ItemCatProduto(objGridInt)
                If lErro <> SUCESSO Then gError 98816
        
            'se for a celula da Tabela A
            Case iGrid_TabelaA_Col
                
                lErro = Saida_Celula_TabelaA(objGridInt)
                If lErro <> SUCESSO Then gError 98818
                
            'se for a celula da Tabela B
            Case iGrid_TabelaB_Col
            
                lErro = Saida_Celula_TabelaB(objGridInt)
                If lErro <> SUCESSO Then gError 98819
       
       End Select

    End If

    'finaliza a saida
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98820
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 98813 To 98819
        
        Case 98820
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cliente(objGridInt As AdmGrid) As Long
'Faz a crítica do campo Cliente que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Cliente

    'instancia objcontrole como o controle Cliente
    Set objGridInt.objControle = Cliente

    'exibe o cliente no grid
    lErro = Traz_Cliente_GridRegras()
    If lErro <> SUCESSO Then gError 98824
    
    If Len(Trim(Cliente.Text)) = 0 Then
        GridRegras.TextMatrix(GridRegras.Row, iGrid_FilialCliente_Col) = STRING_VAZIO
    End If
    
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98821

    'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
    Call Adiciona_Linha_Seguinte

    Saida_Celula_Cliente = SUCESSO

    Exit Function

Erro_Saida_Celula_Cliente:

    Saida_Celula_Cliente = gErr

    Select Case gErr

        Case 98821, 98824
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Traz_Cliente_GridRegras(Optional objcliente As ClassCliente) As Long
'Traz o cliente para o grid...

Dim lErro As Long
Dim iCodFilial As Integer
Dim lCliente As Long

On Error GoTo Erro_Traz_Cliente_GridRegras
    
    'Se objCliente estiver instanciado
    If Not (objcliente Is Nothing) Then
    
        'Guarda no campo Cliente o codigo do cliente contido no obj
        Cliente.Text = objcliente.lCodigo
    
    'se o obj ainda nao foi instanciado
    Else
        
        'instancia o mesmo para utiliza-lo numa possivel futura leitura
        Set objcliente = New ClassCliente
               
    End If
    
    'se o controle Cliente estiver preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
        
        'extrai o codigo
        lCliente = LCodigo_Extrai(Cliente.Text)
        
        'se conseguiu, entao coloca no campo o código para fazer a leitura
        If lCliente > 0 Then
            Cliente.Text = lCliente
        End If
        
        'Le os dados do cliente
        lErro = TP_Cliente_Le3(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 98822

        'Exibe o Codigo do Cliente e seu nome reduzido
        Cliente.Text = objcliente.lCodigo & SEPARADOR & objcliente.sNomeReduzido
        GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col) = Cliente.Text

    End If

    Traz_Cliente_GridRegras = SUCESSO
    
    Exit Function
    
Erro_Traz_Cliente_GridRegras:

    Traz_Cliente_GridRegras = gErr
    
    Select Case gErr
    
        Case 98822
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_RegiaoVenda(objGridInt As AdmGrid) As Long
'Faz a crítica do campo RegiaoVenda que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RegiaoVenda

    'instancia objcontrole como o controle da regiao
    Set objGridInt.objControle = RegiaoVenda

    'exibe a regiao de venda no grid
    lErro = Traz_RegiaoVenda_GridRegras()
    If lErro <> SUCESSO Then gError 98826
    
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98825

    'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
    Call Adiciona_Linha_Seguinte

    Saida_Celula_RegiaoVenda = SUCESSO

    Exit Function

Erro_Saida_Celula_RegiaoVenda:

    Saida_Celula_RegiaoVenda = gErr

    Select Case gErr

        Case 98825, 98826
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Traz_RegiaoVenda_GridRegras(Optional objRegiaoVenda As ClassRegiaoVenda) As Long
'Traz a regiao de venda para o grid...

Dim lErro As Long
Dim iRegiao As Integer

On Error GoTo Erro_Traz_RegiaoVenda_GridRegras
    
    'Se objRegiaoVenda estiver instanciado
    If Not (objRegiaoVenda Is Nothing) Then
    
        'Guarda no campo RegiaoVenda o codigo da regiao contido no obj
        RegiaoVenda.Text = objRegiaoVenda.iCodigo
    
    'se o obj ainda nao foi instanciado
    Else
        
        'instancia o mesmo para utiliza-lo numa possivel futura leitura
        Set objRegiaoVenda = New ClassRegiaoVenda
               
    End If
    
    'se o o controle RegiaoVenda estiver preenchido
    If Len(Trim(RegiaoVenda.Text)) > 0 Then
        
        'extrai o codigo da regiao
        iRegiao = Codigo_Extrai(RegiaoVenda.Text)
        
        'se conseguiu extrair, coloca o codigo no controle
        If iRegiao > 0 Then
            RegiaoVenda.Text = iRegiao
        End If
        
        'Le os dados da regiao de venda
        lErro = CF("TP_RegiaoVenda_Le", RegiaoVenda, objRegiaoVenda)
        If lErro <> SUCESSO Then gError 98827

        'Exibe o Codigo da regiao concatenado com sua descricao
        RegiaoVenda.Text = objRegiaoVenda.iCodigo & SEPARADOR & objRegiaoVenda.sDescricao
        GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col) = RegiaoVenda.Text

    End If

    Traz_RegiaoVenda_GridRegras = SUCESSO
    
    Exit Function
    
Erro_Traz_RegiaoVenda_GridRegras:

    Traz_RegiaoVenda_GridRegras = gErr
    
    Select Case gErr
    
        Case 98827
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_TabelaA(objGridInt As AdmGrid) As Long
'Faz a crítica do campo TabelaA que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TabelaA

    'instancia objcontrole como o controle da tabelaA
    Set objGridInt.objControle = TabelaA

    'Testar se o campo está preenchido
    If Len(Trim(TabelaA.Text)) > 0 Then
    
        'critica o valor digitado
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 98840
            
        'aplica o formato de porcentagem
        objGridInt.objControle.Text = PercentParaDbl(objGridInt.objControle.FormattedText)
            
        'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
        Call Adiciona_Linha_Seguinte
            
    Else
    
        'limpa o conteudo de tabelab
        GridRegras.TextMatrix(GridRegras.Row, iGrid_TabelaB_Col) = STRING_VAZIO
        
    End If
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98841

    Saida_Celula_TabelaA = SUCESSO

    Exit Function

Erro_Saida_Celula_TabelaA:

    Saida_Celula_TabelaA = gErr

    Select Case gErr

        Case 98840, 98841
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Sub Adiciona_Linha_Seguinte()
'adiciona a linha se a corrente for a
'ultima e se a coluna corrente estiver preenchida

    'se for ultima linha do grid habilitada e o campo estiver preenchido
    If GridRegras.Row - GridRegras.FixedRows = objGridRegras.iLinhasExistentes And Len(Trim(GridRegras.TextMatrix(GridRegras.Row, GridRegras.Col))) > 0 Then
        
        'inclui a proxima linha
        objGridRegras.iLinhasExistentes = objGridRegras.iLinhasExistentes + 1

    End If

End Sub

Private Function Saida_Celula_TabelaB(objGridInt As AdmGrid) As Long
'Faz a crítica do campo TabelaB que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_TabelaB

    'instancia objcontrole como o controle da TabelaB
    Set objGridInt.objControle = TabelaB

    'Testar se o campo está preenchido
    If Len(Trim(TabelaB.Text)) > 0 Then
    
        'critica o valor digitado
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 98842
        
        'se tabelaB = tabelaA -> erro
        If PercentParaDbl(GridRegras.TextMatrix(GridRegras.Row, iGrid_TabelaA_Col)) = StrParaDbl(objGridInt.objControle.Text) / 100 Then gError 98843
        
        'aplica o formato de porcentagem
        objGridInt.objControle.Text = PercentParaDbl(objGridInt.objControle.FormattedText)
           
        'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
        Call Adiciona_Linha_Seguinte
    
    End If
       
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98844

    Saida_Celula_TabelaB = SUCESSO

    Exit Function

Erro_Saida_Celula_TabelaB:

    Saida_Celula_TabelaB = gErr

    Select Case gErr

        Case 98842, 98844
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 98843
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAA_IGUAL_TABELAB", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialCliente(objGridInt As AdmGrid) As Long
'Faz a crítica do campo FilialCliente que está deixando de ser o campo corrente

Dim lErro As Long
Dim iIndex As Integer
Dim iCod As Integer

On Error GoTo Erro_Saida_Celula_FilialCliente

    'instancia objcontrole como o controle da FilialCliente
    Set objGridInt.objControle = FilialCliente

    'Se filial cliente está preenchida
    If Len(Trim(FilialCliente.Text)) > 0 Then
    
        'Tenta selecionar na combo
        lErro = Combo_Seleciona_Grid(FilialCliente, iCod)
        If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 98946
    
        'se nao achou, erro..
        If lErro <> SUCESSO Then gError 98960
    
    End If
    
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98845

    'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
    Call Adiciona_Linha_Seguinte
    
    Saida_Celula_FilialCliente = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialCliente:

    Saida_Celula_FilialCliente = gErr

    Select Case gErr

        Case 98845, 98946
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 98960
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_EXISTENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CategoriaProduto(objGridInt As AdmGrid) As Long
'Faz a crítica do campo CategoriaProduto que está deixando de ser o campo corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CategoriaProduto

    'instancia objcontrole como o controle da CategoriaProduto
    Set objGridInt.objControle = CategoriaProduto

    'se a categoria não estiver preenchida, tem q limpar o campo itemcatproduto
    If Len(Trim(CategoriaProduto.Text)) = 0 Then
        GridRegras.TextMatrix(GridRegras.Row, iGrid_ItemCatProduto_Col) = STRING_VAZIO
    End If
        
    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98846
    
    'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
    Call Adiciona_Linha_Seguinte
    
    Saida_Celula_CategoriaProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_CategoriaProduto:

    Saida_Celula_CategoriaProduto = gErr

    Select Case gErr

        Case 98846
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemCatProduto(objGridInt As AdmGrid) As Long
'Faz a crítica do campo ItemCatProduto que está deixando de ser o campo corrente

Dim lErro As Long
Dim sTexto As String

On Error GoTo Erro_Saida_Celula_ItemCatProduto

    'instancia objcontrole como o controle da ItemCatProduto
    Set objGridInt.objControle = ItemCatProduto

    'Se itemcatproduto está preenchido e nao tem item selecionado na combo
    If Len(Trim(ItemCatProduto.Text)) > 0 Then
    
        'atribui a variavel stexto o codigo da combo...
        'a atribuicao eh feita pq a funcao "SComboSeleciona" eh
        'mutuamente exclusiva em relacao ao codigo ou a descricao/nomereduzido
        'que possam estar concatenados ao texto...
        sTexto = SCodigo_Extrai(ItemCatProduto.Text)
        If Len(Trim(sTexto)) > 0 Then ItemCatProduto.Text = sTexto
    
        'Tenta selecionar na combo
        lErro = CF("SCombo_Seleciona", ItemCatProduto)
        If lErro <> SUCESSO And lErro <> 60483 Then gError 98962
    
        If lErro <> SUCESSO Then gError 98961
    
    End If

    'abandona a celula... atribuindo o conteudo do controle
    'ao textmatrix correspondente
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98847
    
    'adiciona a linha se a corrente for a ultima e seu campo estiver preenchido
    Call Adiciona_Linha_Seguinte
    
    Saida_Celula_ItemCatProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemCatProduto:

    Saida_Celula_ItemCatProduto = gErr

    Select Case gErr

        Case 98847, 98962
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 98961
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMCATPRODUTO_NAO_EXISTENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a funcao que ira efetuar a gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 98850

    'limpa a tela apos a gravacao
    Call Limpa_Tela_PlanComissoesInpal

    iAlterado = 0
    iVendedorAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 98850

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0
    iVendedorAlterado = 0

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Grava o registro.. deve ser sempre publica pois sera chamada
'de fora ...

Dim lErro As Long
Dim iIndice As Integer
Dim objComissoesInpalPlan As New ClassComisInpalPlan
Dim objPlanilhaIrma As New ClassComisInpalPlan

On Error GoTo Erro_Gravar_Registro

     'Exibe uma ampulheta como ponteiro do mouse
     GL_objMDIForm.MousePointer = vbHourglass

     'Guarda os dados presentes na tela no obj..
     lErro = Move_Tela_Memoria(objComissoesInpalPlan)
     If lErro <> SUCESSO Then gError 98852

     'Critica os Dados que serao gravados
     lErro = PlanComissoesInpal_Critica(objPlanilhaIrma, objComissoesInpalPlan)
     If lErro <> SUCESSO Then gError 98851

     'Critica a ajuda de custo
     'essa critica eh feita separada pois acessa o BD
'     lErro = CF("ComissoesInpalPlan_AjudaCusto_Critica", objComissoesInpalPlan, lPlanilhaIrma)
'     If lErro <> SUCESSO Then gError 101693

     'verificar se esta sendo alterado um registro ja gravado
     lErro = Trata_Alteracao(objComissoesInpalPlan, objComissoesInpalPlan.lCodigo)
     If lErro <> SUCESSO Then gError 98853

     'grava a planilha de comissoes no BD
     lErro = CF("ComissoesInpal_Grava", objComissoesInpalPlan, objPlanilhaIrma.lCodigo)
     If lErro <> SUCESSO Then gError 98854

     'fechando comando de setas
     Call ComandoSeta_Fechar(Me.Name)

     'Exibe o ponteiro padrão do mouse
     GL_objMDIForm.MousePointer = vbDefault

     Gravar_Registro = SUCESSO
     
     Exit Function

Erro_Gravar_Registro:

    'Exibe o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 98851, 98852, 98853, 98854, 101693
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function PlanComissoesInpal_Critica(objPlanilhaIrma As ClassComisInpalPlan, objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Funcao que faz a critica da tela PlanComissoesInpal
'objPlanilhaIrma eh parametro de INPUT que sera utilizado para ler os dados da planilha irma
'objPlanilhaIrma eh parametro de OUTPUT, pois retorna os dados da planilha irma
'objComissoesInpalPlan eh parametro de INPUT que traz os dados para a critica da planilha
'objComissoesInpalPlan ja vem preenchido com os dados da tela
'Os dados dos 2 objs serao confrontados, apos a leitura de objPlanilhaIrma


Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_PlanComissoesInpal_Critica

    'se codigo nao estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 98953
    
    'se o vendedor nao estiver preenchido => erro
    If Len(Trim(Vendedor.Text)) = 0 Then gError 98855

    'se o percentual de comissao na emissao nao estiver preenchido => erro
    If Len(Trim(PercentualEmissao.ClipText)) = 0 Then gError 98856

    'se nao tiver linha preenchida no grid => erro
    If objGridRegras.iLinhasExistentes = 0 Then gError 98866

    'para cada linha do grid
    For iLinha = 1 To objGridRegras.iLinhasExistentes
   
        'se a tabela A estiver preenchida
        If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_TabelaA_Col))) > 0 Then
    
            'se o valor de tabela a for 0
            If GridRegras.TextMatrix(iLinha, iGrid_TabelaA_Col) = PERCENTUAL_ZERO Then gError 98969
                        
            'se regiao nao estiver preenchida
            If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_RegiaoVenda_Col))) = 0 Then
            
                'se o cliente nao estiver preenchido
                If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_Cliente_Col))) = 0 Then
        
                    'se a categoria do produto nao estiver preenchida
                    If Len(Trim(GridRegras.TextMatrix(iLinha, iGrid_CategoriaProduto_Col))) = 0 Then gError 98860
            
                End If
        
            End If
        
        'se a tabela A nao estiver preenchida..
        Else
            gError 98967
    
        End If
            
        'chama a funcao que verifica se o conteudo de uma linha esta repetido
        lErro = Verifica_Repeteco_GridRegras(iLinha)
        If lErro <> SUCESSO Then gError 98861
    
    Next

    'Coloca os filtros em objPlanilhaIrma
    'o mesmo vendedor da planilha principal + o contrario da flag tecnico da planilha principal
    objPlanilhaIrma.iVendedor = objComissoesInpalPlan.iVendedor
    objPlanilhaIrma.iTecnico = IIf(objComissoesInpalPlan.iTecnico = VENDEDOR_DIRETO, VENDEDOR_INDIRETO, VENDEDOR_DIRETO)
    
    'Le os dados da planilha irma para poder fazer a critica da mesma
    lErro = CF("ComissoesInpalPlan_Le_Vendedor_Tecnico", objPlanilhaIrma)
    If lErro <> SUCESSO And lErro <> 101698 Then gError 101699
    
    'se achou
    If lErro = SUCESSO Then
    
        'se os tipos nao forem iguais ou se as ajudas de custo nao forem iguais
        If Not (objPlanilhaIrma.iTipoAjudaCusto = objComissoesInpalPlan.iTipoAjudaCusto) Or Not (Abs(objPlanilhaIrma.dAjudaCusto - objComissoesInpalPlan.dAjudaCusto) < DELTA_VALORMONETARIO) Then
        
            'manda msg de aviso
            If Rotina_Aviso(vbYesNo, "AVISO_AJUDACUSTOOUTIPOAJUDACUSTO_ALTERA_IRMA", objPlanilhaIrma.lCodigo, objPlanilhaIrma.iVendedor) = vbNo Then gError 101700
            
        Else
        
            'zera o codigo da planilha irma para otimizar (nao fara um update + tarde -> funcao: ComissoesInpalPlan_Grava_EmTrans)
            objPlanilhaIrma.lCodigo = 0
            
        End If
        
    End If
        
    PlanComissoesInpal_Critica = SUCESSO
    
    Exit Function
    
Erro_PlanComissoesInpal_Critica:

    PlanComissoesInpal_Critica = gErr

    Select Case gErr
    
        Case 98855
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_PREENCHIDO", gErr)
                
        Case 98856
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTAGEM_EMISSAO_NAO_PREENCHIDA", gErr)
            
        Case 98860
            Call Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_EM_BRANCO_GRIDREGRAS", gErr, iLinha)
               
        Case 98861, 101699, 101700
                        
        Case 98866
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
                        
        Case 98953
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                        
        Case 98967
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAA_NAO_PREENCHIDA", gErr, iLinha)
                        
        Case 98969
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAA_INVALIDO", gErr, iLinha)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    Exit Function

End Function

Private Function Verifica_Repeteco_GridRegras(iLinhaCorrente As Integer) As Long
'Verifica se existem linhas repetidas no Grid de Regras da linha correnta pra tras
'Sugestao: Chamar dentro de um loop que rode as linhas de forma crescente
'iLinhaCorrente RECEBE (INPUT) o dado da linha a ser analisada

Dim iLinha As Integer
Dim lErro As Long

On Error GoTo Erro_Verifica_Repeteco_GridRegras
   
    'faz com que ilinha aponte pra primeira linha apos a corrente
    'e roda o grid a partir de ilinha comparando com ilinhaCorrente
    For iLinha = iLinhaCorrente + 1 To objGridRegras.iLinhasExistentes
        
        'se as categorias do produto forem iguais
        If GridRegras.TextMatrix(iLinha, iGrid_CategoriaProduto_Col) = GridRegras.TextMatrix(iLinhaCorrente, iGrid_CategoriaProduto_Col) Then
            'se os clientes forem iguais
            If GridRegras.TextMatrix(iLinha, iGrid_Cliente_Col) = GridRegras.TextMatrix(iLinhaCorrente, iGrid_Cliente_Col) Then
                'se as filiais cliente forem iguais
                If GridRegras.TextMatrix(iLinha, iGrid_FilialCliente_Col) = GridRegras.TextMatrix(iLinhaCorrente, iGrid_FilialCliente_Col) Then
                    'se os itemcatprodutos forem iguais
                    If GridRegras.TextMatrix(iLinha, iGrid_ItemCatProduto_Col) = GridRegras.TextMatrix(iLinhaCorrente, iGrid_ItemCatProduto_Col) Then
                        'se as filiais regiaovenda forem iguais... as linhas sao iguais!!!!!
                        If GridRegras.TextMatrix(iLinha, iGrid_RegiaoVenda_Col) = GridRegras.TextMatrix(iLinhaCorrente, iGrid_RegiaoVenda_Col) Then gError 98862
                    End If
                End If
            End If
        End If
        
    Next

    Verifica_Repeteco_GridRegras = SUCESSO

    Exit Function

Erro_Verifica_Repeteco_GridRegras:
    
    Verifica_Repeteco_GridRegras = gErr
    
    Select Case gErr
    
        Case 98862
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LINHA_REPETIDA", gErr, GridRegras.Name, iLinha, iLinhaCorrente)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objComissoesInpalPlan As New ClassComisInpalPlan

On Error GoTo Erro_BotaoExcluir_Click

     'Exibe uma ampulheta como ponteiro do mouse
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Codigo da Planilha foi informado
    If Len(Trim(Codigo.Text)) = 0 Then gError 98909
    
    'Copia a chave para o obj
    objComissoesInpalPlan.lCodigo = StrParaInt(Codigo.ClipText)
    
    'Le os dados do vendedor
    lErro = CF("ComissoesInpalPlan_Le", objComissoesInpalPlan)
    If lErro <> SUCESSO And lErro <> 98762 Then gError 98941
    
    'se nao achou o vendedor => erro
    If lErro <> SUCESSO Then gError 98942
    
    'Pede confirmação para exclusão ao usuário
    If Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_INPALPLAN", objComissoesInpalPlan.lCodigo) = vbYes Then

        'exclui a planilha
        lErro = CF("ComissoesInpalPlan_Exclui", objComissoesInpalPlan)
        If lErro <> SUCESSO Then gError 98912

        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)

        'Limpa a Tela
        Call Limpa_Tela_PlanComissoesInpal

        iAlterado = 0
        iVendedorAlterado = 0

    End If

    'Exibe o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    'Exibe o ponteiro padrão do mouse
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 98909
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 98912
        
        Case 98941
        
        Case 98942
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANILHA_NAO_CADASTRADA", gErr, objComissoesInpalPlan.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    'se o vendedor estiver preenchido guarda o codigo do vendedor no objvendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.iCodigo = StrParaInt(Vendedor.Text)
        
    'chama o browser de vendedor
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

On Error GoTo Erro_objEventoVendedor_evSelecao

    'faz com que o ponteiro objVendedor
    'aponte para obj1
    Set objVendedor = obj1

    'Traz o vendedor selecionado pro grid
    lErro = Traz_Vendedor_Tela(objVendedor)
    If lErro <> SUCESSO Then gError 98935
   
    'exibe a tela... (para ficar na frente do browser)...
    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 98935
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoConsultaCampo_Click()

On Error GoTo Erro_BotaoConsultaCampo_Click
 
    'se nao existem linhas selecionadas => erro
    If GridRegras.Row = 0 Then gError 98943
    
    'faz uma selecao pela coluna do grid
    Select Case GridRegras.Col
    
        'se for a coluna de cliente
        Case iGrid_Cliente_Col
            
            'chama o browser de cliente
            Call Chama_Browser_Cliente
        
        Case iGrid_RegiaoVenda_Col
        
            'chama o browser de regiao
            Call Chama_Browser_RegiaoVenda
            
    End Select
    
    Exit Sub
    
Erro_BotaoConsultaCampo_Click:

    Select Case gErr
    
        Case 98943
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Chama_Browser_Cliente()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    'se o cliente estiver preenchido
    If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col))) > 0 Then
        
        'se o conteudo do campo for de caracteres numericos
        If IsNumeric(GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col)) Then
            
            'guarda o codigo do cliente no objcliente
            objcliente.lCodigo = StrParaLong(GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col))
        
        Else
        
            'guarda o nomereduzido em objcliente
            objcliente.sNomeReduzido = GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col)
        
        End If

    End If
    
    'chama o browser de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Public Sub Chama_Browser_RegiaoVenda()

Dim objRegiao As New ClassRegiaoVenda
Dim colSelecao As Collection

    'se a regiao estiver preenchido
    If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col))) > 0 Then
        
        'se o conteudo do campo for composto por caracteres numericos
        If IsNumeric(GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col)) Then
        
            'guarda o codigo da regiao no objRegiao
            objRegiao.iCodigo = StrParaInt(GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col))

        Else
        
            'guarda a descricao em objcliente
            objRegiao.sDescricao = GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col)

        End If
        
    End If
    
    'chama o browser de Regiao
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiao, objEventoRegiao)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    'faz com que o ponteiro objCliente
    'aponte para obj1
    Set objcliente = obj1

    'Traz o cliente selecionado pro grid
    lErro = Traz_Cliente_GridRegras(objcliente)
    If lErro <> SUCESSO Then gError 98899
   
    'exibe a tela... (para ficar na frente do browser)...
    Me.Show
    
    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr

        Case 98899
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub objEventoRegiao_evSelecao(obj1 As Object)

Dim objRegiaoVenda As ClassRegiaoVenda
Dim lErro As Long

On Error GoTo Erro_objEventoRegiao_evSelecao

    'faz com que o ponteiro objRegiaoVenda
    'aponte para obj1
    Set objRegiaoVenda = obj1

    'Traz a regiao de venda selecionada pro grid
    lErro = Traz_RegiaoVenda_GridRegras(objRegiaoVenda)
    If lErro <> SUCESSO Then gError 98898
   
    'exibe a tela... (para ficar na frente do browser)...
    Me.Show
    
    Exit Sub

Erro_objEventoRegiao_evSelecao:

    Select Case gErr

        Case 98898
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Caso o usuario queira acessar o browser através da tecla F3.
    
    'Se a tecla pressionada foi a tecla de chamda
    'de browser
    If KeyCode = KEYCODE_BROWSER Then
        
        'se controle ativo eh vendedor
        If Me.ActiveControl Is Vendedor Then
            
            'chama o browser de vendedor usando
            'o evento click do label "vendedor"
            Call LabelVendedor_Click
            
        'se controle ativo eh cliente
        ElseIf Me.ActiveControl Is Cliente Then
            
            'chama o browser de cliente
            Call Chama_Browser_Cliente
            
        'se controle ativo eh regiaovenda
        ElseIf Me.ActiveControl Is RegiaoVenda Then
            
            'chama o browser de regiaovenda
            Call Chama_Browser_RegiaoVenda
        
        End If
        
    ElseIf KeyCode = KEYCODE_REPETE_LINHA_GRID Then

        'marca/desmarca check de duplicacao
        Replicacao.Value = 1 - Replicacao.Value
    
    End If
        
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

Dim iIndice As Integer
Dim bAchou As Boolean
Dim iTecla As Integer
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyUp

    'Se a tecla "liberada" foi a tecla que dispara a repetição de campo
    If KeyCode = KEYCODE_REPETE_CAMPO_GRID Then
              
        lErro = Grid_Replica_Campo(objGridRegras, iTecla)
        If lErro <> SUCESSO And lErro <> GRID_CONTEUDO_INVALIDO_PARA_REPLICAR Then gError 94953
        
        'Se não é possível replicar o campo => erro
        If lErro = GRID_CONTEUDO_INVALIDO_PARA_REPLICAR Then gError 94954
    
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
    
    End If
    
    Exit Sub
    
Erro_UserControl_KeyUp:

    Select Case gErr

        Case 94953
        
        Case 94954
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTEUDO_CAMPO_INVALIDO", gErr, objGridRegras.colCampo(GridRegras.Col), GridRegras.Row - 1, GridRegras.Row)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub


Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'chama a teste_salva
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 98896

    'Limpa a tela
    Call Limpa_Tela_PlanComissoesInpal

    'Fecha Comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 98896

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se o Frame atual não corresponde ao TAB clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
    
    End If

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'libera objGridRegras
    Set objGridRegras = Nothing
    
    'libera objEventoVendedor
    Set objEventoVendedor = Nothing
    
    'libera ojbEventoCliente
    Set objEventoCliente = Nothing
    
    'libera objEventoRegiao
    Set objEventoRegiao = Nothing
    
    'Libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Verifica_Linha_Em_Branco() As Boolean
'verifica se a linha corrente esta em branco

    'atribui false a funcao
    Verifica_Linha_Em_Branco = False
    
    'se todos os campos estiverem vazios
    If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_CategoriaProduto_Col))) = 0 Then
        If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col))) = 0 Then
            If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_FilialCliente_Col))) = 0 Then
                If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_ItemCatProduto_Col))) = 0 Then
                    If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_RegiaoVenda_Col))) = 0 Then
                        If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_TabelaA_Col))) = 0 Then
                            If Len(Trim(GridRegras.TextMatrix(GridRegras.Row, iGrid_TabelaB_Col))) = 0 Then
                                
                                'a funcao retorna true
                                Verifica_Linha_Em_Branco = True
                            
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
                           
End Function

Private Sub BotaoProxNum_Click()
'gera um novo numero da planilha...

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'obtem o numero automatico
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_COMISSOESINPALPLAN", "ComissoesInpalPlan", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 98949

    'coloca o numero obtido anteriormente na tela...
    Codigo.Text = lCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 98949

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Trata_Combo_FilialCliente(iLocalChamada As Integer) As Long
                
Dim lErro As Long
Dim iFilial As Integer
'Dim sFilialAux As String
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim objcliente As New ClassCliente
Dim iIndex As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Trata_Combo_FilialCliente

    'Se a rotina foi chamada durante a execução do abondono de célula => sai da função, pois não faz sentido carregar a combo
    If iLocalChamada = ROTINA_GRID_ABANDONA_CELULA Then Exit Function
    
    'Guarda o Codigo da Filial selecionada
    iFilial = Codigo_Extrai(GridRegras.TextMatrix(GridRegras.Row, iGrid_FilialCliente_Col))
    
    If iLocalChamada = FUNCAO_CLIENTE_KEYPRESS Then
    
        'Guarda o código do cliente obtido no controle
        objcliente.lCodigo = Codigo_Extrai(Cliente.Text)
    
    Else
    
        'Guarda o código do cliente obtido no grid
        objcliente.lCodigo = StrParaLong(Codigo_Extrai(GridRegras.TextMatrix(GridRegras.Row, iGrid_Cliente_Col)))
    
    End If
    
    'se o codigo do cliente foi preenchido
    If objcliente.lCodigo > 0 Then

        'Le as filiais do cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6718 Then gError 98782
    
        'se nao achou filial relacionada ao cliente em questao...
        'If lErro <> SUCESSO Then gError 98804 --> nunca vai ocorrer
    
        'Varre a coleção afim de encontrar a filial selecionada anteriormente
        For iIndex = 1 To colCodigoNome.Count
                        
            'Instancia o obj que contém o código e o nome da filial
            Set objCodigoNome = colCodigoNome(iIndex)
            
            'sFilialAux =  objCodigoNome.iCodigo& SEPARADOR & objCodigoNome.sNome
            
            'se a filial for a selecionada anteriormente
            If objCodigoNome.iCodigo = iFilial Then
                
                bAchou = True
                Exit For
            End If
        
        Next
        
        'Se a função não foi chamada a partir do keypress do campo cliente
        If iLocalChamada <> FUNCAO_CLIENTE_KEYPRESS Then
        
            'Carrega a combo FiliaisCliente
            Call CF("Filial_Preenche", FilialCliente, colCodigoNome)
            
            'Se a filial que estava selecionada anteriormente foi encontrada
            'na coleção
            If bAchou Then
                'Seleciona na combo a filial com o índice encontrado acima,
                'ou seja, a mesma filial que estava selecionada anteriormente
                FilialCliente.ListIndex = iIndex - 1
            End If
        
        End If

        'Se o índice é maior do que o número de filiais na coleção
        If iIndex > colCodigoNome.Count Then
        
            'Significa que não encontrou a filial que estava selecionada anteriormente, portanto
            'deve limpar o campo
            GridRegras.TextMatrix(GridRegras.Row, iGrid_FilialCliente_Col) = STRING_VAZIO
            FilialCliente.Text = STRING_VAZIO
        
        End If
    
    End If

    Trata_Combo_FilialCliente = SUCESSO
    
    Exit Function
    
Erro_Trata_Combo_FilialCliente:

    Trata_Combo_FilialCliente = gErr
    
    Select Case gErr

        Case 98782

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Trata_Combo_ItemCatProduto(iLocalChamada As Integer) As Long
                
Dim lErro As Long
Dim iIndex As Integer
Dim bAchou As Boolean
Dim bSelecionado As Boolean
Dim sItemCatProduto As String
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItensCatProduto As New Collection

On Error GoTo Erro_Trata_Combo_ItemCatProduto

    'Se está executando o abandono de célula=> sai da função pois a combo não precisa ser carregada
    If iLocalChamada = ROTINA_GRID_ABANDONA_CELULA Then Exit Function
    
    'Guarda o item de categoria do produto atual para exibi-lo novamente após carregar a combo
    sItemCatProduto = SCodigo_Extrai(GridRegras.TextMatrix(GridRegras.Row, iGrid_ItemCatProduto_Col))
    
    'Se a função foi chamada a partir da função CategoriaProduto_KeyPress
    Select Case iLocalChamada
    
        Case FUNCAO_CATEGORIAPRODUTO_KEYPRESS
                
            'Guarda a categoria de produto que está selecionada no CONTROLE
            objCategoriaProduto.sCategoria = CategoriaProduto.Text
        
        Case Else
        
        'Guarda a categoria de produto que está selecionada no CONTROLE
        objCategoriaProduto.sCategoria = GridRegras.TextMatrix(GridRegras.Row, iGrid_CategoriaProduto_Col)
        
    End Select

    'Limpa a combo para evitar que ela seja carregada com o mesmo conteúdo várias vezes
    ItemCatProduto.Clear
   
    'se a categoria foi preenchida
    If Len(Trim(objCategoriaProduto.sCategoria)) > 0 Then

        'Le os itens
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCatProduto)
        If lErro <> SUCESSO And lErro <> 22541 Then gError 98779
        
        'se nao achou
        If lErro <> SUCESSO Then gError 98802
        
        'Para cada item na coleção
        For iIndex = 1 To colItensCatProduto.Count
            
            'Instancia objCategoriaProdutoItem com os dados do item atual
            Set objCategoriaProdutoItem = colItensCatProduto.Item(iIndex)
            
            'Se o item atual for igual ao item que estava selecionado na combo => indica que o item já foi encontrado
            If objCategoriaProdutoItem.sItem = sItemCatProduto Then bAchou = True
            
            'Se a função foi chamada de um local diferente da do keypress do campo CategoriaProduto
            If iLocalChamada <> FUNCAO_CATEGORIAPRODUTO_KEYPRESS Then
                
                'é preciso carregar a combo de itens de categoria de produto
                ItemCatProduto.AddItem objCategoriaProdutoItem.sItem & SEPARADOR & objCategoriaProdutoItem.sDescricao
                
                'Se o item foi encontrado na coleção e ainda não foi selecionado na combo
                If bAchou And (Not bSelecionado) Then
                
                    'Seleciona-o na combo
                    ItemCatProduto.ListIndex = iIndex - 1
                    
                    'Indica que ele já foi selecionado
                    bSelecionado = True
                    
                End If
            
            End If
            
        Next
        
        'Se o item não foi encontrado na coleção
        If Not bAchou Then
            'Significa que não encontrou o item de categoria que estava selecionado anteriormente, portanto
            'deve limpar o campo
            GridRegras.TextMatrix(GridRegras.Row, iGrid_ItemCatProduto_Col) = STRING_VAZIO
            ItemCatProduto.Text = STRING_VAZIO

        End If
    
    End If
        
    Trata_Combo_ItemCatProduto = SUCESSO
    
    Exit Function
    
Erro_Trata_Combo_ItemCatProduto:

    Trata_Combo_ItemCatProduto = gErr
    
    Select Case gErr

        Case 98802
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_SEM_ITENS", gErr, objCategoriaProduto.sCategoria)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function
