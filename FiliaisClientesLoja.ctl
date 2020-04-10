VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FiliaisClientesLoja 
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   ScaleHeight     =   5670
   ScaleWidth      =   7980
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4125
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   7365
      Begin VB.ComboBox Estado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1140
         TabIndex        =   25
         Top             =   1965
         Width           =   630
      End
      Begin VB.TextBox Endereco 
         Height          =   315
         Left            =   1140
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   750
         Width           =   6015
      End
      Begin VB.ComboBox Pais 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3435
         TabIndex        =   23
         Top             =   1410
         Width           =   1995
      End
      Begin MSMask.MaskEdBox Cidade 
         Height          =   315
         Left            =   3435
         TabIndex        =   26
         Top             =   1965
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CEP 
         Height          =   315
         Left            =   5805
         TabIndex        =   27
         Top             =   1980
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Email 
         Height          =   315
         Left            =   3435
         TabIndex        =   28
         Top             =   3075
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Contato 
         Height          =   315
         Left            =   5805
         TabIndex        =   29
         Top             =   3075
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fax 
         Height          =   315
         Left            =   3435
         TabIndex        =   30
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Bairro 
         Height          =   315
         Left            =   1140
         TabIndex        =   31
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Telefone1 
         Height          =   315
         Left            =   1140
         TabIndex        =   32
         Top             =   2520
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Telefone2 
         Height          =   315
         Left            =   1140
         TabIndex        =   33
         Top             =   3075
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         PromptChar      =   " "
      End
      Begin VB.Label PaisLabel 
         AutoSize        =   -1  'True
         Caption         =   "País:"
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
         Left            =   2850
         TabIndex        =   44
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Left            =   5040
         TabIndex        =   43
         Top             =   3120
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
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
         Left            =   5325
         TabIndex        =   42
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
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
         Left            =   2940
         TabIndex        =   41
         Top             =   2595
         Width           =   405
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "e-mail:"
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
         Left            =   2745
         TabIndex        =   40
         Top             =   3120
         Width           =   570
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Telefone 2:"
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
         Left            =   75
         TabIndex        =   39
         Top             =   3120
         Width           =   1005
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Telefone 1:"
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
         Left            =   75
         TabIndex        =   38
         Top             =   2580
         Width           =   1005
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
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
         Left            =   465
         TabIndex        =   37
         Top             =   1455
         Width           =   585
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         Left            =   345
         TabIndex        =   36
         Top             =   2010
         Width           =   675
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
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
         Left            =   2670
         TabIndex        =   35
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
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
         Left            =   165
         TabIndex        =   34
         Top             =   810
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4215
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   7335
      Begin VB.Frame Frame1 
         Caption         =   "Código Cliente"
         Height          =   1380
         Index           =   0
         Left            =   735
         TabIndex        =   7
         Top             =   240
         Width           =   5835
         Begin MSMask.MaskEdBox CodCliente 
            Height          =   315
            Left            =   1605
            TabIndex        =   8
            Top             =   345
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label ClienteLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1605
            TabIndex        =   46
            Top             =   840
            Width           =   3450
         End
         Begin VB.Label Label6 
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
            Height          =   210
            Left            =   840
            TabIndex        =   45
            Top             =   885
            Width           =   690
         End
         Begin VB.Label CodigoBO 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4080
            TabIndex        =   11
            Top             =   330
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "BackOffice:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   2985
            TabIndex        =   10
            Top             =   375
            Width           =   1020
         End
         Begin VB.Label LabelCodCliente 
            AutoSize        =   -1  'True
            Caption         =   "Loja:"
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
            Left            =   1110
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   9
            Top             =   375
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Código Filial"
         Height          =   1380
         Left            =   720
         TabIndex        =   12
         Top             =   1725
         Width           =   5835
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2400
            Picture         =   "FiliaisClientesLoja.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Numeração Automática"
            Top             =   315
            Width           =   300
         End
         Begin MSMask.MaskEdBox CodFilial 
            Height          =   315
            Left            =   1725
            TabIndex        =   14
            Top             =   300
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Nome 
            Height          =   315
            Left            =   1725
            TabIndex        =   47
            Top             =   840
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            Caption         =   "Nome da Filial:"
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
            Left            =   360
            TabIndex        =   48
            Top             =   900
            Width           =   1275
         End
         Begin VB.Label CodFilialBO 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4155
            TabIndex        =   17
            Top             =   307
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "BackOffice:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   3090
            TabIndex        =   16
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label LabelCodFilial 
            AutoSize        =   -1  'True
            Caption         =   "Loja:"
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
            Left            =   1215
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   15
            Top             =   345
            Width           =   435
         End
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   1725
         TabIndex        =   18
         Top             =   3300
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox RG 
         Height          =   315
         Left            =   1725
         TabIndex        =   19
         Top             =   3720
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Mask            =   "###############"
         PromptChar      =   " "
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "RG:"
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
         Left            =   1335
         TabIndex        =   21
         Top             =   3765
         Width           =   345
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF:"
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
         Left            =   705
         TabIndex        =   20
         Top             =   3345
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5640
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FiliaisClientesLoja.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FiliaisClientesLoja.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FiliaisClientesLoja.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "FiliaisClientesLoja.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4650
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   8202
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
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
Attribute VB_Name = "FiliaisClientesLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'DECLARAÇÃO DE VARIÁVEIS GLOBAIS

Dim iFrameAtual1 As Integer
Dim iFrameAtual2 As Integer
Dim iAlterado As Integer
Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFilialCliente As AdmEvento
Attribute objEventoFilialCliente.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Enderecos = 2

Public Sub LabelCodCliente_Click()

Dim colSelecao As Collection
Dim objCliente As New ClassCliente

    'Preenche ClienteAte com o cliente da tela
    objCliente.lCodigoLoja = StrParaLong(CodCliente.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean
Dim objFilialCliente As New ClassFilialCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objCliente = obj1
    
    objFilialCliente.lCodClienteLoja = objCliente.lCodigoLoja
    objFilialCliente.iCodFilialLoja = FILIAL_MATRIZ
    objFilialCliente.iFilialEmpresaLoja = giFilialEmpresa
    
    'Faz a leitura da Filial Cliente
    lErro = CF("FilialCliente_Le_Loja", objFilialCliente)
    If lErro <> SUCESSO Then gError 112706
    
    'Exibe os dados do Cliente
    lErro = Exibe_Dados_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 112708

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
    
        Case 112708
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160264)

    End Select

    Exit Sub

End Sub

Public Sub LabelCodFilial_Click()

Dim colSelecao As Collection
Dim objFilialCliente As New ClassFilialCliente

    'Preenche ClienteAte com o cliente da tela
    objFilialCliente.lCodCliente = StrParaLong(CodCliente.Text)
    objFilialCliente.iCodFilial = StrParaInt(CodFilial.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("FiliaisClientesLista", colSelecao, objFilialCliente, objEventoFilialCliente)

End Sub

Private Sub objEventoFilialCliente_evSelecao(obj1 As Object)

Dim objFilialCliente As ClassFilialCliente
Dim lErro As Long

On Error GoTo Erro_objEventoFilialCliente_evSelecao

    Set objFilialCliente = obj1

    'Tenta ler Filial de Cliente com a chave passada em objFilialCliente
    lErro = CF("FilialCliente_Le_Loja", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 112607 Then gError 112646

    'Se Filial não existe
    If lErro <> SUCESSO Then gError 112647

    'Exibe dados da Filial na Tela
    lErro = Exibe_Dados_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 112648

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoFilialCliente_evSelecao:

    Select Case gErr

        Case 112646, 112648

        Case 112647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160265)

    End Select

    Exit Sub

End Sub


Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático de Filial
    lErro = CF("FilialClienteLoja_Automatico", CLng(CodCliente.Text), iCodigo)
    If lErro <> SUCESSO Then gError 112649

    'Coloca código gerado na Tela
    CodFilial.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 112649
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160266)
    
    End Select

    Exit Sub

End Sub

Public Sub Bairro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(CodCliente.Text)) = 0 And Len(CodigoBO.Caption) = 0 Then gError 112650

    'Verifica se o código foi preenchido
    If Len(Trim(CodCliente.Text)) <> 0 And Len(CodigoBO.Caption) <> 0 Then gError 117607

    'Verifica se o código da Filial foi preenchido
    If Len(Trim(CodFilial.Text)) = 0 And Len(CodFilialBO.Caption) = 0 Then gError 112651

    'Verifica se o código da Filial foi preenchido
    If Len(Trim(CodFilial.Text)) <> 0 And Len(CodFilialBO.Caption) <> 0 Then gError 117608

    If Len(CodFilialBO.Caption) <> 0 Then gError 117609

    'Verifica se é Matriz
    If CInt(CodFilial.Text) = FILIAL_MATRIZ Then gError 112652

    'o problema é q o codigo do cliente pode já ter sido cadastrado por outra filial
    objCliente.lCodigoLoja = StrParaLong(CodCliente.Text)
    objFilialCliente.lCodClienteLoja = StrParaLong(CodCliente.Text)
    objFilialCliente.iCodFilialLoja = StrParaInt(CodFilial.Text)
    objCliente.lCodigo = StrParaLong(CodigoBO.Caption)
    objFilialCliente.iCodFilial = StrParaInt(CodFilialBO.Caption)

    'Lê os dados do Cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 112653

    'Verifica se Cliente não está cadastrado
    If lErro <> SUCESSO Then gError 112654

    'Lê os dados da Filial Cliente
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 112655

    'Verifica se a Filial Cliente não está cadastrada
    If lErro <> SUCESSO Then gError 112656

    'Envia aviso perguntando se realmente deseja excluir Filial Cliente
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_FILIALCLIENTE")

    If vbMsgRes = vbYes Then

        'Exclui Filial de Cliente
        lErro = CF("FilialCliente_Exclui", objFilialCliente)
        If lErro <> SUCESSO Then gError 112657

'        'Exclui Filial da TreeView
'        Call Arvore_Excluir(Filiais, objFilialCliente)
'
        'Limpa a Tela
        lErro = Limpa_Tela_FiliaisClientesLoja()
        If lErro <> SUCESSO Then gError 112658
        
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 112650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", gErr)

        Case 112651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", gErr)

        Case 112653, 112655, 112657, 112658

        Case 112654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case 112656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 112652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_EXCLUSAO_MATRIZ", gErr)

        Case 117607
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CODIGO_JA_PREENCHIDO", gErr)

        Case 117608
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_CODIGO_JA_PREENCHIDO", gErr)

        Case 117609
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCLIENTE_TRANSFERIDO", gErr, CodFilialBO.Caption)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160267)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama função de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 112659

    'Limpa a Tela
    lErro = Limpa_Tela_FiliaisClientesLoja()
    If lErro <> SUCESSO Then gError 112660
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 112659, 112660

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160268)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se deseja salvar modificações feitas
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 112661

    'Limpa a Tela
    lErro = Limpa_Tela_FiliaisClientesLoja()
    If lErro <> SUCESSO Then gError 112662

    iAlterado = 0

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 112661 'cancelou operacao de gravacao , continua execucao normal
        
        Case 112662 'Tratado na Rotina Chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160269)

    End Select

End Sub

Public Sub CEP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CEP_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CEP, iAlterado)
    
End Sub

Public Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGC.Text))

    Case STRING_CPF 'CPF

        'Critica CPF
        lErro = Cpf_Critica(CGC.Text)
        If lErro <> SUCESSO Then gError 112663

        CGC.Format = "000\.000\.000-00; ; ; "
        CGC.Text = CGC.Text

    Case STRING_CGC 'CGC

        'Critica CGC
        lErro = Cgc_Critica(CGC.Text)
        If lErro <> SUCESSO Then gError 112664

        CGC.Format = "00\.000\.000\/0000-00; ; ; "
        CGC.Text = CGC.Text

    Case Else

        gError 112665

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case gErr

        Case 112663, 112664

        Case 112665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160270)

    End Select


    Exit Sub

End Sub

Public Sub RG_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RG_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(RG, iAlterado)

End Sub

Public Sub Cidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodCliente_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodCliente, iAlterado)

End Sub

Public Sub CodCliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente
Dim iIndice As Integer

On Error GoTo Erro_CodCliente_Validate

    'Verifica se foi preenchido o campo CodCliente
    If Len(Trim(CodCliente.Text)) = 0 Then Exit Sub

    'Critica se é do tipo Long positivo
    lErro = Long_Critica(CodCliente.Text)
    If lErro <> SUCESSO Then gError 112666

    objCliente.lCodigoLoja = CLng(CodCliente.Text)

    'Lê o Cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 112667

    'Coloca o Nome Reduzido na label
    ClienteLabel.Caption = objCliente.sNomeReduzido
    CodigoBO.Caption = objCliente.lCodigo
    
    'Verifica se existe
    If lErro <> SUCESSO Then

        'Envia aviso perguntando se deseja cadastrar novo Cliente
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE")

        If vbMsgRes = vbYes Then

            'Chama tela de Clientes
            lErro = Chama_Tela("Clientes", objCliente)
            If lErro <> SUCESSO Then gError 112668

        Else

            Cancel = True

        End If

    End If

    Exit Sub

Erro_CodCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 112666 To 112668

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160271)

    End Select

    Exit Sub

End Sub

Public Sub CodFilial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodFilial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodFilial, iAlterado)

End Sub

Public Sub CodFilial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodFilial_Validate

    'Verifica se foi preenchido o campo Codigo Filial
    If Len(Trim(CodFilial.Text)) = 0 Then Exit Sub

    'Verifica se é do tipo Inteiro e positivo
    lErro = Inteiro_Critica(CodFilial.Text)
    If lErro <> SUCESSO Then gError 112669

    Exit Sub

Erro_CodFilial_Validate:

    Cancel = True

    Select Case gErr

        Case 112669

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160272)

    End Select

    Exit Sub

End Sub

Public Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub Email_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Endereco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Estado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Estado
    If Estado.Text = Estado.List(Estado.ListIndex) Then Exit Sub

    'Verifica se existe o item no Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Estado)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 112670

    'Não existe o item na ComboBox Estado
    If lErro <> SUCESSO Then gError 112671

    Exit Sub

Erro_Estado_Validate:

    Cancel = True

    Select Case gErr

        Case 112670
    
        Case 112671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, Estado.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160273)
    
    End Select

    Exit Sub

End Sub

Public Sub Fax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub



Public Sub Pais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Pais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Pais.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox Pais
    If Pais.Text = Pais.List(Pais.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Pais, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 112672

    'Nao existe o item com o CODIGO na List da ComboBox
    If lErro = 6730 Then gError 112673

    'Nao existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 112674

    Exit Sub

Erro_Pais_Validate:

    Cancel = True

    Select Case gErr

        Case 112674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Pais.Text))

        Case 112672  'Tratado na rotina chamada

        Case 112673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160274)

    End Select


    Exit Sub

End Sub

Public Sub Telefone1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Telefone2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Function Limpa_Tela_FiliaisClientesLoja() As Long

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim sCodigoCliente As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_FiliaisClientesLoja

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Guarda o código de Cliente
    sCodigoCliente = CodCliente.Text

    'Limpa os MaskedEdit e TextBoxes
    Call Limpa_Tela(Me)

    'Mantém o código de Cliente na Tela
    CodCliente.Text = sCodigoCliente
    
    'Limpa o código da Filial
    CodFilial.Text = ""
    
    'Escolhe Estado da FilialEmpresa
    lErro = CF("Estado_Seleciona", Estado)
    If lErro <> SUCESSO Then gError 112678

    'Seleciona Brasil nas Combos de Pais se existir
    For iIndice2 = 0 To Pais.ListCount - 1

        If Right(Pais.List(iIndice2), 6) = "Brasil" Then
            Pais.ListIndex = iIndice2
            Exit For
        End If

    Next
        
    Limpa_Tela_FiliaisClientesLoja = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_FiliaisClientesLoja:
    
    Limpa_Tela_FiliaisClientesLoja = gErr
    
    Select Case gErr
        
        Case 112678
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160275)

    End Select
    
    Exit Function
        
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Filiais Clientes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FiliaisClientesLoja"
    
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

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoFilialCliente = New AdmEvento
    
    'Incluido por Leo em 11/03/02
    'Implementado pois agora é possível ter constantes cutomizadas em função de tamanhos de campos do BD. AdmLib.ClassConsCust
    Endereco.MaxLength = STRING_ENDERECO
    Bairro.MaxLength = STRING_BAIRRO
    Cidade.MaxLength = STRING_CIDADE
    'Leo até aqui
    
    iFrameAtual1 = 1
    
'    OpcaoEndereco(0) = True
'    OpcaoEndereco(1) = False
'    OpcaoEndereco(2) = False
    
    'Lê cada codigo da tabela Estados e coloca em colCodigo
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 112680

    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For Each vCodigo In colCodigo
        Estado.AddItem vCodigo
    Next

    'Escolhe Estado da FilialEmpresa
    lErro = CF("Estado_Seleciona", Estado)
    If lErro <> SUCESSO Then gError 112681

    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoNome)
    If lErro <> SUCESSO Then gError 112682

    'Preenche cada ComboBox País com os objetos da colecao colCodigoNome
    For Each objCodigoNome In colCodigoNome
        Pais.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Pais.ItemData(Pais.NewIndex) = objCodigoNome.iCodigo
    Next

    'Seleciona Brasil se existir
    For iIndice2 = 0 To Pais.ListCount - 1
        If Right(Pais.List(iIndice2), 6) = "Brasil" Then
            Pais.ListIndex = iIndice2
            Exit For
        End If
    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 112679 To 112682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160276)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objFilialCliente As ClassFilialCliente) As Long

Dim lErro As Long
Dim objFilialClienteEstatistica As New ClassFilialClienteEst
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objFilialCliente Is Nothing) Then

        'Se foi passado código de Cliente
        If objFilialCliente.lCodClienteLoja <> 0 Then

            'Se foi passado código de Filial
            If objFilialCliente.iCodFilialLoja <> 0 Then
                
                objFilialCliente.iFilialEmpresaLoja = giFilialEmpresa
                
                'Tenta ler FilialCliente
                lErro = CF("FilialCliente_Le_Loja", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 112607 Then gError 112683

                'Se a Filial não existir
                If lErro <> SUCESSO Then
                    
                    'Limpa a Tela
                    lErro = Limpa_Tela_FiliaisClientesLoja()
                    If lErro <> SUCESSO Then gError 112684
                    
                    CodCliente.Text = CStr(objFilialCliente.lCodCliente)
                    CodFilial.Text = CStr(objFilialCliente.iCodFilial)
                    
                    
                Else  'Filial existe, então exibe seus dados
    
                    lErro = Exibe_Dados_FilialCliente(objFilialCliente)
                    If lErro <> SUCESSO Then gError 112685

                End If

            Else 'Apenas o código do Cliente foi passado
                
                'Limpa Tela
                lErro = Limpa_Tela_FiliaisClientesLoja()
                If lErro <> SUCESSO Then gError 112686
                
                CodCliente.Text = CStr(objFilialCliente.lCodCliente)

                'Inserido por Leo em 10/01/02
                If Len(Trim(objFilialCliente.sNomeReduzidoCli)) > 0 Then
        
                    ClienteLabel.Caption = objFilialCliente.sNomeReduzidoCli

                End If
                
                If Len(Trim(objFilialCliente.sNome)) > 0 Then
                    Nome.Text = objFilialCliente.sNome
                End If
                'Leo até aqui

            End If

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 112683 To 112686

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160277)

    End Select

    iAlterado = 0

    Exit Function

End Function
                

Public Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Nome_Validate(Cancel As Boolean)

    Nome.Text = Trim(Nome.Text)
    'FilialLabel(1).Caption = Trim(Nome.Text)
    'FilialLabel(2).Caption = Trim(Nome.Text)
    'FilialLabel(3).Caption = Trim(Nome.Text)
    
End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.index <> iFrameAtual1 Then

        If TabStrip_PodeTrocarTab(iFrameAtual1, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.index).Visible = True
        Frame1(iFrameAtual1).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual1 = Opcao.SelectedItem.index
            
        Select Case iFrameAtual1
            
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_ID
                            
            Case TAB_Enderecos
                Parent.HelpContextID = IDH_FILIAIS_CLIENTES_ENDERECOS
            
        End Select
    
    End If

End Sub

Function Exibe_Dados_FilialCliente(objFilialCliente As ClassFilialCliente) As Long
'Exibe os dados da filial do cliente na tela

Dim lErro As Long
Dim objEndereco As ClassEndereco
Dim colEnderecos As New colEndereco
Dim objCliente As New ClassCliente
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_FilialCliente

    objCliente.lCodigoLoja = objFilialCliente.lCodClienteLoja
    objCliente.iFilialEmpresaLoja = giFilialEmpresa
    
    'Lê Cliente a partir do Nome Reduzido
    lErro = CF("Cliente_Le_Loja", objCliente)
    If lErro <> SUCESSO And lErro <> 112606 Then gError 112689

    'Cliente não cadastrado
    If lErro = 112606 Then gError 112690

    ClienteLabel.Caption = objCliente.sNomeReduzido
    CodigoBO.Caption = objCliente.lCodigo
    
    'IDENTIFICACAO :

    CodCliente.Text = CStr(objFilialCliente.lCodClienteLoja)
    CodFilial.Text = CStr(objFilialCliente.iCodFilialLoja)
    Nome.Text = objFilialCliente.sNome
    ClienteLabel.Caption = objCliente.sRazaoSocial
    CodFilialBO.Caption = objFilialCliente.iCodFilial
    
    'INSCRIÇÕES

    'FilialLabel(2).Caption = objFilialCliente.sNome
    Nome.Text = objFilialCliente.sNome
    
    RG.Text = objFilialCliente.sRG
    CGC.Text = objFilialCliente.sCgc
    Call CGC_Validate(bSGECancelDummy)
    
    'ENDERECOS :

    'FilialLabel(0).Caption = objFilialCliente.sNome

    'Lê os dados dos tres tipos de enderecos
    lErro = Enderecos_Le_FiliaisClientes(colEnderecos, objFilialCliente)
    If lErro <> SUCESSO Then gError 112691

    'Preenche as ComboBox relativas aos tres tipos de Endereco do cliente, com os objetos da colecao colEndereco
    iIndice = 0
    
    Set objEndereco = colEnderecos.Item(1)
    
    Endereco.Text = objEndereco.sEndereco
    Bairro.Text = objEndereco.sBairro
    Cidade.Text = objEndereco.sCidade
    CEP.Text = objEndereco.sCEP
    Estado.Text = objEndereco.sSiglaEstado

    If objEndereco.iCodigoPais = 0 Then
        Pais.Text = ""
    Else
        Pais.Text = objEndereco.iCodigoPais
        Call Pais_Validate(bSGECancelDummy)
    End If

    Telefone1.Text = objEndereco.sTelefone1
    Telefone2.Text = objEndereco.sTelefone2
    Fax.Text = objEndereco.sFax
    Email.Text = objEndereco.sEmail
    Contato.Text = objEndereco.sContato
    
Exit Function

    'VENDAS :

    'FilialLabel(1).Caption = objFilialCliente.sNome

Erro_Exibe_Dados_FilialCliente:

    Exibe_Dados_FilialCliente = gErr

    Select Case gErr

        Case 112690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigoLoja)

        Case 112689, 112691

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160278)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim colEndereco As New Collection
Dim colCategoriaItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código e o Codigo do Backoffice esta vazio
    If Len(Trim(CodCliente.Text)) = 0 And Len(CodigoBO.Caption) = 0 Then gError 1212692
    
    'Verifica se foi preenchido o Código e o Codigo do Backoffice simultaneamente
    If Len(Trim(CodCliente.Text)) <> 0 And Len(CodigoBO.Caption) <> 0 Then gError 117603
    
    'Verifica se foi preenchido Codigo da Filial
    If Len(Trim(CodFilial.Text)) = 0 And Len(CodFilialBO.Caption) = 0 Then gError 112693

    'Verifica se foi preenchido Codigo da Filial
    If Len(Trim(CodFilial.Text)) <> 0 And Len(CodFilialBO.Caption) <> 0 Then gError 117604

    'Verifica se foi preenchido o Nome da Filial
    If Len(Trim(Nome.Text)) = 0 Then gError 112694

    'Verifica se foi preenchido o Estado
    If Len(Trim(Endereco.Text)) <> 0 Then
        If Len(Trim(Estado.Text)) = 0 Then gError 112695
    End If

    'Lê os Enderecos e coloca em colEndereco
    lErro = Le_Dados_Enderecos(colEndereco)
    If lErro <> SUCESSO Then gError 112696

    'Lê os dados da tela relativos à Filial Cliente
    lErro = Le_Dados_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 112697
    
    'Se o CGC estiver Preenchido
    If Len(Trim(objFilialCliente.sCgc)) > 0 Then
        'Verifica se tem outro Cliente com o mesmo CGC e dá aviso
        lErro = CF("FilialCliente_Testa_CGC", objFilialCliente.lCodCliente, 0, objFilialCliente.sCgc)
        If lErro <> SUCESSO Then gError 112698
    End If
    
    'Grava FilialCliente no BD
    lErro = CF("FiliaisClientes_Grava", objFilialCliente, colEndereco)
    If lErro <> SUCESSO Then gError 112699

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 112692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODCLIENTE_NAO_PREENCHIDO", gErr)

        Case 112693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODFILIAL_NAO_PREENCHIDO", gErr)

        Case 112694
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMEFILIAL_NAO_PREENCHIDO", gErr)

        Case 112696 To 112699

        Case 112695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_PREENCHIDO", gErr)

        Case 117603
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CODIGO_JA_PREENCHIDO", gErr)

        Case 117604
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_CODIGO_JA_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160279)

    End Select

    Exit Function

End Function

Private Function Le_Dados_Enderecos(colEndereco As Collection) As Long
'Lê os dados relativos ao endereco e coloca em colEndereco

Dim objEndereco As ClassEndereco
Dim iEstadoPreenchido As Integer 'apagar
    
    Set objEndereco = New ClassEndereco

    'Preenche objEndereco com os dados do Endereço
    objEndereco.sEndereco = Trim(Endereco.Text)
    objEndereco.sBairro = Trim(Bairro.Text)
    objEndereco.sCidade = Trim(Cidade.Text)
    objEndereco.sCEP = Trim(CEP.Text)

    'Se o Endereco não estiver Preenchido --> Seta o Estado que esta Preenchido em Algum dos Frames
    If Len(Trim(Endereco.Text)) > 0 Then
        objEndereco.sSiglaEstado = Trim(Estado.Text)
    End If

    If Len(Trim(Pais.Text)) = 0 Then
        objEndereco.iCodigoPais = 0
    Else
        objEndereco.iCodigoPais = Codigo_Extrai(Pais.Text)
    End If

    objEndereco.sTelefone1 = Trim(Telefone1.Text)
    objEndereco.sTelefone2 = Trim(Telefone2.Text)
    objEndereco.sFax = Trim(Fax.Text)
    objEndereco.sEmail = Trim(Email.Text)
    objEndereco.sContato = Trim(Contato.Text)

    'Adiciona objEndereco na coleção
    colEndereco.Add objEndereco

    Le_Dados_Enderecos = SUCESSO

End Function

Private Function Le_Dados_FilialCliente(objFilialCliente As ClassFilialCliente) As Long
'Lê os dados que estão na tela de FiliaisClientes e coloca-os em objFilialCliente

Dim lErro As Long

On Error GoTo Erro_Le_Dados_FilialCliente

    'IDENTIFICACAO :

    objFilialCliente.lCodClienteLoja = StrParaLong(CodCliente.Text)

    objFilialCliente.iCodFilialLoja = StrParaInt(CodFilial.Text)
    objFilialCliente.iFilialEmpresaLoja = giFilialEmpresa
    
    objFilialCliente.lCodCliente = StrParaLong(CodigoBO.Caption)
    objFilialCliente.iCodFilial = StrParaInt(CodFilialBO.Caption)
    
    objFilialCliente.sRG = Trim(RG.Text)
    objFilialCliente.sNome = Trim(Nome.Text)
    objFilialCliente.sCgc = Trim(CGC.Text)

    Le_Dados_FilialCliente = SUCESSO

    Exit Function

Erro_Le_Dados_FilialCliente:

    Le_Dados_FilialCliente = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160280)

    End Select

    Exit Function

End Function

Function Enderecos_Le_FiliaisClientes(colEnderecos As colEndereco, objFilialCliente As ClassFilialCliente) As Long
'Lê cada um dos 3 tipos de enderecos(principal,entrega e cobranca) na tabela Enderecos e coloca na colecao

Dim lErro As Long
Dim objEndereco As ClassEndereco

On Error GoTo Erro_Enderecos_Le_FiliaisClientes

    'Endereco Principal

    Set objEndereco = New ClassEndereco

    objEndereco.lCodigo = objFilialCliente.lEndereco

    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 112702

    colEnderecos.Add objEndereco.sEndereco, objEndereco.sBairro, objEndereco.sCidade, objEndereco.sSiglaEstado, objEndereco.iCodigoPais, objEndereco.sCEP, objEndereco.sTelefone1, objEndereco.sTelefone2, objEndereco.sEmail, objEndereco.sFax, objEndereco.sContato, objEndereco.lCodigo

    'Endereco de Entrega

    Set objEndereco = New ClassEndereco

    objEndereco.lCodigo = objFilialCliente.lEnderecoEntrega

    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 112703

    colEnderecos.Add objEndereco.sEndereco, objEndereco.sBairro, objEndereco.sCidade, objEndereco.sSiglaEstado, objEndereco.iCodigoPais, objEndereco.sCEP, objEndereco.sTelefone1, objEndereco.sTelefone2, objEndereco.sEmail, objEndereco.sFax, objEndereco.sContato, objEndereco.lCodigo

    'Endereco de Cobranca

    Set objEndereco = New ClassEndereco

    objEndereco.lCodigo = objFilialCliente.lEnderecoCobranca

    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then gError 112704

    colEnderecos.Add objEndereco.sEndereco, objEndereco.sBairro, objEndereco.sCidade, objEndereco.sSiglaEstado, objEndereco.iCodigoPais, objEndereco.sCEP, objEndereco.sTelefone1, objEndereco.sTelefone2, objEndereco.sEmail, objEndereco.sFax, objEndereco.sContato, objEndereco.lCodigo

    Enderecos_Le_FiliaisClientes = SUCESSO

Exit Function

Erro_Enderecos_Le_FiliaisClientes:

    Enderecos_Le_FiliaisClientes = gErr

    Select Case gErr

        Case 112702, 112703, 112704

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160281)

    End Select

    Exit Function

End Function



'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA  "
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "FiliaisClientes"

    'Lê os dados da Tela FilialCliente
    lErro = Le_Dados_FilialCliente(objFilialCliente)
    If lErro <> SUCESSO Then gError 112705

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "CodClienteLoja", objFilialCliente.lCodClienteLoja, 0, "CodClienteLoja"
    colCampoValor.Add "CodFilialLoja", objFilialCliente.iCodFilialLoja, 0, "CodFilialLoja"
    colCampoValor.Add "Nome", objFilialCliente.sNome, STRING_FILIAL_CLIENTE_NOME, "Nome"
    colCampoValor.Add "RG", objFilialCliente.sRG, STRING_RG, "RG"
    colCampoValor.Add "CGC", objFilialCliente.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "Endereco", objFilialCliente.lEndereco, 0, "Endereco"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 112705

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160282)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialClienteEstatistica As New ClassFilialClienteEst

On Error GoTo Erro_Tela_Preenche

    objFilialCliente.lCodClienteLoja = colCampoValor.Item("CodClienteLoja").vValor

    'Se código de Cliente estiver preenchido,
    If objFilialCliente.lCodClienteLoja <> 0 Then

        'Passa os dados da coleção para o objeto
        objFilialCliente.iCodFilialLoja = colCampoValor.Item("CodFilialLoja").vValor
        objFilialCliente.sNome = colCampoValor.Item("Nome").vValor
        objFilialCliente.sCgc = colCampoValor.Item("CGC").vValor
        objFilialCliente.sRG = colCampoValor.Item("RG").vValor
        objFilialCliente.lEndereco = colCampoValor.Item("Endereco").vValor
        objFilialCliente.iFilialEmpresaLoja = giFilialEmpresa
                
        'Faz a leitura da estatistica da Filial Cliente
        lErro = CF("FilialCliente_Le_Loja", objFilialCliente)
        If lErro <> SUCESSO Then gError 112706
        
        'Exibe os dados da FilialCliente
        lErro = Exibe_Dados_FilialCliente(objFilialCliente)
        If lErro <> SUCESSO Then gError 112707

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 112706, 112707

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160283)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

 Dim lErro As Long

    Set objEventoCliente = Nothing
    Set objEventoFilialCliente = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    'lErro = ComandoSeta_Liberar(Me.Name)

    'Set objGridCategoria = Nothing

End Sub
