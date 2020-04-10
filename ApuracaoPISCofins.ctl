VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ApuracaoPISCofinsOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7305
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   30
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "ApuracaoPISCofins.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   555
         Picture         =   "ApuracaoPISCofins.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "ApuracaoPISCofins.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1575
         Picture         =   "ApuracaoPISCofins.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valores Nesse Período de Escrituração"
      Height          =   2220
      Left            =   45
      TabIndex        =   34
      Top             =   3705
      Width           =   9450
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   12
         Left            =   7770
         TabIndex        =   35
         Top             =   165
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   13
         Left            =   7770
         TabIndex        =   10
         Top             =   450
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   14
         Left            =   7770
         TabIndex        =   11
         Top             =   735
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   15
         Left            =   7770
         TabIndex        =   12
         Top             =   1020
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   16
         Left            =   7770
         TabIndex        =   13
         Top             =   1305
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   17
         Left            =   7770
         TabIndex        =   14
         Top             =   1590
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   18
         Left            =   7770
         TabIndex        =   36
         Top             =   1875
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Crédito objeto de Pedido de Ressarcimento (PER):"
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
         Index           =   12
         Left            =   2550
         TabIndex        =   43
         Top             =   780
         Width           =   5055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vlr do Crédito por Declaração de Compensação Intermediária:"
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
         Index           =   13
         Left            =   2355
         TabIndex        =   42
         Top             =   1065
         Width           =   5250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo do Crédito Disponível para Utilização:"
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
         Index           =   10
         Left            =   3825
         TabIndex        =   41
         Top             =   195
         Width           =   3780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Crédito descontado:"
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
         Index           =   11
         Left            =   5130
         TabIndex        =   40
         Top             =   465
         Width           =   2475
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Crédito transf. em evento de cisão, fusão ou incorporação:"
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
         Index           =   14
         Left            =   2310
         TabIndex        =   39
         Top             =   1335
         Width           =   5295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor do crédito utilizado por outras formas:"
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
         Index           =   15
         Left            =   3900
         TabIndex        =   38
         Top             =   1605
         Width           =   3705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo de Créditos a Utilizar em Período de Apuração Futuro:"
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
         Index           =   16
         Left            =   2460
         TabIndex        =   37
         Top             =   1905
         Width           =   5145
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Valores - Períodos Anteriores"
      Height          =   1920
      Left            =   45
      TabIndex        =   26
      Top             =   1740
      Width           =   9435
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   6
         Left            =   7755
         TabIndex        =   5
         Top             =   135
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   7
         Left            =   7755
         TabIndex        =   6
         Top             =   420
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   8
         Left            =   7755
         TabIndex        =   27
         Top             =   705
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   9
         Left            =   7755
         TabIndex        =   7
         Top             =   990
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   10
         Left            =   7755
         TabIndex        =   8
         Top             =   1275
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   270
         Index           =   11
         Left            =   7755
         TabIndex        =   9
         Top             =   1560
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   476
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vlr do Crédito por Declaração de Compensação Intermediária:"
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
         Index           =   9
         Left            =   2385
         TabIndex        =   33
         Top             =   1575
         Width           =   5250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vlr do Crédito Utilizado Mediante Pedido de Ressarcimento:"
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
         Index           =   8
         Left            =   2565
         TabIndex        =   32
         Top             =   1290
         Width           =   5070
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Crédito Utilizado Mediante Desconto:"
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
         Index           =   7
         Left            =   3705
         TabIndex        =   31
         Top             =   1020
         Width           =   3930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total do Crédito Apurado:"
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
         Left            =   4950
         TabIndex        =   30
         Top             =   735
         Width           =   2685
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Crédito Extemporâneo Apurado:"
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
         Left            =   4185
         TabIndex        =   29
         Top             =   465
         Width           =   3450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Crédito Apurado:"
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
         Left            =   5445
         TabIndex        =   28
         Top             =   180
         Width           =   2190
      End
   End
   Begin VB.CommandButton BotaoApuracaoCadastradas 
      Caption         =   "Apurações Cadastradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4935
      TabIndex        =   19
      Top             =   75
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apuração"
      Height          =   1200
      Left            =   30
      TabIndex        =   22
      Top             =   525
      Width           =   9450
      Begin VB.ComboBox Mes 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "ApuracaoPISCofins.ctx":0994
         Left            =   1740
         List            =   "ApuracaoPISCofins.ctx":09BF
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   135
         Width           =   1860
      End
      Begin VB.CommandButton BotaoTrazDadosPerAnterior 
         Caption         =   "Traz Dados de Períodos Anteriores"
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
         Left            =   5160
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox Origem 
         Height          =   315
         ItemData        =   "ApuracaoPISCofins.ctx":0A28
         Left            =   1740
         List            =   "ApuracaoPISCofins.ctx":0A32
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox CodCred 
         Height          =   315
         ItemData        =   "ApuracaoPISCofins.ctx":0A98
         Left            =   1740
         List            =   "ApuracaoPISCofins.ctx":0AA8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   7650
      End
      Begin MSMask.MaskEdBox CGC 
         Height          =   315
         Left            =   7350
         TabIndex        =   3
         Top             =   480
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Ano 
         Height          =   315
         Left            =   4350
         TabIndex        =   1
         Top             =   135
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mês:"
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
         Left            =   1275
         TabIndex        =   45
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label2 
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
         Left            =   3915
         TabIndex        =   44
         Top             =   195
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código do Crédito:"
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
         Index           =   3
         Left            =   135
         TabIndex        =   25
         Top             =   870
         Width           =   1575
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ do Cedente:"
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
         Left            =   5760
         TabIndex        =   24
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origem:"
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
         Index           =   2
         Left            =   1020
         TabIndex        =   23
         Top             =   510
         Width           =   660
      End
   End
End
Attribute VB_Name = "ApuracaoPISCofinsOcx"
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
Dim giTipo As Integer
Dim iMesAnt As Integer
Dim iAnoAnt As Integer

'Eventos dos Browses
Private WithEvents objEventoBotaoApuracao As AdmEvento
Attribute objEventoBotaoApuracao.VB_VarHelpID = -1

Function Trata_Parametros(Optional ByVal objApuracao As ClassRegApuracaoPISCofins, Optional ByVal iTipo As Integer = 0) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giTipo = iTipo
    
    'Se foi passada uma Apuração Pis/Cofins como parâmetro
    If Not objApuracao Is Nothing Then
    
        If objApuracao.iTipo <> 0 Then
            giTipo = objApuracao.iTipo
        End If
               
        'Guarda a Filial Empresa
        objApuracao.iFilialEmpresa = giFilialEmpresa
        
        'Traz os dados da Apuração Pis/Cofins para a tela
        lErro = Traz_ApuracaoPisCofins_Tela(objApuracao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If giTipo = REG_APUR_PIS_COFINS_TIPO_PIS Then
        Caption = "Apuração de Pis"
    Else
        Caption = "Apuração de Cofins"
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213024)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    'Eventos dos Browses
    Set objEventoBotaoApuracao = New AdmEvento
        
    'Traz os dados default da Empresa e da Apuração Pis/Cofins
    lErro = Traz_Dados_Default()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213025)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera variáveis globais
    Set objEventoBotaoApuracao = Nothing

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoApuracaoCadastradas_Click()

Dim colSelecao As New Collection
Dim objApuracao As ClassRegApuracaoPISCofins

    colSelecao.Add giTipo
    
    Call Chama_Tela("RegApuracaoPisCofinsLista", colSelecao, objApuracao, objEventoBotaoApuracao)

End Sub

Private Sub objEventoBotaoApuracao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objApuracao As ClassRegApuracaoPISCofins

On Error GoTo Erro_objEventoBotaoApuracao_evSelecao

    Set objApuracao = obj1

    'Traz os dados da Apuração Pis/Cofins para a tela
    lErro = Traz_ApuracaoPisCofins_Tela(objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoBotaoApuracao_evSelecao:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213026)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objApuracao As New ClassRegApuracaoPISCofins

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RegApuracaoPisCofins"

    'Move os dados da tela para memória
    lErro = Move_Tela_Memoria(objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objApuracao.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Tipo", objApuracao.iTipo, 0, "Tipo"
    colCampoValor.Add "FilialEmpresa", objApuracao.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Ano", objApuracao.iAno, 0, "Ano"
    colCampoValor.Add "Mes", objApuracao.iMes, 0, "Mes"
    colCampoValor.Add "OrigCred", objApuracao.iOrigCred, 0, "OrigCred"
    colCampoValor.Add "CNPJCedCred", objApuracao.sCNPJCedCred, STRING_CGC, "CNPJCedCred"
    colCampoValor.Add "CodCred", objApuracao.sCodCred, STRING_MAXIMO, "CodCred"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Tipo", OP_IGUAL, objApuracao.iTipo
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213027)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objApuracao As New ClassRegApuracaoPISCofins

On Error GoTo Erro_Tela_Preenche

    'Carrega objApuracao com os dados passados em colCampoValor
    objApuracao.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    'Traz os dados dos itens de apuração Pis/Cofins para a tela tela
    lErro = Traz_ApuracaoPisCofins_Tela(objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213028)

    End Select

    Exit Sub

End Sub

Function Traz_ApuracaoPisCofins_Tela(objApuracao As ClassRegApuracaoPISCofins) As Long
'Traz os dados da Apuração Pis/Cofins para a tela

Dim lErro As Long

On Error GoTo Erro_Traz_ApuracaoPisCofins_Tela

    lErro = CF("RegApuracaoPisCofins_Le", objApuracao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then

        Call Combo_Seleciona_ItemData(Mes, objApuracao.iMes)

        Ano.PromptInclude = False
        Ano.Text = objApuracao.iAno
        Ano.PromptInclude = True
        
        Call Combo_Seleciona_ItemData(Origem, objApuracao.iOrigCred)
        
        Call CF("sCombo_Seleciona2", CodCred, objApuracao.sCodCred)
        
        CGC.Text = objApuracao.sCNPJCedCred
        Call CGC_Validate(bSGECancelDummy)
       
        Valor(6).Text = Format(objApuracao.dVlCredApu, "Standard")
        Valor(7).Text = Format(objApuracao.dVlCredExtApu, "Standard")
        Valor(8).Text = Format(objApuracao.dVlTotCredApu, "Standard")
        Valor(9).Text = Format(objApuracao.dVlCredDescPAAnt, "Standard")
        Valor(10).Text = Format(objApuracao.dVlCredPerPAAnt, "Standard")
        Valor(11).Text = Format(objApuracao.dVlCredDCompPAAnt, "Standard")
        Valor(12).Text = Format(objApuracao.dSdCredDispEFD, "Standard")
        Valor(13).Text = Format(objApuracao.dVlCredDescEFD, "Standard")
        Valor(14).Text = Format(objApuracao.dVlCredPerEFD, "Standard")
        Valor(15).Text = Format(objApuracao.dVlCredDCompEFD, "Standard")
        Valor(16).Text = Format(objApuracao.dVlCredTrans, "Standard")
        Valor(17).Text = Format(objApuracao.dVlCredOut, "Standard")
        Valor(18).Text = Format(objApuracao.dSdCredFim, "Standard")

    End If
        
    iAlterado = 0
    
    Traz_ApuracaoPisCofins_Tela = SUCESSO
    
    Exit Function

Erro_Traz_ApuracaoPisCofins_Tela:

    Traz_ApuracaoPisCofins_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213029)
    
    End Select
    
    Exit Function
    
End Function

Function Move_Tela_Memoria(objApuracao As ClassRegApuracaoPISCofins) As Long
'Move dados da tela para a memória

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Move_Tela_Memoria

    'Move dados da Apuração
    objApuracao.iFilialEmpresa = giFilialEmpresa
    objApuracao.iTipo = giTipo
    objApuracao.iMes = Mes.ItemData(Mes.ListIndex)
    objApuracao.iAno = StrParaInt(Ano.Text)
    objApuracao.iOrigCred = Codigo_Extrai(Origem.Text)
    objApuracao.sCodCred = Codigo_Extrai(CodCred.Text)
    objApuracao.sCNPJCedCred = Trim(CGC.Text)
        
    objApuracao.dVlCredApu = StrParaDbl(Valor(6).Text)
    objApuracao.dVlCredExtApu = StrParaDbl(Valor(7).Text)
    objApuracao.dVlTotCredApu = StrParaDbl(Valor(8).Text)
    objApuracao.dVlCredDescPAAnt = StrParaDbl(Valor(9).Text)
    objApuracao.dVlCredPerPAAnt = StrParaDbl(Valor(10).Text)
    objApuracao.dVlCredDCompPAAnt = StrParaDbl(Valor(11).Text)
    objApuracao.dSdCredDispEFD = StrParaDbl(Valor(12).Text)
    objApuracao.dVlCredDescEFD = StrParaDbl(Valor(13).Text)
    objApuracao.dVlCredPerEFD = StrParaDbl(Valor(14).Text)
    objApuracao.dVlCredDCompEFD = StrParaDbl(Valor(15).Text)
    objApuracao.dVlCredTrans = StrParaDbl(Valor(16).Text)
    objApuracao.dVlCredOut = StrParaDbl(Valor(17).Text)
    objApuracao.dSdCredFim = StrParaDbl(Valor(18).Text)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213030)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Origem_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Origem
End Sub

Private Sub Origem_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Origem
End Sub

Private Sub Valor_Change(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Valor_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    If Len(Trim(Valor(Index).ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(Valor(Index).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Valor(8).Text = Format(StrParaDbl(Valor(6)) + StrParaDbl(Valor(7)), "Standard")
    Valor(12).Text = Format(StrParaDbl(Valor(8)) - StrParaDbl(Valor(9)) - StrParaDbl(Valor(10)) - StrParaDbl(Valor(11)), "Standard")
    Valor(18).Text = Format(StrParaDbl(Valor(12)) - StrParaDbl(Valor(13)) - StrParaDbl(Valor(14)) - StrParaDbl(Valor(15)) - StrParaDbl(Valor(16)) - StrParaDbl(Valor(17)), "Standard")
    
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213031)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava uma de apuraçao Pis/Cofins
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a tela
    Call Limpa_Tela_ApuracaoPisCofins

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213032)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objApuracao As New ClassRegApuracaoPISCofins

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Valida_Dados_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Grava um Registro de apuração Pis/Cofins
    lErro = CF("RegApuracaoPisCofins_Grava", objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213033)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objApuracao As New ClassRegApuracaoPISCofins

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Valida_Dados_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'Lê a Apuração Pis/Cofins a partir da FilialEmpresa, DataInicial e DataFinal
    lErro = CF("ApuracaoPisCofins_Le", objApuracao)
    If lErro <> SUCESSO And lErro <> 70013 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrou, erro
    If lErro = 70013 Then gError 213034

    'Pede a confirmação da exclusão da apuração de Pis/Cofins
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_REGAPURACAOPISCOFINS")
    If vbMsgRes <> vbNo Then
    
        'Exclui a apuração de Pis/Cofins
        lErro = CF("RegApuracaoPisCofins_Exclui", objApuracao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Limpa a tela
        Call Limpa_Tela_ApuracaoPisCofins
    
        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)
    
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213034
            Call Rotina_Erro(vbOKOnly, "ERRO_REGAPURACAOPISCOFINS_NAO_CADASTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213035)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a tela
    Call Limpa_Tela_ApuracaoPisCofins

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213036)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_ApuracaoPisCofins() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ApuracaoPisCofins

    'Função Genérica que limpa a tela
    Call Limpa_Tela(Me)
    
    'Traz os dados default da Empresa e da Apuração Pis/Cofins
    lErro = Traz_Dados_Default()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Limpa_Tela_ApuracaoPisCofins = SUCESSO
    
    Exit Function
    
Erro_Limpa_Tela_ApuracaoPisCofins:

    Limpa_Tela_ApuracaoPisCofins = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213037)
    
    End Select
    
    Exit Function
    
End Function

Function Traz_Dados_Default() As Long

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_Traz_Dados_Default

    dtData = DateAdd("d", -Day(Date), Date)
    
    Call Combo_Seleciona_ItemData(Mes, Month(dtData))
    Ano.PromptInclude = False
    Ano.Text = CStr(Year(dtData))
    Ano.PromptInclude = True

    lErro = Trata_EFD_Tabelas
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    CodCred.ListIndex = -1
    
    Origem.ListIndex = 0
    CGC.Enabled = False
    
    Traz_Dados_Default = SUCESSO
        
    Exit Function
    
Erro_Traz_Dados_Default:

    Traz_Dados_Default = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213038)
    
    End Select
    
    Exit Function
    
End Function

Function Trata_EFD_Tabelas() As Long

Dim lErro As Long
Dim dtData As Date
Dim sData As String
Dim iMes As Integer, iAno As Integer

On Error GoTo Erro_Trata_EFD_Tabelas

    iAno = StrParaInt(Ano.Text)
    iMes = Mes.ItemData(Mes.ListIndex)

    If iMes <> iMesAnt Or iAno <> iAnoAnt Then

        sData = "01/" & Format(iMes, "00") & "/" & CStr(iAno)
        
        'Pega o último dia do mês
        dtData = StrParaDate(sData)
        dtData = DateAdd("d", -1, DateAdd("m", 1, dtData))
        
        lErro = CF("EFDTabelas_Carrega_Combo", CodCred, "4.3.6", EFD_TABELAS_OBRIG_PISCOFINS, dtData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        iMesAnt = iMes
        iAnoAnt = iAno
        
    End If

    Trata_EFD_Tabelas = SUCESSO
        
    Exit Function
    
Erro_Trata_EFD_Tabelas:

    Trata_EFD_Tabelas = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213053)
    
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    If giTipo = REG_APUR_PIS_COFINS_TIPO_PIS Then
        Caption = "Apuração de Pis"
    Else
        Caption = "Apuração de Cofins"
    End If
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ApuracaoPisCofins"

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

Public Sub CGC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CGC_GotFocus()
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)
End Sub

Public Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate
    
    'Se CGC/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 213039
            
            'Formata e coloca na Tela
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text

        Case STRING_CGC 'CGC
            
            'Critica CGC
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 213040
            
            'Formata e Coloca na Tela
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text

        Case Else
                
            gError 213041

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True

    Select Case Err

        Case 213039, 213040

        Case 213041
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213042)

    End Select

    Exit Sub

End Sub

Private Sub Ano_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)

End Sub

Private Sub Ano_Validate(bCancel As Boolean)

On Error GoTo Erro_Ano_Validate

    'If Len(Trim(Ano.Text)) > 0 Then

        If Ano.Text < 1900 Then gError 213051
        
    'End If
    
    Exit Sub
    
Erro_Ano_Validate:

    Select Case gErr
    
        Case 213051
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_INVALIDO", gErr)
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213052)

    End Select
    
    Exit Sub
    
End Sub

Public Function Valida_Dados_Tela() As Long

Dim lErro As Long

On Error GoTo Erro_Valida_Dados_Tela

    If CodCred.ListIndex = -1 Then gError 213055
    If Codigo_Extrai(Origem.Text) = 2 And Len(Trim(CGC.Text)) = 0 Then gError 213056

    Valida_Dados_Tela = SUCESSO

    Exit Function

Erro_Valida_Dados_Tela:

    Valida_Dados_Tela = gErr

    Select Case gErr
    
        Case 213055
            Call Rotina_Erro(vbOKOnly, "ERRO_COD_CRED_NAO_PREENCHINDO", gErr)
        
        Case 213056
             Call Rotina_Erro(vbOKOnly, "ERRO_CNPJ_CEDENTE_OBRIGATORIO_PARA_ORIG", gErr)
       
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213054)

    End Select
    
    Exit Function

End Function

Sub Trata_Origem()
    
    If Codigo_Extrai(Origem.Text) = 1 Then
        CGC.Enabled = False
        'CGC.PromptInclude = False
        CGC.Text = ""
        'CGC.PromptInclude = True
    Else
        CGC.Enabled = True
    End If
End Sub

Private Sub BotaoTrazDadosPerAnterior_Click()

Dim lErro As Long
Dim objApuracao As New ClassRegApuracaoPISCofins

On Error GoTo Erro_BotaoTrazDadosPerAnterior_Click

    lErro = Valida_Dados_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objApuracao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If objApuracao.iMes = 1 Then
        objApuracao.iMes = 12
        objApuracao.iAno = objApuracao.iAno - 1
    Else
        objApuracao.iMes = objApuracao.iMes - 1
    End If
    
    lErro = CF("RegApuracaoPisCofins_Le", objApuracao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 9999 'Não existe apuração anterior para essa Origem/CNPJ/Cod.Crédito
    
'    Valor(6).Text = Format(objApuracao.dVlCredApu, "Standard")
'    Valor(7).Text = Format(objApuracao.dVlCredExtApu, "Standard")
'    Valor(8).Text = Format(objApuracao.dVlTotCredApu, "Standard")
'    Valor(9).Text = Format(objApuracao.dVlCredDescPAAnt, "Standard")
'    Valor(10).Text = Format(objApuracao.dVlCredPerPAAnt, "Standard")
'    Valor(11).Text = Format(objApuracao.dVlCredDCompPAAnt, "Standard")
'    Valor(12).Text = Format(objApuracao.dSdCredDispEFD, "Standard")
'    Valor(13).Text = Format(objApuracao.dVlCredDescEFD, "Standard")
'    Valor(14).Text = Format(objApuracao.dVlCredPerEFD, "Standard")
'    Valor(15).Text = Format(objApuracao.dVlCredDCompEFD, "Standard")
'    Valor(16).Text = Format(objApuracao.dVlCredTrans, "Standard")
'    Valor(17).Text = Format(objApuracao.dVlCredOut, "Standard")
'    Valor(18).Text = Format(objApuracao.dSdCredFim, "Standard")
    
    Call Valor_Validate(6, bSGECancelDummy)

    Exit Sub

Erro_BotaoTrazDadosPerAnterior_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213057)

    End Select

    Exit Sub

End Sub
