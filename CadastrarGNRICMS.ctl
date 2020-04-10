VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CadastrarGNRICMSOcx 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6525
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4095
      Index           =   1
      Left            =   180
      TabIndex        =   26
      Top             =   840
      Width           =   6225
      Begin VB.Frame Frame4 
         Caption         =   "SPED"
         Height          =   1140
         Left            =   60
         TabIndex        =   45
         Top             =   2895
         Width           =   6105
         Begin MSMask.MaskEdBox CodObrigRecolher 
            Height          =   315
            Left            =   1950
            TabIndex        =   10
            Top             =   270
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Format          =   "000"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodReceita 
            Height          =   315
            Left            =   1950
            TabIndex        =   11
            Top             =   720
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Código da Receita:"
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
            Left            =   270
            TabIndex        =   47
            Top             =   780
            Width           =   1635
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Obr. a Recolher:"
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
            Left            =   45
            TabIndex        =   46
            Top             =   330
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datas"
         Height          =   750
         Left            =   45
         TabIndex        =   32
         Top             =   2070
         Width           =   6120
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   270
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataRef 
            Height          =   285
            Left            =   4425
            TabIndex        =   8
            Top             =   270
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataVencimento 
            Height          =   300
            Left            =   2895
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownDataRef 
            Height          =   300
            Left            =   5400
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Referência:"
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
            Left            =   3420
            TabIndex        =   34
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
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
            Left            =   810
            TabIndex        =   33
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   315
         Left            =   2535
         Picture         =   "CadastrarGNRICMS.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   300
         Width           =   300
      End
      Begin VB.TextBox Convenio 
         Height          =   300
         Left            =   1965
         TabIndex        =   5
         Top             =   1740
         Width           =   1845
      End
      Begin VB.ComboBox TipoGNR 
         Height          =   315
         ItemData        =   "CadastrarGNRICMS.ctx":00EA
         Left            =   1965
         List            =   "CadastrarGNRICMS.ctx":00F4
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1255
         Width           =   1845
      End
      Begin MSComCtl2.UpDown UpDownDataPagamento 
         Height          =   300
         Left            =   2985
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   780
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataPagamento 
         Height          =   300
         Left            =   1965
         TabIndex        =   2
         Top             =   780
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1965
         TabIndex        =   0
         Top             =   300
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
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
         Height          =   195
         Left            =   1290
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   31
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Pagamento:"
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
         Left            =   180
         TabIndex        =   30
         Top             =   838
         Width           =   1755
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Convênio:"
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
         Left            =   1065
         TabIndex        =   29
         Top             =   1793
         Width           =   870
      End
      Begin VB.Label Label2 
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
         Left            =   1485
         TabIndex        =   28
         Top             =   1315
         Width           =   450
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4125
      Index           =   2
      Left            =   165
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   6225
      Begin VB.Frame Frame1 
         Caption         =   "Recolhimento"
         Height          =   1515
         Left            =   390
         TabIndex        =   40
         Top             =   2235
         Width           =   5445
         Begin VB.ComboBox Banco 
            Height          =   315
            ItemData        =   "CadastrarGNRICMS.ctx":0115
            Left            =   1230
            List            =   "CadastrarGNRICMS.ctx":0117
            TabIndex        =   16
            Text            =   "Banco"
            Top             =   210
            Width           =   2625
         End
         Begin MSMask.MaskEdBox Agencia 
            Height          =   285
            Left            =   3780
            TabIndex        =   18
            Top             =   675
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   285
            Left            =   1230
            TabIndex        =   17
            Top             =   675
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   285
            Left            =   1230
            TabIndex        =   19
            Top             =   1110
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   13
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
         Begin VB.Label Label10 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   660
            TabIndex        =   44
            Top             =   1170
            Width           =   510
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
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
            Left            =   450
            TabIndex        =   43
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
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
            Left            =   2985
            TabIndex        =   42
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   555
            TabIndex        =   41
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contribuinte"
         Height          =   1605
         Left            =   390
         TabIndex        =   35
         Top             =   390
         Width           =   5445
         Begin VB.ComboBox UFSubstTrib 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1890
            TabIndex        =   14
            Top             =   1170
            Width           =   735
         End
         Begin VB.ComboBox UFDestino 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4395
            TabIndex        =   15
            Top             =   1200
            Width           =   810
         End
         Begin MSMask.MaskEdBox CGC 
            Height          =   285
            Left            =   1890
            TabIndex        =   12
            Top             =   300
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   285
            Left            =   1890
            TabIndex        =   13
            Top             =   720
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "UF do Substituto:"
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
            Left            =   360
            TabIndex        =   39
            Top             =   1230
            Width           =   1500
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual:"
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
            Left            =   210
            TabIndex        =   38
            Top             =   780
            Width           =   1650
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
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
            Left            =   1410
            TabIndex        =   37
            Top             =   345
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "UF Favorecida:"
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
            Left            =   3030
            TabIndex        =   36
            Top             =   1260
            Width           =   1320
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4260
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CadastrarGNRICMS.ctx":0119
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CadastrarGNRICMS.ctx":0297
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CadastrarGNRICMS.ctx":07C9
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CadastrarGNRICMS.ctx":0953
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4545
      Left            =   150
      TabIndex        =   25
      Top             =   510
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   8017
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
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
Attribute VB_Name = "CadastrarGNRICMSOcx"
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

'Eventos dos Browses
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Function Trata_Parametros(Optional objGNRICMS As ClassGNRICMS) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se o objGNRICMS está preenchido
    If Not (objGNRICMS Is Nothing) Then
    
        'Se foi passado um código como parâmetro
        If objGNRICMS.lCodigo > 0 Then
        
            'Lê dados da Guia de ICMS
            lErro = CF("GNRICMS_Le", objGNRICMS)
            If lErro <> SUCESSO And lErro <> 70070 Then gError 70071
            
            'Se encontrou
            If lErro = SUCESSO Then
                
                'Traz os dados da Guia para a tela
                Call Traz_GNRICMS_Tela(objGNRICMS)
            
            Else
                'Se não coloca o Código na Tela
                Codigo.PromptInclude = False
                Codigo.Text = objGNRICMS.lCodigo
                Codigo.PromptInclude = True
            
            End If
        
        End If
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 70071 'Erro tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144039)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    'Eventos dos Browses
    Set objEventoCodigo = New AdmEvento
    
    'Carrega Bancos
    lErro = Carrega_Bancos()
    If lErro <> SUCESSO Then gError 70072
    
    'Carrega Estados
    lErro = Carrega_Estados()
    If lErro <> SUCESSO Then gError 70073
        
    'Seleciona o Tipo Apuracao como default
    TipoGNR.ListIndex = 0
        
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 70072, 70073 'Erros tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144040)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Carrega_Bancos() As Long
'Carrega a combo de Bancos

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Bancos

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 70074

   'Preenche ComboBox com código e nome dos Bancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        Banco.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Banco.ItemData(Banco.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Bancos = SUCESSO

    Exit Function

Erro_Carrega_Bancos:

    Carrega_Bancos = gErr

    Select Case gErr

        Case 70074 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144041)

    End Select

    Exit Function

End Function

Private Function Carrega_Estados() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox UFDestino

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Estados

    Set colSiglasUF = gcolUFs
    
    'Adiciona na Combo UFSubstTrib
    For iIndice = 1 To colSiglasUF.Count
        UFSubstTrib.AddItem colSiglasUF.Item(iIndice)
    Next

    'Adiciona na Combo UFDestino
    For iIndice = 1 To colSiglasUF.Count
        UFDestino.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_Estados = SUCESSO

    Exit Function

Erro_Carrega_Estados:

    Carrega_Estados = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144042)

    End Select

End Function

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
    Set objEventoCodigo = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Banco_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objBanco As New ClassBanco
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Banco_Validate

    'verifica se foi preenchido o Banco
    If Len(Trim(Banco.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox Banco
    If Banco.Text = Banco.List(Banco.ListIndex) Then Exit Sub

    'tenta Selecionar o banco com aquele codigo
    lErro = Combo_Seleciona(Banco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 69320
    
    If lErro = 6730 Then
    
        objBanco.iCodBanco = iCodigo
        
        'Verifica se o banco esta no BD
        lErro = CF("Banco_Le", objBanco)
        If lErro <> SUCESSO And lErro <> 16091 Then gError 69321
        
        If lErro = 16091 Then gError 69322
        
        Banco.AddItem CStr(objBanco.iCodBanco) & SEPARADOR & objBanco.sNomeReduzido
        Banco.ItemData(Banco.NewIndex) = objBanco.iCodBanco
        
        Banco.ListIndex = Banco.NewIndex
    
    End If
    
    If lErro = 6731 Then gError 69323
        
    Exit Sub

Erro_Banco_Validate:

    Cancel = True

    Select Case gErr

        Case 69320, 69321

        Case 69322
            'Se o banco nao estiver no BD pergunta se quer criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODBANCO_INEXISTENTE", objBanco.iCodBanco)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Bancos", objBanco)
            End If
            
        Case 69323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, Banco.Text)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144043)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático para próxima Guia de ICMS
    lErro = CF("GNRICMS_Codigo_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 70090
    
    Codigo.Text = CStr(lCodigo)
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 70090
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144044)
    
    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim colSelecao As New Collection
Dim objGNRICMS As New ClassGNRICMS
    
    'Se o Codigo está preechido
    If Len(Trim(Codigo.ClipText)) > 0 Then
        objGNRICMS.lCodigo = CLng(Codigo.Text)
    End If
        
    'Chama a Tela que lista a GNRICMSLista
    Call Chama_Tela("GNRICMSLista", colSelecao, objGNRICMS, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objGNRICMS As ClassGNRICMS

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objGNRICMS = obj1
    
    lErro = CF("GNRICMS_Le", objGNRICMS)
    If lErro <> SUCESSO And lErro <> 70070 Then gError 70071

    'Traz os dados da Guia de ICMS para a tela
    Call Traz_GNRICMS_Tela(objGNRICMS)

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 70071

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144045)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objGNRICMS As New ClassGNRICMS

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "GNRICMS"

    'Move os dados da tela para memória
    Call Move_Tela_Memoria(objGNRICMS)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objGNRICMS.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Codigo", objGNRICMS.lCodigo, 0, "Codigo"
    colCampoValor.Add "Tipo", objGNRICMS.iTipo, 0, "Tipo"
    colCampoValor.Add "DataPagto", objGNRICMS.dtDataPagto, 0, "DataPagto"
    colCampoValor.Add "CGCSubstTrib", objGNRICMS.sCGCSubstTrib, STRING_CGC, "CGCSubstTrib"
    colCampoValor.Add "InscricaoEstadual", objGNRICMS.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "UFSubstTrib", objGNRICMS.sUFSubstTrib, STRING_ESTADO, "UFSubstTrib"
    colCampoValor.Add "UFDestino", objGNRICMS.sUFDestino, STRING_ESTADO, "UFDestino"
    colCampoValor.Add "Banco", objGNRICMS.iBanco, 0, "Banco"
    colCampoValor.Add "Agencia", objGNRICMS.iAgencia, 0, "Agencia"
    colCampoValor.Add "Numero", objGNRICMS.sNumero, STRING_NUMERO, "Numero"
    colCampoValor.Add "Valor", objGNRICMS.dValor, 0, "Valor"
    colCampoValor.Add "Vencimento", objGNRICMS.dtVencimento, 0, "Vencimento"
    colCampoValor.Add "DataRef", objGNRICMS.dtDataRef, 0, "DataRef"
    colCampoValor.Add "Convenio", objGNRICMS.sConvenio, STRING_CONVENIO, "Convenio"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144046)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objGNRICMS As New ClassGNRICMS
Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    'Carrega objGNRICMS com os dados passados em colCampoValor
    objGNRICMS.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objGNRICMS.lCodigo = colCampoValor.Item("Codigo").vValor
    objGNRICMS.iTipo = colCampoValor.Item("Tipo").vValor
    objGNRICMS.dtDataPagto = colCampoValor.Item("DataPagto").vValor
    objGNRICMS.sCGCSubstTrib = colCampoValor.Item("CGCSubstTrib").vValor
    objGNRICMS.sInscricaoEstadual = colCampoValor.Item("InscricaoEstadual").vValor
    objGNRICMS.sUFSubstTrib = colCampoValor.Item("UFSubstTrib").vValor
    objGNRICMS.sUFDestino = colCampoValor.Item("UFDestino").vValor
    objGNRICMS.iBanco = colCampoValor.Item("Banco").vValor
    objGNRICMS.iAgencia = colCampoValor.Item("Agencia").vValor
    objGNRICMS.sNumero = colCampoValor.Item("Numero").vValor
    objGNRICMS.dValor = colCampoValor.Item("Valor").vValor
    objGNRICMS.dtVencimento = colCampoValor.Item("Vencimento").vValor
    objGNRICMS.dtDataRef = colCampoValor.Item("DataRef").vValor
    objGNRICMS.sConvenio = colCampoValor.Item("Convenio").vValor

    'Se o NumIntDoc estiver preenchido
    If objGNRICMS.lNumIntDoc <> 0 Then

        lErro = CF("GNRICMS_Le", objGNRICMS)
        If lErro <> SUCESSO And lErro <> 70070 Then gError 70071

        'Traz os dados da Guia ICMS para a tela tela
        Call Traz_GNRICMS_Tela(objGNRICMS)

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 70071

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144047)

    End Select

    Exit Sub

End Sub

Sub Traz_GNRICMS_Tela(objGNRICMS As ClassGNRICMS)
'Traz os dados da Guia de ICMS para a tela

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela_GNRICMS
    
    'Dados Principais
    Codigo.Text = objGNRICMS.lCodigo
    
    DataPagamento.PromptInclude = False
    DataPagamento.Text = Format(objGNRICMS.dtDataPagto, "dd/mm/yy")
    DataPagamento.PromptInclude = True
    
    'Seleciona o Tipo de Guia
    For iIndice = 0 To TipoGNR.ListCount - 1
        If TipoGNR.ItemData(iIndice) = objGNRICMS.iTipo Then
            TipoGNR.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Convenio.Text = objGNRICMS.sConvenio
    
    DataVencimento.PromptInclude = False
    DataVencimento.Text = Format(objGNRICMS.dtVencimento, "dd/mm/yy")
    DataVencimento.PromptInclude = True
    
    DataRef.PromptInclude = False
    DataRef.Text = Format(objGNRICMS.dtDataRef, "dd/mm/yy")
    DataRef.PromptInclude = True
    
    'Complemento
    CGC.Text = objGNRICMS.sCGCSubstTrib
    InscricaoEstadual.Text = objGNRICMS.sInscricaoEstadual
    UFSubstTrib.Text = objGNRICMS.sUFSubstTrib
    UFDestino.Text = objGNRICMS.sUFDestino
        
    'Seleciona o Banco
    For iIndice = 0 To Banco.ListCount - 1
        If Banco.ItemData(iIndice) = objGNRICMS.iBanco Then
            Banco.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Numero.Text = objGNRICMS.sNumero
    
    If objGNRICMS.iAgencia > 0 Then
        Agencia.Text = objGNRICMS.iAgencia
    End If
    
    Valor.Text = Format(objGNRICMS.dValor, "Standard")
    
    CodReceita.PromptInclude = False
    CodReceita.Text = Format(objGNRICMS.sCodReceita, CodReceita.Format)
    CodReceita.PromptInclude = True
    CodObrigRecolher.PromptInclude = False
    CodObrigRecolher.Text = Format(objGNRICMS.sCodObrigRecolher, CodObrigRecolher.Format)
    CodObrigRecolher.PromptInclude = True
    
    iAlterado = 0

End Sub

Function Move_Tela_Memoria(objGNRICMS As ClassGNRICMS) As Long
'Move dados da tela para a memória
    
    'Dados Principais
    objGNRICMS.lCodigo = StrParaLong(Codigo.Text)
    objGNRICMS.dtDataPagto = StrParaDate(DataPagamento.Text)
    
    If TipoGNR.ListIndex <> -1 Then
        objGNRICMS.iTipo = TipoGNR.ItemData(TipoGNR.ListIndex)
    End If
    
    objGNRICMS.sConvenio = Convenio.Text
    objGNRICMS.dtVencimento = StrParaDate(DataVencimento.Text)
    objGNRICMS.dtDataRef = StrParaDate(DataRef.Text)
    
    'Complemento
    objGNRICMS.sCGCSubstTrib = CGC.Text
    objGNRICMS.sInscricaoEstadual = InscricaoEstadual.Text
    objGNRICMS.sUFSubstTrib = UFSubstTrib.Text
    objGNRICMS.sUFDestino = UFDestino.Text
    
    If Banco.ListIndex <> -1 Then
        objGNRICMS.iBanco = Banco.ItemData(Banco.ListIndex)
    End If

    objGNRICMS.sNumero = Numero.Text
    objGNRICMS.iAgencia = StrParaInt(Agencia.Text)
    objGNRICMS.dValor = StrParaDbl(Valor.Text)
    objGNRICMS.iFilialEmpresa = giFilialEmpresa
    objGNRICMS.sCodReceita = Format(CodReceita.Text, CodReceita.Format)
    objGNRICMS.sCodObrigRecolher = Format(CodObrigRecolher.Text, CodObrigRecolher.Format)
    
End Function

Private Sub DataPagamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataPagamento_Validate

    'Se a DataPagamento está preenchida
    If Len(Trim(DataPagamento.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataPagamento.Text)
        If lErro <> SUCESSO Then gError 70075

    End If

    Exit Sub

Erro_DataPagamento_Validate:

    Cancel = True

    Select Case gErr

        Case 70075

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144048)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPagamento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataPagamento_DownClick

    'Se a data está preenchida
    If Len(Trim(DataPagamento.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataPagamento, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70076

    End If

    Exit Sub

Erro_UpDownDataPagamento_DownClick:

    Select Case gErr

        Case 70076

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144049)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPagamento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataPagamento_UpClick

    'Se a data está preenchida
    If Len(Trim(DataPagamento.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataPagamento, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70077

    End If

    Exit Sub

Erro_UpDownDataPagamento_UpClick:

    Select Case gErr

        Case 70077

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144050)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Se a DataVencimento está preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 70078

    End If

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True

    Select Case gErr

        Case 70078

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144051)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    'Se a data está preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70079

    End If

    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case gErr

        Case 70079

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144052)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    'Se a data está preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70080

    End If

    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case gErr

        Case 70080

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144053)

    End Select

    Exit Sub

End Sub

Private Sub DataRef_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRef_Validate

    'Se a DataRef está preenchida
    If Len(Trim(DataRef.ClipText)) > 0 Then

        'Critica seu formato
        lErro = Data_Critica(DataRef.Text)
        If lErro <> SUCESSO Then gError 70081

    End If

    Exit Sub

Erro_DataRef_Validate:

    Cancel = True

    Select Case gErr

        Case 70081

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144054)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataRef_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataRef_DownClick

    'Se a data está preenchida
    If Len(Trim(DataRef.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(DataRef, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 70082

    End If

    Exit Sub

Erro_UpDownDataRef_DownClick:

    Select Case gErr

        Case 70082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144055)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataRef_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataRef_UpClick

    'Se a data está preenchida
    If Len(Trim(DataRef.ClipText)) > 0 Then

        'Aumenta a data em um dia
        lErro = Data_Up_Down_Click(DataRef, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 70083

    End If

    Exit Sub

Erro_UpDownDataRef_UpClick:

    Select Case gErr

        Case 70083

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144056)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Se Valor foi preenchido
    If Len(Trim(Valor.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 70084

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 70084

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144057)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Se Codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then gError 70085

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 70085

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144058)

    End Select

    Exit Sub

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Se o campo foi preenchido
    If Len(Trim(Numero.ClipText)) > 0 Then

        'Critica o Número
        lErro = Valor_Positivo_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError 70086
        
    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 70086
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144059)

    End Select

    Exit Sub

End Sub

Private Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub

    Select Case Len(Trim(CGC.Text))

        Case STRING_CPF 'CPF
    
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 70087
    
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text
    
        Case STRING_CGC  'CGC
    
            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then gError 70088
    
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text
    
    Case Else
        gError 70089

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True

    Select Case gErr

        Case 70087, 70088

        Case 70089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144060)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub DataPagamento_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataPagamento, iAlterado)

End Sub

Private Sub DataVencimento_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataVencimento, iAlterado)

End Sub

Private Sub DataRef_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataRef, iAlterado)

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub CGC_GotFocus()

    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Private Sub Agencia_GotFocus()

    Call MaskEdBox_TrataGotFocus(Agencia, iAlterado)

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoGNR_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Convenio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataRef_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFSubstTrib_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFSubstTrib_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFDestino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFDestino_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UFDestino_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFDestino_Validate

    'Verifica se tem alguma Placa U.F. foi preenchida
    If Len(Trim(UFDestino.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual_CI(UFDestino)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 70138

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 70139

    Exit Sub

Erro_UFDestino_Validate:

    Cancel = True

    Select Case gErr

        Case 70138

        Case 70139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, UFDestino.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144061)

    End Select

    Exit Sub

End Sub

Private Sub UFSubstTrib_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFSubstTrib_Validate

    'Verifica se tem alguma Placa U.F. foi preenchida
    If Len(Trim(UFSubstTrib.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual_CI(UFSubstTrib)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 70140

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 70141

    Exit Sub

Erro_UFSubstTrib_Validate:

    Cancel = True

    Select Case gErr

        Case 70140

        Case 70141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, UFSubstTrib.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144062)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
    
    End If
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava uma de apuraçao ICMS
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 70085

    'Limpa a tela
    Call Limpa_Tela_GNRICMS

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 70085

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144063)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objGNRICMS As New ClassGNRICMS

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 70093
    
    'Verifica se a Data de Pagamento foi preenchida
    If Len(Trim(DataPagamento.ClipText)) = 0 Then gError 70094
    
    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoGNR.Text)) = 0 Then gError 70095
    
    'Verifica se a data de Vencimento foi preenchida
    If Len(Trim(DataVencimento.ClipText)) = 0 Then gError 70096
    
    'Verifica se a Data de Referência foi preenchida
    If Len(Trim(DataRef.ClipText)) = 0 Then gError 70097
    
    'Verifica se o CGC foi preenchido
    If Len(Trim(CGC.ClipText)) = 0 Then gError 70098
    
    'Verifica se o valor foi preenchido
    If Len(Trim(Valor.Text)) = 0 Then gError 70099
    
    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objGNRICMS)

    'Grava um uma Guia de ICMS
    lErro = CF("GNRICMS_Grava", objGNRICMS)
    If lErro <> SUCESSO Then gError 70100

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 70093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 70094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_PAGAMENTO_NAO_PREENCHIDO", gErr)
                
        Case 70095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOGNR_NAO_PREENCHIDO", gErr)
            
        Case 70096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCIMENTO_NAO_PREENCHIDA", gErr)
            
        Case 70097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_REFERENCIA_NAO_PREENCHIDA", gErr)
                
        Case 70098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CGC_NAO_INFORMADO", gErr)
                
        Case 70099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
            
        Case 70100
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144064)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objGNRICMS As New ClassGNRICMS

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 70101
    
    'Guarda Código da Guia de ICMS
    objGNRICMS.lCodigo = CLng(Codigo.Text)
    
    'Lê a Guia de ICMS a partir do código
    lErro = CF("GNRICMS_Le", objGNRICMS)
    If lErro <> SUCESSO And lErro <> 70070 Then gError 70102
    
    'Se a guia não está cadastrada, erro
    If lErro = 70070 Then gError 70103
    
    'Pede a confirmação da exclusão da Guia de ICMS
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_GNRICMS", objGNRICMS.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui a Guia de ICMS
    lErro = CF("GNRICMS_Exclui", objGNRICMS)
    If lErro <> SUCESSO Then gError 70104

    'Limpa a tela
    Call Limpa_Tela_GNRICMS

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 70101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
                
        Case 70102, 70104
        
        Case 70103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GNRICMS_NAO_CADASTRADA", gErr, objGNRICMS.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144065)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 70105

    'Limpa a tela
    Call Limpa_Tela_GNRICMS

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 70105

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144066)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_GNRICMS()

    'Função Genérica que limpa a tela
    Call Limpa_Tela(Me)

    'Limpa o restante da tela
    Codigo.Text = ""
    
    DataPagamento.PromptInclude = False
    DataPagamento.Text = ""
    DataPagamento.PromptInclude = True
    
    TipoGNR.ListIndex = 0
    Convenio.Text = ""
    
    DataVencimento.PromptInclude = False
    DataVencimento.Text = ""
    DataVencimento.PromptInclude = True
    
    DataRef.PromptInclude = False
    DataRef.Text = ""
    DataRef.PromptInclude = True
    
    UFSubstTrib.Text = ""
    UFDestino.Text = ""
    
    CGC.Text = ""
    Banco.ListIndex = -1
    Numero.Text = ""
    Agencia.Text = ""
        
    iAlterado = 0
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cadastro de Guias de ICMS"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CadastrarGNRICMS"

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




Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub CodObrigRecolher_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodReceita_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodObrigRecolher_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodObrigRecolher, iAlterado)
End Sub
