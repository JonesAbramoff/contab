VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ChequePreOcx 
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   8910
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   4440
      TabIndex        =   43
      Top             =   3960
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
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
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TaxaJuros 
      Height          =   225
      Left            =   4545
      TabIndex        =   40
      Top             =   3315
      Width           =   840
      _ExtentX        =   1482
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
      Format          =   "#,#####0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Desconto 
      Height          =   225
      Left            =   7500
      TabIndex        =   26
      Top             =   2610
      Width           =   825
      _ExtentX        =   1455
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
   Begin MSMask.MaskEdBox Multa 
      Height          =   225
      Left            =   6810
      TabIndex        =   25
      Top             =   2625
      Width           =   705
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
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Juros 
      Height          =   225
      Left            =   5910
      TabIndex        =   24
      Top             =   2580
      Width           =   840
      _ExtentX        =   1482
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
   Begin MSMask.MaskEdBox SaldoParcela 
      Height          =   255
      Left            =   3615
      TabIndex        =   22
      Top             =   2520
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   450
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
   Begin MSMask.MaskEdBox ValorReceber 
      Height          =   255
      Left            =   4710
      TabIndex        =   23
      Top             =   2535
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   450
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
   Begin MSMask.MaskEdBox DataVencimentoReal 
      Height          =   225
      Left            =   1140
      TabIndex        =   19
      Top             =   2520
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cheque"
      Height          =   1455
      Left            =   90
      TabIndex        =   35
      Top             =   810
      Width           =   8715
      Begin VB.TextBox Conta 
         Height          =   300
         Left            =   1515
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Agencia 
         Height          =   300
         Left            =   5055
         MaxLength       =   7
         TabIndex        =   7
         Top             =   195
         Width           =   735
      End
      Begin MSMask.MaskEdBox Banco 
         Height          =   300
         Left            =   1515
         TabIndex        =   5
         Top             =   195
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   5070
         TabIndex        =   11
         Top             =   593
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   7215
         TabIndex        =   13
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDownDeposito 
         Height          =   300
         Left            =   6165
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1035
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDeposito 
         Height          =   300
         Left            =   5070
         TabIndex        =   16
         Top             =   1035
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   2610
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   1515
         TabIndex        =   14
         Top             =   1050
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   690
         TabIndex        =   42
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Depositar em:"
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
         Left            =   3795
         TabIndex        =   15
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label19 
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
         Left            =   6615
         TabIndex        =   12
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
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
         Left            =   885
         TabIndex        =   8
         Top             =   652
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
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
         Left            =   4215
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LabelNumero 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4245
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   652
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6660
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ChequePreOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ChequePreOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ChequePreOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ChequePreOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4290
      TabIndex        =   3
      Top             =   270
      Width           =   2190
   End
   Begin VB.CheckBox Selecionada 
      Height          =   225
      Left            =   450
      TabIndex        =   18
      Top             =   2565
      Width           =   675
   End
   Begin MSMask.MaskEdBox ValorParcela 
      Height          =   225
      Left            =   2550
      TabIndex        =   29
      Top             =   4380
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSMask.MaskEdBox Parcela 
      Height          =   225
      Left            =   3045
      TabIndex        =   21
      Top             =   2535
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumTitulo 
      Height          =   225
      Left            =   2295
      TabIndex        =   20
      Top             =   2520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SiglaDocumento 
      Height          =   225
      Left            =   1005
      TabIndex        =   27
      Top             =   4410
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
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
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataVencimento 
      Height          =   225
      Left            =   1530
      TabIndex        =   28
      Top             =   4425
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   900
      TabIndex        =   1
      Top             =   270
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid GridParcelas 
      Height          =   2430
      Left            =   75
      TabIndex        =   17
      Top             =   2565
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   4286
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label TotalRecebido 
      AutoSize        =   -1  'True
      Caption         =   "Total Recebido:"
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
      TabIndex        =   39
      Top             =   5190
      Width           =   1380
   End
   Begin VB.Label ValorRecebidoTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2040
      TabIndex        =   38
      Top             =   5115
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Parcelas"
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
      Left            =   120
      TabIndex        =   36
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3735
      TabIndex        =   2
      Top             =   330
      Width           =   465
   End
   Begin VB.Label LabelCliente 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   330
      Width           =   660
   End
End
Attribute VB_Name = "ChequePreOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim gcolInfoParcRec As Collection
Dim giTrazendoCheque As Integer

Dim gsBancoAnt As String
Dim gsAgenciaAnt As String
Dim gsContaAnt As String

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCheque As AdmEvento
Attribute objEventoCheque.VB_VarHelpID = -1

Dim objGrid1 As AdmGrid
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_NumTitulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_ValorReceber_Col As Integer
Dim iGrid_Juros_Col As Integer
Dim iGrid_Multa_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_SaldoParcela_Col As Integer
Dim iGrid_Selecionada_Col As Integer
Dim iGrid_VencimentoReal_Col As Integer
Dim iGrid_TaxaJuros_Col As Integer
Dim iGrid_FilialEmpresa_Col As Integer 'Inserido por Wagner

Function Trata_Parametros(Optional objChequePre As ClassChequePre) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um cheque selecionado, exibir seus dados
    If Not (objChequePre Is Nothing) Then

        'Verifica se existe
        lErro = CF("ChequePre_Le", objChequePre)
        If lErro <> SUCESSO And lErro <> 17642 Then Error 17643

        'Se encontrou o cheque pre em questão
        If lErro = SUCESSO Then

            lErro = Traz_ChequePre_Tela(objChequePre)
            If lErro <> SUCESSO Then Error 17645

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 17643, 17645

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144516)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Function GridParcelas_Preenche(colInfoParcRec As Collection) As Long
'Traz os dados para o grid

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer, dtDataBaixa As Date
Dim iLinhasExistentes As Integer
Dim objInfoParcRec As New ClassInfoParcRec
Dim objFilialEmpresa As New AdmFiliais 'Inserido por Wagner

On Error GoTo Erro_GridParcelas_Preenche

    dtDataBaixa = MaskedParaDate(DataDeposito)

    iLinha = 0

    'Recoloca o Numero de Linhas caso seja maior que o Numero de Linhas Visiveis na Tela que é 6
    If (gcolInfoParcRec.Count + 1) > NUM_MAXIMO_PARCELAS Then
        GridParcelas.Rows = gcolInfoParcRec.Count + 1
    Else
        GridParcelas.Rows = NUM_MAXIMO_PARCELAS + 1
    End If

    'Renicializa
    Call Grid_Inicializa(objGrid1)

    For Each objInfoParcRec In colInfoParcRec

        iLinha = iLinha + 1

        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objInfoParcRec.dtVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VencimentoReal_Col) = Format(objInfoParcRec.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcRec.sSiglaDocumento
        GridParcelas.TextMatrix(iLinha, iGrid_NumTitulo_Col) = objInfoParcRec.lNumTitulo
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcRec.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objInfoParcRec.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_Selecionada_Col) = objInfoParcRec.iMarcada
        GridParcelas.TextMatrix(iLinha, iGrid_SaldoParcela_Col) = Format(objInfoParcRec.dSaldoParcela, "Standard")

        GridParcelas.TextMatrix(iLinha, iGrid_ValorReceber_Col) = IIf(objInfoParcRec.dValorReceber <> 0, Format(objInfoParcRec.dValorReceber, "Standard"), "")
        GridParcelas.TextMatrix(iLinha, iGrid_Juros_Col) = IIf(objInfoParcRec.dValorJuros <> 0, Format(objInfoParcRec.dValorJuros, "Standard"), "")
        GridParcelas.TextMatrix(iLinha, iGrid_Multa_Col) = IIf(objInfoParcRec.dValorMulta <> 0, Format(objInfoParcRec.dValorMulta, "Standard"), "")
        GridParcelas.TextMatrix(iLinha, iGrid_Desconto_Col) = IIf(objInfoParcRec.dValorDesconto, Format(objInfoParcRec.dValorDesconto, "Standard"), "")

        '#############################################
        'Inserido por Wagner
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objInfoParcRec.iFilialEmpresa
        
        'le o Nome da Filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 131590
                
        If giTipoVersao = VERSAO_FULL Then
            GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        End If
        '#############################################
        
        Call Inicializa_Taxa_Juros(objInfoParcRec, iLinha, dtDataBaixa)

    Next

    objGrid1.iLinhasExistentes = iLinha

    Call Grid_Refresh_Checkbox(objGrid1)
    Call Soma_Valor(objGrid1)

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 131590 'Inserido por Wagner

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144517)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Campos_ChequePre()

    Numero.PromptInclude = False
    Numero.Text = ""
    Numero.PromptInclude = True

    Valor.Text = ""

    'Limpa o campo ValorRecebidoTotal não limpo em Limpa_Tela
    ValorRecebidoTotal.Caption = ""

    iAlterado = 0

End Sub

Private Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO
    Call Trata_Conta
End Sub

Private Sub Banco_Change()

    iAlterado = REGISTRO_ALTERADO
    Call Trata_Conta
End Sub

Private Sub Banco_GotFocus()

    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)

End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Banco_Validate

    'Verifica se foi preenchido o campo Banco
    If Len(Trim(Banco.Text)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(Banco.Text)
    If lErro <> SUCESSO Then Error 35981

    Exit Sub

Erro_Banco_Validate:

    Cancel = True


    Select Case Err

        Case 35981

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144518)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro  As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim objChequePre As New ClassChequePre
Dim sCliente As String

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se os campos essenciais da tela foram preenchidos
    If Len(Trim(Cliente.Text)) = 0 Then gError 35982
    If Len(Trim(Filial.Text)) = 0 Then gError 35983
    If Len(Trim(Banco.Text)) = 0 Then gError 35984
    If Len(Trim(Agencia.Text)) = 0 Then gError 35985
    If Len(Trim(Conta.Text)) = 0 Then gError 35986
    If Len(Trim(Numero.Text)) = 0 Then gError 35987

    sCliente = Cliente.Text

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)

    'Pesquisa se existe filial com o codigo extraido
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 35988

    If lErro = 17660 Then gError 35989

    objChequePre.iFilialEmpresa = giFilialEmpresa
    objChequePre.lCliente = objFilialCliente.lCodCliente
    objChequePre.iFilial = objFilialCliente.iCodFilial
    objChequePre.iBanco = CInt(Banco.Text)
    objChequePre.sAgencia = Agencia.Text
    objChequePre.sContaCorrente = Conta.Text
    objChequePre.lNumero = CLng(Numero.Text)

    'Pede confirmacao da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CHEQUEPRE", objChequePre.lCliente, objChequePre.iFilial, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

    If vbMsgRes = vbYes Then

        GL_objMDIForm.MousePointer = vbHourglass

        'Chama a rotina de exclusao
        lErro = CF("ChequePre_Exclui", objChequePre)
        If lErro <> SUCESSO Then gError 39372

        Call Limpa_Tela_ChequePre

        GL_objMDIForm.MousePointer = vbDefault

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 35982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 35983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 35984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 35985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 35986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 35987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 39372

        Case 35988, 35989

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144519)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'alteracao por tulio281002
    'para fazer com q, qdo a filial do cliente for inexistente
    'ele desvie o fluxo, nao gravando assim o cheque...
    lErro = Trata_FilialCliente(True)
    If lErro <> SUCESSO Then Error 17699

    'Chama a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 17672

    Call Limpa_Tela_Campos_ChequePre

    'chama essa funcao somente para dar refresh nas parcelas...
    'tulio280103
    Call Trata_FilialCliente

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 17672, 17699

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144520)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se algum campo da tela foi modificado
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17782

    'Limpa a tela
    Call Limpa_Tela_ChequePre

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17782

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144521)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = 1

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then Error 17646

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 17647

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            If giTrazendoCheque = 0 Then

                'Seleciona filial na Combo Filial
                If iCodFilial = FILIAL_MATRIZ Then
                    Filial.ListIndex = -1 'Alterado por Wagner
                Else
                    Call CF("Filial_Seleciona", Filial, iCodFilial)
                End If

            End If

        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case Err

        Case 17646

        Case 17647

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144522)

    End Select

    Exit Sub

End Sub

Private Sub Conta_Change()

    iAlterado = REGISTRO_ALTERADO
    Call Trata_Conta
End Sub

Private Sub DataDeposito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDeposito_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDeposito, iAlterado)

End Sub

Private Sub DataDeposito_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDeposito_Validate

    'Verifica se a data de depósito está preenchida
    If Len(Trim(DataDeposito.ClipText)) = 0 Then Exit Sub

    'Verifica se a data final é válida
    lErro = Data_Critica(DataDeposito.Text)
    If lErro <> SUCESSO Then Error 17648

    Exit Sub

Erro_DataDeposito_Validate:

    Cancel = True

    Select Case Err

        Case 17648

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144523)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Verifica se a data emissao é válida
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 126235

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 126235

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144524)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimentoReal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Click

    'Verifica se é uma filial selecionada
    If Filial.ListIndex <> -1 Then

        Call Trata_FilialCliente

    End If

    Exit Sub

Erro_Filial_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144525)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    If Filial.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17655

    'Se não encontrou, extrai o código
    If lErro = 6730 Then

        Call Trata_FilialCliente

    End If

    'Se não encontrou --> Erro
    If lErro = 6731 Then Error 17663

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err

        Case 17655

        Case 17663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144526)

    End Select

    Exit Sub

End Sub

Private Function Trata_FilialCliente(Optional bEhGravacao As Boolean = False) As Long

Dim lErro As Long
Dim sCliente As String
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Trata_FilialCliente

    If giTrazendoCheque = 0 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 17783

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)

        'Pesquisa se existe filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 17784

        'Se não encontrou a Filial --> Erro
        If lErro = 17660 Then Error 17785

        'tulio240103
        'se nao for gravacao
        If bEhGravacao = False Then

            'Carregar as parcelas deste cliente filial não associadas a cheques pre para o grid
            Set gcolInfoParcRec = New Collection

            lErro = CF("ParcelasRec_Le_SemChequePre", objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, gcolInfoParcRec)
            If lErro <> SUCESSO Then Error 17786

            'Limpa o Grid de Parcelas
            Call Grid_Limpa(objGrid1)

            If gcolInfoParcRec.Count > 0 Then
                'Preenche o Grid de Parcelas
                lErro = GridParcelas_Preenche(gcolInfoParcRec)
                If lErro <> SUCESSO Then Error 17791
            End If

        End If

    End If

    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = Err

    Select Case Err

       Case 17783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CHQPRE_NAO_PREENCHIDO", Err)
            Cliente.SetFocus

       Case 17784
            Filial.SetFocus

        Case 17785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE1", Err, objFilialCliente.iCodFilial, sCliente)
            Filial.SetFocus

        Case 17786, 17791

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144527)

    End Select

    Exit Function

End Function

Private Function Critica_Parcelas(ByVal objChequePre As ClassChequePre, ByVal colParcelas As Collection) As Long

Dim lErro As Long
Dim sCliente As String
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult
Dim iIndex As Integer
Dim objInfoParcRec As ClassInfoParcRec
Dim bIncParc As Boolean
Dim colParcelasBD As New Collection
Dim dValorCheque As Double

On Error GoTo Erro_Critica_Parcelas

    bIncParc = False
    
    sCliente = Cliente.Text

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)

    'Pesquisa se existe filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 17784

    'Se não encontrou a Filial --> Erro
    If lErro = 17660 Then gError 17785

    'Carregar as parcelas deste cliente filial não associadas a cheques pre para o grid
    Set colParcelasBD = New Collection

    lErro = CF("ParcelasRec_Le_SemChequePre", objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, colParcelasBD)
    If lErro <> SUCESSO Then gError 17786

    'Limpa o Grid de Parcelas
    'Call Grid_Limpa(objGrid1)

    'verifica a integridade das parcelas
    For iIndex = 1 To colParcelasBD.Count
        
        'percorre as parcelas passadas como parametro
        For Each objInfoParcRec In colParcelas
        
            'se achou a parcela, verifica o saldo
            'If colParcelasBD.Item(iIndex).iNumParcela = objInfoParcRec.iNumParcela And colParcelasBD.Item(iIndex).lNumTitulo = objInfoParcRec.lNumTitulo And colParcelasBD.Item(iIndex).iFilialEmpresa = objInfoParcRec.iFilialEmpresa Then
            If colParcelasBD.Item(iIndex).lNumIntParc = objInfoParcRec.lNumIntParc Then
            
                dValorCheque = 0
            
                'soma os valores dos cheques associados a parcela passada como parametro e que ainda nao foram depositadas
                lErro = CF("ChequeParcelaRec_Le", objChequePre, objInfoParcRec.lNumIntParc, dValorCheque)
                If lErro <> SUCESSO Then gError 188347
            
                'alterar o num de erro dps
                'sugestao: colocar na mensagem de erro
                'a linha do grid em q a parcela incoerente esta e colocar
                'qual o valor no bd e qual o valor no grid
                'para situar melhor o usuario...
                'obs: esse erro significa que o q esta no grid esta
                'diferente do bd, ou seja, o usuario deve chamar
                'a tela novamente para recarregar o grid com os valores corretos
                '... sabemos que eh uma solucao muito ruim, mas como essa situacao eh muito
                'dificil de ocorrer, ela ficou com esse tratamento tosco mesmo...
                If Abs((objInfoParcRec.dSaldoParcela - dValorCheque) - colParcelasBD.Item(iIndex).dSaldoParcela) > DELTA_VALORMONETARIO Then gError 111761
                
                'If objInfoParcRec.dValorReceber + objInfoParcRec.dValorDesconto - objInfoParcRec.dValorJuros - objInfoParcRec.dValorMulta - objInfoParcRec.dSaldoParcela > DELTA_VALORMONETARIO Then gError 205529
                
                'Call Parcela_Atualiza_Grid(objInfoParcRec, colParcelasBD.Item(iIndex).dSaldoParcela)
                
            End If
        
        Next


    Next

    
    'se tem inconsistencia nas parcelas
    'If bIncParc = True Then gError 99999


    If gcolInfoParcRec.Count > 0 Then
        'Preenche o Grid de Parcelas
        lErro = GridParcelas_Preenche(gcolInfoParcRec)
        If lErro <> SUCESSO Then gError 17791
    End If


    Critica_Parcelas = SUCESSO

    Exit Function

Erro_Critica_Parcelas:

    Critica_Parcelas = gErr

    Select Case gErr

       Case 17783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CHQPRE_NAO_PREENCHIDO", gErr)
            Cliente.SetFocus

       Case 17784
            Filial.SetFocus

        Case 17785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE1", gErr, objFilialCliente.iCodFilial, sCliente)
            Filial.SetFocus

        Case 17786, 17791
        
        Case 111761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCONSISTENCIA_PARCELAS", gErr)
            
        Case 205529
            Call Rotina_Erro(vbOKOnly, "ERRO_INCONSISTENCIA_PARCELAS3", gErr, objInfoParcRec.lNumTitulo, objInfoParcRec.iNumParcela, Format(objInfoParcRec.dValorReceber + objInfoParcRec.dValorDesconto - objInfoParcRec.dValorJuros - objInfoParcRec.dValorMulta, "STANDARD"), Format(objInfoParcRec.dSaldoParcela, "STANDARD"))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144528)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoCheque = New AdmEvento

    Set objGrid1 = New AdmGrid

    lErro = Inicializa_Grid_Parcelas(objGrid1)
    If lErro <> SUCESSO Then Error 17664
    
    '###########################
    'Inserido por Wagner
    FilialEmpresa.left = POSICAO_FORA_TELA
    FilialEmpresa.TabStop = False
    '###########################

    'preecher a data emissão com a data atual
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    'preecher a data crédito com a data atual
    DataDeposito.PromptInclude = False
    DataDeposito.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataDeposito.PromptInclude = True

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 17664

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144529)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCheque = Nothing

    Set gcolInfoParcRec = Nothing

    Set objGrid1 = Nothing

End Sub

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub GridParcelas_EnterCell()

    Call Grid_Entrada_Celula(objGrid1, iAlterado)

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGrid1)

End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGrid1)

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid1)

End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGrid1)

End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGrid1)

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub LabelNumero_Click()

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objChequePre As New ClassChequePre
Dim colSelecao As New Collection

On Error GoTo Erro_NumeroLabel_Click

    'Se Cliente estiver vazio, erro
    If Len(Trim(Cliente.Text)) = 0 Then Error 17952

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 17953

    'Preenche objCliente
    objCliente.sNomeReduzido = Cliente.Text

    'Lê o código pelo Nome Reduzido do Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 43075

    'Não achou o Cliente --> erro
    If lErro = 12348 Then Error 43076

    'Preenche objChequePre
    objChequePre.lCliente = objCliente.lCodigo
    objChequePre.iFilial = Codigo_Extrai(Filial.Text)

    'Adiciona filtros: lCliente e iFilial
    colSelecao.Add objChequePre.lCliente
    colSelecao.Add objChequePre.iFilial

    'Chama Tela ChequePreLista
    Call Chama_Tela("ChequePreLista", colSelecao, objChequePre, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 17952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 17953
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 43075

        Case 43076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144530)

    End Select

    Exit Sub

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    If Len(Trim(Numero.ClipText)) > 0 Then

        If Not IsNumeric(Numero.ClipText) Then Error 17670

        If CLng(Numero) < 1 Then Error 17671

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case Err

        Case 17670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_E_NUMERICO", Err, Numero.Text)

        Case 17671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144531)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(Cancel)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objChequePre As ClassChequePre

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objChequePre = obj1

    'Traz os dados de objChequePre para Tela
    lErro = Traz_ChequePre_Tela(objChequePre)
    If lErro <> SUCESSO Then Error 17954

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 17954

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144532)

    End Select

    Exit Sub

End Sub

Private Sub Selecionada_Click()

Dim iAlterado As Integer
Dim iLinha As Integer
Dim objInfoParcRec As New ClassInfoParcRec
Dim iMarcada As Integer
Dim iIndice As Integer
Dim dValorRecebidoTotal As Double
Dim dSomaValor As Double


    iLinha = GridParcelas.Row

    If iLinha > objGrid1.iLinhasExistentes Then Exit Sub

    iMarcada = StrParaInt(GridParcelas.TextMatrix(iLinha, iGrid_Selecionada_Col))

    'obtem obj correspondente a linha do Grid
    Set objInfoParcRec = gcolInfoParcRec.Item(iLinha)

    'atualiza no Obj se a parcela em questão foi marcada ou desmarcada
    objInfoParcRec.iMarcada = iMarcada

    '#########################################
    'Inserido por Wagner
    'Se não foi preenchido o campo Valor a Receber
    If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorReceber_Col))) = 0 Then
        'Se a parcela está marcada
        If iMarcada = MARCADO Then
            'Valor a Receber assume o valor do Saldo Atual
            GridParcelas.TextMatrix(iLinha, iGrid_ValorReceber_Col) = GridParcelas.TextMatrix(iLinha, iGrid_SaldoParcela_Col)
        
            Call Calcula_Multa_Juros_Desc_Parcela(iLinha)
        End If
    End If
    '#########################################

    Call Grid_Refresh_Checkbox(objGrid1)

    Call Soma_Valor(objGrid1)

End Sub

Private Function Soma_Valor(objGridInt As AdmGrid) As Long
'Atualiza o total recebido (selecionado)

Dim iIndice As Integer
Dim dSomaValor As Double

    dSomaValor = 0

    'Loop no GridParcelas
    For iIndice = 1 To objGridInt.iLinhasExistentes

        If StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorReceber_Col)) > 0 And GridParcelas.TextMatrix(iIndice, iGrid_Selecionada_Col) = 1 Then

            'Acumula Valor em dSomaValor
            dSomaValor = dSomaValor + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorReceber_Col))

        End If

    Next

    'Mostra na tela o Valor Total
    ValorRecebidoTotal.Caption = Format(dSomaValor, "Standard")

    Soma_Valor = SUCESSO

    Exit Function

End Function

Private Sub UpDownDeposito_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDeposito_DownClick

    DataDeposito.SetFocus

    If Len(Trim(DataDeposito.ClipText)) > 0 Then

        sData = DataDeposito.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 17653

        DataDeposito.Text = sData

    End If

    Exit Sub

Erro_UpDownDeposito_DownClick:

    Select Case Err

        Case 17653

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144533)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDeposito_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDeposito_UpClick

    DataDeposito.SetFocus

    If Len(Trim(DataDeposito.ClipText)) > 0 Then

        sData = DataDeposito.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 17654

        DataDeposito.Text = sData

    End If

    Exit Sub

Erro_UpDownDeposito_UpClick:

    Select Case Err

        Case 17654

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144534)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    DataEmissao.SetFocus

    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 126232

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 126232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144535)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    DataEmissao.SetFocus

    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 126233

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 126233

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144536)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

Dim iIndice As Integer

Set objGrid1.objForm = Me

    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Pagar")
    objGridInt.colColuna.Add ("Vcto Real")
    objGridInt.colColuna.Add ("Titulo")
    objGridInt.colColuna.Add ("Parc")
    objGridInt.colColuna.Add ("Saldo Atual")
    objGridInt.colColuna.Add ("Valor Receber")
    objGridInt.colColuna.Add ("Taxa Juros")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor Parc")
    objGridInt.colColuna.Add ("Filial Emp") 'Inserido por Wagner
    
    objGridInt.colCampo.Add (Selecionada.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (NumTitulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (SaldoParcela.Name)
    objGridInt.colCampo.Add (ValorReceber.Name)
    objGridInt.colCampo.Add (TaxaJuros.Name)
    objGridInt.colCampo.Add (Juros.Name)
    objGridInt.colCampo.Add (Multa.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (SiglaDocumento.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name) 'Inserido por Wagner
    
    iGrid_Selecionada_Col = 1
    iGrid_VencimentoReal_Col = 2
    iGrid_NumTitulo_Col = 3
    iGrid_Parcela_Col = 4
    iGrid_SaldoParcela_Col = 5
    iGrid_ValorReceber_Col = 6
    iGrid_TaxaJuros_Col = 7
    iGrid_Juros_Col = 8
    iGrid_Multa_Col = 9
    iGrid_Desconto_Col = 10
    iGrid_Tipo_Col = 11
    iGrid_Vencimento_Col = 12
    iGrid_Valor_Col = 13
    iGrid_FilialEmpresa_Col = 14 'Inserido por Wagner

    objGridInt.objGrid = GridParcelas

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    GridParcelas.ColWidth(0) = 400

    'Não permite exclusão ou inclusao no Grid
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGridInt)

    ValorRecebidoTotal.top = GridParcelas.top + GridParcelas.Height
    ValorRecebidoTotal.left = GridParcelas.left
    For iIndice = 0 To iGrid_ValorReceber_Col - 1
        ValorRecebidoTotal.left = ValorRecebidoTotal.left + GridParcelas.ColWidth(iIndice) + GridParcelas.GridLineWidth
    Next

    TotalRecebido.top = ValorRecebidoTotal.top + (ValorRecebidoTotal.Height / 2) - (TotalRecebido.Height / 2)
    TotalRecebido.left = ValorRecebidoTotal.left - TotalRecebido.Width

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Sub UpDown2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(Valor.ClipText)) = 0 Then Exit Sub

    'critica o valor
    lErro = Valor_Positivo_Critica(Valor.Text)
    If lErro <> SUCESSO Then Error 17667

    'Põe o valor formatado na tela
    Valor.Text = Format(Valor.Text, "Fixed")

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case Err

        Case 17667

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144537)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objChequePre As ClassChequePre, colInfoParcRec As Collection) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iIndice As Integer, objInfoParcRec As ClassInfoParcRec

On Error GoTo Erro_Move_Tela_Memoria

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 17691

    'Não achou o Cliente --> erro
    If lErro = 12348 Then Error 43077

    objChequePre.iFilialEmpresa = giFilialEmpresa

    objChequePre.lCliente = objCliente.lCodigo
    objChequePre.iFilial = Codigo_Extrai(Filial.Text)
    If Len(Trim(Banco.Text)) <> 0 Then objChequePre.iBanco = CInt(Banco.Text)
    objChequePre.sAgencia = Agencia.Text
    objChequePre.sContaCorrente = Conta.Text
    If Len(Trim(Numero.Text)) <> 0 Then objChequePre.lNumero = CLng(Numero.Text)
    If Len(Trim(DataDeposito.ClipText)) <> 0 Then
        objChequePre.dtDataDeposito = CDate(DataDeposito.Text)
    Else
        objChequePre.dtDataDeposito = DATA_NULA
    End If
    
    objChequePre.dtDataEmissao = CDate(DataEmissao.Text)
    
    If Len(Trim(Valor.Text)) Then objChequePre.dValor = CDbl(Valor.Text)

    For iIndice = 1 To objGrid1.iLinhasExistentes

        If GridParcelas.TextMatrix(iIndice, iGrid_Selecionada_Col) = "1" Then

            Set objInfoParcRec = gcolInfoParcRec.Item(iIndice)

            With objInfoParcRec
                .dValorReceber = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorReceber_Col))
                .dValorJuros = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Juros_Col))
                .dValorMulta = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Multa_Col))
                .dValorDesconto = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desconto_Col))
            End With
            
            Call colInfoParcRec.Add(objInfoParcRec)

        End If

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 17691, 17692

        Case 43077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144538)

    End Select

    Exit Function

End Function

Private Function Parcela_Atualiza_Grid(objInfoParcRec As ClassInfoParcRec, Optional dValorAtualizado As Double) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Parcela_Atualiza_Grid

    For iIndice = 1 To objGrid1.iLinhasExistentes

        If StrParaLong(GridParcelas.TextMatrix(iIndice, iGrid_NumTitulo_Col)) = objInfoParcRec.lNumTitulo And StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Parcela_Col)) = objInfoParcRec.iNumParcela Then
            
            GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(dValorAtualizado, "Standard")
        
        End If
        
    Next

    Exit Function

Erro_Parcela_Atualiza_Grid:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144539)

    End Select

    Exit Function

End Function


Private Sub Limpa_Tela_ChequePre()

    Call Limpa_Tela(Me)

    Filial.Clear

    Call Grid_Limpa(objGrid1)

    'Limpa o campo ValorRecebidoTotal não limpo em Limpa_Tela
    ValorRecebidoTotal.Caption = ""

    iAlterado = 0

End Sub

Public Function Gravar_Registro() As Long
'Valida os dados do ChequePre para gravação e grava-o

Dim lErro As Long
Dim objChequePre As New ClassChequePre
Dim objInfoParcRec As New ClassInfoParcRec
Dim dtDataDep As Date
Dim dtDataVencimento As Date
Dim dValor As Double
Dim dSomaParcelas As Double
Dim iTesteData As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer, colInfoParcRec As New Collection
Dim objParcelaReceber As New ClassParcelaReceber

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos essenciais da tela foram preenchidos
    If Len(Trim(Cliente.Text)) = 0 Then gError 17675
    If Len(Trim(Filial.Text)) = 0 Then gError 17676
    If Len(Trim(Banco.Text)) = 0 Then gError 17678
    If Len(Trim(Agencia.Text)) = 0 Then gError 17679
    If Len(Trim(Conta.Text)) = 0 Then gError 17680
    If Len(Trim(Numero.Text)) = 0 Then gError 17681
    If Len(Trim(Valor.Text)) = 0 Then gError 17682
    If Len(Trim(DataDeposito.ClipText)) = 0 Then gError 17683
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 126234

    dValor = CDbl(Valor.Text)

    dtDataDep = CDate(DataDeposito.Text)

    If objGrid1.iLinhasExistentes = 0 Then gError 17730

    dSomaParcelas = 0
    iTesteData = 0

    For iIndice = 1 To objGrid1.iLinhasExistentes

        If GridParcelas.TextMatrix(iIndice, iGrid_Selecionada_Col) = "1" Then

            dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

            If iTesteData = 0 Then

                If dtDataVencimento <> dtDataDep Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATAVENCIMENTO_PARCELA_DIFERENTE_DATADEPOSIT", dtDataVencimento, iIndice, dtDataDep)

                    If vbMsgRes = vbNo Then gError 17734

                    iTesteData = 1

                End If

            End If

        End If

    Next

    'Verifica se o valor do cheque corresponde ao Total Recebido
    If dValor <> CDbl(ValorRecebidoTotal.Caption) Then gError 91307

    'Passa os dados da Tela para objChequePre
    lErro = Move_Tela_Memoria(objChequePre, colInfoParcRec)
    If lErro <> SUCESSO Then gError 17684
    
    'verifica se alguem, de outra maquina,
    'ja nao alterou alguma parcela enquanto
    'a mesma estava na tela...
    lErro = Critica_Parcelas(objChequePre, colInfoParcRec)
    If lErro <> SUCESSO Then gError 111762

    'Rotina encarregada de gravar o cheque pre
    lErro = CF("ChequePre_Grava", objChequePre, colInfoParcRec)
    If lErro <> SUCESSO Then gError 17685

    GL_objMDIForm.MousePointer = vbDefault

    iAlterado = 0

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 17675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 17676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 17678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 17679
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 17680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 17681
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 17682
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_CHQPRE_NAO_PREENCHIDO", gErr)

        Case 17683
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADEPOSITO_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 17684, 17685, 17734, 19142, 111762

        Case 17730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", gErr)

        Case 42424
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_GRID_NAO_SELECIONADA", gErr)

        Case 91307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORCHEQUEDIFERENTETOTALRECEBIDO", gErr)
    
        Case 126234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144540)

    End Select

    Exit Function

End Function

Private Function Traz_ChequePre_Tela(objChequePre As ClassChequePre) As Long
'Traz os dados do ChequePre na Tela

Dim lErro As Long
Dim objInfoParcRec As New ClassInfoParcRec
Dim iLinha As Integer
Dim iIndice As Integer
Dim iLinhasExistentes As Integer, Cancel As Boolean

On Error GoTo Erro_Traz_ChequePre_Tela

    giTrazendoCheque = 1

    'Verifica se existe
    lErro = CF("ChequePre_Le", objChequePre)
    If lErro <> SUCESSO And lErro <> 17642 Then gError 126236

    Cliente.Text = objChequePre.lCliente
    Call Cliente_Validate(Cancel)

    Filial.Text = objChequePre.iFilial
    Call Filial_Validate(bSGECancelDummy)

    Banco.Text = CStr(objChequePre.iBanco)
    Agencia.Text = objChequePre.sAgencia
    Conta.Text = objChequePre.sContaCorrente

    Valor.Text = objChequePre.dValor

    DataDeposito.Text = Format(objChequePre.dtDataDeposito, "dd/mm/yy")
    DataEmissao.Text = Format(objChequePre.dtDataEmissao, "dd/mm/yy")

    Numero.PromptInclude = False
    Numero.Text = CStr(objChequePre.lNumero)
    Numero.PromptInclude = True

    Set gcolInfoParcRec = New Collection

    lErro = CF("ParcelasReceber_Le_ChequePre", objChequePre, gcolInfoParcRec)
    If lErro <> SUCESSO Then gError 126237

    lErro = CF("ParcelasRec_Le_SemChequePre", objChequePre.lCliente, objChequePre.iFilial, gcolInfoParcRec)
    If lErro <> SUCESSO Then gError 126238

    'Limpa o Grid
    Call Grid_Limpa(objGrid1)

    lErro = GridParcelas_Preenche(gcolInfoParcRec)
    If lErro <> SUCESSO Then gError 126239

    iAlterado = 0

    giTrazendoCheque = 0

    Traz_ChequePre_Tela = SUCESSO

    Exit Function

Erro_Traz_ChequePre_Tela:

    giTrazendoCheque = 0

    Traz_ChequePre_Tela = gErr

    Select Case gErr

        Case 126236 To 126239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144541)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case iGrid_NumTitulo_Col

                lErro = Saida_Celula_NumTitulo(objGridInt)
                If lErro <> SUCESSO Then gError 57350

            Case iGrid_Parcela_Col

                lErro = Saida_Celula_Parcela(objGridInt)
                If lErro <> SUCESSO Then gError 57351

            Case iGrid_Selecionada_Col

                lErro = Saida_Celula_Selecionada(objGridInt)
                If lErro <> SUCESSO Then gError 57352

            Case iGrid_Tipo_Col

                lErro = Saida_Celula_SiglaDocumento(objGridInt)
                If lErro <> SUCESSO Then gError 57353

            Case iGrid_Valor_Col

                lErro = Saida_Celula_Valor(objGridInt)
                If lErro <> SUCESSO Then gError 57354

            Case iGrid_ValorReceber_Col

                lErro = Saida_Celula_ValorReceber(objGridInt)
                If lErro <> SUCESSO Then gError 91306

            Case iGrid_SaldoParcela_Col

                lErro = Saida_Celula_SaldoParcela(objGridInt)
                If lErro <> SUCESSO Then gError 91307

            Case iGrid_Juros_Col

                lErro = Saida_Celula_Juros(objGridInt)
                If lErro <> SUCESSO Then gError 91330

            Case iGrid_Multa_Col

                lErro = Saida_Celula_Multa(objGridInt)
                If lErro <> SUCESSO Then gError 91331

            Case iGrid_Desconto_Col

                lErro = Saida_Celula_Desconto(objGridInt)
                If lErro <> SUCESSO Then gError 91332

            Case iGrid_Vencimento_Col

                lErro = Saida_Celula_Vencimento(objGridInt)
                If lErro <> SUCESSO Then gError 57355
            
            Case iGrid_TaxaJuros_Col
            
                lErro = Saida_Celula_TaxaJuros(objGridInt)
                If lErro <> SUCESSO Then gError 125314

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 57356

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 57350 To 57355

        Case 57356
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91306, 91307, 91330, 91331, 91332, 125314

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144542)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Selecionada(objGridInt As AdmGrid) As Long
'faz a critica da celula(checkbox) atualiza do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionada

    Set objGridInt.objControle = Selecionada

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 41223

    Saida_Celula_Selecionada = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionada:

    Saida_Celula_Selecionada = Err

    Select Case Err

        Case 41223
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144543)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NumTitulo(objGridInt As AdmGrid) As Long
'faz a critica da celula NumTitulo do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NumTitulo

    Set objGridInt.objControle = NumTitulo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57357

    Saida_Celula_NumTitulo = SUCESSO

    Exit Function

Erro_Saida_Celula_NumTitulo:

    Saida_Celula_NumTitulo = Err

    Select Case Err

        Case 57357
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144544)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Parcela(objGridInt As AdmGrid) As Long
'faz a critica da celula Parcela do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Parcela

    Set objGridInt.objControle = Parcela

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57358

    Saida_Celula_Parcela = SUCESSO

    Exit Function

Erro_Saida_Celula_Parcela:

    Saida_Celula_Parcela = Err

    Select Case Err

        Case 57358
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144545)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SiglaDocumento(objGridInt As AdmGrid) As Long
'faz a critica da celula Tipo do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SiglaDocumento

    Set objGridInt.objControle = SiglaDocumento

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57359

    Saida_Celula_SiglaDocumento = SUCESSO

    Exit Function

Erro_Saida_Celula_SiglaDocumento:

    Saida_Celula_SiglaDocumento = Err

    Select Case Err

        Case 57359
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144546)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'faz a critica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57360

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 57360
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144547)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorReceber(objGridInt As AdmGrid) As Long
'faz a critica da celula ValorReceber do grid que está deixando de ser a corrente

Dim lErro As Long, dValorAnt As Double

On Error GoTo Erro_Saida_Celula_ValorReceber

    Set objGridInt.objControle = ValorReceber

    dValorAnt = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorReceber_Col))
        
    'Se ValorRecebido estiver preenchido
    If Len(Trim(ValorReceber.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(ValorReceber.Text)
        If lErro <> SUCESSO Then gError 91118

        'Põe o valor formatado na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorReceber_Col) = Format(ValorReceber.Text, "standard")

        If dValorAnt <> StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorReceber_Col)) Then
            Call Calcula_Multa_Juros_Desc_Parcela(GridParcelas.Row)
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91304

    'Chama Soma_Valor
    Call Soma_Valor(objGrid1)

    Saida_Celula_ValorReceber = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorReceber:

    Saida_Celula_ValorReceber = gErr

    Select Case gErr

        Case 91304
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91311, 91312, 91118

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144548)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SaldoParcela(objGridInt As AdmGrid) As Long
'faz a critica da celula SaldoParcela do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SaldoParcela

    Set objGridInt.objControle = SaldoParcela

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91305

    Saida_Celula_SaldoParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_SaldoParcela:

    Saida_Celula_SaldoParcela = gErr

    Select Case gErr

        Case 91305
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144549)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Juros(objGridInt As AdmGrid) As Long
'faz a critica da celula Juros do grid que está deixando de ser a corrente

Dim lErro As Long, dValorAnt As Double

On Error GoTo Erro_Saida_Celula_Juros

    Set objGridInt.objControle = Juros

    dValorAnt = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
    
    'Se o Juros estiver preenchida
    If Len(Trim(Juros.Text)) > 0 Then

        'Critica o Juros
        lErro = Valor_NaoNegativo_Critica(Juros.Text)
        If lErro <> SUCESSO Then gError 91336

        'Põe o Juros formatado na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col) = Format(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col), "standard")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91333

    If dValorAnt <> StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col)) Then
    
        Call Recalcula_TaxaJuros(GridParcelas.Row)
        
    End If
    
    Saida_Celula_Juros = SUCESSO

    Exit Function

Erro_Saida_Celula_Juros:

    Saida_Celula_Juros = gErr

    Select Case gErr

        Case 91333, 91336
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144550)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TaxaJuros(objGridInt As AdmGrid) As Long
'faz a critica da celula Juros do grid que está deixando de ser a corrente

Dim lErro As Long, dTaxaAnterior As Double

On Error GoTo Erro_Saida_Celula_TaxaJuros

    Set objGridInt.objControle = TaxaJuros

    dTaxaAnterior = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TaxaJuros_Col))
        
    'Se o Juros estiver preenchida
    If Len(Trim(TaxaJuros.Text)) > 0 Then

        'Critica o Juros
        lErro = Valor_Positivo_Critica(TaxaJuros.Text)
        If lErro <> SUCESSO Then gError 125315

        'Põe o Juros formatado na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TaxaJuros_Col) = Format(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TaxaJuros_Col), "fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 125316
    
    If dTaxaAnterior <> StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TaxaJuros_Col)) Then
    
        lErro = Calcula_Multa_Juros_Desc_Parcela(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 125317
        
    End If
    
    Saida_Celula_TaxaJuros = SUCESSO

    Exit Function

Erro_Saida_Celula_TaxaJuros:

    Saida_Celula_TaxaJuros = gErr

    Select Case gErr

        Case 125315, 125316, 125317
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144551)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Multa(objGridInt As AdmGrid) As Long
'faz a critica da celula Multa do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Multa

    Set objGridInt.objControle = Multa

    'Se Multa estiver preenchida
    If Len(Trim(Multa.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Multa.Text)
        If lErro <> SUCESSO Then gError 91337

        'Põe o valor formatado na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col) = Format(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), "standard")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91334

    Saida_Celula_Multa = SUCESSO

    Exit Function

Erro_Saida_Celula_Multa:

    Saida_Celula_Multa = gErr

    Select Case gErr

        Case 91334, 91337
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144552)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'faz a critica da celula Desconto do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    'Se desconto estiver preenchida
    If Len(Trim(Desconto.Text)) > 0 Then

        'Critica o desconto
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 91338

        'Põe o valor formatado na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col) = Format(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), "standard")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91335

    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 91335, 91338
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 144553)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vencimento(objGridInt As AdmGrid) As Long
'faz a critica da celula Vencimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Vencimento

    Set objGridInt.objControle = DataVencimento

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57361

    Saida_Celula_Vencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = Err

    Select Case Err

        Case 57361
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144554)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CHEQUE_PRE_DATADO
    Set Form_Load_Ocx = Me
    Caption = "Cheque Pré-Datado"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ChequePre"

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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call LabelNumero_Click
        End If

    End If

End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LabelNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumero, Source, X, Y)
End Sub

Private Sub LabelNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumero, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Desconto_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Juros_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaJuros_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Multa_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Multa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Multa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Multa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Multa
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Juros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub Juros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub Juros_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Juros
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TaxaJuros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub TaxaJuros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub TaxaJuros_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Juros
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorReceber_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub ValorReceber_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub ValorReceber_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = ValorReceber
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Calcula_Multa_Juros_Desc_Parcela(iLinha As Integer) As Long
'Preenche para a parcela da linha do grid passada como parametro os valores para multa, juros e descontos

Dim lErro As Long
Dim objInfoParcRec As ClassInfoParcRec, objParcelaReceber As New ClassParcelaReceber
Dim dtDataBaixa As Date
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim objTituloReceber As New ClassTituloReceber
Dim iDias As Integer
Dim dTaxaJuros As Double

On Error GoTo Erro_Calcula_Multa_Juros_Desc_Parcela

    If gcolInfoParcRec.Count >= iLinha Then

        dtDataBaixa = MaskedParaDate(DataDeposito)

        If dtDataBaixa <> DATA_NULA Then
            
            Set objInfoParcRec = gcolInfoParcRec(iLinha)

            objParcelaReceber.lNumIntDoc = objInfoParcRec.lNumIntParc
            lErro = CF("ParcelaReceber_Le", objParcelaReceber)
            If lErro <> SUCESSO And lErro <> 46477 Then gError 125294
            If lErro <> SUCESSO Then gError 125295

            objTituloReceber.lNumIntDoc = objParcelaReceber.lNumIntTitulo
            
            lErro = CF("TituloReceber_Le", objTituloReceber)
            If lErro <> SUCESSO And lErro <> 26061 Then Error 56718
            If lErro <> SUCESSO Then Error 56719
                
            If dtDataBaixa > objInfoParcRec.dtDataVencimentoReal Then

                dTaxaJuros = StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_TaxaJuros_Col)) / 100
                
                lErro = CF("Calcula_Multa_Juros_Parcela2", StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorReceber_Col)), objTituloReceber.dPercMulta, dTaxaJuros, dtDataBaixa, objInfoParcRec.dtDataVencimentoReal, objInfoParcRec.dtVencimento, dValorMulta, dValorJuros)
                If lErro <> SUCESSO Then gError 125299
                
                GridParcelas.TextMatrix(iLinha, iGrid_Multa_Col) = IIf(dValorMulta <> 0, Format(dValorMulta, "Standard"), "")
                GridParcelas.TextMatrix(iLinha, iGrid_Juros_Col) = IIf(dValorJuros <> 0, Format(dValorJuros, "Standard"), "")
                GridParcelas.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
                
             Else

                GridParcelas.TextMatrix(iLinha, iGrid_Multa_Col) = ""
                GridParcelas.TextMatrix(iLinha, iGrid_Juros_Col) = ""

                lErro = CF("Calcula_Desconto_Parcela", objParcelaReceber, dValorDesconto, dtDataBaixa)
                If lErro <> SUCESSO Then gError 125300

                If dValorDesconto > 0 Then
                    GridParcelas.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dValorDesconto, "Standard")
                Else
                    GridParcelas.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
                End If

            End If

        End If

    End If

    Calcula_Multa_Juros_Desc_Parcela = SUCESSO

    Exit Function

Erro_Calcula_Multa_Juros_Desc_Parcela:

    Calcula_Multa_Juros_Desc_Parcela = gErr

    Select Case gErr

        Case 125294, 125296 To 125300

        Case 125295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", gErr)

        Case 56718
        
        Case 56719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", Err, objTituloReceber.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144555)

    End Select

    Exit Function

End Function

Private Function Inicializa_Taxa_Juros(ByVal objInfoParcRec As ClassInfoParcRec, ByVal iLinha As Integer, ByVal dtDataBaixa As Date) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber, objParcelaReceber As New ClassParcelaReceber

On Error GoTo Erro_Inicializa_Taxa_Juros

    'se a parcela nao está associada ao cheque
    If objInfoParcRec.iMarcada = 0 Then
    
        'obter a taxa de juros do titulo
        objParcelaReceber.lNumIntDoc = objInfoParcRec.lNumIntParc
        lErro = CF("ParcelaReceber_Le", objParcelaReceber)
        If lErro <> SUCESSO And lErro <> 46477 Then gError 125294
        If lErro <> SUCESSO Then gError 125295

        objTituloReceber.lNumIntDoc = objParcelaReceber.lNumIntTitulo
        
        lErro = CF("TituloReceber_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 56718
        If lErro <> SUCESSO Then gError 56719
    
        GridParcelas.TextMatrix(iLinha, iGrid_TaxaJuros_Col) = Format(objTituloReceber.dPercJurosDiario * 3000, "Fixed")
    
    Else
    
        If dtDataBaixa > objInfoParcRec.dtDataVencimentoReal And objInfoParcRec.dSaldoParcela > 0 Then
        
            GridParcelas.TextMatrix(iLinha, iGrid_TaxaJuros_Col) = Format(objInfoParcRec.dValorJuros / (objInfoParcRec.dSaldoParcela * (dtDataBaixa - objInfoParcRec.dtVencimento)) * 3000, "Fixed")
            
        Else
        
            GridParcelas.TextMatrix(iLinha, iGrid_TaxaJuros_Col) = ""
    
        End If
        
    End If
    
    Inicializa_Taxa_Juros = SUCESSO
     
    Exit Function
    
Erro_Inicializa_Taxa_Juros:

    Inicializa_Taxa_Juros = gErr
     
    Select Case gErr
          
        Case 56718, 125294 'Alterado por Wagner - Número errado
        
        Case 125295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", gErr)

        Case 56719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144556)
     
    End Select
     
    Exit Function

End Function

Private Function Recalcula_TaxaJuros(ByVal iLinha As Integer) As Long

Dim lErro As Long, objInfoParcRec As ClassInfoParcRec, dtDataBaixa As Date

On Error GoTo Erro_Recalcula_TaxaJuros

    If gcolInfoParcRec.Count >= iLinha Then

        dtDataBaixa = MaskedParaDate(DataDeposito)

        If dtDataBaixa <> DATA_NULA Then
            
            Set objInfoParcRec = gcolInfoParcRec(iLinha)
    
            If dtDataBaixa > objInfoParcRec.dtDataVencimentoReal Then
            
                GridParcelas.TextMatrix(iLinha, iGrid_TaxaJuros_Col) = Format(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_Juros_Col)) / (objInfoParcRec.dSaldoParcela * (dtDataBaixa - objInfoParcRec.dtVencimento)) * 3000, "Fixed")
                
            End If
        
        End If
        
    End If
    
    Recalcula_TaxaJuros = SUCESSO
     
    Exit Function
    
Erro_Recalcula_TaxaJuros:

    Recalcula_TaxaJuros = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144557)
     
    End Select
     
    Exit Function

End Function

'##############################
'Inserido por Wagner
Private Sub FilialEmpresa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = FilialEmpresa
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134023

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134023

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144558)

    End Select
    
    Exit Sub

End Sub

Private Sub Trata_Conta()

Dim lErro As Long
Dim sBanco As String
Dim sAgencia As String
Dim sConta As String
Dim lNumCheques As Long
Dim vbResult As VbMsgBoxResult
Dim objCheque As New ClassChequePre
Dim colSelecao As New Collection
        
On Error GoTo Erro_Trata_Conta

    If gobjCR.iVerificaChqMesmaConta = MARCADO Then

        Call Formata_String_Numero(Banco.Text, sBanco)
        Call Formata_String_Numero(Agencia.Text, sAgencia)
        Call Formata_String_Numero(Conta.Text, sConta)
        
        If sBanco <> "" And sAgencia <> "" And sConta <> "" Then
        
            If StrParaLong(sBanco) <> StrParaLong(gsBancoAnt) Or StrParaLong(sAgencia) <> StrParaLong(gsAgenciaAnt) Or StrParaLong(sConta) <> StrParaLong(gsContaAnt) Then
            
                lErro = CF("ChequePre_Le_DadosConta", Banco.Text, Agencia.Text, Conta.Text, lNumCheques)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                gsBancoAnt = sBanco
                gsAgenciaAnt = sAgencia
                gsContaAnt = sConta
                
                If lNumCheques > 0 Then
                
                    vbResult = Rotina_Aviso(vbYesNo, "AVISO_CONTA_OUTROS_CHEQUES_CAD", lNumCheques)
                    If vbResult = vbYes Then
                    
                        colSelecao.Add Trim(Banco.Text)
                        colSelecao.Add Trim(Agencia.Text)
                        colSelecao.Add Trim(Conta.Text)
                        colSelecao.Add CHEQUEPRE_LOCALIZACAO_BACKOFFICE

                        Call Chama_Tela("ChequesCRLista", colSelecao, objCheque, objEventoCheque, "Banco = ? AND Agencia = ? AND ContaCorrente = ? AND Localizacao = ?")
                    
                    End If
                
                End If
            
            End If
        
        End If
        
    End If

    Exit Sub

Erro_Trata_Conta:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144558)

    End Select
    
    Exit Sub

End Sub
