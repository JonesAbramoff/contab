VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RecebimentoCarne3 
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   6840
      ScaleHeight     =   1755
      ScaleWidth      =   2235
      TabIndex        =   36
      Top             =   360
      Width           =   2295
      Begin VB.CommandButton BotaoFechar 
         Height          =   585
         Left            =   375
         Picture         =   "RecebimentoCarne3.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Fechar"
         Top             =   945
         Width           =   1545
      End
      Begin VB.CommandButton BotaoOk 
         Caption         =   "OK"
         Height          =   585
         Left            =   390
         Picture         =   "RecebimentoCarne3.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   255
         Width           =   1545
      End
   End
   Begin VB.Frame FramePagamento 
      Caption         =   "Pagamento"
      Height          =   2055
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   6525
      Begin VB.CommandButton BotaoTroco 
         Caption         =   "(F3)  Troco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   26
         Top             =   1440
         Width           =   1710
      End
      Begin VB.Label LabelFaltaValor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4320
         TabIndex        =   34
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label LabelFalta 
         AutoSize        =   -1  'True
         Caption         =   "Falta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   33
         Top             =   382
         Width           =   720
      End
      Begin VB.Label LabelPagoValor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   32
         Top             =   915
         Width           =   1710
      End
      Begin VB.Label LabelPago 
         AutoSize        =   -1  'True
         Caption         =   "Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   345
         TabIndex        =   31
         Top             =   937
         Width           =   705
      End
      Begin VB.Label LabelTrocoValor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4320
         TabIndex        =   30
         Top             =   915
         Width           =   1710
      End
      Begin VB.Label LabelTroco 
         AutoSize        =   -1  'True
         Caption         =   "Troco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3390
         TabIndex        =   29
         Top             =   937
         Width           =   780
      End
      Begin VB.Label LabelTotalValor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label LabelTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   27
         Top             =   382
         Width           =   690
      End
   End
   Begin VB.Frame FrameMeiosPagto 
      Caption         =   "Meios de Pagamento"
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   9135
      Begin VB.TextBox F6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4215
         TabIndex        =   2
         Text            =   "(F6)"
         Top             =   1772
         Width           =   360
      End
      Begin VB.TextBox F10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4215
         TabIndex        =   1
         Text            =   "(F10)"
         Top             =   2445
         Width           =   480
      End
      Begin MSMask.MaskEdBox MaskDinheiro 
         Height          =   345
         Left            =   2040
         TabIndex        =   7
         Top             =   450
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskCheques 
         Height          =   345
         Left            =   2040
         TabIndex        =   8
         Top             =   1095
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskCartaoDebito 
         Height          =   345
         Left            =   2040
         TabIndex        =   9
         Top             =   1755
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskOutros 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   2400
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.TextBox F4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4200
         TabIndex        =   3
         Text            =   "(F4)"
         Top             =   1122
         Width           =   330
      End
      Begin VB.CommandButton BotaoCheques 
         Caption         =   "   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4185
         TabIndex        =   5
         Top             =   1077
         Width           =   2400
      End
      Begin VB.CommandButton BotaoCartaoDebito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4200
         TabIndex        =   4
         Top             =   1727
         Width           =   2400
      End
      Begin VB.CommandButton BotaoOutros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4200
         TabIndex        =   6
         Top             =   2400
         Width           =   2400
      End
      Begin VB.Label LabelSoma 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   3960
         TabIndex        =   35
         Top             =   2430
         Width           =   165
      End
      Begin VB.Label LabelDinheiro 
         AutoSize        =   -1  'True
         Caption         =   "Dinheiro :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   24
         Top             =   465
         Width           =   1170
      End
      Begin VB.Label LabelCheques 
         AutoSize        =   -1  'True
         Caption         =   "Cheques :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   23
         Top             =   1125
         Width           =   1230
      End
      Begin VB.Label LabelCDebito 
         AutoSize        =   -1  'True
         Caption         =   "Cartão Débito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   22
         Top             =   1770
         Width           =   1845
      End
      Begin VB.Label LabelOutros 
         AutoSize        =   -1  'True
         Caption         =   "Outros :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   975
         TabIndex        =   21
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label LabelSoma 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3960
         TabIndex        =   20
         Top             =   1770
         Width           =   165
      End
      Begin VB.Label LabelSoma 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   19
         Top             =   1125
         Width           =   165
      End
      Begin VB.Label LabelIgual 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   6690
         TabIndex        =   18
         Top             =   1770
         Width           =   165
      End
      Begin VB.Label LabelIgual 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6690
         TabIndex        =   17
         Top             =   2430
         Width           =   165
      End
      Begin VB.Label LabelIgual 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6690
         TabIndex        =   16
         Top             =   1125
         Width           =   165
      End
      Begin VB.Label LabelOutrosValor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6915
         TabIndex        =   15
         Top             =   2415
         Width           =   2010
      End
      Begin VB.Label LabelCDebitoValor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6915
         TabIndex        =   14
         Top             =   1755
         Width           =   2010
      End
      Begin VB.Label LabelChequeValor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6915
         TabIndex        =   13
         Top             =   1095
         Width           =   2010
      End
      Begin VB.Label LabelDinheiroValor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6915
         TabIndex        =   12
         Top             =   450
         Width           =   2010
      End
      Begin VB.Label LabelIgual 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   6690
         TabIndex        =   11
         Top             =   465
         Width           =   165
      End
   End
End
Attribute VB_Name = "RecebimentoCarne3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjVenda As ClassVenda
Dim gcolcarneParcelasImpressao As Collection
Dim gdSaldoDinheiroAnterior As Double
Dim gdSaldoChequesAnterior As Double
Dim gdSaldoOutrosAnterior As Double
Dim gdSaldoCartaoDebitoAnterior As Double

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Trata_Parametros(colCarneParcelasImpressao As Collection, objVenda As ClassVenda) As Long
    
Dim objCarneParcelasImpressao As ClassCarneParcelasImpressao
Dim dValor As Double
Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se não existe parcelas --> erro.
    If colCarneParcelasImpressao.Count = 0 Then gError 109656
    
    Set gcolcarneParcelasImpressao = colCarneParcelasImpressao
    
    Set gobjVenda = objVenda
    
    'Para cada parcela...
    For Each objCarneParcelasImpressao In gcolcarneParcelasImpressao
        'Calcula o valor a ser baixado
        dValor = dValor + objCarneParcelasImpressao.dParcelaValor - objCarneParcelasImpressao.dDesconto + objCarneParcelasImpressao.dMulta + objCarneParcelasImpressao.dJuros
    Next
    
    'Joga o valor na tela
    LabelTotalValor.Caption = Format(dValor, "standard")
    
    Call Recalcula_Valores
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 10656
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAS_NAO_EXISTENTES", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166251)

    End Select
    
    Exit Function

End Function

Public Sub Form_Load()

On Error GoTo Erro_Form_Load
    
    Set gobjVenda = New ClassVenda
    Set gcolcarneParcelasImpressao = New Collection
    
    'Inicializa os valores com zero
    MaskDinheiro.Text = Format(0, "standard")
    MaskCheques.Text = Format(0, "standard")
    MaskCartaoDebito.Text = Format(0, "standard")
    MaskOutros.Text = Format(0, "standard")
    
    LabelDinheiroValor.Caption = Format(0, "standard")
    LabelChequeValor.Caption = Format(0, "standard")
    LabelCDebitoValor.Caption = Format(0, "standard")
    LabelOutrosValor.Caption = Format(0, "standard")
    LabelTrocoValor.Caption = Format(0, "standard")
                
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166252)

    End Select
    
    Exit Sub

End Sub

Sub Recalcula_Valores()

    'Recalcula quanto já foi pago
    LabelPagoValor.Caption = Format(StrParaDbl(LabelDinheiroValor.Caption) + StrParaDbl(LabelChequeValor.Caption) + StrParaDbl(LabelCDebitoValor.Caption) + StrParaDbl(LabelOutrosValor.Caption), "standard")
            
    'Recalcula o valor que falta
    LabelFaltaValor.Caption = Format(StrParaDbl(LabelTotalValor.Caption) - StrParaDbl(LabelPagoValor.Caption), "standard")
    
    LabelTrocoValor.Caption = Format(0, "standard")
    
    'SE esta valor for negativo...
    If StrParaDbl(LabelFaltaValor.Caption) < 0 Then
        'zera o falta
        LabelFaltaValor.Caption = Format(0, "standard")
        'calcula o troco
        LabelTrocoValor.Caption = Format(StrParaDbl(LabelPagoValor.Caption) - StrParaDbl(LabelTotalValor.Caption), "Standard")
    End If
    
    Exit Sub
        
End Sub

Private Sub Inclui_Movimento(dValor As Double, iTipo As Integer)

Dim objMovimento As New ClassMovimentoCaixa
Dim bAchou As Boolean
Dim iIndice As Integer
    
    bAchou = False
    
    'Para cada movimento da tela
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o Movimento
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        'Se o movimento for do tipo que foi passado e não especificado
        If objMovimento.iTipo = iTipo And objMovimento.iAdmMeioPagto = 0 Then
            'Se o valor a atribuir do movimento for positivo
            If dValor > 0 Then
                'Atribui o novo valor ao movimento
                objMovimento.dValor = dValor
            'Senão
            Else
                'remove o movimento
                gobjVenda.colMovimentosCaixa.Remove (iIndice)
            End If
            bAchou = True
            Exit For
        End If
    Next

    'Se tiver valor a tribuir e o movimento não foi encontrado
    If Not (bAchou) And dValor > 0 Then
        
        'Cria um novo movimento
        Set objMovimento = New ClassMovimentoCaixa
        
        'Preenche o novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = CODIGO_CAIXA_CENTRAL
        objMovimento.iTipo = iTipo
        objMovimento.iParcelamento = COD_A_VISTA
        objMovimento.dHora = CDbl(Time)
        objMovimento.dtDataMovimento = gdtDataHoje
        objMovimento.dValor = dValor
                
        'Adiciona o novo movimento à coleção global da tela
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    End If
    
    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Set gcolcarneParcelasImpressao = New Collection
    
    giRetornoTela = vbCancel
    
    Unload Me
    
End Sub

Private Sub BotaoOk_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoOk_Click

    'Se o valor é insuficiente para pagar
    If StrParaDbl(LabelFaltaValor.Caption) > 0 Then gError 109677
    
    'Calcula o troca da tela
    If StrParaDbl(LabelTrocoValor.Caption) > 0 Then Call Calcula_Troco
        
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub
        
Erro_BotaoOk_Click:
    
    Select Case gErr
    
        Case 109677
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_INSUFICIENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166253)

    End Select
        
    Exit Sub
    
End Sub

Private Sub MaskDinheiro_LostFocus()

    If MaskDinheiro.Text = "" Then MaskDinheiro.Text = "0,00"
    
End Sub

Private Sub MaskDinheiro_GotFocus()
    
    If MaskDinheiro.Text = "0,00" Then MaskDinheiro.Text = ""
        
    If Right(MaskDinheiro.Text, 3) = ",00" Then MaskDinheiro.Text = Format(MaskDinheiro.Text, "#,#")
    
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(MaskDinheiro)
    
    'Guarda o valor presente no campo dinheiro
    gdSaldoDinheiroAnterior = StrParaDbl(MaskDinheiro.Text)
    
End Sub

Private Sub MaskDinheiro_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_MaskDinheiro_Validate
    
    'Se o vaor em dinheiro estiver preenchido
    If Len(Trim(MaskDinheiro.Text)) > 0 Then
        'Verifica se é válido
        lErro = Valor_NaoNegativo_Critica(MaskDinheiro.Text)
        If lErro <> SUCESSO Then gError 99622
    End If
      
    MaskDinheiro.Text = Format(MaskDinheiro.Text, "fixed")
    'Se o valor informado é diferente do que estava anteriormente
    If gdSaldoDinheiroAnterior = StrParaDbl(MaskDinheiro.Text) Then Exit Sub
    
    'COloca o valor formatado na tela
    LabelDinheiroValor.Caption = Format(MaskDinheiro.Text, "fixed")
    
    'Atualiza o movimenot referente ao pagamento em dinheiro
    Call Inclui_Movimento(StrParaDbl(MaskDinheiro.Text), MOVIMENTOCAIXA_RECEB_CARNE_DINHEIRO)

    'recalcula os totais levando em conta o novo valor de pagamento em dinheiro
    Call Recalcula_Valores
    
    Exit Sub
    
Erro_MaskDinheiro_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99622
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 166254)

    End Select

    Exit Sub
    
End Sub
Private Sub MaskCheques_LostFocus()

    If MaskCheques.Text = "" Then MaskCheques.Text = "0,00"
    
End Sub

Private Sub MaskCheques_GotFocus()
    
    If MaskCheques.Text = "0,00" Then MaskCheques.Text = ""
        
    If Right(MaskCheques.Text, 3) = ",00" Then MaskCheques.Text = Format(MaskCheques.Text, "#,#")
    
    'Posiciona o cursor na frente
    Call MaskEdBox_TrataGotFocus(MaskCheques)
    
    'Guarda o valor presente no campo Cheque
    gdSaldoChequesAnterior = StrParaDbl(MaskCheques.Text)
    
End Sub

Private Sub MaskCheques_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim bAchou As Boolean
Dim objCheque As New ClassChequePre
Dim objMovCaixa As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim lNum As Long

On Error GoTo Erro_MaskCheques_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskCheques.Text)) > 0 Then
        'Verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCheques.Text)
        If lErro <> SUCESSO Then gError 99623
        
    End If
        
    MaskCheques.Text = Format(MaskCheques.Text, "fixed")
    
    'Se o valor nesse campo nao foi alterado ==> Sai
    If gdSaldoChequesAnterior = StrParaDbl(MaskCheques.Text) Then Exit Sub
    
    'Exibe o valor formatado na tela
    LabelChequeValor.Caption = Format(StrParaDbl(BotaoCheques.Caption) + StrParaDbl(MaskCheques.Text), "fixed")
    
    'Para cada cheque
    For iIndice = gobjVenda.colCheques.Count To 1 Step -1
        'Pega o cheque
        Set objCheque = gobjVenda.colCheques.Item(iIndice)
        'Se ele for não especificado
        If objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO Then
            'Se o valor de cheque nao especificado for positivo
            If StrParaDbl(MaskCheques.Text) > 0 Then
                'Guarda o valor e numero do Cheque
                objCheque.dValor = StrParaDbl(MaskCheques.Text)
                lNum = objCheque.lNumero
            Else
                'remove o cheque nao especificado
                gobjVenda.colCheques.Remove (iIndice)
            End If
            bAchou = True
        End If
    Next
    
    'Se achou o chque nao especificado
    If bAchou Then
        'Para cada movimento
        For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
            'Pega o movimento
            Set objMovCaixa = gobjVenda.colMovimentosCaixa.Item(iIndice)
            'Se for o cheque
            If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARNE_CHEQUE And objMovCaixa.lNumMovto = lNum Then
                'Se há valor em cheque nao especificado
                If StrParaDbl(MaskCheques.Text) > 0 Then
                    'Atualiza o Valor do moivmento
                    objMovCaixa.dValor = StrParaDbl(MaskCheques.Text)
                Else
                    'Retira o movimento
                    gobjVenda.colMovimentosCaixa.Remove (iIndice)
                End If
            End If
        Next
    
    'Se não achou
    Else
        'Se há valor de cheque a incluir
        If StrParaDbl(MaskCheques.Text) > 0 Then
            'Cria um novo cheque
            Set objCheque = New ClassChequePre
            'Preenche os dados defaults do cheque
            objCheque.dtDataDeposito = gdtDataHoje
            objCheque.dValor = StrParaDbl(MaskCheques.Text)
            objCheque.iFilialEmpresa = giFilialEmpresa
            objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO
                        
            'Adiciona o cheque  na coleção da venda
            gobjVenda.colCheques.Add objCheque
                    
            'criar movimento para o cheque
            Set objMovCaixa = New ClassMovimentoCaixa
        
            'Preenche o novo movcaixa
            objMovCaixa.iFilialEmpresa = giFilialEmpresa
            objMovCaixa.iCaixa = giCodCaixa
            objMovCaixa.iCodOperador = giCodOperador
            objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARNE_CHEQUE
            objMovCaixa.iParcelamento = COD_A_VISTA
            objMovCaixa.dtDataMovimento = gdtDataHoje
            objMovCaixa.dValor = StrParaDbl(MaskCheques.Text)
            objMovCaixa.dHora = CDbl(Time)
            objMovCaixa.lNumMovto = gobjVenda.colCheques.Count
            objMovCaixa.lNumRefInterna = gobjVenda.colCheques.Count
            
            'Adiciona o movimento a coleção de moivmewntos da venda
            gobjVenda.colMovimentosCaixa.Add objMovCaixa
        End If
    End If
    
    Exit Sub
    
    'Se o valor nesse campo nao foi alterado ==> Sai
    If gdSaldoChequesAnterior = StrParaDbl(MaskCheques.Text) Then Exit Sub
    
    bAchou = False
    
    'Se o valor nesse campo nao foi alterado ==> Sai
    If gdSaldoChequesAnterior = StrParaDbl(MaskCheques.Text) Then Exit Sub
    
    'Para cada cheque
    For iIndice = gobjVenda.colCheques.Count To 1 Step -1
        'Pega o cheque
        Set objCheque = gobjVenda.colCheques.Item(iIndice)
        'Se ele for não especificado
        If objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO Then
            'Se o valor de cheque nao especificado for positivo
            If StrParaDbl(MaskCheques.Text) > 0 Then
                'Guarda o valor e numero do Cheque
                objCheque.dValor = StrParaDbl(MaskCheques.Text)
                lNum = objCheque.lNumero
            Else
                'remove o cheque nao especificado
                gobjVenda.colCheques.Remove (iIndice)
            End If
            bAchou = True
        End If
    Next
    
    'Se achou o chque nao especificado
    If bAchou Then
        'Para cada movimento
        For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
            'Pega o movimento
            Set objMovCaixa = gobjVenda.colMovimentosCaixa.Item(iIndice)
            'Se for o cheque
            If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARNE_CHEQUE And objMovCaixa.lNumMovto = lNum Then
                'Se há valor em cheque nao especificado
                If StrParaDbl(MaskCheques.Text) > 0 Then
                    'Atualiza o Valor do moivmento
                    objMovCaixa.dValor = StrParaDbl(MaskCheques.Text)
                Else
                    'Retira o movimento
                    gobjVenda.colMovimentosCaixa.Remove (iIndice)
                End If
            End If
        Next
    
    'Se não achou
    Else
        'Se há valor de cheque a incluir
        If StrParaDbl(MaskCheques.Text) > 0 Then
            'Cria um novo cheque
            Set objCheque = New ClassChequePre
            'Preenche os dados defaults do cheque
            objCheque.dtDataDeposito = gdtDataHoje
            objCheque.dValor = StrParaDbl(MaskCheques.Text)
            objCheque.iFilialEmpresa = giFilialEmpresa
            objCheque.iNaoEspecificado = CHEQUE_NAO_ESPECIFICADO
                        
            'Adiciona o cheque  na coleção da venda
            gobjVenda.colCheques.Add objCheque
                    
            'criar movimento para o cheque
            Set objMovCaixa = New ClassMovimentoCaixa
        
            'Preenche o novo movcaixa
            objMovCaixa.iFilialEmpresa = giFilialEmpresa
            objMovCaixa.iCaixa = giCodCaixa
            objMovCaixa.iCodOperador = giCodOperador
            objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARNE_CHEQUE
            objMovCaixa.iParcelamento = COD_A_VISTA
            objMovCaixa.dtDataMovimento = gdtDataHoje
            objMovCaixa.dValor = StrParaDbl(MaskCheques.Text)
            objMovCaixa.dHora = CDbl(Time)
            objMovCaixa.lNumMovto = gobjVenda.colCheques.Count
            objMovCaixa.lNumRefInterna = gobjVenda.colCheques.Count
            
            'Adiciona o movimento a coleção de moivmewntos da venda
            gobjVenda.colMovimentosCaixa.Add objMovCaixa
        End If
    End If
    
    Exit Sub
    
Erro_MaskCheques_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99623
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 166255)

    End Select

    Exit Sub
    
End Sub

Private Sub MaskCartaoDebito_LostFocus()

    If MaskCartaoDebito.Text = "" Then MaskCartaoDebito.Text = "0,00"
    
End Sub

Private Sub MaskCartaoDebito_GotFocus()
    
    If MaskCartaoDebito.Text = "0,00" Then MaskCartaoDebito.Text = ""
        
    If Right(MaskCartaoDebito.Text, 3) = ",00" Then MaskCartaoDebito.Text = Format(MaskCartaoDebito.Text, "#,#")
    
    'Posiciona o cursor no início
    Call MaskEdBox_TrataGotFocus(MaskCartaoDebito)
    'guarda o valor atual em cartão Débito nao especificado
    gdSaldoCartaoDebitoAnterior = StrParaDbl(MaskCartaoDebito.Text)
    
End Sub

Private Sub MaskCartaoDebito_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskCartaoDebito_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskCartaoDebito.Text)) > 0 Then
        'verifica se é um valor válido
        lErro = Valor_NaoNegativo_Critica(MaskCartaoDebito.Text)
        If lErro <> SUCESSO Then gError 99624
    
    End If
    
    MaskCartaoDebito.Text = Format(MaskCartaoDebito.Text, "fixed")
    
    'Se nao houve alteracao de valor
    If gdSaldoCartaoDebitoAnterior = StrParaDbl(MaskCartaoDebito.Text) Then Exit Sub
    
    'Exibe o valor formatado na tela
    LabelCDebitoValor.Caption = Format(StrParaDbl(BotaoCartaoDebito.Caption) + StrParaDbl(MaskCartaoDebito.Text), "fixed")
    
    'Recalcula os valores
    Call Recalcula_Valores
        
    'Inclui o movimento
    Call Inclui_Movimento(StrParaDbl(MaskCartaoDebito.Text), MOVIMENTOCAIXA_RECEB_CARNE_CARTAODEBITO)
    
    Exit Sub
    
Erro_MaskCartaoDebito_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99624
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 166256)

    End Select

    Exit Sub
    
End Sub


Private Sub MaskOutros_LostFocus()

    If MaskOutros.Text = "" Then MaskOutros.Text = "0,00"
    
End Sub

Private Sub MaskOutros_GotFocus()
    
    If MaskOutros.Text = "0,00" Then MaskOutros.Text = ""
        
    If Right(MaskOutros.Text, 3) = ",00" Then MaskOutros.Text = Format(MaskOutros.Text, "#,#")
    
    'Posiciona o Cursor no Inicio
    Call MaskEdBox_TrataGotFocus(MaskOutros)
    'Guarda o valor que está em outros
    gdSaldoOutrosAnterior = StrParaDbl(MaskOutros.Text)
    
End Sub

Private Sub MaskOutros_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_MaskOutros_Validate
    
    'Se estiver preenchido
    If Len(Trim(MaskOutros.Text)) > 0 Then
        'Verifica se o valor é válido
        lErro = Valor_NaoNegativo_Critica(MaskOutros.Text)
        If lErro <> SUCESSO Then gError 99750
        
    End If
    
    MaskOutros.Text = Format(MaskOutros.Text, "fixed")
    
    'Se o valor não foi alterado ==> Sai
    If gdSaldoOutrosAnterior = StrParaDbl(MaskOutros.Text) Then Exit Sub
    
    'Exibe o valor formatado
    LabelOutrosValor.Caption = Format(StrParaDbl(BotaoOutros.Caption) + StrParaDbl(MaskOutros.Text), "fixed")
        
    'Recalcula os totais
    Call Recalcula_Valores
        
    'Atualiza o Movimento
    Call Inclui_Movimento(StrParaDbl(MaskOutros.Text), MOVIMENTOCAIXA_RECEB_CARNE_OUTROS)
    
    Exit Sub
    
Erro_MaskOutros_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99750
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 166257)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCheques_Click()
    
Dim lErro As Long
Dim objCheques As New ClassChequePre
Dim dTotal As Double
Dim dTotal1 As Double

On Error GoTo Erro_BotaoCheques_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_Tela_Modal("PagamentoCheque", gobjVenda)
        
    'Faz o somatório dos cheques
    For Each objCheques In gobjVenda.colCheques
        If objCheques.iNaoEspecificado = CHEQUE_ESPECIFICADO Then
            'Acumula os especificados
            dTotal = dTotal + objCheques.dValor
        Else
            'Acumula os não especificados
            dTotal1 = dTotal1 + objCheques.dValor
        End If
    Next
        
    'Joga o valor do somatório no botão e na MaskedBox
    BotaoCheques.Caption = Format(dTotal, "Standard")
    If dTotal1 <> 0 Then MaskCheques.Text = Format(dTotal1, "Standard")
    
    'Atualiza o total
    LabelChequeValor.Caption = Format(StrParaDbl(MaskCheques.Text) + StrParaDbl(BotaoCheques.Caption), "Standard")
     
    'Recalcula os Totais
    Call Recalcula_Valores
            
    Exit Sub
    
Erro_BotaoCheques_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166258)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCartaoDebito_Click()
    
Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim dTotal As Double

On Error GoTo Erro_BotaoCartaoDebito_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_Tela_Modal("PagamentoCartao", gobjVenda, MOVIMENTOCAIXA_RECEB_CARNE_CARTAODEBITO)
        
    'Faz o somatório dos CartaoDebito
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_CARNE_CARTAODEBITO And objMovimento.iAdmMeioPagto <> 0 Then dTotal = dTotal + objMovimento.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoCartaoDebito.Caption = Format(dTotal, "Standard")
        
    'Atualiza o total
    LabelCDebitoValor.Caption = Format(StrParaDbl(MaskCartaoDebito.Text) + StrParaDbl(BotaoCartaoDebito.Caption), "Standard")
    
    'Atualiza os totais
    Call Recalcula_Valores
    
    Exit Sub
            
Erro_BotaoCartaoDebito_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166259)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoOutros_Click()
    
Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim dTotal As Double

On Error GoTo Erro_BotaoOutros_Click
    
    'Chama tela de pagamento cheque modal
    Call Chama_Tela_Modal("PagamentoOutros", gobjVenda)
        
    'Faz o somatório dos Outros
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_CARNE_OUTROS And objMovimento.iAdmMeioPagto <> 0 Then dTotal = dTotal + objMovimento.dValor
    Next
    
    'Joga o valor do somatório no botão
    BotaoOutros.Caption = Format(dTotal, "Standard")
        
    'Exibe o valor formatado
    LabelOutrosValor.Caption = Format(StrParaDbl(MaskOutros.Text) + StrParaDbl(BotaoOutros.Caption), "Standard")
    
    'Recalcula os totais
    Call Recalcula_Valores
    
    Exit Sub
            
Erro_BotaoOutros_Click:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166260)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoTroco_Click()
    
Dim lErro As Long
Dim objMovCaixa As New ClassMovimentoCaixa

On Error GoTo Erro_BotaoTroco_Click
    
    'Se não há troco para especificar --> erro.
    If StrParaDbl(LabelTrocoValor.Caption) = 0 Then gError 109671
    
    'Joga o valor do troco no obj
    gobjVenda.objCupomFiscal.dValorTroco = StrParaDbl(LabelTrocoValor.Caption)
    
    'Calcula o  troco
    Call Calcula_Troco
    
    'Chama tela de troco
    Call Chama_Tela_Modal("Troco", gobjVenda)
        
    Exit Sub
        
Erro_BotaoTroco_Click:
    
    Select Case gErr
        
        Case 109671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TROCO_NAO_ESPECIFICADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166261)

    End Select

    Exit Sub
    
End Sub

Private Sub Calcula_Troco()
'Varrer a col de movimentos procurando movimentos de troco (din,carta,c/v)
'Acumula os valores de troco encontrados e se estiver faltando incluir o que falta em um movimento de troco em dinheiro
'Se não encontrar cria um com todo o valor para troco em dinheiro

Dim dTroco As Double
Dim dTroco1 As Double
Dim bAchou As Boolean
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice As Integer

    dTroco = 0
    dTroco1 = 0
    
    'Para cada movimento
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        'Se for do tipo troco em dinheiro
        If objMovimento.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO Then
            'Guarda o valor do movimento em dinheiro
            dTroco1 = objMovimento.dValor
            'Acumula o troco
            dTroco = dTroco + dTroco1
            
            bAchou = True

        'Se for do tipo Ticket
        ElseIf objMovimento.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET Then
            'Acumula o troco
            dTroco = dTroco + objMovimento.dValor
        
        'Se for do tipo ContraVale
        ElseIf objMovimento.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_CONTRAVALE Then
            'Acumula a troco
            dTroco = dTroco + objMovimento.dValor
        End If
    Next
    
    'Se o troco da tela for maior do que o até agora especificado
    If StrParaDbl(LabelTrocoValor.Caption) - dTroco > 0.00001 Then
        'Calcula a diferença
        dTroco = StrParaDbl(LabelTrocoValor.Caption) - dTroco
        
        If bAchou Then
            'Acrescenta a diferenç a o troco em dinheiro
            For Each objMovimento In gobjVenda.colMovimentosCaixa
            'Se for do tipo dinheiro
                If objMovimento.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO Then objMovimento.dValor = objMovimento.dValor + dTroco
            Next
        Else
            'Cria um movimento em dinheiro para a diferença
            Set objMovimento = New ClassMovimentoCaixa
            
            objMovimento.iFilialEmpresa = giFilialEmpresa
            objMovimento.iCaixa = CODIGO_CAIXA_CENTRAL
            objMovimento.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO
            objMovimento.iParcelamento = COD_A_VISTA
            objMovimento.dHora = CDbl(Time)
            objMovimento.dtDataMovimento = gdtDataHoje
            objMovimento.dValor = dTroco
            objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_DINHEIRO
            
            gobjVenda.colMovimentosCaixa.Add objMovimento
        End If
    'se o troco diminuiu
    Else
        'Se o troco em dinheiro for maior ou igual ao novo troco
        If dTroco1 - StrParaDbl(LabelTrocoValor.Caption) > 0.00001 Then
            'Exclui todos os recebimentos em vale e contra-vale
            For iIndice = (gobjVenda.colMovimentosCaixa.Count - 1) To 1 Step -1
                Set gobjVenda = gobjVenda.colMovimentosCaixa.Item(iIndice)
                If gobjVenda.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_CONTRAVALE Or gobjVenda.iTipo = MOVIMENTOCAIXA_TROCO_VALE Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
            Next
            'Joga o valor total do troco no movimento em dinheiro
            dTroco = StrParaDbl(LabelTrocoValor.Caption)
        Else
            'Se for do tipo dinheiro-->update com o valor restante para completar o troco
            dTroco = StrParaDbl(LabelTrocoValor.Caption) - (dTroco - dTroco1)
        End If
        
        For iIndice = (gobjVenda.colMovimentosCaixa.Count - 1) To 1 Step -1
            Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
            If objMovimento.iTipo = MOVIMENTOCAIXA_CARNE_TROCO_DINHEIRO Then objMovimento.dValor = dTroco
        Next
        
    End If
        
    Exit Sub
    
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
   'Clique em F4
    If KeyCode = vbKeyF4 Then
        Call BotaoCheques_Click
    End If
    
    'Clique em F6
    If KeyCode = vbKeyF6 Then
        Call BotaoCartaoDebito_Click
    End If
    
    'Clique em F10
    If KeyCode = vbKeyF10 Then
        Call BotaoOutros_Click
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Recebimento de Carnê - Formas de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RecebimentoCarne3"
    
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

'***** fim do trecho a ser copiado ******
