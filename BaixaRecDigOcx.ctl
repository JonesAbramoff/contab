VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaRecDigOcx 
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   9120
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7212
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   144
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaRecDigOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "BaixaRecDigOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BaixaRecDigOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4620
      Index           =   1
      Left            =   48
      TabIndex        =   0
      Top             =   48
      Width           =   9000
      Begin VB.ComboBox FilialEmpresa 
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         Top             =   1740
         Width           =   1170
      End
      Begin VB.ComboBox ContaCorrente 
         Height          =   288
         Left            =   1512
         TabIndex        =   3
         Top             =   1146
         Width           =   1965
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   2550
         TabIndex        =   5
         Top             =   1695
         Width           =   885
      End
      Begin MSMask.MaskEdBox TotalRecebido 
         Height          =   312
         Left            =   4836
         TabIndex        =   2
         Top             =   612
         Width           =   1632
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox NumDocInf 
         Height          =   312
         Left            =   1512
         TabIndex        =   1
         Top             =   612
         Width           =   792
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Historico 
         Height          =   300
         Left            =   4836
         TabIndex        =   4
         Top             =   1140
         Width           =   4044
         _ExtentX        =   7144
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   225
         Left            =   825
         TabIndex        =   7
         Top             =   2385
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
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
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   225
         Left            =   375
         TabIndex        =   14
         Top             =   3045
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cobrador 
         Height          =   225
         Left            =   5400
         TabIndex        =   18
         Top             =   3390
         Width           =   2160
         _ExtentX        =   3810
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   225
         Left            =   1680
         TabIndex        =   8
         Top             =   2370
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
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
      Begin MSMask.MaskEdBox Saldo 
         Height          =   225
         Left            =   6300
         TabIndex        =   13
         Top             =   2385
         Width           =   1005
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
      Begin MSMask.MaskEdBox ValorDesconto 
         Height          =   225
         Left            =   4290
         TabIndex        =   11
         Top             =   2325
         Width           =   855
         _ExtentX        =   1508
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
      Begin MSMask.MaskEdBox ValorMulta 
         Height          =   225
         Left            =   2460
         TabIndex        =   9
         Top             =   2370
         Width           =   855
         _ExtentX        =   1508
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
      Begin MSMask.MaskEdBox ValorJuros 
         Height          =   225
         Left            =   3375
         TabIndex        =   10
         Top             =   2355
         Width           =   855
         _ExtentX        =   1508
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
      Begin MSMask.MaskEdBox ValorRecebido 
         Height          =   225
         Left            =   5235
         TabIndex        =   12
         Top             =   2325
         Width           =   945
         _ExtentX        =   1667
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
      Begin MSMask.MaskEdBox ValorParcela 
         Height          =   225
         Left            =   1875
         TabIndex        =   15
         Top             =   3060
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox Cliente 
         Height          =   225
         Left            =   3285
         TabIndex        =   16
         Top             =   3015
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   225
         Left            =   4185
         TabIndex        =   17
         Top             =   3015
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   2550
         Left            =   60
         TabIndex        =   19
         Top             =   1590
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   4498
         _Version        =   393216
      End
      Begin MSComCtl2.UpDown UpDownDataBaixa 
         Height          =   300
         Left            =   2688
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   96
         Width           =   216
         _ExtentX        =   370
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixa 
         Height          =   300
         Left            =   1512
         TabIndex        =   31
         Top             =   96
         Width           =   1176
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataCredito 
         Height          =   300
         Left            =   6000
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   96
         Width           =   240
         _ExtentX        =   370
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCredito 
         Height          =   300
         Left            =   4836
         TabIndex        =   33
         Top             =   96
         Width           =   1176
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data Crédito:"
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
         Left            =   3612
         TabIndex        =   35
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data Baixa:"
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
         Left            =   432
         TabIndex        =   34
         Top             =   150
         Width           =   1008
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Left            =   72
         TabIndex        =   29
         Top             =   4284
         Width           =   1428
      End
      Begin VB.Label ContaCorrenteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
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
         Left            =   84
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   1194
         Width           =   1356
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   3372
         TabIndex        =   25
         Top             =   672
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
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
         Left            =   396
         TabIndex        =   26
         Top             =   672
         Width           =   1056
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Histórico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   3924
         TabIndex        =   27
         Top             =   1194
         Width           =   828
      End
      Begin VB.Label TotalRecebidoGrid 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   288
         Left            =   1512
         TabIndex        =   28
         Top             =   4260
         Width           =   1848
      End
   End
End
Attribute VB_Name = "BaixaRecDigOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
    
Dim gcolParcelasRec As ColParcelaReceber
Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim giDataAnterior As Integer
Dim objGridParcelas As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_ValorMulta_Col As Integer
Dim iGrid_ValorJuros_Col As Integer
Dim iGrid_ValorDesconto_Col As Integer
Dim iGrid_ValorRecebido_Col As Integer
Dim iGrid_Saldo_Col As Integer
Dim iGrid_DataVencimento_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Nome_Col As Integer
Dim iGrid_Cobrador_Col As Integer

Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1

Private Const DATA_BAIXA As String = "Data_Baixa"
Private Const DATA_CREDITO As String = "Data_Credito"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"
Private Const NUM_TITULO As String = "Numero_Titulo"
Private Const PARCELA1 As String = "Parcela"
Private Const VALOR_BAIXAR As String = "Valor_Baixar"
Private Const VALOR_DESCONTO As String = "Valor_Desconto"
Private Const VALOR_MULTA As String = "Valor_Multa"
Private Const VALOR_JUROS As String = "Valor_Juros"
Private Const VALOR_RECEBIDO As String = "Valor_Recebido"
Private Const CTACARTEIRACOBRADOR As String = "Cta_CarteiraCobrador"
Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"

'Indicação de número máximo de linhas do Grid
Const NUM_MAX_PARCELAS_BAIXA = 50

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloRec As ClassTituloReceber
Private gobjParcelaRec As ClassParcelaReceber
Private gobjBaixaParcRec As ClassBaixaParcRec
Private gobjBaixaReceber As ClassBaixaReceber

Private gsContaCtaCorrente As String 'conta contabil da conta corrente onde foram depositados os cheques
Private gsContaFilDep As String 'conta contabil da filial recebedora
Private giFilialEmpresaConta As Integer 'filial empresa possuidora da conta corrente utilizada p/o deposito

Private Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 49552

    'Limpa a Tela
    Call Limpa_Tela_BaixaRecDig
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 49552

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143361)

    End Select

End Sub

Private Sub ContaCorrenteLabel_Click()
'chama browse de conta corrente

Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim colSelecao As New Collection

    If Len(Trim(ContaCorrente.Text)) > 0 Then objContasCorrentesInternas.iCodigo = Codigo_Extrai(ContaCorrente.Text)

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContasCorrentesInternas, objEventoContaCorrenteInt)

End Sub

Private Sub DataBaixa_GotFocus()

Dim iDataAux As Integer
    
    iDataAux = giDataAnterior
    Call MaskEdBox_TrataGotFocus(DataBaixa, iAlterado)
    giDataAnterior = iDataAux

End Sub

Private Sub DataCredito_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataCredito, iAlterado)

End Sub

Private Sub NumDocInf_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumDocInf, iAlterado)

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1
    
    ContaCorrente.Text = objContaCorrenteInt.iCodigo
    Call ContaCorrente_Validate(bSGECancelDummy)
    
    Me.Show

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
      
    Set objEventoContaCorrenteInt = New AdmEvento
    
    iFrameAtual = 1
    
    'Carrega list de ComboBox ContaCorrente
    lErro = Carrega_ContaCorrente()
    If lErro <> SUCESSO Then Error 46416
    
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then Error 46420
    
    'Carrega a combo de FiliaisEmpresa
    lErro = Carrega_FilialEmpresa()
    If lErro <> SUCESSO Then Error 20802
    
    Set objGridParcelas = New AdmGrid
    
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 46421
    
    Set gcolParcelasRec = New ColParcelaReceber
    
    DataBaixa.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCredito.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    iAlterado = 0
    giDataAnterior = 0
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    Select Case Err
    
        Case 20802, 46416, 46420, 46421, 49547
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143362)
            
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_ContaCorrente() As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_Carrega_ContaCorrente

    'Lê Codigos, NomesReduzidos de ContasCorrentes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 46418

    'Preeche list de ComboBox
    For Each objCodigoNomeRed In colCodigoNomeRed
        ContaCorrente.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    Carrega_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_ContaCorrente:

    Carrega_ContaCorrente = Err

    Select Case Err

        Case 46418  'já tratado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143363)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

Dim iIndice As Integer

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor Parcela")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Nome Cliente")
    objGridInt.colColuna.Add ("Cobrador")
   
   'campos de edição do grid
   
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Saldo.Name)
    objGridInt.colCampo.Add (ValorMulta.Name)
    objGridInt.colCampo.Add (ValorJuros.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (ValorRecebido.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (Nome.Name)
    objGridInt.colCampo.Add (Cobrador.Name)

    iGrid_Tipo_Col = 1
    iGrid_FilialEmpresa_Col = 2
    iGrid_Numero_Col = 3
    iGrid_Parcela_Col = 4
    iGrid_Saldo_Col = 5
    iGrid_ValorMulta_Col = 6
    iGrid_ValorJuros_Col = 7
    iGrid_ValorDesconto_Col = 8
    iGrid_ValorRecebido_Col = 9
    iGrid_DataVencimento_Col = 10
    iGrid_ValorParcela_Col = 11
    iGrid_Cliente_Col = 12
    iGrid_Nome_Col = 13
    iGrid_Cobrador_Col = 14

    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PARCELAS_BAIXA + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6

    'largura da primeira coluna
    GridParcelas.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
       
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO
    
    Exit Function

End Function

Private Sub DataBaixa_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giDataAnterior = REGISTRO_ALTERADO

End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixa_Validate

    If giDataAnterior <> REGISTRO_ALTERADO Then Exit Sub
    
    'Se a DataBaixa está preenchida
    If Len(DataBaixa.ClipText) > 0 Then

        'Verifica se a DataBaixa é válida
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then Error 46422

        lErro = Calcula_Multa_Juros_Desc_Parcelas
        If lErro <> SUCESSO Then Error 46506
    
        DataCredito.Text = Format(DataBaixa, "dd/mm/yy")
    
    End If
        
    giDataAnterior = 0
    
    Exit Sub

Erro_DataBaixa_Validate:

    Cancel = True


    Select Case Err

        Case 46422

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143364)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjContabAutomatica = Nothing
    Set gobjTituloRec = Nothing
    Set gobjParcelaRec = Nothing
    Set gobjBaixaParcRec = Nothing
    Set gobjBaixaReceber = Nothing
    
    Set objGridParcelas = Nothing

    Set objEventoContaCorrenteInt = Nothing
    
    Set gcolParcelasRec = Nothing
    
End Sub

Private Sub UpDownDataBaixa_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_DownClick

    'Diminui a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 46424

    Exit Sub

Erro_UpDownDataBaixa_DownClick:

    Select Case Err

        Case 46424

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143365)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBaixa_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_UpClick

    'Aumenta a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 46436

    Exit Sub

Erro_UpDownDataBaixa_UpClick:

    Select Case Err

        Case 46436

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143366)

    End Select

    Exit Sub

End Sub

Private Sub TotalRecebido_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TotalRecebido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TotalRecebido_Validate

    If Len(Trim(TotalRecebido.Text)) = 0 Then Exit Sub
    
    lErro = Valor_Positivo_Critica(TotalRecebido.Text)
    If lErro <> SUCESSO Then Error 46425
    
    Exit Sub
    
Erro_TotalRecebido_Validate:

    Cancel = True


    Select Case Err
    
        Case 46425
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143367)
            
    End Select

    Exit Sub

End Sub

Private Sub ContaCorrente_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se a Conta está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o ítem selecionado na ComboBox CodConta
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se o a Conta existe na Combo, e , se existir, seleciona
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 46426

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Passa o Código da Conta para o Obj
        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 46427

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 46428
        
        If giFilialEmpresa <> EMPRESA_TODA Then
        
            'Verifica se a Conta é Filial Empresa corrente
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 46430
        
        End If
        
        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 46429
    
    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 46426, 46427

        Case 46428
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 46429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)
            
        Case 46430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_FILIAL_DIFERENTE", Err, objContaCorrenteInt.iCodigo, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143368)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoDocumento() As Long
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim iTipo As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then Error 46437

    'Carrega a combobox com as Siglas - DescricoesReduzidas lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
        If objTipoDocumento.sSigla = TIPODOC_NF_FATURA_RECEBER Then iTipo = iIndice - 1
    
    Next


    Tipo.ListIndex = iTipo

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = Err

    Select Case Err

        Case 46437

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143369)

    End Select

    Exit Function

End Function

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_GotFocus()
    Call Grid_Recebe_Foco(objGridParcelas)
End Sub

Private Sub GridParcelas_EnterCell()
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
End Sub

Private Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim iLinhasExistentes  As Integer
Dim iLinhaAtual As Integer

    iLinhasExistentes = objGridParcelas.iLinhasExistentes
    iLinhaAtual = GridParcelas.Row
    
    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

    If iLinhasExistentes > objGridParcelas.iLinhasExistentes Then
        If gcolParcelasRec.Count > objGridParcelas.iLinhasExistentes Then
            gcolParcelasRec.Remove iLinhaAtual
        End If
    End If

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridParcelas)
End Sub

Private Sub GridParcelas_RowColChange()
    Call Grid_RowColChange(objGridParcelas)
End Sub

Private Sub GridParcelas_Scroll()
    Call Grid_Scroll(objGridParcelas)
End Sub

Private Sub ValorJuros_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorJuros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorJuros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorJuros_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorJuros
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridParcelas.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialEmpresa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridParcelas.objControle = FilialEmpresa
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub FilialEmpresa_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Parcela_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Parcela_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridParcelas.objControle = Parcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorMulta_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorMulta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorMulta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorMulta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorMulta
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub ValorDesconto_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGridParcelas.objControle = ValorDesconto
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorRecebido_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorRecebido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorRecebido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorRecebido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorRecebido
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        Select Case GridParcelas.Col

            Case iGrid_Tipo_Col
        
                lErro = Saida_Celula_Tipo(objGridInt)
                If lErro <> SUCESSO Then gError 46438
            
            Case iGrid_FilialEmpresa_Col
        
                lErro = Saida_Celula_FilialEmpresa(objGridInt)
                If lErro <> SUCESSO Then gError 20801
            
            Case iGrid_Numero_Col
        
                lErro = Saida_Celula_Numero(objGridInt)
                If lErro <> SUCESSO Then gError 46439
            
            Case iGrid_Parcela_Col
            
                lErro = Saida_Celula_Parcela(objGridInt)
                If lErro <> SUCESSO Then gError 46440
            
            Case iGrid_ValorRecebido_Col
        
                lErro = Saida_Celula_ValorRecebido(objGridInt)
                If lErro <> SUCESSO Then gError 46444
                
            Case iGrid_ValorMulta_Col
                
                lErro = Saida_Celula_ValorMulta(objGridInt)
                If lErro <> SUCESSO Then gError 82262
                
            Case iGrid_ValorJuros_Col
            
                lErro = Saida_Celula_ValorJuros(objGridInt)
                If lErro <> SUCESSO Then gError 82263
                
            Case iGrid_ValorDesconto_Col
            
                lErro = Saida_Celula_ValorDesconto(objGridInt)
                If lErro <> SUCESSO Then gError 82264
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 46445
        
    End If

    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 20801, 46438, 46439, 46440, 46441, 46442, 46443, 46444
        
        Case 82262, 82263, 82264
        
        Case 46445
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 49721
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143370)
    
    End Select

    Exit Function

End Function


Private Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Saida_Celula_Tipo

    Set objGridInt.objControle = Tipo

    If Tipo.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col) Then
    
        If Len(Trim(Tipo.Text)) > 0 Then
        
            If Tipo.ListIndex = -1 Then
            
                lErro = CF("SCombo_Seleciona", Tipo)
                If lErro <> SUCESSO And lErro <> 60483 Then gError 46446
            
                'Se nao encontrar -> Erro
                If lErro = 60483 Then gError 46447
                
            End If
            
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))) > 0 Then
            
                objTituloReceber.lNumTitulo = CLng(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))
                objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
                
                If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col))) = 0 Then
                
                    objTituloReceber.iFilialEmpresa = giFilialEmpresa
                
                    objFilialEmpresa.iCodFilial = giFilialEmpresa
        
                    'Le a Filial Empresa para pegar a descrição
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                    If lErro <> SUCESSO Then gError 71848
        
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                
                Else
                    objTituloReceber.iFilialEmpresa = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col))
                End If
                
                lErro = Carrega_Dados_TituloReceber(objTituloReceber)
                If lErro <> SUCESSO Then gError 46469
                
                
            End If
            
            If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            End If
            
        Else
        
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
        
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Parcela_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""
        
        
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46448
    
    Saida_Celula_Tipo = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = gErr
    
    Select Case gErr
    
        Case 46446, 46448, 46469, 71848
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 46447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, Tipo.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143371)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_FilialEmpresa(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim objFilialEmpresa As New AdmFiliais
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_FilialEmpresa

    Set objGridInt.objControle = FilialEmpresa

    If Len(Trim(FilialEmpresa.Text)) <> 0 Then

        'Verifica se é uma FilialEmpresa selecionada
        If FilialEmpresa.Text <> FilialEmpresa.List(FilialEmpresa.ListIndex) Then

            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialEmpresa, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 71841

            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then

                objFilialEmpresa.iCodFilial = iCodigo

                'Pesquisa se existe FilialEmpresa com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 71842

                'Se não encontrou a FilialEmpresa
                If lErro = 27378 Then gError 71843

                'coloca na tela
                FilialEmpresa.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome

            End If

            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 71844

        End If

        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))) > 0 And Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col))) > 0 Then
        
            objTituloReceber.lNumTitulo = CLng(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))
            objTituloReceber.sSiglaDocumento = SCodigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col))
            objTituloReceber.iFilialEmpresa = Codigo_Extrai(FilialEmpresa.Text)
            
            lErro = Carrega_Dados_TituloReceber(objTituloReceber)
            If lErro <> SUCESSO Then gError 71846
            
            
        End If
                
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
                
    Else
    
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
    
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Parcela_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""
    
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 71845
    
    Saida_Celula_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_FilialEmpresa:

    Saida_Celula_FilialEmpresa = gErr
    
    Select Case gErr
    
        Case 71841, 71842, 71845, 71846
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 71843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialEmpresa.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 71844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialEmpresa.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143372)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Numero(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Saida_Celula_Numero

    Set objGridInt.objControle = Numero
    
    If Len(Trim(Numero.Text)) > 0 Then
    
        lErro = Long_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError 46449
        
        If Numero.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col) Then

            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col))) > 0 Then
                
                objTituloReceber.sSiglaDocumento = SCodigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col))
                objTituloReceber.lNumTitulo = CLng(Numero.Text)
                
                If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col))) = 0 Then
                
                    objTituloReceber.iFilialEmpresa = giFilialEmpresa
                
                    objFilialEmpresa.iCodFilial = giFilialEmpresa
        
                    'Le a Filial Empresa para pegar a descrição
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                    If lErro <> SUCESSO Then gError 71847
        
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                
                Else
                    objTituloReceber.iFilialEmpresa = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col))
                End If
                
                lErro = Carrega_Dados_TituloReceber(objTituloReceber)
                If lErro <> SUCESSO Then gError 46472
            
            End If
        
        End If
        
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    Else
    
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
    
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Parcela_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46450
    
    Saida_Celula_Numero = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Numero:

    Saida_Celula_Numero = gErr
    
    Select Case gErr
    
        Case 46449, 46450, 46472, 71847
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143373)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Parcela(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iParcela As Integer

On Error GoTo Erro_Saida_Celula_Parcela

    Set objGridInt.objControle = Parcela
    
    If Len(Trim(Parcela.Text)) > 0 Then
    
        lErro = Inteiro_Critica(Parcela.Text)
        If lErro <> SUCESSO Then Error 46451
        
        iParcela = StrParaInt(Parcela)
        
        lErro = Carrega_Dados_ParcelaReceber(iParcela)
        If lErro <> SUCESSO Then Error 46473
               
    Else
    
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 46452
    
    Saida_Celula_Parcela = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Parcela:

    Saida_Celula_Parcela = Err
    
    Select Case Err
    
        Case 46451, 46452, 46473
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143374)
            
    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_ValorRecebido(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorRecebido As Double
Dim dValorDesconto As Double
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorSaldo As Double
Dim dSomaTotal As Double

On Error GoTo Erro_Saida_Celula_ValorRecebido

    Set objGridInt.objControle = ValorRecebido

    'Se ValorRecebido está preenchido
    If Len(Trim(ValorRecebido.Text)) <> 0 Then

        'Verifica se ValorRecebido é válido
        lErro = Valor_Positivo_Critica(ValorRecebido.Text)
        If lErro <> SUCESSO Then Error 46459

        dValorRecebido = StrParaDbl(ValorRecebido.Text)

        ValorRecebido.Text = Format(dValorRecebido, "Standard")

    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 46462
    
    Call Grid_Coluna_Soma(iGrid_ValorRecebido_Col, dSomaTotal)
    
    TotalRecebidoGrid.Caption = Format(dSomaTotal, "Standard")
    
    Saida_Celula_ValorRecebido = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorRecebido:

    Saida_Celula_ValorRecebido = Err

    Select Case Err

        Case 46459, 46462
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143375)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorMulta(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorMulta As Double
Dim dSomaTotal As Double
Dim dtDataBaixa  As Date
Dim dtDataVencimento As Date

On Error GoTo Erro_Saida_Celula_ValorMulta

    Set objGridInt.objControle = ValorMulta

    'Se ValorMulta está preenchido
    If Len(Trim(ValorMulta.Text)) <> 0 Then

        'Verifica se ValorMulta é válido
        lErro = Valor_NaoNegativo_Critica(ValorMulta.Text)
        If lErro <> SUCESSO Then gError 82265

        If Len(Trim(DataBaixa.ClipText)) = 0 Then gError 82272
                
        dValorMulta = StrParaDbl(ValorMulta.Text)
        
        dtDataBaixa = CDate(DataBaixa.Text)
        dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col))

        'Verifica se a Multa é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
        If dtDataBaixa <= dtDataVencimento And dValorMulta > 0 Then gError 82274

        ValorMulta.Text = Format(dValorMulta, "Standard")

    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 82266
    
    Saida_Celula_ValorMulta = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorMulta:

    Saida_Celula_ValorMulta = gErr

    Select Case gErr

        Case 82265, 82266
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 82272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_JUROS_DATABAIXA_VAZIA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 82274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_JUROS_INCOMPATIVEL_DATABAIXA_DATAVENC", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143376)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorJuros(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorJuros As Double
Dim dSomaTotal As Double
Dim dtDataBaixa  As Date
Dim dtDataVencimento As Date

On Error GoTo Erro_Saida_Celula_ValorJuros

    Set objGridInt.objControle = ValorJuros

    'Se ValorJuros está preenchido
    If Len(Trim(ValorJuros.Text)) <> 0 Then

        'Verifica se ValorJuros é válido
        lErro = Valor_NaoNegativo_Critica(ValorJuros.Text)
        If lErro <> SUCESSO Then gError 82267

        If Len(Trim(DataBaixa.ClipText)) = 0 Then gError 82273
        
        dValorJuros = StrParaDbl(ValorJuros.Text)

        dtDataBaixa = CDate(DataBaixa.Text)
        dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col))

        'Verifica se o Juros é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
        If dtDataBaixa <= dtDataVencimento And dValorJuros > 0 Then gError 82275

        ValorJuros.Text = Format(dValorJuros, "Standard")

    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 82268
    
    Saida_Celula_ValorJuros = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorJuros:

    Saida_Celula_ValorJuros = gErr

    Select Case gErr

        Case 82267, 82268
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 82273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_JUROS_DATABAIXA_VAZIA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 82275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_JUROS_INCOMPATIVEL_DATABAIXA_DATAVENC", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143377)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorDesconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorDesconto As Double
Dim dSomaTotal As Double
Dim dValorParcela  As Double

On Error GoTo Erro_Saida_Celula_ValorDesconto

    Set objGridInt.objControle = ValorDesconto

    'Se ValorDesconto está preenchido
    If Len(Trim(ValorDesconto.Text)) <> 0 Then

        'Verifica se ValorDesconto é válido
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 82269

        dValorDesconto = StrParaDbl(ValorDesconto.Text)
       
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))) <> 0 Then
            dValorParcela = (StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))) + (StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col))) + (StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col)))
        End If
        
        'Verifica se o Desconto é maior ou igual ao ValorParcela
        If dValorDesconto >= dValorParcela Then gError 82271

        ValorDesconto.Text = Format(dValorDesconto, "Standard")

    End If

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 82270
    
    Saida_Celula_ValorDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorDesconto:

    Saida_Celula_ValorDesconto = gErr

    Select Case gErr

        Case 82269, 82270
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 82271
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_MENOR_DESCONTO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143378)

    End Select

    Exit Function

End Function

Private Function Carrega_Dados_TituloReceber(objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Carrega_Dados_TituloReceber

    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Parcela_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = ""
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = ""

    lErro = CF("TituloReceber_Le_Numero_Sigla", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 46466 Then Error 46467
    If lErro <> SUCESSO Then Error 46468
    
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = objTituloReceber.lCliente
    
    objCliente.lCodigo = objTituloReceber.lCliente
    
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 46470
    If lErro <> SUCESSO Then Error 46471
    
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col) = objCliente.lCodigo
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Nome_Col) = objCliente.sNomeReduzido

    Carrega_Dados_TituloReceber = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_TituloReceber:

    Carrega_Dados_TituloReceber = Err
    
    Select Case Err
    
        Case 46467
        
        Case 46468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", Err, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143379)
            
    End Select
    
    Exit Function

End Function

Sub Rotina_Grid_Enable(iLinha As Integer, objControle As Object, iCaminho As Integer)
   
    'Pesquisa a controle da coluna em questão
    Select Case objControle.Name

        Case Parcela.Name
        
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))) > 0 Then
                Parcela.Enabled = True
            Else
                Parcela.Enabled = False
            End If
            
        Case ValorRecebido.Name, ValorMulta.Name, ValorJuros.Name, ValorDesconto.Name
            
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col))) > 0 Then
                objControle.Enabled = True
            Else
                objControle.Enabled = False
            End If
           
    End Select
    
    Exit Sub

End Sub

Private Function Carrega_Dados_ParcelaReceber(iParcela As Integer) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber
Dim objParcelaCol As ClassParcelaReceber
Dim objCobrador As New ClassCobrador
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Dados_ParcelaReceber
    
    If gcolParcelasRec.Count >= GridParcelas.Row Then
        Set objParcelaReceber = gcolParcelasRec(GridParcelas.Row)
    End If
    
    objTituloReceber.lNumTitulo = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col))
    objTituloReceber.iFilialEmpresa = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col))
    objParcelaReceber.iNumParcela = iParcela
    
    lErro = CF("ParcelaReceber_Le_NumTitulo", objTituloReceber, objParcelaReceber)
    If lErro <> SUCESSO And lErro <> 46477 Then Error 46478
    If lErro <> SUCESSO Then Error 46479
        
    If objParcelaReceber.iStatus <> STATUS_ABERTO Then Error 59183
    If objParcelaReceber.iCobrador = COBRADOR_PROPRIA_EMPRESA And objParcelaReceber.iCarteiraCobranca = CARTEIRA_CHEQUEPRE Then Error 59184
        
    iIndice = 0
    
    'pesquisa para garantir que uma parcela nao esteja sendo entrada em duplicidade
    For Each objParcelaCol In gcolParcelasRec
        iIndice = iIndice + 1
        If iIndice <> GridParcelas.Row Then
            If objParcelaCol.lNumIntDoc = objParcelaReceber.lNumIntDoc Then Error 46513
        End If
    Next
    
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col) = Format(objParcelaReceber.dSaldo, "Standard")
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col) = Format(objParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = Format(objParcelaReceber.dValor, "Standard")
    
    'Preenche o objCobrador
    objCobrador.iCodigo = objParcelaReceber.iCobrador
    
    'le o nomeReduzido do cobrador
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 56793
    
    If lErro = 19294 Then Error 56794
    
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobrador_Col) = CStr(objCobrador.iCodigo) & SEPARADOR & objCobrador.sNomeReduzido
        
    With objParcelaReceber
        
        If gcolParcelasRec.Count < GridParcelas.Row Then
        '####################################
        'ALTERADO POR WAGNER
            gcolParcelasRec.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, 0, .iCarteiraCobranca, .iCobrador, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValorOriginal
        '####################################
        End If
                
    End With
    
    'Se a DataBaixa está preenchida calcula a Multa o juros das parcelas
    If Len(DataBaixa.ClipText) > 0 Then
   
        lErro = Calcula_Multa_Juros_Desc_Parcela(GridParcelas.Row)
        If lErro <> SUCESSO Then Error 49553
    
    End If
    
    Carrega_Dados_ParcelaReceber = SUCESSO

    Exit Function

Erro_Carrega_Dados_ParcelaReceber:

    Carrega_Dados_ParcelaReceber = Err
    
    Select Case Err
    
        Case 46478, 56793
        
        Case 46479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", Err)
        
        Case 46513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_JA_EXISTENTE", Err, iIndice)
        
        Case 49553
        
        Case 56794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", Err, CStr(objCobrador.iCodigo))
            
        Case 59183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_NAO_ABERTA", Err)
        
        Case 59184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_VINCULADA_CHQPRE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143380)
            
    End Select

    Exit Function

End Function

Private Sub Grid_Coluna_Soma(iColuna As Integer, dSomaTotal As Double)

Dim iIndice As Integer

    dSomaTotal = 0

    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        dSomaTotal = dSomaTotal + StrParaDbl(GridParcelas.TextMatrix(iIndice, iColuna))
    
    Next

End Sub

Private Sub Limpa_Tela_BaixaRecDig()

    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridParcelas)
    
    Set gcolParcelasRec = New ColParcelaReceber
    giDataAnterior = 0
    
    ContaCorrente.Text = ""
    TotalRecebidoGrid.Caption = ""
    
    DataBaixa.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCredito.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    gsContaCtaCorrente = ""
    gsContaFilDep = ""
    giFilialEmpresaConta = 0
    
    iAlterado = 0
    
End Sub

Private Sub DataCredito_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCredito_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataCredito_Validate

    'Se a DataCredito está preenchida
    If Len(DataCredito.ClipText) > 0 Then

        'Verifica se a DataCredito é válida
        lErro = Data_Critica(DataCredito.Text)
        If lErro <> SUCESSO Then Error 46480

    'Se a DataCredito não está preenchida
    Else
    
        Error 46481
        
    End If

    Exit Sub

Erro_DataCredito_Validate:

    Cancel = True


    Select Case Err

        Case 46480

        Case 46481
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143381)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCredito_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCredito_DownClick

    'Diminui a DataCredito em 1 dia
    lErro = Data_Up_Down_Click(DataCredito, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 46482

    Exit Sub

Erro_UpDownDataCredito_DownClick:

    Select Case Err

        Case 46482

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143382)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCredito_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCredito_UpClick

    'Aumenta a DataCredito em 1 dia
    lErro = Data_Up_Down_Click(DataCredito, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 46483

    Exit Sub

Erro_UpDownDataCredito_UpClick:

    Select Case Err

        Case 46483

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143383)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 46484

    Call Limpa_Tela_BaixaRecDig
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 46484

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143384)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iQuantidade As Integer, objBaixarRecDig As New ClassBaixaRecDig
Dim colBaixaParcRec As New colBaixaParcRec
Dim objMovCCI As New ClassMovContaCorrente

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(DataBaixa.ClipText)) = 0 Then Error 46485
    If Len(Trim(DataCredito.ClipText)) = 0 Then Error 46486
    If Len(Trim(NumDocInf.Text)) = 0 Then Error 46487
    If Len(Trim(TotalRecebido.Text)) = 0 Then Error 46488
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 46489
    
    If objGridParcelas.iLinhasExistentes = 0 Then Error 46490
    
    iQuantidade = StrParaInt(NumDocInf.Text)
    
    If iQuantidade <> objGridParcelas.iLinhasExistentes Then Error 46491
    
    lErro = Valida_Dados_GridParcelas()
    If lErro <> SUCESSO Then Error 46492
    
    lErro = Move_Tela_Memoria(colBaixaParcRec, objMovCCI)
    If lErro <> SUCESSO Then Error 46493
    
    Set objBaixarRecDig.colBaixaParcRec = colBaixaParcRec
    Set objBaixarRecDig.objMovCCI = objMovCCI
    objBaixarRecDig.objTelaAtualizacao = Me
    
    lErro = CF("BaixaRecDig_Grava", objBaixarRecDig)
    If lErro <> SUCESSO Then Error 46494
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 46485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_SEM_PREENCHIMENTO", Err)
        
        Case 46486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CREDITO_NAO_PREENCHIDA", Err)

        Case 46487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", Err)
            
        Case 46488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTAL_RECEBIDO_NAO_PREENCHIDO", Err)
            
        Case 46489
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PREENCHIDA", Err)
        
        Case 46490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", Err)
        
        Case 46491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_INFORMADA_DIFERENTE_GRID", Err)
            
        Case 46492, 46493, 46494
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143385)
            
    End Select
    
    Exit Function
        
End Function

Function Valida_Dados_GridParcelas() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValorSaldo As Double
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim dValorRecebido As Double
Dim dValorBaixado As Double
Dim dTotalRecebidoGrid As Double
Dim dTotalRecebido As Double

On Error GoTo Erro_Valida_Dados_GridParcelas

    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Tipo_Col))) = 0 Then Error 46495
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_FilialEmpresa_Col))) = 0 Then Error 20803
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Numero_Col))) = 0 Then Error 46496
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Parcela_Col))) = 0 Then Error 46497
        
        dValorSaldo = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Saldo_Col))
        dValorMulta = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorMulta_Col))
        dValorJuros = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorJuros_Col))
        dValorDesconto = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorDesconto_Col))
        dValorRecebido = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorRecebido_Col))
               
        If dValorRecebido + dValorMulta + dValorJuros < dValorDesconto Then Error 46498
        If dValorRecebido = 0 Then Error 46499
    
    Next
    
    dTotalRecebido = StrParaDbl(TotalRecebido.Text)
    dTotalRecebidoGrid = StrParaDbl(TotalRecebidoGrid.Caption)
    
    If Format(dTotalRecebido, "Standard") <> Format(dTotalRecebidoGrid, "Standard") Then Error 46505
        
    Valida_Dados_GridParcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Dados_GridParcelas:

    Valida_Dados_GridParcelas = Err
    
    Select Case Err
    
        Case 20803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_INFORMADO", Err, iIndice)
        
        Case 46495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_PARCELA_NAO_INFORMADO", Err, iIndice)
        
        Case 46496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_PARCELA_NAO_INFORMADO", Err, iIndice)
            
        Case 46497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMPARCELA_NAO_INFORMADO", Err, iIndice)
            
        Case 46498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_SUPERIOR_SOMA_VALOR", Err, iIndice)
        
        Case 46499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_RECEBIDO_PARCELA_NAO_PREENCHIDO", Err, iIndice)
        
        Case 46505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TOTALRECEBIDO_PARCELAS_DIFERENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143386)
    End Select
    
    Exit Function

End Function

Private Function Move_Tela_Memoria(colBaixaParcRec As colBaixaParcRec, objMovCCI As ClassMovContaCorrente) As Long

Dim lErro As Long
Dim objParcelaReceber As ClassParcelaReceber
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim dValorRecebido As Double
Dim dValorBaixado As Double
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim iIndice As Integer
Dim dValorSaldo As Double

On Error GoTo Erro_Move_Tela_Memoria

    iIndice = 0

    For Each objParcelaReceber In gcolParcelasRec
        
        iIndice = iIndice + 1
        
        dValorMulta = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorMulta_Col))
        dValorJuros = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorJuros_Col))
        dValorDesconto = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorDesconto_Col))
        dValorRecebido = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorRecebido_Col))
        dValorSaldo = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Saldo_Col))
        
        dValorBaixado = dValorRecebido - dValorMulta - dValorJuros + dValorDesconto
                
        If dValorBaixado > dValorSaldo Then dValorBaixado = dValorSaldo
        
        colBaixaParcRec.Add 0, 0, objParcelaReceber.lNumIntDoc, objParcelaReceber.iNumParcela, STATUS_LANCADO, dValorMulta, dValorJuros, dValorDesconto, dValorBaixado, dValorRecebido, objParcelaReceber.iCobrador
        
    Next
        
    objContaCorrente.iCodigo = Codigo_Extrai(ContaCorrente.Text)
        
    lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 46500
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 46501
    
    objMovCCI.iFilialEmpresa = objContaCorrente.iFilialEmpresa
    objMovCCI.iCodConta = objContaCorrente.iCodigo
    objMovCCI.iTipo = MOVCCI_RECEBIMENTO_TITULO
    objMovCCI.iExcluido = NAO_EXCLUIDO
    objMovCCI.iTipoMeioPagto = DINHEIRO
    objMovCCI.dtDataBaixa = MaskedParaDate(DataBaixa)
    objMovCCI.dtDataMovimento = MaskedParaDate(DataCredito)
    objMovCCI.dtDataContabil = objMovCCI.dtDataMovimento
    objMovCCI.dValor = TotalRecebidoGrid.Caption
    objMovCCI.sHistorico = Historico.Text
    objMovCCI.iConciliado = NAO_CONCILIADO

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err
    
    Select Case Err
    
        Case 46500
        
        Case 46501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, ContaCorrente.Text)
            ContaCorrente.SetFocus
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143387)
            
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_Multa_Juros_Desc_Parcelas() As Long
'Preenche para todas as parcelas do grid os valores para multa, juros e descontos

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Calcula_Multa_Juros_Desc_Parcelas
    
    For iLinha = 1 To objGridParcelas.iLinhasExistentes
    
        lErro = Calcula_Multa_Juros_Desc_Parcela(iLinha)
        If lErro <> SUCESSO Then Error 56721
    
    Next

    Calcula_Multa_Juros_Desc_Parcelas = SUCESSO
    
    Exit Function

Erro_Calcula_Multa_Juros_Desc_Parcelas:

    Calcula_Multa_Juros_Desc_Parcelas = Err
    
    Select Case Err
    
        Case 56721
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143388)
    
    End Select
    
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, objCliente As New ClassCliente, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor
Dim iLinha As Integer, iConta As Integer, sContaContabil As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, sContaTela As String
Dim objCarteiraCobrador As New ClassCarteiraCobrador, dValor As Double
Dim objBaixaParcRec As ClassBaixaParcRec, objParcelaReceber As ClassParcelaReceber
 
On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case "Valor_Recebido_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorRecebido
            Next
        
        Case "Valor_Baixar_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorBaixado
            Next
        
        Case "Valor_Desconto_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorDesconto
            Next
        
        Case "Valor_Juros_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorJuros
            Next
        
        Case "Valor_Multa_Det"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                objMnemonicoValor.colValor.Add objBaixaParcRec.dValorMulta
            Next
        
        Case "Numero_Titulo_Det"
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaLong(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col))
            Next
        
        Case "Parcela_Det"
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaInt(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col))
            Next
        
        Case "Cliente_Codigo_Det"
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaLong(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))
            Next
        
        Case "Cta_CartCobr_Det"
                        
            For Each objParcelaReceber In gcolParcelasRec
        
                objCarteiraCobrador.iCobrador = objParcelaReceber.iCobrador
                objCarteiraCobrador.iCodCarteiraCobranca = objParcelaReceber.iCarteiraCobranca
                
                lErro = CartCobr_ObtemCtaTela(objCarteiraCobrador, sContaTela)
                If lErro <> SUCESSO Then Error 32272
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Next
            
        Case "Valor_Recebido"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorRecebido
        
        Case "Valor_Baixar"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorBaixado
        
        Case "Valor_Desconto"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorDesconto
        
        Case "Valor_Juros"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorJuros
        
        Case "Valor_Multa"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorMulta
        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
            'calcula-la apenas uma vez e deixa-la guardada
                
            If gsContaCtaCorrente = "" Then
                
                iConta = Codigo_Extrai(ContaCorrente.Text)
                lErro = CF("ContaCorrenteInt_Le", iConta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 56546
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56547
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                
        Case "FilDep_Cta_Transf" 'conta de transferencia da filial do deposito

            If gsContaFilDep = "" Then
            
                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(giFilialEmpresaConta, sContaContabil)
                If lErro <> SUCESSO Then Error 56548
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56549
                    
                Else
                
                    sContaTela = ""
                    
                End If
            
                gsContaFilDep = sContaTela
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaFilDep
        
        Case "FilNaoDep_Cta_Transf" 'conta de transferencia da filial da parcela

                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloRec.iFilialEmpresa, sContaContabil)
                If lErro <> SUCESSO Then Error 56550
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56551
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
        
        Case "Numero_Titulo"
        
            objMnemonicoValor.colValor.Add gobjTituloRec.lNumTitulo
            
        Case "Parcela"
            objMnemonicoValor.colValor.Add gobjParcelaRec.iNumParcela
        
        Case "Cliente_Codigo"
            
            objMnemonicoValor.colValor.Add gobjTituloRec.lCliente
                    
        
        Case "Cliente_Nome"
        
            objCliente.lCodigo = gobjTituloRec.lCliente
            
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO Then Error 32253
            
            objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
        Case "FilialCli_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloRec.iFilial
        
        Case DATA_BAIXA
            If Len(DataBaixa.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataBaixa.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case DATA_CREDITO
            If Len(DataCredito.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataCredito.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case CTACARTEIRACOBRADOR
        
            objCarteiraCobrador.iCobrador = gobjParcelaRec.iCobrador
            objCarteiraCobrador.iCodCarteiraCobranca = gobjParcelaRec.iCarteiraCobranca
            
            If objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA Then
            
                Select Case objCarteiraCobrador.iCodCarteiraCobranca
                
                    Case CARTEIRA_CARTEIRA
                        sCampoGlobal = "CtaReceberCarteira"
                    
                    Case CARTEIRA_CHEQUEPRE
                        sCampoGlobal = "CtaRecChequePre"
                        
                    Case CARTEIRA_JURIDICO
                        sCampoGlobal = "CtaJuridico"
                    
                    Case Else
                        Error 56802
                        
                End Select
                
                objMnemonico.sMnemonico = sCampoGlobal
                lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
                If lErro <> SUCESSO And lErro <> 39690 Then Error 56803
                If lErro <> SUCESSO Then Error 56804
                
                sContaTela = objMnemonico.sValor
                
            Else
            
                lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
                If lErro <> SUCESSO And lErro <> 23551 Then Error 56528
                If lErro <> SUCESSO Then Error 56798
                
                If objCarteiraCobrador.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56529
                
                Else
                
                    sContaTela = ""
                    
                End If
        
            End If
            
            objMnemonicoValor.colValor.Add sContaTela
        
        Case Else
            Error 39695

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 32253, 56527, 56528, 56529, 56530, 56803, 32272
        
        Case 56798, 56802
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56804
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143389)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

Private Function Carrega_FilialEmpresa() As Long
'Carrega a combobox FilialEmpresa

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o Código e o Nome de todas as FiliaisEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 71849

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 71849
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143390)

    End Select

    Exit Function

End Function


Private Function Calcula_Multa_Juros_Desc_Parcela(iLinha As Integer) As Long
'Preenche para a parcela da linha do grid passada como parametro os valores para multa, juros e descontos

Dim lErro As Long
Dim objParcelaReceber As ClassParcelaReceber
Dim dtDataBaixa As Date, dtDataVenctoReal As Date
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim objTituloReceber As New ClassTituloReceber
Dim iDias As Integer

On Error GoTo Erro_Calcula_Multa_Juros_Desc_Parcela
    
    If gcolParcelasRec.Count >= iLinha Then
    
        dtDataBaixa = MaskedParaDate(DataBaixa)
        
        Set objParcelaReceber = gcolParcelasRec(iLinha)
            
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", objParcelaReceber.dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 56785

        If dtDataBaixa > dtDataVenctoReal Then
                    
            lErro = CF("Calcula_Multa_Juros_Parcela", objParcelaReceber, dtDataBaixa, dValorMulta, dValorJuros)
            If lErro <> SUCESSO Then Error 56720
            
            GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col) = Format(dValorMulta, "Standard")
            GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col) = Format(dValorJuros, "Standard")
            GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = ""
                   
         Else
                   
            GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col) = ""
            GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col) = ""
            
            lErro = CF("Calcula_Desconto_Parcela", objParcelaReceber, dValorDesconto, dtDataBaixa)
            If lErro <> SUCESSO Then Error 46511
            
            If dValorDesconto > 0 Then
                GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = Format(dValorDesconto, "Standard")
            Else
                GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = ""
            End If
                   
        End If

    End If

    Calcula_Multa_Juros_Desc_Parcela = SUCESSO
    
    Exit Function

Erro_Calcula_Multa_Juros_Desc_Parcela:

    Calcula_Multa_Juros_Desc_Parcela = Err
    
    Select Case Err
    
        Case 56720, 46511, 56785
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143391)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BAIXA_TITULOS_RECEBER_RECIBOS
    Set Form_Load_Ocx = Me
    Caption = "Baixas de Títulos a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaRecDig"
    
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
        
        If Me.ActiveControl Is ContaCorrente Then
            Call ContaCorrenteLabel_Click
        End If
    
    End If
    
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrenteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrenteLabel, Source, X, Y)
End Sub

Private Sub ContaCorrenteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrenteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub TotalRecebidoGrid_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalRecebidoGrid, Source, X, Y)
End Sub

Private Sub TotalRecebidoGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalRecebidoGrid, Button, Shift, X, Y)
End Sub

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcrec e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iConta As Integer, dValorLivroAux As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcRec = vParams(0)
    Set gobjParcelaRec = vParams(1)
    Set gobjTituloRec = vParams(2)
    Set gobjBaixaReceber = vParams(3)

    'se ainda nao obtive a filial empresa onde vai ser feito o deposito
    If giFilialEmpresaConta = 0 Then
    
        iConta = Codigo_Extrai(ContaCorrente.Text)
    
        lErro = CF("ContaCorrenteInt_Le", iConta, objContasCorrentesInternas)
        If lErro <> SUCESSO Then Error 32243
    
        giFilialEmpresaConta = objContasCorrentesInternas.iFilialEmpresa
        
    End If
    
    'obtem numero de doc para a filial onde vai ser feito o deposito
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32244
        
    'se contabiliza parcela p/parcela
    If gobjCR.iContabSemDet = 0 Then
    
        dValorLivroAux = Round(gobjBaixaParcRec.dValorRecebido + gobjBaixaParcRec.dValorDesconto - gobjBaixaParcRec.dValorJuros - gobjBaixaParcRec.dValorMulta, 2)
    
        'se a filial onde vai ser feito o deposito é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If giFilialEmpresaConta <> gobjTituloRec.iFilialEmpresa And giContabCentralizada = 0 Then
                        
            'grava a contabilizacao na filial onde vai ser feito o deposito
            lErro = objContabAutomatica.Gravar_Registro(Me, "BaixaRecDigFilDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
            If lErro <> SUCESSO Then Error 32245
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloRec.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32246
        
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "BaixaRecDigFilNaoDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloRec.iFilialEmpresa, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32247
        
        Else
        
            'grava a contabilizacao na filial da cta (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "BaixaRecDig", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32248
        
        End If
    
    Else
    
        GridParcelas.Tag = gobjBaixaReceber.colBaixaParcRec.Count
    
        'grava a contabilizacao na filial da cta
        lErro = objContabAutomatica.Gravar_Registro(Me, "BaixaRecDigRes", gobjBaixaReceber.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
        If lErro <> SUCESSO Then Error 32248
    
    End If
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32243 To 32248
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143392)
     
    End Select
     
    Exit Function

End Function

'??? já existe em ctbaixareccancelar.cls
Private Function CartCobr_ObtemCtaTela(objCarteiraCobrador As ClassCarteiraCobrador, sContaTela As String) As Long

Dim lErro As Long, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_CartCobr_ObtemCtaTela

    If objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA Then

        Select Case objCarteiraCobrador.iCodCarteiraCobranca

            Case CARTEIRA_CARTEIRA
                sCampoGlobal = "CtaReceberCarteira"

            Case CARTEIRA_CHEQUEPRE
                sCampoGlobal = "CtaRecChequePre"

            Case CARTEIRA_JURIDICO
                sCampoGlobal = "CtaJuridico"

            Case Else
                Error 56799

        End Select

        objMnemonico.sMnemonico = sCampoGlobal
        lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
        If lErro <> SUCESSO And lErro <> 39690 Then Error 56800
        If lErro <> SUCESSO Then Error 56801

        sContaTela = objMnemonico.sValor

    Else

        lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 49726
        If lErro <> SUCESSO Then Error 56797

        If objCarteiraCobrador.sContaContabil <> "" Then

            lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
            If lErro <> SUCESSO Then Error 56526

        End If

    End If

    CartCobr_ObtemCtaTela = SUCESSO
     
    Exit Function
    
Erro_CartCobr_ObtemCtaTela:

    CartCobr_ObtemCtaTela = gErr
     
    Select Case gErr
          
        Case 49726, 56526, 56800
        
        Case 56797, 56799
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56801
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143393)
     
    End Select
     
    Exit Function

End Function


