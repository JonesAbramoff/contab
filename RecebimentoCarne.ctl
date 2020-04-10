VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl RecebimentoCarne 
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   8055
   Begin VB.Frame FrameParcelas 
      Caption         =   "Parcelas"
      Height          =   3150
      Left            =   90
      TabIndex        =   21
      Top             =   2850
      Width           =   7815
      Begin VB.Frame FrameTotais 
         Caption         =   "Totais"
         Height          =   840
         Left            =   135
         TabIndex        =   29
         Top             =   2175
         Width           =   7530
         Begin VB.Label LabelCancelar 
            AutoSize        =   -1  'True
            Caption         =   "A receber:"
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
            Left            =   4200
            TabIndex        =   33
            Top             =   405
            Width           =   900
         End
         Begin VB.Label LabelCancelarValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5250
            TabIndex        =   32
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label LabelSelecionadasQuant 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2055
            TabIndex        =   31
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label LabelSelecionadas 
            AutoSize        =   -1  'True
            Caption         =   "Selecionadas:"
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
            Left            =   720
            TabIndex        =   30
            Top             =   405
            Width           =   1215
         End
      End
      Begin VB.CheckBox Selecionar 
         Height          =   300
         Left            =   165
         TabIndex        =   28
         Top             =   675
         Width           =   930
      End
      Begin VB.TextBox CarneNumero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   615
         TabIndex        =   27
         Top             =   630
         Width           =   1785
      End
      Begin VB.TextBox ParcelaNumero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   2040
         TabIndex        =   26
         Top             =   885
         Width           =   735
      End
      Begin VB.TextBox ParcelaValor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   3720
         TabIndex        =   25
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox ParcelaDesconto 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   4680
         TabIndex        =   24
         Top             =   750
         Width           =   975
      End
      Begin VB.TextBox ParcelaJuros 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   6600
         TabIndex        =   23
         Top             =   765
         Width           =   975
      End
      Begin VB.TextBox ParcelaMulta 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5640
         TabIndex        =   22
         Top             =   645
         Width           =   975
      End
      Begin MSMask.MaskEdBox ParcelaVencimento 
         Height          =   300
         Left            =   2775
         TabIndex        =   34
         Top             =   720
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         _Version        =   327681
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   1905
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   3360
         _Version        =   393216
      End
   End
   Begin VB.CommandButton BotaoImprimir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1110
      Picture         =   "RecebimentoCarne.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   60
      Width           =   1560
   End
   Begin VB.CommandButton BotaoSair 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5265
      Picture         =   "RecebimentoCarne.ctx":2FCA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   60
      Width           =   1560
   End
   Begin VB.Frame FrameSelecao 
      Caption         =   "Seleção"
      Height          =   2175
      Left            =   90
      TabIndex        =   0
      Top             =   675
      Width           =   7815
      Begin VB.Frame FrameDataVencimento 
         Caption         =   "Data de Vencimento"
         Height          =   735
         Left            =   4080
         TabIndex        =   16
         Top             =   1155
         Width           =   3615
         Begin MSComCtl2.UpDown UpDownVencimentoDe 
            Height          =   300
            Left            =   1440
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox VencimentoDe 
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   315
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   327681
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownVencimentoAte 
            Height          =   300
            Left            =   3240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox VencimentoAte 
            Height          =   300
            Left            =   2280
            TabIndex        =   6
            Top             =   315
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   327681
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelVencimentoDe 
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
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label20 
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
            Height          =   255
            Left            =   1875
            TabIndex        =   19
            Top             =   345
            Width           =   375
         End
      End
      Begin VB.Frame FrameNumeroCarne 
         Caption         =   "Nº do Carnê"
         Height          =   735
         Left            =   180
         TabIndex        =   10
         Top             =   1155
         Width           =   3870
         Begin MSMask.MaskEdBox CarneAte 
            Height          =   300
            Left            =   2325
            TabIndex        =   4
            Top             =   315
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   327681
            MaxLength       =   20
            Mask            =   "99999999999999999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CarneDe 
            Height          =   300
            Left            =   420
            TabIndex        =   3
            Top             =   330
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   327681
            MaxLength       =   20
            Mask            =   "99999999999999999999"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCarneAte 
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
            Height          =   255
            Left            =   1935
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   15
            Top             =   345
            Width           =   375
         End
         Begin VB.Label LabelCarneDe 
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
            Height          =   255
            Left            =   90
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   14
            Top             =   345
            Width           =   390
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Cliente"
         Height          =   765
         Index           =   6
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   4815
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1440
            TabIndex        =   2
            Top             =   300
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   529
            _Version        =   327681
            PromptInclude   =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.Label LabelCliente 
            AutoSize        =   -1  'True
            Caption         =   "Cliente (F3):"
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
            TabIndex        =   7
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.CommandButton BotaoTrazer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5715
         Picture         =   "RecebimentoCarne.ctx":5ACC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   390
         Width           =   1560
      End
   End
   Begin VB.CommandButton BotaoLimpar 
      Height          =   600
      Left            =   3180
      Picture         =   "RecebimentoCarne.ctx":851A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   60
      Width           =   1560
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   36
      RTSEnable       =   -1  'True
      BaudRate        =   2400
      ParitySetting   =   1
      DataBits        =   7
   End
End
Attribute VB_Name = "RecebimentoCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globáis

Public iAlterado  As Integer
Dim iClienteAlterado  As Integer
Dim gsCliente As String
Dim gcolCarneParcelas As Collection
Dim giOperacao As Integer

'variáveis para grid
Dim objGridParcelas As AdmGrid
Dim igrid_Selecionar_Col As Integer
Dim igrid_Carne_Col As Integer
Dim igrid_ParcNumero_Col As Integer
Dim igrid_ParcVencimento_Col As Integer
Dim igrid_ParcValor_Col As Integer
Dim igrid_ParcJuros_Col As Integer
Dim igrid_ParcMulta_Col As Integer
Dim igrid_ParcDesconto_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gbCarneDe As Boolean
Private WithEvents objEventoCarne As AdmEvento
Attribute objEventoCarne.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Function Trata_Parametros(Optional objRecebimentoCarne As ClassRecebimentoCarne) As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

On Error GoTo Erro_Form_Load
    
    Set objGridParcelas = New AdmGrid
    Set objEventoCliente = New AdmEvento
    Set objEventoCarne = New AdmEvento
    Set gcolCarneParcelas = New Collection
    
    FrameParcelas.Caption = FrameParcelas.Caption & " " & gdtDataAtual
    
    Call Inicializa_GridParcelas(objGridParcelas)
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166221)

    End Select
    
    Exit Sub

End Sub

Private Sub Inicializa_GridParcelas(objGridInt As AdmGrid)

On Error GoTo Erro_Inicializa_GridParcelas

    'form do grid
    Set objGridInt.objForm = Me

    'Títulos das Colunas
    objGridInt.colColuna.Add ""
    objGridInt.colColuna.Add "Selecionar"
    objGridInt.colColuna.Add "Carnê"
    objGridInt.colColuna.Add "Parcela"
    objGridInt.colColuna.Add "Vencimento"
    objGridInt.colColuna.Add "Valor"
    objGridInt.colColuna.Add "Desconto"
    objGridInt.colColuna.Add "Multa"
    objGridInt.colColuna.Add "Juros"

    'Controles que participam do Grid
    objGridInt.colCampo.Add Selecionar.Name
    objGridInt.colCampo.Add CarneNumero.Name
    objGridInt.colCampo.Add ParcelaNumero.Name
    objGridInt.colCampo.Add ParcelaVencimento.Name
    objGridInt.colCampo.Add ParcelaValor.Name
    objGridInt.colCampo.Add ParcelaDesconto.Name
    objGridInt.colCampo.Add ParcelaMulta.Name
    objGridInt.colCampo.Add ParcelaJuros.Name

    'Colunas do Grid
    igrid_Selecionar_Col = 1
    igrid_Carne_Col = 2
    igrid_ParcNumero_Col = 3
    igrid_ParcVencimento_Col = 4
    igrid_ParcValor_Col = 5
    igrid_ParcDesconto_Col = 6
    igrid_ParcMulta_Col = 7
    igrid_ParcJuros_Col = 8

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 50

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 400

    'Largura manual para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Exit Sub

Erro_Inicializa_GridParcelas:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166222)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 109625

    'Limpa a tela
    Call Limpa_Tela_Impressaocarne

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 109625

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166223)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colMovimentosCaixa As New Collection
Dim colCarneParcelasImpressao As New Collection
Dim objRecebimentoCarne As New ClassRecebimentoCarne
Dim objVenda As New ClassVenda

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Se pelo menos uma parcela não está selecionada-->erro.
    If StrParaDbl(LabelCancelarValor.Caption) = 0 Then gError 109626
    
    'Pergunta ao usuário qual a operação ele deseja efetuar
    Call Chama_Tela_Modal("RecebimentoCarne2", objRecebimentoCarne)
    
    If giRetornoTela <> vbCancel Then
        
        lErro = Move_Dados_Memoria(colCarneParcelasImpressao)
        If lErro <> SUCESSO Then gError 109627
        
        'DE acordo com a escolha
        Select Case objRecebimentoCarne.iOpcao
            
            'Se for apenas para imprimir...
            Case 0
                'Imprime o carnê
                lErro = CF_ECF("Caixa_Carne_Imprime_ECF", colCarneParcelasImpressao)
                If lErro <> SUCESSO Then gError 109628
            'se for para gravar e imprimir...
            Case 1
                'Obtem a forma de recebimento destas parcelas
                Call Chama_Tela_Modal("RecebimentoCarne3", colCarneParcelasImpressao, objVenda)
                
                If giRetornoTela <> vbCancel Then
                    'Grava no bd
                    lErro = CF("RecebimentoCarne_Grava", colCarneParcelasImpressao, objVenda)
                    If lErro <> SUCESSO Then gError 109629
                    
                    'Imprime o carnê
                    lErro = CF_ECF("Caixa_Carne_Imprime_ECF", colCarneParcelasImpressao)
                    If lErro <> SUCESSO Then gError 109630
                End If
            'Se for para gravar e autenticar...
            Case 2
                'Obtem a forma de recebimento destas parcelas
                Call Chama_Tela_Modal("RecebimentoCarne3", colCarneParcelasImpressao, objVenda)
                
                If giRetornoTela <> vbCancel Then
                    'Grava no bd
                    lErro = CF("RecebimentoCarne_Grava", colCarneParcelasImpressao, objVenda)
                    If lErro <> SUCESSO Then gError 109631
                    
                    'Autentica o recebimento das parcelas
                    lErro = CF("Carne_Autentica", colCarneParcelasImpressao)
                    If lErro <> SUCESSO Then gError 109632
                End If
                
        End Select
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr

    Select Case gErr

        Case 109626
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_SELECIONADA", gErr)
            
        Case 109627 To 109632
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166224)

    End Select

    Exit Function

End Function

Private Function Move_Dados_Memoria(colCarneParcelasImpressao As Collection) As Long

Dim iIndice As Integer
Dim objCarneParcelasImpressao As ClassCarneParcelasImpressao
Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objRecebimentoCarneParc As New ClassRecebimentoCarneParc
Dim objBaixasParcCarne As New ClassBaixasParcCarne

On Error GoTo Erro_Move_Dados_Memoria

    'Para cada linha do grid
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        'Se estiver selecionada
        If GridParcelas.TextMatrix(iIndice, igrid_Selecionar_Col) = MARCADO Then
            Set objCarneParcelasImpressao = New ClassCarneParcelasImpressao
            Set objBaixasParcCarne = New ClassBaixasParcCarne
            
            Set objRecebimentoCarneParc = gcolCarneParcelas.Item(iIndice)
            
            objCarneParcelasImpressao.sCodCarne = GridParcelas.TextMatrix(iIndice, igrid_Carne_Col)
            objCarneParcelasImpressao.iParcelaNumero = StrParaInt(GridParcelas.TextMatrix(iIndice, igrid_ParcNumero_Col))
            objCarneParcelasImpressao.dtDataVencParcela = StrParaDate(GridParcelas.TextMatrix(iIndice, igrid_ParcVencimento_Col))
            objCarneParcelasImpressao.dParcelaValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, igrid_ParcValor_Col))
            objCarneParcelasImpressao.dDesconto = StrParaDbl(GridParcelas.TextMatrix(iIndice, igrid_ParcDesconto_Col))
            objCarneParcelasImpressao.dJuros = StrParaDbl(GridParcelas.TextMatrix(iIndice, igrid_ParcJuros_Col))
            objCarneParcelasImpressao.dMulta = StrParaDbl(GridParcelas.TextMatrix(iIndice, igrid_ParcMulta_Col))
            
            objCliente.sNomeReduzido = Cliente.Text
            
            'Le os dados do cliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 109633
        
            'Cliente não existente
            If lErro = 12348 Then gError 109634
                    
            'Guarda os dados do cliente neste obj
            objCarneParcelasImpressao.sCPFCGCCliente = objCliente.sCgc
            objCarneParcelasImpressao.lCodCliente = objCliente.lCodigo
            objCarneParcelasImpressao.sNomeCliente = objCliente.sNomeReduzido
            
            'Guarda os dados da baixa
            objBaixasParcCarne.lNumIntParcelaCarne = objCarneParcelasImpressao.iParcelaNumero
            objBaixasParcCarne.sCodCarne = objCarneParcelasImpressao.sCodCarne
            objBaixasParcCarne.dValorBaixado = objCarneParcelasImpressao.dParcelaValor
            objBaixasParcCarne.dValorMulta = objCarneParcelasImpressao.dMulta
            objBaixasParcCarne.dValorDesconto = objCarneParcelasImpressao.dDesconto
            objBaixasParcCarne.dValorJuros = objCarneParcelasImpressao.dJuros
            
            'Jogo na col
            Set objCarneParcelasImpressao.objBaixasParcCarne = objBaixasParcCarne
            colCarneParcelasImpressao.Add objCarneParcelasImpressao
        End If
    Next
    
    Move_Dados_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Dados_Memoria:
    
    Move_Dados_Memoria = gErr

    Select Case gErr

        Case 109633
        
        Case 109634
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166225)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'verifica se houve alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 109624

    Call Limpa_Tela_Impressaocarne

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 109624

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166226)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Impressaocarne()

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridParcelas)

    LabelCancelarValor.Caption = ""
    LabelSelecionadasQuant.Caption = ""

    iClienteAlterado = 0
    iAlterado = 0

End Sub

Private Sub BotaoSair_Click()

    Unload Me

End Sub

Sub Verifica_Selecionado()

Dim vbMsgRes As VbMsgBoxResult

    If StrParaDbl(LabelSelecionadasQuant.Caption) > 0 Then
        'Envia aviso perguntando se deseja salvar as alterações
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SALVAR_ALTERACOES_CARNE")

        'Salva antes
        If vbMsgRes = vbYes Then Call Gravar_Registro

    End If

End Sub

Private Function Move_Dados_Leitura_Memoria(objRecebimentoCarne As ClassRecebimentoCarne) As Long

Dim objCliente As New ClassCliente
Dim lErro As Long

On Error GoTo Erro_Move_Dados_Leitura_Memoria

    objCliente.sNomeReduzido = Cliente.Text

    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 109620
        
    'Cliente não existente
    If lErro = 12348 Then gError 109621
    
    objRecebimentoCarne.iStatus = STATUS_LANCADO
    
    'Guarda o código do cliente obtido pelo nome reduzido
    objRecebimentoCarne.lCodCliente = objCliente.lCodigo

    'Guarda código inicial do carne
    If Len(Trim(CarneDe.Text)) > 0 Then objRecebimentoCarne.sCodCarneDe = CarneDe.Text

    'Guarda código final do carne
    If Len(Trim(CarneAte.Text)) > 0 Then objRecebimentoCarne.sCodCarneAte = CarneAte.Text

    'Guarda data inicial do carne
    If Len(Trim(VencimentoDe.ClipText)) > 0 Then
        objRecebimentoCarne.dtDataVenctoDe = StrParaDate(VencimentoDe.Text)
    Else
        objRecebimentoCarne.dtDataVenctoDe = DATA_NULA
    End If

    'Guarda data final do carne
    If Len(Trim(VencimentoAte.ClipText)) > 0 Then
        objRecebimentoCarne.dtDataVenctoAte = StrParaDate(VencimentoAte.Text)
    Else
        objRecebimentoCarne.dtDataVenctoAte = DATA_NULA
    End If
    
    Move_Dados_Leitura_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_Leitura_Memoria:

    Move_Dados_Leitura_Memoria = gErr

    Select Case gErr

        Case 109620
        
        Case 109621
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166227)

    End Select

    Exit Function

End Function

Private Function Traz_ParcelasCarne_Tela(objRecebimentoCarne As ClassRecebimentoCarne) As Long

Dim objCarneParcImpressao As New ClassCarneParcelasImpressao
Dim lErro As Long
Dim iIndice As Integer
Dim colCarneParcelas As New Collection
Dim objRecebimentoCarneParc As ClassRecebimentoCarneParc

On Error GoTo Erro_Traz_ParcelasCarne_Tela

    'Lê as parcelas que serão carregadas
    lErro = CF("RecebimentoCarne_Le_Parcelas", objRecebimentoCarne)
    If lErro <> SUCESSO Then gError 109622
    
    'Para cada Parcela da coleção de parcelas
    For iIndice = 1 To objRecebimentoCarne.colParcelas.Count
    
        Set objRecebimentoCarneParc = objRecebimentoCarne.colParcelas(iIndice)
        'Joga no grid
        objGridParcelas.iLinhasExistentes = objGridParcelas.iLinhasExistentes + 1
        
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_Carne_Col) = objRecebimentoCarneParc.sCarneNumero
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcNumero_Col) = objRecebimentoCarneParc.iParcelaNumero
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcValor_Col) = Format(objRecebimentoCarneParc.dParcelaValor, "standard")
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcVencimento_Col) = objRecebimentoCarneParc.dtParcelaVencimento
        
        objRecebimentoCarneParc.lCliente = objRecebimentoCarne.lCodCliente
        
    Next
    
    'Aponta para a coleção de parcelas carregadas
    Set gcolCarneParcelas = objRecebimentoCarne.colParcelas
    
    Traz_ParcelasCarne_Tela = SUCESSO

    Exit Function

Erro_Traz_ParcelasCarne_Tela:

    Traz_ParcelasCarne_Tela = gErr

    Select Case gErr

        Case 109622
            lErro = Rotina_Aviso(vbOK, "AVISO_SEM_REGISTRO")

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166228)

    End Select

    Exit Function

End Function

Private Sub BotaoTrazer_Click()

Dim lErro As Long
Dim objRecebimentoCarne As New ClassRecebimentoCarne

On Error GoTo Erro_BotaoTrazer_Click

    'Se tiver algo selecionado pergunta se deseja salvar
    Call Verifica_Selecionado

    'Se não tem cliente preenchido -->erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 109616
    
    'data final não pode ser menor que data inicial
    If Len(Trim(VencimentoAte.ClipText)) <> 0 And Len(Trim(VencimentoDe.ClipText)) <> 0 Then
        If StrParaDate(VencimentoDe.Text) > StrParaDate(VencimentoAte.Text) Then gError 109617
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    Call Grid_Limpa(objGridParcelas)
    
    LabelCancelarValor.Caption = ""
    LabelSelecionadasQuant.Caption = ""
    
    'Guarda na memória os dados que serão utilizados para ler as parcelas
    lErro = Move_Dados_Leitura_Memoria(objRecebimentoCarne)
    If lErro <> SUCESSO Then gError 109618

    'Traz as parcelas para a tela
    lErro = Traz_ParcelasCarne_Tela(objRecebimentoCarne)
    If lErro <> SUCESSO Then gError 109619

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoTrazer_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 109616
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO1", gErr, Error$)

        Case 109618, 109619
        
        Case 109617
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166229)

    End Select

    Exit Sub

End Sub

Private Sub CarneDe_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CarneDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CarneDe)

End Sub



Private Sub CarneAte_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CarneAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CarneAte)

End Sub

Private Sub Cliente_Change()
'Determina se Houve Mudança

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cliente_GotFocus()

    gsCliente = Cliente.Text
    
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cliente_Validate
    
    'Se não tem cliente preenchido
    If Len(Trim(Cliente.Text)) = 0 Then
        'Se não tem nada na tela-->sai
        If objGridParcelas.iLinhasExistentes = 0 Then Exit Sub
        'Pergunta se o usuário deseja limpar a tela
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LIMPAR_TELA")
        'Se deseja
        If vbMsgRes = vbYes Then
            Call Limpa_Tela_Impressaocarne
        Else
            'Retorna o cliente anterior
            Cliente.Text = gsCliente
        End If
        iClienteAlterado = 0
        Exit Sub
    End If
    
    'Se o cliente não foi alterado -->sai
    If iClienteAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Passa o cliente com o código, nome reduzido ou CPF/CGC
    lErro = TP_Cliente_Le(Cliente, objCliente, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 109575
    
    If objGridParcelas.iLinhasExistentes = 0 Then Exit Sub
    
    If gsCliente = Cliente.Text Then Exit Sub
    
    'Verifica se deseja trazer os novos dados para a tela
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NOVAS_CONFIGURACOES")
    'Se deseja
    If vbMsgRes = vbYes Then
        Call BotaoTrazer_Click
    Else
        'Retorna o cliente anterior
        Cliente.Text = gsCliente
    End If
        
    iClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr
        
        Case 109575
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166230)
                        
    End Select
   
    Exit Sub

End Sub

Private Sub LabelCarneDe_Click()

Dim objCarne As New ClassCarne
Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCarneDe_Click
    
    gbCarneDe = True
    
    'Se não tem cliente preenchido -->erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 109576
    
    objCliente.sNomeReduzido = Cliente.Text
    
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 109577
        
    'Cliente não existente
    If lErro = 12348 Then gError 109578
    
    sSelecao = "Cliente = ?"

    colSelecao.Add StrParaInt(objCliente.lCodigo)
            
    'Chama Tela CarneLista
    Call Chama_Tela("CarneLojaLista", colSelecao, objCarne, objEventoCarne, sSelecao)
        
    If giRetornoTela = vbOK Then
        CarneDe.PromptInclude = False
        CarneDe.Text = objCarne.sCodBarrasCarne
        CarneDe.PromptInclude = True
    End If
    
    Exit Sub

Erro_LabelCarneDe_Click:

    Select Case gErr

        Case 109576
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO1", gErr, Error$)
        
        Case 109577
        
        Case 109578
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166231)

    End Select

    Exit Sub
        
End Sub

Private Sub LabelCarneAte_Click()

Dim objCarne As New ClassCarne
Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCarneAte_Click
    
    gbCarneDe = False
    
    'Se não tem cliente preenchido -->erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 109579
    
    objCliente.sNomeReduzido = Cliente.Text
    
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 109580
        
    'Cliente não existente
    If lErro = 12348 Then gError 109581
        
    sSelecao = "Cliente = ? AND Status <> 2"

    colSelecao.Add StrParaInt(objCliente.lCodigo)
            
    'Chama Tela CarneLista
    Call Chama_Tela("CarneLojaLista", colSelecao, objCarne, objEventoCarne, sSelecao)
        
    If giRetornoTela = vbOK Then
        CarneAte.PromptInclude = False
        CarneAte.Text = objCarne.sCodBarrasCarne
        CarneAte.PromptInclude = True
    End If
    
    Exit Sub

Erro_LabelCarneAte_Click:

    Select Case gErr

        Case 109579
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr, Error$)
        
        Case 109580
        
        Case 109581
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166232)

    End Select

    Exit Sub
        
End Sub

Private Sub objEventoCarne_evSelecao(obj1 As Object)

Dim objCarne As ClassCarne

    Set objCarne = obj1
    
    If gbCarneDe Then
        CarneDe.Text = objCarne.sCodBarrasCarne
    Else
        CarneAte.Text = objCarne.sCodBarrasCarne
    End If
    
    Me.Show

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1
    
    Cliente.Text = objCliente.lCodigo
    
    Call Cliente_Validate(False)
    
    Me.Show

    Exit Sub

End Sub

Private Sub VencimentoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(VencimentoDe, iAlterado)

End Sub

Private Sub VencimentoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencimentoDe_Validate

    'se a data não estiver preenchida-> sai
    If Len(Trim(VencimentoDe.ClipText)) = 0 Then Exit Sub
    
    'critica a data
    lErro = Data_Critica(VencimentoDe.Text)
    If lErro <> SUCESSO Then gError 109582

    Exit Sub

Erro_VencimentoDe_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109582
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166233)

    End Select

    Exit Sub

End Sub

Private Sub VencimentoDe_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownVencimentoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoDe_DownClick

    'diminui a data
    lErro = Data_Up_Down_Click(VencimentoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 109584
    
    Exit Sub
    
Erro_UpDownVencimentoDe_DownClick:
    
    Select Case gErr
    
        Case 109584
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166234)

    End Select
    
    Exit Sub

End Sub

Private Sub UpDownVencimentoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoDe_UpClick

    'diminui a data
    lErro = Data_Up_Down_Click(VencimentoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 109585
    
    Exit Sub
    
Erro_UpDownVencimentoDe_UpClick:
    
    Select Case gErr
    
        Case 109585
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166235)

    End Select
    
    Exit Sub

End Sub

Private Sub VencimentoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(VencimentoAte, iAlterado)

End Sub

Private Sub VencimentoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VencimentoAte_Validate

    'se a data não estiver preenchida-> sai
    If Len(Trim(VencimentoAte.ClipText)) = 0 Then Exit Sub
    
    'critica a data
    lErro = Data_Critica(VencimentoAte.Text)
    If lErro <> SUCESSO Then gError 109586

    Exit Sub

Erro_VencimentoAte_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 109586
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166236)

    End Select

    Exit Sub

End Sub

Private Sub VencimentoAte_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownVencimentoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoAte_DownClick

    'diminui a data
    lErro = Data_Up_Down_Click(VencimentoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 109587
    
    Exit Sub
    
Erro_UpDownVencimentoAte_DownClick:
    
    Select Case gErr
    
        Case 109587
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166237)

    End Select
    
    Exit Sub

End Sub

Private Sub UpDownVencimentoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimentoAte_UpClick

    'diminui a data
    lErro = Data_Up_Down_Click(VencimentoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 109588
    
    Exit Sub
    
Erro_UpDownVencimentoAte_UpClick:
    
    Select Case gErr
    
        Case 109588
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166238)

    End Select
    
    Exit Sub

End Sub

Private Sub GridParcelas_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_EnterCell()

    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Private Sub GridParcelas_LostFocus()

    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Private Sub Selecionar_Click()
    
    Call Recalcula_Totais
    
End Sub

Private Sub Recalcula_Totais()

Dim iIndice As Integer
Dim iQuant As Integer
Dim dValor As Double
    
    iQuant = 0
    dValor = 0
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        If GridParcelas.TextMatrix(iIndice, igrid_Selecionar_Col) = MARCADO Then
            iQuant = iQuant + 1
            dValor = dValor + StrParaDbl(GridParcelas.TextMatrix(iIndice, igrid_ParcValor_Col))
        End If
    Next
    
    LabelSelecionadasQuant.Caption = iQuant
    LabelCancelarValor.Caption = Format(dValor, "standard")
        
End Sub

Private Sub Selecionar_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Selecionar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Selecionar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub CarneNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CarneNumero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub CarneNumero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcelaNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelaNumero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcelaNumero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcelaVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelaVencimento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcelaVencimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcelaValor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelaValor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcelaValor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcelaMulta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelaMulta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcelaMulta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcelaJuros_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelaJuros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcelaJuros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcelaDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelaDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcelaDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        If objGridInt.objGrid.Name = GridParcelas.Name Then
            lErro = Saida_Celula_GridParcelas(objGridInt)
            If lErro <> SUCESSO Then gError 109589
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 109590

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 109589, 109590

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166239)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas
    
    Select Case objGridInt.objGrid.Col
    
        Case igrid_Carne_Col
            lErro = Saida_Celula_Carne(objGridInt)
            If lErro <> SUCESSO Then gError 109591
        Case igrid_ParcDesconto_Col
            lErro = Saida_Celula_ParcDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 109592
        Case igrid_ParcJuros_Col
            lErro = Saida_Celula_ParcJuros(objGridInt)
            If lErro <> SUCESSO Then gError 109593
        Case igrid_ParcMulta_Col
            lErro = Saida_Celula_ParcMulta(objGridInt)
            If lErro <> SUCESSO Then gError 109594
        Case igrid_ParcNumero_Col
            lErro = Saida_Celula_ParcNumero(objGridInt)
            If lErro <> SUCESSO Then gError 109595
        Case igrid_ParcValor_Col
            lErro = Saida_Celula_ParcValor(objGridInt)
            If lErro <> SUCESSO Then gError 109596
        Case igrid_ParcVencimento_Col
            lErro = Saida_Celula_ParcVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 109597
        Case igrid_Selecionar_Col
            lErro = Saida_Celula_Selecionar(objGridInt)
            If lErro <> SUCESSO Then gError 109598
            
    End Select
            
    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 109591 To 109598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166240)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Carne(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Carne

    Set objGridInt.objControle = CarneNumero
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109599

    Saida_Celula_Carne = SUCESSO

    Exit Function

Erro_Saida_Celula_Carne:

    Saida_Celula_Carne = gErr

    Select Case gErr
        
        Case 109599
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166241)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcDesconto(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcDesconto

    Set objGridInt.objControle = ParcelaDesconto

    'se o valor estiver preenchido
    If Len(Trim(ParcelaDesconto.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ParcelaDesconto.Text))
        If lErro <> SUCESSO Then gError 109600
        
    End If

    ParcelaDesconto.Text = Format(ParcelaDesconto.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109601
    
    Saida_Celula_ParcDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcDesconto:

    Saida_Celula_ParcDesconto = gErr

    Select Case gErr

        Case 109600, 109601
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166242)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcJuros(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcJuros

    Set objGridInt.objControle = ParcelaJuros

    'se o valor estiver preenchido
    If Len(Trim(ParcelaJuros.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ParcelaJuros.Text))
        If lErro <> SUCESSO Then gError 109602
        
    End If

    ParcelaJuros.Text = Format(ParcelaJuros.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109603
    
    Saida_Celula_ParcJuros = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcJuros:

    Saida_Celula_ParcJuros = gErr

    Select Case gErr

        Case 109602, 109603
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166243)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcMulta(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcMulta

    Set objGridInt.objControle = ParcelaMulta

    'se o valor estiver preenchido
    If Len(Trim(ParcelaMulta.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ParcelaMulta.Text))
        If lErro <> SUCESSO Then gError 109604
        
    End If

    ParcelaMulta.Text = Format(ParcelaMulta.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109605
    
    Saida_Celula_ParcMulta = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcMulta:

    Saida_Celula_ParcMulta = gErr

    Select Case gErr

        Case 109604, 109605
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166244)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcValor(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcValor

    Set objGridInt.objControle = ParcelaValor

    'se o valor estiver preenchido
    If Len(Trim(ParcelaValor.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ParcelaValor.Text))
        If lErro <> SUCESSO Then gError 109606
        
    End If

    ParcelaValor.Text = Format(ParcelaValor.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109607
    
    Saida_Celula_ParcValor = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcValor:

    Saida_Celula_ParcValor = gErr

    Select Case gErr

        Case 109606, 109607
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166245)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcNumero(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcNumero

    Set objGridInt.objControle = ParcelaNumero

    'se o valor estiver preenchido
    If Len(Trim(ParcelaNumero.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ParcelaNumero.Text))
        If lErro <> SUCESSO Then gError 109608
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109609
    
    Saida_Celula_ParcNumero = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcNumero:

    Saida_Celula_ParcNumero = gErr

    Select Case gErr

        Case 109608, 109609
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166246)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcVencimento(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcVencimento

    Set objGridInt.objControle = ParcelaVencimento

    'se a data estiver preenchida
    If Len(Trim(ParcelaVencimento.Text)) <> 0 Then

        'verifica se o a data é valida
        lErro = Data_Critica(ParcelaVencimento.Text)
        If lErro <> SUCESSO Then gError 109610
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109611
    
    Saida_Celula_ParcVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcVencimento:

    Saida_Celula_ParcVencimento = gErr

    Select Case gErr

        Case 109610, 109611
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166247)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Selecionar(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionar

    Set objGridInt.objControle = Selecionar

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109612
    
    Saida_Celula_Selecionar = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionar:

    Saida_Celula_Selecionar = gErr

    Select Case gErr

        Case 109612
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166248)

    End Select

    Exit Function

End Function

Private Sub MSComm1_OnComm()
  
Dim objCarne As ClassCarne
Dim objCliente As New ClassCliente
Dim sCarne As String
Dim sCliente As String
Dim iParcela As String
Dim bAchou As Boolean
Dim objRecebimentoCarne As New ClassRecebimentoCarne
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_MSComm1_OnComm

    'Usado para verificar se tem buffer
    If MSComm1.CommEvent <> 2 Then Exit Sub
    
    'Joga o imput numa variável
    sCarne = Mid(MSComm1.Input, 1, 20)
    
    'Se for o código de uma parcela --> guarda a parcela
    If MSComm1.Input > 20 Then iParcela = StrParaInt(Mid(MSComm1.Input, 20, 21))
    
    bAchou = False
    
    'Acha o cliente para este carne
    For Each objCarne In gcolCarne
        If objCarne.sCodBarrasCarne = sCarne Then
            objCliente.lCodigo = objCarne.lCliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO Then gError 109613
            sCliente = objCliente.sNomeReduzido
            bAchou = True
        End If
    Next
    
    'pode não achar o carne
    If Not (bAchou) Then gError 109614
    
    'SE tiver algo na tela e for de cliente diferente-->limpa a tela
    If objGridParcelas.iLinhasExistentes > 0 Then
        If Cliente.Text <> sCliente Then Call Limpa_Tela_Impressaocarne
    End If
    
    'Jogo o cliente na tela
    Cliente.Text = sCliente
    
    objRecebimentoCarne.lCodCliente = objCliente.lCodigo
    objRecebimentoCarne.dtDataVenctoAte = DATA_NULA
    objRecebimentoCarne.dtDataVenctoDe = DATA_NULA
    objRecebimentoCarne.sCodCarneDe = sCarne
    
    'Trago o carne para a tela
    lErro = Traz_ParcelasCarne_Tela(objRecebimentoCarne)
    If lErro <> SUCESSO Then gError 109615
    
    'Se não existir parcela --> sai
    If iParcela = 0 Then Exit Sub
    
    'Seleciona no grid a parcela
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        If GridParcelas.TextMatrix(igrid_Carne_Col, iIndice) = sCarne And GridParcelas.TextMatrix(igrid_ParcNumero_Col, iIndice) = iParcela Then GridParcelas.TextMatrix(igrid_Selecionar_Col, iIndice) = MARCADO
    Next
    
    Exit Sub

Erro_MSComm1_OnComm:

    Select Case gErr

        Case 109613, 109615
            
        Case 109614
            Call Rotina_Erro(vbOKOnly, "ERRO_CARNE_NAO_EXISTENTE", gErr, Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166249)

    End Select

    Exit Sub

End Sub

Public Sub form_unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    Set objGridParcelas = Nothing
    Set objEventoCliente = Nothing
    Set gcolCarneParcelas = Nothing
    
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
   'Clique em F3
    If KeyCode = vbKeyF3 Then
        Call LabelCliente_Click
    End If
    
    'Clique em F4
    If KeyCode = vbKeyF4 Then
        Call BotaoTrazer_Click
    End If
    
    'Clique em F7
    If KeyCode = vbKeyF7 Then
        Call BotaoLimpar_Click
    End If
    

End Sub


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Recebimento de Carnê"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RecebimentoCarne"
    
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
