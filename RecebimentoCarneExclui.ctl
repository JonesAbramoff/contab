VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RecebimentoCarneExclui 
   ClientHeight    =   5940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8160
   Begin VB.CommandButton BotaoLimpar 
      Height          =   600
      Left            =   3270
      Picture         =   "RecebimentoCarneExclui.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1560
   End
   Begin VB.Frame FrameDadosBaixa 
      Caption         =   "Dados da Baixa"
      Height          =   2280
      Left            =   180
      TabIndex        =   13
      Top             =   990
      Width           =   7815
      Begin VB.Frame FrameValores 
         Caption         =   "Valores"
         Height          =   1695
         Left            =   4440
         TabIndex        =   14
         Top             =   300
         Width           =   3135
         Begin VB.Label LabelBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Baixa:"
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
            TabIndex        =   22
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabelBaixaValor 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label LabelDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            Left            =   1750
            TabIndex        =   20
            Top             =   240
            Width           =   885
         End
         Begin VB.Label LabelDescontoValor 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   1750
            TabIndex        =   19
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label LabelMulta 
            AutoSize        =   -1  'True
            Caption         =   "Multa:"
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
            TabIndex        =   18
            Top             =   960
            Width           =   540
         End
         Begin VB.Label LabelMultaValor 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label LabelJuros 
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
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
            Left            =   1750
            TabIndex        =   16
            Top             =   960
            Width           =   525
         End
         Begin VB.Label LabelJurosValor 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   1750
            TabIndex        =   15
            Top             =   1200
            Width           =   1125
         End
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Codigo 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   420
         Width           =   1005
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
         Height          =   195
         Left            =   300
         TabIndex        =   26
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label ClienteNome 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   25
         Top             =   1020
         Width           =   3075
      End
      Begin VB.Label LabelDataBaixa 
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
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label DataBaixa 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1095
         TabIndex        =   23
         Top             =   1635
         Width           =   1005
      End
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
      Left            =   5355
      Picture         =   "RecebimentoCarneExclui.ctx":2C22
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   1560
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
      Left            =   1200
      Picture         =   "RecebimentoCarneExclui.ctx":5724
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1560
   End
   Begin VB.Frame FrameParcelasBaixadas 
      Caption         =   "Parcelas baixadas"
      Height          =   2445
      Left            =   195
      TabIndex        =   0
      Top             =   3390
      Width           =   7815
      Begin VB.TextBox ValorMulta 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   5370
         TabIndex        =   9
         Top             =   690
         Width           =   840
      End
      Begin VB.TextBox ValorJuros 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   6090
         TabIndex        =   8
         Top             =   1050
         Width           =   840
      End
      Begin VB.TextBox ValorDesconto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   4815
         TabIndex        =   7
         Top             =   930
         Width           =   840
      End
      Begin VB.TextBox ParcelaValor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   3150
         TabIndex        =   5
         Top             =   915
         Width           =   870
      End
      Begin VB.TextBox ParcelaNumero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   3
         Top             =   930
         Width           =   735
      End
      Begin VB.TextBox CarneNumero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   105
         TabIndex        =   2
         Top             =   600
         Width           =   1890
      End
      Begin VB.TextBox ValorBaixado 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   3870
         TabIndex        =   6
         Top             =   630
         Width           =   1080
      End
      Begin MSMask.MaskEdBox ParcelaVencimento 
         Height          =   300
         Left            =   2640
         TabIndex        =   4
         Top             =   510
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   2040
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3598
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "RecebimentoCarneExclui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variáveis Globáis

Public iAlterado  As Integer
Dim iClienteAlterado  As Integer
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

'variáveis para grid
Dim objGridParcelas As AdmGrid
Dim igrid_Carne_Col As Integer
Dim igrid_ParcNumero_Col As Integer
Dim igrid_ParcVencimento_Col As Integer
Dim igrid_ParcvalorBaixado_Col As Integer
Dim igrid_ParcValor_Col As Integer
Dim igrid_ParcJuros_Col As Integer
Dim igrid_ParcMulta_Col As Integer
Dim igrid_ParcDesconto_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Trata_Parametros(Optional objbaixasCarne As ClassBaixasCarne) As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    Set objGridParcelas = New AdmGrid
    Set objEventoCodigo = New AdmEvento

    Call Inicializa_GridParcelas(objGridParcelas)

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166262)

    End Select

    Exit Sub

End Sub

Private Sub Inicializa_GridParcelas(objGridInt As AdmGrid)

On Error GoTo Erro_Inicializa_GridParcelas

    'form do grid
    Set objGridInt.objForm = Me

    'Títulos das Colunas
    objGridInt.colColuna.Add ""
    objGridInt.colColuna.Add "Carne"
    objGridInt.colColuna.Add "Parcela"
    objGridInt.colColuna.Add "Vencimento"
    objGridInt.colColuna.Add "Valor"
    objGridInt.colColuna.Add "Valor Baixado"
    objGridInt.colColuna.Add "Desconto"
    objGridInt.colColuna.Add "Multa"
    objGridInt.colColuna.Add "Juros"

    'Controles que participam do Grid
    objGridInt.colCampo.Add CarneNumero.Name
    objGridInt.colCampo.Add ParcelaNumero.Name
    objGridInt.colCampo.Add ParcelaVencimento.Name
    objGridInt.colCampo.Add ParcelaValor.Name
    objGridInt.colCampo.Add ValorBaixado.Name
    objGridInt.colCampo.Add ValorDesconto.Name
    objGridInt.colCampo.Add ValorMulta.Name
    objGridInt.colCampo.Add ValorJuros.Name

    'Colunas do Grid
    igrid_Carne_Col = 1
    igrid_ParcNumero_Col = 2
    igrid_ParcVencimento_Col = 3
    igrid_ParcValor_Col = 4
    igrid_ParcvalorBaixado_Col = 5
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166263)

    End Select

    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se nõao existe linhas no grid --> erro.
    If objGridParcelas.iLinhasExistentes = 0 Then gError 109780

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 109739

    'Limpa a tela
    Call RecebimentoCarneExclui_Limpa

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 109739

        Case 109780
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_NAO_SELECIONADA", gErr, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166264)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objbaixasCarne As New ClassBaixasCarne
Dim colCarneAtualizados As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    objbaixasCarne.lCodigo = StrParaLong(Codigo.Caption)
    objbaixasCarne.iFilialEmpresa = giFilialEmpresa

    'faz a gravação dos dados do cancelamento
    lErro = CF("RecebimentoCarne_Exclui", objbaixasCarne, colCarneAtualizados)
    If lErro <> SUCESSO Then gError 109743

    'Imprime o carnê
    '???lErro = CF("Caixa_Carne_Imprime_ECF", colCarneAtualizados)
    If lErro <> SUCESSO Then gError 109744

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr

    Select Case gErr

        Case 109743, 109744

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166265)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'verifica se houve alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 109738

    Call RecebimentoCarneExclui_Limpa

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 109738

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166266)

    End Select

    Exit Sub

End Sub

Private Sub RecebimentoCarneExclui_Limpa()

    Call Limpa_Tela(Me)

    Call Grid_Limpa(objGridParcelas)

    Call ComandoSeta_Fechar(Me.Name)

    Codigo.Caption = ""
    ClienteNome.Caption = ""
    DataBaixa.Caption = ""
    LabelBaixaValor.Caption = ""
    LabelMultaValor.Caption = ""
    LabelDescontoValor.Caption = ""
    LabelJurosValor.Caption = ""

    iClienteAlterado = 0
    iAlterado = 0

End Sub

Private Sub BotaoSair_Click()

    Unload Me

End Sub

Private Function Traz_BaixasCarne(objbaixasCarne As ClassBaixasCarne) As Long

Dim objCarneParcelasImpressao As ClassCarneParcelasImpressao
Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente

On Error GoTo Erro_Traz_BaixasCarne

    Call RecebimentoCarneExclui_Limpa

    'Le os dados do recebimento do carnet
    lErro = CF("BaixasCarneDetalhes_Le", objbaixasCarne)
    If lErro <> SUCESSO And lErro <> 109704 Then gError 109703

    'se não achou nenhuma parcela --> erro.
    If lErro = 109704 Then gError 109705

    'Para cada Parcela da coleção de parcelas
    For iIndice = 1 To objbaixasCarne.colParcelas.Count

        Set objCarneParcelasImpressao = objbaixasCarne.colParcelas(iIndice)
        'Joga no grid
        objGridParcelas.iLinhasExistentes = objGridParcelas.iLinhasExistentes + 1

        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_Carne_Col) = objCarneParcelasImpressao.sCodCarne
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcNumero_Col) = objCarneParcelasImpressao.iParcelaNumero
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcValor_Col) = Format(objCarneParcelasImpressao.dParcelaValor, "standard")
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcVencimento_Col) = objCarneParcelasImpressao.dtDataVencParcela
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcvalorBaixado_Col) = Format(objCarneParcelasImpressao.dValorBaixado, "standard")
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcDesconto_Col) = Format(objCarneParcelasImpressao.dDesconto, "standard")
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcJuros_Col) = Format(objCarneParcelasImpressao.dJuros, "standard")
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, igrid_ParcMulta_Col) = Format(objCarneParcelasImpressao.dMulta, "standard")

        LabelBaixaValor.Caption = Format(StrParaDbl(LabelBaixaValor.Caption) + objCarneParcelasImpressao.dValorBaixado, "standard")
        LabelDescontoValor.Caption = Format(StrParaDbl(LabelDescontoValor.Caption) + objCarneParcelasImpressao.dDesconto, "standard")
        LabelMultaValor.Caption = Format(StrParaDbl(LabelMultaValor.Caption) + objCarneParcelasImpressao.dMulta, "standard")
        LabelJurosValor.Caption = Format(StrParaDbl(LabelJurosValor.Caption) + objCarneParcelasImpressao.dJuros, "standard")
        objCliente.lCodigo = objCarneParcelasImpressao.lCodCliente

    Next

    lErro = CF("Cliente_Le", objCliente)

    Codigo.Caption = objbaixasCarne.lCodigo
    DataBaixa.Caption = objbaixasCarne.dtDataBaixa
    ClienteNome.Caption = objCliente.sNomeReduzido

    Traz_BaixasCarne = SUCESSO

    Exit Function

Erro_Traz_BaixasCarne:

    Traz_BaixasCarne = gErr

    Select Case gErr

        Case 109703

        Case 109705
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_ENCONTRADA", gErr, objbaixasCarne.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166267)

    End Select

    Exit Function

End Function

Private Sub LabelCodigo_Click()

Dim objbaixasCarne As New ClassBaixasCarne
Dim colSelecao As Collection

    'Chama Tela CarneLista passando parametros para pegar somente os carnews baixados
    Call Chama_Tela("BaixasCarneLista", colSelecao, objbaixasCarne, objEventoCodigo)

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objbaixasCarne As ClassBaixasCarne
Dim objCarneParcelasImpressao As ClassCarneParcelasImpressao
Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objbaixasCarne = obj1

    lErro = Traz_BaixasCarne(objbaixasCarne)
    If lErro <> SUCESSO Then gError 109710

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 109710

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166268)

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

Private Sub ValorBaixado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorBaixado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorBaixado_KeyPress(KeyAscii As Integer)

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

Private Sub ValorMulta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorMulta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorMulta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorJuros_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorJuros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorJuros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        If objGridInt.objGrid.Name = GridParcelas.Name Then
            lErro = Saida_Celula_GridParcelas(objGridInt)
            If lErro <> SUCESSO Then gError 109711
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 109712

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 109711, 109712

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166269)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    Select Case objGridInt.objGrid.Col

        Case igrid_Carne_Col
            lErro = Saida_Celula_Carne(objGridInt)
            If lErro <> SUCESSO Then gError 109713
        Case igrid_ParcDesconto_Col
            lErro = Saida_Celula_ParcDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 109714
        Case igrid_ParcJuros_Col
            lErro = Saida_Celula_ParcJuros(objGridInt)
            If lErro <> SUCESSO Then gError 109715
        Case igrid_ParcMulta_Col
            lErro = Saida_Celula_ParcMulta(objGridInt)
            If lErro <> SUCESSO Then gError 109716
        Case igrid_ParcNumero_Col
            lErro = Saida_Celula_ParcNumero(objGridInt)
            If lErro <> SUCESSO Then gError 109717
        Case igrid_ParcValor_Col
            lErro = Saida_Celula_ParcValor(objGridInt)
            If lErro <> SUCESSO Then gError 109718
        Case igrid_ParcVencimento_Col
            lErro = Saida_Celula_ParcVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 109719
        Case igrid_ParcvalorBaixado_Col
            lErro = Saida_Celula_ValorBaixado(objGridInt)
            If lErro <> SUCESSO Then gError 109720

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 109713 To 109720

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166270)

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
    If lErro <> SUCESSO Then gError 109721

    Saida_Celula_Carne = SUCESSO

    Exit Function

Erro_Saida_Celula_Carne:

    Saida_Celula_Carne = gErr

    Select Case gErr

        Case 109721
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166271)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcDesconto(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcDesconto

    Set objGridInt.objControle = ValorDesconto

    'se o valor estiver preenchido
    If Len(Trim(ValorDesconto.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ValorDesconto.Text))
        If lErro <> SUCESSO Then gError 109722

    End If

    ValorDesconto.Text = Format(ValorDesconto.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109723

    Saida_Celula_ParcDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcDesconto:

    Saida_Celula_ParcDesconto = gErr

    Select Case gErr

        Case 109722, 109723
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166272)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcJuros(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcJuros

    Set objGridInt.objControle = ValorJuros

    'se o valor estiver preenchido
    If Len(Trim(ValorJuros.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ValorJuros.Text))
        If lErro <> SUCESSO Then gError 109724

    End If

    ValorJuros.Text = Format(ValorJuros.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109725

    Saida_Celula_ParcJuros = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcJuros:

    Saida_Celula_ParcJuros = gErr

    Select Case gErr

        Case 109724, 109725
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166273)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcMulta(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcMulta

    Set objGridInt.objControle = ValorMulta

    'se o valor estiver preenchido
    If Len(Trim(ValorMulta.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ValorMulta.Text))
        If lErro <> SUCESSO Then gError 109726

    End If

    ValorMulta.Text = Format(ValorMulta.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109727

    Saida_Celula_ParcMulta = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcMulta:

    Saida_Celula_ParcMulta = gErr

    Select Case gErr

        Case 109726, 109727
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166274)

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
        If lErro <> SUCESSO Then gError 109728

    End If

    ParcelaValor.Text = Format(ParcelaValor.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109729

    Saida_Celula_ParcValor = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcValor:

    Saida_Celula_ParcValor = gErr

    Select Case gErr

        Case 109728, 109729
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166275)

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
        If lErro <> SUCESSO Then gError 109730

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109731

    Saida_Celula_ParcNumero = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcNumero:

    Saida_Celula_ParcNumero = gErr

    Select Case gErr

        Case 109730, 109731
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166276)

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
        If lErro <> SUCESSO Then gError 109732

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109733

    Saida_Celula_ParcVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcVencimento:

    Saida_Celula_ParcVencimento = gErr

    Select Case gErr

        Case 109732, 109733
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166277)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorBaixado(objGridInt As AdmGrid) As Long

Dim dTotal As Double
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorBaixado

    Set objGridInt.objControle = ValorBaixado

    'se o valor estiver preenchido
    If Len(Trim(ValorBaixado.Text)) <> 0 Then

        'verifica se o valor digitado é positivo
        lErro = Valor_Positivo_Critica(Trim(ValorBaixado.Text))
        If lErro <> SUCESSO Then gError 109734

    End If

    ValorBaixado.Text = Format(ValorBaixado.Text, "STANDARD")

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 109735

    Saida_Celula_ValorBaixado = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBaixado:

    Saida_Celula_ValorBaixado = gErr

    Select Case gErr

        Case 109734, 109735
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166278)

    End Select

    Exit Function

End Function

Public Sub form_unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    Set objGridParcelas = Nothing
    Set objEventoCodigo = Nothing

    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   'Clique em F3
    If KeyCode = vbKeyF3 Then
        Call LabelCodigo_Click
    End If

    'Clique em F7
    If KeyCode = vbKeyF7 Then
        Call BotaoLimpar_Click
    End If

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objbaixasCarne As New ClassBaixasCarne

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "BaixasCarne"

    objbaixasCarne.lCodigo = StrParaLong(Codigo.Caption)

    'Preenche a coleção colCampoValor, com nome do campo,
    colCampoValor.Add "Codigo", objbaixasCarne.lCodigo, 0, "Codigo"
    colCampoValor.Add "DataBaixa", objbaixasCarne.dtDataBaixa, 0, "DataBaixa"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166279)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objbaixasCarne As New ClassBaixasCarne

On Error GoTo Erro_Tela_Preenche

    objbaixasCarne.lCodigo = colCampoValor.Item("Codigo").vValor

    If objbaixasCarne.lCodigo > 0 Then

        'Carrega objBaixascarne com os dados passados em colCampoValor
        objbaixasCarne.dtDataBaixa = colCampoValor.Item("DataBaixa").vValor

        'Traz dados para a Tela
        lErro = Traz_BaixasCarne(objbaixasCarne)
        If lErro <> SUCESSO Then gError 109737

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 109737
        'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166280)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Cancela Recebimento de Carne"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RecebimentoCarneExclui"

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

