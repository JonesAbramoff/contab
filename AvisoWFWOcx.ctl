VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AvisoWFWOcx 
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   ScaleHeight     =   6015
   ScaleWidth      =   8580
   Begin VB.Frame Frame3 
      Caption         =   "Opções"
      Height          =   1020
      Left            =   105
      TabIndex        =   21
      Top             =   4920
      Width           =   8415
      Begin VB.CommandButton BotaoCancelar 
         Caption         =   "Cancelar"
         Height          =   525
         Left            =   7455
         Picture         =   "AvisoWFWOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   345
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton BotaoOK 
         Caption         =   "Continuar"
         Height          =   525
         Left            =   6525
         Picture         =   "AvisoWFWOcx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   345
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton BotaoExcluir 
         Caption         =   "Excluir Item"
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
         Left            =   90
         TabIndex        =   25
         Top             =   210
         Width           =   1620
      End
      Begin VB.CommandButton BotaoExcluirTudo 
         Caption         =   "Excluir Tudo"
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
         Left            =   1980
         TabIndex        =   24
         Top             =   195
         Width           =   1605
      End
      Begin VB.CommandButton BotaoLembrar 
         Caption         =   "Lembrar Novamente Em"
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
         Left            =   90
         TabIndex        =   23
         Top             =   585
         Width           =   2370
      End
      Begin VB.ComboBox TipoIntervalo 
         Height          =   315
         ItemData        =   "AvisoWFWOcx.ctx":025C
         Left            =   3390
         List            =   "AvisoWFWOcx.ctx":026C
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   600
         Width           =   1245
      End
      Begin MSMask.MaskEdBox Intervalo 
         Height          =   315
         Left            =   2640
         TabIndex        =   26
         Top             =   600
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalhe"
      Height          =   2145
      Left            =   105
      TabIndex        =   10
      Top             =   2775
      Width           =   8415
      Begin VB.TextBox Msg1 
         BackColor       =   &H8000000F&
         Height          =   1050
         Left            =   1110
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   225
         Visible         =   0   'False
         Width           =   7185
      End
      Begin VB.Label LabelMsg 
         AutoSize        =   -1  'True
         Caption         =   "Aviso:"
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
         TabIndex        =   20
         Top             =   255
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label LabelData 
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
         Left            =   540
         TabIndex        =   19
         Top             =   1395
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label LabelHora 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Left            =   2535
         TabIndex        =   18
         Top             =   1395
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Data1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1110
         TabIndex        =   17
         Top             =   1350
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Hora1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3045
         TabIndex        =   16
         Top             =   1335
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label LabelUsuario 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   5355
         TabIndex        =   15
         Top             =   1395
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Usuario1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6180
         TabIndex        =   14
         Top             =   1335
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label LabelTransacao 
         AutoSize        =   -1  'True
         Caption         =   "Transação:"
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
         Left            =   45
         TabIndex        =   13
         Top             =   1785
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Transacao1 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1110
         TabIndex        =   12
         Top             =   1725
         Visible         =   0   'False
         Width           =   7185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Avisos"
      Height          =   2670
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   8415
      Begin VB.TextBox Msg 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   195
         TabIndex        =   9
         Top             =   510
         Width           =   3240
      End
      Begin VB.TextBox Hora 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4305
         TabIndex        =   8
         Top             =   1005
         Width           =   735
      End
      Begin VB.TextBox Data 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4260
         TabIndex        =   7
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox Usuario 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4650
         TabIndex        =   6
         Top             =   390
         Width           =   900
      End
      Begin VB.TextBox Transacao 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3885
         TabIndex        =   5
         Top             =   1350
         Width           =   1785
      End
      Begin VB.TextBox NumIntDoc 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4380
         TabIndex        =   4
         Top             =   1575
         Width           =   900
      End
      Begin MSFlexGridLib.MSFlexGrid GridAviso 
         Height          =   2070
         Left            =   105
         TabIndex        =   1
         Top             =   465
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   3651
         _Version        =   393216
         Rows            =   8
         Cols            =   5
         AllowBigSelection=   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin VB.Label LabelLembrar 
         AutoSize        =   -1  'True
         Caption         =   "Lembrar Novamente em:"
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
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Label Intervalo1 
         Height          =   195
         Left            =   2265
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
End
Attribute VB_Name = "AvisoWFWOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Const KEYCODE_VERIFICAR_SINTAXE = vbKeyF5

Dim iGrid_Msg_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Hora_Col As Integer
Dim iGrid_Transacao_Col As Integer
Dim iGrid_Usuario_Col As Integer
Dim iGrid_NumIntDoc_Col As Integer

Dim objGridAviso As AdmGrid
Dim iAlterado As Integer

Dim sModulo As String
Dim iTransacao As Integer
Dim giInicioSistema As Integer
Dim giUltimaLinhaExibida As Integer

Const CONTABILIZACAO_OBRIGATORIA = 1
Const CONTABILIZACAO_NAO_OBRIGATORIA = 0

Const TAB_REGRAS = 1
Const TAB_EMAIL = 2
Const TAB_AVISO = 3
Const TAB_LOG = 4

'Mnemônicos
Private Const INICIO_SISTEMA As String = "Inicio_Sistema"
Private Const SAIDA_SISTEMA As String = "Saida_Sistema"
Private Const DATA_ULTIMA_VERDADEIRA As String = "Data_Ult_Exec"
Private Const HORA_ULTIMA_VERDADEIRA As String = "Hora_Ult_Exec"
Private Const DATA_ATUAL As String = "Data_Atual"
Private Const DIA_UTIL As String = "Dia_Util"
Private Const TEXTO_DATA As String = "Texto_Data"
Private Const ULTDIAUTIL As String = "UltDiaUtil"
Private Const LISTA_NFE_NAO_AUTO As String = "ListaNFeNaoAuto"
Private Const LISTA_CANC_NFE_NAO_HOM As String = "ListaCancNFeNaoHom"
Private Const LISTA_CANC_NFE_HOM As String = "ListaCancNFeHom"
Private Const LISTA_CERTIFICADOS_A_VENCER As String = "ListaCertifAVencer"
Private Const NUM_CERTIFICADOS_A_VENCER As String = "NumCertifAVencer"
Private Const LISTA_VISTPRJ_A_VENCER As String = "ListaVistPRJAVencer"
Private Const NUM_VISTPRJ_A_VENCER As String = "NumVistPRJAVencer"
Private Const LISTA_ETAPASPRJ_SEM_VIST As String = "ListaEtapasSemVist"
Private Const NUM_ETAPASPRJ_SEM_VIST As String = "NumEtapasSemVist"

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

Dim iLinha As Integer
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objAvisoWFW As New ClassAvisoWFW
Dim colLinhas As New Collection
Dim colNumIntDoc As New Collection
Dim colAvisoWFW1 As New Collection
Dim sMsg As String
Dim lErro As Long

On Error GoTo Erro_Form_QueryUnload

    For iLinha = 1 To GridAviso.Rows - 1

        lNumIntDoc = StrParaLong(GridAviso.TextMatrix(iLinha, iGrid_NumIntDoc_Col))
        
        lErro = CF("AvisoWFW_Le_NumIntDoc", lNumIntDoc, objAvisoWFW)
        If lErro <> SUCESSO And lErro <> 178220 Then gError 178225

        If lErro = SUCESSO Then
        
            colAvisoWFW1.Add objAvisoWFW
        
            If objAvisoWFW.dIntervalo = 0 Then
                colLinhas.Add iLinha
                colNumIntDoc.Add lNumIntDoc
            End If

        End If

    Next

    If colLinhas.Count > 0 Then
        
        For iIndice = 1 To colLinhas.Count
            If iIndice > 1 Then sMsg = sMsg & ", "
            sMsg = sMsg & colLinhas.Item(iIndice)
        Next
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_AVISOWFW_NAO_MOSTRAR", sMsg)
        
        If vbMsgRes = vbYes Then
            lErro = CF("AvisoWFW_Exclui", colNumIntDoc)
            If lErro <> SUCESSO Then gError 178226
        Else
            Cancel = 1
        End If
        
    End If

    If Cancel = 0 Then

        Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
    End If
    
    If Cancel = 0 Then
    
        lErro = CF("AvisoWFW_Atualiza_DataHoraUlt", colAvisoWFW1)
        If lErro <> SUCESSO Then gError 178193

    End If
    
    Exit Sub
    
Erro_Form_QueryUnload:

    Select Case gErr
    
        Case 178193, 178225, 178226
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178224)
        
    End Select

    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGridAviso = Nothing
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objGridAviso = New AdmGrid
    
    'inicializa o grid de lancamentos padrão
    lErro = Inicializa_Grid_Avisos(objGridAviso)
    If lErro <> SUCESSO Then gError 178181

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 178181
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178182)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Avisos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Avisos
    
    'tela em questão
    Set objGridAviso.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Aviso")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Hora")
    objGridInt.colColuna.Add ("Transação")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add (" ")
   
   'campos de edição do grid
    objGridInt.colCampo.Add (Msg.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (Hora.Name)
    objGridInt.colCampo.Add (Transacao.Name)
    objGridInt.colCampo.Add (Usuario.Name)
    objGridInt.colCampo.Add (NumIntDoc.Name)
    
    iGrid_Msg_Col = 1
    iGrid_Data_Col = 2
    iGrid_Hora_Col = 3
    iGrid_Transacao_Col = 4
    iGrid_Usuario_Col = 5
    iGrid_NumIntDoc_Col = 6
    
    NumIntDoc.Width = 0

    objGridInt.objGrid = GridAviso
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 7
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6
        
    GridAviso.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGridInt)

    GridAviso.ColWidth(6) = 0

    GridAviso.HighLight = flexHighlightAlways
    GridAviso.SelectionMode = flexSelectionByRow

    Inicializa_Grid_Avisos = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Avisos:

    Inicializa_Grid_Avisos = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178183)
        
    End Select

    Exit Function
        
End Function

Private Sub BotaoCancelar_Click()
    giRetornoTela = vbCancel
    Unload Me
End Sub

Private Sub BotaoLembrar_Click()

Dim iLinha As Integer
Dim iUMIntervalo As Integer
Dim dIntervalo As Double
Dim lNumIntDoc As Long
Dim colNumIntDoc As New Collection
Dim lErro As Long

On Error GoTo Erro_BotaoLembrar_Click

        dIntervalo = StrParaDbl(Intervalo.Text)
        
        If dIntervalo <= 0 Then gError 178212
        
        If dIntervalo >= 1000 Then gError 178213
        
        If TipoIntervalo.ListIndex = -1 Then gError 178214

        iUMIntervalo = TipoIntervalo.ItemData(TipoIntervalo.ListIndex)

        For iLinha = GridAviso.Row To GridAviso.RowSel

            lNumIntDoc = StrParaLong(GridAviso.TextMatrix(iLinha, iGrid_NumIntDoc_Col))
    
            colNumIntDoc.Add lNumIntDoc

        Next

        lErro = CF("AvisoWFW_Atualiza_Intervalo", colNumIntDoc, dIntervalo, iUMIntervalo)
        If lErro <> SUCESSO Then gError 178215
    
        Intervalo.Text = ""
        
        Call Cabecalho_Visivel(False)
        
        giUltimaLinhaExibida = 0
        
        Exit Sub
        
Erro_BotaoLembrar_Click:
                
    Select Case gErr
    
        Case 178212
            Call Rotina_Erro(vbOKOnly, "ERRO_WFW_LEMBRAR_ZERADO", gErr)
    
        Case 178213
            Call Rotina_Erro(vbOKOnly, "ERRO_WFW_LEMBRAR_MAIORQUE1000", gErr)
    
        Case 178214
            Call Rotina_Erro(vbOKOnly, "ERRO_WFW_LEMBRAR_NAO_SELECIONADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178216)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    giRetornoTela = vbOK
    Unload Me
End Sub

Public Sub GridAviso_Click()
    
Dim iExecutaEntradaCelula As Integer
Dim lErro As Long

On Error GoTo Erro_GridAviso_Click
    
    Call Grid_Click(objGridAviso, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAviso, iAlterado)
    End If
    
    lErro = Exibe_Linha_Grid()
    If lErro <> SUCESSO Then gError 178373
    
    Exit Sub
    
Erro_GridAviso_Click:

    Select Case gErr
    
        Case 178373
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178224)
        
    End Select

    Exit Sub
    
End Sub

Public Sub GridAviso_GotFocus()
    
    Call Grid_Recebe_Foco(objGridAviso)

End Sub

Public Sub GridAviso_EnterCell()
    
    Call Grid_Entrada_Celula(objGridAviso, iAlterado)
    
End Sub

Public Sub GridAviso_LeaveCell()
    
    Call Saida_Celula(objGridAviso)
    
End Sub

Public Sub GridAviso_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridAviso)
    
End Sub

Public Sub GridAviso_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAviso, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAviso, iAlterado)
    End If

End Sub

Public Sub GridAviso_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridAviso)

End Sub

Public Sub GridAviso_RowColChange()

Dim lErro As Long

On Error GoTo Erro_Grid_Aviso_RowColChange

    Call Grid_RowColChange(objGridAviso)
    
    lErro = Exibe_Linha_Grid()
    If lErro <> SUCESSO Then gError 178375
    
    Exit Sub
    
Erro_Grid_Aviso_RowColChange:

    Select Case gErr
    
        Case 178375
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178376)
        
    End Select

    Exit Sub
    
End Sub

Public Sub GridAviso_Scroll()

    Call Grid_Scroll(objGridAviso)
    
End Sub

Function Trata_Parametros(ByVal iInicioSistema As Integer) As Long

Dim lErro As Long
Dim objTela As Object
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Trata_Parametros
    
    objUsuarios.sCodUsuario = gsUsuario

    lErro = CF("Usuarios_Le", objUsuarios)
    If lErro <> SUCESSO Then gError 20847

    If objUsuarios.iWorkFlowAtivo <> WORKFLOW_ATIVO Then gError 20847
    
    giInicioSistema = iInicioSistema
    
    If iInicioSistema = 2 Then
        giRetornoTela = vbOK
        BotaoCancelar.Visible = True
        BotaoOK.Visible = True
    End If
    
    Set objTela = Me
    
    lErro = CF("WorkFlow_Trata_Transacao", "ADM", objTela, gsUsuario)
    If lErro <> SUCESSO Then gError 178365
    
    Call Limpa_Tela_Aviso
    
    lErro = Carga_Grid()
    If lErro <> SUCESSO Then gError 178184
    
    If objGridAviso.iLinhasExistentes = 0 Then gError 178227
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 178184, 178227, 178365, 130529, 130530, 20847
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178185)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Carga_Grid() As Long
'carrega os avisos no grid

Dim lErro As Long
Dim colAvisoWFW As New Collection
Dim colAvisoWFW1 As New Collection
Dim objAvisoWFW As ClassAvisoWFW
Dim dtDataHoraAtual As Date
Dim lSegundos As Long
Dim iLinha As Integer
Dim dtDataHora As Date

On Error GoTo Erro_Carga_Grid

    dtDataHoraAtual = Date + Time

    lErro = CF("AvisoWFW_Le_Usuario", gsUsuario, colAvisoWFW)
    If lErro <> SUCESSO Then gError 178190

    For Each objAvisoWFW In colAvisoWFW
    
        If objAvisoWFW.dtDataUltAviso <> DATA_NULA Then
    
            dtDataHora = objAvisoWFW.dtDataUltAviso + CDate(objAvisoWFW.dHoraUltAviso)
        
            If objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_MINUTO Then
                
                lSegundos = CLng(objAvisoWFW.dIntervalo * 60)
                
            ElseIf objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_HORA Then
                
                lSegundos = CLng(objAvisoWFW.dIntervalo * 3600)
                
            ElseIf objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_DIA Then
                
                lSegundos = CLng(objAvisoWFW.dIntervalo * 24 * 3600)
        
            ElseIf objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_SEMANA Then
        
                lSegundos = CLng(objAvisoWFW.dIntervalo * 7 * 24 * 3600)
        
            End If
        
            dtDataHora = DateAdd("s", lSegundos, dtDataHora)
    
        End If
    
        If dtDataHora <= dtDataHoraAtual Or objAvisoWFW.dtDataUltAviso = DATA_NULA Then
            iLinha = iLinha + 1
            If iLinha > GridAviso.Rows - 1 Then
                GridAviso.Rows = GridAviso.Rows + 1
                GridAviso.TextMatrix(iLinha, 0) = iLinha
            End If
                
            GridAviso.TextMatrix(iLinha, iGrid_Msg_Col) = objAvisoWFW.sMsg
            GridAviso.TextMatrix(iLinha, iGrid_Data_Col) = Format(objAvisoWFW.dtData, "dd/mm/yy")
            GridAviso.TextMatrix(iLinha, iGrid_Hora_Col) = Format(objAvisoWFW.dHora, "hh:mm:ss")
            GridAviso.TextMatrix(iLinha, iGrid_Transacao_Col) = objAvisoWFW.sTransacaoTela
            GridAviso.TextMatrix(iLinha, iGrid_Usuario_Col) = objAvisoWFW.sUsuarioOrig
            GridAviso.TextMatrix(iLinha, iGrid_NumIntDoc_Col) = objAvisoWFW.lNumIntDoc
            objGridAviso.iLinhasExistentes = objGridAviso.iLinhasExistentes + 1
            colAvisoWFW1.Add objAvisoWFW
        End If
        
    Next

    GridAviso.Row = objGridAviso.iLinhasExistentes
    
    If objGridAviso.iLinhasExistentes >= objGridAviso.iLinhasVisiveis Then
        GridAviso.TopRow = objGridAviso.iLinhasExistentes - objGridAviso.iLinhasVisiveis + 1
    Else
        GridAviso.TopRow = 1
    End If

    Carga_Grid = SUCESSO
    
    Exit Function

Erro_Carga_Grid:

    Carga_Grid = gErr

    Select Case gErr
    
        Case 178190
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178191)
    
    End Select
    
    iAlterado = 0
    
    Exit Function


End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 178192
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 178192
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178057)
        
    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumIntDoc As Long
Dim colNumIntDoc As New Collection
Dim iLinha As Integer
Dim iIndice As Integer
Dim iLinhaInicial As Integer
Dim iLinhaFinal As Integer
Dim iLinha1 As Integer

On Error GoTo Erro_BotaoExcluir_Click

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_AVISOWFW_COL", GridAviso.Row, GridAviso.RowSel)
    
    If vbMsgRes = vbYes Then

        If GridAviso.RowSel < GridAviso.Row Then
            iLinhaInicial = GridAviso.RowSel
            iLinhaFinal = GridAviso.Row
        Else
            iLinhaInicial = GridAviso.Row
            iLinhaFinal = GridAviso.RowSel
        End If

        For iLinha = iLinhaInicial To iLinhaFinal

            lNumIntDoc = StrParaLong(GridAviso.TextMatrix(iLinha, iGrid_NumIntDoc_Col))
    
            If lNumIntDoc = 0 Then
                iLinhaFinal = iLinha - 1
                Exit For
            End If
    
            colNumIntDoc.Add lNumIntDoc

        Next

        lErro = CF("AvisoWFW_Exclui", colNumIntDoc)
        If lErro <> SUCESSO Then gError 178205
    
        iLinha = iLinhaInicial
    
        If iLinhaFinal < objGridAviso.iLinhasExistentes Then
    
            For iIndice = iLinhaFinal + 1 To objGridAviso.iLinhasExistentes
                    
                GridAviso.TextMatrix(iLinha, iGrid_Msg_Col) = GridAviso.TextMatrix(iIndice, iGrid_Msg_Col)
                GridAviso.TextMatrix(iLinha, iGrid_Data_Col) = GridAviso.TextMatrix(iIndice, iGrid_Data_Col)
                GridAviso.TextMatrix(iLinha, iGrid_Hora_Col) = GridAviso.TextMatrix(iIndice, iGrid_Hora_Col)
                GridAviso.TextMatrix(iLinha, iGrid_Transacao_Col) = GridAviso.TextMatrix(iIndice, iGrid_Transacao_Col)
                GridAviso.TextMatrix(iLinha, iGrid_Usuario_Col) = GridAviso.TextMatrix(iIndice, iGrid_Usuario_Col)
                GridAviso.TextMatrix(iLinha, iGrid_NumIntDoc_Col) = GridAviso.TextMatrix(iIndice, iGrid_NumIntDoc_Col)
            
                iLinha = iLinha + 1
            
            Next
    
        End If
        
        For iLinha1 = iLinha To objGridAviso.iLinhasExistentes
            GridAviso.TextMatrix(iLinha1, iGrid_Msg_Col) = ""
            GridAviso.TextMatrix(iLinha1, iGrid_Data_Col) = ""
            GridAviso.TextMatrix(iLinha1, iGrid_Hora_Col) = ""
            GridAviso.TextMatrix(iLinha1, iGrid_Transacao_Col) = ""
            GridAviso.TextMatrix(iLinha1, iGrid_Usuario_Col) = ""
            GridAviso.TextMatrix(iLinha1, iGrid_NumIntDoc_Col) = ""
        Next
        
        objGridAviso.iLinhasExistentes = objGridAviso.iLinhasExistentes - (iLinhaFinal - iLinhaInicial + 1)
        
        
        If objGridAviso.iLinhasExistentes > objGridAviso.iLinhasVisiveis Then
            GridAviso.Rows = objGridAviso.iLinhasExistentes + 1
        Else
            GridAviso.Rows = objGridAviso.iLinhasVisiveis + 1
        End If
        
        
        Call Cabecalho_Visivel(False)
    
        giUltimaLinhaExibida = 0
        
        GridAviso.ColSel = GridAviso.Col
        GridAviso.RowSel = GridAviso.Row
        
    End If
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 178205
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178059)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoExcluirTudo_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lNumIntDoc As Long
Dim colNumIntDoc As New Collection
Dim iLinha As Integer

On Error GoTo Erro_BotaoExcluirTudo_Click

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_AVISOWFW")
    
    If vbMsgRes = vbYes Then

        For iLinha = 1 To objGridAviso.iLinhasExistentes

            lNumIntDoc = StrParaLong(GridAviso.TextMatrix(iLinha, iGrid_NumIntDoc_Col))
    
            colNumIntDoc.Add lNumIntDoc

        Next

        lErro = CF("AvisoWFW_Exclui", colNumIntDoc)
        If lErro <> SUCESSO Then gError 178211
    
        GridAviso.Rows = 7
    
        Call Limpa_Tela_Aviso
    
        giUltimaLinhaExibida = 0
        
    End If
    
    Exit Sub
    
Erro_BotaoExcluirTudo_Click:

    Select Case gErr
    
        Case 178211
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178059)
            
    End Select
    
    Exit Sub
    
End Sub

Sub Limpa_Tela_Aviso()

    objGridAviso.iLinhasExistentes = 0
    Call Grid_Limpa(objGridAviso)
    Call Limpa_Tela(Me)
    Call Cabecalho_Visivel(False)
    
End Sub

Public Sub BotaoLimpar_Click()

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 178102

    Call Limpa_Tela_Aviso
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 178102
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178103)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Avisos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AvisoWFW"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

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


Private Sub Cabecalho_Visivel(bVisivel As Boolean)
    LabelMsg.Visible = bVisivel
    LabelData.Visible = bVisivel
    LabelHora.Visible = bVisivel
    LabelUsuario.Visible = bVisivel
    LabelTransacao.Visible = bVisivel
    LabelLembrar.Visible = bVisivel
    Msg1.Visible = bVisivel
    Data1.Visible = bVisivel
    Hora1.Visible = bVisivel
    Usuario1.Visible = bVisivel
    Transacao1.Visible = bVisivel
    Intervalo1.Visible = bVisivel
    
    If bVisivel = False Then
        Msg1.Text = ""
        Data1.Caption = ""
        Hora1.Caption = ""
        Usuario1.Caption = ""
        Transacao1.Caption = ""
        Intervalo1.Caption = ""
    End If
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long
'Mnemonico de Workflow

Dim lErro As Long
Dim iCodigoF As Integer
Dim sFilial As String
Dim dtDataParam As Date
Dim dtDataUtil As Date
Dim iNumDias As Integer, dtDataDe As Date, colNF As New Collection
Dim objNF As ClassNFiscal, sTexto As String
Dim iNumCertificados As Integer, iNumAux As Integer

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case INICIO_SISTEMA
            If giInicioSistema = 1 Then
                objMnemonicoValor.colValor.Add 1
                giInicioSistema = 0
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case SAIDA_SISTEMA
            If giInicioSistema = 2 Then
                objMnemonicoValor.colValor.Add 1
                giInicioSistema = 0
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case DATA_ULTIMA_VERDADEIRA
            objMnemonicoValor.colValor.Add objContexto.dtDataUltExec

        Case HORA_ULTIMA_VERDADEIRA
            objMnemonicoValor.colValor.Add objContexto.dHoraUltExec * 3600000 * 24

        Case DATA_ATUAL
            objMnemonicoValor.colValor.Add Date
            
        Case DIA_UTIL
            dtDataParam = Forprint_ConvDataVB(objMnemonicoValor.vParam(1))
        
            lErro = CF("DataVencto_Real", dtDataParam, dtDataUtil)
            If lErro <> SUCESSO Then gError 188294
        
            If dtDataUtil = dtDataParam Then
                objMnemonicoValor.colValor.Add 1
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case TEXTO_DATA
            dtDataParam = StrParaDate(objMnemonicoValor.vParam(1))
            objMnemonicoValor.colValor.Add dtDataParam

        Case ULTDIAUTIL
        'pesquisa a ultima data util antes da data passada como parametro
            
            dtDataUtil = DATA_NULA
            dtDataParam = Forprint_ConvDataVB(objMnemonicoValor.vParam(1))

            Do While dtDataUtil <> dtDataParam

                dtDataParam = DateAdd("d", -1, dtDataParam)
            
                lErro = CF("DataVencto_Real", dtDataParam, dtDataUtil)
                If lErro <> SUCESSO Then gError 188297
            
            Loop
            
            objMnemonicoValor.colValor.Add dtDataUtil
            
        Case LISTA_NFE_NAO_AUTO
            
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            dtDataDe = DateAdd("d", -iNumDias, Date)
            
            lErro = CF("NFe_Le_Nao_Autorizadas", giFilialEmpresa, dtDataDe, colNF)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sTexto = ""
            For Each objNF In colNF
                If Len(Trim(sTexto)) > 0 Then sTexto = sTexto & ";"
                sTexto = sTexto & objNF.sSerie & " " & CStr(objNF.lNumNotaFiscal) & " de " & Format(objNF.dtDataEmissao, "dd/mm/yyyy")
            Next
            If Len(Trim(sTexto)) > 150 Then sTexto = left(sTexto, 150) & "..." 'Reduz para não dar erro na máquina de expressão
            
            objMnemonicoValor.colValor.Add sTexto

        Case LISTA_CANC_NFE_NAO_HOM
        
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            dtDataDe = DateAdd("d", -iNumDias, Date)
            
            lErro = CF("NFe_Le_Canc_Nao_Homologados", giFilialEmpresa, dtDataDe, colNF)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sTexto = ""
            For Each objNF In colNF
                If Len(Trim(sTexto)) > 0 Then sTexto = sTexto & ";"
                sTexto = sTexto & objNF.sSerie & " " & CStr(objNF.lNumNotaFiscal) & " de " & Format(objNF.dtDataEmissao, "dd/mm/yyyy")
            Next
            
            If Len(Trim(sTexto)) > 150 Then sTexto = left(sTexto, 150) & "..." 'Reduz para não dar erro na máquina de expressão
            
            objMnemonicoValor.colValor.Add sTexto

        Case LISTA_CANC_NFE_HOM
        
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            dtDataDe = DateAdd("d", -iNumDias, Date)
            
            lErro = CF("NFe_Le_Canc_Homologados", giFilialEmpresa, dtDataDe, colNF)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            sTexto = ""
            For Each objNF In colNF
                If Len(Trim(sTexto)) > 0 Then sTexto = sTexto & ";"
                sTexto = sTexto & objNF.sSerie & " " & CStr(objNF.lNumNotaFiscal) & " de " & Format(objNF.dtDataEmissao, "dd/mm/yyyy")
            Next
            
            If Len(Trim(sTexto)) > 150 Then sTexto = left(sTexto, 150) & "..." 'Reduz para não dar erro na máquina de expressão
            
            objMnemonicoValor.colValor.Add sTexto
            
        Case LISTA_CERTIFICADOS_A_VENCER
        
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            
            lErro = CF("Certificados_Lista_Validade_Texto", iNumDias, iNumCertificados, sTexto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            objMnemonicoValor.colValor.Add sTexto

        Case NUM_CERTIFICADOS_A_VENCER
        
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            
            lErro = CF("Certificados_Lista_Validade_Texto", iNumDias, iNumCertificados, sTexto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            objMnemonicoValor.colValor.Add iNumCertificados
            
        Case LISTA_VISTPRJ_A_VENCER
        
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            
            lErro = CF("VistoriaPRJ_Lista_Validade_Texto", iNumDias, iNumAux, sTexto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            If Len(Trim(sTexto)) > 200 Then sTexto = left(sTexto, 200) & "..." 'Reduz para não dar erro na máquina de expressão

            objMnemonicoValor.colValor.Add sTexto

        Case NUM_VISTPRJ_A_VENCER
        
            iNumDias = Forprint_ConvInt(objMnemonicoValor.vParam(1))
            
            lErro = CF("VistoriaPRJ_Lista_Validade_Texto", iNumDias, iNumAux, sTexto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            objMnemonicoValor.colValor.Add iNumAux
            
        Case LISTA_ETAPASPRJ_SEM_VIST
        
            lErro = CF("EtapasPRJ_Sem_Vistorias_Texto", iNumAux, sTexto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            If Len(Trim(sTexto)) > 200 Then sTexto = left(sTexto, 200) & "..." 'Reduz para não dar erro na máquina de expressão

            objMnemonicoValor.colValor.Add sTexto

        Case NUM_ETAPASPRJ_SEM_VIST
        
            lErro = CF("EtapasPRJ_Sem_Vistorias_Texto", iNumAux, sTexto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            objMnemonicoValor.colValor.Add iNumAux

        Case Else
            gError 188295

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 188294, 188297

        Case 188295
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188296)

    End Select

    Exit Function

End Function

Function Exibe_Linha_Grid() As Long

Dim lNumIntDoc As Long
Dim objAvisoWFW As New ClassAvisoWFW
Dim lErro As Long

On Error GoTo Erro_Exibe_Linha_Grid

    If GridAviso.Row > 0 And Len(GridAviso.TextMatrix(GridAviso.Row, iGrid_Msg_Col)) > 0 And objGridAviso.iLinhasExistentes >= GridAviso.Row Then

        If GridAviso.Row <> giUltimaLinhaExibida Then
            
            giUltimaLinhaExibida = GridAviso.Row

            Msg1.Text = GridAviso.TextMatrix(GridAviso.Row, iGrid_Msg_Col)
            Data1.Caption = Format(GridAviso.TextMatrix(GridAviso.Row, iGrid_Data_Col), "dd/mm/yyyy")
            Hora1.Caption = Format(GridAviso.TextMatrix(GridAviso.Row, iGrid_Hora_Col), "hh:mm:ss")
            Usuario1.Caption = GridAviso.TextMatrix(GridAviso.Row, iGrid_Usuario_Col)
            Transacao1.Caption = GridAviso.TextMatrix(GridAviso.Row, iGrid_Transacao_Col)
            lNumIntDoc = StrParaLong(GridAviso.TextMatrix(GridAviso.Row, iGrid_NumIntDoc_Col))
            
            lErro = CF("AvisoWFW_Le_NumIntDoc", lNumIntDoc, objAvisoWFW)
            If lErro <> SUCESSO And lErro <> 178220 Then gError 178223
            
            If lErro <> SUCESSO Then gError 178222
            
            Intervalo1.Caption = objAvisoWFW.dIntervalo
            
            If objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_MINUTO Then
                
                Intervalo1.Caption = Intervalo1.Caption & " minuto"
                
            ElseIf objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_HORA Then
                
                Intervalo1.Caption = Intervalo1.Caption & " hora"
                
            ElseIf objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_DIA Then
                
                Intervalo1.Caption = Intervalo1.Caption & " dia"
        
            ElseIf objAvisoWFW.iUMIntervalo = AVISOWFW_INTERVALO_SEMANA Then
        
                Intervalo1.Caption = Intervalo1.Caption & " semana"
        
            End If
            
            If objAvisoWFW.dIntervalo > 1 Then
                Intervalo1.Caption = Intervalo1.Caption & "s"
            End If
            
            Call Cabecalho_Visivel(True)
            
        End If
        
    End If

    Exibe_Linha_Grid = SUCESSO
    
    Exit Function
    
Erro_Exibe_Linha_Grid:

    Exibe_Linha_Grid = gErr

    Select Case gErr
    
        Case 178222
            Call Rotina_Erro(vbOKOnly, "ERRO__AVISOWFW_NAO_CADASTRADO", gErr, lNumIntDoc)
    
        Case 178223
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178374)
        
    End Select

End Function
