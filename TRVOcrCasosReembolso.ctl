VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVOcrCasosReemb 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame2 
      Caption         =   "Pré a Receber"
      Height          =   3780
      Left            =   45
      TabIndex        =   30
      Top             =   2415
      Width           =   9375
      Begin VB.TextBox PRDescricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   2715
         MaxLength       =   250
         TabIndex        =   36
         Top             =   600
         Width           =   5190
      End
      Begin MSMask.MaskEdBox PRValor 
         Height          =   255
         Left            =   3375
         TabIndex        =   35
         Top             =   1035
         Width           =   960
         _ExtentX        =   1693
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
         PromptChar      =   " "
      End
      Begin VB.CheckBox PRSel 
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
         Left            =   1410
         TabIndex        =   34
         Top             =   1020
         Width           =   510
      End
      Begin VB.TextBox PRData 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1860
         MaxLength       =   250
         TabIndex        =   33
         Top             =   990
         Width           =   1305
      End
      Begin MSFlexGridLib.MSFlexGrid GridPR 
         Height          =   2715
         Left            =   45
         TabIndex        =   13
         Top             =   210
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   4789
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Total:"
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
         Height          =   225
         Index           =   25
         Left            =   5700
         TabIndex        =   32
         Top             =   3525
         Width           =   2130
      End
      Begin VB.Label Total 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7845
         TabIndex        =   31
         Top             =   3480
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pagamento"
      Height          =   915
      Left            =   45
      TabIndex        =   26
      Top             =   1485
      Width           =   9390
      Begin VB.CheckBox optHistAuto 
         Caption         =   "Histórico automático"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6240
         TabIndex        =   11
         Top             =   210
         Value           =   1  'Checked
         Width           =   2865
      End
      Begin MSComCtl2.UpDown UpDownDataCredito 
         Height          =   300
         Left            =   5820
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCredito 
         Height          =   300
         Left            =   4740
         TabIndex        =   9
         Top             =   195
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Historico 
         Height          =   300
         Left            =   1020
         TabIndex        =   12
         Top             =   540
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.ComboBox ContaCorrente 
         Height          =   315
         Left            =   1020
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   195
         Width           =   2415
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
         Height          =   195
         Left            =   105
         TabIndex        =   29
         Top             =   570
         Width           =   825
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
         Height          =   195
         Left            =   3525
         TabIndex        =   28
         Top             =   255
         Width           =   1140
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
         Left            =   360
         TabIndex        =   27
         Top             =   225
         Width           =   570
      End
   End
   Begin VB.Frame Frame0 
      Caption         =   "Identificação"
      Height          =   900
      Index           =   0
      Left            =   45
      TabIndex        =   20
      Top             =   570
      Width           =   9390
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2385
         Picture         =   "TRVOcrCasosReembolso.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Numeração Automática"
         Top             =   540
         Width           =   300
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4740
         TabIndex        =   1
         Top             =   195
         Width           =   1815
      End
      Begin MSMask.MaskEdBox NumTitulo 
         Height          =   300
         Left            =   1035
         TabIndex        =   4
         Top             =   540
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   300
         Left            =   1035
         TabIndex        =   0
         Top             =   210
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   8985
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   195
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   7890
         TabIndex        =   2
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   7890
         TabIndex        =   7
         Top             =   525
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Ccl 
         Height          =   300
         Left            =   4740
         TabIndex        =   6
         Top             =   540
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   " "
      End
      Begin VB.Label CclLabel 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo:"
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
         Left            =   3240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   600
         Width           =   1440
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   255
         Width           =   660
      End
      Begin VB.Label NumeroLabel 
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
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   585
         Width           =   720
      End
      Begin VB.Label LabelFilial 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
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
         Left            =   4155
         TabIndex        =   23
         Top             =   255
         Width           =   525
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   7080
         TabIndex        =   22
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label17 
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
         Left            =   7320
         TabIndex        =   21
         Top             =   555
         Width           =   510
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5970
      ScaleHeight     =   450
      ScaleWidth      =   3405
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   45
      Width           =   3465
      Begin VB.CommandButton BotaoConsulta 
         Height          =   360
         Left            =   90
         Picture         =   "TRVOcrCasosReembolso.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   1230
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2865
         Picture         =   "TRVOcrCasosReembolso.ctx":1EAC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   2385
         Picture         =   "TRVOcrCasosReembolso.ctx":202A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1890
         Picture         =   "TRVOcrCasosReembolso.ctx":255C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   1395
         Picture         =   "TRVOcrCasosReembolso.ctx":26E6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "TRVOcrCasosReemb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim glNumIntDoc As Long
Dim gcolPR As Collection

Dim objGridPR As AdmGrid
Dim iGrid_PRSel_Col As Integer
Dim iGrid_PRData_Col As Integer
Dim iGrid_PRValor_Col As Integer
Dim iGrid_PRDescricao_Col As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Dim iAlterado As Integer


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Reembolso de antecipação de pagto de seguro"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVOcrCasosReemb"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        End If
    
    End If
    
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
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCcl = Nothing

    Set objGridPR = Nothing
    
    Set gcolPR = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208809)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim sConteudo As String
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load
    
    'Inicializa os Eventos da Tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoCcl = New AdmEvento
    
    Set gcolPR = New Collection

    'Inicializa os Grids da tela
    Set objGridPR = New AdmGrid

    'Faz as inicializações particulares ao GridParcelas
    lErro = Inicializa_Grid_PR(objGridPR)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    'Carrega list de ComboBox ContaCorrente
    lErro = ContaCorrente_Carrega(ContaCorrente)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_CLIENTE_OCR_REEMBOLSO_PADRAO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaLong(sConteudo) <> 0 Then
    
        Cliente.Text = sConteudo
        Call Cliente_Validate(bSGECancelDummy)
        
    End If
    
    'Inicializa Máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Ccl.Mask = sMascaraCcl

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208810)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional ByVal objTitRec As ClassTituloReceber) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTitRec Is Nothing) Then

        lErro = Traz_Reembolso_Tela(objTitRec)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208811)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objTituloReceber As ClassTituloReceber, ByVal colPR As Collection, Optional ByVal bValida As Boolean = True) As Long

Dim lErro As Long, iLinha As Integer
Dim objcliente As New ClassCliente
Dim objParcelaReceber As New ClassParcelaReceber
Dim colParcelaReceber As New colParcelaReceber
Dim sCclFormatada As String, iCclPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o cliente foi digitado
    If Len(Trim(Cliente.ClipText)) > 0 Then
    
        objcliente.sNomeReduzido = Cliente.Text
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
        
        'Não achou o Cliente --> erro
        If lErro <> SUCESSO Then gError 208812
        
        'Guarda o código no objTituloReceber
        objTituloReceber.lCliente = objcliente.lCodigo
        
    End If

    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.lNumTitulo = StrParaLong(NumTitulo.Text)
    objTituloReceber.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objTituloReceber.dValor = StrParaDbl(Valor.Text)
    objTituloReceber.iNumParcelas = 1
    objTituloReceber.iFilialEmpresa = giFilialEmpresa
    objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_OCR_REEM
    objTituloReceber.iMoeda = MOEDA_REAL
    objTituloReceber.dtReajusteBase = DATA_NULA
    
    objParcelaReceber.iNumParcela = 1
    objParcelaReceber.dtDataVencimento = StrParaDate(DataCredito.Text)
    objParcelaReceber.dtDataVencimentoReal = StrParaDate(DataCredito.Text)
    objParcelaReceber.dValor = objTituloReceber.dValor
    objParcelaReceber.dValorOriginal = objTituloReceber.dValor
    objParcelaReceber.dtDataCredito = DATA_NULA
    objParcelaReceber.dtDataDepositoCheque = DATA_NULA
    objParcelaReceber.dtDataEmissaoCheque = DATA_NULA
    objParcelaReceber.dtDataPrevReceb = DATA_NULA
    objParcelaReceber.dtDataProxCobr = DATA_NULA
    objParcelaReceber.dtDataTransacaoCartao = DATA_NULA
    objParcelaReceber.dtDesconto1Ate = DATA_NULA
    objParcelaReceber.dtDesconto2Ate = DATA_NULA
    objParcelaReceber.dtDesconto3Ate = DATA_NULA
    objParcelaReceber.dtValidadeCartao = DATA_NULA
    
    objParcelaReceber.sObservacao = Historico.Text
    objParcelaReceber.iCodConta = Codigo_Extrai(ContaCorrente.Text)

    colParcelaReceber.AddObj objParcelaReceber
    
    Set objTituloReceber.colParcelaReceber = colParcelaReceber
    
    For iLinha = 1 To objGridPR.iLinhasExistentes
        If StrParaInt(GridPR.TextMatrix(iLinha, iGrid_PRSel_Col)) = MARCADO Then
            colPR.Add gcolPR.Item(iLinha)
        End If
    Next
    
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iCclPreenchida <> CCL_PREENCHIDA And bValida Then gError 209160
        
    objTituloReceber.sCcl = sCclFormatada
    objTituloReceber.sNatureza = Ccl.Text
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 208812
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
            
        Case 209160
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_INFORMADO", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208813)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim colPR As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosRecTodos"
    
    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objTituloReceber, colPR, False)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objTituloReceber.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Cliente", objTituloReceber.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objTituloReceber.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloReceber.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloReceber.dtDataEmissao, DATA_NULA, "DataEmissao"

    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "SiglaDocumento", OP_IGUAL, TIPODOC_FATURA_OCR_REEM

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208814)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber

On Error GoTo Erro_Tela_Preenche

    objTituloReceber.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objTituloReceber.lNumIntDoc <> 0 Then

        lErro = Traz_Reembolso_Tela(objTituloReceber)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208815)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim colPR As New Collection
Dim objPR As ClassTRVOcrCasosPreRec
Dim dValorTotal As Double, vbResult As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 208816
    If Len(Trim(Filial.Text)) = 0 Then gError 208817
    If Len(Trim(NumTitulo.ClipText)) = 0 Then gError 208818
    If Len(Trim(Valor.ClipText)) = 0 Then gError 208819
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 208820
    If Len(Trim(DataCredito.ClipText)) = 0 Then gError 208821
    If Codigo_Extrai(ContaCorrente.Text) = 0 Then gError 208822
    '#####################

    'Preenche o objTRVTiposOcorrencia
    lErro = Move_Tela_Memoria(objTituloReceber, colPR)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If colPR.Count = 0 Then gError 208823
    
    dValorTotal = 0
    For Each objPR In colPR
        dValorTotal = dValorTotal + objPR.dValor
    Next
    
    If Abs(dValorTotal - objTituloReceber.dValor) > DELTA_VALORMONETARIO Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRVOCRCASOS_REEMB_VALOR_DIF")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    End If

    lErro = CF("TRVOcrCasosReemb_Grava", objTituloReceber, colPR)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 208816
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 208817
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
    
        Case 208818
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
            
        Case 208819
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 208820
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 208821
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACREDITO_NAO_PREENCHIDA", gErr)

        Case 208822
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PREENCHIDA", gErr)

        Case 208823
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEM_SELECIONADO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208824)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Reembolso() As Long

Dim lErro As Long, sConteudo As String

On Error GoTo Erro_Limpa_Tela_Reembolso

    Total.Caption = ""
    
    NumTitulo.Enabled = True
    
    ContaCorrente.ListIndex = -1
    Filial.Clear
    
    glNumIntDoc = 0
    
    Call Grid_Limpa(objGridPR)

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_CLIENTE_OCR_REEMBOLSO_PADRAO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaLong(sConteudo) <> 0 Then
    
        Cliente.Text = sConteudo
        Call Cliente_Validate(bSGECancelDummy)
        
    End If
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    optHistAuto.Value = vbChecked
    iAlterado = 0

    Limpa_Tela_Reembolso = SUCESSO

    Exit Function

Erro_Limpa_Tela_Reembolso:

    Limpa_Tela_Reembolso = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208825)

    End Select

    Exit Function

End Function

Function Traz_Reembolso_Tela(ByVal objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim colPR As New Collection
Dim colParcelaReceber As New colParcelaReceber
Dim objParcelaReceber As New ClassParcelaReceber
Dim iLinha As Integer

On Error GoTo Erro_Traz_Reembolso_Tela

    'Limpa a tela
    Call Limpa_Tela_Reembolso
    
    'Lê o Título à Receber
    lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 56570 Then gError ERRO_SEM_MENSAGEM
    
    'Coloca o Cliente na Tela
    Cliente.Text = objTituloReceber.lCliente
    Call Cliente_Validate(bSGECancelDummy)

    'Coloca a Filial na Tela
    Filial.Text = objTituloReceber.iFilial
    Call Filial_Validate(bSGECancelDummy)
        
    'Coloca os demais dados do Título a Receber na Tela
    NumTitulo.Text = CStr(objTituloReceber.lNumTitulo)
    Call DateParaMasked(DataEmissao, objTituloReceber.dtDataEmissao)
    Valor.Text = Format(objTituloReceber.dValor, "Standard")

    'Lê as Parcelas a Receber vinculadas ao Título
    lErro = CF("ParcelasReceber_Le_Todas", objTituloReceber, colParcelaReceber)
    If lErro <> SUCESSO And lErro <> 58990 Then gError ERRO_SEM_MENSAGEM

    'Lê as Parcelas a Receber vinculadas ao Título
    lErro = CF("TRVOcrCasosPreRec_Le_Vinculados", objTituloReceber.lNumIntDoc, colPR)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objParcelaReceber.lNumIntTitulo = objTituloReceber.lNumIntDoc
    objParcelaReceber.iNumParcela = 1

    'Lê as Parcelas a Receber vinculadas ao Título
    lErro = CF("ParcelaReceberBaixada_Le_SemNumIntDoc", objParcelaReceber)
    If lErro <> SUCESSO And lErro <> 28567 Then gError ERRO_SEM_MENSAGEM
    
    objParcelaReceber.sObservacao = colParcelaReceber.Item(1).sObservacao
    
    Call Combo_Seleciona_ItemData(ContaCorrente, objParcelaReceber.iCodConta)

    Call DateParaMasked(DataCredito, objParcelaReceber.dtDataVencimento)
    
    Historico.Text = objParcelaReceber.sObservacao
    
    NumTitulo.Enabled = False
    glNumIntDoc = objTituloReceber.lNumIntDoc
    
    lErro = Traz_PR_Tela(colPR, MARCADO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    optHistAuto.Value = vbUnchecked
    iAlterado = 0

    Traz_Reembolso_Tela = SUCESSO

    Exit Function

Erro_Traz_Reembolso_Tela:

    Traz_Reembolso_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208826)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_Reembolso

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208827)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208828)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Reembolso

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208829)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If glNumIntDoc = 0 Then gError 208830

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TITULORECEBER", StrParaLong(NumTitulo.Text))

    If vbMsgRes = vbYes Then
    
        objTituloReceber.lNumIntDoc = glNumIntDoc

        lErro = CF("TRVOcrCasosReemb_Exclui", objTituloReceber)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Limpa Tela
        Call Limpa_Tela_Reembolso

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 208830
            Call Rotina_Erro(vbOKOnly, "ERRO_EXC_NECESS_TRAZER_DOC_TELA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208832)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_PR(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Descrição")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (PRSel.Name)
    objGridInt.colCampo.Add (PRData.Name)
    objGridInt.colCampo.Add (PRValor.Name)
    objGridInt.colCampo.Add (PRDescricao.Name)

    'Colunas do Grid
    iGrid_PRSel_Col = 1
    iGrid_PRData_Col = 2
    iGrid_PRValor_Col = 3
    iGrid_PRDescricao_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridPR

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 1000 + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridPR.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_PR = SUCESSO

    Exit Function

End Function

Private Sub GridPR_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPR, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPR, iAlterado)
    End If

End Sub

Private Sub GridPR_GotFocus()
    Call Grid_Recebe_Foco(objGridPR)
End Sub

Private Sub GridPR_EnterCell()
    Call Grid_Entrada_Celula(objGridPR, iAlterado)
End Sub

Private Sub GridPR_LeaveCell()
    Call Saida_Celula(objGridPR)
End Sub

Private Sub GridPR_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPR, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPR, iAlterado)
    End If

End Sub

Private Sub GridPR_RowColChange()
    Call Grid_RowColChange(objGridPR)
End Sub

Private Sub GridPR_Scroll()
    Call Grid_Scroll(objGridPR)
End Sub

Private Sub GridPR_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridPR)

End Sub

Private Sub GridPR_LostFocus()
    Call Grid_Libera_Foco(objGridPR)
End Sub

Public Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Cliente_Preenche
End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear
        
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208833)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Preenche()

Static sNomeReduzidoParte As String
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208834)

    End Select
    
    Exit Sub

End Sub

Public Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 208835

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM

        If lErro = 17660 Then gError 208836

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
    
    End If

    'Não encontrou a STRING
    If lErro = 6731 Then gError 208837

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208835
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 208836
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 208837
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208838)

    End Select

    Exit Sub

End Sub

Public Sub NumTitulo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colPR As New Collection

On Error GoTo Erro_NumTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumTitulo.ClipText)) <> 0 Then

        'Critica se é Long positivo
        lErro = Long_Critica(NumTitulo.ClipText)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objGridPR.iLinhasExistentes = 0 Then
        
            'Lê as Parcelas a Receber vinculadas ao Título
            lErro = CF("TRVOcrCasosPreRec_Le_Vinculados", 0, colPR)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Traz_PR_Tela(colPR, DESMARCADO)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
    Else
        Call Grid_Limpa(objGridPR)
        Total.Caption = "0,00"
    End If
    
    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208839)

    End Select

    Exit Sub

End Sub

Public Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)

End Sub

Public Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate
    
    If StrParaDate(DataEmissao.Text) <> DATA_NULA Then
    
        'Critica a data digitada
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208840)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
End Sub

Public Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Valor_Validate
        
    'Verifica se algum valor foi digitado
    If Len(Trim(Valor.ClipText)) <> 0 Then

        'Critica se é valor positivo
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Põe o valor formatado na tela
        Valor.Text = Format(Valor.Text, "Standard")
    
    End If
    
    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208841)

    End Select

    Exit Sub

End Sub

Private Sub PRSel_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Soma_Coluna_Grid(objGridPR, iGrid_PRValor_Col, Total, False, iGrid_PRSel_Col)
    Call Traz_Historico
End Sub

Public Sub PRSel_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPR)
End Sub

Public Sub PRSel_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPR)
End Sub

Public Sub PRSel_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPR.objControle = PRSel
    lErro = Grid_Campo_Libera_Foco(objGridPR)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridPR.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_PRSel_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, PRSel)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                     
            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 208842

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208842
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208843)

    End Select

    Exit Function

End Function

Private Function ContaCorrente_Carrega(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_ContaCorrente_Carrega

    'Lê Codigos, NomesReduzidos de ContasCorrentes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preeche list de ComboBox
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    ContaCorrente_Carrega = SUCESSO

    Exit Function

Erro_ContaCorrente_Carrega:

    ContaCorrente_Carrega = Err

    Select Case Err

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208844)

    End Select

    Exit Function

End Function

Public Sub ContaCorrente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContaCorrente_Validate(Cancel As Boolean)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Passa o Código da Conta para o Obj
        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then gError ERRO_SEM_MENSAGEM

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then gError 208845
        
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Verifica se a Conta é Filial Empresa corrente
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then gError 208846
        
        End If
        
        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then gError 208847
    
    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208845
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            End If

        Case 208846
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_FILIAL_DIFERENTE", gErr, objContaCorrenteInt.iCodigo, giFilialEmpresa)
        
        Case 208847
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, ContaCorrente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208848)

    End Select

    Exit Sub

End Sub

Public Sub DataCredito_Change()
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataCredito_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataCredito_Validate

    'Se a DataCredito está preenchida
    If Len(DataCredito.ClipText) > 0 Then

        'Verifica se a DataCredito é válida
        lErro = Data_Critica(DataCredito.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DataCredito_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208849)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCredito_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCredito_DownClick

    'Diminui a DataCredito em 1 dia
    lErro = Data_Up_Down_Click(DataCredito, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataCredito_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208850)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCredito_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCredito_UpClick

    'Aumenta a DataCredito em 1 dia
    lErro = Data_Up_Down_Click(DataCredito, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDataCredito_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208851)

    End Select

    Exit Sub

End Sub

Public Sub DataCredito_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataCredito, iAlterado)
End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208852)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208853)

    End Select

    Exit Sub

End Sub

Public Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
    optHistAuto.Value = vbUnchecked
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Me.Show

    'Preenche Cliente na tela com NomeReduzido
    Set objcliente = obj1
    
    Cliente.Text = CStr(objcliente.sNomeReduzido)

    'Chama Validate de Cliente
    Call Cliente_Validate(bCancel)
    
    Exit Sub

End Sub

Public Sub NumeroLabel_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_NumeroLabel_Click

    objTituloReceber.lNumTitulo = StrParaLong(NumTitulo.Text)

    sSelecao = "SiglaDocumento = ?"
    colSelecao.Add TIPODOC_FATURA_OCR_REEM

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TitRecTodosTFLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)
            
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208854)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloReceber = obj1

    'Traz os dados de objTituloReceber para tela
    lErro = Traz_Reembolso_Tela(objTituloReceber)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208855)

    End Select

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Function Traz_PR_Tela(ByVal colPR As Collection, ByVal iFlag As Integer) As Long

Dim lErro As Long
Dim objPR As ClassTRVOcrCasosPreRec
Dim iLinha As Integer

On Error GoTo Erro_Traz_PR_Tela

    iLinha = 0
    For Each objPR In colPR
        iLinha = iLinha + 1
        GridPR.TextMatrix(iLinha, iGrid_PRData_Col) = Format(objPR.dtData, "dd/mm/yyyy")
        GridPR.TextMatrix(iLinha, iGrid_PRValor_Col) = Format(objPR.dValor, "STANDARD")
        GridPR.TextMatrix(iLinha, iGrid_PRDescricao_Col) = objPR.sDescricao
        GridPR.TextMatrix(iLinha, iGrid_PRSel_Col) = CStr(iFlag)
    Next
    
    Set gcolPR = colPR
 
    objGridPR.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridPR)
    
    Call Soma_Coluna_Grid(objGridPR, iGrid_PRValor_Col, Total, False, iGrid_PRSel_Col)

    iAlterado = 0

    Traz_PR_Tela = SUCESSO

    Exit Function

Erro_Traz_PR_Tela:

    Traz_PR_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208856)

    End Select

    Exit Function

End Function

Function Traz_Historico() As Long

Dim lErro As Long
Dim iLinha As Integer, sTexto As String, sSubTexto As String, iPOS As Integer
Dim objPR As ClassTRVOcrCasosPreRec

On Error GoTo Erro_Traz_Historico

    If optHistAuto.Value = vbChecked Then
    
        sTexto = "Reemb. Ocrs Assist.: "
        For iLinha = 1 To objGridPR.iLinhasExistentes
            If StrParaInt(GridPR.TextMatrix(iLinha, iGrid_PRSel_Col)) = MARCADO Then
                Set objPR = gcolPR.Item(iLinha)
                iPOS = InStr(1, objPR.sDescricao, "Vou:")
                If Len(Trim(sSubTexto)) > 0 Then sSubTexto = sSubTexto & ","
                sSubTexto = sSubTexto & Mid(left(objPR.sDescricao, iPOS - 2), 6)
            End If
        Next
        sTexto = left(sTexto & sSubTexto, STRING_HISTORICO)
        
        Historico.Text = sTexto
        optHistAuto.Value = vbChecked
    
    End If
    
    Traz_Historico = SUCESSO

    Exit Function

Erro_Traz_Historico:

    Traz_Historico = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208857)

    End Select

    Exit Function

End Function

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("Config_ObterNumInt_Trans", "TRVConfig", TRVCONFIG_PROX_NUM_TITREC, lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NumTitulo.Text = CStr(lCodigo)
    Call NumTitulo_Validate(bSGECancelDummy)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208858)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objFat As New ClassTituloReceber

On Error GoTo Erro_BotaoConsulta_Click

    If glNumIntDoc <> 0 Then

        objFat.lNumIntDoc = glNumIntDoc
        
        Call Chama_Tela(TRV_TIPO_DOC_DESTINO_TITREC_TELA, objFat)
        
    End If
    
    Exit Sub

Erro_BotaoConsulta_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208872)

    End Select
    
    Exit Sub
    
End Sub

Public Sub CclLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
'Preenche Ccl

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209157)

    End Select

    Exit Sub

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(Ccl.Text)) = 0 Then Exit Sub

    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError ERRO_SEM_MENSAGEM

    If lErro = 5703 Then gError 209158

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 209158
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209159)

    End Select

    Exit Sub

End Sub
