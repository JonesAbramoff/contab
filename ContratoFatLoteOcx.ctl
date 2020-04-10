VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ContratoFatLoteOcx 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   7545
   Begin VB.PictureBox Picture 
      Height          =   555
      Left            =   5550
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   90
      Width           =   1620
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ContratoFatLoteOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "ContratoFatLoteOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "ContratoFatLoteOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5745
      Left            =   180
      TabIndex        =   16
      Top             =   90
      Width           =   7155
      Begin VB.Frame Frame5 
         Caption         =   "Data de Cobrança"
         Height          =   840
         Left            =   105
         TabIndex        =   30
         Top             =   1665
         Width           =   6840
         Begin MSMask.MaskEdBox DataCobrIni 
            Height          =   300
            Left            =   870
            TabIndex        =   5
            Top             =   360
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownCobrIni 
            Height          =   300
            Left            =   1980
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   345
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataCobrFim 
            Height          =   300
            Left            =   4395
            TabIndex        =   7
            Top             =   330
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownCobrFim 
            Height          =   300
            Left            =   5505
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   330
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   510
            TabIndex        =   32
            Top             =   390
            Width           =   360
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00000080&
            Height          =   180
            Left            =   3990
            TabIndex        =   31
            Top             =   375
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Clientes"
         Height          =   810
         Left            =   105
         TabIndex        =   23
         Top             =   3975
         Width           =   6840
         Begin MSMask.MaskEdBox ClienteIni 
            Height          =   300
            Left            =   855
            TabIndex        =   13
            Top             =   330
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteFim 
            Height          =   300
            Left            =   4065
            TabIndex        =   14
            Top             =   330
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label ClienteIniLabel 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   390
            Width           =   315
         End
         Begin VB.Label ClienteFimLabel 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   3660
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   390
            Width           =   360
         End
      End
      Begin VB.ComboBox TipoNFiscal 
         Height          =   315
         ItemData        =   "ContratoFatLoteOcx.ctx":080A
         Left            =   1005
         List            =   "ContratoFatLoteOcx.ctx":080C
         TabIndex        =   0
         Top             =   525
         Width           =   2835
      End
      Begin VB.CommandButton BotaoPreviaFat 
         Caption         =   "Prévia do Faturamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   15
         TabIndex        =   15
         Top             =   5025
         Width           =   2730
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contratos"
         Height          =   1245
         Left            =   105
         TabIndex        =   19
         Top             =   2610
         Width           =   6840
         Begin MSMask.MaskEdBox ContratoIni 
            Height          =   315
            Left            =   870
            TabIndex        =   9
            Top             =   285
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContratoFim 
            Height          =   315
            Left            =   870
            TabIndex        =   11
            Top             =   795
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label DescContratoIni 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2235
            TabIndex        =   10
            Top             =   285
            Width           =   4410
         End
         Begin VB.Label DescContratoFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2235
            TabIndex        =   12
            Top             =   795
            Width           =   4410
         End
         Begin VB.Label ContratoFimLabel 
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
            Height          =   210
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   825
            Width           =   360
         End
         Begin VB.Label ContratoIniLabel 
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
            Height          =   210
            Left            =   510
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   20
            Top             =   330
            Width           =   360
         End
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Top             =   1155
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   2085
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1170
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   300
         Left            =   4500
         TabIndex        =   3
         Top             =   1155
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownVencimento 
         Height          =   300
         Left            =   5610
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1170
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label TipoNFiscalLabel 
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
         Left            =   480
         TabIndex        =   22
         Top             =   540
         Width           =   450
      End
      Begin VB.Label DataEmissaoLabel 
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
         Height          =   180
         Left            =   165
         TabIndex        =   18
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label DataVencimentoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Data Ref. Vencimento:"
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
         Left            =   2505
         TabIndex        =   17
         Top             =   1185
         Width           =   1950
      End
   End
End
Attribute VB_Name = "ContratoFatLoteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents objEventoClienteIni As AdmEvento
Attribute objEventoClienteIni.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoContratoFim As AdmEvento
Attribute objEventoContratoFim.VB_VarHelpID = -1
Private WithEvents objEventoContratoIni As AdmEvento
Attribute objEventoContratoIni.VB_VarHelpID = -1

'#######################################################
Const DOCINFO_RECIBO = 171

Const STRING_NFFATURA = "Nota Fiscal de Fatura"
Const STRING_RECIBO = "Recibo de Fatura"
'#######################################################

Dim iAlterado As Integer

'HElp
Const IDH_RASTROPRODNFFAT = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 129750

    Call Limpa_Tela_Faturamento

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 129750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155041)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 129751

    Call Limpa_Tela_Faturamento

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 129751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155042)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objGeracaoFatContrato As ClassGeracaoFatContrato) As Long

Dim lErro As Long
Dim objClienteIni As New ClassCliente
Dim objClienteFim As New ClassCliente

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o Cliente Inicial foi preenchido
    If Len(Trim(ClienteIni.ClipText)) > 0 Then

        objClienteIni.sNomeReduzido = ClienteIni.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objClienteIni)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 131062

        If lErro = SUCESSO Then objGeracaoFatContrato.lClienteIni = objClienteIni.lCodigo
                            
    End If
    
    'Verifica se o Cliente Final foi preenchido
    If Len(Trim(ClienteFim.ClipText)) > 0 Then

        objClienteFim.sNomeReduzido = ClienteFim.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objClienteFim)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 131063

        If lErro = SUCESSO Then objGeracaoFatContrato.lClienteFim = objClienteFim.lCodigo
                            
    End If

    objGeracaoFatContrato.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objGeracaoFatContrato.dtDataCobrIni = StrParaDate(DataCobrIni.Text)
    objGeracaoFatContrato.dtDataCobrFim = StrParaDate(DataCobrFim.Text)
    objGeracaoFatContrato.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objGeracaoFatContrato.dtDataRefVencimento = StrParaDate(DataVencimento.Text)

    objGeracaoFatContrato.sContratoIni = ContratoIni.Text
    objGeracaoFatContrato.sContratoFim = ContratoFim.Text

    objGeracaoFatContrato.iFilialEmpresa = giFilialEmpresa
    objGeracaoFatContrato.sUsuario = gsUsuario
    objGeracaoFatContrato.dtDataGeracao = gdtDataAtual

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 131062 To 131063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155043)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim lNumAlmoxarifados As Long

On Error GoTo Erro_Form_Load

    Set objEventoClienteIni = New AdmEvento
    Set objEventoContratoIni = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set objEventoContratoFim = New AdmEvento

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    lErro = Carrega_ComboTipo(TipoNFiscal)
    If lErro <> SUCESSO Then gError 129752

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 129752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155044)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoClienteIni = Nothing
    Set objEventoContratoIni = Nothing
    Set objEventoClienteFim = Nothing
    Set objEventoContratoFim = Nothing
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()
    'COMENTADO
    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
    'COMENTADO
    'gi_ST_SetaIgnoraClick = 1

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTROPRODNFFAT
    Set Form_Load_Ocx = Me
    Caption = "Faturamento de Contrato"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ContratoFatLote"

End Function

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

Private Sub BotaoPreviaFat_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = PreencherRelOp()
    If lErro <> SUCESSO Then gError 129784
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 129784

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155045)

    End Select

    Exit Sub

End Sub

Private Sub ClienteFim_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClienteIni_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClienteIni_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_ClienteIni_Validate

    'Se Cliente está preenchido
    If Len(Trim(ClienteIni.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(ClienteIni, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 129753

    End If
    
    Exit Sub

Erro_ClienteIni_Validate:
        
    Cancel = True

    Select Case gErr
    
        Case 129753
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155046)

    End Select

    Exit Sub

End Sub

Private Sub ClienteFim_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_ClienteFim_Validate

    'Se Cliente está preenchido
    If Len(Trim(ClienteFim.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(ClienteFim, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 129754

    End If
    
    Exit Sub

Erro_ClienteFim_Validate:
        
    Cancel = True

    Select Case gErr
    
        Case 129754
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155047)

    End Select

    Exit Sub

End Sub

Private Sub ContratoFim_Change()
    
    iAlterado = REGISTRO_ALTERADO
    DescContratoFim.Caption = ""

End Sub

Private Sub ContratoFim_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_ContratoFim_Validate

    If Len(Trim(ContratoFim.Text)) = 0 Then Exit Sub

    objContrato.sCodigo = ContratoFim.Text
    objContrato.iFilialEmpresa = giFilialEmpresa

    lErro = Traz_ContratoFim_Tela(objContrato)
    If lErro <> SUCESSO Then gError 129755
    
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_ContratoFim_Validate:

    Cancel = True

    Select Case gErr

        Case 129755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155048)

    End Select
    
End Sub

Private Sub ContratoIni_Change()
    
    iAlterado = REGISTRO_ALTERADO
    DescContratoIni.Caption = ""

End Sub

Private Sub ContratoIni_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_ContratoIni_Validate

    If Len(Trim(ContratoIni.Text)) = 0 Then Exit Sub

    objContrato.iFilialEmpresa = giFilialEmpresa
    objContrato.sCodigo = ContratoIni.Text

    If lErro = SUCESSO Then
    
        lErro = Traz_ContratoIni_Tela(objContrato)
        If lErro <> SUCESSO Then gError 129756
        
    End If

    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_ContratoIni_Validate:

    Cancel = True

    Select Case gErr

        Case 129756

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155049)

    End Select
    
End Sub


Private Sub DataEmissao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub DataVencimento_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoNFiscal_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub



'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ContratoIni Then
            Call ContratoIniLabel_Click
        ElseIf Me.ActiveControl Is ContratoFim Then
            Call ContratoFimLabel_Click
        ElseIf Me.ActiveControl Is ClienteIni Then
            Call ClienteIniLabel_Click
        ElseIf Me.ActiveControl Is ClienteFim Then
            Call ClienteFimLabel_Click
        End If
          
    End If

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

Private Sub objEventoContratoIni_evSelecao(obj1 As Object)

Dim objContrato As ClassContrato
Dim bCancel As Boolean

    Set objContrato = obj1

    ContratoIni.Text = objContrato.sCodigo

    Call ContratoIni_Validate(bCancel)

    Exit Sub
    
End Sub

Private Sub objEventoContratoFim_evSelecao(obj1 As Object)

Dim objContrato As ClassContrato
Dim bCancel As Boolean

    Set objContrato = obj1

    ContratoFim.Text = objContrato.sCodigo

    Call ContratoFim_Validate(bCancel)

    Exit Sub
    
End Sub

Private Sub objEventoClienteIni_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteIni.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call ClienteIni_Validate(bCancel)

    Exit Sub

End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    ClienteFim.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call ClienteFim_Validate(bCancel)

    Exit Sub

End Sub

Private Sub ClienteIniLabel_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_ClienteIniLabel_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteIni.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(ClienteIni.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = ClienteIni.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteIni, "", sOrdenacao)

    Exit Sub
    
Erro_ClienteIniLabel_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155050)
    
    End Select
    
End Sub

Private Sub ClienteFimLabel_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_ClienteFimLabel_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(ClienteFim.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(ClienteFim.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = ClienteFim.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoClienteFim, "", sOrdenacao)

    Exit Sub
    
Erro_ClienteFimLabel_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155051)
    
    End Select
    
End Sub

Private Sub ContratoIniLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(ContratoIni.Text)) > 0 Then
        objContrato.sCodigo = ContratoIni.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    End If
    
    

    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoContratoIni)

End Sub

Private Sub ContratoFimLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(ContratoFim.Text)) > 0 Then
        objContrato.sCodigo = ContratoFim.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    End If

    

    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoContratoFim)

End Sub

Private Function Traz_ContratoIni_Tela(objContrato As ClassContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_ContratoIni_Tela

    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129757
    
    'Contrato Não Cadastrado
    If lErro = 129332 Then gError 129758
    
    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132902
   
    With objContrato
                   
        ContratoIni.Text = .sCodigo
        DescContratoIni.Caption = .sDescricao
   
    End With
         
    Traz_ContratoIni_Tela = SUCESSO

    Exit Function

Erro_Traz_ContratoIni_Tela:
     
    Traz_ContratoIni_Tela = gErr

    Select Case gErr
    
        Case 129757

        Case 129758
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, ContratoIni.Text)
        
        Case 132902
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155052)

    End Select

    Exit Function

End Function

Private Function Traz_ContratoFim_Tela(objContrato As ClassContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_ContratoFim_Tela
    
    lErro = CF("Contrato_Le", objContrato)
    If lErro <> SUCESSO And lErro <> 129332 Then gError 129759
    
    'Contrato Não Cadastrado
    If lErro = 129332 Then gError 129760

    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132903

    ContratoFim.Text = objContrato.sCodigo
    DescContratoFim.Caption = objContrato.sDescricao
         
    Traz_ContratoFim_Tela = SUCESSO

    Exit Function

Erro_Traz_ContratoFim_Tela:
     
    Traz_ContratoFim_Tela = gErr

    Select Case gErr
    
        Case 129759
        
        Case 129760
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, ContratoFim.Text)

        Case 132903
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155053)

    End Select

    Exit Function

End Function

Private Sub DataEmissao_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 129763

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 129763

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155054)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataVencimento, iAlterado)

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataVencimento.Text)
    If lErro <> SUCESSO Then gError 129764

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True

    Select Case gErr

        Case 129764

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155055)

    End Select

    Exit Sub

End Sub

Private Sub DataCobrIni_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCobrIni_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataCobrIni, iAlterado)

End Sub

Private Sub DataCobrIni_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCobrIni_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataCobrIni.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataCobrIni.Text)
    If lErro <> SUCESSO Then gError 136226

    Exit Sub

Erro_DataCobrIni_Validate:

    Cancel = True

    Select Case gErr

        Case 136226

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155056)

    End Select

    Exit Sub

End Sub

Private Sub DataCobrFim_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCobrFim_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataCobrFim, iAlterado)

End Sub

Private Sub DataCobrFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCobrFim_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataCobrFim.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataCobrFim.Text)
    If lErro <> SUCESSO Then gError 136227

    Exit Sub

Erro_DataCobrFim_Validate:

    Cancel = True

    Select Case gErr

        Case 136227

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155057)

    End Select

    Exit Sub

End Sub

'####################################################
'INÍCIO DOS BOTÕES UPDOWN
'####################################################
Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129769

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 129769

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155058)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 129770

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 129770

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155059)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 129771

    Exit Sub

Erro_UpDownVencimento_DownClick:

    Select Case gErr

        Case 129771

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155060)

    End Select

    Exit Sub

End Sub

Private Sub UpDownVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencimento_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 129772

    Exit Sub

Erro_UpDownVencimento_UpClick:

    Select Case gErr

        Case 129772

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155061)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCobrIni_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownCobrIni_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataCobrIni, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 136025

    Exit Sub

Erro_UpDownCobrIni_DownClick:

    Select Case gErr

        Case 136025

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155062)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCobrIni_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownCobrIni_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataCobrIni, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 136026

    Exit Sub

Erro_UpDownCobrIni_UpClick:

    Select Case gErr

        Case 136026

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155063)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCobrFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownCobrFim_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataCobrFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 136027

    Exit Sub

Erro_UpDownCobrFim_DownClick:

    Select Case gErr

        Case 136027

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155064)

    End Select

    Exit Sub

End Sub

Private Sub UpDownCobrFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownCobrFim_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataCobrFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 136028

    Exit Sub

Erro_UpDownCobrFim_UpClick:

    Select Case gErr

        Case 136028

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155065)

    End Select

    Exit Sub

End Sub
'####################################################
'FIM DOS BOTÕES UPDOWN
'####################################################


'#######################################################################
'INÍCIO DO SCRIPT PARA MODO DE EDICAO
'#######################################################################
Private Sub DescContratoIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescContratoIni, Source, X, Y)
End Sub

Private Sub DescContratoIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescContratoIni, Button, Shift, X, Y)
End Sub

Private Sub DescContratoFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescContratoFim, Source, X, Y)
End Sub

Private Sub DescContratoFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescContratoFim, Button, Shift, X, Y)
End Sub

Private Sub ContratoIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContratoIni, Source, X, Y)
End Sub

Private Sub ContratoIniLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContratoIniLabel, Source, X, Y)
End Sub

Private Sub ContratoIniLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContratoIniLabel, Button, Shift, X, Y)
End Sub

Private Sub ContratoFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContratoFim, Source, X, Y)
End Sub

Private Sub ContratoFimLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContratoFimLabel, Source, X, Y)
End Sub

Private Sub ContratoFimLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContratoFimLabel, Button, Shift, X, Y)
End Sub

Private Sub ClienteIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteIni, Source, X, Y)
End Sub

Private Sub ClienteIniLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteIniLabel, Source, X, Y)
End Sub

Private Sub ClienteIniLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteIniLabel, Button, Shift, X, Y)
End Sub

Private Sub ClienteFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteFim, Source, X, Y)
End Sub

Private Sub ClienteFimLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteFimLabel, Source, X, Y)
End Sub

Private Sub ClienteFimLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteFimLabel, Button, Shift, X, Y)
End Sub

Private Sub DataVencimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVencimento, Source, X, Y)
End Sub

Private Sub DataVencimentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVencimentoLabel, Source, X, Y)
End Sub

Private Sub DataVencimentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataVencimentoLabel, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissaoLabel, Source, X, Y)
End Sub

Private Sub DataEmissaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissaoLabel, Button, Shift, X, Y)
End Sub

Private Sub TipoNFiscal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoNFiscal, Source, X, Y)
End Sub

Private Sub TipoNFiscalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoNFiscalLabel, Source, X, Y)
End Sub

Private Sub TipoNFiscalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoNFiscalLabel, Button, Shift, X, Y)
End Sub
'#######################################################################
'FIM DO SCRIPT PARA MODO DE EDICAO
'#######################################################################

Private Sub Limpa_Tela_Faturamento()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Faturamento

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
          
    DescContratoIni.Caption = ""
    DescContratoFim.Caption = ""
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    TipoNFiscal.Text = DOCINFO_NFISFS & SEPARADOR & STRING_NFFATURA
          
    iAlterado = 0
 
    Exit Sub

Erro_Limpa_Tela_Faturamento:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155066)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboTipo(objCombo As ComboBox) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_ComboTipo
   
    objCombo.AddItem DOCINFO_NFISFS & SEPARADOR & STRING_NFFATURA
    objCombo.ItemData(TipoNFiscal.NewIndex) = DOCINFO_NFISFS

'    objCombo.AddItem DOCINFO_RECIBO & SEPARADOR & STRING_RECIBO
'    objCombo.ItemData(TipoNFiscal.NewIndex) = DOCINFO_RECIBO

    objCombo.Text = DOCINFO_NFISFS & SEPARADOR & STRING_NFFATURA

    Carrega_ComboTipo = SUCESSO

    Exit Function

Erro_Carrega_ComboTipo:

    Carrega_ComboTipo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155067)

    End Select

    Exit Function
    
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim sNomeArqParam As String
Dim objGeracaoFatContrato As New ClassGeracaoFatContrato

On Error GoTo Erro_Gravar_Registro

    lErro = Formata_E_Critica_Dados(objGeracaoFatContrato)
    If lErro <> SUCESSO Then gError 129797
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 129923
    
    lErro = CF("Rotina_NFiscalContrato_Gera", sNomeArqParam, objGeracaoFatContrato)
    If lErro <> SUCESSO Then gError 129781
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 129781, 129797, 129923
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155068)

    End Select

    Exit Function

End Function

Private Function PreencherRelOp() As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lNumIntRel As Long
Dim objGeracaoFatContrato As New ClassGeracaoFatContrato
Dim sNomeArqParam As String
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Dados(objGeracaoFatContrato)
    If lErro <> SUCESSO Then gError 129785

'    lErro = Sistema_Preparar_Batch(sNomeArqParam)
'    If lErro <> SUCESSO Then gError 129922
'
'    lErro = CF("Rotina_NFiscalContrato_Prepara", sNomeArqParam, lNumIntRel, objGeracaoFatContrato)
'    If lErro <> SUCESSO Then gError 129786
    
    lErro = CF("NFiscalContrato_Prepara", lNumIntRel, objGeracaoFatContrato)
    If lErro <> SUCESSO Then gError 129786
    
    lErro = objRelatorio.ExecutarDireto("Faturamento de Contratos", "", 0, "RELCONT", "DDATAEMISSAO", CStr(objGeracaoFatContrato.dtDataEmissao), "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 129787
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 129785, 129786, 129992, 129787

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155069)

    End Select

End Function


Public Function Formata_E_Critica_Dados(objGeracaoFatContrato As ClassGeracaoFatContrato) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Dados

    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 129773
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 129774
    If Len(Trim(DataVencimento.ClipText)) = 0 Then gError 129775
    If Len(Trim(DataCobrIni.ClipText)) = 0 Then gError 129776
    If Len(Trim(DataCobrFim.ClipText)) = 0 Then gError 129777
    
    If StrParaDate(DataCobrFim.Text) < StrParaDate(DataCobrIni.Text) Then gError 129778
    
    If Len(Trim(ContratoFim.Text)) <> 0 And Len(Trim(ContratoIni.Text)) Then
        If ContratoFim.Text < ContratoIni.Text Then gError 129779
    End If
    
    lErro = Move_Tela_Memoria(objGeracaoFatContrato)
    If lErro <> SUCESSO Then gError 129782
    
    If objGeracaoFatContrato.lClienteIni <> 0 And objGeracaoFatContrato.lClienteFim <> 0 Then
        If objGeracaoFatContrato.lClienteFim < objGeracaoFatContrato.lClienteIni Then gError 129780
    End If
        
    If objGeracaoFatContrato.dtDataRefVencimento < objGeracaoFatContrato.dtDataEmissao Then gError 201400
    
    Formata_E_Critica_Dados = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Dados:

    Formata_E_Critica_Dados = gErr

    Select Case gErr
        
        Case 129773
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case 129774
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 129775
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCIMENTO_NAO_PREENCHIDA", gErr)
        
        Case 129776, 129777
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_REFERENCIA_NAO_PREENCHIDA", gErr)
                
        Case 129778
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case 129779
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATOINICIAL_MAIOR_CONTRATOFINAL", gErr)
        
        Case 129780
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)

        Case 129782
        
        Case 201400
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCIMENTO_MENOR_EMISSAO_FATCONTRATO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155070)

    End Select

    Exit Function

End Function

Private Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_TipoNFiscal_Validate

    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub

    If TipoNFiscal.Text = TipoNFiscal.List(TipoNFiscal.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO Then gError 129878

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 129878
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCAL_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155071)

    End Select

    Exit Sub

End Sub
