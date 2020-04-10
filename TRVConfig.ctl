VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVConfig 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ScaleHeight     =   5370
   ScaleWidth      =   6165
   Begin VB.Frame Frame5 
      Caption         =   "Vouchers Pago no Cartão com valor acima"
      Height          =   645
      Left            =   30
      TabIndex        =   34
      Top             =   4620
      Width           =   5970
      Begin MSMask.MaskEdBox PercFatorDevCMCC 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   240
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "##0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "do valor junto a CMCC"
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
         Left            =   1785
         TabIndex        =   37
         Top             =   300
         Width           =   1920
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Devolver"
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
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bloqueios para o faturamento"
      Height          =   585
      Left            =   30
      TabIndex        =   31
      Top             =   3990
      Width           =   5970
      Begin MSMask.MaskEdBox PrazoMinPagto 
         Height          =   315
         Left            =   4260
         TabIndex        =   13
         Top             =   150
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "dias"
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
         Height          =   315
         Index           =   6
         Left            =   4305
         TabIndex        =   33
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Prazo mínimo para faturamento de a pagar:"
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
         Height          =   315
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   210
         Width           =   4020
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Liberação de valores para assistência"
      Height          =   630
      Left            =   30
      TabIndex        =   28
      Top             =   3330
      Width           =   5970
      Begin MSMask.MaskEdBox AssistData 
         Height          =   315
         Left            =   1245
         TabIndex        =   10
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownAssistData 
         Height          =   300
         Left            =   2565
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox AssistLimite 
         Height          =   315
         Left            =   4275
         TabIndex        =   12
         Top             =   210
         Width           =   1590
         _ExtentX        =   2805
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Limite diário:"
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
         Height          =   315
         Index           =   5
         Left            =   3060
         TabIndex        =   30
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Início:"
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
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   255
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   210
      Left            =   3945
      TabIndex        =   25
      Top             =   135
      Visible         =   0   'False
      Width           =   165
      Begin MSMask.MaskEdBox ProxTitPag 
         Height          =   315
         Left            =   1635
         TabIndex        =   26
         Top             =   45
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -210
         Top             =   -105
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Próximo número de título a pagar:"
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
         Height          =   315
         Index           =   0
         Left            =   510
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   105
         Visible         =   0   'False
         Width           =   3420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Geração de OVER"
      Height          =   870
      Left            =   30
      TabIndex        =   21
      Top             =   2445
      Width           =   5970
      Begin MSMask.MaskEdBox Cliente 
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Top             =   210
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercComiCliOver 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Top             =   525
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorTarifaCartaoOver 
         Height          =   285
         Left            =   4275
         TabIndex        =   9
         Top             =   525
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tarifa Cartão:"
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
         Left            =   3075
         TabIndex        =   24
         Top             =   585
         Width           =   1170
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "% de comis.:"
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
         TabIndex        =   23
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label LabelCliente 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton BotaoModeloFatCartao 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5460
      TabIndex        =   4
      Top             =   1515
      Width           =   555
   End
   Begin VB.TextBox ModeloFatCartao 
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   5385
   End
   Begin VB.TextBox ModeloFat 
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   5385
   End
   Begin VB.CommandButton BotaoModeloFat 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5460
      TabIndex        =   2
      Top             =   915
      Width           =   555
   End
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5445
      TabIndex        =   6
      Top             =   2070
      Width           =   555
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   2130
      Width           =   5400
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4215
      ScaleHeight     =   495
      ScaleWidth      =   1035
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   1095
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   555
         Picture         =   "TRVConfig.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRVConfig.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox ProxTitRec 
      Height          =   315
      Left            =   2415
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo p/ ger. de Fat. de Cartão e Nota de Crédito em html:"
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
      Index           =   3
      Left            =   75
      TabIndex        =   20
      Top             =   1335
      Width           =   5460
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo padrão para geração de Faturas em html:"
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
      Index           =   1
      Left            =   75
      TabIndex        =   19
      Top             =   720
      Width           =   4500
   End
   Begin VB.Label Label1 
      Caption         =   "Diretório padrão para geração das faturas html:"
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
      Height          =   210
      Index           =   2
      Left            =   75
      TabIndex        =   18
      Top             =   1905
      Width           =   4965
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Próximo número do Título:"
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
      Height          =   315
      Left            =   45
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "TRVConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

'Variáveis Globais
Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Form_Load

    Set objEventoCliente = New AdmEvento
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192315
    
    ProxTitRec.PromptInclude = False
    ProxTitRec.Text = sConteudo
    ProxTitRec.PromptInclude = True

    lErro = CF("TRVConfig_Le", TRVCONFIG_PROX_NUM_TITPAG, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192316

    ProxTitPag.PromptInclude = False
    ProxTitPag.Text = sConteudo
    ProxTitPag.PromptInclude = True

    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192317
    
    ModeloFat.Text = sConteudo

    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192317
    
    ModeloFatCartao.Text = sConteudo

    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    NomeDiretorio.Text = sConteudo
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_CLIENTE_OVER, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    If StrParaLong(sConteudo) <> 0 Then
        Cliente.Text = sConteudo
        Call Cliente_Validate(bSGECancelDummy)
    End If
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_PERC_COMIS_CLI_OVER, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    PercComiCliOver.Text = CStr(100 * StrParaDbl(sConteudo))

    lErro = CF("TRVConfig_Le", TRVCONFIG_TAR_CARTAO_NOVO_CLI_OVER, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    ValorTarifaCartaoOver.Text = Format(StrParaDbl(sConteudo), "STANDARD")
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_ASSISTENCIA_LIMITE_DIARIO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    AssistLimite.Text = Format(StrParaDbl(sConteudo), "STANDARD")
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_ASSISTENCIA_DATA_INICIO_LIB, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    Call DateParaMasked(AssistData, StrParaDate(sConteudo))
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_PRAZO_MIN_PARA_PAGTO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    PrazoMinPagto.PromptInclude = False
    PrazoMinPagto.Text = sConteudo
    PrazoMinPagto.PromptInclude = True
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_PERC_FATOR_DEV_CMCC, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    PercFatorDevCMCC.Text = CStr(100 * StrParaDbl(sConteudo))
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
    
        Case 192315 To 192318

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192319)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objEventoCliente = Nothing
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Configurações"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVConfig"

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

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 192320

    iAlterado = 0
    
    Unload Me
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 192320

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192321)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colConfig As New Collection

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Preenche o objFaturamento
    lErro = Move_Tela_Memoria(colConfig)
    If lErro <> SUCESSO Then gError 192322

    lErro = CF("TRVConfig_Grava", colConfig)
    If lErro <> SUCESSO Then gError 192323

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 192322, 192323

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192324)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colConfig As Collection) As Long

Dim lErro As Long
Dim objTRVConfig As ClassTRVConfig

On Error GoTo Erro_Move_Tela_Memoria

'    Set objTRVConfig = New ClassTRVConfig
'
'    objTRVConfig.sCodigo = TRVCONFIG_PROX_NUM_TITREC
'    objTRVConfig.sConteudo = ProxTitRec.Text
'
'    colConfig.Add objTRVConfig
'
'    Set objTRVConfig = New ClassTRVConfig
'
'    objTRVConfig.sCodigo = TRVCONFIG_PROX_NUM_TITPAG
'    objTRVConfig.sConteudo = ProxTitPag.Text
'
'    colConfig.Add objTRVConfig

    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_DIRETORIO_MODELO_FAT_HTML
    objTRVConfig.sConteudo = ModeloFat.Text
    
    colConfig.Add objTRVConfig
    
    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_DIRETORIO_FAT_HTML
    objTRVConfig.sConteudo = NomeDiretorio.Text
    
    colConfig.Add objTRVConfig
    
    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO
    objTRVConfig.sConteudo = ModeloFatCartao.Text
    
    colConfig.Add objTRVConfig

    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_CLIENTE_OVER
    objTRVConfig.sConteudo = CStr(LCodigo_Extrai(Cliente.Text))
    
    colConfig.Add objTRVConfig

    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_PERC_COMIS_CLI_OVER
    
    If Len(Trim(PercComiCliOver.Text)) > 0 Then
        objTRVConfig.sConteudo = CStr(CDbl(PercComiCliOver.Text) / 100)
    Else
        objTRVConfig.sConteudo = "0"
    End If
    
    colConfig.Add objTRVConfig
    
    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_TAR_CARTAO_NOVO_CLI_OVER
    objTRVConfig.sConteudo = ValorTarifaCartaoOver.Text
    
    colConfig.Add objTRVConfig

    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_ASSISTENCIA_DATA_INICIO_LIB
    objTRVConfig.sConteudo = Format(StrParaDate(AssistData.Text), "dd/mm/yyyy")
    
    colConfig.Add objTRVConfig

    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_ASSISTENCIA_LIMITE_DIARIO
    objTRVConfig.sConteudo = AssistLimite.Text
    
    colConfig.Add objTRVConfig
    
    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_PRAZO_MIN_PARA_PAGTO
    objTRVConfig.sConteudo = Trim(PrazoMinPagto.Text)
    
    colConfig.Add objTRVConfig
    
    Set objTRVConfig = New ClassTRVConfig
    
    objTRVConfig.sCodigo = TRVCONFIG_PERC_FATOR_DEV_CMCC
    
    If Len(Trim(PercFatorDevCMCC.Text)) > 0 Then
        objTRVConfig.sConteudo = CStr(CDbl(PercFatorDevCMCC.Text) / 100)
    Else
        objTRVConfig.sConteudo = "0"
    End If
    
    colConfig.Add objTRVConfig

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192325)

    End Select

End Function

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub BotaoModeloFat_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoModeloFat_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Html Files" & _
    "(*.html)|*.html"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    ModeloFat.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoModeloFat_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física dos arquivos .html"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPOS As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPOS = InStr(1, NomeDiretorio.Text, "/")
        If iPOS = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192328)

    End Select

    Exit Sub

End Sub

Private Sub BotaoModeloFatCartao_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoModeloFatCartao_Click
    
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Html Files" & _
    "(*.html)|*.html"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    ModeloFatCartao.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoModeloFatCartao_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le2(Cliente, objcliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190637)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTarifaCartaoOver_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorTarifaCartaoOver_GotFocus()
    Call MaskEdBox_TrataGotFocus(ValorTarifaCartaoOver, iAlterado)
End Sub

Private Sub ValorTarifaCartaoOver_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTarifaCartaoOver_Validate

    'Veifica se ValorTarifaCartaoOver está preenchida
    If Len(Trim(ValorTarifaCartaoOver.Text)) <> 0 Then

       'Critica a ValorTarifaCartaoOver
       lErro = Valor_Positivo_Critica(ValorTarifaCartaoOver.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    

    Exit Sub

Erro_ValorTarifaCartaoOver_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub PercComiCliOver_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiCliOver_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercComiCliOver, iAlterado)
End Sub

Private Sub PercComiCliOver_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiCliOver_Validate

    'Veifica se PercComiCliOver está preenchida
    If Len(Trim(PercComiCliOver.Text)) <> 0 Then

       'Critica a PercComiCliOver
       lErro = Porcentagem_Critica(PercComiCliOver.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    Exit Sub

Erro_PercComiCliOver_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub AssistLimite_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AssistLimite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AssistLimite_Validate

    'Veifica se AssistLimite está preenchida
    If Len(Trim(AssistLimite.Text)) <> 0 Then

       'Critica a AssistLimite
       lErro = Valor_Positivo_Critica(AssistLimite.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_AssistLimite_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownAssistData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAssistData_DownClick

    AssistData.SetFocus

    If Len(AssistData.ClipText) > 0 Then

        sData = AssistData.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 190609

        AssistData.Text = sData

    End If

    Exit Sub

Erro_UpDownAssistData_DownClick:

    Select Case gErr

        Case 190609

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190610)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAssistData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAssistData_UpClick

    AssistData.SetFocus

    If Len(Trim(AssistData.ClipText)) > 0 Then

        sData = AssistData.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 190611

        AssistData.Text = sData

    End If

    Exit Sub

Erro_UpDownAssistData_UpClick:

    Select Case gErr

        Case 190611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190612)

    End Select

    Exit Sub

End Sub

Private Sub AssistData_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AssistData, iAlterado)
    
End Sub

Private Sub AssistData_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AssistData_Validate

    If Len(Trim(AssistData.ClipText)) <> 0 Then

        lErro = Data_Critica(AssistData.Text)
        If lErro <> SUCESSO Then gError 190613

    End If

    Exit Sub

Erro_AssistData_Validate:

    Cancel = True

    Select Case gErr

        Case 190613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190614)

    End Select

    Exit Sub

End Sub

Private Sub AssistData_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercFatorDevCMCC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercFatorDevCMCC_GotFocus()
    Call MaskEdBox_TrataGotFocus(PercFatorDevCMCC, iAlterado)
End Sub

Private Sub PercFatorDevCMCC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercFatorDevCMCC_Validate

    'Veifica se PercComiCliOver está preenchida
    If Len(Trim(PercFatorDevCMCC.Text)) <> 0 Then

       'Critica a PercComiCliOver
       lErro = Porcentagem_Critica(PercFatorDevCMCC.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    
    Exit Sub

Erro_PercFatorDevCMCC_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub
