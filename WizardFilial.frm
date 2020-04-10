VERSION 5.00
Begin VB.Form frmWizardFilial 
   Appearance      =   0  'Flat
   Caption         =   "Configuração"
   ClientHeight    =   5445
   ClientLeft      =   555
   ClientTop       =   915
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WizardFilial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8415
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   2
      Left            =   15
      TabIndex        =   7
      Tag             =   "2006"
      Top             =   15
      Width           =   8310
      Begin VB.ComboBox EstoqueAno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardFilial.frx":014A
         Left            =   4350
         List            =   "WizardFilial.frx":0172
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1635
         Width           =   855
      End
      Begin VB.ComboBox EstoqueMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardFilial.frx":01BE
         Left            =   1425
         List            =   "WizardFilial.frx":01E9
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1635
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
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
         Index           =   1
         Left            =   3855
         TabIndex        =   17
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mês:"
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
         Index           =   0
         Left            =   930
         TabIndex        =   18
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label8 
         Caption         =   $"WizardFilial.frx":0252
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   8
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   5040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Estoque"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   195
         TabIndex        =   20
         Top             =   135
         Width           =   2355
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1770
         Index           =   10
         Left            =   5520
         Picture         =   "WizardFilial.frx":02FB
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 0"
      Enabled         =   0   'False
      Height          =   4830
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8310
      Begin VB.Label Label14 
         Caption         =   $"WizardFilial.frx":E905
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   150
         TabIndex        =   12
         Top             =   3105
         Width           =   7905
      End
      Begin VB.Label Label12 
         Caption         =   "As próximas telas permitirão que você configure o funcionamento do sistema de acordo com as opções escolhidas."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   3000
         TabIndex        =   13
         Top             =   1875
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "A Configuração da Filial está sendo iniciada."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   3000
         TabIndex        =   14
         Top             =   375
         Width           =   5055
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2145
         Index           =   0
         Left            =   120
         Picture         =   "WizardFilial.frx":E9D8
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Termino da Instalação"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Tag             =   "3000"
      Top             =   15
      Width           =   8310
      Begin VB.Label Label10 
         Caption         =   "Pressione o botão ""Terminar"" para que suas configurações sejam gravadas."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   780
         TabIndex        =   15
         Top             =   2655
         Width           =   4275
      End
      Begin VB.Label lblStep 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "A Configuração da Filial está encerrada. "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   5
         Left            =   780
         TabIndex        =   16
         Tag             =   "3001"
         Top             =   630
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   5
         Left            =   5655
         Picture         =   "WizardFilial.frx":20472
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.Frame fraStep 
      Caption         =   "Frame5"
      Height          =   1815
      Index           =   0
      Left            =   -10000
      TabIndex        =   8
      Top             =   375
      Width           =   2490
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   4875
      Width           =   8415
      Begin VB.CommandButton cmdNav 
         Caption         =   "Terminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   7140
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Prosseguir >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   5745
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< Voltar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   4620
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   3450
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Ajuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   105
         X2              =   8254
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   8254
         Y1              =   30
         Y2              =   30
      End
   End
End
Attribute VB_Name = "frmWizardFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 4

Const MENSAGEM_TERMINO_CONFIG_FILIAL1 = "A Configuração da Filial "
Const MENSAGEM_TERMINO_CONFIG_FILIAL2 = " da Empresa "
Const MENSAGEM_TERMINO_CONFIG_FILIAL3 = " está encerrada."
Const MENSAGEM_INICIO_CONFIG_FILIAL1 = "A Configuração da Filial "
Const MENSAGEM_INICIO_CONFIG_FILIAL2 = " da Empresa "
Const MENSAGEM_INICIO_CONFIG_FILIAL3 = " está sendo iniciada."


Const RES_ERROR_MSG = 30000

'BASE VALUE FOR HELP FILE FOR THIS WIZARD:
Const HELP_BASE = 1000
Const HELP_FILE = "MYWIZARD.HLP"

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_FINISH = 3

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Configuração da Filial "
Const INTRO_KEY = "Tela de Introdução"
Const SHOW_INTRO = "Exibir Introdução"
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean

Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer
Dim objConfiguraADM1 As ClassConfiguraADM

Private Sub cmdNav_Click(Index As Integer)
    
Dim nAltStep As Integer
Dim lHelpTopic As Long
Dim rc As Long
Dim lErro As Long
    
On Error GoTo Erro_cmdNav_Click

    Select Case Index
        Case BTN_HELP
            mbHelpStarted = True
            lHelpTopic = HELP_BASE + 10 * (1 + mnCurStep)
            rc = WinHelp(Me.hwnd, HELP_FILE, HELP_CONTEXT, lHelpTopic)
        
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
LABEL_BTN_BACK:
            nAltStep = mnCurStep - 1
            lErro = SetStep(nAltStep, DIR_BACK)
            If lErro = 44865 Then GoTo LABEL_BTN_BACK
            
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
LABEL_BTN_NEXT:
            nAltStep = mnCurStep + 1
            lErro = SetStep(nAltStep, DIR_NEXT)
            If lErro = 44865 Then GoTo LABEL_BTN_NEXT
            
            
        Case BTN_FINISH
      
            lErro = Gravar_Registro()
            If lErro <> SUCESSO Then Error 44847
            
            objConfiguraADM1.iConfiguracaoOK = True
            
            Unload Me
            
'            If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
'                frmConfirm.Show vbModal
'            End If
        
    End Select
    
    Exit Sub
    
Erro_cmdNav_Click:

    Select Case Err

        Case 44847

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175897)

    End Select

    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        cmdNav_Click BTN_HELP
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim iMesAtual As Integer
    Dim iAnoAtual As Integer
    'init all vars
    mbFinishOK = False
    
    iMesAtual = Month(Date)
    
    'seleciona o mes atual como o mes de inicializacao do estoque
    For i = 0 To EstoqueMes.ListCount - 1
        If EstoqueMes.ItemData(i) = iMesAtual Then
            EstoqueMes.ListIndex = i
            Exit For
        End If
    Next
    
    iAnoAtual = Year(Date)
    
    'seleciona o ano atual como o ano de inicializacao do estoque
    For i = 0 To EstoqueAno.ListCount - 1
        If CInt(EstoqueAno.List(i)) = iAnoAtual Then
            EstoqueAno.ListIndex = i
            Exit For
        End If
    Next
    
    For i = STEP_1 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    'Load All string info for Form
    LoadResStrings Me
    
    'Determine 1st Step:
    If GetSetting(APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString) = SHOW_INTRO Then
        SetStep STEP_INTRO, DIR_NEXT
    Else
        SetStep STEP_1, DIR_NONE
    End If
    
    
End Sub

Private Function SetStep(nStep As Integer, nDirection As Integer) As Long
  
Dim lErro As Long
  
On Error GoTo Erro_SetSetp
  
    Select Case nStep
    
        Case STEP_INTRO
            
        Case STEP_1
            Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA
            Label11.Caption = MENSAGEM_INICIO_CONFIG_FILIAL1 & gsNomeFilialEmpresa & MENSAGEM_INICIO_CONFIG_FILIAL2 & gsNomeEmpresa & MENSAGEM_INICIO_CONFIG_FILIAL3
      
        Case STEP_2
            Me.HelpContextID = IDH_CONFIGURACAO_FILIAL_EMPRESA_EST
        
        Case STEP_FINISH
            lblStep(5).Caption = MENSAGEM_TERMINO_CONFIG_FILIAL1 & gsNomeFilialEmpresa & MENSAGEM_TERMINO_CONFIG_FILIAL2 & gsNomeEmpresa & MENSAGEM_TERMINO_CONFIG_FILIAL3
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
    SetStep = SUCESSO

    Exit Function

Erro_SetSetp:

    SetStep = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175898)

    End Select

    Exit Function
  
End Function

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = STEP_1 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    Me.Caption = FRM_TITLE & gsNomeFilialEmpresa & " da Empresa " & gsNomeEmpresa
'    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
'    If chkSaveSettings(0).Value = vbChecked Then
      
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
      
'    End If
    Set objConfiguraADM1 = Nothing
    
    If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

Private Function Gravar_Registro() As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lTransacaoDic As Long
Dim lConexao As Long

On Error GoTo Erro_Gravar_Registro
    
    iAlterado = 0
    
    lConexao = GL_lConexaoDic
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexao)
    If lTransacaoDic = 0 Then Error 44961
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44867
    
    lErro = CTB_Exercicio_Gravar_Registro()
    If lErro <> SUCESSO Then Error 44868
    
    lErro = CR_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then Error 41930
    
    lErro = EST_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then Error 41931
    
    lErro = FAT_Filial_Gravar_Registro()
    If lErro <> SUCESSO Then Error 41932
    
    lErro = CF("ModuloFilEmp_Atualiza_Configurado",glEmpresa, giFilialEmpresa, objConfiguraADM1.colModulosConfigurar)
    If lErro <> SUCESSO Then Error 44957
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 44869
    
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then Error 44962
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = Err
    
    Select Case Err

        Case 44867
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 44868, 44957, 44961, 44962, 41930, 41931, 41932

        Case 44869
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175899)

    End Select

    If Err <> 44962 Then Call Transacao_Rollback
    Call Transacao_RollbackExt(lTransacaoDic)

    Exit Function
    
End Function

Private Function Valida_Step(sModulo As String) As Long

Dim vModulo As Variant

    For Each vModulo In objConfiguraADM1.colModulosConfigurar

        If sModulo = vModulo Then
            Valida_Step = SUCESSO
            Exit Function
        End If
        
    Next
    
    Valida_Step = 44863

End Function

Function Trata_Parametros(objConfiguraADM As ClassConfiguraADM) As Long

On Error GoTo Erro_Trata_Parametros

    Set objConfiguraADM1 = objConfiguraADM
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175900)
    
    End Select
    
    Exit Function

End Function

Private Function CTB_Exercicio_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_CTB_Exercicio_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE)

    If lErro = SUCESSO Then
        
        lErro = CF("Exercicio_Instalacao_Filial",giFilialEmpresa)
        If lErro <> SUCESSO Then Error 44866
        
    End If
    
    CTB_Exercicio_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CTB_Exercicio_Gravar_Registro:
    
    CTB_Exercicio_Gravar_Registro = Err
    
    Select Case Err

        Case 44866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175901)

    End Select

    Exit Function
    
End Function

Private Function EST_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_EST_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_ESTOQUE)

    If lErro = SUCESSO Then
        
        lErro = CF("EST_Instalacao_Filial",giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41934
        
        objEstoqueMes.iFilialEmpresa = giFilialEmpresa
        objEstoqueMes.iAno = CInt(EstoqueAno.Text)
        objEstoqueMes.iMes = EstoqueMes.ItemData(EstoqueMes.ListIndex)
        
        lErro = CF("EstoqueMes_Insere",objEstoqueMes)
        If lErro <> SUCESSO Then Error 44969
        
    End If
    
    EST_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_EST_Filial_Gravar_Registro:
    
    EST_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41934, 44969

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175902)

    End Select

    Exit Function

End Function

Private Function CR_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_CR_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTASARECEBER)

    If lErro = SUCESSO Then
        
        lErro = CF("CR_Instalacao_Filial",giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41933
        
    End If
    
    CR_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CR_Filial_Gravar_Registro:
    
    CR_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175903)

    End Select

    Exit Function
    
End Function

Private Function FAT_Filial_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_FAT_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_FATURAMENTO)

    If lErro = SUCESSO Then
        
        lErro = CF("FAT_Instalacao_Filial",giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41935
        
    End If
    
    FAT_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_FAT_Filial_Gravar_Registro:
    
    FAT_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41935

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175904)

    End Select

    Exit Function
    
End Function




Private Sub lblStep_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(lblStep(Index), Source, X, Y)
End Sub

Private Sub lblStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblStep(Index), Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label3(Index), Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3(Index), Button, Shift, X, Y)
End Sub


Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

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

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

