VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl SpedFcontOcx 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ScaleHeight     =   2880
   ScaleWidth      =   6795
   Begin VB.Frame Frame3 
      Caption         =   "Demais Informações"
      Height          =   1020
      Left            =   150
      TabIndex        =   25
      Top             =   780
      Width           =   6420
      Begin VB.ComboBox SituacaoEspecial 
         Height          =   315
         ItemData        =   "SpedFcontOcx.ctx":0000
         Left            =   1740
         List            =   "SpedFcontOcx.ctx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   4305
      End
      Begin VB.ComboBox IndIniPer 
         Height          =   315
         ItemData        =   "SpedFcontOcx.ctx":004C
         Left            =   1740
         List            =   "SpedFcontOcx.ctx":005C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   4305
      End
      Begin MSMask.MaskEdBox NumOrd 
         Height          =   300
         Left            =   1605
         TabIndex        =   12
         Top             =   -165
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Situação Especial:"
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
         Left            =   90
         TabIndex        =   28
         Top             =   660
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ind. Iní. Período:"
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
         TabIndex        =   27
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número do Diário:"
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
         Left            =   45
         TabIndex        =   26
         Top             =   -135
         Visible         =   0   'False
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localização do arquivo"
      Height          =   990
      Left            =   150
      TabIndex        =   22
      Top             =   1830
      Width           =   6420
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   7
         Top             =   615
         Width           =   4185
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   1740
         TabIndex        =   6
         Top             =   255
         Width           =   4185
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5910
         TabIndex        =   8
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo:"
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
         Left            =   1005
         TabIndex        =   24
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
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
         Left            =   930
         TabIndex        =   23
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Outros"
      Height          =   3120
      Left            =   150
      TabIndex        =   19
      Top             =   2865
      Width           =   6420
      Begin MSMask.MaskEdBox ContaOutros 
         Height          =   315
         Left            =   1605
         TabIndex        =   13
         Top             =   195
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
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
      Begin MSComctlLib.TreeView TvwContas 
         Height          =   2055
         Left            =   180
         TabIndex        =   14
         Top             =   915
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   3625
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conta Outros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   315
         TabIndex        =   21
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label LabelPlanoConta 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Contas"
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
         Left            =   195
         TabIndex        =   20
         Top             =   615
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   105
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1110
         Picture         =   "SpedFcontOcx.ctx":0120
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   607
         Picture         =   "SpedFcontOcx.ctx":029E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   345
         Left            =   105
         Picture         =   "SpedFcontOcx.ctx":07D0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gera o arquivo"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   690
      Left            =   150
      TabIndex        =   15
      Top             =   0
      Width           =   4740
      Begin MSComCtl2.UpDown UpDownPeriodoDe 
         Height          =   330
         Left            =   1665
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoDe 
         Height          =   315
         Left            =   675
         TabIndex        =   0
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownPeriodoAte 
         Height          =   330
         Left            =   3750
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoAte 
         Height          =   330
         Left            =   2775
         TabIndex        =   2
         Top             =   255
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPeriodoAte 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   2355
         TabIndex        =   17
         Top             =   300
         Width           =   450
      End
      Begin VB.Label LabelPeriodoDe 
         Appearance      =   0  'Flat
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
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   300
         Width           =   390
      End
   End
End
Attribute VB_Name = "SpedFcontOcx"
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

Dim iListIndexDefault As Integer


'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim sDiretorio As String
Dim dtData As Date
Dim lNumOrd As Long
Dim sNomeArqParam As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_BotaoGerar_Click
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 203084
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 203085
    'If Len(Trim(NumOrd.Text)) = 0 Then gError 203111
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text & NomeArquivo.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\" & NomeArquivo.Text
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatorios estao preenchidos
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then gError 203086
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then gError 203087
    If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 203088
    
    If IndIniPer.ItemData(IndIniPer.ListIndex) = 0 Then
        If Month(StrParaDate(PeriodoDe.Text)) <> 1 Or Day(StrParaDate(PeriodoDe.Text)) <> 1 Then gError 211079
    End If
    If Codigo_Extrai(SituacaoEspecial.Text) = 0 Then
        If Month(StrParaDate(PeriodoAte.Text)) <> 12 Or Day(StrParaDate(PeriodoAte.Text)) <> 31 Then gError 211080
    End If
    
    lNumOrd = StrParaLong(NumOrd.Text)
    
    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaOutros.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then gError 207367
            
    'conta não cadastrada
    If lErro = 5700 Then gError 207368
    
    If Len(ContaOutros.ClipText) > 0 Then
        If objPlanoConta.iNaturezaSped <> 9 Then gError 207371
    End If
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 203089
    
    lErro = CF("Rotina_Sped_FCont", sNomeArqParam, giFilialEmpresa, sDiretorio, StrParaDate(PeriodoDe.Text), StrParaDate(PeriodoAte.Text), lNumOrd, objPlanoConta.sConta, IndIniPer.ItemData(IndIniPer.ListIndex), Codigo_Extrai(SituacaoEspecial.Text))
    If lErro <> SUCESSO Then gError 203253
        
    Call BotaoLimpar_Click
   
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub
    
Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 203084
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_INFORMADO", gErr)
        
        Case 203085
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
        
        Case 203086
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA1", gErr)
        
        Case 203087
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case 203088
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 203089, 203253, 207367
        
        Case 203111
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_DIARIO_NAO_INFORMADO", gErr)
        
        Case 207368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err)
        
        Case 207371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NATUREZA_SPED_OUTROS", Err)
            
        Case 211079
            Call Rotina_Erro(vbOKOnly, "ERRO_FCONT_REGRA_DT_INICIO_ESCRITURACAO", gErr)
        
        Case 211080
            Call Rotina_Erro(vbOKOnly, "ERRO_FCONT_REGRA_DT_FINAL_ESCRITURACAO", gErr)
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203090)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)
    
    IndIniPer.ListIndex = 1
    SituacaoEspecial.ListIndex = -1

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
   
    NomeDiretorio.Text = CurDir
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203091)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraConta As String

On Error GoTo Erro_Form_Load
    
    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
    If lErro <> SUCESSO Then gError 207369

    
    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 207370
    
    ContaOutros.Mask = sMascaraConta
    
    IndIniPer.ListIndex = 1
    
    
'    iListIndexDefault = Drive1.ListIndex
    
'    If Len(Trim(CurDir)) > 0 Then
'        Dir1.Path = CurDir
'        Drive1.Drive = Left(CurDir, 2)
'    End If
'
'    NomeDiretorio.Text = Dir1.Path
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case 207369, 207370
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203092)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203093)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203094)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Sped Diário"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "SpedDiario"

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


Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização do arquivos"
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

Private Sub NumOrd_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumOrd, iAlterado)
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

Function Trata_Parametros(Optional obj1 As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203095)

    End Select

    Exit Function

End Function

'Private Sub Dir1_Change()
'
'     NomeDiretorio.Text = Dir1.Path
'
'End Sub

'Private Sub Drive1_Change()
'
'On Error GoTo Erro_Drive1_Change
'
'    Dir1.Path = Drive1.Drive
'
'    Exit Sub
'
'Erro_Drive1_Change:
'
'    Select Case Err
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 203096)
'
'    End Select
'
'    Drive1.ListIndex = iListIndexDefault
'
'    Exit Sub
'
'End Sub
'
'Private Sub Drive1_GotFocus()
'
'    iListIndexDefault = Drive1.ListIndex
'
'End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 203097

'    Drive1.Drive = Mid(NomeDiretorio.Text, 1, 2)
'
'    Dir1.Path = NomeDiretorio.Text

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 76, 203097
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203098)

    End Select

    Exit Sub

End Sub


Private Sub PeriodoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoDe_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoDe.Text)
    If lErro <> SUCESSO Then gError 203099

    Exit Sub

Erro_PeriodoDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 203099
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203100)
            
    End Select
    
    Exit Sub

End Sub

Private Sub PeriodoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoAte_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoAte.Text)
    If lErro <> SUCESSO Then gError 203101

    Exit Sub

Erro_PeriodoAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 203101
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203102)
            
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 203103

    End If

    Exit Sub

Erro_UpDownPeriodoDe_DownClick:

    Select Case gErr

        Case 203103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203104)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 203105

    End If

    Exit Sub

Erro_UpDownPeriodoDe_UpClick:

    Select Case gErr

        Case 203105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203106)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 203107

    End If

    Exit Sub

Erro_UpDownPeriodoAte_DownClick:

    Select Case gErr

        Case 203107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203108)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 203109

    End If

    Exit Sub

Erro_UpDownPeriodoAte_UpClick:

    Select Case gErr

        Case 203109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203110)

    End Select

    Exit Sub

End Sub

Private Sub ContaOutros_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaOutros_Validate

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaOutros.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then gError 207363
            
    'conta não cadastrada
    If lErro = 5700 Then gError 207364
    
    If Len(ContaOutros.ClipText) > 0 Then
        If objPlanoConta.iNaturezaSped <> 5 Then gError 207366
    End If
    
    Exit Sub

Erro_ContaOutros_Validate:

    Cancel = True


    Select Case gErr
    
        Case 207363
        
        Case 207364
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaOutros.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Chama_Tela("PlanoConta", objPlanoConta)

            End If
            
        Case 207366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NATUREZA_SPED_OUTROS", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 207365)
    
    End Select
    
    Exit Sub

End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim cControl As Control
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    If sCaracterInicial <> "A" Then Error 20299
    
    sConta = right(Node.Key, Len(Node.Key) - 1)
    
    sContaEnxuta = String(STRING_CONTA, 0)

    'volta mascarado apenas os caracteres preenchidos
    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then Error 20300

    ContaOutros.PromptInclude = False
    ContaOutros.Text = sContaEnxuta
    ContaOutros.PromptInclude = True

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 20299

        Case 20300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143122)

    End Select

    Exit Sub

End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 40798
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40798
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143123)
        
    End Select
        
    Exit Sub
    
End Sub

