VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SpedECFOcx 
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ScaleHeight     =   1905
   ScaleWidth      =   6795
   Begin VB.Frame Frame2 
      Caption         =   "Localização do arquivo"
      Height          =   990
      Left            =   150
      TabIndex        =   19
      Top             =   750
      Width           =   6495
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   5
         Top             =   615
         Width           =   4185
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   1740
         TabIndex        =   4
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
         TabIndex        =   6
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Outros"
      Height          =   3120
      Left            =   150
      TabIndex        =   16
      Top             =   2865
      Width           =   6420
      Begin MSMask.MaskEdBox ContaOutros 
         Height          =   315
         Left            =   1605
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   615
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   105
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1110
         Picture         =   "SpedECFOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   607
         Picture         =   "SpedECFOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   345
         Left            =   105
         Picture         =   "SpedECFOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Gera o arquivo"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   690
      Left            =   150
      TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   300
         Width           =   390
      End
   End
End
Attribute VB_Name = "SpedECFOcx"
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
Dim sNomeArqParam As String
Dim objSped As New ClassSpedECFSel

On Error GoTo Erro_BotaoGerar_Click
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 213949
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 213950
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text & NomeArquivo.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\" & NomeArquivo.Text
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatorios estao preenchidos
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then gError 213951
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then gError 213952
    If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 213953
    
    objSped.sDiretorio = sDiretorio
    objSped.iFilialEmpresa = giFilialEmpresa
    objSped.dtDataIni = StrParaDate(PeriodoDe.Text)
    objSped.dtDataFim = StrParaDate(PeriodoAte.Text)
       
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("Rotina_Sped_ECF", sNomeArqParam, objSped)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Call BotaoLimpar_Click
   
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub
    
Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 213949
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_INFORMADO", gErr)
        
        Case 213950
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
        
        Case 213951
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA1", gErr)
        
        Case 213952
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case 213953
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213954)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)
    
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
   
    Call Define_Default
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213955)

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

On Error GoTo Erro_Form_Load

    Call Define_Default
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213956)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213957)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213958)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Sped ECF"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "SpedECF"

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213959)

    End Select

    Exit Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213960)

    End Select

    Exit Function

End Function

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 213961

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 76, 213961
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213962)

    End Select

    Exit Sub

End Sub

Private Sub PeriodoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoDe_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoDe.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_PeriodoDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213964)
            
    End Select
    
    Exit Sub

End Sub

Private Sub PeriodoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoAte_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoAte.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_PeriodoAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213965)
            
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_UpDownPeriodoDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213966)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_UpDownPeriodoDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213967)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_UpDownPeriodoAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213968)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_UpDownPeriodoAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213969)

    End Select

    Exit Sub

End Sub

Private Sub Define_Default()
Dim dtDataIni As Date, dtDataFim As Date

    dtDataIni = StrParaDate("01/01/" & CStr(Year(gdtDataAtual) - 1))
    dtDataFim = StrParaDate("31/12/" & CStr(Year(gdtDataAtual) - 1))
    NomeDiretorio.Text = CurDir & "\"
    Call DateParaMasked(PeriodoDe, dtDataIni)
    Call DateParaMasked(PeriodoAte, dtDataFim)
    NomeArquivo.Text = "ECF_" & CStr(Year(gdtDataAtual) - 1) & ".txt"
End Sub
