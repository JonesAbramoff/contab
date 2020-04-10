VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl SATConfig 
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   ScaleHeight     =   4455
   ScaleWidth      =   8550
   Begin VB.TextBox PortaImpressora 
      Height          =   330
      Left            =   6030
      MaxLength       =   20
      TabIndex        =   20
      ToolTipText     =   "Senha definida pelo contribuinte no software de ativação com 8 a 32 caracteres"
      Top             =   2595
      Width           =   735
   End
   Begin VB.ComboBox ModeloImpressora 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "SATConfig.ctx":0000
      Left            =   2205
      List            =   "SATConfig.ctx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2625
      Width           =   3075
   End
   Begin VB.CommandButton BotaoTesteImpressora 
      Caption         =   "Testar"
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
      Left            =   6960
      TabIndex        =   18
      Top             =   2580
      Width           =   930
   End
   Begin VB.CommandButton BotaoLogo 
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
      Left            =   6675
      TabIndex        =   17
      Top             =   3780
      Width           =   555
   End
   Begin VB.CommandButton BotaoNomeArqDll 
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
      Left            =   6675
      TabIndex        =   16
      Top             =   1995
      Width           =   555
   End
   Begin VB.CommandButton BotaoDirXml 
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
      Left            =   6675
      TabIndex        =   15
      Top             =   1470
      Width           =   555
   End
   Begin VB.TextBox NomeArqLogo 
      Height          =   375
      Left            =   2190
      MaxLength       =   200
      TabIndex        =   13
      Top             =   3795
      Width           =   4485
   End
   Begin VB.CheckBox EmTeste 
      Caption         =   "Em Teste"
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
      Left            =   3825
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.ComboBox LayoutImpressao 
      Height          =   315
      ItemData        =   "SATConfig.ctx":00A8
      Left            =   2220
      List            =   "SATConfig.ctx":00AF
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3225
      Width           =   3375
   End
   Begin VB.TextBox NomeArqDll 
      Height          =   375
      Left            =   2220
      MaxLength       =   200
      TabIndex        =   8
      Top             =   1995
      Width           =   4425
   End
   Begin VB.TextBox DirArqXml 
      Height          =   330
      Left            =   2220
      MaxLength       =   200
      TabIndex        =   6
      Top             =   1485
      Width           =   4440
   End
   Begin VB.CheckBox EmuladorSefaz 
      Caption         =   "Usando o emulador da Sefaz SP"
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
      Left            =   390
      TabIndex        =   5
      Top             =   345
      Width           =   3105
   End
   Begin VB.TextBox CodigoAtivacao 
      Height          =   330
      Left            =   2235
      MaxLength       =   32
      TabIndex        =   3
      ToolTipText     =   "Senha definida pelo contribuinte no software de ativação com 8 a 32 caracteres"
      Top             =   840
      Width           =   2955
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7020
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   270
      Width           =   1140
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "SATConfig.ctx":00BB
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   555
         Picture         =   "SATConfig.ctx":0215
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5235
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label12 
      Caption         =   "Impressora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      TabIndex        =   22
      Top             =   2670
      Width           =   1065
   End
   Begin VB.Label Label13 
      Caption         =   "Porta:"
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
      Left            =   5430
      TabIndex        =   21
      Top             =   2655
      Width           =   555
   End
   Begin VB.Label Label6 
      Caption         =   "Logo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1500
      TabIndex        =   14
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "Layout de Impressão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   3270
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "DLL do SAT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      TabIndex        =   9
      Top             =   2085
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Diretório XMLs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   7
      Top             =   1485
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Código de Ativação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   390
      TabIndex        =   4
      Top             =   915
      Width           =   1800
   End
End
Attribute VB_Name = "SATConfig"
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

Public Function SATConfig_Grava() As Long

Dim objConfiguracaoSAT As New ClassConfiguracaoSAT, lErro As Long

On Error GoTo Erro_SATConfig_Grava

    If Len(Trim(CodigoAtivacao.Text)) < 7 Or Len(Trim(CodigoAtivacao.Text)) > 32 Then gError 201548
    
    'grava na memória
    Call Move_Tela_Memoria(objConfiguracaoSAT)
    
    'grava no bd
    lErro = CF_ECF("ConfiguracaoSAT_Grava", objConfiguracaoSAT)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjSATInfo.Copia(objConfiguracaoSAT)
    
    SATConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_SATConfig_Grava:
    
    SATConfig_Grava = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 201548
            Call Rotina_ErroECF(vbOKOnly, ERRO_SAT_CODIGO_ATIVACAO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144089)
    
    End Select
    
    Exit Function

End Function

Private Sub Move_Tela_Memoria(objConfiguracaoSAT As ClassConfiguracaoSAT)

    With objConfiguracaoSAT
        .iEmuladorSefaz = EmuladorSefaz.Value
        .iEmTeste = EmTeste.Value
        .sCodigoDeAtivacao = Trim(CodigoAtivacao.Text)
        .sDirArqXml = Trim(DirArqXml.Text)
        .sNomeArqDLL = Trim(NomeArqDll.Text)
        .iModeloImpressora = ModeloImpressora.ItemData(ModeloImpressora.ListIndex)
        .iLayoutImpressao = LayoutImpressao.ItemData(LayoutImpressao.ListIndex)
        .sNomeArqLogo = Trim(NomeArqLogo.Text)
        .sPortaImpressora = Trim(PortaImpressora.Text)
    End With
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    'grava as configurações no arquivo e na memória
    lErro = SATConfig_Grava()
    If lErro <> SUCESSO Then gError 109486

    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 109486
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144091)
    
    End Select

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a gravar registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 109485
    
    'fecha a tela
    Unload Me

    Exit Sub
    
Erro_BotaoGravar_Click:
    
    Select Case gErr
    
        Case 109485
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144093)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Load()

Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    EmuladorSefaz.Value = gobjSATInfo.iEmuladorSefaz
    EmTeste.Value = gobjSATInfo.iEmTeste
    CodigoAtivacao.Text = gobjSATInfo.sCodigoDeAtivacao
    DirArqXml.Text = gobjSATInfo.sDirArqXml
    NomeArqDll.Text = gobjSATInfo.sNomeArqDLL
    Call Combo_Seleciona_ItemData(ModeloImpressora, gobjSATInfo.iModeloImpressora)
    Call Combo_Seleciona_ItemData(LayoutImpressao, gobjSATInfo.iLayoutImpressao)
    NomeArqLogo.Text = gobjSATInfo.sNomeArqLogo
    PortaImpressora.Text = gobjSATInfo.sPortaImpressora
       
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144094)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 144095)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Configurações SAT"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "SATConfig"

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****

Public Function objParent() As Object

    Set objParent = Parent
    
End Function

Private Sub BotaoDirXml_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoDirXml_Click

    szTitle = "Localização física dos arquivos .xml"
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
       
        DirArqXml.Text = sBuffer
        Call DirArqXml_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoDirXml_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Private Sub DirArqXml_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_DirArqXml_Validate

    If Len(Trim(DirArqXml.Text)) = 0 Then Exit Sub
    
    If right(DirArqXml.Text, 1) <> "\" And right(DirArqXml.Text, 1) <> "/" Then
        iPos = InStr(1, DirArqXml.Text, "/")
        If iPos = 0 Then
            DirArqXml.Text = DirArqXml.Text & "\"
        Else
            DirArqXml.Text = DirArqXml.Text & "/"
        End If
    End If

    If Len(Trim(Dir(DirArqXml.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_DirArqXml_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_ErroECF(vbOKOnly, ERRO_DIRETORIO_INVALIDO, gErr, DirArqXml.Text)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 192328)

    End Select

    Exit Sub

End Sub

Private Sub BotaoNomeArqDll_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoNomeArqDll_Click
    
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "dll Files (*.dll)|*.dll"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    NomeArqDll.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoNomeArqDll_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoLogo_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoLogo_Click
    
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "bmp Files (*.bmp)|*.bmp"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    NomeArqLogo.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoLogo_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub BotaoTesteImpressora_Click()

Dim lErro As Long, bRelAberto As Boolean
Dim sMensagem As String
Dim objTela As Object

On Error GoTo Erro_BotaoTesteImpressora_Click

    Set objTela = Me
    
    lErro = AFRAC_AbrirRelatorioGerencial(RELGER_TESTE_IMPRESSAO, objTela)
    bRelAberto = True
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abrir Teste Impressao")
    If lErro <> SUCESSO Then gError 133088
    
    sMensagem = "123456789012345678901234567890123456789012345678"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
    
    sMensagem = "Teste de impressão"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
            
    'guilhotina
    lErro = AFRAC_AcionarGuilhotina("P")
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acionar Guilhotina")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sMensagem = "123456789012345678901234567890123456789012345678"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
    
    sMensagem = "Teste de impressão"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
            
    'guilhotina
    lErro = AFRAC_AcionarGuilhotina("P")
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Acionar Guilhotina")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sMensagem = "123456789012345678901234567890123456789012345678"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
    
    sMensagem = "Teste de impressão"
    lErro = AFRAC_ImprimirRelatorioGerencial(sMensagem, objTela)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Teste Impressao")
    If lErro <> SUCESSO Then gError 99893
            
    lErro = AFRAC_FecharRelatorioGerencial(objTela)
    bRelAberto = False
    
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Fechar Teste Impressao")
    'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub
    
Erro_BotaoTesteImpressora_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 99893
            Call AFRAC_FecharRelatorioGerencial(objTela)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 149439)

    End Select

    If bRelAberto Then Call AFRAC_FecharRelatorioGerencial(objTela)

    Exit Sub

End Sub

