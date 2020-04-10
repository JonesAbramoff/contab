VERSION 5.00
Begin VB.UserControl RelOpImportXMLOcx 
   Appearance      =   0  'Flat
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   6570
   Begin VB.CheckBox CadProd 
      Caption         =   "Cadastrar produtos não encontrados"
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
      Left            =   180
      TabIndex        =   12
      Top             =   1620
      Width           =   6000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localização da pasta com os XMLs dessa filial empresa"
      Height          =   720
      Left            =   180
      TabIndex        =   10
      Top             =   840
      Width           =   6195
      Begin VB.TextBox NomeDiretorio 
         Height          =   315
         Left            =   975
         TabIndex        =   1
         Top             =   270
         Width           =   4590
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
         Left            =   5565
         TabIndex        =   2
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpImportXML.ctx":0000
      Left            =   1005
      List            =   "RelOpImportXML.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2265
      Picture         =   "RelOpImportXML.ctx":0004
      TabIndex        =   3
      Top             =   1965
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5235
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpImportXML.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpImportXML.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   510
         Picture         =   "RelOpImportXML.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   -180
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   30
         Picture         =   "RelOpImportXML.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   -165
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   315
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "RelOpImportXMLOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const SW_HIDE = 0
Private Const STARTF_USESHOWWINDOW = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

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

Private Type STARTUPINFO
     cb As Long
     lpReserved As String
     lpDesktop As String
     lpTitle As String
     dwX As Long
     dwY As Long
     dwXSize As Long
     dwYSize As Long
     dwXCountChars As Long
     dwYCountChars As Long
     dwFillAttribute As Long
     dwFlags As Long
     wShowWindow As Integer
     cbReserved2 As Integer
     lpReserved2 As Long
     hStdInput As Long
     hStdOutput As Long
     hStdError As Long
  End Type

  Private Type PROCESS_INFORMATION
     hProcess As Long
     hThread As Long
     dwProcessID As Long
     dwThreadID As Long
  End Type
   
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call Trata_Default
       
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209839)

    End Select

    Exit Sub

End Sub

Private Function Trata_Default() As Long

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Trata_Default

    lErro = CF("Config_Le", "CRFATConfig", CRFATCFG_DIR_IMPORT_XML_NFE, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    NomeDiretorio.Text = sConteudo
    Call NomeDiretorio_Validate(bSGECancelDummy)
        
    Trata_Default = SUCESSO

    Exit Function

Erro_Trata_Default:

    Trata_Default = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209840)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError ERRO_SEM_MENSAGEM
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209841)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 209842
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If ComboOpcoes.ListCount <> 0 Then
        ComboOpcoes.ListIndex = 0
    End If
  
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
                
        Case 209842
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209843)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 209844
    bCancel = False
    Call NomeDiretorio_Validate(bCancel)
    If bCancel Then gError ERRO_SEM_MENSAGEM
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
    
        Case 209844
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209845)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    If ComboOpcoes.Visible Then
        ComboOpcoes.Text = ""
        ComboOpcoes.SetFocus
    End If
    
    Call Trata_Default
    
    If ComboOpcoes.ListCount <> 0 Then
        ComboOpcoes.ListIndex = 0
    End If
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209846)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sDiretorio As String
Dim lRetorno As Long
Dim objCRFATConfig As New ClassCRFATConfig
Dim sConteudo As String
Dim objRelatorio As New AdmRelatorio
Dim sNomeExeImport As String

On Error GoTo Erro_PreencherRelOp

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    objCRFATConfig.iFilialEmpresa = EMPRESA_TODA
    objCRFATConfig.iTipo = 0
    objCRFATConfig.sCodigo = CRFATCFG_DIR_IMPORT_XML_NFE
    objCRFATConfig.sDescricao = DESC_CRFATCFG_DIR_IMPORT_XML_NFE
    objCRFATConfig.sConteudo = NomeDiretorio.Text
    
    lErro = CF("CRFATConfig_Grava2", objCRFATConfig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sDiretorio = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
    sDiretorio = left(sDiretorio, lRetorno)
    
    'Obtem o nome customizado para máquina
    sNomeExeImport = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "NomePgmImportXml", "", sNomeExeImport, 255, NOME_ARQUIVO_ADM)
    sNomeExeImport = left(sNomeExeImport, lRetorno)
    
    'Se não tem customizado para máquina, busca customizado para o cliente (no DIC)
    If Len(Trim(sNomeExeImport)) = 0 Then
    
        lErro = CF("Controle_Le", CONTROLE_NOME_PGM_IMPORTXML, sNomeExeImport)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If Len(Trim(sNomeExeImport)) = 0 Then sNomeExeImport = "Import_Xml"
    
    Call ExecCmd(sDiretorio & sNomeExeImport & " " & CStr(glEmpresa))
    
    lErro = CF("Importa_Xml_NFe_Cust")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If CadProd.Value = vbChecked Then Call Chama_Tela("ImportXMLCadProd")
    
    lErro = CF("Config_Le", "FATConfig", "IMPORTAR_PV_APP_PPL", giFilialEmpresa, sConteudo)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If lErro = SUCESSO And StrParaInt(sConteudo) = MARCADO Then
    
        objRelatorio.sOpcao = "?*?"
        objRelatorio.Rel_Menu_Executar ("Log de Integração")
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    GL_objMDIForm.MousePointer = vbDefault

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209847)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 209848

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        ComboOpcoes.Text = ""
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 209848
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209849)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Unload Me

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209850)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 209851

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 209851
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209852)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
          
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209853)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NOTAS_REC
    Set Form_Load_Ocx = Me
    Caption = "Importação de arquivos xmls de NFe, CTe e/ou DI"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpImportXml"
    
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

Public Sub Unload(objme As Object)
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
    
    End If

End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física dos arquivos .xmls"
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209854)

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

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 209855

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 209855, 76, 75, 52
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209856)

    End Select

    Exit Sub

End Sub

Public Function ExecCmd(cmdline$)

Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim RET&
    
    start.dwFlags = STARTF_USESHOWWINDOW
    start.wShowWindow = SW_HIDE
    
    start.cb = Len(start)
    
    RET& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

    RET& = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, RET&)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    ExecCmd = RET&
End Function
