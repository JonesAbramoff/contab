VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRPConfig 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   KeyPreview      =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   5355
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
      Left            =   4560
      TabIndex        =   5
      Top             =   2805
      Width           =   555
   End
   Begin VB.TextBox ModeloFatCartao 
      Height          =   285
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2835
      Width           =   4050
   End
   Begin VB.TextBox ModeloFat 
      Height          =   285
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1770
      Width           =   4050
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1740
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
      Left            =   4560
      TabIndex        =   7
      Top             =   3585
      Width           =   555
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   375
      TabIndex        =   6
      Top             =   3615
      Width           =   4050
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4035
      ScaleHeight     =   495
      ScaleWidth      =   1035
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   135
      Width           =   1095
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   555
         Picture         =   "TRPConfig.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRPConfig.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox ProxTitRec 
      Height          =   315
      Left            =   3840
      TabIndex        =   0
      Top             =   900
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox ProxTitPag 
      Height          =   315
      Left            =   1635
      TabIndex        =   1
      Top             =   360
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
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Modelo para geração de Faturas de Cartão e Nota de Crédito em html:"
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
      Height          =   390
      Index           =   3
      Left            =   390
      TabIndex        =   15
      Top             =   2340
      Width           =   4500
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
      Height          =   300
      Index           =   1
      Left            =   390
      TabIndex        =   14
      Top             =   1455
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
      Height          =   495
      Index           =   2
      Left            =   390
      TabIndex        =   13
      Top             =   3255
      Width           =   4965
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
      TabIndex        =   12
      Top             =   105
      Visible         =   0   'False
      Width           =   3420
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
      Left            =   360
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   945
      Width           =   3420
   End
End
Attribute VB_Name = "TRPConfig"
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

'Variáveis Globais
Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Form_Load
    
    lErro = CF("TRPConfig_Le", TRPCONFIG_PROX_NUM_TITREC, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192315
    
    ProxTitRec.PromptInclude = False
    ProxTitRec.Text = sConteudo
    ProxTitRec.PromptInclude = True

    lErro = CF("TRPConfig_Le", TRPCONFIG_PROX_NUM_TITPAG, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192316

    ProxTitPag.PromptInclude = False
    ProxTitPag.Text = sConteudo
    ProxTitPag.PromptInclude = True

    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_MODELO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192317
    
    ModeloFat.Text = sConteudo

    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192317
    
    ModeloFatCartao.Text = sConteudo

    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192318
    
    NomeDiretorio.Text = sConteudo

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
    
'
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

    Name = "TRPConfig"

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

    lErro = CF("TRPConfig_Grava", colConfig)
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
Dim objTRPConfig As ClassTRPConfig

On Error GoTo Erro_Move_Tela_Memoria

    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_PROX_NUM_TITREC
    objTRPConfig.sConteudo = ProxTitRec.Text
    
    colConfig.Add objTRPConfig

    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_PROX_NUM_TITPAG
    objTRPConfig.sConteudo = ProxTitPag.Text
    
    colConfig.Add objTRPConfig

    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_DIRETORIO_MODELO_FAT_HTML
    objTRPConfig.sConteudo = ModeloFat.Text
    
    colConfig.Add objTRPConfig
    
    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_DIRETORIO_FAT_HTML
    objTRPConfig.sConteudo = NomeDiretorio.Text
    
    colConfig.Add objTRPConfig
    
    Set objTRPConfig = New ClassTRPConfig
    
    objTRPConfig.sCodigo = TRPCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO
    objTRPConfig.sConteudo = ModeloFatCartao.Text
    
    colConfig.Add objTRPConfig

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
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
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
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
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

