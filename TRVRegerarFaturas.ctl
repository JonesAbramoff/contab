VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRVRegerarFaturas 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   5025
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      Height          =   1515
      Index           =   0
      Left            =   255
      TabIndex        =   18
      Top             =   2640
      Width           =   4530
      Begin VB.ComboBox TipoDocSeleciona 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "TRVRegerarFaturas.ctx":0000
         Left            =   1470
         List            =   "TRVRegerarFaturas.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   900
         Width           =   2955
      End
      Begin VB.OptionButton TipoDocTodos 
         Caption         =   "Todos"
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
         Left            =   405
         TabIndex        =   6
         Top             =   420
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton TipoDocApenas 
         Caption         =   "Apenas:"
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
         Left            =   405
         TabIndex        =   7
         Top             =   930
         Width           =   1050
      End
   End
   Begin VB.TextBox ModeloFat 
      Height          =   285
      Left            =   255
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4710
      Visible         =   0   'False
      Width           =   3975
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
      Left            =   4245
      TabIndex        =   5
      Top             =   4635
      Visible         =   0   'False
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
      Left            =   4245
      TabIndex        =   3
      Top             =   2055
      Width           =   555
   End
   Begin VB.TextBox NomeDiretorio 
      Height          =   285
      Left            =   255
      TabIndex        =   2
      Top             =   2115
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3210
      ScaleHeight     =   495
      ScaleWidth      =   1530
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   75
      Width           =   1590
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   90
         Picture         =   "TRVRegerarFaturas.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gerar o documento .html"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1050
         Picture         =   "TRVRegerarFaturas.ctx":0446
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "TRVRegerarFaturas.ctx":05C4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Faturas"
      Height          =   1005
      Index           =   1
      Left            =   285
      TabIndex        =   12
      Top             =   690
      Width           =   4530
      Begin MSMask.MaskEdBox FaturaDe 
         Height          =   300
         Left            =   720
         TabIndex        =   0
         Top             =   390
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FaturaAte 
         Height          =   300
         Left            =   2985
         TabIndex        =   1
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelInicio 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   330
         TabIndex        =   14
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LabelFim 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2550
         TabIndex        =   13
         Top             =   435
         Width           =   360
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Modelo da Fatura em html:"
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
      Height          =   300
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   4365
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Localização física dos arquivos html:"
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
      Height          =   300
      Index           =   2
      Left            =   255
      TabIndex        =   16
      Top             =   1785
      Width           =   3225
   End
End
Attribute VB_Name = "TRVRegerarFaturas"
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

'Variáveis globais
Dim iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call Default_Tela
    
    Call Carrega_TipoDocumento(TipoDocSeleciona)
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192846)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long
'A tela não espera recebimento de parâmetros, portanto, essa função sempre retorna sucesso
    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    'Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Libera os objetos e coleções globais

End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub FaturaDe_GotFocus()
    Call MaskEdBox_TrataGotFocus(FaturaDe, iAlterado)
End Sub

Private Sub FaturaAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(FaturaAte, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa a tela
    Call Limpa_Tela_RegerarFaturas

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192847)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub


'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub FaturaDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub FaturaAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***

'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

Private Sub Limpa_Tela_RegerarFaturas()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_RegerarFaturas

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    
    Call Default_Tela
    
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_RegerarFaturas:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192848)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Regera as faturas em .html"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVRegerarFaturas"

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
'**** fim do trecho a ser copiado *****

'*** TRATAMENTO PARA MODO DE EDIÇÃO - INÍCIO ***
Private Sub LabelInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelInicio, Button, Shift, X, Y)
End Sub

Private Sub LabelInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelInicio, Source, X, Y)
End Sub

Private Sub LabelFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFim, Button, Shift, X, Y)
End Sub

Private Sub LabelFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFim, Source, X, Y)
End Sub

'*** TRATAMENTO PARA MODO DE EDIÇÃO - FIM ***
Sub BotaoGerar_Click()

Dim lErro As Long
Dim sSiglaDoc As String

On Error GoTo Erro_BotaoGerar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If StrParaLong(FaturaDe.Text) = 0 Then gError 192849
    If StrParaLong(FaturaAte.Text) = 0 Then gError 192850
    If StrParaLong(FaturaDe.Text) > StrParaLong(FaturaAte.Text) Then gError 192851
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 192852
    If Len(Trim(ModeloFat.Text)) = 0 Then gError 192853
    
    
    If TipoDocApenas.Value = True Then
        sSiglaDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
    Else
        sSiglaDoc = ""
    End If
    
    lErro = CF("TRVFaturas_Regera_Html", StrParaLong(FaturaDe.Text), StrParaLong(FaturaAte.Text), ModeloFat.Text, NomeDiretorio.Text, sSiglaDoc)
    If lErro <> SUCESSO Then gError 192854
    
    GL_objMDIForm.MousePointer = vbDefault

    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")

    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 192849
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERODE_NAO_PREENCHIDO", gErr)
            FaturaDe.SetFocus

        Case 192850
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMEROATE_NAO_PREENCHIDO", gErr)
            FaturaAte.SetFocus
    
        Case 192851
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MENOR_NUMERO_DE", gErr)
            FaturaDe.SetFocus
    
        Case 192852
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
        
        Case 192853
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)
            ModeloFat.SetFocus
            
        Case 192854
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192855)

    End Select

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192856)

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

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 192857

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 192857, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192858)

    End Select

    Exit Sub

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

Sub Default_Tela()

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Default_Tela
   
    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192859
    
    NomeDiretorio.Text = sConteudo
    Call NomeDiretorio_Validate(bSGECancelDummy)
    
    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_MODELO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192860
    
    ModeloFat.Text = sConteudo
    
    TipoDocTodos.Value = True
    
    Call TipoDocTodos_Click

    Exit Sub

Erro_Default_Tela:

    Select Case gErr
    
        Case 192859, 192860

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192861)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 192303
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 192303

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192304)

    End Select

    Exit Function

End Function

Public Sub TipoDocApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True

End Sub

Public Sub TipoDocTodos_Click()

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

End Sub
