VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl SpedFiscalOCx 
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   LockControls    =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   6795
   Begin VB.Frame Frame3 
      Caption         =   "Livro Registro de Controle da Produção e do Estoque (RCPE):"
      Height          =   510
      Left            =   150
      TabIndex        =   23
      Top             =   3810
      Width           =   6525
      Begin VB.CheckBox IncluiRCPE 
         Caption         =   "Empresa obrigada a escriturar o RCPE"
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
         Left            =   1005
         TabIndex        =   7
         ToolTipText     =   "O inventário deverá ser apresentado no arquivo da EFD-ICMS/IPI, no segundo mês subsequente ao evento."
         Top             =   195
         Width           =   3750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Registro de inventário"
      Height          =   2055
      Left            =   150
      TabIndex        =   19
      Top             =   1740
      Width           =   6525
      Begin VB.CheckBox FiltroNatureza 
         Caption         =   $"SpedFiscalOcx.ctx":0000
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
         Left            =   990
         TabIndex        =   6
         ToolTipText     =   "O inventário deverá ser apresentado no arquivo da EFD-ICMS/IPI, no segundo mês subsequente ao evento."
         Top             =   1320
         Value           =   1  'Checked
         Width           =   5460
      End
      Begin VB.ComboBox MotivoRegInv 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "SpedFiscalOcx.ctx":00A6
         Left            =   1005
         List            =   "SpedFiscalOcx.ctx":00B9
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   615
         Width           =   5445
      End
      Begin VB.CheckBox IncluiRegInv 
         Caption         =   "Incluir no Arquivo"
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
         Left            =   1005
         TabIndex        =   4
         ToolTipText     =   "O inventário deverá ser apresentado no arquivo da EFD-ICMS/IPI, no segundo mês subsequente ao evento."
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label DataInv 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1005
         TabIndex        =   22
         Top             =   1005
         Width           =   1185
      End
      Begin VB.Label LabelRegInv 
         Caption         =   "Refer.:"
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo:"
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
         Left            =   330
         TabIndex        =   20
         Top             =   675
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destino"
      Height          =   960
      Left            =   150
      TabIndex        =   15
      Top             =   750
      Width           =   6525
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
         Left            =   6060
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   165
         Width           =   360
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   990
         TabIndex        =   2
         Top             =   195
         Width           =   5070
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   990
         MaxLength       =   50
         TabIndex        =   3
         Top             =   555
         Width           =   3405
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
         Left            =   135
         TabIndex        =   18
         Top             =   240
         Width           =   795
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
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data Emissão"
      Height          =   675
      Left            =   150
      TabIndex        =   12
      Top             =   30
      Width           =   4650
      Begin VB.ComboBox Mes 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "SpedFiscalOcx.ctx":01C2
         Left            =   735
         List            =   "SpedFiscalOcx.ctx":01ED
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1860
      End
      Begin MSMask.MaskEdBox Ano 
         Height          =   315
         Left            =   3315
         TabIndex        =   1
         Top             =   255
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2790
         TabIndex        =   14
         Top             =   315
         Width           =   405
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   315
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   5010
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   1680
      Begin VB.CommandButton BotaoGerar 
         Height          =   345
         Left            =   105
         Picture         =   "SpedFiscalOcx.ctx":0256
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gera o arquivo"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   607
         Picture         =   "SpedFiscalOcx.ctx":0698
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1110
         Picture         =   "SpedFiscalOcx.ctx":0BCA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "SpedFiscalOCx"
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

Dim iMesAnt As Integer
Dim iAnoAnt As Integer

Private WithEvents objEventoBotaoInventario As AdmEvento
Attribute objEventoBotaoInventario.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub Ano_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Ano, iAlterado)

End Sub

Private Sub Ano_Validate(bCancel As Boolean)

On Error GoTo Erro_Ano_Validate

    If Len(Trim(Ano.Text)) > 0 Then

        If Ano.Text < 1900 Then gError 204110
        
    End If
    
    If iAnoAnt <> StrParaInt(Ano.Text) Then
        Call Trata_DataInv
        iAnoAnt = StrParaInt(Ano.Text)
    End If
    
    Exit Sub
    
Erro_Ano_Validate:

    Select Case gErr
    
        Case 204110
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_INVALIDO", gErr)
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204111)

    End Select
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim sDiretorio As String
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim sNomeArqParam As String
Dim iIncluiRegInv As Integer, iMotivoRegInv As Integer
Dim dtDataInv As Date, iFiltroNatureza As Integer, iIncluiRCPE As Integer

On Error GoTo Erro_BotaoGerar_Click
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 204112
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 204113
    If Len(Ano.Text) = 0 Then gError 204114
    If Len(Mes.Text) = 0 Then gError 204115
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text & NomeArquivo.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\" & NomeArquivo.Text
    End If
    
    dtDataIni = CDate("01/" & Mes.ItemData(Mes.ListIndex) & "/" & Ano.Text)
    
    dtDataFim = DateAdd("m", 1, dtDataIni)
    dtDataFim = DateAdd("d", -1, dtDataFim)
    
    iMotivoRegInv = Codigo_Extrai(MotivoRegInv.Text)
    If IncluiRegInv.Value = vbChecked Then
        iIncluiRegInv = MARCADO
    Else
        iIncluiRegInv = DESMARCADO
    End If
    
    If Len(Trim(DataInv.Caption)) <> 0 Then
        dtDataInv = StrParaDate(DataInv.Caption)
    Else
        dtDataInv = DATA_NULA
    End If
    
    If FiltroNatureza.Value = vbChecked Then
        iFiltroNatureza = MARCADO
    Else
        iFiltroNatureza = DESMARCADO
    End If
    
    If IncluiRCPE.Value = vbChecked Then
        iIncluiRCPE = MARCADO
    Else
        iIncluiRCPE = DESMARCADO
    End If
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 204116
    
    lErro = CF("Rotina_SpedFiscal", sNomeArqParam, giFilialEmpresa, sDiretorio, dtDataIni, dtDataFim, iIncluiRegInv, iMotivoRegInv, dtDataInv, iFiltroNatureza, iIncluiRCPE)
    If lErro <> SUCESSO Then gError 204117
        
    Call BotaoLimpar_Click
   

    Exit Sub
    
Erro_BotaoGerar_Click:

    Select Case gErr
    
        Case 204112
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_INFORMADO", gErr)
        
        Case 204113
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
        
        Case 204114
            Call Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREECHIDO", gErr)
        
        Case 204115
            Call Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREECHIDO", gErr)
        
        Case 204116, 204117
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204118)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)
    
    Mes.ListIndex = -1
    
    iAnoAnt = 0
    iMesAnt = 0
    
    DataInv.Caption = ""
    
    MotivoRegInv.ListIndex = 0

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204119)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoBotaoInventario = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    MotivoRegInv.ListIndex = 0
    
    Set objEventoBotaoInventario = New AdmEvento
    
    iMesAnt = -1
    
    Call Trata_RegInv
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204120)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204121)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204122)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Sped Fiscal"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "SpedFiscal"

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204123)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204124)

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

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 204125

'    Drive1.Drive = Mid(NomeDiretorio.Text, 1, 2)
'
'    Dir1.Path = NomeDiretorio.Text

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 76, 204125
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204126)

    End Select

    Exit Sub

End Sub

Private Sub Mes_Change()
    Call Trata_RegInv
End Sub

Private Sub Mes_Click()
    Call Trata_RegInv
End Sub

Private Sub Trata_RegInv()

Dim lErro As Long, iMes As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Trata_RegInv

    objFilialEmpresa.iCodFilial = giFilialEmpresa
    
    'Lê dados da FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
    If Mes.ListIndex <> -1 Then
        iMes = Mes.ItemData(Mes.ListIndex)
    Else
        iMes = 0
    End If
    
    If iMesAnt <> iMes Then
        iMesAnt = iMes
        
        If iMes = 2 Or objFilialEmpresa.sCNAE = "4681801" Or objFilialEmpresa.sCNAE = "4681802" Then
            IncluiRegInv.Value = vbChecked
        Else
            IncluiRegInv.Value = vbUnchecked
        End If
        
        Call Trata_DataInv
        
    End If

    Exit Sub

Erro_Trata_RegInv:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204126)

    End Select

    Exit Sub

End Sub

Private Sub IncluiRegInv_Click()
    If IncluiRegInv.Value = vbUnchecked Then
        MotivoRegInv.ListIndex = 0
        MotivoRegInv.Enabled = False
    Else
        MotivoRegInv.Enabled = True
    End If
End Sub

Private Sub Trata_DataInv()

Dim dtDataIni As Date, dtDataAux As Date

On Error GoTo Erro_Trata_DataInv

    If Len(Trim(Ano.Text)) <> 0 And Len(Trim(Mes.Text)) <> 0 Then
    
        dtDataIni = CDate("01/" & Mes.ItemData(Mes.ListIndex) & "/" & Ano.Text)
        dtDataAux = DateAdd("m", -1, dtDataIni)  'Retorna mais um mês
        dtDataAux = DateAdd("d", -Day(dtDataAux), dtDataAux) 'Retorna para o último dia do mês anterior

        DataInv.Caption = Format(dtDataAux, "dd/mm/yyyy")

    Else
    
        DataInv.Caption = ""

    End If

    Exit Sub

Erro_Trata_DataInv:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204126)

    End Select

    Exit Sub

End Sub

Private Sub LabelRegInv_Click()

Dim colSelecao As New Collection
Dim objRegInventario As ClassRegInventario

    Call Chama_Tela("RegInventarioLista", colSelecao, objRegInventario, objEventoBotaoInventario)

End Sub

Private Sub objEventoBotaoInventario_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRegInventario As ClassRegInventario

On Error GoTo Erro_objEventoBotaoInventario_evSelecao

    Set objRegInventario = obj1

    DataInv.Caption = Format(objRegInventario.dtData, "dd/mm/yyyy")

    Me.Show

    Exit Sub

Erro_objEventoBotaoInventario_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159281)

    End Select

    Exit Sub

End Sub
