VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVRelsEmExcel 
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3270.122
   ScaleMode       =   0  'User
   ScaleWidth      =   6503.613
   Begin VB.ComboBox Tipo 
      Height          =   315
      ItemData        =   "TRVRelsEmExcel.ctx":0000
      Left            =   1215
      List            =   "TRVRelsEmExcel.ctx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   3285
   End
   Begin VB.Frame Frame3 
      Caption         =   "Localização"
      Height          =   1230
      Left            =   180
      TabIndex        =   14
      Top             =   1770
      Width           =   6075
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   2550
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   315
         Width           =   4050
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
         Left            =   5130
         TabIndex        =   6
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   ".csv"
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
         Index           =   5
         Left            =   3630
         TabIndex        =   18
         Top             =   750
         Width           =   840
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   4
         Left            =   270
         TabIndex        =   17
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   16
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de emissão dos vouchers"
      Height          =   855
      Left            =   165
      TabIndex        =   11
      Top             =   765
      Width           =   6105
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   300
         Left            =   2175
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissaoDe 
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissaoAte 
         Height          =   300
         Left            =   3225
         TabIndex        =   3
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   300
         Left            =   4395
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   2820
         TabIndex        =   13
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   660
         TabIndex        =   12
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5115
      ScaleHeight     =   495
      ScaleWidth      =   1095
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   1155
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   75
         Picture         =   "TRVRelsEmExcel.ctx":0028
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gerar a planilha selecionada"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "TRVRelsEmExcel.ctx":046A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   2
      Left            =   330
      TabIndex        =   15
      Top             =   255
      Width           =   810
   End
End
Attribute VB_Name = "TRVRelsEmExcel"
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

Public Sub Padrao_Tela()

Dim iMes As Integer
Dim iAno As Integer
Dim dtData As Date
Dim dtData1 As Date

On Error GoTo Erro_Padrao_Tela

    iMes = Month(gdtDataAtual)
    iAno = Year(gdtDataAtual)

    dtData = CDate("01/" & iMes & "/" & iAno)
    dtData = DateAdd("m", -1, dtData)

    dtData1 = CDate("1/" & iMes & "/" & iAno)
    dtData1 = DateAdd("d", -1, dtData1)
    
    DataEmissaoDe.PromptInclude = False
    DataEmissaoDe.Text = Format(dtData, "dd/mm/yy")
    DataEmissaoDe.PromptInclude = True
    
    DataEmissaoAte.PromptInclude = False
    DataEmissaoAte.Text = Format(dtData1, "dd/mm/yy")
    DataEmissaoAte.PromptInclude = True
    
    Exit Sub

Erro_Padrao_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200716)

    End Select


    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
   
    Call Padrao_Tela

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200717)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()
   'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Relatórios em Excel"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVRelsEmExcel"

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

Private Sub DataEmissaoDe_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoDe)

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoDe.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 200718

    Exit Sub

Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 200718

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200719)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoAte_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoAte)

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 200720

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 200720

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200721)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Click()
    Select Case Codigo_Extrai(Tipo.Text)
        Case TRVRELSPARAEXCEL_TIPO_PLANILHA_META
            NomeArquivo.Text = TRVRELSPARAEXCEL_TIPO_PLANILHA_META_NOMEARQ
    End Select
End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 200722

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 200722

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200723)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 200724

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 200724

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200725)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 200726

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 200726

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200727)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 200728

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 200728

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200729)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRVRelsEmExcel

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200730)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_TRVRelsEmExcel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVRelsEmExcel

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
    
    Call Padrao_Tela
 
    Exit Sub

Erro_Limpa_Tela_TRVRelsEmExcel:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200731)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    'Call Tela_QueryUnload(Me, iAlterado, UnloadMode, Cancel, iTelaCorrenteAtiva)

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

Private Sub BotaoGerar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click

    If Codigo_Extrai(Tipo.Text) = 0 Then gError 200732

    If StrParaDate(DataEmissaoDe.Text) = DATA_NULA Then gError 200733
    If StrParaDate(DataEmissaoAte.Text) = DATA_NULA Then gError 200734
    
    If StrParaDate(DataEmissaoDe.Text) > StrParaDate(DataEmissaoAte.Text) Then gError 200735
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 200736
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 200737
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = CF("TRVRelsEmExcel_Gera", Codigo_Extrai(Tipo.Text), StrParaDate(DataEmissaoDe.Text), StrParaDate(DataEmissaoAte.Text), NomeDiretorio.Text, NomeArquivo.Text)
    If lErro <> SUCESSO Then gError 200738
    
    GL_objMDIForm.MousePointer = vbDefault
       
    Call Limpa_Tela_TRVRelsEmExcel
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")
    
    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 200732
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TIPO_NAO_PREENCHIDO", gErr)

        Case 200733
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INIC_NAO_PREENCHIDA", gErr)

        Case 200734
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)

        Case 200735
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAODE_MAIOR_DATAEMISSAOATE", gErr, DataEmissaoDe.Text, DataEmissaoAte.Text)
       
        Case 200736
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
       
        Case 200737
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
       
        Case 200738
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200739)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200740)

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

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 200741

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 200741, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200742)

    End Select

    Exit Sub

End Sub
