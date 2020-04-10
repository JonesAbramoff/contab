VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl PRJRelsEmExcel 
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   ScaleHeight     =   4185.156
   ScaleMode       =   0  'User
   ScaleWidth      =   6503.613
   Begin VB.Frame Filtros 
      Caption         =   "Filtros"
      Height          =   915
      Left            =   165
      TabIndex        =   23
      Top             =   3180
      Width           =   6075
      Begin VB.Frame Frame5 
         Caption         =   "Projeto"
         Height          =   630
         Left            =   60
         TabIndex        =   24
         Top             =   195
         Width           =   5925
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   3315
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   195
            Width           =   2550
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   765
            TabIndex        =   26
            Top             =   210
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   60
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa:"
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
            Index           =   41
            Left            =   2730
            TabIndex        =   27
            Top             =   240
            Width           =   570
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modelo"
      Height          =   675
      Left            =   165
      TabIndex        =   19
      Top             =   1425
      Width           =   6105
      Begin VB.CommandButton BotaoModelo 
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
         Left            =   5430
         TabIndex        =   21
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox Modelo 
         Height          =   285
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   4545
      End
      Begin VB.Label Label1 
         Caption         =   "Modelo:"
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
         Index           =   6
         Left            =   165
         TabIndex        =   22
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      ItemData        =   "PRJRelsEmExcel.ctx":0000
      Left            =   1020
      List            =   "PRJRelsEmExcel.ctx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   3285
   End
   Begin VB.Frame Frame3 
      Caption         =   "Localização"
      Height          =   1035
      Left            =   165
      TabIndex        =   14
      Top             =   2130
      Width           =   6090
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   900
         TabIndex        =   7
         Top             =   630
         Width           =   2550
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   900
         TabIndex        =   5
         Top             =   255
         Width           =   4530
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
         Left            =   5430
         TabIndex        =   6
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   ".xls"
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
         Left            =   3495
         TabIndex        =   18
         Top             =   660
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
         Left            =   135
         TabIndex        =   17
         Top             =   630
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
         Left            =   60
         TabIndex        =   16
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   615
      Left            =   165
      TabIndex        =   11
      Top             =   765
      Width           =   6105
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   300
         Left            =   1995
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   195
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissaoDe 
         Height          =   300
         Left            =   870
         TabIndex        =   1
         Top             =   210
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
         Top             =   195
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
         Top             =   195
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
         Top             =   240
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
         Left            =   525
         TabIndex        =   12
         Top             =   240
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
         Picture         =   "PRJRelsEmExcel.ctx":003D
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gerar a planilha selecionada"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "PRJRelsEmExcel.ctx":047F
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   135
      TabIndex        =   15
      Top             =   255
      Width           =   810
   End
End
Attribute VB_Name = "PRJRelsEmExcel"
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

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo


Dim glNumIntPRJ As Long
Dim glNumIntPRJEtapa As Long
                   
Dim sProjetoAnt As String
Dim sEtapaAnt As String

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Padrao_Tela()

Dim iMes As Integer
Dim iAno As Integer
Dim dtData As Date
Dim dtData1 As Date
Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Padrao_Tela

    'Call TodasCtas_Click

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
    
    Tipo.ListIndex = 0
    Call Tipo_Click
    
    sConteudo = ""
    lErro = CF("CRFatConfig_Le", "PRJ_EXPORTA_EXCEL_DIR", 0, sConteudo)
    If lErro <> SUCESSO And lErro <> 61454 Then gError ERRO_SEM_MENSAGEM
    
    NomeDiretorio.Text = sConteudo
    Call NomeDiretorio_Validate(bSGECancelDummy)
        
    Exit Sub

Erro_Padrao_Tela:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200716)

    End Select


    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
   
    Call Padrao_Tela
    
    'Call PreencheComboContas
    
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

    Name = "PRJRelsEmExcel"

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
Dim lErro As Long, sConteudo As String
    Select Case Codigo_Extrai(Tipo.Text)
        Case 1
            NomeArquivo.Text = "MovFinancPrj_" & Format(Date, "yyyymmdd") & "_" & Format(Now, "hhmmss")
    End Select
    sConteudo = ""
    lErro = CF("CRFatConfig_Le", "PRJ_EXPORTA_EXCEL_MODELO" + CStr(Codigo_Extrai(Tipo.Text)), 0, sConteudo)
    If lErro = SUCESSO Then Modelo.Text = sConteudo
    
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

    Call Limpa_Tela_PRJRelsEmExcel

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200730)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_PRJRelsEmExcel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PRJRelsEmExcel

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
    
    glNumIntPRJ = 0
    glNumIntPRJEtapa = 0
                   
    sProjetoAnt = ""
    sEtapaAnt = ""
    
    Call Padrao_Tela

    Exit Sub

Erro_Limpa_Tela_PRJRelsEmExcel:

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
Dim objFluxo As New ClassFluxoPRJ

On Error GoTo Erro_BotaoGerar_Click

    If Codigo_Extrai(Tipo.Text) = 0 Then gError 200732

    If StrParaDate(DataEmissaoDe.Text) = DATA_NULA Then gError 200733
    If StrParaDate(DataEmissaoAte.Text) = DATA_NULA Then gError 200734
    
    If StrParaDate(DataEmissaoDe.Text) > StrParaDate(DataEmissaoAte.Text) Then gError 200735
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 200736
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 200737
    If Len(Trim(Modelo.Text)) = 0 Then gError 211607
    
    objFluxo.dtDataAte = StrParaDate(DataEmissaoAte.Text)
    objFluxo.dtDataDe = StrParaDate(DataEmissaoDe.Text)
    'objFluxo.iCodConta = Codigo_Extrai(ContaCorrente.Text)
    objFluxo.iFilialEmpresa = giFilialEmpresa
    objFluxo.lNumIntEtapa = glNumIntPRJEtapa
    objFluxo.lNumIntPRJ = glNumIntPRJ
    objFluxo.sDiretorio = NomeDiretorio.Text
    objFluxo.sModelo = Modelo.Text
    objFluxo.sNomeArquivo = NomeArquivo.Text
    objFluxo.iTipo = Codigo_Extrai(Tipo.Text)
    
    GL_objMDIForm.MousePointer = vbHourglass
       
    lErro = CF("PRJRelsEmExcel_Gera", objFluxo)
    If lErro <> SUCESSO Then gError 200738
    
    GL_objMDIForm.MousePointer = vbDefault
       
    Call Limpa_Tela_PRJRelsEmExcel
    
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
       
        Case 211607
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)
       
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
Dim iPos As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPos = InStr(1, NomeDiretorio.Text, "/")
        If iPos = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 200741

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True
    
    NomeDiretorio.SetFocus

    Select Case gErr

        Case 200741, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200742)

    End Select

    Exit Sub

End Sub

Private Sub BotaoModelo_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoModelo_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "*.xls"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    Modelo.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoModelo_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Sub Projeto_GotFocus()
    Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Sub Projeto_Validate(Cancel As Boolean)
    Call ProjetoTela_Validate(Cancel)
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call ProjetoTela_Validate(Cancel)
End Sub

Public Function ProjetoTela_Validate(Cancel As Boolean) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProjeto As New ClassProjetos
Dim vbResult As VbMsgBoxResult
Dim colItensPRJCR As New Collection
Dim objItemPRJCR As New ClassItensPRJCR
Dim objPRJCR As ClassPRJCR
Dim colPRJCR As New Collection
Dim bPossuiDocOriginal As Boolean
Dim objNF As New ClassNFiscal
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer

On Error GoTo Erro_ProjetoTela_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Or sEtapaAnt <> SCodigo_Extrai(Etapa.Text) Then

        If Len(Trim(Projeto.ClipText)) > 0 Then
                
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError 194310
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194311
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194312
            
            If sProjetoAnt <> Projeto.Text Then
                Call gobjTelaProjetoInfo.Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
            End If
            
            If Len(Trim(Etapa.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194313
            
            End If
                          
            glNumIntPRJ = objProjeto.lNumIntDoc
            glNumIntPRJEtapa = objEtapa.lNumIntDoc
            
        Else
        
            glNumIntPRJ = 0
            glNumIntPRJEtapa = 0
            
            Etapa.Clear
            
        End If
        
        sProjetoAnt = Projeto.Text
        sEtapaAnt = SCodigo_Extrai(Etapa.Text)
        
    End If
    
    ProjetoTela_Validate = SUCESSO
    
    Exit Function

Erro_ProjetoTela_Validate:

    ProjetoTela_Validate = gErr

    Cancel = True

    Select Case gErr
    
        Case 194310, 194311, 194313
        
        Case 194312
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194314)

    End Select

    Exit Function

End Function

'Function PreencheComboContas() As Long
'
'Dim lErro As Long
'Dim colCodigoNomeConta As New AdmColCodigoNome
'Dim objCodigoNomeConta As New AdmCodigoNome
'
'On Error GoTo Erro_PreencheComboContas
'
'    'Carrega a Coleção de Contas
'    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeConta)
'    If lErro <> SUCESSO Then Error 59745
'
'    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
'    For Each objCodigoNomeConta In colCodigoNomeConta
'
'        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
'        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo
'
'    Next
'
'    PreencheComboContas = SUCESSO
'
'    Exit Function
'
'Erro_PreencheComboContas:
'
'    PreencheComboContas = Err
'
'    Select Case Err
'
'        Case 59745
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168808)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub ContaCorrente_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim iCodigo As Integer
'
'On Error GoTo Erro_ContaCorrente_Validate
'
'    'Verifica se foi preenchida a ComboBox
'    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub
'
'    'Verifica se está preenchida com o item selecionado na ComboBox
'    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub
'
'    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 59768
'
'    'Não existe o ítem com a STRING na List da ComboBox
'    If lErro <> SUCESSO Then Error 59769
'
'    Exit Sub
'
'Erro_ContaCorrente_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 59768 'Tratado na rotina chamada
'
'        Case 59769
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", Err, ContaCorrente.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 168816)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TodasCtas_Click()
'
'Dim lErro As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_TodasCtas_Click
'
'    'Limpa e desabilita a ComboTipo
'    ContaCorrente.ListIndex = -1
'    ContaCorrente.Enabled = False
'    TodasCtas.Value = True
'
'    Exit Sub
'
'Erro_TodasCtas_Click:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168817)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ApenasCta_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_OptionUmTipo_Click
'
'    'Limpa Combo Tipo e Abilita
'    ContaCorrente.ListIndex = -1
'    ContaCorrente.Enabled = True
'    ContaCorrente.SetFocus
'
'    Exit Sub
'
'Erro_OptionUmTipo_Click:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168819)
'
'    End Select
'
'    Exit Sub
'
'End Sub
