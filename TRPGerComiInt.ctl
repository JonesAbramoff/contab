VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPGerComiInt 
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3885
   ScaleMode       =   0  'User
   ScaleWidth      =   6503.613
   Begin VB.CommandButton BotaoComissao 
      Caption         =   "Gerações já Realizadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   135
      TabIndex        =   5
      ToolTipText     =   "Consulta as gerações já realizadas"
      Top             =   3150
      Width           =   1470
   End
   Begin VB.Frame Frame3 
      Caption         =   "Localização dos arquivos"
      Height          =   1080
      Left            =   135
      TabIndex        =   19
      Top             =   945
      Width           =   5970
      Begin VB.TextBox NomeDiretorio 
         Height          =   345
         Left            =   1080
         TabIndex        =   3
         Top             =   420
         Width           =   4200
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
         Left            =   5295
         TabIndex        =   4
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label2 
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
         Height          =   405
         Left            =   195
         TabIndex        =   20
         Top             =   465
         Width           =   1140
      End
   End
   Begin VB.CheckBox Previa 
      Caption         =   "Prévia"
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
      Left            =   3165
      TabIndex        =   2
      Top             =   390
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Geração"
      Height          =   825
      Left            =   135
      TabIndex        =   12
      Top             =   2190
      Width           =   6015
      Begin VB.Label Hora 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5040
         TabIndex        =   18
         Top             =   315
         Width           =   885
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3180
         TabIndex        =   17
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Usuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1080
         TabIndex        =   16
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
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
         Index           =   6
         Left            =   4530
         TabIndex        =   15
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Index           =   5
         Left            =   2610
         TabIndex        =   14
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Index           =   2
         Left            =   270
         TabIndex        =   13
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período de emissão dos vouchers"
      Height          =   690
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   2865
      Begin MSMask.MaskEdBox DataEmissaoAte 
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   270
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
         Left            =   2250
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
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
         Left            =   630
         TabIndex        =   11
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4425
      ScaleHeight     =   495
      ScaleWidth      =   1665
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   180
      Width           =   1725
      Begin VB.CommandButton BotaoGerar 
         Height          =   360
         Left            =   105
         Picture         =   "TRPGerComiInt.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gerar a comissão interna"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TRPGerComiInt.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "TRPGerComiInt.ctx":05CC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "TRPGerComiInt"
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

Private WithEvents objEventoGerComissao As AdmEvento
Attribute objEventoGerComissao.VB_VarHelpID = -1

Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Public Sub Padrao_Tela()

On Error GoTo Erro_Padrao_Tela
    
    DataEmissaoAte.PromptInclude = False
    DataEmissaoAte.Text = Format(DateAdd("d", -1, Date), "dd/mm/yy")
    DataEmissaoAte.PromptInclude = True
    
    Previa.Value = vbUnchecked
    
    Exit Sub

Erro_Padrao_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoGerComissao = New AdmEvento
    
    Call Padrao_Tela

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197260)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoGerComissao = Nothing
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()
   Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Geração de Comissão Interna"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPGerComiInt"

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

Private Sub DataEmissaoAte_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissaoAte_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 197263

    Exit Sub

Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197264)

    End Select

    Exit Sub

End Sub

Private Sub NomeDiretorio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 197277

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 197277

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197278)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 197279

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 197279

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197280)

    End Select

    Exit Sub

End Sub

Private Sub BotaoComissao_Click()

Dim lErro As Long
Dim objTRPGerComiInt As New ClassTRPGerComiInt
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoComissao_Click

    Call Chama_Tela("TRPGerComiIntLista", colSelecao, objTRPGerComiInt, objEventoGerComissao)

    Exit Sub

Erro_BotaoComissao_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197289)

    End Select

    Exit Sub

End Sub

Private Sub objEventoGerComissao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRPGerComiInt As ClassTRPGerComiInt

On Error GoTo Erro_objEventoGerComissao_evSelecao

    Set objTRPGerComiInt = obj1

    'Mostra os dados do TRPGerComiInt na tela
    lErro = Traz_TRPGerComiInt_Tela(objTRPGerComiInt)
    If lErro <> SUCESSO Then gError 197290

    Me.Show

    Exit Sub

Erro_objEventoGerComissao_evSelecao:

    Select Case gErr

        Case 197290

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197291)

    End Select

    Exit Sub

End Sub

Function Traz_TRPGerComiInt_Tela(objTRPGerComiInt As ClassTRPGerComiInt) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Traz_TRPGerComiInt_Tela

    Call Limpa_Tela_TRPGerComiInt

    'Lê o TRPGerComiInt que está sendo Passado
    lErro = CF("TRPGerComiInt_Le", objTRPGerComiInt)
    If lErro <> SUCESSO And lErro <> 197295 Then gError 197297
    
    If lErro = SUCESSO Then
        
        If objTRPGerComiInt.dtDataEmiAte <> DATA_NULA Then
            DataEmissaoAte.PromptInclude = False
            DataEmissaoAte.Text = Format(objTRPGerComiInt.dtDataEmiAte, "dd/mm/yy")
            DataEmissaoAte.PromptInclude = True
        End If


        Usuario.Caption = objTRPGerComiInt.sUsuario
        Data.Caption = Format(objTRPGerComiInt.dtDataGeracao, "dd/mm/yyyy")
        Hora.Caption = Format(objTRPGerComiInt.dHoraGeracao, "hh:mm:ss")
        NomeDiretorio.Text = objTRPGerComiInt.sDiretorio
        If objTRPGerComiInt.iPrevia = MARCADO Then
            Previa.Value = vbChecked
        Else
            Previa.Value = vbUnchecked
        End If
        
    End If

    iAlterado = 0

    Traz_TRPGerComiInt_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPGerComiInt_Tela:

    Traz_TRPGerComiInt_Tela = gErr

    Select Case gErr

        Case 197297

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197298)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_TRPGerComiInt

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 197299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197300)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_TRPGerComiInt()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRPGerComiInt

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela(Me)
          
    Usuario.Caption = ""
    Data.Caption = ""
    Hora.Caption = ""
    
    Call Padrao_Tela
    
    iAlterado = 0
 
    Exit Sub

Erro_Limpa_Tela_TRPGerComiInt:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197301)

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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
          
    End If

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

Private Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objTRPGerComiInt As New ClassTRPGerComiInt
    
On Error GoTo Erro_BotaoExcluir_Click
    
    If Len(Trim(Data.Caption)) = 0 Then gError 197419

    objTRPGerComiInt.dtDataGeracao = StrParaDate(Data.Caption)
    
    lErro = CF("TRPGerComiInt_Exclui", objTRPGerComiInt)
    If lErro <> SUCESSO Then gError 197420
    
    Call Limpa_Tela_TRPGerComiInt
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 197419
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case 197420
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197420)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim objTRPGerComiInt As New ClassTRPGerComiInt

On Error GoTo Erro_BotaoGerar_Click

    lErro = Formata_E_Critica_Dados(objTRPGerComiInt)
    If lErro <> SUCESSO Then gError 197304
    
    GL_objMDIForm.MousePointer = vbHourglass
          
    lErro = CF("TRPGerComiInt_Gera", objTRPGerComiInt)
    If lErro <> SUCESSO Then gError 197306
        
    GL_objMDIForm.MousePointer = vbDefault
    
    If objTRPGerComiInt.iPrevia = 0 Then Call Limpa_Tela_TRPGerComiInt
    
    Exit Sub

Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 197304 To 197306
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197307)

    End Select

    Exit Sub

End Sub

Public Function Formata_E_Critica_Dados(objTRPGerComiInt As ClassTRPGerComiInt) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Dados

    objTRPGerComiInt.dtDataGeracao = Date
    
    If StrParaDate(DataEmissaoAte.Text) = DATA_NULA Then gError 200816
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 200817

    'Lê o TRPGerComiInt que está sendo Passado
    lErro = CF("TRPGerComiInt_Le", objTRPGerComiInt)
    If lErro <> SUCESSO And lErro <> 197895 Then gError 200818

    If lErro = SUCESSO Then gError 200819

    objTRPGerComiInt.sUsuario = gsUsuario
    objTRPGerComiInt.dHoraGeracao = Time
    objTRPGerComiInt.dtDataEmiAte = StrParaDate(DataEmissaoAte.Text)
    objTRPGerComiInt.sDiretorio = NomeDiretorio.Text
    
    If Previa.Value = vbChecked Then
        objTRPGerComiInt.iPrevia = MARCADO
    Else
        objTRPGerComiInt.iPrevia = DESMARCADO
    End If

    Formata_E_Critica_Dados = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Dados:

    Formata_E_Critica_Dados = gErr

    Select Case gErr
        
        Case 200816
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_NAO_PREENCHIDA", gErr)
            DataEmissaoAte.SetFocus
            
        Case 200817
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
        
        Case 200818
        
        Case 200819
            Call Rotina_Erro(vbOKOnly, "ERRO_GERCOMIINT_DATA_EXISTENTE", gErr, Date)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200820)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPGerComiInt"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "DataGeracao", StrParaDate(Data.Caption), 0, "DataGeracao"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197454)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTRPGerComiInt As New ClassTRPGerComiInt

On Error GoTo Erro_Tela_Preenche

    objTRPGerComiInt.dtDataGeracao = colCampoValor.Item("DataGeracao").vValor

    If objTRPGerComiInt.dtDataGeracao <> DATA_NULA Then
    
        lErro = Traz_TRPGerComiInt_Tela(objTRPGerComiInt)
        If lErro <> SUCESSO Then gError 197455
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 197455

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197456)

    End Select

    Exit Function

End Function

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
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
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
    
    If Right(NomeDiretorio.Text, 1) <> "\" And Right(NomeDiretorio.Text, 1) <> "/" Then
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
