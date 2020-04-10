VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl VistoriaPRJ 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ForeColor       =   &H00000080&
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame FramePRJ 
      Caption         =   "Projeto"
      Height          =   1935
      Left            =   90
      TabIndex        =   20
      Top             =   630
      Width           =   9315
      Begin VB.ComboBox Etapa 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1050
         Width           =   7845
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   285
         Left            =   1380
         TabIndex        =   22
         Top             =   255
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label PRJDtFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   34
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fim:"
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
         Left            =   7545
         TabIndex        =   33
         Top             =   1470
         Width           =   360
      End
      Begin VB.Label PRJDtIni 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6150
         TabIndex        =   32
         Top             =   1455
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
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
         Left            =   5565
         TabIndex        =   31
         Top             =   1485
         Width           =   555
      End
      Begin VB.Label PRJResp 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1380
         TabIndex        =   30
         Top             =   1455
         Width           =   3840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
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
         Left            =   135
         TabIndex        =   29
         Top             =   1485
         Width           =   1155
      End
      Begin VB.Label PRJDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1380
         TabIndex        =   28
         Top             =   645
         Width           =   7815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   390
         TabIndex        =   27
         Top             =   690
         Width           =   915
      End
      Begin VB.Label PRJCli 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5130
         TabIndex        =   26
         Top             =   255
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   4410
         TabIndex        =   25
         Top             =   285
         Width           =   660
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
         ForeColor       =   &H00000080&
         Height          =   180
         Left            =   645
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   285
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
         ForeColor       =   &H00000080&
         Height          =   180
         Index           =   62
         Left            =   750
         TabIndex        =   23
         Top             =   1110
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados"
      Height          =   2985
      Left            =   105
      TabIndex        =   5
      Top             =   3465
      Width           =   9315
      Begin VB.TextBox Laudo 
         Height          =   2250
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   16
         Top             =   630
         Width           =   7755
      End
      Begin VB.ComboBox Responsavel 
         Height          =   315
         Left            =   1380
         TabIndex        =   15
         Top             =   225
         Width           =   3210
      End
      Begin VB.CheckBox RespUsu 
         Caption         =   "Usuário do Sistema"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   14
         Top             =   270
         Width           =   2400
      End
      Begin VB.Label Label1 
         Caption         =   "Laudo:"
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
         Height          =   330
         Index           =   4
         Left            =   705
         TabIndex        =   18
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
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
         Index           =   11
         Left            =   135
         TabIndex        =   17
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   750
      Left            =   90
      TabIndex        =   4
      Top             =   2670
      Width           =   9315
      Begin VB.CommandButton BotaoLimpaCodigo 
         Height          =   300
         Left            =   2190
         Picture         =   "VistoriaPRJ.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Limpar o Número"
         Top             =   270
         Width           =   345
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   4605
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   3450
         TabIndex        =   7
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataValidade 
         Height          =   300
         Left            =   7290
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataValidade 
         Height          =   315
         Left            =   6135
         TabIndex        =   11
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1395
         TabIndex        =   35
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Validade:"
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
         Index           =   0
         Left            =   5250
         TabIndex        =   12
         Top             =   330
         Width           =   810
      End
      Begin VB.Label LabelCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   705
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   330
         Width           =   660
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   8
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6660
      ScaleHeight     =   495
      ScaleWidth      =   2640
      TabIndex        =   3
      Top             =   60
      Width           =   2700
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1575
         Picture         =   "VistoriaPRJ.ctx":0532
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoAnexos 
         Height          =   360
         Left            =   90
         Picture         =   "VistoriaPRJ.ctx":0A64
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Anexar Arquivos"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2055
         Picture         =   "VistoriaPRJ.ctx":0BFA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1095
         Picture         =   "VistoriaPRJ.ctx":0D78
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   585
         Picture         =   "VistoriaPRJ.ctx":0F02
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "VistoriaPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim glNumIntPRJ As Long
Dim glNumIntPRJEtapa As Long
                   
Dim sProjetoAnt As String
Dim sEtapaAnt As String

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo
Dim gobjAnexos As ClassAnexos
                   
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Public iAlterado As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Vistorias de uma etapa do Projeto"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "VistoriaPRJ"
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        End If
        
    End If

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
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_UnLoad

    Set objEventoCodigo = Nothing
    Set gobjTelaProjetoInfo = Nothing
    Set gobjAnexos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213631)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me

    Set gobjAnexos = New ClassAnexos

    Set objEventoCodigo = New AdmEvento
        
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    Call Data_Validate(bSGECancelDummy)
    
    lErro = Carrega_Usuarios(Responsavel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213632)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objVistPRJ As ClassPRJEtapaVistorias) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objVistPRJ Is Nothing) Then

        lErro = Traz_VistPRJ_Tela(objVistPRJ)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213633)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objVistPRJ As ClassPRJEtapaVistorias) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
   
    objVistPRJ.lNumIntPRJEtapa = glNumIntPRJEtapa
    objVistPRJ.lCodigo = StrParaLong(Codigo.Caption)
    objVistPRJ.dtData = StrParaDate(Data.Text)
    objVistPRJ.dtDataValidade = StrParaDate(DataValidade.Text)
    objVistPRJ.sResponsavel = Responsavel.Text
    objVistPRJ.sLaudo = Laudo.Text
    
    Set objVistPRJ.objAnexos = gobjAnexos

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213634)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "VistoriaPRJ"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntPRJEtapa", glNumIntPRJEtapa, 0, "NumIntPRJEtapa"
    colCampoValor.Add "Codigo", StrParaLong(Codigo.Caption), 0, "Codigo"
    
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213635)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objVistPRJ As New ClassPRJEtapaVistorias

On Error GoTo Erro_Tela_Preenche

    objVistPRJ.lCodigo = colCampoValor.Item("Codigo").vValor
    objVistPRJ.lNumIntPRJEtapa = colCampoValor.Item("NumIntPRJEtapa").vValor

    lErro = Traz_VistPRJ_Tela(objVistPRJ)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213636)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objVistPRJ As New ClassPRJEtapaVistorias

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If glNumIntPRJEtapa = 0 Then gError 213637
    If Len(Trim(Data.ClipText)) = 0 Then gError 213638

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objVistPRJ)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Critica_Dados(objVistPRJ)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Trata_Alteracao(objVistPRJ, objVistPRJ.lNumIntPRJEtapa, objVistPRJ.lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Grava a etapa no Banco de Dados
    lErro = CF("PRJEtapaVistorias_Grava", objVistPRJ)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
       
        Case 213637
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ETAPA_NAO_PREENCHIDO2", gErr)
            Etapa.SetFocus
            
        Case 213638
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213639)

    End Select

    Exit Function

End Function

Function Critica_Dados(ByVal objVistPRJ As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Critica_Dados

    GL_objMDIForm.MousePointer = vbDefault

    Critica_Dados = SUCESSO

    Exit Function

Erro_Critica_Dados:

    Critica_Dados = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213640)

    End Select

    Exit Function

End Function

Function Limpa_Tela_VistPRJ() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_VistPRJ

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Set gobjAnexos = New ClassAnexos
    
    glNumIntPRJ = 0
    glNumIntPRJEtapa = 0
            
    Etapa.Clear
        
    sProjetoAnt = ""
    sEtapaAnt = ""
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    Call Data_Validate(bSGECancelDummy)
    
    PRJCli.Caption = ""
    PRJDesc.Caption = ""
    PRJResp.Caption = ""
    PRJDtIni.Caption = ""
    PRJDtFim.Caption = ""
    Codigo.Caption = ""

    iAlterado = 0

    Limpa_Tela_VistPRJ = SUCESSO

    Exit Function

Erro_Limpa_Tela_VistPRJ:

    Limpa_Tela_VistPRJ = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213641)

    End Select

    Exit Function

End Function

Function Traz_VistPRJ_Tela(ByVal objVistPRJAux As ClassPRJEtapaVistorias) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas
Dim objVistPRJ As New ClassPRJEtapaVistorias
Dim bNova As Boolean

On Error GoTo Erro_Traz_VistPRJ_Tela

    Call Limpa_Tela_VistPRJ
    
    bNova = True
    
    objVistPRJ.lNumIntDoc = objVistPRJAux.lNumIntDoc
    objVistPRJ.lNumIntPRJEtapa = objVistPRJAux.lNumIntPRJEtapa
    objVistPRJ.lCodigo = objVistPRJAux.lCodigo

    'Lê a Etapa que está sendo Passada
    lErro = CF("PRJEtapaVistorias_Le", objVistPRJ)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then bNova = False
    
    objEtapa.lNumIntDoc = objVistPRJ.lNumIntPRJEtapa
    
    If objEtapa.lNumIntDoc <> 0 Then
    
        lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 213642
        
    End If
    
    glNumIntPRJEtapa = objEtapa.lNumIntDoc
    
    objProjeto.lNumIntDoc = objEtapa.lNumIntDocPRJ
    
    If objProjeto.lNumIntDoc <> 0 Then
    
        lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 213643
        
    End If
    
    glNumIntPRJ = objProjeto.lNumIntDoc
        
    lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sProjetoAnt = Projeto.Text
    sEtapaAnt = objEtapa.sCodigo
        
    Call gobjTelaProjetoInfo.Trata_Etapa(glNumIntPRJ, Etapa)
    
    Call CF("SCombo_Seleciona2", Etapa, sEtapaAnt)
    sEtapaAnt = ""
    Call ProjetoTela_Validate(bSGECancelDummy)
    
    If Not bNova Then
    
        Codigo.Caption = CStr(objVistPRJ.lCodigo)
    
        Data.PromptInclude = False
        Data.Text = Format(objVistPRJ.dtData, "dd/mm/yy")
        Data.PromptInclude = True
    
        DataValidade.PromptInclude = False
        DataValidade.Text = Format(objVistPRJ.dtDataValidade, "dd/mm/yy")
        DataValidade.PromptInclude = True
        
        Laudo.Text = objVistPRJ.sLaudo
        
        Responsavel.Text = objVistPRJ.sResponsavel
        Call Responsavel_Validate(bSGECancelDummy)
        
        Set gobjAnexos = objVistPRJ.objAnexos
        If gobjAnexos Is Nothing Then
            Set gobjAnexos = New ClassAnexos
            gobjAnexos.iTipoDoc = ANEXO_TIPO_VISTPRJ
            gobjAnexos.lNumIntDoc = objVistPRJ.lNumIntDoc
        End If
        lErro = CF("Anexos_Le", gobjAnexos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Traz_VistPRJ_Tela = SUCESSO

    Exit Function

Erro_Traz_VistPRJ_Tela:

    Traz_VistPRJ_Tela = gErr

    Select Case gErr

        Case 213642
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)
        
        Case 213643
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.lNumIntDoc)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213644)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_VistPRJ

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213645)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213646)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_VistPRJ

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213647)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objVistPRJ As New ClassPRJEtapaVistorias
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If glNumIntPRJEtapa = 0 Then gError 213648
    If Len(Trim(Codigo.Caption)) = 0 Then gError 213649
    
    lErro = Move_Tela_Memoria(objVistPRJ)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PRJETAPAVISTORIAS", objVistPRJ.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("PRJEtapaVistorias_Exclui", objVistPRJ)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Limpa Tela
        Call Limpa_Tela_VistPRJ

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
           
        Case 213648
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ETAPA_NAO_PREENCHIDO2", gErr)
            Etapa.SetFocus
        
        Case 213649
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213650)

    End Select

    Exit Sub

End Sub

Private Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelCodigo_Click()

Dim objVistPRJ As New ClassPRJEtapaVistorias
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    If glNumIntPRJEtapa = 0 Then gError 213651

    'Preenche objFornecedor com NomeReduzido da tela
    objVistPRJ.lCodigo = StrParaLong(Codigo.Caption)
    
    colSelecao.Add glNumIntPRJEtapa

    Call Chama_Tela("VistoriaPRJLista", colSelecao, objVistPRJ, objEventoCodigo, "NumIntPRJEtapa = ?")
    
    Exit Sub
    
Erro_LabelCodigo_Click:

    Select Case gErr
        
        Case 213651
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ETAPA_NAO_PREENCHIDO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213652)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData
        
        Call Data_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213653)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData
        
        Call Data_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213654)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Select Case gobjFAT.iPRJTipoPrazoValidVist
        
            Case 0 'Não automático
            
            Case 1 'Anos
                Call DateParaMasked(DataValidade, DateAdd("yyyy", gobjFAT.iPRJPrazoValidVist, StrParaDate(Data.Text)))
            
            Case 2 'meses
                Call DateParaMasked(DataValidade, DateAdd("m", gobjFAT.iPRJPrazoValidVist, StrParaDate(Data.Text)))
            
            Case 3 'semanas
                Call DateParaMasked(DataValidade, DateAdd("ww", gobjFAT.iPRJPrazoValidVist, StrParaDate(Data.Text)))
            
            Case 4 'dias
                Call DateParaMasked(DataValidade, DateAdd("d", gobjFAT.iPRJPrazoValidVist, StrParaDate(Data.Text)))
            
        End Select
    
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213655)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataValidade_DownClick()

Dim lErro As Long
Dim sDataValidade As String

On Error GoTo Erro_UpDownDataValidade_DownClick

    DataValidade.SetFocus

    If Len(DataValidade.ClipText) > 0 Then

        sDataValidade = DataValidade.Text

        lErro = Data_Diminui(sDataValidade)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataValidade.Text = sDataValidade
        
        Call DataValidade_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataValidade_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213656)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataValidade_UpClick()

Dim lErro As Long
Dim sDataValidade As String

On Error GoTo Erro_UpDownDataValidade_UpClick

    DataValidade.SetFocus

    If Len(Trim(DataValidade.ClipText)) > 0 Then

        sDataValidade = DataValidade.Text

        lErro = Data_Aumenta(sDataValidade)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataValidade.Text = sDataValidade
        
        Call DataValidade_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataValidade_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213657)

    End Select

    Exit Sub

End Sub

Private Sub DataValidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataValidade, iAlterado)
    
End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidade_Validate

    If Len(Trim(DataValidade.ClipText)) <> 0 Then

        lErro = Data_Critica(DataValidade.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213658)

    End Select

    Exit Sub

End Sub

Private Sub DataValidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Laudo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVistPRJ As ClassPRJEtapaVistorias

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objVistPRJ = obj1

    'Mostra os dados do CentrodeTrabalho na tela
    lErro = Traz_VistPRJ_Tela(objVistPRJ)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213659)

    End Select

    Exit Sub

End Sub

'##################################################################
'Tem que colocar o código para o modo de edição aqui
Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub


Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelProjeto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProjeto, Source, X, Y)
End Sub

Private Sub LabelProjeto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProjeto, Button, Shift, X, Y)
End Sub
'##################################################################

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Private Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Projeto_GotFocus()
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
Dim bPossuiDocOriginal As Boolean
Dim objEtapa As New ClassPRJEtapas
Dim sProjeto As String
Dim iProjetoPreenchido As Integer
Dim objCliente As New ClassCliente, lCliente As Long

On Error GoTo Erro_ProjetoTela_Validate

    'Se alterou o projeto
    If sProjetoAnt <> Projeto.Text Or sEtapaAnt <> SCodigo_Extrai(Etapa.Text) Then
    
        Codigo.Caption = ""

        If Len(Trim(Projeto.ClipText)) > 0 Then
                
            lErro = Projeto_Formata(Projeto.Text, sProjeto, iProjetoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 213660
            
            If sProjetoAnt <> Projeto.Text Then
                Call gobjTelaProjetoInfo.Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
            End If
            
            If Len(Trim(Etapa.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            End If
                          
            glNumIntPRJ = objProjeto.lNumIntDoc
            glNumIntPRJEtapa = objEtapa.lNumIntDoc
            
            PRJDesc.Caption = objProjeto.sDescricao
            
            If objEtapa.lCliente <> 0 Then
                lCliente = objEtapa.lCliente
            Else
                lCliente = objProjeto.lCliente
            End If
            
            If lCliente <> 0 Then
            
                Set objCliente = New ClassCliente
                
                objCliente.lCodigo = lCliente
                
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
                
                PRJCli.Caption = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sRazaoSocial
           
            Else
            
                PRJCli.Caption = ""
            
            End If
            
            If objEtapa.sResponsavel <> "" Then
                PRJResp.Caption = objEtapa.sResponsavel
            Else
                PRJResp.Caption = objProjeto.sResponsavel
            End If
            
            If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.lNumIntDocPRJ <> 0 Then
                PRJDtIni.Caption = Format(objEtapa.dtDataInicio, "dd/mm/yyyy")
            ElseIf objProjeto.dtDataInicio <> DATA_NULA Then
                PRJDtIni.Caption = Format(objProjeto.dtDataInicio, "dd/mm/yyyy")
            Else
                PRJDtIni.Caption = ""
            End If
            
            If objEtapa.dtDataInicio <> DATA_NULA And objEtapa.lNumIntDocPRJ <> 0 Then
                PRJDtFim.Caption = Format(objEtapa.dtDataFim, "dd/mm/yyyy")
            ElseIf objProjeto.dtDataInicio <> DATA_NULA Then
                PRJDtFim.Caption = Format(objProjeto.dtDataFim, "dd/mm/yyyy")
            Else
                PRJDtFim.Caption = ""
            End If
            
        Else
        
            glNumIntPRJ = 0
            glNumIntPRJEtapa = 0
            
            Etapa.Clear
            
            PRJDesc.Caption = ""
            PRJCli.Caption = ""
            PRJResp.Caption = ""
            PRJDtIni.Caption = ""
            PRJDtFim.Caption = ""
            
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
    
        Case ERRO_SEM_MENSAGEM
        
        Case 213660
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 213661)

    End Select

    Exit Function

End Function

Public Sub Responsavel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Responsavel_Validate

    RespUsu.Value = vbUnchecked

    'Verifica se Responsavel está preenchida
    If Len(Trim(Responsavel.Text)) <> 0 Then

        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = Responsavel.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then RespUsu.Value = vbChecked

    End If

    Exit Sub

Erro_Responsavel_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213662)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpaCodigo_Click()
    Codigo.Caption = ""
End Sub

Private Function Carrega_Usuarios(ByVal objCombo As Object) As Long

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objUsuarios In colUsuarios
        objCombo.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213663)

    End Select

    Exit Function

End Function

Public Sub BotaoAnexos_Click()
   Call Chama_Tela_Modal("Anexos", gobjAnexos)
End Sub
