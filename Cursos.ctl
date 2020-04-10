VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Cursos 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5190
      Index           =   1
      Left            =   75
      TabIndex        =   14
      Top             =   630
      Width           =   9270
      Begin VB.Frame Frame2 
         Caption         =   "Certificados"
         Height          =   1815
         Index           =   2
         Left            =   225
         TabIndex        =   25
         Top             =   3255
         Width           =   8805
         Begin VB.CommandButton BotaoMarcarTodos 
            Height          =   480
            Left            =   7920
            Picture         =   "Cursos.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   225
            Width           =   780
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Height          =   480
            Left            =   7920
            Picture         =   "Cursos.ctx":101A
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   720
            Width           =   780
         End
         Begin VB.ListBox Certificados 
            Columns         =   3
            Height          =   1410
            ItemData        =   "Cursos.ctx":21FC
            Left            =   225
            List            =   "Cursos.ctx":21FE
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   255
            Width           =   7665
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Básicos"
         Height          =   2265
         Index           =   0
         Left            =   225
         TabIndex        =   20
         Top             =   90
         Width           =   8805
         Begin VB.TextBox Detalhamento 
            Height          =   1035
            Left            =   1625
            MaxLength       =   250
            TabIndex        =   1
            Top             =   690
            Width           =   7095
         End
         Begin VB.TextBox Responsavel 
            Height          =   315
            Left            =   1625
            MaxLength       =   250
            TabIndex        =   2
            Top             =   1860
            Width           =   7095
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2535
            Picture         =   "Cursos.ctx":2200
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Numeração Automática"
            Top             =   225
            Width           =   300
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1620
            TabIndex        =   0
            Top             =   225
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigo 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label LabelDetalhamento 
            Alignment       =   1  'Right Justify
            Caption         =   "Detalhamento:"
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
            Left            =   135
            TabIndex        =   23
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label LabelResponsavel 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   135
            TabIndex        =   22
            Top             =   1890
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Duração"
         Height          =   690
         Index           =   1
         Left            =   225
         TabIndex        =   15
         Top             =   2475
         Width           =   8805
         Begin MSMask.MaskEdBox DataInicio 
            Height          =   315
            Left            =   1590
            TabIndex        =   3
            Top             =   225
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataInicio 
            Height          =   300
            Left            =   2910
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataConclusao 
            Height          =   315
            Left            =   4905
            TabIndex        =   4
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataConclusao 
            Height          =   300
            Left            =   6225
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelDataInicio 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   930
            TabIndex        =   19
            Top             =   255
            Width           =   600
         End
         Begin VB.Label LabelDataConclusao 
            Alignment       =   1  'Right Justify
            Caption         =   "Conclusão:"
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
            Left            =   3300
            TabIndex        =   18
            Top             =   270
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5190
      Index           =   2
      Left            =   75
      TabIndex        =   28
      Top             =   630
      Visible         =   0   'False
      Width           =   9270
      Begin VB.Frame Frame6 
         Caption         =   "Participantes"
         Height          =   5055
         Left            =   105
         TabIndex        =   29
         Top             =   135
         Width           =   9105
         Begin VB.CheckBox MOAprovado 
            DragMode        =   1  'Automatic
            Height          =   270
            Left            =   5415
            TabIndex        =   33
            Top             =   3435
            Width           =   1185
         End
         Begin VB.TextBox MOAvaliacao 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3420
            TabIndex        =   32
            Top             =   3885
            Width           =   3480
         End
         Begin VB.TextBox MODesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1335
            TabIndex        =   31
            Top             =   2880
            Width           =   2610
         End
         Begin VB.TextBox MOCodigo 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   120
            MaxLength       =   20
            TabIndex        =   30
            Top             =   2910
            Width           =   705
         End
         Begin VB.CommandButton BotaoMO 
            Caption         =   "Participantes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Abre o Browse de Máquinas, Habilidades e Processos"
            Top             =   4530
            Width           =   2100
         End
         Begin MSFlexGridLib.MSFlexGrid GridMO 
            Height          =   2160
            Left            =   105
            TabIndex        =   10
            Top             =   225
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   3810
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7305
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Cursos.ctx":22EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Cursos.ctx":2444
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Cursos.ctx":25CE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Cursos.ctx":2B00
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5595
      Left            =   45
      TabIndex        =   13
      Top             =   285
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Curso/Exame"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Participantes"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Cursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1

Dim objGridMO As AdmGrid
Dim iGrid_MOCod_Col As Integer
Dim iGrid_MODesc_Col As Integer
Dim iGrid_MOAprovado_Col As Integer
Dim iGrid_MOAvaliacao_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cursos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Cursos"

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

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoMO = Nothing
    Set objGridMO = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213255)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoMO = New AdmEvento
    Set objGridMO = New AdmGrid
    
    iFrameAtual = 1
    
    lErro = Inicializa_GridMO(objGridMO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê o Código e a Descrição de cada Tipo de Mão-de-Obra
    lErro = CF("Cod_Nomes_Le", "Certificados", "Codigo", "Sigla", STRING_MAXIMO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 137558

    'preenche a ListBox certificados com os objetos da colecao
    For Each objCodigoDescricao In colCodigoDescricao
        Certificados.AddItem objCodigoDescricao.sNome
        Certificados.ItemData(Certificados.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213256)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCursos As ClassCursos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCursos Is Nothing) Then

        lErro = Traz_Cursos_Tela(objCursos)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213257)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCursos As ClassCursos) As Long

Dim lErro As Long, iIndice As Integer
Dim objMO As ClassCursoMO
Dim objCertificado As ClassCursoCertificados

On Error GoTo Erro_Move_Tela_Memoria

    objCursos.lCodigo = StrParaLong(Codigo.Text)
    objCursos.iFilialEmpresa = giFilialEmpresa
    objCursos.sDetalhamento = Detalhamento.Text
    objCursos.sResponsavel = Responsavel.Text
    objCursos.dtDataInicio = StrParaDate(DataInicio.Text)
    objCursos.dtDataConclusao = StrParaDate(DataConclusao.Text)
    
    Set objCursos.colCertificados = New Collection
    For iIndice = 0 To Certificados.ListCount - 1
        If Certificados.Selected(iIndice) Then
            Set objCertificado = New ClassCursoCertificados
            objCertificado.lCodCertificado = Certificados.ItemData(iIndice)
            objCursos.colCertificados.Add objCertificado
        End If
    Next
    
    For iIndice = 1 To objGridMO.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridMO.TextMatrix(iIndice, iGrid_MOCod_Col))) <> 0 Then
        
            Set objMO = New ClassCursoMO
        
            objMO.sAvaliacao = GridMO.TextMatrix(iIndice, iGrid_MOAvaliacao_Col)
            objMO.iAprovado = StrParaInt(GridMO.TextMatrix(iIndice, iGrid_MOAprovado_Col))
            objMO.iCodMO = StrParaInt(GridMO.TextMatrix(iIndice, iGrid_MOCod_Col))
        
            objCursos.colMOCursos.Add objMO
        
        End If
        
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213258)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCursos As New ClassCursos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Cursos"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objCursos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCursos.lCodigo, 0, "Codigo"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213259)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCursos As New ClassCursos

On Error GoTo Erro_Tela_Preenche

    objCursos.lCodigo = colCampoValor.Item("Codigo").vValor

    objCursos.iFilialEmpresa = giFilialEmpresa

    If objCursos.lCodigo <> 0 And objCursos.iFilialEmpresa <> 0 Then

        lErro = Traz_Cursos_Tela(objCursos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213260)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCursos As New ClassCursos

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 213261
    '#####################

    'Preenche o objCursos
    lErro = Move_Tela_Memoria(objCursos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Trata_Alteracao(objCursos, objCursos.lCodigo, objCursos.iFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Grava o/a Cursos no Banco de Dados
    lErro = CF("Cursos_Grava", objCursos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 213261
            Call Rotina_Erro(vbOKOnly, "ERRO_CURSOS_CODIGO_NAO_PREENCHIDO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213262)

    End Select

    Exit Function

End Function

Private Sub BotaoDesmarcarTodos_Click()
Dim iLinha As Integer
    For iLinha = 0 To Certificados.ListCount - 1
        Certificados.Selected(iLinha) = False
    Next
End Sub

Private Sub BotaoMarcarTodos_Click()
Dim iLinha As Integer
    For iLinha = 0 To Certificados.ListCount - 1
        Certificados.Selected(iLinha) = True
    Next
End Sub

Function Limpa_Tela_Cursos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Cursos

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Call Grid_Limpa(objGridMO)
    
    Call BotaoDesmarcarTodos_Click

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_Cursos = SUCESSO

    Exit Function

Erro_Limpa_Tela_Cursos:

    Limpa_Tela_Cursos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213263)

    End Select

    Exit Function

End Function

Function Traz_Cursos_Tela(objCursos As ClassCursos) As Long

Dim lErro As Long, iIndice As Integer
Dim objMO As ClassCursoMO
Dim objTipoMO As ClassTiposDeMaodeObra
Dim objCertificado As ClassCursoCertificados

On Error GoTo Erro_Traz_Cursos_Tela

    Call Limpa_Tela_Cursos

    'Lê o Cursos que está sendo Passado
    lErro = CF("Cursos_Le", objCursos)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then

        If objCursos.lCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objCursos.lCodigo)
            Codigo.PromptInclude = True
        End If

        Detalhamento.Text = objCursos.sDetalhamento
        Responsavel.Text = objCursos.sResponsavel

        If objCursos.dtDataInicio <> DATA_NULA Then
            DataInicio.PromptInclude = False
            DataInicio.Text = Format(objCursos.dtDataInicio, "dd/mm/yy")
            DataInicio.PromptInclude = True
        End If

        If objCursos.dtDataConclusao <> DATA_NULA Then
            DataConclusao.PromptInclude = False
            DataConclusao.Text = Format(objCursos.dtDataConclusao, "dd/mm/yy")
            DataConclusao.PromptInclude = True
        End If
        
        For Each objCertificado In objCursos.colCertificados
            For iIndice = 0 To Certificados.ListCount - 1
                If objCertificado.lCodCertificado = Certificados.ItemData(iIndice) Then
                    Certificados.Selected(iIndice) = True
                    Exit For
                End If
            Next
        Next
        
        iIndice = 0
        For Each objMO In objCursos.colMOCursos
            iIndice = iIndice + 1
            Set objTipoMO = New ClassTiposDeMaodeObra
            objTipoMO.iCodigo = objMO.iCodMO
            'Lê o TiposDeMaodeObra que está sendo Passado
            lErro = CF("TiposDeMaodeObra_Le", objTipoMO)
            If lErro <> SUCESSO And lErro <> 137598 Then gError ERRO_SEM_MENSAGEM
        
            GridMO.TextMatrix(iIndice, iGrid_MOAprovado_Col) = CStr(objMO.iAprovado)
            GridMO.TextMatrix(iIndice, iGrid_MOAvaliacao_Col) = objMO.sAvaliacao
            GridMO.TextMatrix(iIndice, iGrid_MOCod_Col) = CStr(objTipoMO.iCodigo)
            GridMO.TextMatrix(iIndice, iGrid_MODesc_Col) = objTipoMO.sDescricao
        Next
        objGridMO.iLinhasExistentes = objCursos.colMOCursos.Count
        Call Grid_Refresh_Checkbox(objGridMO)

    End If

    iAlterado = 0

    Traz_Cursos_Tela = SUCESSO

    Exit Function

Erro_Traz_Cursos_Tela:

    Traz_Cursos_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213264)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_Cursos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213265)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213266)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_Cursos

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213267)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCursos As New ClassCursos
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 213268
    '#####################

    objCursos.lCodigo = StrParaLong(Codigo.Text)
    objCursos.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CURSOS", objCursos.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Cursos_Exclui", objCursos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Limpa Tela
        Call Limpa_Tela_Cursos

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 213268
            Call Rotina_Erro(vbOKOnly, "ERRO_CURSOS_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213269)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213270)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Detalhamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Detalhamento_Validate

    'Verifica se Detalhamento está preenchida
    If Len(Trim(Detalhamento.Text)) <> 0 Then

       '#######################################
       'CRITICA Detalhamento
       '#######################################

    End If

    Exit Sub

Erro_Detalhamento_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213271)

    End Select

    Exit Sub

End Sub

Private Sub Detalhamento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Responsavel_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Responsavel_Validate

    'Verifica se Responsavel está preenchida
    If Len(Trim(Responsavel.Text)) <> 0 Then

       '#######################################
       'CRITICA Responsavel
       '#######################################

    End If

    Exit Sub

Erro_Responsavel_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213272)

    End Select

    Exit Sub

End Sub

Private Sub Responsavel_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_DownClick

    DataInicio.SetFocus

    If Len(DataInicio.ClipText) > 0 Then

        sData = DataInicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        DataInicio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213273)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_UpClick

    DataInicio.SetFocus

    If Len(Trim(DataInicio.ClipText)) > 0 Then

        sData = DataInicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataInicio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213274)

    End Select

    Exit Sub

End Sub

Private Sub DataInicio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicio, iAlterado)
    
End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate

    If Len(Trim(DataInicio.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicio.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213275)

    End Select

    Exit Sub

End Sub

Private Sub DataInicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataConclusao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataConclusao_DownClick

    DataConclusao.SetFocus

    If Len(DataConclusao.ClipText) > 0 Then

        sData = DataConclusao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataConclusao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataConclusao_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213276)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataConclusao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataConclusao_UpClick

    DataConclusao.SetFocus

    If Len(Trim(DataConclusao.ClipText)) > 0 Then

        sData = DataConclusao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataConclusao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataConclusao_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213277)

    End Select

    Exit Sub

End Sub

Private Sub DataConclusao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataConclusao, iAlterado)
    
End Sub

Private Sub DataConclusao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataConclusao_Validate

    If Len(Trim(DataConclusao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataConclusao.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataConclusao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213278)

    End Select

    Exit Sub

End Sub

Private Sub DataConclusao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCursos As ClassCursos

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCursos = obj1

    'Mostra os dados do Cursos na tela
    lErro = Traz_Cursos_Tela(objCursos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213279)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objCursos As New ClassCursos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objCursos.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("CursosLista", colSelecao, objCursos, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213280)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_PROX_CURSOS", "Cursos", "Codigo", lCodigo)
    If lErro <> SUCESSO And lErro <> 25191 Then gError ERRO_SEM_MENSAGEM
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213281)
    
    End Select

    Exit Sub

End Sub

Private Sub GridMO_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Private Sub GridMO_EnterCell()
    Call Grid_Entrada_Celula(objGridMO, iAlterado)
End Sub

Private Sub GridMO_GotFocus()
    Call Grid_Recebe_Foco(objGridMO)
End Sub

Private Sub GridMO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMO)
End Sub

Private Sub GridMO_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Private Sub GridMO_LeaveCell()
    Call Saida_Celula(objGridMO)
End Sub

Private Sub GridMO_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridMO)
End Sub

Private Sub GridMO_RowColChange()
    Call Grid_RowColChange(objGridMO)
End Sub

Private Sub GridMO_Scroll()
    Call Grid_Scroll(objGridMO)
End Sub

Public Sub MOCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MOCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Public Sub MOCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Public Sub MOCodigo_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMO.objControle = MOCodigo
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub MOAvaliacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MOAvaliacao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Public Sub MOAvaliacao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Public Sub MOAvaliacao_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMO.objControle = MOAvaliacao
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub MOAprovado_Click()

Dim lErro As Long

On Error GoTo Erro_MOAprovado_Click
    
    'Verifica se é alguma linha válida
    If GridMO.Row > objGridMO.iLinhasExistentes Then Exit Sub

    lErro = Grid_Refresh_Checkbox(objGridMO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_MOAprovado_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213282)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridMO.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MOCod_Col
                
                    lErro = Saida_Celula_MOCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Case iGrid_MOAvaliacao_Col
    
                    lErro = Saida_Celula_Padrao(objGridInt, MOAvaliacao)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End Select

        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 213283

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 213283
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213284)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim lCodMO As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Guardo o valor do Codigo do Tipo de MO
    lCodMO = GridMO.TextMatrix(GridMO.Row, iGrid_MOCod_Col)
    
    Select Case objControl.Name
    
        Case MOCodigo.Name
            objControl.Enabled = True
            
        Case MODesc.Name
            objControl.Enabled = False
            
        Case MOAvaliacao.Name
            If lCodMO <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case MOAprovado.Name
            If lCodMO <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 213285)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridMO(objGrid As AdmGrid) As Long

Dim iIndice As Integer

On Error GoTo Erro_Inicializa_GridMO

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Código")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Aprovado")
    objGrid.colColuna.Add ("Avaliação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (MOCodigo.Name)
    objGrid.colCampo.Add (MODesc.Name)
    objGrid.colCampo.Add (MOAprovado.Name)
    objGrid.colCampo.Add (MOAvaliacao.Name)

    'Colunas do Grid
    iGrid_MOCod_Col = 1
    iGrid_MODesc_Col = 2
    iGrid_MOAprovado_Col = 3
    iGrid_MOAvaliacao_Col = 4

    objGrid.objGrid = GridMO

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 11

    'Largura da primeira coluna
    GridMO.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMO = SUCESSO

    Exit Function

Erro_Inicializa_GridMO:

    Inicializa_GridMO = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 213286)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOCodigo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodTipoMO As Integer
Dim iLinha As Integer
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_MOCodigo

    Set objGridInt.objControle = MOCodigo
                    
    'Se o campo foi preenchido
    If Len(MOCodigo.Text) > 0 Then

        'Verifica se há algum produto repetido no grid
        For iLinha = 1 To objGridInt.iLinhasExistentes
            
            If iLinha <> GridMO.Row Then
                                                    
                If StrParaInt(GridMO.TextMatrix(iLinha, iGrid_MOCod_Col)) = StrParaInt(MOCodigo.Text) Then
                    iCodTipoMO = StrParaInt(MOCodigo.Text)
                    MOCodigo.Text = ""
                    gError 213287
                End If
                    
            End If
                           
        Next
        
        Set objTiposDeMaodeObra = New ClassTiposDeMaodeObra
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(MOCodigo.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError ERRO_SEM_MENSAGEM
    
        If lErro <> SUCESSO Then gError 213288

        GridMO.TextMatrix(GridMO.Row, iGrid_MODesc_Col) = objTiposDeMaodeObra.sDescricao
        GridMO.TextMatrix(GridMO.Row, iGrid_MOAprovado_Col) = CStr(MARCADO)
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
        Call Grid_Refresh_Checkbox(objGridInt)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134132

    Saida_Celula_MOCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCodigo:

    Saida_Celula_MOCodigo = gErr

    Select Case gErr

        Case 213287
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, CStr(iCodTipoMO), iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 213288
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOCodigo.SetFocus
        
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 213289)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Sub BotaoMO_Click()

Dim lErro As Long
Dim objMO As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is MOCodigo Then
    
        objMO.iCodigo = StrParaInt(MOCodigo.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridMO.Row = 0 Then gError 213353

        objMO.iCodigo = StrParaInt(GridMO.TextMatrix(GridMO.Row, iGrid_MOCod_Col))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objMO, objEventoMO)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 213353
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213354)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMO As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoMO_evSelecao

    Set objMO = obj1

    'Verifica se há algum produto repetido no grid
    For iLinha = 1 To objGridMO.iLinhasExistentes
        If iLinha < GridMO.Row Then
            If StrParaInt(GridMO.TextMatrix(iLinha, iGrid_MOCod_Col)) = objMO.iCodigo Then
                MOCodigo.Text = ""
                gError 213355
            End If
        End If
    Next
    
    MOCodigo.Text = CStr(objMO.iCodigo)
    
    If Not (Me.ActiveControl Is MOCodigo) Then
    
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCod_Col) = CStr(objMO.iCodigo)
        GridMO.TextMatrix(GridMO.Row, iGrid_MODesc_Col) = objMO.sDescricao
        GridMO.TextMatrix(GridMO.Row, iGrid_MOAprovado_Col) = CStr(MARCADO)
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridMO.iLinhasExistentes Then
            objGridMO.iLinhasExistentes = objGridMO.iLinhasExistentes + 1
        End If
        
        Call Grid_Refresh_Checkbox(objGridMO)
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr

        Case 213355
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_REPETIDO", gErr, objMO.iCodigo, iLinha)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213356)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    
        If Me.ActiveControl Is MOCodigo Then Call BotaoMO_Click
    
    ElseIf KeyCode = KEYCODE_PROXIMO_NUMERO Then
        
        Call BotaoProxNum_Click
        
    End If
    
End Sub
