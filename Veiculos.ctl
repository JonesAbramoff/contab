VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl VeiculosOcx 
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   KeyPreview      =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   7320
   Begin VB.CommandButton BotaoEntregas 
      Caption         =   "Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   90
      TabIndex        =   11
      Top             =   4920
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      Caption         =   "Outros"
      Height          =   630
      Left            =   105
      TabIndex        =   30
      Top             =   4095
      Width           =   7110
      Begin MSMask.MaskEdBox CustoHora 
         Height          =   315
         Left            =   1365
         TabIndex        =   10
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Custo p/ hora:"
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
         Index           =   6
         Left            =   60
         TabIndex        =   31
         Top             =   255
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Disponibilidade"
      Height          =   615
      Left            =   90
      TabIndex        =   23
      Top             =   3390
      Width           =   7125
      Begin MSMask.MaskEdBox DispPadraoDe 
         Height          =   315
         Left            =   1365
         TabIndex        =   8
         Top             =   210
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DispPadraoAte 
         Height          =   315
         Left            =   4170
         TabIndex        =   9
         Top             =   210
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   " "
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3750
         TabIndex        =   25
         Top             =   255
         Width           =   345
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   990
         TabIndex        =   24
         Top             =   255
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Capacidade"
      Height          =   630
      Left            =   75
      TabIndex        =   18
      Top             =   2670
      Width           =   7140
      Begin MSMask.MaskEdBox CapacidadeKg 
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Top             =   180
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VolumeM3 
         Height          =   315
         Left            =   4170
         TabIndex        =   7
         Top             =   180
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "m3"
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
         Index           =   3
         Left            =   4995
         TabIndex        =   22
         Top             =   195
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "kg"
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
         Index           =   0
         Left            =   2205
         TabIndex        =   21
         Top             =   225
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Volume:"
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
         Index           =   2
         Left            =   3375
         TabIndex        =   20
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso:"
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
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   2025
      Left            =   60
      TabIndex        =   17
      Top             =   585
      Width           =   7155
      Begin VB.ComboBox PlacaUF 
         Height          =   315
         Left            =   5055
         TabIndex        =   32
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox Proprio 
         Caption         =   "Próprio"
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
         Left            =   4950
         TabIndex        =   4
         Top             =   1170
         Width           =   1965
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   3210
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2265
         Picture         =   "Veiculos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   240
         Width           =   300
      End
      Begin VB.TextBox Descricao 
         Height          =   315
         Left            =   1395
         MaxLength       =   100
         TabIndex        =   2
         Top             =   690
         Width           =   5500
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1395
         TabIndex        =   0
         Top             =   240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Placa 
         Height          =   315
         Left            =   1395
         TabIndex        =   5
         Top             =   1575
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "U.F. da Placa:"
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
         Index           =   53
         Left            =   3750
         TabIndex        =   33
         Top             =   1620
         Width           =   1245
      End
      Begin VB.Label LabelPlaca 
         Alignment       =   1  'Right Justify
         Caption         =   "Placa:"
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
         Left            =   600
         TabIndex        =   29
         Top             =   1605
         Width           =   765
      End
      Begin VB.Label LabelTipo 
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
         Height          =   315
         Left            =   375
         TabIndex        =   28
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label LabelDescricao 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   990
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5115
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "Veiculos.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "Veiculos.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Veiculos.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Veiculos.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "VeiculosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1


Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Veículos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Veiculos"

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
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205293)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPOS_VEICULOS, Tipo)
    If lErro <> SUCESSO Then gError 205153
    
    'Carrega os Estados
    lErro = Carrega_PlacaUF()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Proprio.Value = vbChecked

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205294)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objVeiculos As ClassVeiculos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objVeiculos Is Nothing) Then

        lErro = Traz_Veiculos_Tela(objVeiculos)
        If lErro <> SUCESSO Then gError 205295

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 205295

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205296)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objVeiculos As ClassVeiculos) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objVeiculos.lCodigo = StrParaLong(Codigo.Text)
    objVeiculos.sDescricao = Descricao.Text
    objVeiculos.lTipo = LCodigo_Extrai(Tipo.Text)
    If Proprio.Value = vbChecked Then
        objVeiculos.iProprio = MARCADO
    Else
        objVeiculos.iProprio = DESMARCADO
    End If
    objVeiculos.sPlaca = Placa.Text
    objVeiculos.sPlacaUF = PlacaUF.Text
    objVeiculos.dCapacidadeKg = StrParaDbl(CapacidadeKg.Text)
    objVeiculos.dVolumeM3 = StrParaDbl(VolumeM3.Text)
    objVeiculos.dCustoHora = StrParaDbl(CustoHora.Text)
    
    If Len(Trim(DispPadraoDe.ClipText)) > 0 Then
        objVeiculos.dDispPadraoDe = CDate(DispPadraoDe.Text)
    End If
    If Len(Trim(DispPadraoAte.ClipText)) > 0 Then
        objVeiculos.dDispPadraoAte = CDate(DispPadraoAte.Text)
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205297)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objVeiculos As New ClassVeiculos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Veiculos"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objVeiculos)
    If lErro <> SUCESSO Then gError 205298

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objVeiculos.lCodigo, 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 205298

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205299)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objVeiculos As New ClassVeiculos

On Error GoTo Erro_Tela_Preenche

    objVeiculos.lCodigo = colCampoValor.Item("Codigo").vValor

    If objVeiculos.lCodigo <> 0 Then

        lErro = Traz_Veiculos_Tela(objVeiculos)
        If lErro <> SUCESSO Then gError 205300

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 205300

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205301)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objVeiculos As New ClassVeiculos

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 205302
    '#####################

    'Preenche o objVeiculos
    lErro = Move_Tela_Memoria(objVeiculos)
    If lErro <> SUCESSO Then gError 205303

    lErro = Trata_Alteracao(objVeiculos, objVeiculos.lCodigo)
    If lErro <> SUCESSO Then gError 205304

    'Grava o/a Veiculos no Banco de Dados
    lErro = CF("Veiculos_Grava", objVeiculos)
    If lErro <> SUCESSO Then gError 205305

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205302
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205303, 205304, 205305

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205306)

    End Select

    Exit Function

End Function

Function Limpa_Tela_Veiculos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Veiculos

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Proprio.Value = vbChecked
    Tipo.ListIndex = -1
    PlacaUF.Text = ""

    iAlterado = 0

    Limpa_Tela_Veiculos = SUCESSO

    Exit Function

Erro_Limpa_Tela_Veiculos:

    Limpa_Tela_Veiculos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205307)

    End Select

    Exit Function

End Function

Function Traz_Veiculos_Tela(objVeiculos As ClassVeiculos) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Veiculos_Tela

    Call Limpa_Tela_Veiculos

    If objVeiculos.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objVeiculos.lCodigo)
        Codigo.PromptInclude = True
    End If

    'Lê o Veiculos que está sendo Passado
    lErro = CF("Veiculos_Le", objVeiculos)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205308

    If lErro = SUCESSO Then

        If objVeiculos.lCodigo <> 0 Then
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objVeiculos.lCodigo)
            Codigo.PromptInclude = True
        End If

        Descricao.Text = objVeiculos.sDescricao
        
        Call Combo_Seleciona_ItemData(Tipo, objVeiculos.lTipo)

        If objVeiculos.iProprio = MARCADO Then
            Proprio.Value = vbChecked
        Else
            Proprio.Value = vbUnchecked
        End If

        Placa.Text = objVeiculos.sPlaca
        PlacaUF.Text = objVeiculos.sPlacaUF

        If objVeiculos.dCapacidadeKg <> 0 Then
            CapacidadeKg.PromptInclude = False
            CapacidadeKg.Text = Formata_Estoque(objVeiculos.dCapacidadeKg)
            CapacidadeKg.PromptInclude = True
        End If

        If objVeiculos.dVolumeM3 <> 0 Then
            VolumeM3.PromptInclude = False
            VolumeM3.Text = Formata_Estoque(objVeiculos.dVolumeM3)
            VolumeM3.PromptInclude = True
        End If

        If objVeiculos.dCustoHora <> 0 Then
            CustoHora.PromptInclude = False
            CustoHora.Text = Format(objVeiculos.dCustoHora, "STANDARD")
            CustoHora.PromptInclude = True
        End If

        If objVeiculos.dDispPadraoDe <> 0 Then
            DispPadraoDe.PromptInclude = False
            DispPadraoDe.Text = Format(objVeiculos.dDispPadraoDe, DispPadraoDe.Format)
            DispPadraoDe.PromptInclude = True
        End If

        If objVeiculos.dDispPadraoAte <> 0 Then
            DispPadraoAte.PromptInclude = False
            DispPadraoAte.Text = Format(objVeiculos.dDispPadraoAte, DispPadraoAte.Format)
            DispPadraoAte.PromptInclude = True
        End If

    End If

    iAlterado = 0

    Traz_Veiculos_Tela = SUCESSO

    Exit Function

Erro_Traz_Veiculos_Tela:

    Traz_Veiculos_Tela = gErr

    Select Case gErr

        Case 205308

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205309)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 205310

    'Limpa Tela
    Call Limpa_Tela_Veiculos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 205310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205311)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205312)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 205313

    Call Limpa_Tela_Veiculos

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 205313

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205314)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objVeiculos As New ClassVeiculos
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 205315
    '#####################

    objVeiculos.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_VEICULOS", objVeiculos.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("Veiculos_Exclui", objVeiculos)
        If lErro <> SUCESSO Then gError 205316

        'Limpa Tela
        Call Limpa_Tela_Veiculos

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 205315
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 205316

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205317)

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
       If lErro <> SUCESSO Then gError 205318

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 205318

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205319)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205320)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

'    'Verifica se Tipo está preenchida
'    If Len(Trim(Tipo.Text)) <> 0 Then
'
'       'Critica a Tipo
'       lErro = Long_Critica(Tipo.Text)
'       If lErro <> SUCESSO Then gError 205321
'
'    End If

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True

    Select Case gErr

        Case 205321

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205322)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Placa_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Placa_Validate

    'Verifica se Placa está preenchida
    If Len(Trim(Placa.Text)) <> 0 Then

       '#######################################
       'CRITICA Placa
       '#######################################

    End If

    Exit Sub

Erro_Placa_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205325)

    End Select

    Exit Sub

End Sub

Private Sub Placa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CapacidadeKg_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CapacidadeKg_Validate

    'Verifica se CapacidadeKg está preenchida
    If Len(Trim(CapacidadeKg.Text)) <> 0 Then

       'Critica a CapacidadeKg
       lErro = Valor_Positivo_Critica(CapacidadeKg.Text)
       If lErro <> SUCESSO Then gError 205326

       CapacidadeKg.Text = Formata_Estoque(StrParaDbl(CapacidadeKg.Text))

    End If

    Exit Sub

Erro_CapacidadeKg_Validate:

    Cancel = True

    Select Case gErr

        Case 205326

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205327)

    End Select

    Exit Sub

End Sub

Private Sub CapacidadeKg_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CapacidadeKg, iAlterado)
    
End Sub

Private Sub CapacidadeKg_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VolumeM3_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeM3_Validate

    'Verifica se VolumeM3 está preenchida
    If Len(Trim(VolumeM3.Text)) <> 0 Then

       'Critica a VolumeM3
       lErro = Valor_Positivo_Critica(VolumeM3.Text)
       If lErro <> SUCESSO Then gError 205328
       
       VolumeM3.Text = Formata_Estoque(StrParaDbl(VolumeM3.Text))

    End If

    Exit Sub

Erro_VolumeM3_Validate:

    Cancel = True

    Select Case gErr

        Case 205328

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205329)

    End Select

    Exit Sub

End Sub

Private Sub VolumeM3_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VolumeM3, iAlterado)
    
End Sub

Private Sub VolumeM3_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CustoHora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CustoHora_Validate

    'Verifica se CustoHora está preenchida
    If Len(Trim(CustoHora.Text)) <> 0 Then

       'Critica a CustoHora
       lErro = Valor_Positivo_Critica(CustoHora.Text)
       If lErro <> SUCESSO Then gError 205330

       CustoHora.Text = Format(StrParaDbl(CustoHora.Text), "STANDARD")

    End If

    Exit Sub

Erro_CustoHora_Validate:

    Cancel = True

    Select Case gErr

        Case 205330

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205331)

    End Select

    Exit Sub

End Sub

Private Sub CustoHora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CustoHora, iAlterado)
    
End Sub

Private Sub CustoHora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DispPadraoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DispPadraoDe_Validate

    'Verifica se DispPadraoDe está preenchida
    If Len(Trim(DispPadraoDe.ClipText)) <> 0 Then

       'Critica a DispPadraoDe
       lErro = Hora_Critica(DispPadraoDe.Text)
       If lErro <> SUCESSO Then gError 205332

    End If

    Exit Sub

Erro_DispPadraoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 205332

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205333)

    End Select

    Exit Sub

End Sub

Private Sub DispPadraoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DispPadraoDe, iAlterado)
    
End Sub

Private Sub DispPadraoDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DispPadraoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DispPadraoAte_Validate

    'Verifica se DispPadraoAte está preenchida
    If Len(Trim(DispPadraoAte.ClipText)) <> 0 Then

       'Critica a DispPadraoAte
       lErro = Hora_Critica(DispPadraoAte.Text)
       If lErro <> SUCESSO Then gError 205334

    End If

    Exit Sub

Erro_DispPadraoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 205334

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205335)

    End Select

    Exit Sub

End Sub

Private Sub DispPadraoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DispPadraoAte, iAlterado)
    
End Sub

Private Sub DispPadraoAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVeiculos As ClassVeiculos

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objVeiculos = obj1

    'Mostra os dados do Veiculos na tela
    lErro = Traz_Veiculos_Tela(objVeiculos)
    If lErro <> SUCESSO Then gError 205336

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 205336


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205337)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objVeiculos As New ClassVeiculos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objVeiculos.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("VeiculosLista", colSelecao, objVeiculos, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205338)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_VEICULO", "Veiculos", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 205339

    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 205339
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 205340)
    
    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
    
    End If

End Sub

Private Sub BotaoEntregas_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sFiltro As String

On Error GoTo Erro_BotaoEntregas_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        colSelecao.Add StrParaLong(Codigo.Text)
        sFiltro = "Veiculo = ?"

    End If

    Call Chama_Tela("MapaDeEntregaLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub

Erro_BotaoEntregas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205338)

    End Select

    Exit Sub
    
End Sub

Public Sub PlacaUF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PlacaUF_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual_CI(PlacaUF)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 46527

    'Se não encontrar --> Erro
    If lErro = 58583 Then gError 46528

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 46527

        Case 46528
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157957)

    End Select

    Exit Sub

End Sub

Private Function Carrega_PlacaUF() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_PlacaUF

    Set colSiglasUF = gcolUFs
    
    'Adiciona na Combo PlacaUF
    For iIndice = 1 To colSiglasUF.Count
        PlacaUF.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_PlacaUF = SUCESSO

    Exit Function

Erro_Carrega_PlacaUF:

    Carrega_PlacaUF = gErr

    Select Case gErr

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158064)

    End Select

End Function

