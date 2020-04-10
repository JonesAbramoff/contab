VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ApontamentoPRJ 
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ForeColor       =   &H00000080&
   KeyPreview      =   -1  'True
   ScaleHeight     =   6495
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame FrameMP 
      Caption         =   "Produtos"
      Height          =   1845
      Index           =   3
      Left            =   60
      TabIndex        =   40
      Top             =   2715
      Width           =   9405
      Begin MSMask.MaskEdBox MPUM 
         Height          =   270
         Left            =   3660
         TabIndex        =   49
         Top             =   495
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   15
         Format          =   "#,##0.0#"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton BotaoMP 
         Caption         =   "Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   9
         Top             =   1455
         Width           =   1600
      End
      Begin VB.TextBox MPDescricao 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1635
         MaxLength       =   50
         TabIndex        =   43
         Top             =   480
         Width           =   1950
      End
      Begin VB.TextBox MPOBS 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   6765
         MaxLength       =   50
         TabIndex        =   41
         Top             =   540
         Width           =   1890
      End
      Begin MSMask.MaskEdBox MPCustoT 
         Height          =   270
         Left            =   5880
         TabIndex        =   42
         Top             =   540
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MPProduto 
         Height          =   270
         Left            =   150
         TabIndex        =   44
         Top             =   465
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MPQuantidade 
         Height          =   270
         Left            =   4305
         TabIndex        =   45
         Top             =   495
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   15
         Format          =   "#,##0.0#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MPCusto 
         Height          =   270
         Left            =   4995
         TabIndex        =   46
         Top             =   510
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridMP 
         Height          =   270
         Left            =   30
         TabIndex        =   8
         Top             =   190
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   476
         _Version        =   393216
         Rows            =   10
      End
      Begin VB.Label Label1 
         Caption         =   "Custo Total:"
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
         Index           =   24
         Left            =   6390
         TabIndex        =   48
         Top             =   1530
         Width           =   1050
      End
      Begin VB.Label MPCustoTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7455
         TabIndex        =   47
         Top             =   1470
         Width           =   1410
      End
   End
   Begin VB.Frame FrameMO 
      Caption         =   "Mão de Obra"
      Height          =   1845
      Index           =   3
      Left            =   60
      TabIndex        =   30
      Top             =   840
      Width           =   9405
      Begin VB.CommandButton BotaoMO 
         Caption         =   "Mão de Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   7
         Top             =   1455
         Width           =   1600
      End
      Begin VB.TextBox MOCodigo 
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   240
         MaxLength       =   20
         TabIndex        =   34
         Top             =   600
         Width           =   645
      End
      Begin VB.TextBox MODescricao 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   900
         MaxLength       =   50
         TabIndex        =   33
         Top             =   600
         Width           =   2235
      End
      Begin VB.TextBox MOOBS 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   6765
         MaxLength       =   50
         TabIndex        =   31
         Top             =   615
         Width           =   2010
      End
      Begin MSMask.MaskEdBox MOCustoT 
         Height          =   270
         Left            =   3525
         TabIndex        =   32
         Top             =   735
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MOQuantidade 
         Height          =   270
         Left            =   5295
         TabIndex        =   35
         Top             =   1005
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MOHoras 
         Height          =   270
         Left            =   6285
         TabIndex        =   36
         Top             =   765
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MOCusto 
         Height          =   270
         Left            =   3570
         TabIndex        =   37
         Top             =   1035
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridMO 
         Height          =   315
         Left            =   30
         TabIndex        =   6
         Top             =   190
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   556
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         Caption         =   "Custo Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   29
         Left            =   6300
         TabIndex        =   39
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label MOCustoTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7455
         TabIndex        =   38
         Top             =   1470
         Width           =   1410
      End
   End
   Begin VB.Frame FrameMaq 
      Caption         =   "Máquinas"
      Height          =   1845
      Index           =   3
      Left            =   60
      TabIndex        =   21
      Top             =   4590
      Width           =   9405
      Begin VB.CommandButton BotaoMaq 
         Caption         =   "Máquinas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   11
         Top             =   1455
         Width           =   1600
      End
      Begin VB.TextBox MaqCodigo 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   165
         MaxLength       =   20
         TabIndex        =   24
         Top             =   690
         Width           =   2220
      End
      Begin VB.TextBox MaqOBS 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   6255
         MaxLength       =   50
         TabIndex        =   22
         Top             =   705
         Width           =   2655
      End
      Begin MSMask.MaskEdBox MaqCustoT 
         Height          =   270
         Left            =   5115
         TabIndex        =   23
         Top             =   705
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaqCusto 
         Height          =   270
         Left            =   4050
         TabIndex        =   25
         Top             =   705
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaqHoras 
         Height          =   270
         Left            =   2925
         TabIndex        =   26
         Top             =   735
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaqQuantidade 
         Height          =   270
         Left            =   1950
         TabIndex        =   27
         Top             =   720
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridMaq 
         Height          =   345
         Left            =   30
         TabIndex        =   10
         Top             =   190
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   609
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         Caption         =   "Custo Total:"
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
         Index           =   33
         Left            =   6375
         TabIndex        =   29
         Top             =   1515
         Width           =   1050
      End
      Begin VB.Label MaqCustoTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7455
         TabIndex        =   28
         Top             =   1470
         Width           =   1410
      End
   End
   Begin VB.ComboBox Etapa 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   420
      Width           =   2970
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1890
      Picture         =   "ApontamentoPRJ.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   75
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ApontamentoPRJ.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ApontamentoPRJ.ctx":0268
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ApontamentoPRJ.ctx":079A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ApontamentoPRJ.ctx":0924
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   300
      Left            =   1110
      TabIndex        =   0
      Top             =   60
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "999999"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataInicio 
      Height          =   300
      Left            =   5340
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   4185
      TabIndex        =   2
      Top             =   60
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Projeto 
      Height          =   300
      Left            =   1110
      TabIndex        =   4
      Top             =   435
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   375
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   20
      Top             =   465
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
      Height          =   195
      Index           =   41
      Left            =   3540
      TabIndex        =   19
      Top             =   465
      Width           =   570
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
      Left            =   3615
      TabIndex        =   18
      Top             =   90
      Width           =   480
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
      Left            =   405
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   90
      Width           =   660
   End
End
Attribute VB_Name = "ApontamentoPRJ"
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

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoMaq As AdmEvento
Attribute objEventoMaq.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1
Private WithEvents objEventoMP As AdmEvento
Attribute objEventoMP.VB_VarHelpID = -1

Public iAlterado As Integer

'Grid de Matéria prima
Dim objGridMP As AdmGrid
Dim iGrid_MPProduto_Col As Integer
Dim iGrid_MPDescricao_Col As Integer
Dim iGrid_MPCusto_Col As Integer
Dim iGrid_MPUM_Col As Integer
Dim iGrid_MPQuantidade_Col As Integer
Dim iGrid_MPCustoT_Col As Integer
Dim iGrid_MPOBS_Col As Integer

'Grid de máquinas
Dim objGridMaq As AdmGrid
Dim iGrid_MaqCodigo_Col As Integer
'Dim iGrid_MaqDescricao_Col As Integer
Dim iGrid_MaqCusto_Col As Integer
Dim iGrid_MaqHoras_Col As Integer
Dim iGrid_MaqQuantidade_Col As Integer
Dim iGrid_MaqCustoT_Col As Integer
Dim iGrid_MaqOBS_Col As Integer

'Grid de mão de obra
Dim objGridMO As AdmGrid
Dim iGrid_MOCodigo_Col As Integer
Dim iGrid_MODescricao_Col As Integer
Dim iGrid_MOCusto_Col As Integer
Dim iGrid_MOHoras_Col As Integer
Dim iGrid_MOQuantidade_Col As Integer
Dim iGrid_MOCustoT_Col As Integer
Dim iGrid_MOOBS_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Apontamento do Projeto"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "ApontamentoPRJ"
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

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelCodigo_Click()

Dim objApontPRJ As New ClassApontPRJ
Dim colSelecao As New Collection

    'Preenche objFornecedor com NomeReduzido da tela
    objApontPRJ.lCodigo = StrParaLong(Codigo.Text)

    Call Chama_Tela("ApontamentoPRJLista", colSelecao, objApontPRJ, objEventoCodigo)
    
End Sub

Private Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is MPProduto Then
            Call BotaoMP_Click
        ElseIf Me.ActiveControl Is MOCodigo Then
            Call BotaoMO_Click
        ElseIf Me.ActiveControl Is MaqCodigo Then
            Call BotaoMaq_Click
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

    Set gobjTelaProjetoInfo = Nothing

    Set objEventoCodigo = Nothing
    Set objEventoMaq = Nothing
    Set objEventoMO = Nothing
    Set objEventoMP = Nothing
    
    Set objGridMP = Nothing
    Set objGridMO = Nothing
    Set objGridMaq = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_UnLoad:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194533)

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

    Set objEventoCodigo = New AdmEvento
    Set objEventoMaq = New AdmEvento
    Set objEventoMO = New AdmEvento
    Set objEventoMP = New AdmEvento
    
    Set objGridMP = New AdmGrid

    lErro = Inicializa_GridMP(objGridMP)
    If lErro <> SUCESSO Then gError 194534
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MPProduto)
    If lErro <> SUCESSO Then gError 194535

    Set objGridMO = New AdmGrid

    lErro = Inicializa_GridMO(objGridMO)
    If lErro <> SUCESSO Then gError 194536
    
    Set objGridMaq = New AdmGrid
    
    lErro = Inicializa_GridMaq(objGridMaq)
    If lErro <> SUCESSO Then gError 194537
    
    lErro = Inicializa_Mascara_Projeto(Projeto)
    If lErro <> SUCESSO Then gError 194538
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 194534 To 194538

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194539)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objApontPRJ Is Nothing) Then

        lErro = Traz_ApontPRJ_Tela(objApontPRJ)
        If lErro <> SUCESSO Then gError 194540

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 194540

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194541)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria
   
    objApontPRJ.lNumIntDocPRJ = glNumIntPRJ
    objApontPRJ.lNumIntDocEtapa = glNumIntPRJEtapa
    'objApontPRJ.sDescricao = Descricao.Text
    'objApontPRJ.sObservacao = Descricao.Text
    objApontPRJ.dtData = Data.Text
    objApontPRJ.lCodigo = StrParaLong(Codigo.Text)
    
    lErro = Move_MP_Memoria(objApontPRJ)
    If lErro <> SUCESSO Then gError 194542

    lErro = Move_MO_Memoria(objApontPRJ)
    If lErro <> SUCESSO Then gError 194543

    lErro = Move_Maq_Memoria(objApontPRJ)
    If lErro <> SUCESSO Then gError 194544

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 194542 To 194544
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194545)

    End Select

    Exit Function

End Function

Function Move_MP_Memoria(objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer
Dim objApontProdPRJ As ClassApontProdPRJ

On Error GoTo Erro_Move_MP_Memoria

    For iIndice = 1 To objGridMP.iLinhasExistentes
    
        Set objApontProdPRJ = New ClassApontProdPRJ
        
        lErro = CF("Produto_Formata", GridMP.TextMatrix(iIndice, iGrid_MPProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 194546
        
        objApontProdPRJ.sProduto = sProdutoFormatado
        objApontProdPRJ.dQtd = StrParaDbl(GridMP.TextMatrix(iIndice, iGrid_MPQuantidade_Col))
        objApontProdPRJ.dCusto = StrParaDbl(GridMP.TextMatrix(iIndice, iGrid_MPCustoT_Col))
        objApontProdPRJ.sUM = GridMP.TextMatrix(iIndice, iGrid_MPUM_Col)
        objApontProdPRJ.iSeq = iIndice
        objApontProdPRJ.sOBS = GridMP.TextMatrix(iIndice, iGrid_MPOBS_Col)
        
        objApontPRJ.colMateriaPrima.Add objApontProdPRJ
    
    Next
    
    Move_MP_Memoria = SUCESSO

    Exit Function

Erro_Move_MP_Memoria:

    Move_MP_Memoria = gErr

    Select Case gErr
    
        Case 194546

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194547)

    End Select

    Exit Function

End Function

Function Move_MO_Memoria(objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim objApontMOPRJ As ClassApontMOPRJ

On Error GoTo Erro_Move_MO_Memoria

    For iIndice = 1 To objGridMO.iLinhasExistentes
    
        Set objApontMOPRJ = New ClassApontMOPRJ
        
        objApontMOPRJ.iCodMO = StrParaInt(GridMO.TextMatrix(iIndice, iGrid_MOCodigo_Col))
        objApontMOPRJ.iQtd = StrParaInt(GridMO.TextMatrix(iIndice, iGrid_MOQuantidade_Col))
        objApontMOPRJ.dHoras = StrParaDbl(GridMO.TextMatrix(iIndice, iGrid_MOHoras_Col))
        objApontMOPRJ.dCusto = StrParaDbl(GridMO.TextMatrix(iIndice, iGrid_MOCustoT_Col))
        objApontMOPRJ.iSeq = iIndice
        objApontMOPRJ.sOBS = GridMO.TextMatrix(iIndice, iGrid_MOOBS_Col)
        
        objApontPRJ.colMaoDeObra.Add objApontMOPRJ
    
    Next

    Move_MO_Memoria = SUCESSO

    Exit Function

Erro_Move_MO_Memoria:

    Move_MO_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194548)

    End Select

    Exit Function

End Function

Function Move_Maq_Memoria(objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim iTipo As Integer
Dim iIndice As Integer
Dim objApontMaqPRJ As ClassApontMaqPRJ
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Move_Maq_Memoria

    For iIndice = 1 To objGridMaq.iLinhasExistentes
    
        Set objApontMaqPRJ = New ClassApontMaqPRJ
        Set objMaquina = New ClassMaquinas
        
        objMaquina.sNomeReduzido = GridMaq.TextMatrix(iIndice, iGrid_MaqCodigo_Col)
        
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 194549
        
        objApontMaqPRJ.iCodMaq = objMaquina.iCodigo
        objApontMaqPRJ.iQtd = StrParaInt(GridMaq.TextMatrix(iIndice, iGrid_MaqQuantidade_Col))
        objApontMaqPRJ.dHoras = StrParaDbl(GridMaq.TextMatrix(iIndice, iGrid_MaqHoras_Col))
        objApontMaqPRJ.dCusto = StrParaDbl(GridMaq.TextMatrix(iIndice, iGrid_MaqCustoT_Col))
        objApontMaqPRJ.iSeq = iIndice
        objApontMaqPRJ.sOBS = GridMaq.TextMatrix(iIndice, iGrid_MaqOBS_Col)
        
        objApontPRJ.colMaquinas.Add objApontMaqPRJ
    
    Next

    Move_Maq_Memoria = SUCESSO

    Exit Function

Erro_Move_Maq_Memoria:

    Move_Maq_Memoria = gErr

    Select Case gErr
    
        Case 194549

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194550)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ApontPRJ"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocPRJ", glNumIntPRJ, 0, "NumIntDocPRJ"
    colCampoValor.Add "NumIntDocEtapa", glNumIntPRJEtapa, 0, "NumIntDocEtapa"
    colCampoValor.Add "Codigo", StrParaLong(Codigo.Text), 0, "Codigo"
    
    'Filtros para o Sistema de Setas

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 194551

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194552)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objApontPRJ As New ClassApontPRJ

On Error GoTo Erro_Tela_Preenche

    objApontPRJ.lCodigo = colCampoValor.Item("Codigo").vValor
    objApontPRJ.lNumIntDocPRJ = colCampoValor.Item("NumIntDocPRJ").vValor
    objApontPRJ.lNumIntDocEtapa = colCampoValor.Item("NumIntDocEtapa").vValor

    lErro = Traz_ApontPRJ_Tela(objApontPRJ)
    If lErro <> SUCESSO Then gError 194553

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 194553

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194554)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objApontPRJ As New ClassApontPRJ

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 194555
    If Len(Trim(Projeto.ClipText)) = 0 Then gError 194556
    If Len(Trim(Data.ClipText)) = 0 Then gError 194557
    If Len(Trim(Etapa.Text)) = 0 Then gError 194558

    'Preenche o objProjetos
    lErro = Move_Tela_Memoria(objApontPRJ)
    If lErro <> SUCESSO Then gError 194559
    
    lErro = Critica_Dados(objApontPRJ)
    If lErro <> SUCESSO Then gError 194560

    lErro = Trata_Alteracao(objApontPRJ, objApontPRJ.lCodigo)
    If lErro <> SUCESSO Then gError 194561

    'Grava a etapa no Banco de Dados
    lErro = CF("ApontPRJ_Grava", objApontPRJ)
    If lErro <> SUCESSO Then gError 194562

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 194555
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_APONT_PRJ_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
        
        Case 194556
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRJ_NAO_PREENCHIDO", gErr)
            Projeto.SetFocus
            
        Case 194557
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_APONTAMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus
            
        Case 194558
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_ETAPA_NAO_PREENCHIDO2", gErr)
            Etapa.SetFocus
            
        Case 194559 To 194562

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194563)

    End Select

    Exit Function

End Function

Function Critica_Dados(ByVal objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim objApontMaqPRJ As ClassApontMaqPRJ
Dim objApontProdPRJ As ClassApontProdPRJ
Dim objApontMOPRJ As ClassApontMOPRJ
Dim iLinha As Integer

On Error GoTo Erro_Critica_Dados
     
    For Each objApontMOPRJ In objApontPRJ.colMaoDeObra
    
        iLinha = objApontMOPRJ.iSeq
        
        If objApontMOPRJ.dHoras = 0 Then gError 194564
        If objApontMOPRJ.iQtd = 0 Then gError 194565
    
    Next
    
    For Each objApontMaqPRJ In objApontPRJ.colMaquinas
    
        iLinha = objApontMaqPRJ.iSeq
    
        If objApontMaqPRJ.dHoras = 0 Then gError 194566
        If objApontMaqPRJ.iQtd = 0 Then gError 194567
    
    Next
    
    For Each objApontProdPRJ In objApontPRJ.colMateriaPrima
    
        iLinha = objApontProdPRJ.iSeq
   
        If objApontProdPRJ.dQtd = 0 Then gError 194568
    
    Next

    GL_objMDIForm.MousePointer = vbDefault

    Critica_Dados = SUCESSO

    Exit Function

Erro_Critica_Dados:

    Critica_Dados = gErr

    Select Case gErr
        
        Case 194564
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MOHORAS_NAO_PREECHIDA", gErr, "", iLinha)

        Case 194565
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MOQTD_NAO_PREECHIDA", gErr, "", iLinha)

        Case 194566
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MAQHORAS_NAO_PREECHIDA", gErr, "", iLinha)

        Case 194567
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MAQQTD_NAO_PREECHIDA", gErr, "", iLinha)

        Case 194568
            Call Rotina_Erro(vbOKOnly, "ERRO_ETAPA_MPQUANTIDADE_NAO_PREECHIDA", gErr, "", iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194569)

    End Select

    Exit Function

End Function

Function Limpa_Tela_ApontPRJ() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_ApontPRJ

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridMP)
    Call Grid_Limpa(objGridMO)
    Call Grid_Limpa(objGridMaq)
    
    glNumIntPRJ = 0
    glNumIntPRJEtapa = 0
            
    Etapa.Clear
        
    sProjetoAnt = ""
    sEtapaAnt = ""
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    MPCustoTotal.Caption = ""
    MOCustoTotal.Caption = ""
    MaqCustoTotal.Caption = ""

    iAlterado = 0

    Limpa_Tela_ApontPRJ = SUCESSO

    Exit Function

Erro_Limpa_Tela_ApontPRJ:

    Limpa_Tela_ApontPRJ = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194570)

    End Select

    Exit Function

End Function

Function Traz_ApontPRJ_Tela(objApontPRJ As ClassApontPRJ) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjetos
Dim objEtapa As New ClassPRJEtapas

On Error GoTo Erro_Traz_ApontPRJ_Tela

    Call Limpa_Tela_ApontPRJ

    'Lê a Etapa que está sendo Passada
    lErro = CF("ApontPRJ_Le", objApontPRJ)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194571

    Codigo.PromptInclude = False
    Codigo.Text = objApontPRJ.lCodigo
    Codigo.PromptInclude = True

    If lErro = SUCESSO Then
    
        objProjeto.lNumIntDoc = objApontPRJ.lNumIntDocPRJ
        
        If objProjeto.lNumIntDoc <> 0 Then
        
            lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194572
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194573
            
        End If
        
        glNumIntPRJ = objProjeto.lNumIntDoc
    
        objEtapa.lNumIntDoc = objApontPRJ.lNumIntDocEtapa
        
        If objEtapa.lNumIntDoc <> 0 Then
        
            lErro = CF("PRJEtapas_Le_NumIntDoc", objEtapa)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194574
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194575
            
        End If
        
        glNumIntPRJEtapa = objEtapa.lNumIntDoc
        
        lErro = Retorno_Projeto_Tela(Projeto, objProjeto.sCodigo)
        If lErro <> SUCESSO Then gError 194576
        
        sProjetoAnt = Projeto.Text
        
        sEtapaAnt = objEtapa.sCodigo
            
        Call gobjTelaProjetoInfo.Trata_Etapa(glNumIntPRJ, Etapa)
        
        Call CF("SCombo_Seleciona2", Etapa, sEtapaAnt)
    
        Data.PromptInclude = False
        Data.Text = Format(objApontPRJ.dtData, "dd/mm/yy")
        Data.PromptInclude = True
    
        lErro = Traz_MP_Tela(objApontPRJ.colMateriaPrima)
        If lErro <> SUCESSO Then gError 194577
    
        lErro = Traz_MO_Tela(objApontPRJ.colMaoDeObra)
        If lErro <> SUCESSO Then gError 194578
    
        lErro = Traz_Maq_Tela(objApontPRJ.colMaquinas)
        If lErro <> SUCESSO Then gError 194579

    End If

    iAlterado = 0

    Traz_ApontPRJ_Tela = SUCESSO

    Exit Function

Erro_Traz_ApontPRJ_Tela:

    Traz_ApontPRJ_Tela = gErr

    Select Case gErr

        Case 194571, 194572, 194574, 194576 To 194579
        
        Case 194573
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO", gErr, objProjeto.lNumIntDoc)
            
        Case 194575
            Call Rotina_Erro(vbOKOnly, "ERRO_PRJETAPAS_NAO_CADASTRADO", gErr, objEtapa.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194580)

    End Select

    Exit Function

End Function

Function Traz_MP_Tela(ByVal colMP As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objApontProdPRJ As ClassApontProdPRJ
Dim sProdutoMascarado As String
Dim dCustoTotal As Double
Dim objProduto As ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_MP_Tela
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objApontProdPRJ In colMP
        
        iLinha = iLinha + 1
        
        Set objProduto = New ClassProduto
                
        lErro = Mascara_RetornaProdutoTela(objApontProdPRJ.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 194581
               
        MPProduto.PromptInclude = False
        MPProduto.Text = sProdutoMascarado
        MPProduto.PromptInclude = True
        
        lErro = CF("Produto_Critica2", MPProduto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 194582
        
        GridMP.TextMatrix(iLinha, iGrid_MPCusto_Col) = Format(objApontProdPRJ.dCusto / objApontProdPRJ.dQtd, "STANDARD")
        GridMP.TextMatrix(iLinha, iGrid_MPDescricao_Col) = objProduto.sDescricao
        GridMP.TextMatrix(iLinha, iGrid_MPProduto_Col) = MPProduto.Text
        GridMP.TextMatrix(iLinha, iGrid_MPQuantidade_Col) = Formata_Estoque(objApontProdPRJ.dQtd)
        GridMP.TextMatrix(iLinha, iGrid_MPUM_Col) = objApontProdPRJ.sUM
        GridMP.TextMatrix(iLinha, iGrid_MPCustoT_Col) = Format(objApontProdPRJ.dCusto, "STANDARD")
        GridMP.TextMatrix(iLinha, iGrid_MPOBS_Col) = objApontProdPRJ.sOBS
        
        dCustoTotal = dCustoTotal + objApontProdPRJ.dCusto
    
    Next

    MPCustoTotal.Caption = Format(dCustoTotal, "STANDARD")

    objGridMP.iLinhasExistentes = colMP.Count

    Traz_MP_Tela = SUCESSO

    Exit Function

Erro_Traz_MP_Tela:

    Traz_MP_Tela = gErr

    Select Case gErr
    
        Case 194581, 194582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194583)

    End Select

    Exit Function

End Function

Function Traz_MO_Tela(ByVal colMO As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objApontMOPRJ As ClassApontMOPRJ
Dim dCustoTotal As Double
Dim objMO As ClassTiposDeMaodeObra

On Error GoTo Erro_Traz_MO_Tela
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objApontMOPRJ In colMO
        
        iLinha = iLinha + 1
        
        Set objMO = New ClassTiposDeMaodeObra
        
        objMO.iCodigo = objApontMOPRJ.iCodMO
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objMO)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 194584
        
        GridMO.TextMatrix(iLinha, iGrid_MOCusto_Col) = Format(objApontMOPRJ.dCusto / objApontMOPRJ.dHoras / objApontMOPRJ.iQtd, "STANDARD")
        GridMO.TextMatrix(iLinha, iGrid_MODescricao_Col) = objMO.sDescricao
        GridMO.TextMatrix(iLinha, iGrid_MOCodigo_Col) = objApontMOPRJ.iCodMO
        GridMO.TextMatrix(iLinha, iGrid_MOQuantidade_Col) = objApontMOPRJ.iQtd
        GridMO.TextMatrix(iLinha, iGrid_MOHoras_Col) = Formata_Estoque(objApontMOPRJ.dHoras)
        GridMO.TextMatrix(iLinha, iGrid_MOCustoT_Col) = Format(objApontMOPRJ.dCusto, "STANDARD")
        GridMO.TextMatrix(iLinha, iGrid_MOOBS_Col) = objApontMOPRJ.sOBS
        
        dCustoTotal = dCustoTotal + objApontMOPRJ.dCusto
    
    Next

    MOCustoTotal.Caption = Format(dCustoTotal, "STANDARD")

    objGridMO.iLinhasExistentes = colMO.Count

    Traz_MO_Tela = SUCESSO

    Exit Function

Erro_Traz_MO_Tela:

    Traz_MO_Tela = gErr

    Select Case gErr
    
        Case 194584

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194585)

    End Select

    Exit Function

End Function

Function Traz_Maq_Tela(ByVal colMaq As Collection) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objApontMaqPRJ As ClassApontMaqPRJ
Dim objMaquina As ClassMaquinas
Dim dCustoTotal As Double

On Error GoTo Erro_Traz_Maq_Tela
    
    'Exibe os dados da coleção de Competencias na tela
    For Each objApontMaqPRJ In colMaq
        
        iLinha = iLinha + 1
        
        Set objMaquina = New ClassMaquinas
        
        objMaquina.iCodigo = objApontMaqPRJ.iCodMaq
        objMaquina.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("Maquinas_Le", objMaquina)
        If lErro <> SUCESSO And lErro <> 103090 Then gError 194586
        
        GridMaq.TextMatrix(iLinha, iGrid_MaqCusto_Col) = Format(objApontMaqPRJ.dCusto / objApontMaqPRJ.dHoras / objApontMaqPRJ.iQtd, "STANDARD")
        'GridMaq.TextMatrix(iLinha, iGrid_MaqDescricao_Col) = objMaquina.sDescricao
        GridMaq.TextMatrix(iLinha, iGrid_MaqCodigo_Col) = objMaquina.sNomeReduzido
        GridMaq.TextMatrix(iLinha, iGrid_MaqQuantidade_Col) = objApontMaqPRJ.iQtd
        GridMaq.TextMatrix(iLinha, iGrid_MaqHoras_Col) = Formata_Estoque(objApontMaqPRJ.dHoras)
        GridMaq.TextMatrix(iLinha, iGrid_MaqCustoT_Col) = Format(objApontMaqPRJ.dCusto, "STANDARD")
        GridMaq.TextMatrix(iLinha, iGrid_MaqOBS_Col) = objApontMaqPRJ.sOBS
    
        dCustoTotal = dCustoTotal + objApontMaqPRJ.dCusto
        
    Next
    
    MaqCustoTotal.Caption = Format(dCustoTotal, "STANDARD")

    objGridMaq.iLinhasExistentes = colMaq.Count

    Traz_Maq_Tela = SUCESSO

    Exit Function

Erro_Traz_Maq_Tela:

    Traz_Maq_Tela = gErr

    Select Case gErr
    
        Case 194586

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194587)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 194588

    'Limpa Tela
    Call Limpa_Tela_ApontPRJ

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 194588

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194589)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194590)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 194591

    Call Limpa_Tela_ApontPRJ

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 194591

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194592)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objApontPRJ As New ClassApontPRJ
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 194593
    
    lErro = Move_Tela_Memoria(objApontPRJ)
    If lErro <> SUCESSO Then gError 194594

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_APONTPRJ", objApontPRJ.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("ApontPRJ_Exclui", objApontPRJ)
        If lErro <> SUCESSO Then gError 194595

        'Limpa Tela
        Call Limpa_Tela_ApontPRJ

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
        
        Case 194593
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_APONT_PRJ_NAO_PREENCHIDO", gErr)
        
        Case 194594, 194595

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194596)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 194597

        Data.Text = sData
        
        Call Data_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 194597

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194598)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 194599

        Data.Text = sData
        
        Call Data_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 194599

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194600)

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
        If lErro <> SUCESSO Then gError 194601
    
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 194601

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194602)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objApontPRJ As ClassApontPRJ

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objApontPRJ = obj1

    'Mostra os dados do CentrodeTrabalho na tela
    lErro = Traz_ApontPRJ_Tela(objApontPRJ)
    If lErro <> SUCESSO Then gError 194603
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 194603

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194604)

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

Private Function Inicializa_GridMP(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("Custo")
    objGrid.colColuna.Add ("Custo T.")
    objGrid.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MPProduto.Name)
    objGrid.colCampo.Add (MPDescricao.Name)
    objGrid.colCampo.Add (MPUM.Name)
    objGrid.colCampo.Add (MPQuantidade.Name)
    objGrid.colCampo.Add (MPCusto.Name)
    objGrid.colCampo.Add (MPCustoT.Name)
    objGrid.colCampo.Add (MPOBS.Name)
       
    'Colunas do Grid
    iGrid_MPProduto_Col = 1
    iGrid_MPDescricao_Col = 2
    iGrid_MPUM_Col = 3
    iGrid_MPQuantidade_Col = 4
    iGrid_MPCusto_Col = 5
    iGrid_MPCustoT_Col = 6
    iGrid_MPOBS_Col = 7
    
    objGrid.objGrid = GridMP

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridMP.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMP = SUCESSO

End Function

Private Sub GridMP_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMP, iAlterado)
    End If

End Sub

Private Sub GridMP_GotFocus()
    Call Grid_Recebe_Foco(objGridMP)
End Sub

Private Sub GridMP_EnterCell()
    Call Grid_Entrada_Celula(objGridMP, iAlterado)
End Sub

Private Sub GridMP_LeaveCell()
    Call Saida_Celula(objGridMP)
End Sub

Private Sub GridMP_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMP, iAlterado)
    End If

End Sub

Private Sub GridMP_RowColChange()
    Call Grid_RowColChange(objGridMP)
End Sub

Private Sub GridMP_Scroll()
    Call Grid_Scroll(objGridMP)
End Sub

Private Sub GridMP_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMP)

    Call Soma_Coluna_Grid(objGridMP, iGrid_MPCustoT_Col, MPCustoTotal)
End Sub

Private Sub GridMP_LostFocus()
    Call Grid_Libera_Foco(objGridMP)
End Sub

Private Function Inicializa_GridMaq(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Máquina")
    'objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Custo UN/h")
    objGrid.colColuna.Add ("Custo T.")
    objGrid.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MaqCodigo.Name)
    'objGrid.colCampo.Add (MaqDescricao.Name)
    objGrid.colCampo.Add (MaqQuantidade.Name)
    objGrid.colCampo.Add (MaqHoras.Name)
    objGrid.colCampo.Add (MaqCusto.Name)
    objGrid.colCampo.Add (MaqCustoT.Name)
    objGrid.colCampo.Add (MaqOBS.Name)
    
    'Colunas do Grid
    iGrid_MaqCodigo_Col = 1
    'iGrid_MaqDescricao_Col = 2
    iGrid_MaqQuantidade_Col = 2
    iGrid_MaqHoras_Col = 3
    iGrid_MaqCusto_Col = 4
    iGrid_MaqCustoT_Col = 5
    iGrid_MaqOBS_Col = 6
    
    objGrid.objGrid = GridMaq

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridMaq.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaq = SUCESSO

End Function

Private Sub GridMaq_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq, iAlterado)
    End If

End Sub

Private Sub GridMaq_GotFocus()
    Call Grid_Recebe_Foco(objGridMaq)
End Sub

Private Sub GridMaq_EnterCell()
    Call Grid_Entrada_Celula(objGridMaq, iAlterado)
End Sub

Private Sub GridMaq_LeaveCell()
    Call Saida_Celula(objGridMaq)
End Sub

Private Sub GridMaq_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaq, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaq, iAlterado)
    End If

End Sub

Private Sub GridMaq_RowColChange()
    Call Grid_RowColChange(objGridMaq)
End Sub

Private Sub GridMaq_Scroll()
    Call Grid_Scroll(objGridMaq)
End Sub

Private Sub GridMaq_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMaq)
    
    Call Soma_Coluna_Grid(objGridMaq, iGrid_MaqCustoT_Col, MaqCustoTotal)

End Sub

Private Sub GridMaq_LostFocus()
    Call Grid_Libera_Foco(objGridMaq)
End Sub

Private Function Inicializa_GridMO(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Código")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Qtde")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Custo UN/h")
    objGrid.colColuna.Add ("Custo T.")
    objGrid.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (MOCodigo.Name)
    objGrid.colCampo.Add (MODescricao.Name)
    objGrid.colCampo.Add (MOQuantidade.Name)
    objGrid.colCampo.Add (MOHoras.Name)
    objGrid.colCampo.Add (MOCusto.Name)
    objGrid.colCampo.Add (MOCustoT.Name)
    objGrid.colCampo.Add (MOOBS.Name)
    
    'Colunas do Grid
    iGrid_MOCodigo_Col = 1
    iGrid_MODescricao_Col = 2
    iGrid_MOQuantidade_Col = 3
    iGrid_MOHoras_Col = 4
    iGrid_MOCusto_Col = 5
    iGrid_MOCustoT_Col = 6
    iGrid_MOOBS_Col = 7
    
    objGrid.objGrid = GridMO

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridMO.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMO = SUCESSO

End Function

Private Sub GridMO_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Private Sub GridMO_GotFocus()
    Call Grid_Recebe_Foco(objGridMO)
End Sub

Private Sub GridMO_EnterCell()
    Call Grid_Entrada_Celula(objGridMO, iAlterado)
End Sub

Private Sub GridMO_LeaveCell()
    Call Saida_Celula(objGridMO)
End Sub

Private Sub GridMO_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMO, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMO, iAlterado)
    End If

End Sub

Private Sub GridMO_RowColChange()
    Call Grid_RowColChange(objGridMO)
End Sub

Private Sub GridMO_Scroll()
    Call Grid_Scroll(objGridMO)
End Sub

Private Sub GridMO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMO)

    Call Soma_Coluna_Grid(objGridMO, iGrid_MOCustoT_Col, MOCustoTotal)
End Sub

Private Sub GridMO_LostFocus()
    Call Grid_Libera_Foco(objGridMO)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
    If objGridInt.objGrid.Name = GridMP.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
            
                Case iGrid_MPProduto_Col
                
                    lErro = Saida_Celula_MPProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 194605

                Case iGrid_MPQuantidade_Col
                
                    lErro = Saida_Celula_MPQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 194606

                Case iGrid_MPCusto_Col
                
                    lErro = Saida_Celula_MPCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 194607

                Case iGrid_MPOBS_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, MPOBS)
                    If lErro <> SUCESSO Then gError 194608

            End Select
                    
                
        ElseIf objGridInt.objGrid.Name = GridMO.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MOCodigo_Col
                
                    lErro = Saida_Celula_MOCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 194609

                Case iGrid_MOQuantidade_Col
                
                    lErro = Saida_Celula_MOQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 194610

                Case iGrid_MOHoras_Col
                
                    lErro = Saida_Celula_MOHoras(objGridInt)
                    If lErro <> SUCESSO Then gError 194611

                Case iGrid_MOCusto_Col
                
                    lErro = Saida_Celula_MOCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 194612

                Case iGrid_MOOBS_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, MOOBS)
                    If lErro <> SUCESSO Then gError 194613
        
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridMaq.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_MaqCodigo_Col
                
                    lErro = Saida_Celula_MaqCodigo(objGridInt)
                    If lErro <> SUCESSO Then gError 194614

                Case iGrid_MaqQuantidade_Col
                
                    lErro = Saida_Celula_MaqQuantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 194615

                Case iGrid_MaqHoras_Col
                
                    lErro = Saida_Celula_MaqHoras(objGridInt)
                    If lErro <> SUCESSO Then gError 194616

                Case iGrid_MaqCusto_Col
                
                    lErro = Saida_Celula_MaqCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 194617

                Case iGrid_MaqOBS_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, MaqOBS)
                    If lErro <> SUCESSO Then gError 194618
                    
            End Select
                        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 194619

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 194604 To 194618
            'erros tratatos nas rotinas chamadas
        
        Case 194619
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194620)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatadoMP As String
Dim iProdutoPreenchidoMP As Integer
Dim iMaquinaPreenchida As Integer
Dim iTipoMaoDeObraPreenchida As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    lErro = CF("Produto_Formata", GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col), sProdutoFormatadoMP, iProdutoPreenchidoMP)
    If lErro <> SUCESSO Then gError 194621
    
    If Len(Trim(GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col))) > 0 Then
        iTipoMaoDeObraPreenchida = MARCADO
    Else
        iTipoMaoDeObraPreenchida = DESMARCADO
    End If
    
    If Len(Trim(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCodigo_Col))) > 0 Then
        iMaquinaPreenchida = MARCADO
    Else
        iMaquinaPreenchida = DESMARCADO
    End If
    
        
    Select Case objControl.Name
    

        Case MPProduto.Name
        
            If iProdutoPreenchidoMP = PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MPQuantidade.Name, MPCusto.Name, MPOBS.Name

            If iProdutoPreenchidoMP <> PRODUTO_PREENCHIDO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
                        
        Case MOCodigo.Name
        
            If iTipoMaoDeObraPreenchida = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MOQuantidade.Name, MOCusto.Name, MOHoras.Name, MOOBS.Name

            If iTipoMaoDeObraPreenchida <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case MaqCodigo.Name
        
            If iMaquinaPreenchida = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case MaqQuantidade.Name, MaqCusto.Name, MaqHoras.Name, MaqOBS.Name

            If iMaquinaPreenchida <> MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 194621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194622)

    End Select

    Exit Sub

End Sub

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
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
            If lErro <> SUCESSO Then gError 194623
            
            objProjeto.sCodigo = sProjeto
            objProjeto.iFilialEmpresa = giFilialEmpresa
            
            'Le
            lErro = CF("Projetos_Le", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194624
            
            'Se não encontrou => Erro
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 194625
            
            If sProjetoAnt <> Projeto.Text Then
                Call gobjTelaProjetoInfo.Trata_Etapa(objProjeto.lNumIntDoc, Etapa)
            End If
            
            If Len(Trim(Etapa.Text)) > 0 Then
            
                objEtapa.lNumIntDocPRJ = objProjeto.lNumIntDoc
                objEtapa.sCodigo = SCodigo_Extrai(Etapa.Text)
            
                lErro = CF("PrjEtapas_Le", objEtapa)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194626
            
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
    
        Case 194623, 194624, 194626
        
        Case 194625
            Call Rotina_Erro(vbOKOnly, "ERRO_PROJETOS_NAO_CADASTRADO2", gErr, objProjeto.sCodigo, objProjeto.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194627)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_PreencheMP(objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_ProdutoLinha_PreencheMP

    GridMP.TextMatrix(GridMP.Row, iGrid_MPDescricao_Col) = objProduto.sDescricao
    GridMP.TextMatrix(GridMP.Row, iGrid_MPUM_Col) = objProduto.sSiglaUMEstoque
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridMP.Row - GridMP.FixedRows = objGridMP.iLinhasExistentes Then
        objGridMP.iLinhasExistentes = objGridMP.iLinhasExistentes + 1
    End If

    ProdutoLinha_PreencheMP = SUCESSO

    Exit Function

Erro_ProdutoLinha_PreencheMP:

    ProdutoLinha_PreencheMP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 194628)

    End Select

    Exit Function

End Function

Private Sub MPProduto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPProduto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPProduto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPProduto
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPQuantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPQuantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPQuantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPDescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPDescricao
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPCusto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPCusto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPCusto
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoMP_Click()

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMP_Click

    If Me.ActiveControl Is MPProduto Then
        sProduto = MPProduto.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMP.Row = 0 Then gError 194629
        
        sProduto = GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col)
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 194630
    
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""
    
    objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutosKitLista", colSelecao, objProduto, objEventoMP)
    
    Exit Sub

Erro_BotaoMP_Click:

    Select Case gErr

        Case 194629
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 194630
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194631)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMP_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objProdutoKit As New ClassProdutoKit

On Error GoTo Erro_objEventoMP_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 194632
        
    MPProduto.PromptInclude = False
    MPProduto.Text = sProdutoMascarado
    MPProduto.PromptInclude = True
        
    If Not (Me.ActiveControl Is MPProduto) Then
        
        GridMP.TextMatrix(GridMP.Row, iGrid_MPProduto_Col) = MPProduto.Text

        lErro = ProdutoLinha_PreencheMP(objProduto)
        If lErro <> SUCESSO Then gError 194633
        
    End If
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMP_evSelecao:

    Select Case gErr

        Case 194632, 194633
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194634)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_MPProduto(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sProdutoMascarado As String

On Error GoTo Erro_Saida_Celula_MPProduto

    Set objGridInt.objControle = MPProduto

    lErro = CF("Produto_Formata", MPProduto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 194635
    
    'se o produto foi preenchido
    If Len(Trim(MPProduto.ClipText)) <> 0 Then
        
        lErro = CF("Produto_Critica2", MPProduto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 194636
        
        'mascara produto escolhido
        lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 194637

        MPProduto.PromptInclude = False
        MPProduto.Text = sProdutoMascarado
        MPProduto.PromptInclude = True
        
        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                
            'se é um produto gerencial  ==> erro
            If lErro = 25043 Then gError 194638
            
            'se não está cadastrado
            If lErro = 25041 Then gError 194639
        
             'Preenche a linha do grid
            lErro = ProdutoLinha_PreencheMP(objProduto)
            If lErro <> SUCESSO Then gError 194640
            
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194641

    Saida_Celula_MPProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_MPProduto:

    Saida_Celula_MPProduto = gErr

    Select Case gErr

        Case 194635, 194636, 194637, 194640, 194641
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 194638
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 194639
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", MPProduto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = MPProduto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194642)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPQuantidade

    Set objGridInt.objControle = MPQuantidade
    
    'se a quantidade foi preenchida
    If Len(MPQuantidade.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MPQuantidade.Text)
        If lErro <> SUCESSO Then gError 194643
    
        MPQuantidade.Text = Formata_Estoque(MPQuantidade.Text)
    
        GridMP.TextMatrix(GridMP.Row, iGrid_MPCustoT_Col) = Format(StrParaDbl(GridMP.TextMatrix(GridMP.Row, iGrid_MPCusto_Col)) * StrParaDbl(MPQuantidade.Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194644
    
    Call Soma_Coluna_Grid(objGridInt, iGrid_MPCustoT_Col, MPCustoTotal)
    
    Saida_Celula_MPQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MPQuantidade:

    Saida_Celula_MPQuantidade = gErr

    Select Case gErr
    
        Case 194643
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MPQuantidade.SetFocus
    
        Case 194644
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194645)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MPCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MPCusto

    Set objGridInt.objControle = MPCusto()
    
    'se a quantidade foi preenchida
    If Len(MPCusto().ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MPCusto().Text)
        If lErro <> SUCESSO Then gError 194646
    
        MPCusto().Text = Format(MPCusto().Text, "STANDARD")
    
        GridMP.TextMatrix(GridMP.Row, iGrid_MPCustoT_Col) = Format(StrParaDbl(GridMP.TextMatrix(GridMP.Row, iGrid_MPQuantidade_Col)) * StrParaDbl(MPCusto().Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194647

    Call Soma_Coluna_Grid(objGridInt, iGrid_MPCustoT_Col, MPCustoTotal)
    
    Saida_Celula_MPCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_MPCusto:

    Saida_Celula_MPCusto = gErr

    Select Case gErr
           
        Case 194646
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MPCusto.SetFocus

        Case 194647
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194648)

    End Select

    Exit Function

End Function

Private Sub MOCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOCodigo
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOQuantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOQuantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOQuantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MODescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MODescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MODescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MODescricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MODescricao
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOCusto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOCusto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOCusto
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOHoras_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOHoras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOHoras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOHoras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOHoras
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_MOCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra

On Error GoTo Erro_Saida_Celula_MOCodigo

    Set objGridInt.objControle = MOCodigo
    
    'Se o campo foi preenchido
    If Len(MOCodigo.Text) > 0 Then
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(MOCodigo.Text)
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 194649
    
        If lErro <> SUCESSO Then gError 194650

        GridMO.TextMatrix(GridMO.Row, iGrid_MODescricao_Col) = objTiposDeMaodeObra.sDescricao
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCusto_Col) = Format(objTiposDeMaodeObra.dCustoHora, "STANDARD")
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194651

    Saida_Celula_MOCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCodigo:

    Saida_Celula_MOCodigo = gErr

    Select Case gErr
    
        Case 194649, 1946501
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 194650
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194652)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOQuantidade

    Set objGridInt.objControle = MOQuantidade
    
    'se a quantidade foi preenchida
    If Len(MOQuantidade.ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(MOQuantidade.Text)
        If lErro <> SUCESSO Then gError 194653
        
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCustoT_Col) = Format(StrParaDbl(GridMO.TextMatrix(GridMO.Row, iGrid_MOCusto_Col)) * StrParaDbl(GridMO.TextMatrix(GridMO.Row, iGrid_MOHoras_Col)) * StrParaDbl(MOQuantidade.Text), "STANDARD")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194654

    Saida_Celula_MOQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MOQuantidade:

    Saida_Celula_MOQuantidade = gErr

    Select Case gErr
    
        Case 194653
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOQuantidade.SetFocus

        Case 194654
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194655)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOHoras(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOHoras

    Set objGridInt.objControle = MOHoras
    
    'se a quantidade foi preenchida
    If Len(MOHoras.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MOHoras.Text)
        If lErro <> SUCESSO Then gError 194656
    
        MOHoras.Text = Formata_Estoque(MOHoras.Text)
        
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCustoT_Col) = Format(StrParaDbl(GridMO.TextMatrix(GridMO.Row, iGrid_MOCusto_Col)) * StrParaDbl(GridMO.TextMatrix(GridMO.Row, iGrid_MOQuantidade_Col)) * StrParaDbl(MOHoras.Text), "STANDARD")
    
    End If
    
    Call Soma_Coluna_Grid(objGridInt, iGrid_MOCustoT_Col, MOCustoTotal)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194657

    Saida_Celula_MOHoras = SUCESSO

    Exit Function

Erro_Saida_Celula_MOHoras:

    Saida_Celula_MOHoras = gErr

    Select Case gErr
    
        Case 194656
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOHoras.SetFocus
        
        Case 194657
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194658)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MOCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MOCusto

    Set objGridInt.objControle = MOCusto()
    
    'se a quantidade foi preenchida
    If Len(MOCusto().ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MOCusto().Text)
        If lErro <> SUCESSO Then gError 194659
    
        MOCusto().Text = Format(MOCusto().Text, "STANDARD")
    
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCustoT_Col) = Format(StrParaDbl(GridMO.TextMatrix(GridMO.Row, iGrid_MOHoras_Col)) * StrParaDbl(GridMO.TextMatrix(GridMO.Row, iGrid_MOQuantidade_Col)) * StrParaDbl(MOCusto().Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194660

    Call Soma_Coluna_Grid(objGridInt, iGrid_MOCustoT_Col, MOCustoTotal)

    Saida_Celula_MOCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_MOCusto:

    Saida_Celula_MOCusto = gErr

    Select Case gErr
    
        Case 194659
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MOCusto.SetFocus

        Case 194660
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194661)

    End Select

    Exit Function

End Function

Private Sub MaqCodigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqCodigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqCodigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqCodigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqCodigo
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqQuantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqQuantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqQuantidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'Private Sub MaqDescricao_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Private Sub MaqDescricao_GotFocus()
'    Call Grid_Campo_Recebe_Foco(objGridMaq)
'End Sub
'
'Private Sub MaqDescricao_KeyPress(KeyAscii As Integer)
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
'End Sub
'
'Private Sub MaqDescricao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridMaq.objControle = MaqDescricao
'    lErro = Grid_Campo_Libera_Foco(objGridMaq)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub

Private Sub MaqCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqCusto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqCusto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqCusto
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MaqHoras_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqHoras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqHoras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqHoras_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqHoras
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_MaqCodigo(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas

On Error GoTo Erro_Saida_Celula_MaqCodigo

    Set objGridInt.objControle = MaqCodigo
    
    'Se o campo foi preenchido
    If Len(Trim(MaqCodigo.Text)) > 0 Then
    
        Set objMaquinas = New ClassMaquinas
    
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", MaqCodigo, objMaquinas)
        If lErro <> SUCESSO Then gError 194662
        
        'GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqDescricao_Col) = objMaquinas.sDescricao
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCusto_Col) = Format(objMaquinas.dCustoHora, "STANDARD")

        'verifica se precisa preencher o grid com uma nova linha
        If GridMaq.Row - GridMaq.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194663

    Saida_Celula_MaqCodigo = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqCodigo:

    Saida_Celula_MaqCodigo = gErr

    Select Case gErr
    
        Case 194662, 194663
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194664)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqQuantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqQuantidade

    Set objGridInt.objControle = MaqQuantidade
    
    'se a quantidade foi preenchida
    If Len(MaqQuantidade.ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(MaqQuantidade.Text)
        If lErro <> SUCESSO Then gError 194665
        
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCustoT_Col) = Format(StrParaDbl(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCusto_Col)) * StrParaDbl(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqHoras_Col)) * StrParaDbl(MaqQuantidade.Text), "STANDARD")
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194666

    Saida_Celula_MaqQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqQuantidade:

    Saida_Celula_MaqQuantidade = gErr

    Select Case gErr
           
        Case 194665
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MaqQuantidade.SetFocus

        Case 194666
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194667)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqHoras(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqHoras

    Set objGridInt.objControle = MaqHoras
    
    'se a quantidade foi preenchida
    If Len(MaqHoras.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MaqHoras.Text)
        If lErro <> SUCESSO Then gError 194668
    
        MaqHoras.Text = Formata_Estoque(MaqHoras.Text)
    
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCustoT_Col) = Format(StrParaDbl(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCusto_Col)) * StrParaDbl(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqQuantidade_Col)) * StrParaDbl(MaqHoras.Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194669

    Call Soma_Coluna_Grid(objGridInt, iGrid_MaqCustoT_Col, MaqCustoTotal)

    Saida_Celula_MaqHoras = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqHoras:

    Saida_Celula_MaqHoras = gErr

    Select Case gErr

        Case 194668
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MaqHoras.SetFocus

        Case 194669
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194670)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaqCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaqCusto

    Set objGridInt.objControle = MaqCusto()
    
    'se a quantidade foi preenchida
    If Len(MaqCusto().ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(MaqCusto().Text)
        If lErro <> SUCESSO Then gError 194671
    
        MaqCusto().Text = Format(MaqCusto().Text, "STANDARD")
    
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCustoT_Col) = Format(StrParaDbl(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqHoras_Col)) * StrParaDbl(GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqQuantidade_Col)) * StrParaDbl(MaqCusto().Text), "STANDARD")
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194672
    
    Call Soma_Coluna_Grid(objGridInt, iGrid_MaqCustoT_Col, MaqCustoTotal)

    Saida_Celula_MaqCusto = SUCESSO

    Exit Function

Erro_Saida_Celula_MaqCusto:

    Saida_Celula_MaqCusto = gErr

    Select Case gErr

        Case 194671
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            MaqCusto.SetFocus

        Case 194672
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194673)

    End Select

    Exit Function

End Function

Private Sub BotaoMO_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is MOCodigo Then
        objTiposDeMaodeObras.iCodigo = StrParaInt(MOCodigo)
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridMO.Row = 0 Then gError 194674

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col))
        
    End If

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoMO)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 194674
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194675)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer

On Error GoTo Erro_objEventoMO_evSelecao

    Set objTiposDeMaodeObra = obj1
    
    MOCodigo.Text = CStr(objTiposDeMaodeObra.iCodigo)
    
    If Not (Me.ActiveControl Is MOCodigo) Then
    
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCodigo_Col) = CStr(objTiposDeMaodeObra.iCodigo)
        GridMO.TextMatrix(GridMO.Row, iGrid_MODescricao_Col) = objTiposDeMaodeObra.sDescricao
        GridMO.TextMatrix(GridMO.Row, iGrid_MOCusto_Col) = objTiposDeMaodeObra.dCustoHora
    
        'verifica se precisa preencher o grid com uma nova linha
        If GridMO.Row - GridMO.FixedRows = objGridMO.iLinhasExistentes Then
            objGridMO.iLinhasExistentes = objGridMO.iLinhasExistentes + 1
        End If
        
    End If

    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194676)

    End Select

    Exit Sub

End Sub

Private Sub BotaoMaq_Click()

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaq_Click

    Set objMaquinas = New ClassMaquinas

    If Me.ActiveControl Is MaqCodigo Then
        objMaquinas.sNomeReduzido = MaqCodigo.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridMaq.Row = 0 Then gError 194677
        objMaquinas.sNomeReduzido = GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCodigo_Col)
    End If
    
    'Le a Máquina no BD a partir do NomeReduzido
    lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
    If lErro <> SUCESSO And lErro <> 103100 Then gError 194678
    
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaq, , "Nome Reduzido")

    Exit Sub

Erro_BotaoMaq_Click:

    Select Case gErr

        Case 194677
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 194678

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194679)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMaq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_objEventoMaq_evSelecao

    Set objMaquinas = obj1

    'Lê o Maquinas
    lErro = CF("TP_Maquina_Le", MaqCodigo, objMaquinas)
    If lErro <> SUCESSO Then gError 194680
    
    'Mostra os dados do Maquinas na tela
    MaqCodigo.Text = objMaquinas.sNomeReduzido
    
    If Not (Me.ActiveControl Is MaqCodigo) Then
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCodigo_Col) = objMaquinas.sNomeReduzido
        'GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqDescricao_Col) = objMaquinas.sDescricao
        GridMaq.TextMatrix(GridMaq.Row, iGrid_MaqCusto_Col) = objMaquinas.dCustoHora
    End If
    
    'verifica se precisa preencher o grid com uma nova linha
    If GridMaq.Row - GridMaq.FixedRows = objGridMaq.iLinhasExistentes Then
        objGridMaq.iLinhasExistentes = objGridMaq.iLinhasExistentes + 1
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaq_evSelecao:

    Select Case gErr

        Case 194680
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194681)

    End Select

    Exit Sub

End Sub

Function Soma_Coluna_Grid(ByVal objGrid As AdmGrid, ByVal iColuna As Integer, ByVal objControle As Object) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValor As Double

On Error GoTo Erro_Soma_Coluna_Grid
            
    For iIndice = 1 To objGrid.iLinhasExistentes
        dValor = dValor + StrParaDbl(objGrid.objGrid.TextMatrix(iIndice, iColuna))
    Next
    
    objControle.Caption = Format(dValor, "STANDARD")
        
    Soma_Coluna_Grid = SUCESSO

    Exit Function

Erro_Soma_Coluna_Grid:

    Soma_Coluna_Grid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194682)

    End Select

    Exit Function

End Function

Private Sub MaqOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MaqOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMaq)
End Sub

Private Sub MaqOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaq)
End Sub

Private Sub MaqOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaq.objControle = MaqOBS
    lErro = Grid_Campo_Libera_Foco(objGridMaq)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MPOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MPOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMP)
End Sub

Private Sub MPOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMP)
End Sub

Private Sub MPOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMP.objControle = MPOBS
    lErro = Grid_Campo_Libera_Foco(objGridMP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub MOOBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MOOBS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMO)
End Sub

Private Sub MOOBS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMO)
End Sub

Private Sub MOOBS_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMO.objControle = MOOBS
    lErro = Grid_Campo_Libera_Foco(objGridMO)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194683

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 194683
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194684)

    End Select

    Exit Function

End Function

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("ApontPRJ_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 194685
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 194685

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194686)
    
    End Select

    Exit Sub
    
End Sub

