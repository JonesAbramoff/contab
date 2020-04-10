VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl InfoAdicDocItemOcx 
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   LockControls    =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   10185
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5280
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   1410
      Visible         =   0   'False
      Width           =   9930
      Begin VB.Frame FrameExp 
         Caption         =   "Detalhamento de Exportação"
         Height          =   4020
         Left            =   270
         TabIndex        =   27
         Top             =   1215
         Width           =   9240
         Begin VB.Frame FrameDE 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Enabled         =   0   'False
            Height          =   360
            Left            =   60
            TabIndex        =   41
            Top             =   195
            Width           =   7725
            Begin VB.ComboBox NumRE 
               Height          =   315
               Left            =   6030
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   0
               Width           =   1710
            End
            Begin MSMask.MaskEdBox NumDE 
               Height          =   315
               Left            =   2370
               TabIndex        =   14
               Tag             =   "Número da declaração de exportação padrão caso a informação não seja preenchida no item"
               Top             =   30
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   11
               Mask            =   "###########"
               PromptChar      =   " "
            End
            Begin VB.Label Label1Ret 
               AutoSize        =   -1  'True
               Caption         =   "Registro de Exportação:"
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
               Index           =   12
               Left            =   3915
               TabIndex        =   43
               Top             =   60
               Width           =   2055
            End
            Begin VB.Label LabelDE 
               AutoSize        =   -1  'True
               Caption         =   "Declaração de Exportação:"
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
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   42
               Top             =   60
               Width           =   2325
            End
         End
         Begin VB.CheckBox optDEPadrao 
            Caption         =   "Usar Padrão da NF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7815
            TabIndex        =   16
            Top             =   150
            Value           =   1  'Checked
            Width           =   1380
         End
         Begin MSMask.MaskEdBox QuantExport 
            Height          =   255
            Left            =   7680
            TabIndex        =   26
            Top             =   2310
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ChvNFe 
            Height          =   255
            Left            =   2775
            TabIndex        =   30
            Top             =   1395
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   54
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#### #### #### #### #### #### #### #### #### #### ####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumRegistExport 
            Height          =   255
            Left            =   2670
            TabIndex        =   29
            Top             =   2445
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumDrawback 
            Height          =   255
            Left            =   300
            TabIndex        =   28
            Top             =   2460
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridExp 
            Height          =   2280
            Left            =   75
            TabIndex        =   17
            Top             =   570
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   4022
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label TotalExport 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7005
            TabIndex        =   31
            Top             =   3645
            Width           =   930
         End
      End
      Begin VB.Frame FrameCompra 
         Caption         =   "Compra"
         Height          =   630
         Left            =   255
         TabIndex        =   23
         Top             =   0
         Width           =   5460
         Begin VB.TextBox Pedido 
            Height          =   315
            Left            =   1185
            MaxLength       =   15
            TabIndex        =   10
            Top             =   225
            Width           =   1755
         End
         Begin MSMask.MaskEdBox ItemPedido 
            Height          =   300
            Left            =   3945
            TabIndex        =   11
            Top             =   240
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pedido:"
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
            Index           =   9
            Left            =   450
            TabIndex        =   25
            Top             =   270
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Item:"
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
            Index           =   0
            Left            =   3465
            TabIndex        =   24
            Top             =   270
            Width           =   435
         End
      End
      Begin VB.Frame FrameFat 
         Caption         =   "Faturamento"
         Height          =   645
         Left            =   6165
         TabIndex        =   20
         Top             =   -15
         Width           =   3345
         Begin MSComCtl2.UpDown UpDownDataLimFat 
            Height          =   300
            Left            =   2355
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataLimFat 
            Height          =   300
            Left            =   1275
            TabIndex        =   12
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Limite:"
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
            Left            =   600
            TabIndex        =   22
            Top             =   285
            Width           =   570
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Configurações"
         Height          =   585
         Left            =   255
         TabIndex        =   9
         Top             =   630
         Width           =   9255
         Begin VB.CheckBox IncluiValorTotal 
            Caption         =   "O valor do item compõem o valor total de produtos e serviços"
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
            Left            =   1185
            TabIndex        =   13
            Top             =   255
            Value           =   1  'Checked
            Width           =   5595
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5280
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   1410
      Width           =   9930
      Begin VB.Frame Frame3 
         Caption         =   "Texto Adicional"
         Height          =   4935
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   9810
         Begin VB.TextBox TextoAd 
            Height          =   4425
            Left            =   225
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   345
            Width           =   9360
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Identificação"
      Height          =   945
      Left            =   60
      TabIndex        =   2
      Top             =   75
      Width           =   8160
      Begin VB.TextBox Produto 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   210
         Width           =   5670
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6945
         TabIndex        =   40
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Index           =   7
         Left            =   6360
         TabIndex        =   39
         Top             =   600
         Width           =   510
      End
      Begin VB.Label PrecoUnitario 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4965
         TabIndex        =   38
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Preço Unitário:"
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
         Left            =   3645
         TabIndex        =   37
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Qtde 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2415
         TabIndex        =   36
         Top             =   570
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
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
         Index           =   4
         Left            =   1320
         TabIndex        =   35
         Top             =   615
         Width           =   1050
      End
      Begin VB.Label UM 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   735
         TabIndex        =   34
         Top             =   570
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "UM:"
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
         Index           =   1
         Left            =   330
         TabIndex        =   32
         Top             =   630
         Width           =   360
      End
      Begin VB.Label Item 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   735
         TabIndex        =   5
         Top             =   210
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
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
         Index           =   3
         Left            =   1620
         TabIndex        =   4
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
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
         TabIndex        =   3
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      Picture         =   "InfoAdicDocItemOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   885
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9210
      Picture         =   "InfoAdicDocItemOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   165
      Width           =   885
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5670
      Left            =   60
      TabIndex        =   6
      Top             =   1065
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   10001
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mensagem"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outros"
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
Attribute VB_Name = "InfoAdicDocItemOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim objGridExp As AdmGrid
Dim iGrid_NumDrawback_Col As Integer
Dim iGrid_ChvNFe_Col As Integer
Dim iGrid_NumRegistExport_Col As Integer
Dim iGrid_QuantExport_Col As Integer

Private sNumDEAnt As String

Private gobjInfoAdicDocItem As ClassInfoAdicDocItem
Private gobjInfoAdicDoc As ClassInfoAdic
Private gobjTela As Object
Private gsTipoTela As String

Dim iFrameAtual As Integer

Private WithEvents objEventoDE As AdmEvento
Attribute objEventoDE.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Informações Adicionais do Item"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "InfoAdicDocItem"
    
End Function

Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
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

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    lErro = Move_Tela_Memoria()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iAlterado = REGISTRO_ALTERADO Then
        gobjTela.iAlterado = iAlterado
    End If
    
    Unload Me
    
    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211246)

    End Select

    Exit Sub
    
End Sub

Private Sub Form_Unload()
    Set gobjInfoAdicDocItem = Nothing
    Set gobjInfoAdicDoc = Nothing
    Set objEventoDE = Nothing
    Set gobjTela = Nothing
End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
 
    iAlterado = 0
    iFrameAtual = 1
    
    Set objGridExp = New AdmGrid
    Set objEventoDE = New AdmEvento
    
    lErro = Inicializa_GridExp(objGridExp)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211247)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(ByVal objInfoAdicDocItem As ClassInfoAdicDocItem, ByVal objTela As Object, Optional ByVal sTipoTela As String = TIPO_SAIDA, Optional ByVal objInfoAdicDoc As ClassInfoAdic) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjInfoAdicDocItem = objInfoAdicDocItem
    Set gobjInfoAdicDoc = objInfoAdicDoc
    Set gobjTela = objTela
    gsTipoTela = sTipoTela
    
    If gsTipoTela = TIPO_ENTRADA Then
        FrameFat.Enabled = False
    End If
    
    Item.Caption = CStr(objInfoAdicDocItem.iItem)
    Produto.Text = objInfoAdicDocItem.sProduto & SEPARADOR & objInfoAdicDocItem.sDescProd
    
    Qtde.Caption = Formata_Estoque(objInfoAdicDocItem.dQuantidade)
    PrecoUnitario.Caption = Format(objInfoAdicDocItem.dPrecoUnitario, "STANDARD")
    ValorTotal.Caption = Format(objInfoAdicDocItem.dValorTotal, "STANDARD")
    UM.Caption = objInfoAdicDocItem.sUM
    
    lErro = Traz_InfoAdic_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211248)
    
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria() As Long

Dim lErro As Long, iLinha As Integer
Dim objDetExport As ClassInfoAdicDocItemDetExp
Dim objDE As New ClassDEInfo

On Error GoTo Erro_Move_Tela_Memoria

    If StrParaDbl(Qtde.Caption) - StrParaDbl(TotalExport.Caption) < -QTDE_ESTOQUE_DELTA Then gError 213674

    gobjInfoAdicDocItem.dtDataLimiteFaturamento = StrParaDate(DataLimFat.Text)
    gobjInfoAdicDocItem.iIncluiValorTotal = IIf(IncluiValorTotal = vbChecked, MARCADO, DESMARCADO)
    gobjInfoAdicDocItem.lItemPedCompra = StrParaLong(ItemPedido.Text)
    gobjInfoAdicDocItem.sMsg = TextoAd.Text
    gobjInfoAdicDocItem.sNumPedidoCompra = Pedido.Text
    
    If optDEPadrao.Value = vbUnchecked Then
    
        objDE.sNumero = Trim(NumDE.ClipText)
        
        If Len(Trim(objDE.sNumero)) > 0 Then
        
            lErro = CF("DEInfo_Le", objDE)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            gobjInfoAdicDocItem.lNumIntDE = objDE.lNumIntDoc
            
            gobjInfoAdicDocItem.sNumRE = Trim(NumRE.Text)
        
        End If
    Else
        gobjInfoAdicDocItem.lNumIntDE = 0
        gobjInfoAdicDocItem.sNumRE = ""
    End If
    
    Set gobjInfoAdicDocItem.colDetExportacao = New Collection
    
    For iLinha = 1 To objGridExp.iLinhasExistentes
        
        Set objDetExport = New ClassInfoAdicDocItemDetExp
        
        objDetExport.sChvNFe = Replace(GridExp.TextMatrix(iLinha, iGrid_ChvNFe_Col), " ", "")
        objDetExport.sNumDrawback = GridExp.TextMatrix(iLinha, iGrid_NumDrawback_Col)
        objDetExport.sNumRegistExport = GridExp.TextMatrix(iLinha, iGrid_NumRegistExport_Col)
        objDetExport.dQuantExport = StrParaDbl(GridExp.TextMatrix(iLinha, iGrid_QuantExport_Col))

        gobjInfoAdicDocItem.colDetExportacao.Add objDetExport
    
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 213674
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_EXPORT_MAIOR_QTDE_ITEM", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211249)

    End Select

    Exit Function
    
End Function

Private Function Traz_InfoAdic_Tela() As Long

Dim lErro As Long, iLinha As Integer
Dim objDetExport As ClassInfoAdicDocItemDetExp
Dim objDE As New ClassDEInfo, sNumRE As String

On Error GoTo Erro_Traz_InfoAdic_Tela

    DataLimFat.PromptInclude = False
    If gobjInfoAdicDocItem.dtDataLimiteFaturamento <> DATA_NULA Then
        DataLimFat.Text = Format(gobjInfoAdicDocItem.dtDataLimiteFaturamento, "dd/mm/yy")
    Else
        DataLimFat.Text = ""
    End If
    DataLimFat.PromptInclude = True
    
    If gobjInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
        IncluiValorTotal.Value = vbChecked
    Else
        IncluiValorTotal.Value = vbUnchecked
    End If
    
    ItemPedido.PromptInclude = False
    If gobjInfoAdicDocItem.lItemPedCompra <> 0 Then
        ItemPedido.Text = CStr(gobjInfoAdicDocItem.lItemPedCompra)
    Else
        ItemPedido.Text = ""
    End If
    ItemPedido.PromptInclude = True
    TextoAd.Text = gobjInfoAdicDocItem.sMsg
    Pedido.Text = gobjInfoAdicDocItem.sNumPedidoCompra
    
    If gobjInfoAdicDocItem.lNumIntDE <> 0 Then
        optDEPadrao.Value = vbUnchecked
        objDE.lNumIntDoc = gobjInfoAdicDocItem.lNumIntDE
        sNumRE = gobjInfoAdicDocItem.sNumRE
    Else
        optDEPadrao.Value = vbChecked
        If Not (gobjInfoAdicDoc Is Nothing) Then
            If Not (gobjInfoAdicDoc.objExportacao Is Nothing) Then
                objDE.lNumIntDoc = gobjInfoAdicDoc.objExportacao.lNumIntDE
                sNumRE = gobjInfoAdicDoc.objExportacao.sNumRE
            End If
        End If
    End If
    Call Trata_DE_Padrao
    
    If objDE.lNumIntDoc > 0 Then
    
        lErro = CF("DEInfo_Le", objDE)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            NumDE.PromptInclude = False
            NumDE.Text = objDE.sNumero
            NumDE.PromptInclude = True
            Call NumDE_Validate(bSGECancelDummy)
            
            Call CF("SCombo_Seleciona2", NumRE, sNumRE)
        End If
        
    End If
    
    iLinha = 0
    For Each objDetExport In gobjInfoAdicDocItem.colDetExportacao
        iLinha = iLinha + 1
        
        ChvNFe.PromptInclude = False
        ChvNFe.Text = objDetExport.sChvNFe
        ChvNFe.PromptInclude = True
        
        GridExp.TextMatrix(iLinha, iGrid_ChvNFe_Col) = ChvNFe.Text
        GridExp.TextMatrix(iLinha, iGrid_NumDrawback_Col) = objDetExport.sNumDrawback
        GridExp.TextMatrix(iLinha, iGrid_NumRegistExport_Col) = objDetExport.sNumRegistExport
        GridExp.TextMatrix(iLinha, iGrid_QuantExport_Col) = Formata_Estoque(objDetExport.dQuantExport)
    
        ChvNFe.PromptInclude = False
        ChvNFe.Text = ""
        ChvNFe.PromptInclude = True
    
    Next
    objGridExp.iLinhasExistentes = iLinha
    
    Call Soma_Coluna_Grid(objGridExp, iGrid_QuantExport_Col, TotalExport, True)
    
    Traz_InfoAdic_Tela = SUCESSO

    Exit Function

Erro_Traz_InfoAdic_Tela:

    Traz_InfoAdic_Tela = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211250)

    End Select

    Exit Function
    
End Function

Private Sub DataLimFat_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataLimFat, iAlterado)
End Sub

Private Sub DataLimFat_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataLimFat_Validate

    If Len(Trim(DataLimFat.ClipText)) <> 0 Then

        lErro = Data_Critica(DataLimFat.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Exit Sub

Erro_DataLimFat_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211251)

    End Select

    Exit Sub

End Sub

Private Sub DataLimFat_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataLimFat_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataLimFat_DownClick

    DataLimFat.SetFocus

    If Len(DataLimFat.ClipText) > 0 Then

        sData = DataLimFat.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataLimFat.Text = sData
        
        Call DataLimFat_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataLimFat_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211252)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataLimFat_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataLimFat_UpClick

    DataLimFat.SetFocus

    If Len(Trim(DataLimFat.ClipText)) > 0 Then

        sData = DataLimFat.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataLimFat.Text = sData
        
        Call DataLimFat_Validate(bSGECancelDummy)

    End If

    Exit Sub

Erro_UpDownDataLimFat_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211253)

    End Select

    Exit Sub

End Sub

Private Sub TextoAd_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemPedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Pedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ItemPedido_GotFocus()
    Call MaskEdBox_TrataGotFocus(ItemPedido, iAlterado)
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If
    
End Sub

Public Sub GridExp_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridExp, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExp, iAlterado)
    End If
    
End Sub

Public Sub GridExp_EnterCell()
    Call Grid_Entrada_Celula(objGridExp, iAlterado)
End Sub

Public Sub GridExp_GotFocus()
    Call Grid_Recebe_Foco(objGridExp)
End Sub

Public Sub GridExp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridExp)
    
End Sub

Public Sub GridExp_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridExp, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExp, iAlterado)
    End If
    
End Sub

Public Sub GridExp_LeaveCell()
    Call Saida_Celula(objGridExp)
End Sub

Public Sub GridExp_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridExp)
End Sub

Public Sub GridExp_RowColChange()
    Call Grid_RowColChange(objGridExp)
End Sub

Public Sub GridExp_Scroll()
    Call Grid_Scroll(objGridExp)
End Sub

Public Sub NumDrawback_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumDrawback_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExp)
End Sub

Public Sub NumDrawback_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExp)
End Sub

Public Sub NumDrawback_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExp.objControle = NumDrawback
    lErro = Grid_Campo_Libera_Foco(objGridExp)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub NumRegistExport_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumRegistExport_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExp)
End Sub

Public Sub NumRegistExport_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExp)
End Sub

Public Sub NumRegistExport_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExp.objControle = NumRegistExport
    lErro = Grid_Campo_Libera_Foco(objGridExp)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ChvNFe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ChvNFe_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExp)
End Sub

Public Sub ChvNFe_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExp)
End Sub

Public Sub ChvNFe_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExp.objControle = ChvNFe
    lErro = Grid_Campo_Libera_Foco(objGridExp)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub QuantExport_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub QuantExport_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExp)
End Sub

Public Sub QuantExport_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExp)
End Sub

Public Sub QuantExport_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridExp.objControle = QuantExport
    lErro = Grid_Campo_Libera_Foco(objGridExp)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col
           
            Case iGrid_NumDrawback_Col
                lErro = Saida_Celula_Padrao(objGridInt, NumDrawback, True)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case iGrid_NumRegistExport_Col
                lErro = Saida_Celula_Padrao(objGridInt, NumRegistExport, True)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            Case iGrid_ChvNFe_Col
                lErro = Saida_Celula_ChvNFe(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case iGrid_QuantExport_Col
                lErro = Saida_Celula_Valor(objGridInt, QuantExport)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                Call Soma_Coluna_Grid(objGridExp, iGrid_QuantExport_Col, TotalExport, True)

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 213678

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 213678
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213679)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridExp(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Núm.Drawback")
    objGrid.colColuna.Add ("Reg.Export.Ind.")
    objGrid.colColuna.Add ("Chave da NF-e recebida para exportação")
    objGrid.colColuna.Add ("Qtde Exp.")

    'Controles que participam do Grid
    objGrid.colCampo.Add (NumDrawback.Name)
    objGrid.colCampo.Add (NumRegistExport.Name)
    objGrid.colCampo.Add (ChvNFe.Name)
    objGrid.colCampo.Add (QuantExport.Name)

    'Colunas do Grid
    iGrid_NumDrawback_Col = 1
    iGrid_NumRegistExport_Col = 2
    iGrid_ChvNFe_Col = 3
    iGrid_QuantExport_Col = 4

    objGrid.objGrid = GridExp

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 200 + 1

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridExp.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)
    
    Call Soma_Coluna_Grid(objGridExp, iGrid_QuantExport_Col, TotalExport, True)

    Inicializa_GridExp = SUCESSO

End Function

Private Function Saida_Celula_ChvNFe(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sChvNFe As String
Dim sUF As String, sAno As String, sMes As String, sCNPJEmi As String, sModelo As String
Dim sSerie As String, sNumero As String, sTipoEmissao As String, sCodNumerico As String
Dim sDV As String, sDVCalc As String

On Error GoTo Erro_Saida_Celula_ChvNFe

    Set objGridInt.objControle = ChvNFe
    
    sChvNFe = ChvNFe.ClipText
    
    If Len(Trim(sChvNFe)) <> 0 Then
        
        If Len(Trim(sChvNFe)) <> STRING_NFE_CHNFE Then gError 213680
            
        sUF = Mid(sChvNFe, 1, 2)
        sAno = Mid(sChvNFe, 3, 2)
        sMes = Mid(sChvNFe, 5, 2)
        sCNPJEmi = Mid(sChvNFe, 7, 14)
        sModelo = Mid(sChvNFe, 21, 2)
        sSerie = Mid(sChvNFe, 23, 3)
        sNumero = Mid(sChvNFe, 26, 9)
        sTipoEmissao = Mid(sChvNFe, 35, 1)
        sCodNumerico = Mid(sChvNFe, 36, 8)
        sDV = Mid(sChvNFe, 44, 1)
        
        lErro = CF("Calcula_DV11", left(sChvNFe, 43), 9, sDVCalc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If sDV <> sDVCalc Then gError 213682 'Chave inválida
        
        lErro = Cgc_Critica(sCNPJEmi)
        If lErro <> SUCESSO Then gError 213683 'Chave inválida
        
        'GridExp.TextMatrix(GridExp.Row, iGrid_ChvNFe_Col) = ChvNFe.Text
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ChvNFe = SUCESSO

    Exit Function

Erro_Saida_Celula_ChvNFe:

    Saida_Celula_ChvNFe = gErr

    Select Case gErr

        Case 213680
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHV_TAM_INVALIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 213682, 213683
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHV_INVALIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213681)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function


Public Sub NumDE_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumDE, iAlterado)
End Sub

Public Sub NumDE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumRE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub LabelDE_Click()

Dim objDE As New ClassDEInfo
Dim colSelecao As Collection

    objDE.sNumero = NumDE.Text

    'Chama a Tela de PaisesLista
    Call Chama_Tela_Modal("DEInfoLista", colSelecao, objDE, objEventoDE)

End Sub

Private Sub objEventoDE_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDE As ClassDEInfo

On Error GoTo Erro_objEventoDE_evSelecao

    Set objDE = obj1

    NumDE.PromptInclude = False
    NumDE.Text = objDE.sNumero
    NumDE.PromptInclude = True
    Call NumDE_Validate(bSGECancelDummy)
    
    Call CF("SCombo_Seleciona2", NumRE, objDE.sNumRegistro)

    Me.Show

    Exit Sub

Erro_objEventoDE_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216037)

    End Select

    Exit Sub

End Sub

Private Sub NumDE_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDE As New ClassDEInfo
Dim objRE As ClassDERegistro

On Error GoTo Erro_NumDE_Validate

    If sNumDEAnt <> NumDE.ClipText Then

        If Len(NumDE.ClipText) > 0 Then
    
            objDE.sNumero = NumDE.ClipText
            
            lErro = CF("DEInfo_Le", objDE)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216036
            
            NumRE.Clear
            NumRE.AddItem ""
            For Each objRE In objDE.colRE
                NumRE.AddItem objRE.sNumRegistro
            Next
            
            If objDE.colRE.Count = 1 Then NumRE.ListIndex = 1
        
        Else
        
            NumRE.Clear
        
        End If
        
        sNumDEAnt = NumDE.ClipText
        
    End If
    
    Exit Sub
    
Erro_NumDE_Validate:

    Cancel = True

    Select Case gErr
    
        Case 216036
            Call Rotina_Erro(vbOKOnly, "ERRO_DEINFO_NAO_CADASTRADO", gErr, objDE.sNumero)
    
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216038)
        
    End Select

    Exit Sub
    
End Sub

Private Sub optDEPadrao_Click()
    Call Trata_DE_Padrao
End Sub

Sub Trata_DE_Padrao()

Dim lErro As Long
Dim objDE As New ClassDEInfo, sNumRE As String

On Error GoTo Erro_Trata_DE_Padrao

    If optDEPadrao.Value = vbChecked Then
    
        FrameDE.Enabled = False
        If Not (gobjInfoAdicDoc Is Nothing) Then
            If Not (gobjInfoAdicDoc.objExportacao Is Nothing) Then
                objDE.lNumIntDoc = gobjInfoAdicDoc.objExportacao.lNumIntDE
                sNumRE = gobjInfoAdicDoc.objExportacao.sNumRE
            End If
        End If
        
        If objDE.lNumIntDoc > 0 Then
        
            lErro = CF("DEInfo_Le", objDE)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
                NumDE.PromptInclude = False
                NumDE.Text = objDE.sNumero
                NumDE.PromptInclude = True
                Call NumDE_Validate(bSGECancelDummy)
                
                Call CF("SCombo_Seleciona2", NumRE, sNumRE)
            End If
            
        Else
            NumDE.PromptInclude = False
            NumDE.Text = ""
            NumDE.PromptInclude = True
            Call NumDE_Validate(bSGECancelDummy)
        End If
    Else
        FrameDE.Enabled = True
    End If
    
    Exit Sub
    
Erro_Trata_DE_Padrao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216039)
        
    End Select

    Exit Sub
End Sub
