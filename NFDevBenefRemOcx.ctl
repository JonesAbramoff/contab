VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl NFDevBenefRemOcx 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   8205
   Begin VB.Frame Frame1 
      Caption         =   "Ordens de Produção utilizadas no cálculo"
      Height          =   3570
      Index           =   2
      Left            =   165
      TabIndex        =   0
      Top             =   1230
      Visible         =   0   'False
      Width           =   7770
      Begin MSMask.MaskEdBox OPQuantDev 
         Height          =   225
         Left            =   4800
         TabIndex        =   1
         Top             =   2070
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox OPQuantADevolver 
         Height          =   225
         Left            =   4095
         TabIndex        =   2
         Top             =   2595
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox OPConsumo 
         Height          =   225
         Left            =   2910
         TabIndex        =   3
         Top             =   2520
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox ProdutoOP 
         Height          =   225
         Left            =   1200
         TabIndex        =   4
         Top             =   2685
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox OP 
         Height          =   225
         Left            =   360
         TabIndex        =   5
         Top             =   2070
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid GridOP 
         Height          =   1725
         Left            =   75
         TabIndex        =   6
         Top             =   210
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   3043
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
      Height          =   525
      Left            =   3720
      Picture         =   "NFDevBenefRemOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4995
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notas Fiscais de Remessa para Beneficiamento devolvidas"
      Height          =   3570
      Index           =   1
      Left            =   165
      TabIndex        =   7
      Top             =   1230
      Width           =   7770
      Begin MSMask.MaskEdBox NFQuantDev 
         Height          =   225
         Left            =   5925
         TabIndex        =   8
         Top             =   2010
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox NFSaldoADevolver 
         Height          =   225
         Left            =   4080
         TabIndex        =   9
         Top             =   2070
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox NFQuantidade 
         Height          =   225
         Left            =   2700
         TabIndex        =   10
         Top             =   2070
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Item 
         Height          =   225
         Left            =   3300
         TabIndex        =   11
         Top             =   2460
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox NumNF 
         Height          =   225
         Left            =   1110
         TabIndex        =   12
         Top             =   2070
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox Serie 
         Height          =   225
         Left            =   360
         TabIndex        =   13
         Top             =   2055
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid GridRem 
         Height          =   1770
         Left            =   75
         TabIndex        =   14
         Top             =   210
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   225
         Left            =   1950
         TabIndex        =   15
         Top             =   2415
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4050
      Left            =   45
      TabIndex        =   16
      Top             =   885
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7144
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Remessa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produção"
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
   Begin VB.Label Label4 
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
      Height          =   210
      Left            =   6960
      TabIndex        =   25
      Top             =   540
      Width           =   420
   End
   Begin VB.Label UM 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   7380
      TabIndex        =   24
      Top             =   495
      Width           =   675
   End
   Begin VB.Label Label2 
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
      Height          =   210
      Left            =   210
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Produto 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1050
      TabIndex        =   22
      Top             =   90
      Width           =   7005
   End
   Begin VB.Label QtdeTotalDev 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2175
      TabIndex        =   21
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label3 
      Caption         =   "Qtde Total a Devolver:"
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
      Left            =   165
      TabIndex        =   20
      Top             =   525
      Width           =   2175
   End
   Begin VB.Label QtdeDev 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5820
      TabIndex        =   19
      Top             =   495
      Width           =   1050
   End
   Begin VB.Label Label5 
      Caption         =   "Qtde Total sendo Devolvida:"
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
      Left            =   3330
      TabIndex        =   18
      Top             =   540
      Width           =   2640
   End
End
Attribute VB_Name = "NFDevBenefRemOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjItemNF As ClassItemNF
Dim iFrameAtual As Integer

Dim objGridOP As AdmGrid
Dim iGrid_OP_Col As Integer
Dim iGrid_ProdutoOP_Col As Integer
Dim iGrid_OPConsumo_Col As Integer
Dim iGrid_OPQuantDev_Col As Integer
Dim iGrid_OPQuantADevolver_Col As Integer

Dim objGridRem As AdmGrid
Dim iGrid_Serie_Col As Integer
Dim iGrid_NumNF_Col As Integer
Dim iGrid_DataEmissao_Col As Integer
Dim iGrid_Item_Col As Integer
Dim iGrid_NFQuantidade_Col As Integer
Dim iGrid_NFSaldoADevolver_Col As Integer
Dim iGrid_NFQuantDev_Col As Integer

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Detalhamento da devolução"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFDevBenefRemOcx"
    
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

Public Function Trata_Parametros(objItemNF As ClassItemNF) As Long

Dim lErro As Long
Dim objCli As New ClassCliente
Dim objFilCli As New ClassFilialCliente

On Error GoTo Erro_Trata_Parametros

    Set gobjItemNF = objItemNF
    
    Set objGridOP = New AdmGrid
    Set objGridRem = New AdmGrid
    
    iFrameAtual = 1

    'inicializacao do grid
    Call Inicializa_Grid_OP(objGridOP)
    Call Inicializa_Grid_Rem(objGridRem)
        
    lErro = Traz_Dados_Tela(objItemNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
         
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209983)

    End Select

    Exit Function
    
End Function

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209984)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_OP(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_OP

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Ordem de Produção")
    objGridInt.colColuna.Add ("Produto OP")
    objGridInt.colColuna.Add ("Consumido")
    objGridInt.colColuna.Add ("Devolvido")
    objGridInt.colColuna.Add ("Devolver")

    'campos de edição do grid
    objGridInt.colCampo.Add (OP.Name)
    objGridInt.colCampo.Add (ProdutoOP.Name)
    objGridInt.colCampo.Add (OPConsumo.Name)
    objGridInt.colCampo.Add (OPQuantDev.Name)
    objGridInt.colCampo.Add (OPQuantADevolver.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_OP_Col = 1
    iGrid_ProdutoOP_Col = 2
    iGrid_OPConsumo_Col = 3
    iGrid_OPQuantDev_Col = 4
    iGrid_OPQuantADevolver_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridOP

    'Largura da primeira coluna
    GridOP.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 11
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_OP = SUCESSO

    Exit Function

Erro_Inicializa_Grid_OP:

    Inicializa_Grid_OP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209985)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Rem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Rem

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Núm.NF")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Total")
    objGridInt.colColuna.Add ("Saldo Atual")
    objGridInt.colColuna.Add ("Devolver")

    'campos de edição do grid
    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (NumNF.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (Item.Name)
    objGridInt.colCampo.Add (NFQuantidade.Name)
    objGridInt.colCampo.Add (NFSaldoADevolver.Name)
    objGridInt.colCampo.Add (NFQuantDev.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Serie_Col = 1
    iGrid_NumNF_Col = 2
    iGrid_DataEmissao_Col = 3
    iGrid_Item_Col = 4
    iGrid_NFQuantidade_Col = 5
    iGrid_NFSaldoADevolver_Col = 6
    iGrid_NFQuantDev_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridRem

    'Largura da primeira coluna
    GridRem.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 11
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Rem = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Rem:

    Inicializa_Grid_Rem = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209986)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set gobjItemNF = Nothing
    Set objGridOP = Nothing
    Set objGridRem = Nothing
End Sub

Private Sub GridOP_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOP, iAlterado)
    End If

End Sub

Private Sub GridOP_GotFocus()
    Call Grid_Recebe_Foco(objGridOP)
End Sub

Private Sub GridOP_EnterCell()
    Call Grid_Entrada_Celula(objGridOP, iAlterado)
End Sub

Private Sub GridOP_LeaveCell()
    Call Saida_Celula(objGridOP)
End Sub

Private Sub GridOP_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOP, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOP, iAlterado)
    End If

End Sub

Private Sub GridOP_RowColChange()
    Call Grid_RowColChange(objGridOP)
End Sub

Private Sub GridOP_Scroll()
    Call Grid_Scroll(objGridOP)
End Sub

Private Sub GridOP_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridOP)
End Sub

Private Sub GridOP_LostFocus()
    Call Grid_Libera_Foco(objGridOP)
End Sub

Private Sub GridRem_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRem, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRem, iAlterado)
    End If

End Sub

Private Sub GridRem_GotFocus()
    Call Grid_Recebe_Foco(objGridRem)
End Sub

Private Sub GridRem_EnterCell()
    Call Grid_Entrada_Celula(objGridRem, iAlterado)
End Sub

Private Sub GridRem_LeaveCell()
    Call Saida_Celula(objGridRem)
End Sub

Private Sub GridRem_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRem, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRem, iAlterado)
    End If

End Sub

Private Sub GridRem_RowColChange()
    Call Grid_RowColChange(objGridRem)
End Sub

Private Sub GridRem_Scroll()
    Call Grid_Scroll(objGridRem)
End Sub

Private Sub GridRem_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridRem)
End Sub

Private Sub GridRem_LostFocus()
    Call Grid_Libera_Foco(objGridRem)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 209987

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            'erros tratatos nas rotinas chamadas

        Case 209987
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209988)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Else
            objControl.Enabled = False

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209989)

    End Select

    Exit Sub

End Sub

Private Sub BotaoOK_Click()
    Unload Me
End Sub

Private Function Traz_Dados_Tela(ByVal objItemNF As ClassItemNF) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objItemOPDev As ClassNFDevBenefInsumo
Dim objItemNFRem As ClassItemNF
Dim objItemNFRemBD As ClassItemNF
Dim objNFRemBD As ClassNFiscal
Dim dFatorUM As Double, objProduto As ClassProduto
Dim dSaldoADevolver As Double, dQtdeDev As Double, dQtdeTotalDev As Double
Dim sProdMask As String, iIndice As Integer

On Error GoTo Erro_Traz_Dados_Tela

    Call Grid_Limpa(objGridOP)
    Call Grid_Limpa(objGridRem)
    
    Set objProduto = New ClassProduto
    objProduto.sCodigo = objItemNF.sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    Produto.Caption = objProduto.sCodigo & SEPARADOR & objProduto.sDescricao
    
    iIndice = 0
    For Each objItemNFRem In objItemNF.colItensNFDevBenefRem
        iIndice = iIndice + 1
        Set objItemNFRemBD = New ClassItemNF
        Set objNFRemBD = New ClassNFiscal
        objItemNFRemBD.lNumIntDoc = objItemNFRem.lNumIntDoc
        
        lErro = CF("ItemNFiscal_Le", objItemNFRemBD)
        If lErro <> SUCESSO And lErro <> 35225 Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNFRemBD.sUnidadeMed, objItemNF.sUnidadeMed, dFatorUM)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objNFRemBD.lNumIntDoc = objItemNFRemBD.lNumIntNF
        
        lErro = CF("NFiscal_Le", objNFRemBD)
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("ItemNFiscalRem_Le_SaldoADevolver", objItemNFRemBD.lNumIntDoc, dSaldoADevolver)
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
        
        GridRem.TextMatrix(iIndice, iGrid_Serie_Col) = objNFRemBD.sSerie
        GridRem.TextMatrix(iIndice, iGrid_NumNF_Col) = CStr(objNFRemBD.lNumNotaFiscal)
        GridRem.TextMatrix(iIndice, iGrid_DataEmissao_Col) = Format(objNFRemBD.dtDataEmissao, "dd/mm/yyyy")
        GridRem.TextMatrix(iIndice, iGrid_Item_Col) = objItemNFRemBD.iItem
        
        'Está tudo na UM de remessa, tem que colocar na UM da devolução
        GridRem.TextMatrix(iIndice, iGrid_NFQuantidade_Col) = Formata_Estoque(objItemNFRemBD.dQuantidade * dFatorUM)
        GridRem.TextMatrix(iIndice, iGrid_NFSaldoADevolver_Col) = Formata_Estoque(dSaldoADevolver * dFatorUM)
        GridRem.TextMatrix(iIndice, iGrid_NFQuantDev_Col) = Formata_Estoque(objItemNFRem.dQuantidade * dFatorUM)
    Next
    objGridRem.iLinhasExistentes = iIndice
    
    iIndice = 0
    For Each objItemOPDev In objItemNF.colItensNFDevBenefItemOP
        iIndice = iIndice + 1
        Set objItemOP = New ClassItemOP
        objItemOP.lNumIntDoc = objItemOPDev.lNumIntItemOP
        
        lErro = CF("ItemOP_Le_NumIntDoc", objItemOP, True)
        If lErro <> SUCESSO And lErro <> 1 Then gError ERRO_SEM_MENSAGEM
        
        Call Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdMask)
        
        GridOP.TextMatrix(iIndice, iGrid_OP_Col) = objItemOP.sCodigo
        GridOP.TextMatrix(iIndice, iGrid_ProdutoOP_Col) = sProdMask
        GridOP.TextMatrix(iIndice, iGrid_OPConsumo_Col) = Formata_Estoque(objItemOPDev.dQuantidade)
        GridOP.TextMatrix(iIndice, iGrid_OPQuantDev_Col) = Formata_Estoque(objItemOPDev.dQuantDevolvida)
        GridOP.TextMatrix(iIndice, iGrid_OPQuantADevolver_Col) = Formata_Estoque(objItemOPDev.dQuantADevolver)
    
        dQtdeDev = dQtdeDev + objItemOPDev.dQuantADevolver
        dQtdeTotalDev = dQtdeTotalDev + (objItemOPDev.dQuantidade - objItemOPDev.dQuantDevolvida)
    Next
    objGridOP.iLinhasExistentes = iIndice
    
    QtdeDev.Caption = Formata_Estoque(dQtdeDev)
    QtdeTotalDev.Caption = Formata_Estoque(dQtdeTotalDev)
    UM.Caption = objItemNF.sUnidadeMed
    
    Traz_Dados_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Tela:

    Traz_Dados_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209990)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209991)

    End Select

    Exit Sub

End Sub

