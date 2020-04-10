VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl NFD2 
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   8970
   Begin VB.Frame Frame2 
      Caption         =   "Novo Item"
      Height          =   945
      Left            =   75
      TabIndex        =   28
      Top             =   1590
      Width           =   8760
      Begin VB.CommandButton BotaoIncluir 
         Caption         =   "(F4)  Incluir"
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
         Left            =   7275
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   525
         Width           =   1350
      End
      Begin VB.TextBox DiscrIncluir 
         Height          =   330
         Left            =   4050
         TabIndex        =   5
         Top             =   180
         Width           =   4590
      End
      Begin MSMask.MaskEdBox PrecoUnitIncluir 
         Height          =   300
         Left            =   4050
         TabIndex        =   7
         Top             =   585
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QtdeIncluir 
         Height          =   300
         Left            =   1170
         TabIndex        =   6
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoNomeRed 
         Height          =   315
         Left            =   1170
         TabIndex        =   4
         Top             =   195
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelProduto 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2700
         TabIndex        =   33
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   75
         TabIndex        =   32
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total: "
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
         Left            =   5415
         TabIndex        =   31
         Top             =   630
         Width           =   555
      End
      Begin VB.Label LabelTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6000
         TabIndex        =   30
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Discriminação:"
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
         Left            =   2715
         TabIndex        =   29
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destinatário"
      Height          =   780
      Left            =   75
      TabIndex        =   27
      Top             =   780
      Width           =   8760
      Begin VB.TextBox Destinatario 
         Height          =   540
         Left            =   120
         MaxLength       =   250
         TabIndex        =   3
         Top             =   195
         Width           =   8505
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6690
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   195
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "NFD2.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "NFD2.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "F7 - Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "NFD2.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "F6 - Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "NFD2.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "F5- Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameIdentificacao 
      Caption         =   "Identificação"
      Height          =   705
      Left            =   75
      TabIndex        =   15
      Top             =   90
      Width           =   6420
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   2385
         TabIndex        =   1
         Top             =   255
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   4935
         TabIndex        =   2
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
      Begin MSComCtl2.UpDown SpinData 
         Height          =   315
         Left            =   6120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   195
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox Serie 
         Height          =   315
         Left            =   750
         TabIndex        =   0
         Top             =   255
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Index           =   3
         Left            =   210
         TabIndex        =   23
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
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
         Left            =   4020
         TabIndex        =   22
         Top             =   315
         Width           =   750
      End
      Begin VB.Label LabelNum 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   1620
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   285
         Width           =   705
      End
   End
   Begin VB.Frame FrameItens 
      Caption         =   "Itens Cadastrados"
      Height          =   3375
      Left            =   75
      TabIndex        =   14
      Top             =   2580
      Width           =   8760
      Begin VB.TextBox Produto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   3900
         TabIndex        =   35
         Top             =   1500
         Width           =   1305
      End
      Begin MSMask.MaskEdBox Qtde 
         Height          =   240
         Left            =   5160
         TabIndex        =   26
         Top             =   1785
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.TextBox Discr 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   240
         Left            =   600
         TabIndex        =   25
         Top             =   1800
         Width           =   3375
      End
      Begin MSMask.MaskEdBox PrecoUnit 
         Height          =   240
         Left            =   6270
         TabIndex        =   24
         Top             =   1620
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PrecoTotal 
         Height          =   240
         Left            =   7335
         TabIndex        =   18
         Top             =   1830
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   2700
         Left            =   75
         TabIndex        =   9
         Top             =   210
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   4763
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label ValorTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7350
         TabIndex        =   20
         Top             =   2970
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
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
         Left            =   6300
         TabIndex        =   19
         Top             =   3030
         Width           =   1005
      End
   End
End
Attribute VB_Name = "NFD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iAlterado As Integer
Dim iAlteradoItem As Integer

'Variáveis Globais
Dim objGridItens As AdmGrid

Dim iGrid_Qtde_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Discr_Col As Integer
Dim iGrid_PrecoUnit_Col As Integer
Dim iGrid_ValorTotal_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal - Modelo d2"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "NFD2"
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

Public Sub Form_Load()
'Função inicialização da Tela
Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Instanciar o objGridCheques para apontar para uma posição de memória
    Set objGridItens = New AdmGrid
    
    'Inicialização de Grid Cheques
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call DateParaMasked(DataEmissao, Date)
    
    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamada
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213430)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing

End Sub

Function Inicializa_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_GridItens

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Qtde")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Discriminação da Mercadoria")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Qtde.Name)
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (Discr.Name)
    objGridInt.colCampo.Add (PrecoUnit.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    
    'Colunas do Grid
    iGrid_Qtde_Col = 1
    iGrid_Produto_Col = 2
    iGrid_Discr_Col = 3
    iGrid_PrecoUnit_Col = 4
    iGrid_ValorTotal_Col = 5
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_LINHAS_GRID

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    'objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_GridItens = SUCESSO

    Exit Function

Erro_Inicializa_GridItens:

    Inicializa_GridItens = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213431)

    End Select

    Exit Function

End Function

Private Sub GridItens_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    
    Call Recalcula_Totais

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_LostFocus()
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub Qtde_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Qtde_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Qtde_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Qtde_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Qtde
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Discr_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Discr_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Discr_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Discr_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Discr
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoUnit_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrecoUnit_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub PrecoUnit_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub PrecoUnit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnit
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrecoTotal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 213432

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 213432
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213433)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    'Fechar a Tela
    Unload Me

End Sub


Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case KEYCODE_BROWSER
            If Me.ActiveControl Is ProdutoNomeRed Then
                Call BotaoProdutos_Click
            Else
                Call LabelNum_Click
            End If

        Case vbKeyF4
            If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
            Call BotaoIncluir_Click

        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoGravar) Then Exit Sub
            Call BotaoGravar_Click
            
        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoExcluir) Then Exit Sub
            Call BotaoExcluir_Click
            
        Case vbKeyF7
            If Not TrocaFoco(Me, BotaoLimpar) Then Exit Sub
            Call BotaoLimpar_Click
            
        Case vbKeyF8
            If Not TrocaFoco(Me, BotaoFechar) Then Exit Sub
            Call BotaoFechar_Click
            
    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213434)

    End Select

    Exit Sub

End Sub

Private Sub LabelNum_Click()

Dim objNF As New ClassNFiscal
Dim lErro As Long
    
On Error GoTo Erro_LabelNum_Click
    
    'Chama tela de MovimentoBoletoLista
    Call Chama_TelaECF_Modal("NFD2Lista", objNF)
    
    If Not (objNF Is Nothing) Then
        'Verifica se o Codvendedor está preenchido e joga na coleção
        
        If objNF.lNumNotaFiscal <> 0 Then
            
            'Função que traz o MovimentoBoleto para a Tela
            lErro = Traz_NFD2_Tela(objNF)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
    End If
    
    Exit Sub

Erro_LabelNum_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213435)
    
    End Select

    Exit Sub

End Sub

Function Traz_NFD2_Tela(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim objItem As ClassItemNF

On Error GoTo Erro_Traz_NFD2_Tela

    lErro = CF_ECF("NFD2_Le", objNF)
    If lErro <> SUCESSO And lErro <> 107850 Then gError ERRO_SEM_MENSAGEM
    
    Numero.Text = objNF.lNumNotaFiscal
    Serie.Text = objNF.sSerie
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objNF.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    Destinatario.Text = objNF.sDestino

    'Limpa Grid
    Call Grid_Limpa(objGridItens)
    
    'Preencher o GridCheque com os Dados dos Cheques Selecionados com o Numero do Movimento passado como parâmetro
    For Each objItem In objNF.colItens
        
        iIndice = iIndice + 1
        
        lErro = Incluir_Item_Grid(objItem.dQuantidade, objItem.sProdutoXml, objItem.sDescricaoItem, objItem.dPrecoUnitario)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
   
    'Carrega o Grid Recalcula Totais
    lErro = Recalcula_Totais()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Traz_NFD2_Tela = SUCESSO
    
    Exit Function

Erro_Traz_NFD2_Tela:

    Traz_NFD2_Tela = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213436)

    End Select

    Exit Function

End Function

Function Incluir_Item_Grid(ByVal dQuantidade As Double, ByVal sProduto As String, ByVal sDescricaoItem As String, ByVal dPrecoUnitario As Double) As Long

Dim lErro As Long

On Error GoTo Erro_Incluir_Item_Grid
    
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Discr_Col) = sDescricaoItem
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Produto_Col) = sProduto
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_PrecoUnit_Col) = Format(dPrecoUnitario, "STANDARD")
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Qtde_Col) = Formata_Estoque(dQuantidade)
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_ValorTotal_Col) = Format(dPrecoUnitario * dQuantidade, "STANDARD")

    objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

    Incluir_Item_Grid = SUCESSO
    
    Exit Function

Erro_Incluir_Item_Grid:

    Incluir_Item_Grid = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213437)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Função que efeuara a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Limpa_Tela_NFD2
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213438)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Botão que Exclui um Movimento de Sangria

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objNF As New ClassNFiscal

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica os dados obrigatórios foram Preenchidos
    If Len(Trim(Numero.Text)) = 0 Then gError 213439
    If Len(Trim(Serie.Text)) = 0 Then gError 213440
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 213441

    'Pergunta se deseja Realmente Excluir o Movimento
    vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_EXCLUSAO_NFD2, Numero.Text)
    If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM
    
    objNF.lNumNotaFiscal = StrParaDbl(Numero.Text)
    objNF.sSerie = Trim(Serie.Text)
    objNF.dtDataEmissao = StrParaDate(DataEmissao.Text)

    'Atualiza os Dados na Memória
    lErro = CF_ECF("NFD2_Exclui", objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'Função Que Limpa a Tela
    Call Limpa_Tela_NFD2

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213439
            Call Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_PREENCHIDO, gErr)

        Case 213440
            Call Rotina_ErroECF(vbOKOnly, ERRO_SERIE_NAO_PREENCHIDA, gErr)
        
        Case 213441
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213442)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objNF As New ClassNFiscal
Dim objNFBD As New ClassNFiscal

On Error GoTo Erro_Gravar_Registro

    'Verifica os dados obrigatórios foram Preenchidos
    If Len(Trim(Numero.Text)) = 0 Then gError 213443
    If Len(Trim(Serie.Text)) = 0 Then gError 213444
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 213445
    
    'Guarda os dados na memoria que serão inseridos
    lErro = Move_NFD2_Memoria(objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objNF.colItens.Count = 0 Then gError 213446
    
    objNFBD.lNumNotaFiscal = objNF.lNumNotaFiscal
    objNFBD.sSerie = objNF.sSerie
    objNFBD.dtDataEmissao = objNF.dtDataEmissao
    
    lErro = CF_ECF("NFD2_Le", objNFBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro <> ERRO_LEITURA_SEM_DADOS Then
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_ALTERACAO_NFD2, Numero.Text)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM
    End If
    
    'Função que grava os dados no arquivao
    lErro = CF_ECF("NFD2_Grava", objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 213443
            Call Rotina_ErroECF(vbOKOnly, ERRO_NUMERO_NAO_PREENCHIDO, gErr)

        Case 213444
            Call Rotina_ErroECF(vbOKOnly, ERRO_SERIE_NAO_PREENCHIDA, gErr)
        
        Case 213445
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
               
        Case 213446
            Call Rotina_ErroECF(vbOKOnly, ERRO_NENHUM_ITEM_GRID, gErr)
               
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213447)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_NFD2()
'Função que Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFD2

    'Limpa os Controles básico da Tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridItens)
    
    ValorTotal.Caption = ""
    LabelTotal.Caption = ""
    
    Exit Sub
    
Erro_Limpa_Tela_NFD2:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213448)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo BotaoLimpar_Click

    'Função que Lima a Tela
    Call Limpa_Tela_NFD2

    Exit Sub

BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213449)

    End Select

    Exit Sub

End Sub

Function Move_NFD2_Memoria(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim objItem As ClassItemNF
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_Move_NFD2_Memoria
    
    objNF.lNumNotaFiscal = StrParaDbl(Numero.Text)
    objNF.sSerie = Trim(Serie.Text)
    objNF.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNF.sDestino = Destinatario.Text
    
    'Preencher o GridCheque com os Dados dos Cheques Selecionados com o Numero do Movimento passado como parâmetro
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objItem = New ClassItemNF
        
        objItem.iItem = iIndice
        objItem.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_Discr_Col)
        objItem.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Qtde_Col))
        objItem.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnit_Col))
        objItem.sProdutoXml = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        
        sProduto = objItem.sProdutoXml
        
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto, objProduto)
        
        objItem.sProduto = objProduto.sCodigo

        objNF.ColItensNF.Add1 objItem
        
    Next
   
    Move_NFD2_Memoria = SUCESSO
    
    Exit Function

Erro_Move_NFD2_Memoria:

    Move_NFD2_Memoria = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213450)

    End Select

    Exit Function

End Function

Function Recalcula_Totais() As Long
'Calcula o Total

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotal As Double

On Error GoTo Erro_Recalcula_Totais

    'Para todos os Cheque do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorTotal_Col))
    Next
            
    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Recalcula_Totais = SUCESSO

    Exit Function

Erro_Recalcula_Totais:

    Recalcula_Totais = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213451)

    End Select

    Exit Function

End Function

Private Sub BotaoIncluir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoIncluir_Click
    
    'verificação do preenchimento dos campos
    If Len(Trim(ProdutoNomeRed.Text)) = 0 Then gError 213455
    If Len(Trim(DiscrIncluir.Text)) = 0 Then gError 213453
    If StrParaDbl(QtdeIncluir.Text) = 0 Then gError 213452
    If StrParaDbl(PrecoUnitIncluir.Text) = 0 Then gError 213454
    
    lErro = Incluir_Item_Grid(StrParaDbl(QtdeIncluir.Text), ProdutoNomeRed.Text, DiscrIncluir.Text, StrParaDbl(PrecoUnitIncluir.Text))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    Call Recalcula_Totais
    
    'Limpa os campos da tela
    QtdeIncluir.Text = ""
    DiscrIncluir.Text = ""
    PrecoUnitIncluir.Text = ""
    ProdutoNomeRed.Text = ""
    
    LabelTotal.Caption = ""
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr
    
        Case 213452
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_NAO_PREENCHIDO1, gErr)
    
        Case 213453
            Call Rotina_ErroECF(vbOKOnly, ERRO_DISCRIMINACAO_NAO_PREENCHIDA, gErr)
    
        Case 213454
            Call Rotina_ErroECF(vbOKOnly, ERRO_PRECO_NAO_PREENCHIDO, gErr)
    
        Case 213455
            Call Rotina_ErroECF(vbOKOnly, ERRO_PRODUTO_NAO_PREENCHIDO1, gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213455)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_SpinData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213456)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_SpinData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 213457)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataEmissao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DataEmissao_Validate
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
        
    Exit Sub
    
Erro_DataEmissao_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213458)

    End Select

    Exit Sub
    
End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus()
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Numero_Validate
    
    If Len(Trim(Numero.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Numero.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
        
    Exit Sub
    
Erro_Numero_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213459)

    End Select

    Exit Sub
    
End Sub

Private Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Destinatario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub QtdeIncluir_GotFocus()
    Call MaskEdBox_TrataGotFocus(QtdeIncluir, iAlteradoItem)
End Sub

Private Sub QtdeIncluir_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_QtdeIncluir_Validate
    
    If Len(Trim(QtdeIncluir.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(QtdeIncluir.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        QtdeIncluir.Text = Format(StrParaDbl(QtdeIncluir.Text), "STANDARD")
        
    End If
    
    LabelTotal.Caption = Format(StrParaDbl(PrecoUnitIncluir.Text) * StrParaDbl(QtdeIncluir.Text), "STANDARD")
        
    Exit Sub
    
Erro_QtdeIncluir_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213460)

    End Select

    Exit Sub
    
End Sub

Private Sub PrecoUnitIncluir_GotFocus()
    Call MaskEdBox_TrataGotFocus(PrecoUnitIncluir, iAlteradoItem)
End Sub

Private Sub PrecoUnitIncluir_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_PrecoUnitIncluir_Validate
    
    If Len(Trim(PrecoUnitIncluir.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(PrecoUnitIncluir.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
        PrecoUnitIncluir.Text = Format(StrParaDbl(PrecoUnitIncluir.Text), "STANDARD")
     
    End If
       
    LabelTotal.Caption = Format(StrParaDbl(PrecoUnitIncluir.Text) * StrParaDbl(QtdeIncluir.Text), "STANDARD")
        
    Exit Sub
    
Erro_PrecoUnitIncluir_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213461)

    End Select

    Exit Sub
    
End Sub

Private Sub ProdutoNomeRed_GotFocus()
    ProdutoNomeRed.SelStart = 0
    ProdutoNomeRed.SelLength = Len(ProdutoNomeRed.Text)
End Sub

Private Sub ProdutoNomeRed_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_ProdutoNomeRed_Validate
    
    Parent.MousePointer = vbHourglass
    
    If Len(Trim(ProdutoNomeRed.Text)) <> 0 Then
    
        sProduto = ProdutoNomeRed.Text
        
        Call TP_Produto_Le_Col(gaobjProdutosReferencia, gaobjProdutosCodBarras, gaobjProdutosNome, sProduto, objProduto)
                
        'caso o produto não seja encontrado
        If objProduto Is Nothing Then gError ERRO_SEM_MENSAGEM
        
        DiscrIncluir.Text = objProduto.sDescricao
        PrecoUnitIncluir.Text = Format(objProduto.dPrecoLoja, "STANDARD")
        
    Else
    
        DiscrIncluir.Text = ""
        PrecoUnitIncluir.Text = ""
        
    End If
    
    Parent.MousePointer = vbDefault
    
    Exit Sub

Erro_ProdutoNomeRed_Validate:

    Cancel = True

    Parent.MousePointer = vbDefault

    Select Case gErr
                
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213511)

    End Select
    
    Exit Sub

End Sub

Public Sub LabelProduto_Click()

    Call BotaoProdutos_Click

End Sub

Public Sub BotaoProdutos_Click()
'Chama o browser do ProdutoLojaLista
'So traz produtos onde codigo de barras ou referencia está preenchida

Dim objProduto As New ClassProduto

On Error GoTo Erro_BotaoProdutos_Click
    
    objProduto.sNomeReduzido = ProdutoNomeRed.Text
    
    'Chama tela de ProdutosLista
    Call Chama_TelaECF_Modal("ProdutosLista", objProduto)
        
    UserControl.Refresh
    
    If giRetornoTela = vbOK Then
        If Len(Trim(objProduto.sReferencia)) > 0 Then
            ProdutoNomeRed.Text = objProduto.sReferencia
        Else
            ProdutoNomeRed.Text = objProduto.sCodigoBarras
        End If
        Call ProdutoNomeRed_Validate(False)
    End If
    
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 213512)

    End Select

    Exit Sub

End Sub
