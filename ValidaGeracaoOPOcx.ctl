VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ValidaGeracaoOP 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ScaleHeight     =   5370
   ScaleWidth      =   3630
   Begin VB.CommandButton BotaoOk 
      Caption         =   "Prosseguir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   518
      Picture         =   "ValidaGeracaoOPOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4710
      Width           =   1170
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
      Height          =   555
      Left            =   1943
      Picture         =   "ValidaGeracaoOPOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4710
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   3270
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   3255
      Begin MSMask.MaskEdBox Produto 
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   375
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataPrevisaoInicio 
         Height          =   255
         Left            =   1875
         TabIndex        =   6
         Top             =   390
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridProdutos 
         Height          =   2910
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   5133
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"ValidaGeracaoOPOcx.ctx":025C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   210
      TabIndex        =   3
      Top             =   3495
      Width           =   3165
   End
End
Attribute VB_Name = "ValidaGeracaoOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGridProd As New AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DataInicProd_Col As Integer

Private Sub BotaoOK_Click()

    giRetornoTela = vbOK
     
    Unload Me
    
End Sub

Private Sub BotaoCancela_Click()

    giRetornoTela = vbCancel
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giRetornoTela = vbCancel

    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175667)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridProd = Nothing
    
End Sub

Function Trata_Parametros(objOrdemProd As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim colItensOP As New Collection
Dim objItemOP As ClassItemOP
Dim iIndice As Integer
Dim sProdutoForm As String

On Error GoTo Erro_Trata_Parametros
    
    For Each objItemOP In objOrdemProd.colItens
        If objItemOP.dtDataInicioProd < Date Then colItensOP.Add objItemOP
    Next
    

    lErro = Inicializa_Grid_Produtos(colItensOP.Count)
    If lErro <> SUCESSO Then Error 62643
    
    iIndice = 0
    
    For Each objItemOP In colItensOP
        iIndice = iIndice + 1
        GridProdutos.TextMatrix(iIndice, 0) = ""
        
        lErro = Mascara_MascararProduto(objItemOP.sProduto, sProdutoForm)
        If lErro <> SUCESSO Then Error 62644
        
        GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoForm
        GridProdutos.TextMatrix(iIndice, iGrid_DataInicProd_Col) = Format(objItemOP.dtDataInicioProd, "dd/mm/yyyy")
    Next
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 62643, 62644
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175668)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Valida Geração da Ordem de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ValidaGeracaoOP"
    
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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub


Private Function Inicializa_Grid_Produtos(iLinhas As Integer) As Long

Dim iIndice As Integer

    Set objGridProd = New AdmGrid

    'tela em questão
    Set objGridProd.objForm = Me

    'titulos do grid
    objGridProd.colColuna.Add ("")
    objGridProd.colColuna.Add ("Produto")
    objGridProd.colColuna.Add ("PrevisãoInicio")

    'Controles que participam do Grid
    objGridProd.colCampo.Add (Produto.Name)
    objGridProd.colCampo.Add (DataPrevisaoInicio.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DataInicProd_Col = 2
    
    objGridProd.objGrid = GridProdutos

    'Todas as linhas do grid
    objGridProd.objGrid.Rows = iLinhas + 1

    If iLinhas > 10 Then
        objGridProd.iLinhasVisiveis = 10
    Else
        objGridProd.iLinhasVisiveis = iLinhas
    End If

    'Largura da primeira coluna
    GridProdutos.ColWidth(0) = 0

    objGridProd.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGridProd)

    Inicializa_Grid_Produtos = SUCESSO

End Function

