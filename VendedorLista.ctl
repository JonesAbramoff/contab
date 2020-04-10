VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl VendedorLista 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   4425
   Begin VB.CommandButton BotaoFecha 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   2265
      Picture         =   "VendedorLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1830
   End
   Begin VB.CommandButton BotaoSeleciona 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   300
      Picture         =   "VendedorLista.ctx":0272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridVendedor 
      Height          =   3015
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5318
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "VendedorLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVendedor As ClassVendedor
Dim iAlterado As Integer

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Codigo_Col As Integer
Dim iGrid_Nome_Col As Integer

Public Sub Form_Load()
    
    Set gobjVendedor = New ClassVendedor
        
    iGrid_Codigo_Col = 0
    iGrid_Nome_Col = 1
    
    GridVendedor.TextMatrix(0, iGrid_Codigo_Col) = "Código"
    GridVendedor.TextMatrix(0, iGrid_Nome_Col) = "Nome"
    
    If gcolVendedores.Count > 8 Then
        GridVendedor.Rows = gcolVendedores.Count + 1
    Else
        GridVendedor.Rows = 9
    End If
    
    Call Preenche_Grid_Vendedor
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175734)

    End Select

    Exit Sub

End Sub

Function Preenche_Grid_Vendedor() As Long

Dim objVendedor As ClassVendedor
Dim iIndice As Integer

    For Each objVendedor In gcolVendedores
        
        iIndice = iIndice + 1
                
        GridVendedor.TextMatrix(iIndice, iGrid_Codigo_Col) = objVendedor.iCodigo
        GridVendedor.TextMatrix(iIndice, iGrid_Nome_Col) = objVendedor.sNomeReduzido
        
    Next
    
End Function

Private Sub BotaoFecha_Click()
    
    giRetornoTela = vbCancel
    Unload Me
    
End Sub

Private Sub BotaoSeleciona_Click()

Dim lErro As Long
Dim objProduto As New ClassVendedor
Dim obj1 As Object

On Error GoTo Erro_BotaoSeleciona_Click
    
    If GridVendedor.Row = 0 Or GridVendedor.Row > gcolVendedores.Count Then Exit Sub
    
    gobjVendedor.iCodigo = StrParaInt(GridVendedor.TextMatrix(GridVendedor.Row, iGrid_Codigo_Col))
    gobjVendedor.sNomeReduzido = GridVendedor.TextMatrix(GridVendedor.Row, iGrid_Nome_Col)
    
    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case Err
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 175735)

    End Select

    Exit Sub

End Sub

Private Sub GridVendedor_DblClick()
    
    Call BotaoSeleciona_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
   
End Sub

Function Trata_Parametros(objVendedor As ClassVendedor) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjVendedor = objVendedor
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err

        Case Else
        
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 175736)

    End Select

    Exit Function
    
End Function
'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Vendedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "VendedorLista"
    
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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property

'**** fim do trecho a ser copiado *****




Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        Call BotaoSeleciona_Click
    End If
    
    
    'Clique em F8
    If KeyCode = vbKeyEscape Then
        Call BotaoFecha_Click
    End If
  
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

