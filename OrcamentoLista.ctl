VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl OrcamentoLista 
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   8145
   Begin VB.CommandButton BotaoSeleciona 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   735
      Left            =   1935
      Picture         =   "OrcamentoLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3210
      Width           =   1830
   End
   Begin VB.CommandButton BotaoFecha 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   735
      Left            =   3870
      Picture         =   "OrcamentoLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3210
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridOrcamento 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   5054
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "OrcamentoLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjOrcamento As ClassOrcamentoLoja
Dim iAlterado As Integer
Dim gdQuant As Double

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Codigo_Col As Integer
'Dim iGrid_DataValidade_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_DAV_Col As Integer

Public Sub Form_Load()
    
    
    iGrid_Codigo_Col = 0
    iGrid_DAV_Col = 1
    iGrid_Cliente_Col = 2
    
    GridOrcamento.TextMatrix(0, iGrid_Codigo_Col) = "Código"
'    GridOrcamento.TextMatrix(0, iGrid_DataValidade_Col) = "Data Validade"
    GridOrcamento.TextMatrix(0, iGrid_Cliente_Col) = "Cliente"
    GridOrcamento.TextMatrix(0, iGrid_DAV_Col) = "Número DAV"
    
    GridOrcamento.ColWidth(0) = 800
    GridOrcamento.ColWidth(1) = 1200
    GridOrcamento.ColWidth(2) = 5000
    
    
    
    Call Preenche_Grid_Orcamento
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163690)

    End Select

    Exit Sub

End Sub

Function Preenche_Grid_Orcamento() As Long

Dim objVenda As ClassVenda
Dim iIndice As Integer
Dim objCliente As ClassCliente
Dim colOrcamento As New Collection
Dim lErro As Long

On Error GoTo Erro_Preenche_Grid_Orcamento

    'Função Que le os orcamentos
    lErro = CF_ECF("OrcamentoECF_Le1", colOrcamento)
    If lErro <> SUCESSO Then gError 105857

    GridOrcamento.Rows = colOrcamento.Count + 1

    For Each objVenda In colOrcamento
        
        iIndice = iIndice + 1
        
        GridOrcamento.TextMatrix(iIndice, iGrid_Codigo_Col) = objVenda.objCupomFiscal.lNumOrcamento
        GridOrcamento.TextMatrix(iIndice, iGrid_DAV_Col) = objVenda.objCupomFiscal.lNumeroDAV
'        GridOrcamento.TextMatrix(iIndice, iGrid_DataValidade_Col) = Format(objVenda.objCupomFiscal.dtDataEmissao + objVenda.objCupomFiscal.lDuracao, "dd/mm/yyyy")
        For Each objCliente In gcolCliente
            If objCliente.lCodigo = objVenda.objCupomFiscal.lCliente Then GridOrcamento.TextMatrix(iIndice, iGrid_Cliente_Col) = objCliente.sNomeReduzido
        Next
            
    Next
    
    gdQuant = iIndice
    
    Preenche_Grid_Orcamento = SUCESSO
    
    Exit Function
    
Erro_Preenche_Grid_Orcamento:

    Preenche_Grid_Orcamento = gErr
    
    Select Case gErr
    
        Case 105857
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 163691)

    End Select

    Exit Function
    
End Function

Private Sub BotaoFecha_Click()

    Set gobjOrcamento = Nothing
    giRetornoTela = vbCancel
    Unload Me
    
End Sub

Private Sub BotaoSeleciona_Click()

On Error GoTo Erro_BotaoSeleciona_Click
    
    If GridOrcamento.Row = 0 Or GridOrcamento.Row > gdQuant Then Exit Sub
    
    gobjOrcamento.lNumOrcamento = StrParaLong(GridOrcamento.TextMatrix(GridOrcamento.Row, iGrid_Codigo_Col))
'    gobjOrcamento.lDuracao = StrParaLong(StrParaDate(GridOrcamento.TextMatrix(GridOrcamento.Row, iGrid_DataValidade_Col)) - gdtDataHoje)
        
    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case Err
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 163692)

    End Select

    Exit Sub

End Sub

Private Sub GridOrcamento_DblClick()
    
    Call BotaoSeleciona_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

      
End Sub

Function Trata_Parametros(objOrcamento As ClassOrcamentoLoja) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjOrcamento = objOrcamento
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err

        Case Else
        
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 163693)

    End Select

    Exit Function
    
End Function
'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Orcamentos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OrcamentoLista"
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

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



