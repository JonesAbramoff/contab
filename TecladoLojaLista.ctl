VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl TecladoLojaLista 
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   4245
   Begin VB.CommandButton BotaoSeleciona 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   225
      Picture         =   "TecladoLojaLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3225
      Width           =   1830
   End
   Begin VB.CommandButton BotaoFecha 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   2175
      Picture         =   "TecladoLojaLista.ctx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3225
      Width           =   1830
   End
   Begin MSFlexGridLib.MSFlexGrid GridTeclado 
      Height          =   3015
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5318
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
Attribute VB_Name = "TecladoLojaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjTeclado As ClassTecladoProduto
Dim iAlterado As Integer

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Teclado_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_Padrao_Col As Integer

Public Sub Form_Load()
    
    Set gobjTeclado = New ClassTecladoProduto
        
    iGrid_Teclado_Col = 0
    iGrid_Descricao_Col = 1
    iGrid_Padrao_Col = 2
    
    GridTeclado.TextMatrix(0, iGrid_Teclado_Col) = "Teclado"
    GridTeclado.TextMatrix(0, iGrid_Descricao_Col) = "Descrição"
    GridTeclado.TextMatrix(0, iGrid_Padrao_Col) = "Padrão"
    
    If gcolTeclados.Count > 8 Then
        GridTeclado.Rows = gcolTeclados.Count + 1
    Else
        GridTeclado.Rows = 9
    End If
    
    Call Preenche_Grid_Teclado
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    Select Case gErr
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 174554)

    End Select

    Exit Sub

End Sub

Function Preenche_Grid_Teclado() As Long

Dim objTeclado As ClassTecladoProduto
Dim iIndice As Integer

    For Each objTeclado In gcolTeclados
        
        If objTeclado.iTeclado = giTeclado Then
            iIndice = iIndice + 1
               
            GridTeclado.TextMatrix(iIndice, iGrid_Teclado_Col) = objTeclado.iCodigo
            GridTeclado.TextMatrix(iIndice, iGrid_Descricao_Col) = objTeclado.sDescricao
            If objTeclado.iPadrao = TECLADO_PADRAO Then
                GridTeclado.TextMatrix(iIndice, iGrid_Padrao_Col) = "Sim"
            Else
                GridTeclado.TextMatrix(iIndice, iGrid_Padrao_Col) = "Não"
            End If
        End If
    Next
    
End Function

Private Sub BotaoFecha_Click()
    
    giRetornoTela = vbCancel
    Unload Me
    
End Sub

Private Sub BotaoSeleciona_Click()

Dim lErro As Long
Dim objTeclado As New ClassTeclado
Dim obj1 As Object

On Error GoTo Erro_BotaoSeleciona_Click
    
    If GridTeclado.Row = 0 Or GridTeclado.Row > gcolTeclados.Count Then Exit Sub
    
    gobjTeclado.iCodigo = StrParaInt(GridTeclado.TextMatrix(GridTeclado.Row, iGrid_Teclado_Col))
    gobjTeclado.sDescricao = GridTeclado.TextMatrix(GridTeclado.Row, iGrid_Descricao_Col)
    
    Unload Me
    
    giRetornoTela = vbOK
    
    Exit Sub

Erro_BotaoSeleciona_Click:

    Select Case Err
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 174555)

    End Select

    Exit Sub

End Sub

Private Sub GridTeclado_DblClick()
    
    Call BotaoSeleciona_Click
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
   
End Sub

Function Trata_Parametros(objTeclado As ClassTecladoProduto) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjTeclado = objTeclado
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err

        Case Else
        
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error$, 174556)

    End Select

    Exit Function
    
End Function
'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Teclados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TecladoLojaLista"
    
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






