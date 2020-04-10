VERSION 5.00
Begin VB.UserControl AlocacaoProdutoSaidaOcx 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LockControls    =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   6360
   Begin VB.ListBox Tratamento 
      Height          =   645
      ItemData        =   "AlocacaoProdutoSaidaOcx.ctx":0000
      Left            =   120
      List            =   "AlocacaoProdutoSaidaOcx.ctx":000D
      TabIndex        =   0
      Top             =   1305
      Width           =   6045
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   5310
      Picture         =   "AlocacaoProdutoSaidaOcx.ctx":005E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   5310
      Picture         =   "AlocacaoProdutoSaidaOcx.ctx":01B8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "A quantidade total reservada do item é menor do que a quantidade a reservar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   150
      TabIndex        =   4
      Top             =   210
      Width           =   4680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tratamento"
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
      Left            =   135
      TabIndex        =   3
      Top             =   1035
      Width           =   975
   End
End
Attribute VB_Name = "AlocacaoProdutoSaidaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Dim gcolItemPedido As colItemPedido
'Dim gobjTelaPai As Form
'Dim gobjGrid As AdmGrid
'Dim giListIndexAnterior As Integer

Dim gobjGenerico As AdmGenerico

Dim iGrid_QuantReservada_Col As Integer
Dim iGrid_Responsavel_Col As Integer

Private Sub BotaoCancela_Click()

    giRetornoTela = vbCancel
    Unload Me

End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim dFator As Double

On Error GoTo Erro_BotaoOK_Click
    
    Select Case Tratamento.ListIndex
        Case NENHUMA_SELECAO
            gobjGenerico.vVariavel = NENHUMA_SELECAO
            giRetornoTela = vbCancel

        Case SELECAO_OK
            gobjGenerico.vVariavel = SELECAO_OK
            giRetornoTela = vbOK

        Case CANCELA_ACIMA_DA_RESERVADA
            gobjGenerico.vVariavel = CANCELA_ACIMA_DA_RESERVADA
            giRetornoTela = vbOK

        Case NAO_RESERVAR_PRODUTO
            gobjGenerico.vVariavel = NAO_RESERVAR_PRODUTO
            giRetornoTela = vbOK

    End Select
    
    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142761)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Tratamento.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142762)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjGenerico = Nothing
    
End Sub

Function Trata_Parametros(objGenerico As AdmGenerico) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjGenerico = objGenerico

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142763)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If UnloadMode <> vbFormCode Then giRetornoTela = vbCancel

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ALOCACAO_PRODUTO_SAIDA
    Set Form_Load_Ocx = Me
    Caption = "Saída de Reserva de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AlocacaoProdutoSaida"
    
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



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

