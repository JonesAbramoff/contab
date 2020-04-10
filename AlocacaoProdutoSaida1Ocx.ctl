VERSION 5.00
Begin VB.UserControl AlocacaoProdutoSaida1Ocx 
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   LockControls    =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   6375
   Begin VB.ListBox Tratamento 
      Height          =   645
      ItemData        =   "AlocacaoProdutoSaida1Ocx.ctx":0000
      Left            =   150
      List            =   "AlocacaoProdutoSaida1Ocx.ctx":000D
      TabIndex        =   0
      Top             =   1350
      Width           =   6015
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   5310
      Picture         =   "AlocacaoProdutoSaida1Ocx.ctx":005E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   5310
      Picture         =   "AlocacaoProdutoSaida1Ocx.ctx":01B8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   840
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
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   225
      Width           =   4680
   End
End
Attribute VB_Name = "AlocacaoProdutoSaida1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
'Dim gobjItemPedido As ClassItemPedido
'Dim gobjTelaPai As Form
Dim gobjGenerico As AdmGenerico

Function Trata_Parametros(objGenerico As AdmGenerico) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjGenerico = objGenerico
    
''''    Set gobjTelaPai = objTela
''''    Set gobjItemPedido = objItemPedido

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142758)

    End Select

    Exit Function

End Function

Private Sub BotaoCancela_Click()

    giRetornoTela = vbCancel
    Unload Me

End Sub


Private Sub BotaoOK_Click()

Dim lErro As Long
Dim dFator As Double

On Error GoTo Erro_BotaoOK_Click

    giRetornoTela = vbOK
    
    Select Case Tratamento.ListIndex

        Case NENHUMA_SELECAO
            gobjGenerico.vVariavel = NENHUMA_SELECAO
            giRetornoTela = vbCancel

        Case Else
            gobjGenerico.vVariavel = Tratamento.ListIndex

    End Select

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142759)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If UnloadMode <> vbFormCode Then giRetornoTela = vbCancel

End Sub

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142760)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjGenerico = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ALOCACAO_PRODUTO_SAIDA1
    Set Form_Load_Ocx = Me
    Caption = "Saída de Reserva de Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AlocacaoProdutoSaida1"
    
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



Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

