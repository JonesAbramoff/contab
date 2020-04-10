VERSION 5.00
Begin VB.UserControl FechamentoMesEst1 
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ScaleHeight     =   1320
   ScaleWidth      =   5760
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5295
      Top             =   210
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   4140
      Picture         =   "FechamentoMesEst1.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   870
   End
   Begin VB.Label labelPasso 
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
      Left            =   435
      TabIndex        =   3
      Top             =   165
      Width           =   2055
   End
   Begin VB.Label ProdutosProcessados 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2505
      TabIndex        =   2
      Top             =   585
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Produtos Processados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   495
      TabIndex        =   1
      Top             =   615
      Width           =   1995
   End
End
Attribute VB_Name = "FechamentoMesEst1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iCancela As Integer
Public objEstoque As ClassEstoqueMes
Dim gbSemAviso As Boolean

Private Sub Cancelar_Click()

    iCancela = CANCELA

End Sub

Public Sub Form_Load()

    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEstoque = Nothing
    
End Sub

Function Trata_Parametros(Optional objEstoqueMes As ClassEstoqueMes, Optional bSemAviso As Boolean = False) As Long

    Set objEstoque = objEstoqueMes
    gbSemAviso = bSemAviso

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If iCancela <> CANCELA Then
        iCancela = CANCELA
        Cancel = 1
    End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()

Dim lErro As Long, objAux As Object

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0
    ProdutosProcessados.Caption = 0
    iCancela = CANCELA_BATCH

    Set objAux = Me
    'Rotina onde chamara as funcoes que alteram o BD
    lErro = CF("Rotina_FechamentoMes", objEstoque, iCancela, objAux)
    If lErro <> SUCESSO Then Error 40683

    'avisa o termino do fechamento do mes
    If Not gbSemAviso Then Call Rotina_Aviso(vbOKOnly, "AVISO_TERMINO_ABERTURA_MES")
    
    iCancela = CANCELA

    Unload Me

    Exit Sub

Erro_Timer1_Timer:

    Select Case Err
        
        Case 40683
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160190)
    
    End Select
    
    iCancela = CANCELA

    Unload Me
    
    Exit Sub
    
End Sub



'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FECHAMENTO_MES_ESTOQUE1
    Set Form_Load_Ocx = Me
    Caption = "Abertura de Mês - Estoque - Acompanhamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FechamentoMesEst1"
    
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

'**** fim do trecho a ser copiado *****



Private Sub ProdutosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutosProcessados, Source, X, Y)
End Sub

Private Sub ProdutosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutosProcessados, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

