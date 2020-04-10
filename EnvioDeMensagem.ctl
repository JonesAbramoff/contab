VERSION 5.00
Begin VB.UserControl EnvioDeMensagem 
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   ScaleHeight     =   1395
   ScaleWidth      =   5475
   Begin VB.CommandButton BotaoNao 
      Caption         =   "Não"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton BotaoSimTodos 
      Caption         =   "Sim para Todos"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton BotaoNaoTodos 
      Caption         =   "Não para Todos"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton BotaoSim 
      Caption         =   "Sim"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancelar 
      Height          =   375
      Left            =   4920
      Picture         =   "EnvioDeMensagem.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancelar"
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Texto 
      Caption         =   $"EnvioDeMensagem.ctx":0342
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "EnvioDeMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declarações Globais
Dim iAlterado As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()


Public Sub Trata_Parametros()
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159472)

    End Select

    Exit Sub

End Sub


Private Sub BotaoCancelar_Click()
    
    giRetornoTela = vbCancel
    Unload Me
End Sub

Private Sub BotaoNao_Click()
    
    giRetornoTela = vbNo
    Unload Me
End Sub

Private Sub BotaoNaoTodos_Click()

    giRetornoTela = vbAbort
    Unload Me
End Sub

Private Sub BotaoSim_Click()
    
    giRetornoTela = vbYes
    Unload Me
End Sub

Private Sub BotaoSimTodos_Click()
    
    giRetornoTela = vbIgnore
    Unload Me
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Mensagem"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "EnvioDeMensagem"
    
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

