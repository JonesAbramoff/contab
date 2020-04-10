VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CargaBalanca 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ScaleHeight     =   1110
   ScaleWidth      =   3720
   Begin VB.CommandButton BotaoCarga 
      Caption         =   "Fazer a Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   465
      TabIndex        =   2
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   315
      Left            =   2835
      Picture         =   "CargaBalanca.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Fechar"
      Top             =   105
      Width           =   780
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "CargaBalanca"
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


Private Sub BotaoCarga_Click()

Dim objBarraProgresso As Object
Dim lErro As Long

On Error GoTo Erro_BotaoCarga_Click

    Set objBarraProgresso = BarraProgresso

    lErro = CF("Gera_Arquivo_CargaBalanca", giFilialEmpresa, 1, objBarraProgresso)
    If lErro <> SUCESSO Then gError 178897

    Exit Sub
    
Erro_BotaoCarga_Click:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178898)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160716)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Carga da Balança"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CargaBalanca"
    
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

