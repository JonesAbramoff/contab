VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl Fininvest 
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ScaleHeight     =   1185
   ScaleWidth      =   3390
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
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
      Left            =   165
      TabIndex        =   1
      Top             =   720
      Width           =   1485
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "(Esc)  Cancelar"
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
      Left            =   1770
      TabIndex        =   0
      Top             =   735
      Width           =   1485
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   315
      Left            =   1365
      TabIndex        =   2
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   780
      TabIndex        =   3
      Top             =   180
      Width           =   510
   End
End
Attribute VB_Name = "Fininvest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Valor_Validate
    
    If StrParaDbl(Valor.Text) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 133836
        
    End If
    
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 133836
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 160352)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros() As Long
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Private Sub BotaoCancelar_Click()

    Unload Me
    
End Sub

Private Sub BotaoOk_Click()
    
Dim lErro As Long
Dim objFormMsg As Object
Dim objTela As Object

On Error GoTo Erro_BotaoOk_Click

    If StrParaDbl(Valor.Text) = 0 Then gError 133837

    Set objTela = Me
    Set objFormMsg = MsgTEF
    
    lErro = CF_ECF("TEF_Fininvest", StrParaDbl(Valor.Text), objFormMsg, objTela)
    If lErro <> SUCESSO Then gError 133838
    
    Unload Me
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 133837
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO1, gErr)
        
        Case 133838
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 160353)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Clique em f5
    If KeyCode = vbKeyF5 Then
        If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
        Call BotaoOk_Click
    End If

    'Clique em esc
    If KeyCode = vbKeyEscape Then
        If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
        Call BotaoCancelar_Click
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Fininvest"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CancelaItem"
    
End Function

Public Function objParent() As Object

    Set objParent = Parent
    
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'***** fim do trecho a ser copiado ******



