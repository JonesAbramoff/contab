VERSION 5.00
Begin VB.UserControl RecebimentoCarne2 
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4695
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      Height          =   585
      Left            =   2700
      Picture         =   "RecebimentoCarne2.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   165
      Width           =   1545
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      Height          =   585
      Left            =   615
      Picture         =   "RecebimentoCarne2.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   165
      Width           =   1545
   End
   Begin VB.Frame FrameSelecionaOpcao 
      Caption         =   "Selecione uma opção:"
      Height          =   1935
      Left            =   105
      TabIndex        =   0
      Top             =   930
      Width           =   4455
      Begin VB.OptionButton GravarAutenticar 
         Caption         =   "Gravar e autenticar as parcelas selecionadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   4215
      End
      Begin VB.OptionButton GravarImprimir 
         Caption         =   "Gravar e imprimir as parcelas selecionadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.OptionButton ApenasImprimir 
         Caption         =   "Apenas imprimir as parcelas selecionadas"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "RecebimentoCarne2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjRecebimentoCarne As ClassRecebimentoCarne

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Trata_Parametros(objRecebimentoCarne As ClassRecebimentoCarne) As Long
        
    Set gobjRecebimentoCarne = objRecebimentoCarne
    
    Trata_Parametros = SUCESSO

End Function

Public Sub Form_Load()

On Error GoTo Erro_Form_Load
    
    Set gobjRecebimentoCarne = New ClassRecebimentoCarne
    
    'default
    ApenasImprimir.Value = True
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166250)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoCancelar_Click()

    giRetornoTela = vbCancel
    
    Unload Me
    
End Sub

Private Sub BotaoOk_Click()
    
    If ApenasImprimir.Value = True Then
        gobjRecebimentoCarne.iOpcao = 0
    Else
        If GravarImprimir.Value = True Then
            gobjRecebimentoCarne.iOpcao = 1
        Else
            gobjRecebimentoCarne.iOpcao = 2
        End If
    End If
    
    giRetornoTela = vbOK
    
    Unload Me
    
End Sub

Private Sub form_unload()
    Unload Me
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Recebimento de Carnê - Operação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RecebimentoCarne2"
    
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

