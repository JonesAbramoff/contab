VERSION 5.00
Begin VB.UserControl NFCeOffline 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   ScaleHeight     =   1080
   ScaleWidth      =   5805
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4290
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   270
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "NFCeOffline.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "F8 - Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "NFCeOffline.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "F5 - Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label LabelQtdeArqs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   1
      Top             =   270
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Quantidade de Arquivos a enviar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   330
      Width           =   2940
   End
End
Attribute VB_Name = "NFCeOffline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub AtualizaQtdeXmlsPendentes()
Dim lErro As Long, lQtdeArqs As Long

    'preencher a qtde de arquivos pendentes
    lErro = CF_ECF("NFCeOffline_QtdePendente", lQtdeArqs)
    If lErro = SUCESSO Or lErro = 201586 Then LabelQtdeArqs.Caption = CStr(lQtdeArqs)
    
    If lErro <> SUCESSO Or lQtdeArqs = 0 Then BotaoGravar.Enabled = False
    
    If lErro = 201586 Then Call Rotina_ErroECF(vbOKOnly, ERRO_NFCE_OFFLINE_NAO_TRANSMITIDA, 201586)
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Call AtualizaQtdeXmlsPendentes

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163638)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros() As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163639)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "NFCe - Enviar Xmls Pendentes"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFCeOffline"

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

Private Sub BotaoGravar_Click()
    
Dim lErro As Long, lQtdeArqs As Long

On Error GoTo Erro_Form_Load

    If StrParaLong(LabelQtdeArqs.Caption) = 0 Then gError 201565
    
    'preencher a qtde de arquivos pendentes
    lErro = CF_ECF("NFCE_Enviar_Offline")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call AtualizaQtdeXmlsPendentes

    Call Rotina_AvisoECF(vbOK, AVISO_NFCE_XMLS_PENDENTES_ENVIADOS)
    
    Exit Sub

Erro_Form_Load:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 201565
            Call Rotina_ErroECF(vbOKOnly, ERRO_NFCE_SEM_OFFLINE_PENDENTE, gErr)
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163638)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Função que Incrementa o Código Atravez da Tecla F2
Dim lErro As Long

On Error GoTo Erro_UserControl_KeyDown

    Select Case KeyCode

        Case vbKeyF8
            
            Call BotaoFechar_Click

    End Select

    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163640)

    End Select

    Exit Sub

End Sub

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

