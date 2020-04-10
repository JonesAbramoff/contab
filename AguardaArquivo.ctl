VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl AguardaArquivo 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ScaleHeight     =   2010
   ScaleWidth      =   6975
   Begin VB.CommandButton BotaoCancelar 
      Height          =   615
      Left            =   2670
      Picture         =   "AguardaArquivo.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   495
      Top             =   1230
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   375
      TabIndex        =   0
      Top             =   510
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "AguardaArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iCancelaBatch As Integer
Dim gsArquivo As String

Public Sub Form_Load()

    ProgressBar1.Max = 1200
    ProgressBar1.Min = 1
    ProgressBar1.Value = 1
    Parent.top = Screen.Height / 3
    Parent.left = Screen.Width / 3
    
    Timer1.Interval = 100

    lErro_Chama_Tela = SUCESSO
    
    
End Sub


Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Aguardando geração do arquivo"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "AguardaArquivo"

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

Private Sub Timer1_Timer()
    
Dim vbMsgRes As VbMsgBoxResult
Dim sArq As String
    
    sArq = Dir(gsArquivo)
    
    If Len(sArq) > 0 Then
        iCancelaBatch = CANCELA_BATCH
        Unload Me
        Exit Sub
    End If
    
    If ProgressBar1.Value < ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Value + 1
    Else
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_DESEJA_PROSSEGUIR_ACOMPANHAMENTO_ARQUIVO, gsArquivo)
        If vbMsgRes = vbYes Then
            ProgressBar1.Value = 1
        Else
            iCancelaBatch = CANCELA_BATCH
            Unload Me
            Exit Sub
        End If
    End If
    
    DoEvents
    
    If iCancelaBatch = CANCELA_BATCH Then
        vbMsgRes = Rotina_AvisoECF(vbYesNo, AVISO_CANCELAR_ACOMPANHAMENTO_ARQUIVO, gsArquivo)
        If vbMsgRes = vbYes Then Unload Me
        iCancelaBatch = 0
    End If
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim lErro As Long
    
On Error GoTo Erro_UserControl_KeyDown
    
    Select Case KeyCode
    
        Case vbKeyF8
            Call BotaoCancelar_Click
    
    End Select
    
    Exit Sub

Erro_UserControl_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 210062)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    If iCancelaBatch <> CANCELA_BATCH Then
        iCancelaBatch = CANCELA_BATCH
    End If

End Sub

Private Sub BotaoCancelar_Click()
    iCancelaBatch = CANCELA_BATCH
End Sub

Function Trata_Parametros(sArquivo As String) As Long
    
    gsArquivo = sArquivo
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

