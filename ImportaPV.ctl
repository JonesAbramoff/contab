VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ImportaPV 
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   ScaleHeight     =   3225
   ScaleWidth      =   5295
   Begin VB.Timer Timer1 
      Left            =   4260
      Top             =   2775
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1815
      TabIndex        =   1
      Top             =   2790
      Width           =   1740
   End
   Begin VB.ListBox Mensagem 
      Height          =   1425
      Left            =   165
      TabIndex        =   0
      Top             =   1125
      Width           =   4905
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   135
      TabIndex        =   2
      Top             =   150
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1275
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "ImportaPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()
Public giStop As Integer

Public Sub Form_Load()
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
End Sub

Function Trata_Parametros() As Long

    Timer1.Interval = 1

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Importação de Pedido de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ImportaPV"
    
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

Private Sub BotaoCancelar_Click()
    giStop = 1
End Sub

Private Sub Timer1_Timer()

Dim sArq As String
Dim objProgresso As Object
Dim objMsg As Object
Dim objTela As Object
Dim lErro As Long

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

    CommonDialog1.Flags = cdlOFNExplorer Or cdlOFNAllowMultiselect Or cdlOFNFileMustExist
    CommonDialog1.DialogTitle = "Favor informar o arquivo para importação do pedido de venda"
    CommonDialog1.Filter = "Excel (*.xls)|*.xls|Todos os Arquivo (*.*)|*.*"
    CommonDialog1.ShowOpen
    sArq = CommonDialog1.FileName

    Set objProgresso = ProgressBar1
    
    Set objMsg = Mensagem
    
    Set objTela = Me

    If Len(sArq) > 0 Then
    
        lErro = CF("Importa_Dados_Excel", giFilialEmpresa, sArq, objMsg, objProgresso, objTela)
        If lErro <> SUCESSO Then gError 197681
        
        GL_objMDIForm.MousePointer = vbDefault
        
        Call Rotina_Aviso(vbOKOnly, "IMPORTACAO_COMPLETADA_SUCESSSO")
    
    End If
    
    Unload Me
    
    Exit Sub
    
Erro_Timer1_Timer:

    Select Case gErr
        
        Case 197681
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197682)
        
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





