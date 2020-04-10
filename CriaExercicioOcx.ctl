VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl CriaExercicioOcx 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LockControls    =   -1  'True
   ScaleHeight     =   1485
   ScaleWidth      =   4365
   Begin VB.CommandButton BotaoCancel 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2385
      Picture         =   "CriaExercicioOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   765
      Width           =   975
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   960
      Picture         =   "CriaExercicioOcx.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   765
      Width           =   975
   End
   Begin MSMask.MaskEdBox Exercicio 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   120
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   120
      Top             =   120
      Width           =   1005
      ForeColor       =   128
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Caption         =   "Exercício : "
   End
End
Attribute VB_Name = "CriaExercicioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Responsavel: Mario
'Revisado em 20/8/98

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objExercicio1 As ClassExercicio

Private Sub BotaoCancel_Click()
    Unload Me
    objExercicio1.sNomeExterno = ""
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(Exercicio.ClipText)) = 0 Then Error 11372
    
    objExercicio1.sNomeExterno = Exercicio.Text

    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case Err
    
        Case 11372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_PREENCHIDO", Err)
            Exercicio.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155286)
        
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Load()
    
    lErro_Chama_Tela = SUCESSO
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objExercicio1 = Nothing
    
End Sub


Function Trata_Parametros(objExercicio As ClassExercicio) As Long

    Set objExercicio1 = objExercicio

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CRIAR_EXERCICIO
    Set Form_Load_Ocx = Me
    Caption = "Criar Exercício"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CriaExercicio"
    
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



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

