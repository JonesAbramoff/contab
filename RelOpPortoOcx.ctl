VERSION 5.00
Begin VB.UserControl RelOpPortoOcx 
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   ScaleHeight     =   1440
   ScaleWidth      =   3090
   Begin VB.Frame Frame1 
      Caption         =   "Porto"
      Height          =   630
      Left            =   60
      TabIndex        =   2
      Top             =   390
      Width           =   2955
      Begin VB.TextBox Porto 
         Height          =   270
         Left            =   540
         TabIndex        =   0
         Top             =   225
         Width           =   1890
      End
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "Ok"
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Top             =   1110
      Width           =   1035
   End
   Begin VB.CommandButton BotaoCancel 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   1725
      TabIndex        =   3
      Top             =   1110
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Conhecimento Frete:"
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   1575
   End
   Begin VB.Label CFrete 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   1650
      TabIndex        =   4
      Top             =   90
      Width           =   1335
   End
End
Attribute VB_Name = "RelOpPortoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim sRetorno As String
Dim gobjCompserv As ClassCompServ
Dim iAlterado As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
       
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCompServ As ClassCompServ) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Set gobjCompserv = objCompServ
    
    CFrete.Caption = gobjCompserv.lCodigo
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Sub BotaoCancel_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoOk_Click()
    
   
    gobjCompserv.sPorto = Porto.Text
    
    Unload Me
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Porto"
    Call Form_Load

End Function

Public Sub Form_Unload(Cancel As Integer)
  
    Set gobjCompserv = Nothing
    
End Sub

Public Function Name() As String
    
    Name = "RelOpPorto"

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
'''    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******
    
