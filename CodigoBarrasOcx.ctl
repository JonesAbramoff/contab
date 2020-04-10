VERSION 5.00
Begin VB.UserControl CodigoBarrasOcx 
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1350
   ScaleWidth      =   2895
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   372
      Picture         =   "CodigoBarrasOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   732
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      Height          =   510
      Left            =   1548
      Picture         =   "CodigoBarrasOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   732
      Width           =   855
   End
   Begin VB.TextBox Codigo 
      Height          =   288
      Left            =   156
      TabIndex        =   0
      Top             =   180
      Width           =   2616
   End
End
Attribute VB_Name = "CodigoBarrasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjProduto As ClassProduto
Dim glErro As Long

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)


End Sub

Private Sub BotaoCancela_Click()
    gobjProduto.sCodigoBarras = "Cancel"
    Unload Me
End Sub

Public Sub Form_Load()
    
On Error GoTo Erro_Form_Load
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 199065)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros(objProduto As ClassProduto) As Long

    Set gobjProduto = objProduto

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoOK_Click()

Dim objProduto As New ClassProduto
Dim lErro As Long

On Error GoTo Erro_BotaoOK_Click

    If Len(Codigo.Text) = 12 Then
    
        Codigo.Text = "4" & Right(Codigo.Text, 11)
    
        lErro = CF("ProdutoCodBarras_Le12", Codigo.Text, gobjProduto)
        If lErro <> SUCESSO And lErro <> 199873 Then gError 199878
        
        If lErro = SUCESSO Then
            Codigo.Text = gobjProduto.sCodigoBarras
        End If
    Else

        lErro = CF("ProdutoCodBarras_Le", Codigo.Text, gobjProduto)
        If lErro <> SUCESSO And lErro <> 193965 Then gError 199879
    
    End If



'    If Len(Codigo.Text) = 12 Then
'        Codigo.Text = "0" & Codigo.Text
'    End If
'
'    gobjProduto.sCodigoBarras = Codigo.Text
'
'    lErro = CF("ProdutoCodBarras_Le", Codigo.Text, gobjProduto)
'    If lErro <> SUCESSO And lErro <> 193965 Then gError 199062
    
    'se o codigo de barras nao existir
    If lErro <> SUCESSO Then gError 199063
    
    gobjProduto.lErro = SUCESSO
    
    Unload Me

    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 199063
            If gobjProduto.lErro = 0 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_BARRAS_NAO_ENCONTRADO", gErr, Codigo.Text)
            Else
                gobjProduto.lErro = gErr
                Unload Me
            End If

        Case 199878, 199879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199064)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANAIS_VENDA
    Set Form_Load_Ocx = Me
    Caption = "Código de Barras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CanalDeVenda"
    
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


