VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTAnotacoes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Private gobjAnotacao As ClassAnotacoes

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Titulo() As Object
     Set Titulo = objUserControl.Controls("Titulo")
End Property

Public Property Get Anotacao() As Object
     Set Anotacao = objUserControl.Controls("Anotacao")
End Property

Public Property Get LabelTexto() As Object
     Set LabelTexto = objUserControl.Controls("LabelTexto")
End Property

Public Property Get LabelTitulo() As Object
     Set LabelTitulo = objUserControl.Controls("LabelTitulo")
End Property

'--- fim dos properties get dos controles da tela

Public Function Trata_Parametros(ByVal objAnotacoes As ClassAnotacoes) As Long
'espera o Id e a Origem de ques está chamando a tela
    
Dim lErro As Long
    
On Error GoTo Erro_Trata_Parametros
        
    Set gobjAnotacao = objAnotacoes

    Titulo.Text = objAnotacoes.sTitulo
    Anotacao.Text = objAnotacoes.sTextoCompleto
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PLEO_VB", gErr, Error)
    
    End Select
    
    Exit Function
    
End Function

Private Function Move_Tela_Memoria(ByVal objAnotacoes As ClassAnotacoes) As Long

    objAnotacoes.sTitulo = Trim(Titulo.Text)
    objAnotacoes.sTextoCompleto = Anotacao.Text

End Function

Public Sub BotaoFechar_Click()
'fecha a tela

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142819)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

    Call Move_Tela_Memoria(gobjAnotacao)
    
    Unload Me

End Sub

Public Sub BotaoLimpar_Click()
'limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa a tela
    Call Limpa_Tela_Anotacoes

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142820)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()
'carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142821)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_Anotacoes()
'limpa toda a tela, exeto o frame de id/origem

    'limpa toda a tela
    Call Limpa_Tela(Me)
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    '???
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set gobjAnotacao = Nothing
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Anotações"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Anotacoes2"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub Unload(objme As Object)
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





Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        

        
    End If

End Sub
