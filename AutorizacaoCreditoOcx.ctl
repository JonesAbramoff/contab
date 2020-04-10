VERSION 5.00
Begin VB.UserControl AutorizacaoCreditoOcx 
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   LockControls    =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   3630
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
      Left            =   690
      Picture         =   "AutorizacaoCreditoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton BotaoCancela 
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
      Left            =   2115
      Picture         =   "AutorizacaoCreditoOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crédito a ser concedido"
      Height          =   1440
      Left            =   195
      TabIndex        =   5
      Top             =   150
      Width           =   3255
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   6
         Top             =   435
         Width           =   660
      End
      Begin VB.Label LabelCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   390
         Width           =   1980
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
         Height          =   195
         Left            =   435
         TabIndex        =   8
         Top             =   975
         Width           =   510
      End
      Begin VB.Label LabelValor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1065
         TabIndex        =   9
         Top             =   930
         Width           =   1980
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Responsável pela autorização"
      Height          =   1425
      Left            =   180
      TabIndex        =   4
      Top             =   1800
      Width           =   3270
      Begin VB.TextBox TextSenha 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   930
         Width           =   2010
      End
      Begin VB.ComboBox ComboUsuario 
         Height          =   315
         Left            =   1050
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   330
         TabIndex        =   10
         Top             =   990
         Width           =   615
      End
      Begin VB.Label UsuariosLabel 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   405
         TabIndex        =   11
         Top             =   435
         Width           =   555
      End
   End
End
Attribute VB_Name = "AutorizacaoCreditoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objAutorizacaoCredito As ClassAutorizacaoCredito

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objUsuarios As New ClassUsuarios
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim dValorCreditoSolicitado As Double
Dim objValorLiberadoCredito As New ClassValorLiberadoCredito

On Error GoTo Erro_BotaoOK_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Usuario foi Preenchido
    If Len(ComboUsuario.Text) = 0 Then Error 44434
    
    'Verifica se digitou a senha
    If Len(TextSenha.Text) = 0 Then Error 44435
    
    objUsuarios.sCodUsuario = ComboUsuario.Text

    'le os dados do usuário
    lErro = CF("Usuarios_Le",objUsuarios)
    If lErro <> SUCESSO And lErro <> 44433 Then Error 44436
    
    'se o usuário não está cadastrado ==> erro.
    If lErro = 44433 Then Error 44437
    
    'se a senha não for a que está cadastada ==> erro
    If TextSenha.Text <> objUsuarios.sSenha Then Error 44438
    
    If giTipoVersao = VERSAO_FULL Then
    
        objLiberacaoCredito.sCodUsuario = objUsuarios.sCodUsuario
        
        'Lê a liberacao de credito a partir do código do usuario.
        lErro = CF("LiberacaoCredito_Le",objLiberacaoCredito)
        If lErro <> SUCESSO And lErro <> 36968 Then Error 44440
        
        'se não foi encontrado autorização para o usuario liberar credito
        If lErro = 36968 Then Error 44441
            
        dValorCreditoSolicitado = CDbl(LabelValor.Caption)
            
        'se o valor do crédito solicitado ultrapassar o limite de credito que o usuario pode conceder por operacao
        If dValorCreditoSolicitado > objLiberacaoCredito.dLimiteOperacao Then Error 44442
            
        objValorLiberadoCredito.sCodUsuario = objUsuarios.sCodUsuario
        objValorLiberadoCredito.iAno = Year(gdtDataAtual)
            
        'Lê a estatistica de liberação de credito de um usuario em um determinado ano
        lErro = CF("ValorLiberadoCredito_Le",objValorLiberadoCredito)
        If lErro <> SUCESSO And lErro <> 36973 Then Error 44443
            
        'se o valor do pedido ultrapassar o valor mensal que o usuario tem capacidade de liberar
        If dValorCreditoSolicitado > objLiberacaoCredito.dLimiteMensal - objValorLiberadoCredito.adValorLiberado(Month(gdtDataAtual)) Then Error 44445
    
    End If
    
    objAutorizacaoCredito.iCreditoAutorizado = CREDITO_APROVADO
    objAutorizacaoCredito.sCodUsuario = ComboUsuario.Text
        
    GL_objMDIForm.MousePointer = vbDefault
     
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 44434
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)
    
        Case 44435
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", Err)
    
        Case 44436, 44440, 44443
    
        Case 44437
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, objUsuarios.sCodUsuario)
    
        Case 44438
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_INVALIDA", Err)
    
        Case 44441
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_INEXISTENTE", Err, objLiberacaoCredito.sCodUsuario)
    
        Case 44442
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEOPERACAO", Err, objLiberacaoCredito.sCodUsuario)
    
        Case 44445
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEMENSAL", Err, objLiberacaoCredito.sCodUsuario)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143182)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 143183)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objAutorizacaoCredito = Nothing
    
End Sub

Function Trata_Parametros(Optional objAutorizacaoCredito1 As ClassAutorizacaoCredito) As Long

Dim lErro As Long
Dim iIndice As Integer, objCliente As New ClassCliente
Dim colUsuarios As New Collection
Dim objUsuarios As ClassUsuarios
Dim colUsuariosComLiberacao As New Collection
Dim objUsuariosComLiberacao As ClassUsuarios

On Error GoTo Erro_Trata_Parametros

    If IsMissing(objAutorizacaoCredito1) Then Error 44429
    
    'Preenche Cliente e VAlor
    objCliente.lCodigo = objAutorizacaoCredito1.lCliente
    lErro = CF("Cliente_Le",objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then Error 59357
    If lErro <> SUCESSO Then Error 59358
    
    LabelCliente.Caption = objAutorizacaoCredito1.lCliente & SEPARADOR & objCliente.sNomeReduzido
    LabelValor.Caption = Format(objAutorizacaoCredito1.dValor, "Standard")
    
    'Preenche a Combo com Usuarios que tem alçada Superior ao Valor da Operacao
    
    'Le todos os usuarios para esta Filial Empresa
    lErro = CF("UsuariosFilialEmpresa_Le_Todos",colUsuarios)
    If lErro <> SUCESSO Then Error 44446
    
    If giTipoVersao = VERSAO_FULL Then
        'Le todos os Usuarios que tem Liberacao de Crédito por Operacao e Mensal
        lErro = CF("Usuarios_Com_LiberacaoCredito_Le",colUsuarios, objAutorizacaoCredito1.dValor, colUsuariosComLiberacao)
        If lErro <> SUCESSO Then Error 58581
    
        'Preenche a combo de Usuarios
        For Each objUsuariosComLiberacao In colUsuariosComLiberacao
            ComboUsuario.AddItem objUsuariosComLiberacao.sCodUsuario
        Next
        
    ElseIf giTipoVersao = VERSAO_LIGHT Then
        
        'Preenche a combo de Usuarios
        For Each objUsuarios In colUsuarios
            ComboUsuario.AddItem objUsuarios.sCodUsuario
        Next
    
    End If
    
    'Seta o objeto global a Tela (objAutorizacaoCredito)
    Set objAutorizacaoCredito = objAutorizacaoCredito1
    
    objAutorizacaoCredito.iCreditoAutorizado = CREDITO_RECUSADO
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 44446, 59357, 58581 'Tratados nas Rotinas Chamadas
    
        Case 59358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objCliente.lCodigo)
        
        Case 44429
            lErro = Rotina_Erro(vbOKOnly, "TELA_AUTCRED_CHAMADA_SEM_PARAMETRO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143184)
    
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_AUTORIZACAO_CREDITO
    Set Form_Load_Ocx = Me
    Caption = "Autorização de Crédito"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AutorizacaoCredito"
    
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

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelValor, Source, X, Y)
End Sub

Private Sub LabelValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelValor, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub UsuariosLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UsuariosLabel, Source, X, Y)
End Sub

Private Sub UsuariosLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UsuariosLabel, Button, Shift, X, Y)
End Sub

