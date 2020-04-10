VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl MensagensOcx 
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   6270
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1785
      Picture         =   "MensagensOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   315
      Width           =   300
   End
   Begin VB.ListBox Mensagens 
      Height          =   2010
      ItemData        =   "MensagensOcx.ctx":00EA
      Left            =   135
      List            =   "MensagensOcx.ctx":00EC
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1665
      Width           =   6000
   End
   Begin VB.TextBox Descricao 
      Height          =   345
      Left            =   1215
      MaxLength       =   250
      TabIndex        =   2
      Top             =   930
      Width           =   4920
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "MensagensOcx.ctx":00EE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "MensagensOcx.ctx":0248
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "MensagensOcx.ctx":03D2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "MensagensOcx.ctx":0904
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   300
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mensagens"
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
      Left            =   180
      TabIndex        =   9
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem:"
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
      Left            =   135
      TabIndex        =   10
      Top             =   975
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   480
      TabIndex        =   11
      Top             =   330
      Width           =   675
   End
End
Attribute VB_Name = "MensagensOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico","CPRConfig", NUM_PROX_MENSAGEM, "Mensagens", "Codigo", lNumAuto)
    If lErro <> SUCESSO Then Error 57550
    
    Codigo.Text = CStr(lNumAuto)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57550
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162715)
    
    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim objMensagem As New ClassMensagem

    objMensagem.iCodigo = colCampoValor.Item("Codigo").vValor

    If objMensagem.iCodigo <> 0 Then
    
        'Coloca colCampoValor na Tela
        'Conversão de tipagem para a tipagem da tela se necessário
        Codigo.Text = CStr(colCampoValor.Item("Codigo").vValor)
        Descricao.Text = colCampoValor.Item("Descricao").vValor
    
    End If
    
    iAlterado = 0
    
End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim iCodigo As Integer
Dim objMensagem As New ClassMensagem
    
    'Informa tabela associada à Tela
    sTabela = "Mensagens"
    
    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Codigo.Text) > 0 Then
        objMensagem.iCodigo = CInt(Codigo.Text)
    Else
        objMensagem.iCodigo = 0
    End If
    
       objMensagem.sDescricao = Descricao.Text
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objMensagem.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objMensagem.sDescricao, STRING_NFISCAL_MENSAGEM, "Descricao"

End Sub

Private Sub Mensagens_Adiciona(objMensagem As ClassMensagem)
        
Dim sEspacos As String
Dim sListBoxItem As String
    
    'Espacos para completar o tamanho STRING_CODIGO_MENSAGEM
    sEspacos = Space(STRING_CODIGO_MENSAGEM - Len(CStr(objMensagem.iCodigo)))
    
    'Concatena o código com a descrição da Mensagem
    sListBoxItem = sEspacos & CStr(objMensagem.iCodigo) & SEPARADOR & objMensagem.sDescricao
    Mensagens.AddItem (sListBoxItem)
    Mensagens.ItemData(Mensagens.NewIndex) = objMensagem.iCodigo
        
End Sub

Private Sub Mensagens_Exclui(iCodigo As Integer)
'Exclui da ListBox

Dim iIndice As Integer

    For iIndice = 0 To Mensagens.ListCount - 1
    
        If Mensagens.ItemData(iIndice) = iCodigo Then
            
            Mensagens.RemoveItem (iIndice)
            Exit For
            
        End If
    
    Next

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objMensagem As New ClassMensagem
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a Mensagem foi informada
    If Len(Codigo.Text) = 0 Then Error 39697
    
    objMensagem.iCodigo = CInt(Codigo.Text)

    'Verifica se a Mensagem existe
    lErro = CF("Mensagem_Le",objMensagem)
    If lErro <> SUCESSO And lErro <> 19234 Then Error 39699
    
    'Mensagem não está cadastrada
    If lErro = 19234 Then Error 39700
    
    'Pede confirmação para exclusão ao usuário
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_MENSAGEM")
            
    If vbMsgRes = vbYes Then
        
        'exclui a Mensagem
        lErro = CF("Mensagem_Exclui",objMensagem.iCodigo)
        If lErro Then Error 39701
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        'Exclui a Mensagem da ListBox
        Call Mensagens_Exclui(objMensagem.iCodigo)
    
        'Limpa a Tela
        Call Limpa_Tela(Me)
            
        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
                    
        Case 39697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_INFORMADA", Err)
            
        Case 39700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", Err, objMensagem.iCodigo)
            
        Case 39699, 39701
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 162716)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 39723
    
    'Limpa a Tela
    Call Limpa_Tela(Me)
        
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 39723
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162717)

     End Select
        
     Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objMensagem As New ClassMensagem
Dim iCodigo As Integer
Dim iOperacao As Integer 'Dúvida

On Error GoTo Erro_BotaoGravar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados da Mensagem foram informados
    If Len(Codigo.Text) = 0 Then Error 39725
    If Len(Trim(Descricao.Text)) = 0 Then Error 39726
    
    'Verifica se  não começa por CARACTER_MENSAGEM
    If Left(Trim(Descricao.Text), 1) = CARACTER_MENSAGEM Then Error 39727

    'Preenche objeto Mensagem
    objMensagem.iCodigo = CInt(Codigo.Text)
    objMensagem.sDescricao = Trim(Descricao.Text)
        
    lErro = Trata_Alteracao(objMensagem, objMensagem.iCodigo)
    If lErro <> SUCESSO Then Error 32333
        
    'Grava a Mensagem no banco de dados
    lErro = CF("Mensagem_Grava",objMensagem, iOperacao) 'Dúvida
    If lErro <> SUCESSO Then Error 39728
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'remove o item da lista de Mensagens
    Call Mensagens_Exclui(objMensagem.iCodigo)
    
    'insere o item na lista de Mensagens
    Call Mensagens_Adiciona(objMensagem)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_BotaoGravar_Click:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 32333
    
        Case 39725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO", Err)
            Codigo.SetFocus
            
        Case 39726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_INFORMADA", Err)
            Descricao.SetFocus
            
        Case 39727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_COM_CARACTER_INICIAL_ERRADO", Err)
            Descricao.SetFocus
        
        Case 39728
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162718)

     End Select
        
     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
 
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 39729

    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 39729
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162719)

     End Select
        
     Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 57974
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 57974 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162720)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colMensagem As New Collection
Dim sListBoxItem As String
Dim objMensagem As ClassMensagem

On Error GoTo Erro_Mensagem_Form_Load
        
    'Preenche a ListBox com Mensagens existentes no BD
    lErro = CF("Mensagem_Le_Todas",colMensagem)
    If lErro <> SUCESSO Then Error 39731
    
    For Each objMensagem In colMensagem
    
        'Espaços que faltam para completar tamanho STRING_CODIGO_MENSAGEM
        sListBoxItem = Space(STRING_CODIGO_MENSAGEM - Len(CStr(objMensagem.iCodigo)))
        
        'Concatena Codigo e Descricao da Mensagem
        sListBoxItem = sListBoxItem & CStr(objMensagem.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & objMensagem.sDescricao
    
        Mensagens.AddItem sListBoxItem
        Mensagens.ItemData(Mensagens.NewIndex) = objMensagem.iCodigo
        
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Mensagem_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 39731
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162721)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objMensagem As ClassMensagem) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim sEspacos As String
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma Mensagem selecionada, exibir seus dados
    If Not (objMensagem Is Nothing) Then
        
        'Verifica se a Mensagem existe
        lErro = CF("Mensagem_Le",objMensagem)
        If lErro <> 19234 And lErro <> SUCESSO Then Error 39732
        
        If lErro = SUCESSO Then
        
            'Mensagem está cadastrada
            Codigo.Text = CStr(objMensagem.iCodigo)
            Descricao.Text = objMensagem.sDescricao
            
        Else
        
            'Limpa a Tela
            Call Limpa_Tela(Me)
        
            'Mensagem não está cadastrada
            Codigo.Text = CStr(objMensagem.iCodigo)
            
        End If
                
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 39732
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 162722)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Mensagens_DblClick()

Dim lErro As Long
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Mensagens_DblClick
    
    objMensagem.iCodigo = Mensagens.ItemData(Mensagens.ListIndex)
    
    'Verifica se a Mensagem existe
    lErro = CF("Mensagem_Le",objMensagem)
    If lErro <> 19234 And lErro <> SUCESSO Then Error 39734
    
    If lErro = SUCESSO Then 'Mensagem está cadastrada
            
        'Preenche campos da Tela
        Codigo.Text = CStr(objMensagem.iCodigo)
        Descricao.Text = objMensagem.sDescricao
                
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
    Else 'Mensagem não existe
    
        'Exclui da ListBox
        Mensagens.RemoveItem (Mensagens.ListIndex)
        
    End If
    
    iAlterado = 0
    
    Exit Sub
    
Erro_Mensagens_DblClick:

    Select Case Err
            
        Case 39734
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 162723)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Mensagens_KeyPress(KeyAscii As Integer)

    'Se há Mensagem selecionada
    If Mensagens.ListIndex <> -1 Then

        If KeyAscii = ENTER_KEY Then
    
            Call Mensagens_DblClick
    
        End If
        
    End If

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_MENSAGEM
    Set Form_Load_Ocx = Me
    Caption = "Mensagens"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Mensagens"
    
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
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub



Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

