VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl HistMovCtaOcx 
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   6630
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1905
      Picture         =   "HistMovCtaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   270
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4365
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "HistMovCtaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "HistMovCtaOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "HistMovCtaOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "HistMovCtaOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox Historicos 
      Height          =   1620
      ItemData        =   "HistMovCtaOcx.ctx":0A7E
      Left            =   120
      List            =   "HistMovCtaOcx.ctx":0A80
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox DescHistPadrao 
      Height          =   315
      Left            =   1230
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   990
      Width           =   5235
   End
   Begin MSMask.MaskEdBox HistPadrao 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   255
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Height          =   195
      Left            =   120
      Top             =   1560
      Width           =   885
      ForeColor       =   0
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
      Caption         =   "Históricos"
   End
   Begin VB.Label Label2 
      Height          =   195
      Left            =   210
      Top             =   1050
      Width           =   945
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
      Caption         =   "Descrição:"
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   480
      Top             =   300
      Width           =   675
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
      Caption         =   "Código:"
   End
End
Attribute VB_Name = "HistMovCtaOcx"
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
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera Código de Histórico automático.
    lErro = CF("HistMovCta_Automatico",iCodigo)
    If lErro <> SUCESSO Then Error 57549

    HistPadrao.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57549
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161773)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objHistMovCta As New ClassHistMovCta
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código do Histórico foi informado
    If Len(HistPadrao.Text) = 0 Then Error 15045

    objHistMovCta.iCodigo = CInt(HistPadrao.Text)

    'Verifica se o Código do Histórico existe
    lErro = CF("HistMovCta_Le",objHistMovCta)
    If lErro <> SUCESSO And lErro <> 15011 Then Error 15047

    'Histórico não está cadastrado
    If lErro = 15011 Then Error 15046

    'Confirma a exclusão
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_HISTMOVCTA", objHistMovCta.iCodigo)

    If vbMsgRet = vbYes Then

        'Exclui o Histórico
        lErro = CF("HistMovCta_Exclui",objHistMovCta.iCodigo)
        If lErro <> SUCESSO Then Error 15050

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        'Exclui o Histórico da ListBox
        Call ListaHistoricos_Exclui(HistPadrao.Text)

        'Limpa a Tela
        Call Limpa_Tela(Me)
        
        HistPadrao.Text = ""

        'Zerar iAlterado
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_INFORMADO", Err)
            HistPadrao.SetFocus

        Case 15046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)

        Case 15047, 15050

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161774)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava HistMovCta
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 15205

    'Limpa a tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 15205

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161775)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alteração e confirma se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 15063

    'Limpa a Tela
    Call Limpa_Tela(Me)

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Zerar iAlterado
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 15063

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161776)

     End Select

     Exit Sub

End Sub

Private Sub DescHistPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sListBoxItem As String
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodNome As AdmCodigoNome

On Error GoTo Erro_HistMovCta_Form_Load

    'Coloca todos os Históricos Padrão na coleção
    lErro = CF("Cod_Nomes_Le","HistPadraoMovConta", "Codigo", "Descricao", STRING_HISTORICO, colCodigoNome)
    If lErro <> SUCESSO Then Error 15000

    'Preenche a ListBox com Históricos existentes na coleção
    For Each objCodNome In colCodigoNome

        'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
        sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(CStr(objCodNome.iCodigo)))

        'Concatena Código e Descrição do Histórico
        sListBoxItem = sListBoxItem & CStr(objCodNome.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & Trim(objCodNome.sNome)

        Historicos.AddItem sListBoxItem
        Historicos.ItemData(Historicos.NewIndex) = objCodNome.iCodigo

    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_HistMovCta_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 15000

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161777)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objHistMovCta As ClassHistMovCta) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um Histórico selecionado
    If Not (objHistMovCta Is Nothing) Then

        'Verifica se o Histórico existe
        lErro = CF("HistMovCta_Le",objHistMovCta)
        If lErro <> 15011 And lErro <> SUCESSO Then Error 15002

        'Se Histórico está cadastrado
        If lErro = SUCESSO Then

            'Mantém o Código do Histórico na tela e adiciona a Descrição
            HistPadrao.Text = CStr(objHistMovCta.iCodigo)
            DescHistPadrao.Text = objHistMovCta.sDescricao

            'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
            sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(HistPadrao.Text))

            'Concatena para comparar com ítens da ListBox
            sListBoxItem = sListBoxItem & HistPadrao.Text & SEPARADOR & DescHistPadrao.Text

            'Seleciona Histórico na ListBox
            For iIndice = 0 To Historicos.ListCount - 1

                If Historicos.List(iIndice) = sListBoxItem Then
                    Historicos.ListIndex = iIndice
                    Exit For
                End If

            Next

        'Se Histórico não está cadastrado
        Else

            Call Limpa_Tela(Me)

            'Mantém o Código do Histórico na tela
            HistPadrao.Text = CStr(objHistMovCta.iCodigo)

        End If

    End If

    'Zerar iAlterado
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 15002

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161778)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Private Sub Historicos_DblClick()

Dim lErro As Long
Dim sCodigo As String
Dim sListBoxItem As String
Dim objHistMovCta As New ClassHistMovCta
Dim lSeparadorPosicao As Long

    'Pega a String do ítem selecionado
    sListBoxItem = Historicos.List(Historicos.ListIndex)

    'Acha a posição do separador (-)
    lSeparadorPosicao = InStr(sListBoxItem, SEPARADOR)

    'Preenche Código e Descrição do Histórico na Tela
    HistPadrao.Text = Trim(Left(sListBoxItem, lSeparadorPosicao - 1))
    DescHistPadrao.Text = Mid(sListBoxItem, lSeparadorPosicao + 1)
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Historicos_KeyPress(KeyAscii As Integer)

    'Se não tiver nenhum ítem selecionado na lista
    If Historicos.ListIndex = -1 Then Exit Sub

    'Se a tecla pressionada for Enter
    If KeyAscii = ENTER_KEY Then

        'Executa o mesmo procedimento que o duplo click
        Call Historicos_DblClick

    End If

End Sub

Private Sub HistPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Gravar_Registro() As Long
'Grava HistMovCta

Dim lErro As Long
Dim objHistMovCta As New ClassHistMovCta
Dim iCodigo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados do Histórico foram informados
    If Len(HistPadrao.Text) = 0 Then Error 15028
    If Len(Trim(DescHistPadrao.Text)) = 0 Then Error 15029

    'Verifica se a Descrição do Histórico não começa por CARACTER_HISTPADRAO
    If Left(Trim(DescHistPadrao.Text), 1) = CARACTER_HISTPADRAO Then Error 15030

    'Preenche objeto objHistMovCta
    objHistMovCta.iCodigo = CInt(HistPadrao.Text)
    objHistMovCta.sDescricao = Trim(DescHistPadrao.Text)

    'Grava o Histórico no Banco de Dados
    lErro = CF("HistMovCta_Grava",objHistMovCta)
    If lErro <> SUCESSO Then Error 15031

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Remove o ítem da lista de históricos, se já existir
    Call ListaHistoricos_Exclui(objHistMovCta.iCodigo)

    'Insere o ítem na lista de históricos
    Call ListaHistoricos_Adiciona(objHistMovCta)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 15028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_INFORMADO", Err)

        Case 15029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_INFORMADA", Err)

        Case 15030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_COM_CARACTER_INICIAL_ERRADO", Err)

        Case 15031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161779)

     End Select

     Exit Function

End Function

Private Sub ListaHistoricos_Adiciona(objHistMovCta As ClassHistMovCta)
'Adiciona ítem na ListBox Historicos

Dim sListBoxItem As String

    'Espacos para completar o tamanho STRING_CODIGO_HISTORICO
    sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(CStr(objHistMovCta.iCodigo)))

    'Concatena o código com a descrição do Histórico
    sListBoxItem = sListBoxItem & CStr(objHistMovCta.iCodigo) & SEPARADOR & objHistMovCta.sDescricao

    'Adiciona o ítem na ListBox
    Historicos.AddItem (sListBoxItem)
    Historicos.ItemData(Historicos.NewIndex) = objHistMovCta.iCodigo

End Sub

Private Sub ListaHistoricos_Exclui(iCodigo As Integer)
'Exclui ítem da ListBox Historicos

Dim iIndice As Integer

    'Percorre todos os itens da ListBox
    For iIndice = 0 To Historicos.ListCount - 1

        'Se o ItemData do ítem for igual ao Código passado em iCodigo
        If Historicos.ItemData(iIndice) = iCodigo Then

            'Remove o ítem
            Historicos.RemoveItem (iIndice)
            Exit For

        End If

    Next

End Sub

Public Function Traz_Hist_Tela(objHistMovCta As ClassHistMovCta) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Hist_Tela

    HistPadrao.Text = CStr(objHistMovCta.iCodigo)
    DescHistPadrao.Text = objHistMovCta.sDescricao
    
    iAlterado = 0
    
    Traz_Hist_Tela = SUCESSO
    
    Exit Function

Erro_Traz_Hist_Tela:

    Traz_Hist_Tela = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161780)
    
    End Select
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Tela_Extrai

    sTabela = "HistPadraoMovConta"
    
    If Len(Trim(HistPadrao.ClipText)) > 0 Then
        objHistMovCta.iCodigo = CInt(HistPadrao.Text)
    Else
        objHistMovCta.iCodigo = 0
    End If

    If Len(DescHistPadrao.Text) > 0 Then
        objHistMovCta.sDescricao = DescHistPadrao.Text
    Else
        objHistMovCta.sDescricao = String(STRING_HISTORICO, 0)
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'Valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objHistMovCta.iCodigo, 0, "Codigo"
    colCampoValor.Add "Descricao", objHistMovCta.sDescricao, STRING_HISTORICO, "Descricao"
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161781)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Tela_Preenche

    objHistMovCta.iCodigo = colCampoValor.Item("Codigo").vValor

    If objHistMovCta.iCodigo <> 0 Then
        
        objHistMovCta.sDescricao = colCampoValor.Item("Descricao").vValor
    
        lErro = Traz_Hist_Tela(objHistMovCta)
        If lErro <> SUCESSO Then Error 34667
    
    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case Err
    
        Case 34667
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161782)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub HistPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HistPadrao, iAlterado)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_HISTORICO_EXTRATO_CONTA_CORRENTE
    Set Form_Load_Ocx = Me
    Caption = "Históricos para o Extrato de Conta Corrente"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "HistMovCta"
    
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

Private Sub HistPadrao_Validate(Cancel As Boolean)

On Error GoTo Erro_HistPadrao_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(HistPadrao.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(HistPadrao.Text) Then Error 55960

        'Verifica se codigo é menor que um
        If CInt(HistPadrao.Text) < 1 Then Error 55961

    End If

    Exit Sub

Erro_HistPadrao_Validate:

    Cancel = True

    Select Case Err

        Case 55960, 55961
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO1", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161783)

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

