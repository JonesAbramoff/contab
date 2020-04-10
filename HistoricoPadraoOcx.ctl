VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl HistoricoPadraoOcx 
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   6300
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   2250
      Picture         =   "HistoricoPadraoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   315
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "HistoricoPadraoOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "HistoricoPadraoOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "HistoricoPadraoOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "HistoricoPadraoOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox DescHistPadrao 
      Height          =   345
      Left            =   1155
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   930
      Width           =   4995
   End
   Begin VB.ListBox Historicos 
      Height          =   2010
      ItemData        =   "HistoricoPadraoOcx.ctx":0A7E
      Left            =   150
      List            =   "HistoricoPadraoOcx.ctx":0A80
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1665
      Width           =   6000
   End
   Begin MSMask.MaskEdBox HistPadrao 
      Height          =   315
      Left            =   1680
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Histórico Padrão:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   330
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   930
      Width           =   945
   End
   Begin VB.Label LblTituloTvw 
      AutoSize        =   -1  'True
      Caption         =   "Históricos"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "HistoricoPadraoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer


Const CLIENTE = "C"
Const FORNECEDOR = "F"

'--------------------------------------------------------------

Const STRING_ERROS_DESCRICAO = 255



Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iHistPadrao As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático.
    lErro = CF("HistPadrao_Automatico", iHistPadrao)
    If lErro <> SUCESSO Then Error 57514
        
    HistPadrao.Text = CStr(iHistPadrao)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57514
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161784)
    
    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objHistPadrao As New ClassHistPadrao

    objHistPadrao.iHistPadrao = colCampoValor.Item("HistPadrao").vValor

    If objHistPadrao.iHistPadrao <> 0 Then
    
        'Coloca colCampoValor na Tela
        'Conversão de tipagem para a tipagem da tela se necessário
        HistPadrao.Text = CStr(colCampoValor.Item("HistPadrao").vValor)
        DescHistPadrao.Text = colCampoValor.Item("DescHistPadrao").vValor
    
        iAlterado = 0
        
    End If
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim iHistPadrao As Integer
Dim objHistPadrao As New ClassHistPadrao
    
    'Informa tabela associada à Tela
    sTabela = "HistPadrao"
    
    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(HistPadrao.Text)) > 0 Then
        objHistPadrao.iHistPadrao = CInt(HistPadrao.Text)
    Else
        objHistPadrao.iHistPadrao = 0
    End If
    
    objHistPadrao.sDescHistPadrao = DescHistPadrao.Text
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "HistPadrao", objHistPadrao.iHistPadrao, 0, "HistPadrao"
    colCampoValor.Add "DescHistPadrao", objHistPadrao.sDescHistPadrao, STRING_HISTORICO, "DescHistPadrao"

End Sub

Private Sub Historicos_Adiciona(objHistPadrao As ClassHistPadrao)
        
Dim sEspacos As String
Dim sListBoxItem As String
    
    'Espacos para completar o tamanho STRING_CODIGO_HISTORICO
    sEspacos = Space(STRING_CODIGO_HISTORICO - Len(CStr(objHistPadrao.iHistPadrao)))
    
    'Concatena o código com a descrição do Histórico
    sListBoxItem = sEspacos & CStr(objHistPadrao.iHistPadrao) & SEPARADOR & objHistPadrao.sDescHistPadrao
    Historicos.AddItem (sListBoxItem)
    Historicos.ItemData(Historicos.NewIndex) = objHistPadrao.iHistPadrao
    
        
End Sub

Private Sub Historicos_Exclui(iHistPadrao As Integer)

Dim iIndice As Integer

    For iIndice = 0 To Historicos.ListCount - 1
    
        If Historicos.ItemData(iIndice) = iHistPadrao Then
        
            Historicos.RemoveItem (iIndice)
            Exit For
        
        End If
    
    Next

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao
Dim vbMsgRet As VbMsgBoxResult
Dim iHistPadrao As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o HistPadrao foi informado
    If Len(HistPadrao.Text) = 0 Then Error 6140
    
    objHistPadrao.iHistPadrao = CInt(HistPadrao.Text)

    'Verifica se o HistPadrao existe
    lErro = CF("HistPadrao_Le", objHistPadrao)
    If lErro <> SUCESSO And lErro <> 5446 Then Error 6142
    
    'HistPadrao não está cadastrado
    If lErro = 5446 Then Error 6141
    
    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "EXCLUSAO_HISTPADRAO")
            
    If vbMsgRet = vbYes Then
        
        'exclui o Histórico Padrão
        lErro = CF("HistPadrao_Exclui", objHistPadrao.iHistPadrao)
        If lErro <> SUCESSO Then Error 6143
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        'Exclui o Histórico da ListBox
        Call Historicos_Exclui(objHistPadrao.iHistPadrao)
    
        'Limpa a Tela
        Call Limpa_Tela(Me)
            
        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
                    
        Case 6140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_INFORMADO", Err)
            HistPadrao.SetFocus
            
        Case 6141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            HistPadrao.SetFocus
            
        Case 6142, 6143
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 161785)
        
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
    If lErro <> SUCESSO Then Error 9933
    
    'Limpa a Tela
    Call Limpa_Tela(Me)
        
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 9933
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161786)

     End Select
        
     Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao
Dim iHistPadrao As Integer
Dim iOperacao As Integer

On Error GoTo Erro_BotaoGravar_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados de Histórico foram informados
    If Len(HistPadrao.Text) = 0 Then Error 6138
    If Len(Trim(DescHistPadrao.Text)) = 0 Then Error 6176
    
    'Verifica se Descrição não começa por CARACTER_HISTPADRAO
    If left(Trim(DescHistPadrao.Text), 1) = CARACTER_HISTPADRAO Then Error 6213
    
    'Preenche objeto HistóricoPadrão
    objHistPadrao.iHistPadrao = CInt(HistPadrao.Text)
    
    'verifica se o codigo do historico não está zerado
    If objHistPadrao.iHistPadrao = 0 Then Error 55697
        
    objHistPadrao.sDescHistPadrao = Trim(DescHistPadrao.Text)
                        
    lErro = Trata_Alteracao(objHistPadrao, objHistPadrao.iHistPadrao)
    If lErro <> SUCESSO Then Error 32299
                                            
    'grava o HistPadrao no banco de dados
    lErro = CF("HistPadrao_Grava", objHistPadrao, iOperacao)
    If lErro <> SUCESSO Then Error 6139
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'remove o item da lista de historicos
    Call Historicos_Exclui(objHistPadrao.iHistPadrao)
    
    'insere o item na lista de historicos
    Call Historicos_Adiciona(objHistPadrao)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_BotaoGravar_Click:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 6138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_INFORMADO", Err)
            HistPadrao.SetFocus
            
        Case 6176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_INFORMADA", Err)
            DescHistPadrao.SetFocus
    
        Case 6139
        
        Case 6213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_COM_CARACTER_INICIAL_ERRADO", Err)
            DescHistPadrao.SetFocus
            
        Case 32299
            
        Case 55697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_ZERADO", Err)
            HistPadrao.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161787)

     End Select
        
     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
 
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17060

    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    HistPadrao.Text = ""
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 17060
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161788)

     End Select
        
     Exit Sub

End Sub

Private Sub DescHistPadrao_Change()

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
Dim colHistPadrao As New Collection
Dim sListBoxItem As String
Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_HistPadrao_Form_Load
        
    'Preenche a ListBox com Históricos Padrões existentes no BD
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 6170
    
    For Each objHistPadrao In colHistPadrao
    
        'Espaços que faltam para completar tamanho STRING_CODIGO_HISTORICO
        sListBoxItem = Space(STRING_CODIGO_HISTORICO - Len(CStr(objHistPadrao.iHistPadrao)))
        
        'Concatena Codigo e Nome do HistPadrao
        sListBoxItem = sListBoxItem & CStr(objHistPadrao.iHistPadrao)
        sListBoxItem = sListBoxItem & SEPARADOR & objHistPadrao.sDescHistPadrao
    
        Historicos.AddItem sListBoxItem
        Historicos.ItemData(Historicos.NewIndex) = objHistPadrao.iHistPadrao
        
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_HistPadrao_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 6170
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161789)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objHistPadrao As ClassHistPadrao) As Long

Dim lErro As Long
Dim sListBoxItem As String
Dim sEspacos As String
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há um HistPadrao selecionado, exibir seus dados
    If Not (objHistPadrao Is Nothing) Then
        
        'Verifica se o HistPadrao existe
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> 5446 And lErro <> SUCESSO Then Error 6126
        
        If lErro = SUCESSO Then
        
            'HistPadrao está cadastrado
            HistPadrao.Text = CStr(objHistPadrao.iHistPadrao)
            DescHistPadrao.Text = objHistPadrao.sDescHistPadrao
            
                            
        Else
        
            'Limpa a Tela
            Call Limpa_Tela(Me)
        
            'HistPadrao não está cadastrado
            HistPadrao.Text = CStr(objHistPadrao.iHistPadrao)
            
        End If
                
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 6126
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161790)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Historicos_DblClick()

Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Historicos_DblClick
    
    objHistPadrao.iHistPadrao = Historicos.ItemData(Historicos.ListIndex)
    
    'Verifica se a Histórico existe
    lErro = CF("HistPadrao_Le", objHistPadrao)
    If lErro <> 5446 And lErro <> SUCESSO Then Error 6214
    
    'Se Histórico está cadastrada
    If lErro = SUCESSO Then
            
        'Preenche campos da Tela
        HistPadrao.Text = CStr(objHistPadrao.iHistPadrao)
        DescHistPadrao.Text = objHistPadrao.sDescHistPadrao
                
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
    Else
        'Histórico não existe
    
        'Exclui da ListBox
        Historicos.RemoveItem (Historicos.ListIndex)
        
    End If
    
    iAlterado = 0
    
    Exit Sub
    
Erro_Historicos_DblClick:

    Select Case Err
            
        Case 6214
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161791)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Historicos_KeyPress(KeyAscii As Integer)

    'Se há Histórico selecionado
    If Historicos.ListIndex <> -1 Then

        If KeyAscii = ENTER_KEY Then
    
            Call Historicos_DblClick
    
        End If
        
    End If

End Sub

Private Sub HistPadrao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub HistPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HistPadrao, iAlterado)

End Sub

Private Sub HistPadrao_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_HistPadrao_Validate

    If Len(HistPadrao.Text) > 0 Then

        lErro = Long_Critica(HistPadrao.Text)
        If lErro <> SUCESSO Then Error 55694
    
    End If
    
    Exit Sub
    
Erro_HistPadrao_Validate:

    Cancel = True

    Select Case Err
    
        Case 55694
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161792)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_HISTORICO_PADRAO
    Set Form_Load_Ocx = Me
    Caption = "Histórico Padrão"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "HistoricoPadrao"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub



Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LblTituloTvw_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblTituloTvw, Source, X, Y)
End Sub

Private Sub LblTituloTvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblTituloTvw, Button, Shift, X, Y)
End Sub


'====================================================================
'Mover trecho de código para Rotinas Contab
'====================================================================


Function Extrai_CliForn_Filial(sHistorico As String, lCodigo As Long, iCodFilial As Integer, sTipo As String) As Long
'Função que extrai o código do cliente ou fornecedor e a filial passados em sHistorico

Dim lErro As Long

On Error GoTo Erro_Extrai_CliForn_Filial

    'Extrai tipo cliente ou fornecedor
    sTipo = Mid(sHistorico, 1, 1)
    'Extrai o código do cliente ou fornecedor
    lCodigo = CLng(LCodigo_Extrai(Mid(sHistorico, 2)))
    'Extrai o código da filial
    iCodFilial = StrParaInt(Nome_Extrai(sHistorico))

Extrai_CliForn_Filial = SUCESSO

    Exit Function

Erro_Extrai_CliForn_Filial:

    Extrai_CliForn_Filial = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161793)

        End Select

    Exit Function

End Function


Public Function Verifica_TituloReceber(lCliente As Long, iFilial As Integer) As Long
'Verifica se existem Titulos a receber baixados ou não

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim lNumTitulo As Long

On Error GoTo Erro_Verifica_TituloReceber

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87267

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 87268

    'Faz leitura na tabela TitulosRec, através da chave Clientes e Filial
    lErro = Comando_Executar(lComando, "SELECT NumTitulo FROM TitulosRec WHERE Cliente = ? AND Filial = ?", lNumTitulo, lCliente, iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 87269

    'Busca Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87270

    'Se não achou Titulos a Receber, proucura em Titulos a Receber baixados
    If lErro = AD_SQL_SEM_DADOS Then

        'Faz leitura na tabela TitulosRecBaixados, através da chave Clientes e Filial
        lErro = Comando_Executar(lComando1, "SELECT NumTitulo FROM TitulosRecBaixados WHERE Cliente = ? AND Filial = ?", lNumTitulo, lCliente, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 87271

        'Busca Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87272

        'Se não achou registros dispara erro
        If lErro = AD_SQL_SEM_DADOS Then gError 87273

    End If

    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

Verifica_TituloReceber = SUCESSO

    Exit Function

Erro_Verifica_TituloReceber:

    Verifica_TituloReceber = gErr

    Select Case gErr

        Case 87267, 87268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)

        Case 87269, 87270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC1", gErr, Error)

        Case 87271, 87272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS1", gErr, Error)

        Case 87273
            'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161794)

    End Select

    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function



Public Function Atualiza_MvPerCli_Trans(iFilialEmpresa As Integer, iExercicio As Integer, lCliente As Long, iFilial As Integer, dValor As Double) As Long
'Atualiza saldo em SldIni na tabela MvPerCli em Trans

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim dSldInic As Double

On Error GoTo Erro_Atualiza_MvPerCli_Trans

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87274

    lComando1 = Comando_Abrir()
    If lComando = 0 Then gError 87285

    'Faz leitura na tabela MvPerCli com exercício posterior ao encontrado na tabela de lancamentos
    lErro = Comando_ExecutarPos(lComando, "SELECT SldIni FROM MvPerCli WHERE FilialEmpresa = ? AND Exercicio = ? AND Cliente = ? AND Filial = ?", 0, dSldInic, iFilialEmpresa, iExercicio + 1, lCliente, iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 87276

    'Busca o primeiro registro que satisfaz a condição
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87277

    'Se não achou
    If lErro <> AD_SQL_SUCESSO Then

        'Insere registro no BD
        lErro = Comando_Executar(lComando, "INSERT INTO MvPerCli (SLDIni, FilialEmpresa, Exercicio, Cliente, Filial) VALUES (?,?,?,?,?)", dValor, iFilialEmpresa, iExercicio + 1, lCliente, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 87278

    Else

        'Atualiza registro no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCli SET SldIni = SldIni + ?, FilialEmpresa = ?, Exercicio = ?, Cliente = ?, Filial = ?", lComando, dValor, iFilialEmpresa, iExercicio + 1, lCliente, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 87279

    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

Atualiza_MvPerCli_Trans = SUCESSO

    Exit Function

Erro_Atualiza_MvPerCli_Trans:

    Atualiza_MvPerCli_Trans = gErr

    Select Case gErr

        Case 87274, 87285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)

        Case 87276, 87277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCLI", gErr, iFilialEmpresa, iExercicio + 1, lCliente, iFilial)

        Case 87278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCLI", gErr, iFilialEmpresa, iExercicio + 1, lCliente, iFilial)

        Case 87279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCLI", gErr, iFilialEmpresa, iExercicio + 1, lCliente, iFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161795)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Public Function Atualiza_MvPerCliForn_SldInic(iFilialEmpresa As Integer, sOrigem As String, iExercicio As Integer, iPeriodoLan As Integer, lDoc As Long) As Long
'Função que atualiza o campo SldIni da tabela MvPerForn

Dim lErro As Long
Dim lComando As Long
Dim lTransacao As Long
Dim sHistorico As String
Dim dValor As Double
Dim lCodigo As Long
Dim iCodFilial As Integer
Dim sTipo As String
Dim lCodCliente As Long

On Error GoTo Erro_Atualiza_MvPerCliForn_SldInic

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87291

    'Abre Transação
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then gError 87292

    'Inicia String
    sHistorico = String(STRING_HISTORICO, 0)

    'Faz leitura na tabela lancamentos nos campos histórico e valor
    lErro = Comando_Executar(lComando, "SELECT Historico, Valor FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ?", sHistorico, dValor, iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 87293

    'Verifica se o registro existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87294

    'Se não encontrou registros dispara erro
    If lErro = AD_SQL_SEM_DADOS Then gError 87295

    'Enquanto existirem Lançamentos
    Do While lErro = AD_SQL_SUCESSO

        'Se histórico não estiver preenchido erro
        If Len(Trim(sHistorico)) <> 0 Then

            'Extrai o código do Cliente ou Fornecedor e Filial
            lErro = Extrai_CliForn_Filial(sHistorico, lCodigo, iCodFilial, sTipo)
            If lErro <> SUCESSO Then gError 87296
    
            If sTipo = FORNECEDOR Then
    
                'Função que verifica se existe titulos a pagar cadastrados
                lErro = Verifica_TituloPagar(lCodigo, iCodFilial)
                If lErro <> SUCESSO And lErro <> 87273 Then gError 87297
    
                'Caso não sejam encontrados titulos, dispara erro
                If lErro = 87273 Then gError 87298
    
'                'Atualiza / Insere MvPerForn p/FilialEmpresa
'                lErro = Atualiza_MvPerForn_Trans(iFilialEmpresa, iExercicio, lCodigo, iCodFilial, dValor)
'                If lErro <> SUCESSO Then gError 87299
'
'                'Atualiza / Insere MvPerForn p/EMPRESA_TODA
'                lErro = Atualiza_MvPerForn_Trans(EMPRESA_TODA, iExercicio, lCodigo, iCodFilial, dValor)
'                If lErro <> SUCESSO Then gError 87301
    
            Else
                'Função que verifica se existe titulos a receber cadastrados
                lErro = Verifica_TituloReceber(lCodigo, iCodFilial)
                If lErro <> SUCESSO And lErro <> 87273 Then gError 87265
    
                'Caso não sejam encontrados titulos, dispara erro
                If lErro = 87273 Then gError 87283
    
'                'Atualiza / Insere MvPerCli p/FilialEmpresa
'                lErro = Atualiza_MvPerCli_Trans(iFilialEmpresa, iExercicio, lCodigo, iCodFilial, dValor)
'                If lErro <> SUCESSO Then gError 87266
'
'                'Atualiza / Insere MvPerCli p/EMPRESA_TODA
'                lErro = Atualiza_MvPerCli_Trans(EMPRESA_TODA, iExercicio, lCodigo, iCodFilial, dValor)
'                If lErro <> SUCESSO Then gError 87300
    
            End If
    
        End If
        
        'Busca próximo registro de acordo com chave passada no Select
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87302

    Loop

    'Confirma Transação
'    lErro = Transacao_Commit()
'    If lErro <> SUCESSO Then gError 87303

    'Fecha Comando
    Call Comando_Fechar(lComando)

Atualiza_MvPerCliForn_SldInic = SUCESSO

    Exit Function

Erro_Atualiza_MvPerCliForn_SldInic:

    Atualiza_MvPerCliForn_SldInic = gErr
    
    Select Case gErr

        Case 87283
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOREC_INEXISTENTE", gErr, lCodigo, iCodFilial)
            MsgBox ("cliente: " & CStr(lCodigo) & " filial: " & CStr(iCodFilial))
            Resume Next
            
        Case 87291
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)

        Case 87292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr, Error)

        Case 87293, 87294, 87302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr, Error)

        Case 87295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_INEXISTENTE", gErr, iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc)

        Case 87265, 87266, 87296, 87297, 87299, 87300, 87301
            'Erros já tratados na rotina

        Case 87298
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOREC_INEXISTENTE", gErr, lCodigo, iCodFilial)
            MsgBox ("fornecedor: " & CStr(lCodigo) & " filial: " & CStr(iCodFilial))
            Resume Next

        Case 87303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr, Error)

        Case 87306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTORICOFOR_NULO", gErr, lCodigo, iCodFilial, dValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161796)

    End Select

'    'Desfaz Transação
'    Call Transacao_Rollback

    'Fecha Comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Atualiza_MvPerForn_Trans(iFilialEmpresa As Integer, iExercicio As Integer, lFornecedor As Long, iFilial As Integer, dValor As Double) As Long
'Atualiza saldo em SldIni na tabela MvPerForn em Trans

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim dSldInic As Double

On Error GoTo Erro_Atualiza_MvPerForn_Trans

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87304

    lComando1 = Comando_Abrir()
    If lComando = 0 Then gError 87305

    'Faz leitura na tabela MvPerForn com exercício posterior ao encontrado na tabela de lancamentos
    lErro = Comando_ExecutarPos(lComando, "SELECT SldIni FROM MvPerForn WHERE FilialEmpresa = ? AND Exercicio = ? AND Fornecedor = ? AND Filial = ?", 0, dSldInic, iFilialEmpresa, iExercicio + 1, lFornecedor, iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 87306

    'Busca o primeiro registro que satisfaz a condição
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87307

    'Se não achou
    If lErro <> AD_SQL_SUCESSO Then

        'Insere registro no BD
        lErro = Comando_Executar(lComando, "INSERT INTO MvPerForn (SLDIni, FilialEmpresa, Exercicio, Fornecedor, Filial) VALUES (?,?,?,?,?)", dValor, iFilialEmpresa, iExercicio + 1, lFornecedor, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 87308

    Else

        'Atualiza registro no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerForn SET SldIni = SldIni + ?, FilialEmpresa = ?, Exercicio = ?, Fornecedor = ?, Filial = ?", lComando, dValor, iFilialEmpresa, iExercicio + 1, lFornecedor, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 87309

    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Atualiza_MvPerForn_Trans = SUCESSO

    Exit Function

Erro_Atualiza_MvPerForn_Trans:

    Atualiza_MvPerForn_Trans = gErr

    Select Case gErr

        Case 87304, 87305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)

        Case 87306, 87307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERFORN", gErr, iFilialEmpresa, iExercicio + 1, lFornecedor, iFilial)

        Case 87308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERFORN", gErr, iFilialEmpresa, iExercicio + 1, lFornecedor, iFilial)

        Case 87309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERFORN", gErr, iFilialEmpresa, iExercicio + 1, lFornecedor, iFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161797)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function Verifica_TituloPagar(lFornecedor As Long, iFilial As Integer) As Long
'Verifica se existem Titulos a pagar baixados ou não

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim lNumTitulo As Long

On Error GoTo Erro_Verifica_TituloPagar

    'Abre Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87267

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 87268

    'Faz leitura na tabela TitulosPag, através da chave Clientes e Filial
    lErro = Comando_Executar(lComando, "SELECT NumTitulo FROM TitulosPag WHERE Fornecedor = ? AND Filial = ?", lNumTitulo, lFornecedor, iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 87269

    'Busca Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87270

    'Se não achou Titulos a Pagar, proucura em Titulos a Pagar baixados
    If lErro = AD_SQL_SEM_DADOS Then

        'Faz leitura na tabela TitulosPagBaixados, através da chave Clientes e Filial
        lErro = Comando_Executar(lComando1, "SELECT NumTitulo FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ?", lNumTitulo, lFornecedor, iFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 87271

        'Busca Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87272

        'Se não achou registros dispara erro
        If lErro = AD_SQL_SEM_DADOS Then gError 87273

    End If

    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

Verifica_TituloPagar = SUCESSO

    Exit Function

Erro_Verifica_TituloPagar:

    Verifica_TituloPagar = gErr

    Select Case gErr

        Case 87267, 87268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr, Error)

        Case 87269, 87270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG1", gErr, Error) 'WW Verificar parametros da constante CRFAT

        Case 87271, 87272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGBAIXADOS1", gErr, Error) 'WW Verificar parametros da constante CRFAT

        Case 87273
            'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161798)

    End Select

    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

'Fernando copiado da tela de Transferencia
Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)

    Nome_Extrai = sString

    Exit Function

End Function

'====================================================================

'??? Jones

'Autor: William
'Data de Inicio: 27/03/2001
'Data de Término: 27/03/2001
'Função: RotinaErro_Formata
'Descrição: Recebe linha de comando da função Rotina_Erro
'           e retorna a mesma formata com aspas no segundo
'           parâmetro da função (onde é especificada a Constante)


Public Function RotinaErro_Formata(sRotinaErro As String) As String
'Retorna constante da função passada como parâmetro entre aspas

Dim iPosConst1 As Integer
Dim iPosConst2 As Integer
Dim sInicString As String
Dim sFinString As String
Dim sConstante As String
Dim sAspas As String
Dim lErro As Long

On Error GoTo Erro_RotinaErro_Formata

    'Verifica se função foi passada em sRotinaErro
    If Len(Trim(sRotinaErro)) <> 0 Then
    
        'Atribui aspas a variável, onde as mesma serão concatenadas + a diante
        sAspas = """"
            
        'Retira posição da primeira virgula(delimita inicio do segundo parâmetro)
        iPosConst1 = InStr(1, sRotinaErro, ",")
        
        'Retira posição da segunda virgula(delimita final do segundo parâmetro)
        iPosConst2 = InStr(iPosConst1 + 1, sRotinaErro, ",")
        
        'Extrai a parte inicial do texto para + adiante ser concatenado
        sInicString = left(sRotinaErro, iPosConst1)
        
        'Extrai a parte final do texto para + adiante ser concatenado
        sFinString = Mid(sRotinaErro, iPosConst2)
        
        'Extrai apenas a Constante (segundo Parametro da função)
        sConstante = Mid(sRotinaErro, iPosConst1 + 1, iPosConst2 - (iPosConst1 + 1))
        
        'Atribui Aspas a Constante (segundo Parametro da função)
        sConstante = sAspas + sConstante + sAspas
        
        'Monta a chamada a função devidamente formatada
        sRotinaErro = sInicString + sConstante + sFinString
        
        'Atribui o valor de retorno da função
        RotinaErro_Formata = sRotinaErro
        
    End If
    
RotinaErro_Formata = SUCESSO

    Exit Function
    
Erro_RotinaErro_Formata:

    RotinaErro_Formata = gErr
    
    Select Case gErr
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161799)
        
    End Select
    
    Exit Function
    
    
End Function

'========================================================================
'??? Jones

'Autor: William
'Data de Inicio: 28/03/01
'Data de Término: 28/03/01
'Função: Gravar_ErroBD
'Descrição: le arquivo texto com constantes de erro e descrição, e inclui
'           registros na tabela de Erros

'obs.: função subdividida em Arq_ConstErros_Le, Formata_ConstErros, Erros_Insere_Atualiza


Function Gravar_ErroBD() As Long
'Chama funções responsáveis pela gravação dos erros

Dim colArquivo As New Collection
Dim lErro As Long

On Error GoTo Erro_Gravar_ErroBD

    'Le arquivo de erros
    lErro = Arq_ConstErros_Le(colArquivo)
    If lErro <> SUCESSO Then gError 87522
    
    'Insere / Atualiza tabela no BD
    lErro = Erros_Insere_Atualiza(colArquivo)
    If lErro <> SUCESSO Then gError 87521
    
Gravar_ErroBD = SUCESSO

    Exit Function
    
Erro_Gravar_ErroBD:

    Gravar_ErroBD = gErr
    
    Select Case gErr
    
        Case 87521, 87522
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161800)
        
    End Select
    
    Exit Function

End Function


Public Function Arq_ConstErros_Le(colArquivo As Collection) As Long
'Extrai dados do arquivo ConstErros

Dim sArquivo As String
Dim lErro As Long

On Error GoTo Erro_Arq_ConstErros_Le

    'Abre arquivo
    Open "C:\CONTAB\CONSTERROS.TXT" For Input As #1
    
    'Para cada linha de ConstErros atribuir a sArquivo
    Do While Not EOF(1)
        Line Input #1, sArquivo
        If Len(Trim(sArquivo)) > 0 Then
            colArquivo.Add sArquivo
        End If
    Loop
    
    'Fecha arquivo
    Close #1
    
Arq_ConstErros_Le = SUCESSO

    Exit Function
    
Erro_Arq_ConstErros_Le:

    Arq_ConstErros_Le = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161801)
            
    End Select
    
    'Fecha Arquivo - Saida por erro
    Close #1
    
    Exit Function
    
End Function

Function Erros_Insere_Atualiza(colArquivo As Collection) As Long
'Grava Erro no BD

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sCodigo As String
Dim sDescricao As String
Dim sDescricao1 As String
Dim iIndice As Integer
Dim sArquivo As String

On Error GoTo Erro_Erros_Insere_Atualiza

    'Abre Comandos
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 87513
           
    lComando1 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando1 = 0 Then gError 87514
                      
    'Abre Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then gError 87515

    'Para cada registro em colArquivo
    For iIndice = 1 To colArquivo.Count
        
        'Atribui registro a sArquivo
        sArquivo = colArquivo.Item(iIndice)
                                        
        'Extrai o código e a descrição para Variaveis separadas
        lErro = Formata_ConstErros(sArquivo, sCodigo, sDescricao)
        If lErro <> SUCESSO Then gError 87253
                    
        'Inicia String
        sDescricao1 = String(STRING_ERROS_DESCRICAO, 0)
              
        'Le tabela erros por chave - codigo
        lErro = Comando_ExecutarPos(lComando, "SELECT Descricao FROM Erros WHERE Codigo = ?", 0, sDescricao1, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 87516
    
        'busca o primeiro registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87517
        
        'Se não encontrou
        If lErro <> AD_SQL_SUCESSO Then
                
            '===> Insere
            lErro = Comando_Executar(lComando, "INSERT INTO Erros (Codigo, Descricao) VALUES (?,?)", sCodigo, sDescricao)
            If lErro <> AD_SQL_SUCESSO Then gError 87518
            
        ''Else
            
            '===> Atualiza
            ''If Trim(sDescricao) <> Trim(sDescricao1) Then
                       
                ''lErro = Comando_ExecutarPos(lComando1, "UPDATE Erros SET Descricao = ?", lComando, sDescricao)
                ''If lErro <> AD_SQL_SUCESSO Then gError 87519
            
            ''End If
            
        End If
            
    Next

    'Confirma transacao
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 87520

    'Fecha Comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

Erros_Insere_Atualiza = SUCESSO

    Exit Function
    
Erro_Erros_Insere_Atualiza:

    Erros_Insere_Atualiza = gErr

    Select Case gErr
    
        Case 87513, 87514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 87515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 87516, 87517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ERROS", gErr)

        Case 87518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ERROS", gErr, sCodigo)

        Case 87519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ERROS", gErr, sCodigo)

        Case 87520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case 87253

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161802)

    End Select
        
    'Desfaz transação
    Call Transacao_RollbackDic
    
    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
        
    Exit Function
        
End Function


Public Function Formata_ConstErros(sArquivo As String, sCodigo As String, sDescricao As String) As Long
'Separa sArquivo para que possa ser gravado no BD

Dim iPosConst1 As Integer
Dim iPosConst2 As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_ConstErros

    If Len(Trim(sArquivo)) > 0 Then
                
        'Retira posição da primeira Aspa(delimita inicio da Descrição)
        iPosConst1 = InStr(1, sArquivo, """")
        
        'Retira posição da segunda Aspa(delimita final da Descrição)
        iPosConst2 = InStr(iPosConst1 + 1, sArquivo, """")
        
        'Extrai o código
        sCodigo = left(sArquivo, iPosConst1 - 1)
                
        'Extrai a Descrição
        sDescricao = Mid(sArquivo, iPosConst1 + 1, iPosConst2 - (iPosConst1 + 1))
                
    End If

Formata_ConstErros = SUCESSO

    Exit Function
    
Erro_Formata_ConstErros:

    Formata_ConstErros = gErr

    Select Case gErr
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161803)
        
    End Select
    
    Exit Function

End Function

