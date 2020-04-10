VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ContaCategoriaOcx 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6270
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1500
      Picture         =   "ContaCategoriaOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   225
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ContaCategoriaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ContaCategoriaOcx.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ContaCategoriaOcx.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ContaCategoriaOcx.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListaCategoria 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2100
      Width           =   6000
   End
   Begin VB.TextBox Nome 
      Height          =   345
      Left            =   915
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   825
      Width           =   2565
   End
   Begin VB.CheckBox Apuracao 
      Caption         =   "Faz parte da Apuração"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   930
      TabIndex        =   0
      Top             =   210
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
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
      Left            =   180
      TabIndex        =   10
      Top             =   870
      Width           =   585
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
      Left            =   195
      TabIndex        =   11
      Top             =   270
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Categorias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   1860
      Width           =   945
   End
End
Attribute VB_Name = "ContaCategoriaOcx"
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

    'Gera codigo automático.
    lErro = CF("ContaCategoria_Automatico",iCodigo)
    If lErro <> SUCESSO Then Error 57512

    Codigo.PromptInclude = False
    Codigo.Text = CStr(iCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57512
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154946)
    
    End Select

    Exit Sub

End Sub

Private Sub ListaCategoria_Insere(objContaCategoria As ClassContaCategoria)
'Inclui a categoria fornecida por objContaCategoria na listbox ListaCategoria

Dim sListBoxItem As String
    
    'Espacos para completar o tamanho STRING_CODIGO_CONTA_CATEGORIA
    sListBoxItem = Space(STRING_CONTA_CATEGORIA_CODIGO - Len(CStr(objContaCategoria.iCodigo)))
    
    'Concatena o código com o nome da Categoria
    sListBoxItem = sListBoxItem & CStr(objContaCategoria.iCodigo) & SEPARADOR & objContaCategoria.sNome
    ListaCategoria.AddItem (sListBoxItem)
    ListaCategoria.ItemData(ListaCategoria.NewIndex) = objContaCategoria.iCodigo
        
End Sub

Private Sub ListaCategoria_Exclui(iCodigo As Integer)
'Exclui sCategoria de ListaCategoria

Dim iIndice As Integer
Dim sListBoxItem As String

    For iIndice = 0 To ListaCategoria.ListCount - 1
    
        If ListaCategoria.ItemData(iIndice) = iCodigo Then
        
            ListaCategoria.RemoveItem (iIndice)
            Exit For
        
        End If
    
    Next

End Sub

Private Sub Apuracao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objContaCategoria As New ClassContaCategoria
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Codigo foi informado
    If Len(Trim(Codigo.ClipText)) = 0 Then Error 9700

    objContaCategoria.iCodigo = CInt(Codigo.Text)

    'Verifica se a Categoria existe
    lErro = CF("ContaCategoria_Le",objContaCategoria)
    If lErro <> SUCESSO And lErro <> 9651 Then Error 9653
        
    'Categoria não está cadastrada
    If lErro = 9651 Then Error 9654
    
    'Pede confirmação para exclusão
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_CONTACATEGORIA")
            
    If vbMsgRet = vbYes Then
        
        'Exclui a Categoria
        lErro = CF("ContaCategoria_Exclui",objContaCategoria.iCodigo)
        If lErro <> SUCESSO Then Error 9655
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        'Exclui a Categoria da ListBox ListaCategoria
        Call ListaCategoria_Exclui(objContaCategoria.iCodigo)
    
        'Limpa a Tela
        Call Limpa_Tela(Me)
        
        Apuracao.Value = 0
                    
        iAlterado = 0
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
                    
        Case 9653, 9655
        
        Case 9654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, objContaCategoria.iCodigo)
            Codigo.SetFocus
            
        Case 9700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CATEGORIA_NAO_INFORMADO", Err)
            Codigo.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 154947)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

    Call Gravar_Registro
    
End Sub

Public Function Gravar_Registro() As Long
'grava a categoria

Dim lErro As Long
Dim objContaCategoria As New ClassContaCategoria
Dim iApuracao As Integer
Dim iOperacao As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os dados da Categoria foram informados
    If Len(Trim(Codigo.ClipText)) = 0 Then Error 9671
    
    If Len(Trim(Nome.Text)) = 0 Then Error 9701
    
    'Preenche objeto
    objContaCategoria.iCodigo = CInt(Codigo.Text)
    
    'o codigo da categoria não pode ser zero
    If objContaCategoria.iCodigo = 0 Then Error 55696
    
    objContaCategoria.sNome = Nome.Text
    objContaCategoria.iApuracao = Apuracao.Value
        
    lErro = Trata_Alteracao(objContaCategoria, objContaCategoria.iCodigo)
    If lErro <> SUCESSO Then Error 32297
                
    'grava a Categoria no banco de dados
    lErro = CF("ContaCategoria_Grava",objContaCategoria)
    If lErro <> SUCESSO Then Error 9672
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'retira a categoria da listbox, se estiver lá
    Call ListaCategoria_Exclui(objContaCategoria.iCodigo)
    
    'insere a categoria na listbox
    Call ListaCategoria_Insere(objContaCategoria)
            
    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    Apuracao.Value = 0
            '
    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
         
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 9671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CATEGORIA_NAO_INFORMADO", Err)
            Codigo.SetFocus
            
        Case 9672
        
        Case 9701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_CATEGORIA_NAO_INFORMADO", Err)
            Nome.SetFocus
        
        Case 32297
        
        Case 55696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CATEGORIA_ZERADO", Err)
            Codigo.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154948)

     End Select
        
     Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
 
    'se tiver tido alguma alteração na tela, pergunta se quer salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 9673

    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    Apuracao.Value = 0
        
    Codigo.PromptInclude = False
    Codigo.Text = ""
    Codigo.PromptInclude = True
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 9673
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154949)

     End Select
        
     Exit Sub

End Sub

Private Sub Categoria_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
    
Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) > 0 Then

        lErro = Long_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 55695
    
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True

    Select Case Err
    
        Case 55695
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154950)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colContaCategoria As New Collection
Dim objContaCategoria As ClassContaCategoria
Dim sListBoxItem As String

On Error GoTo Erro_Form_Load
        
    'Le todas as categorias existentes no BD
    lErro = CF("ContaCategoria_Le_Todos",colContaCategoria)
    If lErro <> SUCESSO Then Error 9674
    
    For Each objContaCategoria In colContaCategoria
    
        sListBoxItem = Space(STRING_CONTA_CATEGORIA_CODIGO - Len(CStr(objContaCategoria.iCodigo)))
        
        'Concatena Codigo e Nome da Categoria
        sListBoxItem = sListBoxItem & CStr(objContaCategoria.iCodigo)
        sListBoxItem = sListBoxItem & SEPARADOR & objContaCategoria.sNome
    
        ListaCategoria.AddItem sListBoxItem
        ListaCategoria.ItemData(ListaCategoria.NewIndex) = objContaCategoria.iCodigo
        
    Next
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
            
        Case 9674
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154951)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objContaCategoria As ClassContaCategoria) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma Categoria selecionada, exibir seus dados
    If Not (objContaCategoria Is Nothing) Then
        
        'Verifica se o HistPadrao existe
        lErro = CF("ContaCategoria_Le",objContaCategoria)
        If lErro <> 9651 And lErro <> SUCESSO Then Error 9679
        
        If lErro = SUCESSO Then
        
            'A Categoria em questão está cadastrada
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objContaCategoria.iCodigo)
            Codigo.PromptInclude = True
            Nome.Text = objContaCategoria.sNome
            Apuracao.Value = objContaCategoria.iApuracao
                            
        Else
        
            Call Limpa_Tela(Me)
            
            Apuracao.Value = 0
            
            Codigo.PromptInclude = False
            Codigo.Text = CStr(objContaCategoria.iCodigo)
            Codigo.PromptInclude = True
            
        End If
        '
    End If

    iAlterado = 0
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 9679
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154952)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub ListaCategoria_DblClick()

Dim lErro As Long
Dim sHistPadrao As String
Dim sListBoxItem As String
Dim objContaCategoria As New ClassContaCategoria

On Error GoTo Erro_ListaCategoria_DblClick
    
    objContaCategoria.iCodigo = ListaCategoria.ItemData(ListaCategoria.ListIndex)
    
    'Verifica se a Categoria existe
    lErro = CF("ContaCategoria_Le",objContaCategoria)
    If lErro <> 9651 And lErro <> SUCESSO Then Error 9680
    
    'Se a Categoria está cadastrada
    If lErro = SUCESSO Then
            
        'Preenche campos da Tela
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objContaCategoria.iCodigo)
        Codigo.PromptInclude = True
        Nome.Text = objContaCategoria.sNome
        Apuracao.Value = objContaCategoria.iApuracao
                
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        
        iAlterado = 0
        
    Else
    
        'Categoria não existe
        'Exclui da ListBox
        ListaCategoria.RemoveItem (ListaCategoria.ListIndex)
        
    End If
    
    
 
    Exit Sub
    
Erro_ListaCategoria_DblClick:

    Select Case Err
            
        Case 9680
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154953)

    End Select
    
    Exit Sub
    
End Sub

Private Sub ListaCategoria_KeyPress(KeyAscii As Integer)
    
    'Se há categoria selecionada
    If ListaCategoria.ListIndex <> -1 Then
        
        If KeyAscii = ENTER_KEY Then
    
            Call ListaCategoria_DblClick
    
        End If
        
    End If
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objContaCategoria As New ClassContaCategoria

    'Informa tabela associada à Tela
    sTabela = "ContaCategoria"

    'Realiza conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    If Len(Trim(Codigo.Text)) > 0 Then
        objContaCategoria.iCodigo = CInt(Codigo.Text)
    Else
        objContaCategoria.iCodigo = 0
    End If
    
    objContaCategoria.sNome = Nome.Text
    objContaCategoria.iApuracao = Apuracao.Value
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objContaCategoria.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objContaCategoria.sNome, STRING_CONTA_CATEGORIA_NOME, "Nome"
    colCampoValor.Add "Apuracao", objContaCategoria.iApuracao, 0, "Apuracao"

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim objContaCategoria As New ClassContaCategoria

    objContaCategoria.iCodigo = colCampoValor.Item("Codigo").vValor

    If objContaCategoria.iCodigo <> 0 Then

        Codigo.PromptInclude = False
        Codigo.Text = colCampoValor.Item("Codigo").vValor
        Codigo.PromptInclude = True
        
        Nome.Text = colCampoValor.Item("Nome").vValor
        Apuracao.Value = colCampoValor.Item("Apuracao").vValor
          
        iAlterado = 0

    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CATEGORIA_CONTA
    Set Form_Load_Ocx = Me
    Caption = "Categoria de Conta"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ContaCategoria"
    
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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

