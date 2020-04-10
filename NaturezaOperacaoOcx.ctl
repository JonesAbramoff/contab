VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl NaturezaOperacaoOcx 
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   6615
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4320
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "NaturezaOperacaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "NaturezaOperacaoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "NaturezaOperacaoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "NaturezaOperacaoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ListBox ListaNatureza 
      Height          =   2010
      ItemData        =   "NaturezaOperacaoOcx.ctx":0994
      Left            =   120
      List            =   "NaturezaOperacaoOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   6336
   End
   Begin VB.TextBox Descricao 
      Height          =   675
      Left            =   150
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1035
      Width           =   6300
   End
   Begin VB.TextBox DescrNF 
      Height          =   324
      Left            =   2415
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1905
      Width           =   4035
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   270
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   0
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
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
      Left            =   180
      TabIndex        =   12
      Top             =   810
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Naturezas de Operação"
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
      Left            =   165
      TabIndex        =   11
      Top             =   2385
      Width           =   2025
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
      Left            =   330
      TabIndex        =   10
      Top             =   330
      Width           =   660
   End
   Begin VB.Label LblDescrNF 
      AutoSize        =   -1  'True
      Caption         =   "Descrição na Nota Fiscal:"
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
      Left            =   165
      TabIndex        =   9
      Top             =   1965
      Width           =   2220
   End
End
Attribute VB_Name = "NaturezaOperacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Sub Traz_Natureza_Tela(objNaturezaOperacao As ClassNaturezaOp)

    'Natureza de Operação está cadastrada
    Codigo.Text = objNaturezaOperacao.sCodigo
    Descricao.Text = objNaturezaOperacao.sDescricao
    DescrNF.Text = objNaturezaOperacao.sDescrNF

    iAlterado = 0
    
End Sub

Function Trata_Parametros(Optional objNaturezaOperacao As ClassNaturezaOp) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    'Se há uma Natureza de Operação selecionada, exibir seus dados
    If Not (objNaturezaOperacao Is Nothing) Then

        'Verifica se a Natureza de Operação existe
        lErro = CF("NaturezaOperacao_Le", objNaturezaOperacao)
        If lErro <> SUCESSO And lErro <> 17958 Then Error 17963

        'Se não encontrou a Natureza de Operação em questão
        If lErro = 17958 Then
            Codigo.Text = objNaturezaOperacao.sCodigo
        Else
            Call Traz_Natureza_Tela(objNaturezaOperacao)
        End If
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 17963
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163195)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objNaturezaOperacao As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código foi informado
    If Len(Trim(Codigo.Text)) = 0 Then Error 17978

    objNaturezaOperacao.sCodigo = Trim(Codigo.Text)

    'Verifica se a Natureza de Operação existe
    lErro = CF("NaturezaOperacao_Le", objNaturezaOperacao)
    If lErro <> SUCESSO And lErro <> 17958 Then Error 17979

    'Natureza de Operação não está cadastrada
    If lErro = 17958 Then Error 17980

    'Pede confirmação para exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NATUREZAOP", objNaturezaOperacao.sCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a Natureza de Operação
        lErro = CF("NaturezaOperacao_Exclui", objNaturezaOperacao.sCodigo)
        If lErro <> SUCESSO Then Error 17981

        'Exclui a Natureza de Operação da ListBox
        Call ListaNatureza_Exclui(objNaturezaOperacao.sCodigo)

        Call Limpa_Tela_Natureza
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
              Codigo.SetFocus

        Case 17980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objNaturezaOperacao.sCodigo)
            Codigo.SetFocus

        Case 17979, 17981

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 163196)

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
    If lErro <> SUCESSO Then Error 19372

    Call Limpa_Tela_Natureza

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 19372

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163197)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17977

    Call Limpa_Tela_Natureza

Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17977

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163198)

     End Select

     Exit Sub

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

        If Not IsNumeric(Codigo.ClipText) Then Error 17991

        If CLng(Codigo) < 100 Then Error 17992

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 17991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_E_NUMERICO", Err, Codigo.Text)

        Case 17992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO", Err, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163199)

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
Dim colNaturezaOperacao As New Collection
Dim objNaturezaOperacao As New ClassNaturezaOp
Dim sEspacos As String
Dim sNaturezaOp As String

On Error GoTo Erro_Form_Load

    'Preenche a ListBox com Natureza de Operacao existentes no BD
    lErro = CF("NaturezaOperacao_Le_Todos", colNaturezaOperacao)
    If lErro <> SUCESSO Then Error 17954
        
    For Each objNaturezaOperacao In colNaturezaOperacao
        
        'Espacos para completar o tamanho STRING_NATUREZAOP_CODIGO
        sEspacos = Space(STRING_NATUREZAOP_CODIGO - Len(objNaturezaOperacao.sCodigo))
    
        sNaturezaOp = sEspacos & objNaturezaOperacao.sCodigo & SEPARADOR & objNaturezaOperacao.sDescricao
        
        'Insere na ListBox Código e Descrição
        ListaNatureza.AddItem sNaturezaOp

    Next

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 17954

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163200)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub ListaNatureza_DblClick()

Dim lErro As Long
Dim objNaturezaOperacao As New ClassNaturezaOp
Dim sCodigo As String
Dim sListBoxItem As String

On Error GoTo Erro_ListaNatureza_DblClick

    'Se não há Natureza de Operação selecionada sai da rotina
    If ListaNatureza.ListIndex = -1 Then Exit Sub

    sListBoxItem = ListaNatureza.List(ListaNatureza.ListIndex)
    sCodigo = Trim(SCodigo_Extrai(sListBoxItem))
    objNaturezaOperacao.sCodigo = sCodigo

    'Verifica se a Natureza de Operação existe
    lErro = CF("NaturezaOperacao_Le", objNaturezaOperacao)
    If lErro <> SUCESSO And lErro <> 17958 Then Error 17965

    'Natureza de Operação está cadastrada
    If lErro = SUCESSO Then

        Call Traz_Natureza_Tela(objNaturezaOperacao)

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

    'Natureza de Operação não existe
    Else

        'Exclui da ListBox
        ListaNatureza.RemoveItem (ListaNatureza.ListIndex)

    End If

    Exit Sub

Erro_ListaNatureza_DblClick:

    Select Case Err

        Case 17965

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163201)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim sCodigo As String
Dim objNaturezaOperacao As New ClassNaturezaOp

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se dados da Natureza de Operação foram informados
    If Len(Trim(Codigo.Text)) = 0 Then Error 17966

    If Len(Trim(Descricao.Text)) = 0 Then Error 17967
    
    If Len(Trim(DescrNF.Text)) = 0 Then Error 29758

    'Preenche objNaturezaOperacao
    objNaturezaOperacao.sCodigo = Trim(Codigo.Text)
    objNaturezaOperacao.sDescricao = Trim(Descricao.Text)
    objNaturezaOperacao.sDescrNF = Trim(DescrNF.Text)

    lErro = Trata_Alteracao(objNaturezaOperacao, objNaturezaOperacao.sCodigo)
    If lErro <> SUCESSO Then Error 32318

    'Grava a Natureza de Operação no banco de dados
    lErro = CF("NaturezaOperacao_Grava", objNaturezaOperacao)
    If lErro <> SUCESSO Then Error 17968

    Call ListaNatureza_Exclui(objNaturezaOperacao.sCodigo)
    Call ListaNatureza_Adiciona(objNaturezaOperacao)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 17967
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_NAO_PREENCHIDA", Err)

        Case 29758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAONF_NAO_PREENCHIDA", Err)

        Case 17968, 32318

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 163202)

     End Select

     Exit Function

End Function

Private Sub ListaNatureza_Adiciona(objNaturezaOperacao As ClassNaturezaOp)

Dim iIndice As Integer
Dim sListBoxItem As String
Dim sEspacos As String

    'Espacos para completar o tamanho STRING_NATUREZAOP_CODIGO
    sEspacos = Space(STRING_NATUREZAOP_CODIGO - Len(objNaturezaOperacao.sCodigo))
    
    objNaturezaOperacao.sCodigo = sEspacos & objNaturezaOperacao.sCodigo
    
    For iIndice = 0 To ListaNatureza.ListCount - 1
        
        If SCodigo_Extrai(ListaNatureza.List(iIndice)) > objNaturezaOperacao.sCodigo Then Exit For
    Next
    
    sListBoxItem = sEspacos & objNaturezaOperacao.sCodigo & SEPARADOR & objNaturezaOperacao.sDescricao
    ListaNatureza.AddItem sListBoxItem, iIndice
    ListaNatureza.ItemData(iIndice) = objNaturezaOperacao.sCodigo

End Sub

Private Sub ListaNatureza_Exclui(sCodigo As String)

Dim iIndice As Integer
Dim sListBox As String

    For iIndice = 0 To ListaNatureza.ListCount - 1
        
        sListBox = Trim(SCodigo_Extrai(ListaNatureza.List(iIndice)))
        
        If sListBox = sCodigo Then

            ListaNatureza.RemoveItem (iIndice)

            Exit For

        End If

    Next

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objCampoValor As AdmCampoValor
Dim sCodigo As String

    'Informa tabela associada à Tela
    sTabela = "NaturezaOp"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", Trim(Codigo.Text), STRING_NATUREZAOP_CODIGO, "Codigo"
    colCampoValor.Add "Descricao", Descricao.Text, STRING_NATUREZAOP_DESCRICAO, "Descricao"
    colCampoValor.Add "DescrNF", DescrNF.Text, STRING_NATUREZAOP_DESCRNF, "DescrNF"

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim sListBoxItem As String
Dim iIndice As Integer

    'Coloca colCampoValor na Tela
    'Conversão de tipagem para a tipagem da tela se necessário
    Codigo.Text = colCampoValor.Item("Codigo").vValor
    Descricao.Text = colCampoValor.Item("Descricao").vValor
    DescrNF.Text = colCampoValor.Item("DescrNF").vValor

    'Concatena para comparar com ítens da ListBox
    sListBoxItem = Codigo.Text & SEPARADOR & Descricao.Text

    'Seleciona Natureza de Operação na ListBox
    For iIndice = 0 To ListaNatureza.ListCount - 1

        If ListaNatureza.List(iIndice) = sListBoxItem Then
            ListaNatureza.ListIndex = iIndice

            Exit For
        End If

    Next

    iAlterado = 0
    
End Sub

Sub Limpa_Tela_Natureza()
Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa a Tela
    Call Limpa_Tela(Me)
    
    iAlterado = 0

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_NATUREZA_OPERACAO
    Set Form_Load_Ocx = Me
    Caption = "Naturezas de Operação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NaturezaOperacao"
    
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

Private Sub LblDescrNF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblDescrNF, Source, X, Y)
End Sub

Private Sub LblDescrNF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblDescrNF, Button, Shift, X, Y)
End Sub

