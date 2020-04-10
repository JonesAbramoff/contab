VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpRelOperadores 
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   6705
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4733
      Picture         =   "RelOpreloperadores.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   945
      Width           =   1605
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "RelOpreloperadores.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOpreloperadores.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpreloperadores.ctx":07B2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpreloperadores.ctx":093C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpreloperadores.ctx":0A96
      Left            =   1080
      List            =   "RelOpreloperadores.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2670
   End
   Begin VB.Frame FrameOperador 
      Caption         =   "Operador"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      Begin MSMask.MaskEdBox OperadorDe 
         Height          =   315
         Left            =   900
         TabIndex        =   11
         Top             =   285
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox OperadorAte 
         Height          =   315
         Left            =   900
         TabIndex        =   12
         Top             =   765
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelOperadorAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
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
         Left            =   360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   825
         Width           =   360
      End
      Begin VB.Label LabelOperadorDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   345
         Width           =   315
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelOperadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOperador As AdmEvento
Attribute objEventoOperador.VB_VarHelpID = -1

Dim giOperadorDe As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_NF
    Set Form_Load_Ocx = Me
    Caption = "Relação de Operadores"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpRelOperadores"

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

Private Sub LabelOperadorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOperadorDe, Source, X, Y)
End Sub

Private Sub LabelOperadorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOperadorDe, Button, Shift, X, Y)
End Sub
Private Sub LabelOperadorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOperadorDe, Source, X, Y)
End Sub

Private Sub LabelOperadorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOperadorAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub ComboOpcoes_Click()
    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoOperador = New AdmEvento
   
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172591)

    End Select

    Exit Sub

End Sub

Private Sub OperadorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_OperadorDe_Validate

    'Se o campo está preenchido
    If Len(Trim(OperadorDe.Text)) > 0 Then
   
        'Tenta ler o codigo do Operador
        lErro = CF("TP_Operador_Le", OperadorDe, objOperador)
        If lErro <> SUCESSO And lErro <> 117117 And lErro <> 117119 Then gError 117098
        
        'Se o operador não foi encontrado => erro
        If lErro = 117117 Or lErro = 117119 Then gError 127092

    End If
    
    giOperadorDe = 1
    
    Exit Sub

Erro_OperadorDe_Validate:

    Cancel = True
    
    Select Case gErr

        Case 117098
        
        Case 127092
             Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_ENCONTRADO", gErr, OperadorDe.Text)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172592)

    End Select

End Sub

Private Sub OperadorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOperador As New ClassOperador

On Error GoTo Erro_OperadorAte_Validate

    'Se o campo está preenchido
    If Len(Trim(OperadorAte.Text)) > 0 Then
        
        'Tenta ler o Código do Operador
        lErro = CF("TP_Operador_Le", OperadorAte, objOperador)
        If lErro <> SUCESSO And lErro <> 117117 And lErro <> 117119 Then gError 117099

        'Se o operador não foi encontrado => erro
        If lErro = 117117 Or lErro = 117119 Then gError 127093

    End If
    
    giOperadorDe = 0
 
    Exit Sub

Erro_OperadorAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 117099
            
        Case 127093
             Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_NAO_ENCONTRADO", gErr, OperadorDe.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172593)

    End Select

End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set objEventoOperador = Nothing
End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sOperador_De As String, sOperador_Ate As String) As Long
'Monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   'Inclui a seleção por operador inicial
   If sOperador_De <> "" Then sExpressao = "Operador >= " & Forprint_ConvInt(StrParaInt(sOperador_De))

   'Inclui a seleção por operador final
   If sOperador_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Operador <= " & Forprint_ConvInt(StrParaInt(sOperador_Ate))

    End If
    
    'Inclui a seleção por filialempresa
    If giFilialEmpresa <> EMPRESA_TODA Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(giFilialEmpresa)
    End If
      
    
    'passa a expressão completa para o obj
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If
    
    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172594)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 117100
    
    'Exibe Operador inicial
    lErro = objRelOpcoes.ObterParametro("NOPINIC", sParam)
    If lErro <> SUCESSO Then gError 117101
    
    bSGECancelDummy = False
    OperadorDe.PromptInclude = False
    OperadorDe.Text = sParam
    OperadorDe.PromptInclude = True
    Call OperadorDe_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then OperadorDe.Text = ""
    
    'Exibe Operador final
    lErro = objRelOpcoes.ObterParametro("NOPFIM", sParam)
    If lErro <> SUCESSO Then gError 117102
    
    bSGECancelDummy = False
    OperadorAte.PromptInclude = False
    OperadorAte.Text = sParam
    OperadorAte.PromptInclude = True
    Call OperadorAte_Validate(bSGECancelDummy)
    If bSGECancelDummy = True Then OperadorAte.Text = ""
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 117100 To 117102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172595)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sOperador_De As String
Dim sOperador_Ate As String

On Error GoTo Erro_PreencherRelOp
       
    lErro = Formata_E_Critica_Parametros(sOperador_De, sOperador_Ate)
    If lErro <> SUCESSO Then gError 117103
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 117104
             
    lErro = objRelOpcoes.IncluirParametro("NOPINIC", sOperador_De)
    If lErro <> AD_BOOL_TRUE Then gError 117105
    
    lErro = objRelOpcoes.IncluirParametro("TOPINIC", Trim(OperadorDe.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117106

    lErro = objRelOpcoes.IncluirParametro("NOPFIM", sOperador_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 117107
    
    lErro = objRelOpcoes.IncluirParametro("TOPFIM", Trim(OperadorAte.Text))
    If lErro <> AD_BOOL_TRUE Then gError 117108
               
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sOperador_De, sOperador_Ate)
    If lErro <> SUCESSO Then gError 117109
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 117103 To 117109
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172596)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 117110
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 117111
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 117111
        
        Case 117110
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172597)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is OperadorDe Then
            Call LabelOperadorDe_Click
            
        ElseIf Me.ActiveControl Is OperadorAte Then
            Call LabelOperadorAte_Click
            
        End If
    
    End If

End Sub

Private Function Formata_E_Critica_Parametros(sOperador_De As String, sOperador_Ate As String) As Long

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
   
    'critica Operador Inicial e Final
    If OperadorDe.ClipText <> "" Then
        sOperador_De = CStr(Codigo_Extrai(OperadorDe.ClipText))
    Else
        sOperador_De = ""
    End If
    
    If OperadorAte.ClipText <> "" Then
        sOperador_Ate = CStr(Codigo_Extrai(OperadorAte.ClipText))
    Else
        sOperador_Ate = ""
    End If
            
    If sOperador_De <> "" And sOperador_Ate <> "" Then
        
        'Se o Operador Inicial for maior que o final --> erro
        If CInt(sOperador_De) > CInt(sOperador_Ate) Then gError 117112
        
    End If
         
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 117112
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERADOR_INICIAL_MAIOR", gErr)
            OperadorDe.SetFocus
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172598)

    End Select

    Exit Function

End Function

Private Sub LabelOperadorDe_Click()
Dim objOperador As New ClassOperador
Dim colSelecao As Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelOperadorDe_Click
    
    giOperadorDe = 1
    
    'Se é possível extrair o código do operador
    If Codigo_Extrai(OperadorDe.Text) > 0 Then
        
        'Guarda o código do operador
        objOperador.iCodigo = Codigo_Extrai(OperadorDe.Text)
        
        'Indica que os registros no browser serão ordenados por código
        sOrdenacao = "Código"
    
    'Senão, ou seja, se está digitado o nome do operador
    Else
    
        'Guarda o nome do operador
        objOperador.sNome = OperadorDe.Text
        
        'Indica que os registros no browser serão ordenados por nome
        sOrdenacao = "Nome do Operador"
    
    End If

    'Chama Tela OperadorLista
    Call Chama_Tela("OperadorLista", colSelecao, objOperador, objEventoOperador, "", sOrdenacao)
    
     Exit Sub

Erro_LabelOperadorDe_Click:

    LabelOperadorDe = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172599)

    End Select

    Exit Sub

End Sub

Private Sub LabelOperadorAte_Click()

Dim objOperador As New ClassOperador
Dim colSelecao As Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelOperadorAte_Click
    
    giOperadorDe = 0

    'Se é possível extrair o código do operador
    If Codigo_Extrai(OperadorAte.Text) > 0 Then
        
        'Guarda o código do operador
        objOperador.iCodigo = Codigo_Extrai(OperadorAte.Text)
        
        'Indica que os registros no browser serão ordenados por código
        sOrdenacao = "Código"
    
    'Senão, ou seja, se está digitado o nome do operador
    Else
    
        'Guarda o nome do operador
        objOperador.sNome = OperadorAte.Text
        
        'Indica que os registros no browser serão ordenados por nome
        sOrdenacao = "Nome do Operador"
    
    End If

    'Chama Tela OperadorLista
    Call Chama_Tela("OperadorLista", colSelecao, objOperador, objEventoOperador, "", sOrdenacao)
  
     Exit Sub

Erro_LabelOperadorAte_Click:

    LabelOperadorAte = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172600)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 117090
           
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 117090

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172601)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Nome da opção de Relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 117091

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 117092

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 117093

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 117094
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 117091
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 117092, 117093, 117094

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172602)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 117095

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 117096

        'Retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa as opções da tela
        Call BotaoLimpar_Click

        ComboOpcoes.Text = ""

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 117095
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 117096

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172603)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 117097

    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 117097

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172604)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOperador_evSelecao(obj1 As Object)

Dim objOperador As ClassOperador

    Set objOperador = obj1
    
    'Preenche campo Vendedor
    If giOperadorDe = 1 Then
        OperadorDe.Text = CStr(objOperador.iCodigo)
        OperadorDe_Validate (bSGECancelDummy)
    Else
        OperadorAte.Text = CStr(objOperador.iCodigo)
        OperadorAte_Validate (bSGECancelDummy)
    End If

    Me.Show

     Exit Sub
End Sub

