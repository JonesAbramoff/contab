VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl RelOpPosAplicOcx 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   6270
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
         Picture         =   "RelOpPosAplicOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPosAplicOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPosAplicOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPosAplicOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
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
      Height          =   585
      Left            =   4260
      Picture         =   "RelOpPosAplicOcx.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPosAplicOcx.ctx":0A96
      Left            =   840
      List            =   "RelOpPosAplicOcx.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aplicações"
      Height          =   1320
      Left            =   180
      TabIndex        =   9
      Top             =   795
      Width           =   3570
      Begin MSMask.MaskEdBox CodigoInicial 
         Height          =   300
         Left            =   1755
         TabIndex        =   1
         Top             =   345
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodigoFinal 
         Height          =   300
         Left            =   1740
         TabIndex        =   2
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label CodigoFinalLabel 
         AutoSize        =   -1  'True
         Caption         =   "Código Final:"
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
         Left            =   525
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   11
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label CodigoInicialLabel 
         AutoSize        =   -1  'True
         Caption         =   "Código Inicial:"
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
         Left            =   435
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   10
         Top             =   375
         Width           =   1230
      End
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPosAplicOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private WithEvents objEventoCodigoInicial As AdmEvento
Attribute objEventoCodigoInicial.VB_VarHelpID = -1
Private WithEvents objEventoCodigoFinal As AdmEvento
Attribute objEventoCodigoFinal.VB_VarHelpID = -1

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros de uma opcao salva anteriormente e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 23353

    'Pega parametros e exibe
    lErro = objRelOpcoes.ObterParametro("TCODINI", sParam)
    If lErro <> SUCESSO Then Error 23354
    
    CodigoInicial.Text = CStr(sParam)
    
    lErro = objRelOpcoes.ObterParametro("TCODFIM", sParam)
    If lErro <> SUCESSO Then Error 23355
    
    CodigoFinal.Text = CStr(sParam)
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 23353, 23354, 23355

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171221)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 23348
    
    'Se os códigos inicial e final foram preenchidos
    If Len(Trim(CodigoFinal.Text)) <> 0 And Len(Trim(CodigoInicial.Text)) <> 0 Then
    
        'Verificar se o código final é maior
        If CInt(CodigoFinal.Text) < CInt(CodigoInicial.Text) Then Error 23349
        
    End If
    
    'Pegar parametros da tela
    lErro = objRelOpcoes.IncluirParametro("TCODINI", CodigoInicial.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23351

    lErro = objRelOpcoes.IncluirParametro("TCODFIM", CodigoFinal.Text)
    If lErro <> AD_BOOL_TRUE Then Error 23352
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 23350

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 23348
        
        Case 23349
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_INICIAL_MAIOR", Err, Error$)
            CodigoInicial.SetFocus
        
        Case 23350, 23351, 23352
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171222)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoCodigoInicial = Nothing
    Set objEventoCodigoFinal = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29904
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche combo com as opções de relatório
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 23362

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 23362
        
        Case 29904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171223)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 23356

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPPOSAPLIC")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 23357

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 23356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 23357

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171224)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23358

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 23358

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171225)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'grava os parametros informados no preenchimento da tela associando-os a um "nome de opção"

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 23359

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 23360

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 23361

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 57759
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 23359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 23360, 23361, 57759

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171226)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    
    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub CodigoFinalLabel_Click()

Dim objAplicacao As New ClassAplicacao
Dim colSelecao As Collection

    If Len(Trim(CodigoFinal.Text)) = 0 Then
        objAplicacao.lCodigo = 0
    Else
        objAplicacao.lCodigo = CLng(CodigoFinal.Text)
    End If

    Call Chama_Tela("AplicacaoLista", colSelecao, objAplicacao, objEventoCodigoFinal)

End Sub

Private Sub CodigoInicialLabel_Click()

Dim objAplicacao As New ClassAplicacao
Dim colSelecao As Collection

    If Len(Trim(CodigoInicial.Text)) = 0 Then
        objAplicacao.lCodigo = 0
    Else
        objAplicacao.lCodigo = CLng(CodigoInicial.Text)
    End If

    Call Chama_Tela("AplicacaoLista", colSelecao, objAplicacao, objEventoCodigoInicial)

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_OpcoesRel_Form_Load

    Set objEventoCodigoInicial = New AdmEvento
    Set objEventoCodigoFinal = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_OpcoesRel_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171227)

    End Select

    Unload Me

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção que será incluida dinamicamente para a execucao do relatorio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If CodigoInicial.Text <> "" Then sExpressao = "Codigo >= " & Forprint_ConvInt(CInt(CodigoInicial.Text))

    If CodigoFinal.Text <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Codigo <= " & Forprint_ConvInt(CInt(CodigoFinal.Text))
    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171228)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POS_APLIC
    Set Form_Load_Ocx = Me
    Caption = "Posição das Aplicações"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPosAplic"
    
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

Private Sub objEventoCodigoFinal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAplicacao As ClassAplicacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objAplicacao = obj1

    CodigoFinal.Text = CStr(objAplicacao.lCodigo)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171229)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigoInicial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAplicacao As ClassAplicacao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objAplicacao = obj1

    CodigoInicial.Text = CStr(objAplicacao.lCodigo)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 171230)

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

Public Sub Unload(objme As Object)
    
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
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodigoInicial Then
            Call CodigoInicialLabel_Click
        ElseIf Me.ActiveControl Is CodigoFinal Then
            Call CodigoFinalLabel_Click
        End If
    
    End If

End Sub


Private Sub CodigoFinalLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoFinalLabel, Source, X, Y)
End Sub

Private Sub CodigoFinalLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoFinalLabel, Button, Shift, X, Y)
End Sub

Private Sub CodigoInicialLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoInicialLabel, Source, X, Y)
End Sub

Private Sub CodigoInicialLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoInicialLabel, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

