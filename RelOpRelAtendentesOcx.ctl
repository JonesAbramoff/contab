VERSION 5.00
Begin VB.UserControl RelOpRelAtendentesOcx 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   ScaleHeight     =   2295
   ScaleWidth      =   6405
   Begin VB.Frame FrameAtendentes 
      Caption         =   "Atendentes"
      Height          =   1140
      Left            =   360
      TabIndex        =   10
      Top             =   840
      Width           =   3645
      Begin VB.ComboBox AtendenteAte 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox AtendenteDe 
         Height          =   315
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label LabelAtendenteAte 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   780
         Width           =   360
      End
      Begin VB.Label LabelAtendenteDe 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpRelAtendentesOcx.ctx":0000
      Left            =   945
      List            =   "RelOpRelAtendentesOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2730
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
      Height          =   600
      Left            =   4245
      Picture         =   "RelOpRelAtendentesOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   870
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   8
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpRelAtendentesOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpRelAtendentesOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpRelAtendentesOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpRelAtendentesOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
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
      Left            =   240
      TabIndex        =   9
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpRelAtendentesOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega a combo AtendenteDe
    lErro = CF("Carrega_Atendentes", AtendenteDe)
    If lErro <> SUCESSO Then gError 131400
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", AtendenteAte)
    If lErro <> SUCESSO Then gError 131401
    
    AtendenteDe.Text = ""
    AtendenteAte.Text = ""
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 131400 To 131401
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172365)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 131402

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche a Combo Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 131403

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 131402
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case 131403

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172366)

    End Select

    Exit Function

End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO***
Public Sub AtendenteDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteDe_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteDe)
    If lErro <> SUCESSO Then gError 131404
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 131405
    
    Exit Sub
    
Erro_AtendenteDe_Validate:

    Cancel = True
    
    Select Case gErr
                
        Case 131404

        Case 131405
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172367)

    End Select

End Sub

Public Sub AtendenteAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AtendenteAte_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("Atendente_Validate", AtendenteAte)
    If lErro <> SUCESSO Then gError 131406
    
    'Se os atendentes foram preenchidos e o atendente de for maior que o atendente até => erro
    If Len(Trim(AtendenteDe.Text)) > 0 And Len(Trim(AtendenteAte.Text)) > 0 And Codigo_Extrai(AtendenteDe.Text) > Codigo_Extrai(AtendenteAte.Text) Then gError 131407
    
    Exit Sub

Erro_AtendenteAte_Validate:

    Cancel = True
    
    Select Case gErr

        Case 131406
        
        Case 131407
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTEDE_MAIOR_ATENDENTEATE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172368)

    End Select

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click
    
    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131408
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 131408

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172369)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
     'Limpa a tela
    Call LimpaRelatorioAtendentes
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr

        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172370)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 131409

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_ATENDENTES")

    If vbMsgRes = vbYes Then

        'Exclui o elemento do banco de dados
        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 131410

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'Limpa a tela
        lErro = LimpaRelatorioAtendentes()
        If lErro <> SUCESSO Then gError 131411
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 131409
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 131410, 131411

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172371)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 131412

    'Faz a chamada da função que irá realizar o preenchimento do objeto RelOpcoes
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 131413

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    'Grava no banco de dados
    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 131414
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 131415
    
    'Limpa a tela
    lErro = LimpaRelatorioAtendentes()
    If lErro <> SUCESSO Then gError 131416
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 131412
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 131413 To 131416
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172372)

    End Select

    Exit Sub

End Sub
'*** EVENTO CLICK DOS CONTROLES - FIM ***
Private Function LimpaRelatorioAtendentes()
'Limpa a tela RelOpRelacClientes

Dim lErro As Long

On Error GoTo Erro_LimpaRelatorioAtendentes

    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 131417
    
    ComboOpcoes.Text = ""
    
    LimpaRelatorioAtendentes = SUCESSO
    
    Exit Function
    
Erro_LimpaRelatorioAtendentes:

    LimpaRelatorioAtendentes = gErr
    
    Select Case gErr
    
        Case 131417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172373)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o objRelOp com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iIndice As Integer
Dim sAtendente_De As String
Dim sAtendente_Ate As String

On Error GoTo Erro_PreencherRelOp
   
    'Critica os valores preenchidos pelo usuário
    lErro = Formata_E_Critica_Parametros(sAtendente_De, sAtendente_Ate)
    If lErro <> SUCESSO Then gError 131418
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 131419
        
    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("NATENDDE", sAtendente_De)
    If lErro <> AD_BOOL_TRUE Then gError 131420

    'Inclui o atendente inicial
    lErro = objRelOpcoes.IncluirParametro("TATENDDE", AtendenteDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131662

    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("NATENDATE", sAtendente_Ate)
    If lErro <> AD_BOOL_TRUE Then gError 131421
    
    'Inclui o código final
    lErro = objRelOpcoes.IncluirParametro("TATENDATE", AtendenteAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 131663

    
    'Faz a chamada da função que irá montar a expressão
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sAtendente_De, sAtendente_Ate)
    If lErro <> SUCESSO Then gError 131422
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 131418 To 131422, 131662 To 131663
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172374)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sAtendente_De As String, sAtendente_Ate As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
     
    'Verifica se o atendente inicial foi preenchido
    If AtendenteDe.Text <> "" Then
        sAtendente_De = CStr(LCodigo_Extrai(AtendenteDe.Text))
    Else
        sAtendente_De = ""
    End If
    
    'Verifica se o atendente final foi preenchido
    If AtendenteAte.Text <> "" Then
        sAtendente_Ate = CStr(LCodigo_Extrai(AtendenteAte.Text))
    Else
        sAtendente_Ate = ""
    End If
    
    'Verifica se o atendente inicial é menor que o final, se não for --> ERRO
    If sAtendente_De <> "" And sAtendente_Ate <> "" Then
        If CInt(sAtendente_De) > CInt(sAtendente_Ate) Then gError 131423
    End If
        
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                     
        Case 131423
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_INICIAL_MAIOR_FINAL", gErr)
            AtendenteDe.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172375)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sAtendente_De As String, sAtendente_Ate As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Monta_Expressao_Selecao
      
    'Verifica se o atendente final foi preenchido
    If sAtendente_De <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente >= " & Forprint_ConvInt(CInt(sAtendente_De))
        
    End If
    
    'Verifica se o atendente final foi preenchido
    If sAtendente_Ate <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Atendente <= " & Forprint_ConvInt(CInt(sAtendente_Ate))
        
    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO
    
    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172376)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim iTipo As Integer
Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 131424
        
    'Preenche o atendente inicial
    lErro = objRelOpcoes.ObterParametro("NATENDDE", sParam)
    If lErro <> SUCESSO Then gError 131425
    
    AtendenteDe.Text = sParam
    Call AtendenteDe_Validate(bSGECancelDummy)
    
    'Preenche o atendente final
    lErro = objRelOpcoes.ObterParametro("NATENDATE", sParam)
    If lErro <> SUCESSO Then gError 131426
    
    AtendenteAte.Text = sParam
    Call AtendenteAte_Validate(bSGECancelDummy)
                
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 131424 To 131426
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 172377)

    End Select

    Exit Function

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITPAG_L
    Set Form_Load_Ocx = Me
    Caption = "Atendentes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRelAtendentes"
    
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



