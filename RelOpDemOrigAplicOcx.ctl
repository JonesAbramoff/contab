VERSION 5.00
Begin VB.UserControl RelOpDemOrigAplicOcx 
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   ScaleHeight     =   2490
   ScaleWidth      =   7425
   Begin VB.CommandButton BotaoConfigura 
      Caption         =   "Configurar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4725
      Picture         =   "RelOpDemOrigAplicOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1665
      Width           =   1455
   End
   Begin VB.ComboBox ComboModelos 
      Height          =   315
      ItemData        =   "RelOpDemOrigAplicOcx.ctx":0BA2
      Left            =   1380
      List            =   "RelOpDemOrigAplicOcx.ctx":0BA4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
      Width           =   2775
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
      Height          =   720
      Left            =   4725
      Picture         =   "RelOpDemOrigAplicOcx.ctx":0BA6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   825
      Width           =   1455
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpDemOrigAplicOcx.ctx":0CA8
      Left            =   1395
      List            =   "RelOpDemOrigAplicOcx.ctx":0CAA
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   825
      Width           =   1860
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDemOrigAplicOcx.ctx":0CAC
      Left            =   1410
      List            =   "RelOpDemOrigAplicOcx.ctx":0CAE
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   225
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4440
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemOrigAplicOcx.ctx":0CB0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemOrigAplicOcx.ctx":0E2E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemOrigAplicOcx.ctx":1360
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemOrigAplicOcx.ctx":14EA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
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
      Left            =   645
      TabIndex        =   12
      Top             =   1485
      Width           =   690
   End
   Begin VB.Label Label6 
      Caption         =   "Exercicio:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   870
      Width           =   855
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
      Left            =   720
      TabIndex        =   10
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpDemOrigAplicOcx"
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
Dim giFocoInicial As Integer

Function MostraExercicio(iExercicio As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicio

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next

    MostraExercicio = SUCESSO

    Exit Function

Erro_MostraExercicio:

    MostraExercicio = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168152)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 60698

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 60699

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 60698
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)

        Case 60699

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168153)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long

On Error GoTo Erro_PreencherRelOp

    'Exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 60700

    'Verifica se Modelo foi preenchido
    If ComboModelos.ListIndex = -1 Then Error 60701

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 60702

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 60703

    lErro = objRelOpcoes.IncluirParametro("TEXERCICIO", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 60704

    lErro = objRelOpcoes.IncluirParametro("TMODELO", ComboModelos.Text)
    If lErro <> AD_BOOL_TRUE Then Error 60705

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 60700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 60701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", Err)
            ComboModelos.SetFocus

        Case 60702, 60703, 60704, 60705

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168154)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 60706

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 60707

    iExercicio = CInt(sParam)

    lErro = MostraExercicio(iExercicio)
    If lErro <> SUCESSO Then Error 60708

    lErro = objRelOpcoes.ObterParametro("TMODELO", sParam)
    If lErro <> SUCESSO Then Error 60709

    ComboModelos.Text = sParam

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 60706, 60707, 60708, 60709

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168155)

    End Select

    Exit Function

End Function

Private Sub BotaoConfigura_Click()

Dim iIndice As Integer
Dim lErro As Long
Dim colModelos As New Collection

On Error GoTo Erro_BotaoConfigura_Click

    Call Chama_Tela("RelDREConfig", RELDOAR, "Configuração da Demonstração das Origens e Aplicações de Recursos")

    Exit Sub

Erro_BotaoConfigura_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168156)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 60710

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui",gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 60711

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 60712


    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 60710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 60711, 60712

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168157)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim lCodIdExec As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 60713

    lErro = gobjRelOpcoes.IncluirParametro("NIDEXECREL", CStr(lCodIdExec))
    If lErro <> AD_BOOL_TRUE Then Error 60714

    lErro = RelDOAR_Calcula(ComboModelos.List(ComboModelos.ListIndex), ComboExercicio.ItemData(ComboExercicio.ListIndex), giFilialEmpresa)
    If lErro <> SUCESSO Then Error 60715

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 60713, 60714, 60715

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168158)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long, iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 60716

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 60717

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava",gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 60718

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 60719

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 60716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 60717, 60718, 60719

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168159)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 60720

    ComboOpcoes.Text = ""

    ComboOpcoes.SetFocus

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 60720

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168160)

    End Select

    Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection
Dim colModelos As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos",colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 60721

    For iIndice = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iIndice)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next

    lErro = MostraExercicio(giExercicioAtual)
    If lErro <> SUCESSO Then Error 60722

    'Le os Modelos na tabela RelDRE
    lErro = CF("RelDRE_Le_Modelos_Distintos",RELDOAR, colModelos)
    If lErro <> SUCESSO And lErro <> 47101 Then Error 60723

    'se  encontrou alguém
    If lErro = SUCESSO Then

        'preenche a combo Modelos
        For iIndice = 1 To colModelos.Count
            ComboModelos.AddItem colModelos.Item(iIndice)
        Next

        ComboModelos.ListIndex = -1


    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 60721, 60722, 60723

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168161)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Private Function RelDOAR_Calcula(ByVal sModelo As String, ByVal iExercicio As Integer, ByVal iFilialEmpresa As Integer) As Long
'Calcula Valor correspondente ao Modelo

Dim lErro As Long
Dim iOperacao As Integer
Dim dValorExercAtual As Double
Dim dValorExercAnt As Double
Dim colRelDRE As New Collection
Dim colRelDREConta As New Collection
Dim colRelDREFormula As New Collection
Dim objRelDRE As New ClassRelDRE
Dim objRelDRE1 As New ClassRelDRE
Dim objRelDREConta As New ClassRelDREConta
Dim objRelDREFormula As New ClassRelDREFormula

On Error GoTo Erro_RelDOAR_Calcula

    'Lê registros da tabela RelDRE para o Modelo passado como parâmetro
    lErro = CF("RelDRE_Le_Modelo",RELDOAR, sModelo, colRelDRE)
    If lErro <> SUCESSO Then Error 60724

    'Lê registros da tabela RelDREConta para o Modelo passado como parâmetro
    lErro = CF("RelDREConta_Le_Modelo",RELDOAR, sModelo, colRelDREConta)
    If lErro <> SUCESSO Then Error 60725

    'Lê registros da tabela RelDREFormula para o Modelo passado como parâmetro
    lErro = CF("RelDREFormula_Le_Modelo",RELDOAR, sModelo, colRelDREFormula)
    If lErro <> SUCESSO Then Error 60726

    For Each objRelDRE In colRelDRE

        If objRelDRE.iTipo = DRE_TIPO_CONTA Then

            For Each objRelDREConta In colRelDREConta

                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREConta.iCodigo = objRelDRE.iCodigo Then

                    If objRelDRE.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then

                        'Calcula o Valor
                        lErro = CF("MvPerCta_Calcula_Valor",iFilialEmpresa, iExercicio - 1, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorExercAtual, dValorExercAnt)
                        If lErro <> SUCESSO Then Error 60729

                    Else

                        'Calcula o Valor
                        lErro = CF("MvPerCta_Calcula_Valor",iFilialEmpresa, iExercicio, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorExercAtual, dValorExercAnt)
                        If lErro <> SUCESSO Then Error 60727

                    End If

                    'Aculmula o Valor
                    objRelDRE.dValor = objRelDRE.dValor + dValorExercAtual
                    objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + dValorExercAnt

                End If

            Next

        ElseIf objRelDRE.iTipo = DRE_TIPO_FORMULA Then

            iOperacao = DRE_OPERACAO_SOMA

            For Each objRelDREFormula In colRelDREFormula

                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREFormula.iCodigo = objRelDRE.iCodigo Then

                    For Each objRelDRE1 In colRelDRE

                        'Pesquisar na coleção RelDRE o elemento que tenha o Código igual ao
                        'código da fórmula do elemento da coleção RelDREFormula
                        If objRelDRE1.iCodigo = objRelDREFormula.iCodigoFormula Then

                            If iOperacao = DRE_OPERACAO_SOMA Then
                                objRelDRE.dValor = objRelDRE.dValor + objRelDRE1.dValor
                                objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + objRelDRE1.dValorExercAnt
                            Else
                                objRelDRE.dValor = objRelDRE.dValor - objRelDRE1.dValor
                                objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt - objRelDRE1.dValorExercAnt
                            End If

                            iOperacao = objRelDREFormula.iOperacao

                            Exit For

                        End If

                    Next

                End If

            Next

        End If

    Next

    'Grava o Valor acumulado na tabela RelDRE
    lErro = CF("RelDRE_Grava_Valor",RELDOAR, colRelDRE)
    If lErro <> SUCESSO Then Error 60728

    RelDOAR_Calcula = SUCESSO

    Exit Function

Erro_RelDOAR_Calcula:

    RelDOAR_Calcula = Err

    Select Case Err

        Case 60724, 60725, 60726, 60727, 60728, 60729

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168162)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_DEMONST_ORIGENS_APLIC
    Set Form_Load_Ocx = Me
    Caption = "Demonstração Comparativa das Origens e Aplicações de Recursos"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpDemOrigAplic"

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





Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

