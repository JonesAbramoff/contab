VERSION 5.00
Begin VB.UserControl RelOpMutacaoPLOcx 
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   ScaleHeight     =   2550
   ScaleWidth      =   6255
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
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
      Left            =   4275
      Picture         =   "RelOpMutacaoPLOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1665
      Width           =   1455
   End
   Begin VB.ComboBox ComboModelos 
      Height          =   315
      ItemData        =   "RelOpMutacaoPLOcx.ctx":0BA2
      Left            =   1080
      List            =   "RelOpMutacaoPLOcx.ctx":0BA4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1425
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
      Left            =   4275
      Picture         =   "RelOpMutacaoPLOcx.ctx":0BA6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   825
      Width           =   1455
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpMutacaoPLOcx.ctx":0CA8
      Left            =   1080
      List            =   "RelOpMutacaoPLOcx.ctx":0CAA
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   825
      Width           =   1860
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpMutacaoPLOcx.ctx":0CAC
      Left            =   1080
      List            =   "RelOpMutacaoPLOcx.ctx":0CAE
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   225
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3990
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpMutacaoPLOcx.ctx":0CB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpMutacaoPLOcx.ctx":0E2E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpMutacaoPLOcx.ctx":1360
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpMutacaoPLOcx.ctx":14EA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Período:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2070
      Width           =   735
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
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   855
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
      Left            =   285
      TabIndex        =   11
      Top             =   1470
      Width           =   690
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
      Left            =   270
      TabIndex        =   10
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "RelOpMutacaoPLOcx"
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
Dim giCarregando As Integer

Const CEL_TIPO_CONTA = 0
Const CEL_TIPO_FORMULA = 1
Const CEL_TIPO_TITULO = 2

Const REL_OPERACAO_SOMA = 0
Const REL_OPERACAO_SUBTRAI = 1

Const CONTAS_EXERCICIO_ATUAL = 0
Const CONTAS_EXERCICIO_ANTERIOR = 1

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170153)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 60647
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 60648

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 60647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case 60648
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170154)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim iPeriodo As Integer

On Error GoTo Erro_PreencherRelOp

    'Exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 60649
    
    'Verifica se Modelo foi preenchido
    If ComboModelos.ListIndex = -1 Then Error 60650

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 60651

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 60652
    
    lErro = objRelOpcoes.IncluirParametro("TEXERCICIO", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 60653

    lErro = objRelOpcoes.IncluirParametro("TMODELO", ComboModelos.Text)
    If lErro <> AD_BOOL_TRUE Then Error 60654
    
    If ComboPeriodo.ListIndex <> -1 Then
        iPeriodo = ComboPeriodo.ItemData(ComboPeriodo.ListIndex)
    Else
        iPeriodo = 0
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(iPeriodo))
    If lErro <> AD_BOOL_TRUE Then Error 60654

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 60649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 60650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", Err)
            ComboModelos.SetFocus
            
        Case 60651, 60652, 60653, 60654

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170155)
            
    End Select
    
    Exit Function
    
End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer
Dim sParam As String, iPeriodo As Integer

On Error GoTo Erro_PreencherParametrosNaTela
    
    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 60655

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 60656

    iExercicio = CInt(sParam)

    lErro = MostraExercicio(iExercicio)
    If lErro <> SUCESSO Then Error 60657
    
    'período
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 60657

    iPeriodo = CInt(sParam)

    If iPeriodo <> 0 Then
        lErro = MostraExercicioPeriodo(iExercicio, iPeriodo)
        If lErro <> SUCESSO Then Error 60657
    End If
    
    lErro = objRelOpcoes.ObterParametro("TMODELO", sParam)
    If lErro <> SUCESSO Then Error 60658

    ComboModelos.Text = sParam
    
    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 60655, 60656, 60657, 60658

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170156)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoConfigura_Click()
    
Dim iIndice As Integer
Dim lErro As Long
Dim colModelos As New Collection

On Error GoTo Erro_BotaoConfigura_Click
    
    Call Chama_Tela("RelDMPLConfig", RELDMPL, "Configuração da Demonstração de Mutação do Patrimônio Líquido")
    
    Exit Sub

Erro_BotaoConfigura_Click:
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170157)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 60659

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 60660

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa as opções da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 60661
    

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 60659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 60660, 60661

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170158)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim lCodIdExec As Long
Dim iPeriodo As Integer

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 60662

    lErro = gobjRelOpcoes.IncluirParametro("NIDEXECREL", CStr(lCodIdExec))
    If lErro <> AD_BOOL_TRUE Then Error 60663
    
    If ComboPeriodo.ListIndex <> -1 Then
        iPeriodo = ComboPeriodo.ItemData(ComboPeriodo.ListIndex)
    Else
        iPeriodo = 0
    End If
    
    lErro = RelDMPL_Calcula(ComboModelos.List(ComboModelos.ListIndex), ComboExercicio.ItemData(ComboExercicio.ListIndex), giFilialEmpresa, iPeriodo)
    If lErro <> SUCESSO Then Error 60664
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err
        
        Case 60662, 60663, 60664
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170159)

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
    If ComboOpcoes.Text = "" Then Error 60665

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 60666

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 60667

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 60668
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 60665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 60666, 60667, 60668
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170160)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 60669
    
    ComboOpcoes.Text = ""
    
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 60669
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170161)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click
    
    'se esta vazia
    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    'se não estiver carregando do BD
    If giCarregando = CANCELA Then
        
        'preenche a combo com periodo 1
        lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1)
        If lErro <> SUCESSO Then Error 13408
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 13408

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168185)

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

    giCarregando = CANCELA

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 60670

    For iIndice = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iIndice)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next

    lErro = MostraExercicio(giExercicioAtual)
    If lErro <> SUCESSO Then Error 60671

    'Le os Modelos na tabela RelDRE
    lErro = CF("RelDMPL_Le_Modelos_Distintos", RELDMPL, colModelos)
    If lErro <> SUCESSO Then Error 60672
    
    'preenche a combo Modelos
    For iIndice = 1 To colModelos.Count
        ComboModelos.AddItem colModelos.Item(iIndice)
    Next
        
    ComboModelos.ListIndex = -1
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 60670, 60671, 60672
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170162)

    End Select

    'Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Private Function RelDMPL_Calcula(ByVal sModelo As String, ByVal iExercicio As Integer, ByVal iFilialEmpresa As Integer, Optional ByVal iPeriodo As Integer) As Long
'Calcula Valor correspondente ao Modelo

Dim lErro As Long
Dim iOperacao As Integer
Dim dValorExercAtual As Double
Dim colRel As New Collection
Dim colRelConta As New Collection
Dim colRelFormula As New Collection
Dim objRel As New ClassRelDMPL
Dim objRel1 As New ClassRelDMPL
Dim objRelConta As New ClassRelDMPLConta
Dim objRelFormula As New ClassRelDMPLFormula
    
On Error GoTo Erro_RelDMPL_Calcula
    
    'Lê registros da tabela RelDMPL para o Modelo passado como parâmetro
    lErro = CF("RelDMPL_Le_Modelo", RELDMPL, sModelo, colRel)
    If lErro <> SUCESSO Then Error 60673
    
    'Lê registros da tabela RelDMPLConta para o Modelo passado como parâmetro
    lErro = CF("RelDMPLConta_Le_Modelo", RELDMPL, sModelo, colRelConta)
    If lErro <> SUCESSO Then Error 60674
    
    'Lê registros da tabela RelDMPLFormula para o Modelo passado como parâmetro
    lErro = CF("RelDMPLFormula_Le_Modelo", RELDMPL, sModelo, colRelFormula)
    If lErro <> SUCESSO Then Error 60675
    
    For Each objRel In colRel
        
        If objRel.iTipo = CEL_TIPO_CONTA Then
        
            For Each objRelConta In colRelConta
                
                'Se o elemento tiver a mesma linha e coluna do elemento da coleção colRel
                If objRelConta.iLinha = objRel.iLinha And objRelConta.iColuna = objRel.iColuna Then
                    
                    If iPeriodo = 0 Then
                        If objRel.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then
                            'Calcula o Valor
                            lErro = CF("MvPerCta_Calcula_Valor1", iFilialEmpresa, iExercicio - 1, objRelConta.sContaInicial, objRelConta.sContaFinal, dValorExercAtual)
                            If lErro <> SUCESSO Then Error 60679
                        Else
                            'Calcula o Valor
                            lErro = CF("MvPerCta_Calcula_Valor1", iFilialEmpresa, iExercicio, objRelConta.sContaInicial, objRelConta.sContaFinal, dValorExercAtual)
                            If lErro <> SUCESSO Then Error 60697
                        End If
                    Else
                         If objRel.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then
                            'Calcula o Valor
                            lErro = CF("MvPerCta_Calcula_Valor1_Periodo", iFilialEmpresa, iExercicio - 1, iPeriodo, objRelConta.sContaInicial, objRelConta.sContaFinal, dValorExercAtual)
                            If lErro <> SUCESSO Then Error 60679
                        Else
                            'Calcula o Valor
                            lErro = CF("MvPerCta_Calcula_Valor1_Periodo", iFilialEmpresa, iExercicio, iPeriodo, objRelConta.sContaInicial, objRelConta.sContaFinal, dValorExercAtual)
                            If lErro <> SUCESSO Then Error 60697
                        End If
                    End If
                    
                    'Aculmula o Valor
                    objRel.dValor = objRel.dValor + dValorExercAtual
                    
                End If
                
            Next
            
        ElseIf objRel.iTipo = CEL_TIPO_FORMULA Then
        
            iOperacao = REL_OPERACAO_SOMA
            
            For Each objRelFormula In colRelFormula
                
                'Se o elemento tiver a memsa linha e coluna do elemento da coleção colRel
                If objRelFormula.iLinha = objRel.iLinha And objRelFormula.iColuna = objRel.iColuna Then
                                        
                    For Each objRel1 In colRel
                        
                        'Pesquisar na coleção RelDRE o elemento que tenha o Código igual ao
                        'código da fórmula do elemento da coleção RelDREFormula
                        If objRelFormula.iLinhaFormula = objRel1.iLinha And objRelFormula.iColunaFormula = objRel1.iColuna Then
                        
                            If iOperacao = REL_OPERACAO_SOMA Then
                                objRel.dValor = objRel.dValor + objRel1.dValor
                            Else
                                objRel.dValor = objRel.dValor - objRel1.dValor
                            End If
                            
                            iOperacao = objRelFormula.iOperacao
                
                            Exit For
                            
                        End If
                        
                    Next
                        
                End If
                
            Next
        
        End If
        
    Next
    
    'Grava o Valor acumulado na tabela RelDMPL
    lErro = CF("RelDMPL_Grava_Valor", RELDMPL, colRel)
    If lErro <> SUCESSO Then Error 60692
    
    RelDMPL_Calcula = SUCESSO
    
    Exit Function
      
Erro_RelDMPL_Calcula:

    RelDMPL_Calcula = Err
    
    Select Case Err
    
        Case 60673, 60674, 60675, 60679, 60692, 60697
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170163)
        
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_MUTACAO_PL
    Set Form_Load_Ocx = Me
    Caption = "Demonstrativo de Mutação do Patrimônio Líquido"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpMutacaoPL"
    
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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

Function PreencheComboPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'lê os períodos do exercício 'iExercicio' preenchendo o combo de período
'seleciona o período 'iPeriodo'

Dim lErro As Long, iConta As Integer
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheComboPeriodo

    ComboPeriodo.Clear

    'inicializar os periodos do exercicio selecionado no combo de exercícios
    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 13391

    ComboPeriodo.AddItem " "
    ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = 0

    For iConta = 1 To colPeriodos.Count
        Set objPeriodo = colPeriodos.Item(iConta)
        ComboPeriodo.AddItem objPeriodo.sNomeExterno
        ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = objPeriodo.iPeriodo
    Next

'    'mostra o período
'    For iConta = 0 To ComboPeriodo.ListCount - 1
'        If ComboPeriodo.ItemData(iConta) = iPeriodo Then
'            ComboPeriodo.ListIndex = iConta
'            Exit For
'        End If
'    Next

    PreencheComboPeriodo = SUCESSO

    Exit Function

Erro_PreencheComboPeriodo:

    PreencheComboPeriodo = Err

    Select Case Err

        Case 13391

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168175)

    End Select

    Exit Function

End Function

Function MostraExercicioPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodo
    
    'seta cargando para que não execute o click
    giCarregando = OK

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next
    
    'mostra o período
    For iConta = 0 To ComboPeriodo.ListCount - 1
        If ComboPeriodo.ItemData(iConta) = iPeriodo Then
            ComboPeriodo.ListIndex = iConta
            Exit For
        End If
    Next

'    lErro = PreencheComboPeriodo(iExercicio, iPeriodo)
'    If lErro <> SUCESSO Then Error 13390

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = Err

    Select Case Err

        Case 13390

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168174)

    End Select

    Exit Function

End Function
