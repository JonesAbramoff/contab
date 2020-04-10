VERSION 5.00
Begin VB.UserControl RelOpDemResAcumOcx 
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   ScaleHeight     =   2280
   ScaleWidth      =   6285
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
      Left            =   4035
      Picture         =   "RelOpDemResAcumOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   930
      Width           =   1815
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpDemResAcumOcx.ctx":0102
      Left            =   1035
      List            =   "RelOpDemResAcumOcx.ctx":0104
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   270
      Width           =   2655
   End
   Begin VB.TextBox NivelMaximo 
      Height          =   285
      Left            =   5550
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1875
      Width           =   255
   End
   Begin VB.CheckBox CheckAnaliticas 
      Caption         =   "Exibir Contas Analíticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1830
      Width           =   2385
   End
   Begin VB.ComboBox ComboExercicio 
      Height          =   315
      ItemData        =   "RelOpDemResAcumOcx.ctx":0106
      Left            =   1035
      List            =   "RelOpDemResAcumOcx.ctx":0108
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   795
      Width           =   1860
   End
   Begin VB.ComboBox ComboPeriodo 
      Height          =   315
      ItemData        =   "RelOpDemResAcumOcx.ctx":010A
      Left            =   1425
      List            =   "RelOpDemResAcumOcx.ctx":010C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1350
      Width           =   1470
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpDemResAcumOcx.ctx":010E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpDemResAcumOcx.ctx":028C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpDemResAcumOcx.ctx":07BE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpDemResAcumOcx.ctx":0948
         Style           =   1  'Graphical
         TabIndex        =   1
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
      Left            =   360
      TabIndex        =   14
      Top             =   315
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Até o Período:"
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
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   1395
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Nível Máximo de Conta:"
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
      Left            =   3450
      TabIndex        =   12
      Top             =   1890
      Width           =   2055
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
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "RelOpDemResAcumOcx"
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

Function MostraExercicioPeriodo(iExercicio As Integer, iPeriodo As Integer) As Long
'mostra o exercício 'iExercicio' no combo de exercícios
'chama PreencheComboPeriodo

Dim iConta As Integer, lErro As Long

On Error GoTo Erro_MostraExercicioPeriodo

    giCarregando = OK

    For iConta = 0 To ComboExercicio.ListCount - 1
        If ComboExercicio.ItemData(iConta) = iExercicio Then
            ComboExercicio.ListIndex = iConta
            Exit For
        End If
    Next

    lErro = PreencheComboPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 13343

    MostraExercicioPeriodo = SUCESSO

    Exit Function

Erro_MostraExercicioPeriodo:

    MostraExercicioPeriodo = Err

    Select Case Err

        Case 13343

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167274)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO Then Error 13344

    For iConta = 1 To colPeriodos.Count
        Set objPeriodo = colPeriodos.Item(iConta)
        ComboPeriodo.AddItem objPeriodo.sNomeExterno
        ComboPeriodo.ItemData(ComboPeriodo.NewIndex) = objPeriodo.iPeriodo
    Next

    'mostra o período
    For iConta = 0 To ComboPeriodo.ListCount - 1
        If ComboPeriodo.ItemData(iConta) = iPeriodo Then
            ComboPeriodo.ListIndex = iConta
            Exit For
        End If
    Next

    PreencheComboPeriodo = SUCESSO

    Exit Function

Erro_PreencheComboPeriodo:

    PreencheComboPeriodo = Err

    Select Case Err

        Case 13344

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167275)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 16937
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
       
    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 13369
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 13369
        
        Case 16937
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167276)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a expressão de seleção

Dim sExpressao As String
Dim lErro As Long
Dim sExprCtasPatr As String

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If CheckAnaliticas.Value = 0 Then sExpressao = "TipoConta=1"

    If Len(NivelMaximo.Text) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "NivelConta <= " & Forprint_ConvInt(CInt(NivelMaximo.Text))

    End If

'    lErro = CF("Monta_Filtro_Ctas_Patrimoniais", objRelOpcoes, sExprCtasPatr)
'    If lErro <> SUCESSO Then Error 7165
'
'    If Len(sExprCtasPatr) <> 0 Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & sExprCtasPatr
'
'    End If
        
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case 7165
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167277)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sCheck As String

On Error GoTo Erro_PreencherRelOp

    'exercício não pode ser vazio
    If ComboExercicio.Text = "" Then Error 13362

    'período não pode ser vazio
    If ComboPeriodo.Text = "" Then Error 13363

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 13346

    'transforma o valor do check box CheckAnaliticas em "S" ou "N"
    If CheckAnaliticas.Value = 0 Then
        sCheck = "N"
    Else
        sCheck = "S"
    End If
    
    lErro = objRelOpcoes.IncluirParametro("TANALITICA", sCheck)
    If lErro <> AD_BOOL_TRUE Then Error 13347
    
    lErro = objRelOpcoes.IncluirParametro("NCTANIVMAX", NivelMaximo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 13348

    lErro = objRelOpcoes.IncluirParametro("NEXERCICIO", CStr(ComboExercicio.ItemData(ComboExercicio.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13349

    lErro = objRelOpcoes.IncluirParametro("TTITAUX1", ComboExercicio.Text)
    If lErro <> AD_BOOL_TRUE Then Error 19395

    lErro = objRelOpcoes.IncluirParametro("NPERIODO", CStr(ComboPeriodo.ItemData(ComboPeriodo.ListIndex)))
    If lErro <> AD_BOOL_TRUE Then Error 13350

    lErro = objRelOpcoes.IncluirParametro("TTITAUX2", ComboPeriodo.Text)
    If lErro <> AD_BOOL_TRUE Then Error 19396

    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then Error 13351

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err
    
    Select Case Err

        Case 13362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_VAZIO", Err)
            ComboExercicio.SetFocus

        Case 13363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", Err)
            ComboPeriodo.SetFocus

        Case 13346

        Case 13347, 13348, 13349, 13350, 19395, 19396

        Case 13351

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167278)
            
    End Select
    
    Exit Function
    
End Function

Private Sub NivelMaximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NivelMaximo_Validate

    'lote inicial deve estar entre 1 e 9999
    If NivelMaximo.Text <> "" Then
        lErro = Valor_Critica(NivelMaximo.Text)
        If lErro <> SUCESSO Then Error 54886
    End If

    Exit Sub

Erro_NivelMaximo_Validate:

    Cancel = True


    Select Case Err
        
        Case 54886
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167279)

    End Select

    Exit Sub
    
End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long, iExercicio As Integer, iPeriodo As Integer
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then Error 13352

    'imprimir contas analíticas
    lErro = objRelOpcoes.ObterParametro("TANALITICA", sParam)
    If lErro <> SUCESSO Then Error 13353

    If sParam = "S" Then CheckAnaliticas.Value = 1
    If sParam = "N" Then CheckAnaliticas.Value = 0
    
    'limitar nível máximo de conta
    lErro = objRelOpcoes.ObterParametro("NCTANIVMAX", sParam)
    If lErro <> SUCESSO Then Error 13354

    NivelMaximo.Text = sParam

    'exercício
    lErro = objRelOpcoes.ObterParametro("NEXERCICIO", sParam)
    If lErro <> SUCESSO Then Error 13355

    iExercicio = CInt(sParam)

    'período
    lErro = objRelOpcoes.ObterParametro("NPERIODO", sParam)
    If lErro <> SUCESSO Then Error 13356

    iPeriodo = CInt(sParam)

    lErro = MostraExercicioPeriodo(iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 13357

    PreencherParametrosNaTela = SUCESSO

    Exit Function
    
Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err
    
    Select Case Err

        Case 13352

        Case 13353, 13354, 13355, 13356

        Case 13357

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167280)

    End Select
    
    Exit Function
    
End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 13358

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 13359

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex
    
        'limpa os campos da tela
        lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then Error 47035
        
        CheckAnaliticas.Value = 0

    End If

    Exit Sub
    
Erro_BotaoExcluir_Click:
    
    Select Case Err

        Case 13358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 13359, 47035

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167281)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13360

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 13360

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167282)

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
    If ComboOpcoes.Text = "" Then Error 13361

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 13364

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 13365

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then Error 47033
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 13361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 13364

        Case 13365, 47033
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167283)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then Error 47032
    
    CheckAnaliticas.Value = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 47032
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167284)

    End Select

    Exit Sub

End Sub

Private Sub ComboExercicio_Click()

Dim lErro As Long

On Error GoTo Erro_ComboExercicio_Click

    If ComboExercicio.ListIndex = -1 Then Exit Sub
    
    If giCarregando = CANCELA Then

        lErro = PreencheComboPeriodo(ComboExercicio.ItemData(ComboExercicio.ListIndex), 1)
        If lErro <> SUCESSO Then Error 13366
    
    End If
    
    giCarregando = CANCELA
    
    Exit Sub

Erro_ComboExercicio_Click:

    Select Case Err

        Case 13366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167285)

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

Dim lErro As Long, iConta As Integer
Dim objExercicio As ClassExercicio
Dim colExerciciosAbertos As New Collection

On Error GoTo Erro_Form_Load

    giCarregando = CANCELA

    'ler os exercicios abertos
    lErro = CF("Exercicios_Le_Todos", colExerciciosAbertos)
    If lErro <> SUCESSO Then Error 13370
    
    For iConta = 1 To colExerciciosAbertos.Count
        Set objExercicio = colExerciciosAbertos.Item(iConta)
        ComboExercicio.AddItem objExercicio.sNomeExterno
        ComboExercicio.ItemData(ComboExercicio.NewIndex) = objExercicio.iExercicio
    Next
    
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 13370

        Case 13371

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167286)

    End Select

    Unload Me

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_BALANCO_PATR_PERIODO
    Set Form_Load_Ocx = Me
    Caption = "Demonstração do Resultado Acumulado"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpDemResAcum"
    
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


