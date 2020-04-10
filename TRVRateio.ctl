VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRVRateio 
   ClientHeight    =   6870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   6870
   ScaleWidth      =   8190
   Begin VB.CheckBox CheckExcluiRateio 
      Caption         =   "Exclui rateios antigos gerados por esta rotina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   420
      TabIndex        =   16
      Top             =   840
      Width           =   6645
   End
   Begin VB.ComboBox Exercicio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1125
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   1590
   End
   Begin VB.ComboBox Periodo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   255
      Width           =   1590
   End
   Begin VB.CommandButton BotaoCcl 
      Caption         =   "Centros de Custo/Lucro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4455
      TabIndex        =   7
      Top             =   6435
      Width           =   2325
   End
   Begin VB.CommandButton BotaoPlanoConta 
      Caption         =   "Plano de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   855
      TabIndex        =   6
      Top             =   6435
      Width           =   1875
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5400
      ScaleHeight     =   495
      ScaleWidth      =   2610
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   135
      Width           =   2670
      Begin VB.CommandButton BotaoGerar 
         Height          =   345
         Left            =   1605
         Picture         =   "TRVRateio.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gerar os rateios"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TRVRateio.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoPesquisa 
         Height          =   360
         Left            =   585
         Picture         =   "TRVRateio.ctx":059C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Trazer a tela de cadastro"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2115
         Picture         =   "TRVRateio.ctx":08AE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TRVRateio.ctx":0A2C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.TextBox Descricao 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3360
      MaxLength       =   37
      TabIndex        =   4
      Top             =   1875
      Width           =   3405
   End
   Begin MSMask.MaskEdBox Ccl 
      Height          =   225
      Left            =   2220
      TabIndex        =   3
      Top             =   1785
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Conta 
      Height          =   225
      Left            =   405
      TabIndex        =   2
      Top             =   1830
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridRateio 
      Height          =   2400
      Left            =   330
      TabIndex        =   5
      Top             =   1215
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4233
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Exercício:"
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
      TabIndex        =   13
      Top             =   300
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2850
      TabIndex        =   12
      Top             =   315
      Width           =   750
   End
End
Attribute VB_Name = "TRVRateio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Descricao_Col As Integer

Dim objGridRateio As AdmGrid
Dim iAlterado As Integer
Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Public Sub Form_Load()

Dim lErro As Long
Dim objExercicio As ClassExercicio
Dim colExercicios As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Set objGridRateio = New AdmGrid
    
    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
    
    lErro = Inicializa_Grid_Rateio(objGridRateio)
    If lErro <> SUCESSO Then gError 197554
    
    'ler os exercicios do banco de dados
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then gError 197555

    'carrega combobox Exercicio
    For iIndice = 1 To colExercicios.Count
        Set objExercicio = colExercicios.Item(iIndice)
        Exercicio.AddItem objExercicio.sNomeExterno
        Exercicio.ItemData(Exercicio.NewIndex) = objExercicio.iExercicio
    Next
    
    lErro = Traz_Doc_Tela()
    If lErro <> SUCESSO Then gError 197627
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 197554, 197555, 197627
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197556)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long
Dim iExcluiRateio As Integer

On Error GoTo Erro_BotaGerar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 197644
    
    iExcluiRateio = CheckExcluiRateio.Value
    
    Call Chama_Tela_Modal("TRVImpRateioOff", iExcluiRateio)
    
    Call Limpa_Tela_Rateio

    iAlterado = 0

    
    Exit Sub
    
Erro_BotaGerar_Click:

    Select Case gErr
    
        Case 197641, 197644
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197642)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub Exercicio_Click()

Dim lErro As Long
Dim iExercicio As Integer

On Error GoTo Erro_Exercicio_Click

    If Exercicio.ListIndex = -1 Then
        Periodo.Clear
    Else
    
        iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
    
        lErro = Preenche_ComboPeriodo(iExercicio)
        If lErro <> SUCESSO Then gError 197557
    
        Periodo.ListIndex = -1
    End If
    
    Exit Sub
    
Erro_Exercicio_Click:

    Select Case gErr
    
        Case 197557
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197558)
        
    End Select
    
    Exit Sub
    
End Sub

Private Function Preenche_ComboPeriodo(iExercicio As Integer)

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo
Dim iIndice As Integer

On Error GoTo Erro_Preenche_ComboPeriodo

    lErro = CF("Periodo_Le_Todos_Exercicio", giFilialEmpresa, iExercicio, colPeriodos)
    If lErro <> SUCESSO Then gError 197559

    Periodo.Clear

    For Each objPeriodo In colPeriodos

        Periodo.AddItem objPeriodo.sNomeExterno
        Periodo.ItemData(Periodo.NewIndex) = objPeriodo.iPeriodo

    Next

    Exit Function

Erro_Preenche_ComboPeriodo:

    Select Case gErr

        Case 197559

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197560)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Rateio(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Rateio
    
    'Form do Grid
    Set objGridInt.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Descrição")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    
    iGrid_Conta_Col = 1
    iGrid_Ccl_Col = 2
    iGrid_Descricao_Col = 3
    
    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then gError 197561
    
    objGridInt.objGrid = GridRateio
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 1001
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20
        
    GridRateio.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Rateio = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Rateio:

    Inicializa_Grid_Rateio = gErr
    
    Select Case gErr
    
        Case 197561
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 197562)
        
    End Select

    Exit Function
        
End Function

Private Function Inicializa_Mascaras() As Long
'inicializa as mascaras de conta e centro de custo

Dim sMascaraConta As String
Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascaras
   
    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 197563
    
    Conta.Mask = sMascaraConta
   
    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 197564
    
    Ccl.Mask = sMascaraCcl

    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = gErr
    
    Select Case gErr
    
        Case 197563, 197564
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 197565)
        
    End Select

    Exit Function
    
End Function

Private Sub BotaoPesquisa_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoPesquisa_Click
    
    lErro = Traz_Doc_Tela()
    If lErro <> SUCESSO Then gError 197566
    
    Exit Sub
    
Erro_BotaoPesquisa_Click:

    Select Case gErr
    
        Case 197566
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197567)
        
    End Select
    
    Exit Sub

End Sub

Private Sub Conta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRateio)
      
End Sub

Private Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRateio)

End Sub

Private Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRateio.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Ccl_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridRateio)
      
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRateio)

End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRateio.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Descricao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRateio)

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridRateio)
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRateio.objControle = Descricao
    lErro = Grid_Campo_Libera_Foco(objGridRateio)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Function Trata_Parametros() As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197568)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Sub BotaoCcl_Click()
'Quando clicada ativa a tela de browser RateioOffLista

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
Dim sCclOrigem As String
Dim iCclPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoCcl_Click
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)
    
    Exit Sub
    
Erro_BotaoCcl_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197569)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim objCcl As ClassCcl
Dim sCclMascarado As String
Dim lErro As Long
    
On Error GoTo Erro_objEventoCcl_evSelecao
    
    Set objCcl = obj1
    
    'mascara o centro de custo
    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 197570
    
    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True
    
    GridRateio.TextMatrix(GridRateio.Row, iGrid_Ccl_Col) = Ccl.Text
    
    iAlterado = 0
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197570
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 197571)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub GridRateio_LeaveCell()
    Call Saida_Celula(objGridRateio)
End Sub

Private Sub GridRateio_EnterCell()
    Call Grid_Entrada_Celula(objGridRateio, iAlterado)
End Sub

Private Sub GridRateio_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridRateio, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRateio, iAlterado)
    End If
    
End Sub

Private Sub GridRateio_GotFocus()
    Call Grid_Recebe_Foco(objGridRateio)
End Sub

Private Sub GridRateio_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridRateio)
    
End Sub

Private Sub GridRateio_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRateio, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRateio, iAlterado)
    End If

End Sub

Private Sub GridRateio_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridRateio)

End Sub

Private Sub GridRateio_RowColChange()

    Call Grid_RowColChange(objGridRateio)
       
End Sub

Private Sub GridRateio_Scroll()

    Call Grid_Scroll(objGridRateio)
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridRateio.Col
    
            Case iGrid_Conta_Col
            
                lErro = Saida_Celula_Conta(objGridInt)
                If lErro <> SUCESSO Then gError 197572
                
            Case iGrid_Ccl_Col
            
                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then gError 197573
                
            Case iGrid_Descricao_Col
            
                lErro = Saida_Celula_Descricao(objGridInt)
                If lErro <> SUCESSO Then gError 197574
                
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 197575
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
            
        Case 197572 To 197574
    
        Case 197575
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197576)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim sContaFormatada As String
Dim sContaMascarada As String
Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica", Conta.Text, Conta.ClipText, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then gError 197577
    
    'se é uma conta simples, coloca a conta normal no lugar da conta simples
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 197578
        
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
        
    'se não encontrou a conta simples
    ElseIf lErro = 44033 Or lErro = 44037 Then

        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 197579
                    
        'conta não cadastrada
        If lErro = 5700 Then gError 197580
                
    End If
    
    'Verifica se a conta foi preenchida
    If Len(Conta.ClipText) > 0 Then
    
'        'verifica se a conta passada como parametro coincide com a conta credito
'        lErro = Testa_Conta_Credito(Conta.Text)
'        If lErro <> SUCESSO Then Error 55761
'
        'verifica se a conta passada como parametro coincide com as contas em lançamentos
'        lErro = Testa_Conta_Grid_Contas(Conta.Text, TESTA_LINHA_ATUAL)
'        If lErro <> SUCESSO Then Error 55763
        
        'verifica se a conta passada como parametro tem associacao com o centro de custo em questao
        lErro = Testa_Assoc_ContaCcl(sContaFormatada, objContaCcl)
        If lErro <> SUCESSO And lErro <> 197596 Then gError 197581
        
        'se está faltando a associacao da conta com o centro de custo
        If lErro <> SUCESSO Then gError 197582
        
        If GridRateio.Row - GridRateio.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
                        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 197583

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = gErr
    
    Select Case gErr
            
        Case 197577, 197579, 197581, 197583
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 197578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 197580
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Conta.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
                
            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
                      
        Case 197582
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", Conta.Text, GridRateio.TextMatrix(GridRateio.Row, iGrid_Ccl_Col))

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197584)
    
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula ccl do grid que está deixando de ser a corrente

Dim sCclFormatada As String
Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sConta As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl
                
    'critica o formato do ccl, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 197585
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then gError 197586
                
    'se o centro de custo foi preenchido
    If Len(Ccl.ClipText) > 0 Then
    
        'se a conta foi informada
        If Len(GridRateio.TextMatrix(GridRateio.Row, iGrid_Conta_Col)) > 0 Then
    
            'verificar se a associação da conta com o centro de custo em questão está cadastrada
            sConta = GridRateio.TextMatrix(GridRateio.Row, iGrid_Conta_Col)
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError 197587
        
            If iContaPreenchida = CONTA_PREENCHIDA Then
        
                objContaCcl.sConta = sContaFormatada
                objContaCcl.sCcl = sCclFormatada
            
                lErro = CF("ContaCcl_Le", objContaCcl)
                If lErro <> SUCESSO And lErro <> 5871 Then gError 197588
            
                'associação Conta x Centro de Custo/Lucro não cadastrada
                If lErro = 5871 Then gError 197589
        
            End If
        
        End If
                        
        If GridRateio.Row - GridRateio.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 197590

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr
    
    Select Case gErr
    
        Case 197585, 197587, 197588, 197590
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 197586
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
            
        Case 197589
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, Ccl.Text)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197591)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Descricao(objGridInt As AdmGrid) As Long
'faz a critica da celula de percentual do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim dValor As Double

On Error GoTo Erro_Saida_Celula_Descricao

    Set objGridInt.objControle = Descricao
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 197592
                        
    Saida_Celula_Descricao = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Descricao:

    Saida_Celula_Descricao = gErr
    
    Select Case gErr
    
        Case 197592
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197593)
        
    End Select

    Exit Function

End Function

Private Function Testa_Assoc_ContaCcl(sContaFormatada As String, objContaCcl As ClassContaCcl) As Long
'verifica se a conta passada como parametro tem associacao com o centro de custo em questao

Dim lErro As Long
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim sCcl As String

On Error GoTo Erro_Testa_Assoc_ContaCcl

    'se o centro de custo foi preenchido
    If Len(GridRateio.TextMatrix(GridRateio.Row, iGrid_Ccl_Col)) > 0 Then
    
        'verifica se a associação da conta com o centro de custo está cadastrado
        sCcl = GridRateio.TextMatrix(GridRateio.Row, iGrid_Ccl_Col)

        lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 197594

        objContaCcl.sConta = sContaFormatada
        objContaCcl.sCcl = sCclFormatada

        lErro = CF("ContaCcl_Le", objContaCcl)
        If lErro <> SUCESSO And lErro <> 5871 Then gError 197595

        'associação Conta x Centro de Custo/Lucro não cadastrada
        If lErro = 5871 Then gError 197596
        
    End If
        
    Testa_Assoc_ContaCcl = SUCESSO
    
    Exit Function

Erro_Testa_Assoc_ContaCcl:

    Testa_Assoc_ContaCcl = gErr

    Select Case gErr

        Case 197594 To 197596

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197597)

    End Select
    
    Exit Function

End Function

Private Function Traz_Doc_Tela() As Long
'traz os dados do rateio do banco de dados para a tela

Dim lErro As Long
Dim colTRVRateio As New Collection
Dim objTRVRateio As ClassTRVRateio
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim dTotal As Double
Dim iIndice As Integer
Dim objCcl As New ClassCcl
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Traz_Doc_Tela

    Call Limpa_Tela_Rateio
        
    'Coloca em colRateioOff Todos os dados do rateio passado em objRateioOff
    lErro = CF("TRVRateio_Le", colTRVRateio)
    If lErro <> SUCESSO Then gError 197598
        
    If colTRVRateio.Count > 0 Then
    
        Set objTRVRateio = colTRVRateio(1)
        
        'mostra o Exercicio
        For iIndice = 0 To Exercicio.ListCount - 1
            If Exercicio.ItemData(iIndice) = objTRVRateio.iExercicio Then
                Exercicio.ListIndex = iIndice
                Exit For
            End If
        Next
            
        'mostra o periodo
        For iIndice = 0 To Periodo.ListCount - 1
            If Periodo.ItemData(iIndice) = objTRVRateio.iPeriodo Then
                Periodo.ListIndex = iIndice
                Exit For
            End If
        Next
        
        iIndice = 0
        
        For Each objTRVRateio In colTRVRateio
         
            iIndice = iIndice + 1
         
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
             
            lErro = Mascara_RetornaContaEnxuta(objTRVRateio.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 197599
             
            'move as conta envolvidas no rateio para o Grid
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
             
            'coloca a conta na tela
            GridRateio.TextMatrix(iIndice, iGrid_Conta_Col) = Conta.Text
                      
             
            'mascara o centro de custo
            sCclMascarado = String(STRING_CCL, 0)
               
            If objTRVRateio.sCcl <> "" Then
            
                lErro = Mascara_MascararCcl(objTRVRateio.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then gError 197600
        
            Else
              sCclMascarado = ""
                   
            End If
                    
            'coloca o centro de custo na tela
            GridRateio.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado
                
            'coloca o percentual na tela
            GridRateio.TextMatrix(iIndice, iGrid_Descricao_Col) = objTRVRateio.sDescricao
                
            objGridRateio.iLinhasExistentes = objGridRateio.iLinhasExistentes + 1
                
        Next
     
    End If
     
    iAlterado = 0
    
    Traz_Doc_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = gErr

    Select Case gErr
        
        Case 197598
        
        Case 197599
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objTRVRateio.sConta)
            
        Case 197600
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objTRVRateio.sCcl)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197601)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoCcl = Nothing
    Set objEventoConta = Nothing
    Set objGridRateio = Nothing
    
End Sub

Private Sub BotaoGravar_Click()
    
Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 197602
    
    Call Limpa_Tela_Rateio

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 197602
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 197603)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
            
Dim lErro As Long
Dim colTRVRateio As New Collection
Dim colContas As New Collection
Dim objTRVRateio As New ClassTRVRateio
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim sContaOrigem As String
Dim sContaCredito As String
Dim sCclOrigem As String
Dim sConta As String
Dim dTotal As Double

On Error GoTo Erro_Gravar_Registro
            
    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Exercicio.Text) = 0 Then gError 197604
    
    If Len(Periodo.Text) = 0 Then gError 197605
    
    'Verifica se pelo menos uma linha do Grid está preenchida
    If objGridRateio.iLinhasExistentes = 0 Then gError 197606
    
    'Preenche a colTRVRateio com as informacoes contidas no Grid
    lErro = Grid_Rateio(colTRVRateio)
    If lErro <> SUCESSO Then gError 197607
           
    lErro = CF("TRVRateio_Grava", colTRVRateio)
    If lErro <> SUCESSO Then gError 197608
         
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 197604
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_SELECIONADO", gErr)

        Case 197605
            Call Rotina_Erro(vbOKOnly, "ERRO_PERIODO_VAZIO", gErr)
        
        Case 197606
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDITENS_VAZIO", gErr)
        
        Case 197607, 197608
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 197609)
            
    End Select
    
    Exit Function
    
End Function

Private Function Grid_Rateio(colTRVRateio As Collection) As Long
'Armazena os dados do grid em colTRVRateio

Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim objTRVRateio As ClassTRVRateio
Dim objTRVRateio1 As ClassTRVRateio
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_Grid_Rateio

    For iIndice1 = 1 To objGridRateio.iLinhasExistentes
        
        Set objTRVRateio = New ClassTRVRateio
            
        objTRVRateio.iSeq = StrParaInt(GridRateio.TextMatrix(iIndice1, 0))
            
        objTRVRateio.iExercicio = Exercicio.ItemData(Exercicio.ListIndex)
        
        objTRVRateio.iPeriodo = Periodo.ItemData(Periodo.ListIndex)
            
        sConta = GridRateio.TextMatrix(iIndice1, iGrid_Conta_Col)
        
        If Len(Trim(sConta)) = 0 Then gError 197610
        
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 197611
            
        'Armazena a conta
        objTRVRateio.sConta = sContaFormatada
    
        'Armazena a descricao
        objTRVRateio.sDescricao = GridRateio.TextMatrix(iIndice1, iGrid_Descricao_Col)
        
        If Len(objTRVRateio.sDescricao) = 0 Then gError 197612
                        
        sCcl = GridRateio.TextMatrix(iIndice1, iGrid_Ccl_Col)
        
        lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
        If lErro <> SUCESSO Then gError 197613
            
        If iCclPreenchida <> CCL_PREENCHIDA Then gError 197614
        
        objTRVRateio.sCcl = sCclFormatada
                        
        iIndice2 = 0
        
        For Each objTRVRateio1 In colTRVRateio
        
            iIndice2 = iIndice2 + 1
            
            If objTRVRateio1.sConta = objTRVRateio.sConta And objTRVRateio1.sCcl = objTRVRateio.sCcl Then gError 197615
            
            If objTRVRateio1.sDescricao = objTRVRateio.sDescricao Then gError 197616
            
        Next
            
        colTRVRateio.Add objTRVRateio
        
    Next
    
    Grid_Rateio = SUCESSO

    Exit Function

Erro_Grid_Rateio:

    Grid_Rateio = gErr

    Select Case gErr
    
        Case 197610
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_GRID_NAO_PREENCHIDA", gErr, iIndice1)
    
        Case 197611, 197613
                
        Case 197612
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_GRID_NAO_PREENCHIDO", gErr, iIndice1)
        
        Case 197614
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_GRID_NAO_PREENCHIDO", gErr, iIndice1)
        
        Case 197615
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_GRID_DUPLICADO", gErr, iIndice1, iIndice2)
        
        Case 197616
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_GRID_DUPLICADO", gErr, iIndice1, iIndice2)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197617)
            
    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim iCodigo As Integer
Dim lErro As Long
Dim objRateioOff As New ClassRateioOff
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 197618

    Call Limpa_Tela_Rateio

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 197618

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 197619)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Sub Limpa_Tela_Rateio()

Dim lErro As Long

    Call Grid_Limpa(objGridRateio)
    
    Exercicio.ListIndex = -1
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
'    Parent.HelpContextID =
    Set Form_Load_Ocx = Me
    Caption = "Geração de Novos Rateios"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRVRateio"
    
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
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Conta Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Conta Then
            Call BotaoPlanoConta_Click
        End If
    
    End If

End Sub

Private Sub BotaoPlanoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaMascarada As String

On Error GoTo Erro_objEventoConta_evSelecao

    Set objPlanoConta = obj1
    
    'mascara a conta
    sContaMascarada = String(STRING_CONTA, 0)
     
    lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
    If lErro <> SUCESSO Then gError 197620
     
    'move as conta envolvidas no rateio para o Grid
    Conta.PromptInclude = False
    Conta.Text = sContaMascarada
    Conta.PromptInclude = True
     
    'coloca a conta na tela
    GridRateio.TextMatrix(GridRateio.Row, iGrid_Conta_Col) = Conta.Text
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197620
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197621)
        
    End Select

    Exit Sub

End Sub

