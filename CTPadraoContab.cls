VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPadraoContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Const KEYCODE_VERIFICAR_SINTAXE = vbKeyF5

Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim iGrid_Aglutina_Col As Integer
Dim iGrid_Gerencial_Col As Integer
Dim iGrid_ContaCusto_Col As Integer
Dim iGrid_CclCusto_Col As Integer
Dim iGrid_DebitoCusto_Col As Integer
Dim iGrid_CreditoCusto_Col As Integer
Dim iGrid_HistoricoCusto_Col As Integer
Dim iGrid_ProdutoCusto_Col As Integer
Dim iGrid_AglutinaCusto_Col As Integer
Dim iGrid_GerencialCusto_Col As Integer
Dim objGridCusto As AdmGrid
Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim sModulo As String
Dim iTransacao As Integer
Dim iSubTipo As Integer
Dim sModelo As String

Const CONTABILIZACAO_OBRIGATORIA = 1
Const CONTABILIZACAO_NAO_OBRIGATORIA = 0

Const TAB_LANCAMENTOS = 1
Const TAB_CUSTOS = 2

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoHistorico As AdmEvento
Attribute objEventoHistorico.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Debito() As Object
     Set Debito = objUserControl.Controls("Debito")
End Property

Public Property Get Credito() As Object
     Set Credito = objUserControl.Controls("Credito")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get Conta() As Object
     Set Conta = objUserControl.Controls("Conta")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get Aglutina() As Object
     Set Aglutina = objUserControl.Controls("Aglutina")
End Property

Public Property Get Gerencial() As Object
     Set Gerencial = objUserControl.Controls("Gerencial")
End Property

Public Property Get GridLancamentos() As Object
     Set GridLancamentos = objUserControl.Controls("GridLancamentos")
End Property

Public Property Get HistoricoCusto() As Object
     Set HistoricoCusto = objUserControl.Controls("HistoricoCusto")
End Property

Public Property Get ContaCusto() As Object
     Set ContaCusto = objUserControl.Controls("ContaCusto")
End Property

Public Property Get CclCusto() As Object
     Set CclCusto = objUserControl.Controls("CclCusto")
End Property

Public Property Get CreditoCusto() As Object
     Set CreditoCusto = objUserControl.Controls("CreditoCusto")
End Property

Public Property Get DebitoCusto() As Object
     Set DebitoCusto = objUserControl.Controls("DebitoCusto")
End Property

Public Property Get AglutinaCusto() As Object
     Set AglutinaCusto = objUserControl.Controls("AglutinaCusto")
End Property

Public Property Get GerencialCusto() As Object
     Set GerencialCusto = objUserControl.Controls("GerencialCusto")
End Property

Public Property Get ProdutoCusto() As Object
     Set ProdutoCusto = objUserControl.Controls("ProdutoCusto")
End Property

Public Property Get GridCusto() As Object
     Set GridCusto = objUserControl.Controls("GridCusto")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Modulo() As Object
     Set Modulo = objUserControl.Controls("Modulo")
End Property

Public Property Get Transacao() As Object
     Set Transacao = objUserControl.Controls("Transacao")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get BotaoCcl() As Object
     Set BotaoCcl = objUserControl.Controls("BotaoCcl")
End Property

Public Property Get BotaoHistorico() As Object
     Set BotaoHistorico = objUserControl.Controls("BotaoHistorico")
End Property

Public Property Get Modelo() As Object
     Set Modelo = objUserControl.Controls("Modelo")
End Property

Public Property Get Operadores() As Object
     Set Operadores = objUserControl.Controls("Operadores")
End Property

Public Property Get Funcoes() As Object
     Set Funcoes = objUserControl.Controls("Funcoes")
End Property

Public Property Get Mnemonicos() As Object
     Set Mnemonicos = objUserControl.Controls("Mnemonicos")
End Property

Public Property Get BotaoProduto() As Object
     Set BotaoProduto = objUserControl.Controls("BotaoProduto")
End Property

Public Property Get Checkbox_Verifica_Sintaxe() As Object
     Set Checkbox_Verifica_Sintaxe = objUserControl.Controls("Checkbox_Verifica_Sintaxe")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get Padrao() As Object
     Set Padrao = objUserControl.Controls("Padrao")
End Property

Public Property Get Obrigatoriedade() As Object
    Set Obrigatoriedade = objUserControl.Controls("Obrigatoriedade")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoCcl_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Public Sub BotaoHistorico_Click()

Dim colSelecao As Collection
Dim objHistPadrao As New ClassHistPadrao

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHistorico)

End Sub

Public Sub BotaoPlanoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)
    
End Sub

Public Sub BotaoProduto_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

End Sub

Public Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CclCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CclCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCusto)
    
End Sub

Public Sub CclCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)
    
End Sub

Public Sub CclCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = CclCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Conta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
End Sub

Public Sub Conta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Public Sub Conta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Conta
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
End Sub

Public Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ContaCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContaCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCusto)
    
End Sub

Public Sub ContaCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)
    
End Sub

Public Sub ContaCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = ContaCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Credito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Credito_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Public Sub Credito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Public Sub Credito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Credito
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CreditoCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CreditoCusto_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGridCusto)

End Sub

Public Sub CreditoCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)

End Sub

Public Sub CreditoCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = CreditoCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Debito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Debito_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Public Sub Debito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Public Sub Debito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Debito
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DebitoCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DebitoCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCusto)

End Sub

Public Sub DebitoCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)

End Sub

Public Sub DebitoCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = DebitoCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Aglutina_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Public Sub Aglutina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Public Sub Aglutina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Aglutina
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AglutinaCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCusto)

End Sub

Public Sub AglutinaCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)

End Sub

Public Sub AglutinaCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = AglutinaCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Gerencial_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Public Sub Gerencial_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Public Sub Gerencial_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Gerencial
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GerencialCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCusto)

End Sub

Public Sub GerencialCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)

End Sub

Public Sub GerencialCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = GerencialCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub ProdutoCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ProdutoCusto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCusto)

End Sub

Public Sub ProdutoCusto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)

End Sub

Public Sub ProdutoCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = ProdutoCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Set objEventoConta = Nothing
    Set objEventoHistorico = Nothing
    Set objEventoCcl = Nothing
    Set objEventoProduto = Nothing
    Set objGrid1 = Nothing
    Set objGridCusto = Nothing
    
End Sub

Public Sub Funcoes_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaFuncao As New ClassFormulaFuncao
Dim lPos As Long
Dim sFuncao As String
    
On Error GoTo Erro_Funcoes_Click
    
    objFormulaFuncao.sFuncaoCombo = Funcoes.Text
    
    'retorna os dados da funcao passada como parametro
    lErro = CF("FormulaFuncao_Le", objFormulaFuncao)
    If lErro <> SUCESSO And lErro <> 36088 Then Error 36216
    
    Descricao.Text = objFormulaFuncao.sFuncaoDesc
    
    lPos = InStr(1, Funcoes.Text, "(")
    If lPos = 0 Then
        sFuncao = Funcoes.Text
    Else
        sFuncao = Mid(Funcoes.Text, 1, lPos)
    End If
    
    lErro = Funcoes1(sFuncao)
    If lErro <> SUCESSO Then Error 33604
    
    Exit Sub
    
Erro_Funcoes_Click:

    Select Case Err
    
        Case 33604, 36216
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157859)
            
    End Select
        
    Exit Sub

End Sub

Public Sub GridCusto_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridCusto, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCusto, iAlterado)
    End If

End Sub

Public Sub GridCusto_EnterCell()
    
    Call Grid_Entrada_Celula(objGridCusto, iAlterado)

End Sub

Public Sub GridCusto_GotFocus()
    
    Call Grid_Recebe_Foco(objGridCusto)

End Sub

Public Sub GridCusto_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCusto)

End Sub

Public Sub GridCusto_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCusto, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        
        Call Grid_Entrada_Celula(objGridCusto, iAlterado)
        
    End If

End Sub

Public Sub GridCusto_LeaveCell()
    
    Call Saida_Celula(objGridCusto)

End Sub

Public Sub GridCusto_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridCusto)

End Sub

Public Sub GridCusto_RowColChange()

    Call Grid_RowColChange(objGridCusto)

End Sub

Public Sub GridCusto_Scroll()

    Call Grid_Scroll(objGridCusto)

End Sub

Public Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Historico_GotFocus()
    
Dim iPos As Integer
    
    Call Grid_Campo_Recebe_Foco(objGrid1)
    
    If Len(Historico.Text) > 0 Then
        iPos = InStr(Historico.Text, CARACTER_HISTORICO_PARAM)
        If iPos > 0 Then
            Historico.SelStart = iPos - 1
            Historico.SelLength = 1
        End If
    End If
    
End Sub

Public Sub Historico_KeyPress(KeyAscii As Integer)

Dim iInicio As Integer
Dim iPos As Integer
Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Historico_KeyPress

    'se digitou ENTER
    If KeyAscii = vbKeyReturn Then
        
        If Len(Historico.Text) > 0 Then
        
            If Left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(Historico.Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44076
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44077
                
                If lErro = 5446 Then Error 44078
        
                Historico.Text = Chr(34) & objHistPadrao.sDescHistPadrao & Chr(34)
                Historico.SelStart = 0
                
            End If
    
            If Historico.SelText = CARACTER_HISTORICO_PARAM Then
                iInicio = Historico.SelStart + 2
            Else
                iInicio = Historico.SelStart
            End If
        
            If iInicio = 0 Then iInicio = 1
        
            iPos = InStr(iInicio, Historico.Text, CARACTER_HISTORICO_PARAM)
            If iPos > 0 Then
                Historico.SelStart = iPos - 1
                Historico.SelLength = 1
                Exit Sub
            End If
        End If
    End If

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
    Exit Sub
    
Erro_Historico_KeyPress:

    Select Case Err
    
        Case 44076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
        
        Case 44077

        Case 44078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157860)
    
    End Select

    Exit Sub
    
End Sub

Public Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Historico
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridLancamentos_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    
End Sub

Public Sub GridLancamentos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Public Sub GridLancamentos_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Public Sub GridLancamentos_LeaveCell()
    
    Call Saida_Celula(objGrid1)
    
End Sub

Public Sub GridLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid1)
    
End Sub

Public Sub GridLancamentos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Public Sub GridLancamentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Public Sub GridLancamentos_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Public Sub GridLancamentos_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = TAB_LANCAMENTOS
    
    Set objEventoConta = New AdmEvento
    Set objEventoHistorico = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoProduto = New AdmEvento
            
    Set objGrid1 = New AdmGrid
    Set objGridCusto = New AdmGrid
    
    sModulo = ""
    iTransacao = 0
    sModelo = ""
    
    'inicializa o grid de lancamentos padrão
    lErro = Inicializa_Grid_Lancamento(objGrid1)
    If lErro <> SUCESSO Then Error 36076

    'Inicializa o grid de custo
    lErro = Inicializa_Grid_Custo(objGridCusto)
    If lErro <> SUCESSO Then Error 33588

    'carrega a combobox de modulos
    lErro = Carga_Combobox_Modulo()
    If lErro <> SUCESSO Then Error 36100

    'carrega a combobox de funcoes
    lErro = Carga_Combobox_Funcoes()
    If lErro <> SUCESSO Then Error 36089
    
    'carrega a combobox de operadores
    lErro = Carga_Combobox_Operadores()
    If lErro <> SUCESSO Then Error 36099
    
    Modelo.Clear
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 33588, 36076, 36089, 36099, 36100
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157861)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157862)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Carga_Combobox_Funcoes() As Long
'carrega a combobox que contem as funcoes disponiveis

Dim colFormulaFuncao As New Collection
Dim objFormulaFuncao As ClassFormulaFuncao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Funcoes
        
    'leitura das funcoes no BD
    lErro = CF("FormulaFuncao_Le_Todos", colFormulaFuncao)
    If lErro <> SUCESSO Then Error 36078
    
    For Each objFormulaFuncao In colFormulaFuncao
        
        Funcoes.AddItem objFormulaFuncao.sFuncaoCombo
                
    Next
    
    Carga_Combobox_Funcoes = SUCESSO

    Exit Function

Erro_Carga_Combobox_Funcoes:

    Carga_Combobox_Funcoes = Err

    Select Case Err

        Case 36078
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157863)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Operadores() As Long
'carrega a combobox que contem os operadores disponiveis

Dim colFormulaOperador As New Collection
Dim objFormulaOperador As ClassFormulaOperador
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Operadores
        
    'leitura dos operadores no BD
    lErro = CF("FormulaOperador_Le_Todos", colFormulaOperador)
    If lErro <> SUCESSO Then Error 36090
    
    For Each objFormulaOperador In colFormulaOperador
        
        Operadores.AddItem objFormulaOperador.sOperadorCombo
                
    Next
    
    Carga_Combobox_Operadores = SUCESSO

    Exit Function

Erro_Carga_Combobox_Operadores:

    Carga_Combobox_Operadores = Err

    Select Case Err

        Case 36090
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157864)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Mnemonicos(sModulo As String, iTransacao As Integer) As Long
'carrega a combobox que contem os mnemonicos disponiveis para a transacao selecionada.

Dim colMnemonico As New Collection
Dim objMnemonico As ClassMnemonico
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Mnemonicos
        
    Mnemonicos.Enabled = True
    Mnemonicos.Clear
        
    'leitura dos mnemonicos no BD para o modulo/transacao em questão
    lErro = CF("MnemonicoCTB_Le", gcolModulo.Sigla(sModulo), iTransacao, colMnemonico)
    If lErro <> SUCESSO Then Error 36077
    
    For Each objMnemonico In colMnemonico
        
        Mnemonicos.AddItem objMnemonico.sMnemonicoCombo
                
    Next
    
    Carga_Combobox_Mnemonicos = SUCESSO

    Exit Function

Erro_Carga_Combobox_Mnemonicos:

    Carga_Combobox_Mnemonicos = Err

    Select Case Err

        Case 36077
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157865)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Modulo() As Long
'carrega a combobox com  os módulos disponiveis para o sistema

Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_Carga_Combobox_Modulo
        
    For iIndice = 1 To gcolModulo.Count
        If gcolModulo.Item(iIndice).iAtivo = MODULO_ATIVO Then
            If gcolModulo.Item(iIndice).sSigla <> MODULO_ADM And gcolModulo.Item(iIndice).sSigla <> MODULO_CONTABILIDADE And gcolModulo.Item(iIndice).sSigla <> MODULO_PCP Then
                Modulo.AddItem gcolModulo.Item(iIndice).sNome
            End If
        End If
    Next
    
    Carga_Combobox_Modulo = SUCESSO

    Exit Function

Erro_Carga_Combobox_Modulo:

    Carga_Combobox_Modulo = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157866)

    End Select
    
    Exit Function

End Function

Private Function Carga_Combobox_Transacao(sModulo As String) As Long
'carrega a combobox que contem as transacoes disponiveis para o modulo selecionado.

Dim colTransacao As New Collection
Dim objTransacao As ClassTransacaoCTB
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Transacao
        
    Transacao.Enabled = True
    Transacao.Clear
        
    'leitura das contas no BD
    lErro = CF("TransacaoCTB_Le_Todos", gcolModulo.Sigla(sModulo), colTransacao)
    If lErro <> SUCESSO Then Error 36107
    
    For Each objTransacao In colTransacao
        
        Transacao.AddItem objTransacao.sTransacaoTela
        Transacao.ItemData(Transacao.NewIndex) = objTransacao.iCodigo
                
    Next
    
    Modelo.Clear
    Modelo.Enabled = False
    sModelo = ""
    iTransacao = 0
    
    Call Grid_Limpa(objGrid1)
    GridLancamentos.Enabled = False
    
    Mnemonicos.Clear
    Mnemonicos.Enabled = False
    
    Carga_Combobox_Transacao = SUCESSO

    Exit Function

Erro_Carga_Combobox_Transacao:

    Carga_Combobox_Transacao = Err

    Select Case Err

        Case 36107
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157867)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Lancamento(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Lancamento
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("Centro de Custo/Lucro")
    objGridInt.colColuna.Add ("Débito")
    objGridInt.colColuna.Add ("Crédito")
    objGridInt.colColuna.Add ("Histórico")
    objGridInt.colColuna.Add ("Aglutina")
    If giContabGerencial = 1 Then objGridInt.colColuna.Add ("Status")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Debito.Name)
    objGridInt.colCampo.Add (Credito.Name)
    objGridInt.colCampo.Add (Historico.Name)
    objGridInt.colCampo.Add (Aglutina.Name)
    If giContabGerencial = 1 Then objGridInt.colCampo.Add (Gerencial.Name)
    
    'indica onde estão situadas as colunas do grid
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_Conta_Col = 1
        iGrid_Ccl_Col = 2
        iGrid_Debito_Col = 3
        iGrid_Credito_Col = 4
        iGrid_Historico_Col = 5
        iGrid_Aglutina_Col = 6
    Else
        iGrid_Conta_Col = 1
        '999 indica que não está sendo usado
        iGrid_Ccl_Col = 999
        iGrid_Debito_Col = 2
        iGrid_Credito_Col = 3
        iGrid_Historico_Col = 4
        iGrid_Aglutina_Col = 5
        Ccl.Visible = False
    End If
    
    If giContabGerencial = 1 Then
        iGrid_Gerencial_Col = iGrid_Aglutina_Col + 1
    Else
        Gerencial.Visible = False
    End If
    
    objGridInt.objGrid = GridLancamentos
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
        
    GridLancamentos.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Lancamento = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Lancamento:

    Inicializa_Grid_Lancamento = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157868)
        
    End Select

    Exit Function
        
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        If objGridInt.objGrid.Name = GridLancamentos.Name Then
    
            Select Case GridLancamentos.Col
        
                Case iGrid_Conta_Col
                
                    lErro = Saida_Celula_Conta(objGridInt)
                    If lErro <> SUCESSO Then gError 36125
                    
                Case iGrid_Ccl_Col
                
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 36126
                    
                Case iGrid_Credito_Col
                
                    lErro = Saida_Celula_Credito(objGridInt)
                    If lErro <> SUCESSO Then gError 36127
                    
                Case iGrid_Debito_Col
                
                    lErro = Saida_Celula_Debito(objGridInt)
                    If lErro <> SUCESSO Then gError 36128
    
                Case iGrid_Historico_Col
                
                    lErro = Saida_Celula_Historico(objGridInt)
                    If lErro <> SUCESSO Then gError 36129
    
                Case iGrid_Aglutina_Col
                
                    lErro = Saida_Celula_Aglutina(objGridInt)
                    If lErro <> SUCESSO Then gError 36400
    
                Case iGrid_Gerencial_Col
                
                    lErro = Saida_Celula_Gerencial(objGridInt)
                    If lErro <> SUCESSO Then gError 188060
    
    
            End Select
    
        ElseIf objGridInt.objGrid.Name = GridCusto.Name Then
        
            Select Case GridCusto.Col
        
                Case iGrid_ContaCusto_Col
                
                    lErro = Saida_Celula_ContaCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 33589
                    
                Case iGrid_CclCusto_Col
                
                    lErro = Saida_Celula_CclCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 33590
                    
                Case iGrid_CreditoCusto_Col
                
                    lErro = Saida_Celula_CreditoCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 33591
                    
                Case iGrid_DebitoCusto_Col
                
                    lErro = Saida_Celula_DebitoCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 33592
    
                Case iGrid_HistoricoCusto_Col
                
                    lErro = Saida_Celula_HistoricoCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 33593
    
                Case iGrid_ProdutoCusto_Col
                
                    lErro = Saida_Celula_ProdutoCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 36395
    
                Case iGrid_AglutinaCusto_Col
                
                    lErro = Saida_Celula_AglutinaCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 36401
    
                Case iGrid_GerencialCusto_Col
                
                    lErro = Saida_Celula_GerencialCusto(objGridInt)
                    If lErro <> SUCESSO Then gError 188061
    
            End Select
            
        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 36130
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 33589, 33590, 33591, 33592, 33593, 36125, 36126, 36127, 36128, 36129, 36395, 36400, 36401, 188060, 188061
    
        Case 36130
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157869)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta
    
    If Len(Trim(Conta.Text)) > 0 Then

        If Checkbox_Verifica_Sintaxe.Value = 1 Then

            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, Conta.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36145
            
        End If
                
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36146

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = Err
    
    Select Case Err
    
        Case 36145
            Conta.SelStart = iInicio
            Conta.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 36146
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157870)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula ccl do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl
                
    If Len(Trim(Ccl.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, Ccl.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36147
            
        End If
                
'        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
'            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
'        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36148

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err
    
    Select Case Err
    
        Case 36147
            Ccl.SelStart = iInicio
            Ccl.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 36148
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157871)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Credito(objGridInt As AdmGrid) As Long
'faz a critica da celula credito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_Credito

    Set objGridInt.objControle = Credito

    If Len(Trim(Credito.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, Credito.Text, TIPO_NUMERICO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36149
                
        End If
        
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36150
                
    Saida_Celula_Credito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Credito:

    Saida_Celula_Credito = Err
    
    Select Case Err
    
        Case 36149
            Credito.SelStart = iInicio
            Credito.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 36150
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157872)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Debito(objGridInt As AdmGrid) As Long
'faz a critica da celula debito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_Debito

    Set objGridInt.objControle = Debito

    If Len(Trim(Debito.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, Debito.Text, TIPO_NUMERICO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36151
                
        End If
        
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36152

    Saida_Celula_Debito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Debito:

    Saida_Celula_Debito = Err
    
    Select Case Err
    
        Case 36151
            Debito.SelStart = iInicio
            Debito.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 36152
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157873)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'faz a critica da celula historico do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sValor As String
Dim objHistPadrao As New ClassHistPadrao


On Error GoTo Erro_Saida_Celula_Historico

    Set objGridInt.objControle = Historico

    If Len(Trim(Historico.Text)) > 0 Then
    
        If Left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
        
            sValor = Trim(Mid(Historico.Text, 2))
            
            lErro = Valor_Inteiro_Critica(sValor)
            If lErro <> SUCESSO Then Error 44087
            
            objHistPadrao.iHistPadrao = CInt(sValor)
                    
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then Error 44088
            
            If lErro = 5446 Then Error 44089
    
            Historico.Text = Chr(34) + objHistPadrao.sDescHistPadrao + Chr(34)
            Historico.SelStart = 0
            
        End If
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, Historico.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36153
                
        End If
        
        
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36154

    Saida_Celula_Historico = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = Err
    
    Select Case Err
    
    
        Case 36153
            Historico.SelStart = iInicio
            Historico.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 36154, 44088
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 44087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 44089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157874)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Aglutina(objGridInt As AdmGrid) As Long
'faz a critica da celula aglutina do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_Aglutina

    Set objGridInt.objControle = Aglutina

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36402

    Saida_Celula_Aglutina = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Aglutina:

    Saida_Celula_Aglutina = Err
    
    Select Case Err
    
        Case 36402
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157875)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Gerencial(objGridInt As AdmGrid) As Long
'faz a critica da celula Gerencial do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Gerencial

    Set objGridInt.objControle = Gerencial

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188062

    Saida_Celula_Gerencial = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Gerencial:

    Saida_Celula_Gerencial = gErr
    
    Select Case gErr
    
        Case 188062
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188063)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_GerencialCusto(objGridInt As AdmGrid) As Long
'faz a critica da celula GerencialCusto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_GerencialCusto

    Set objGridInt.objControle = GerencialCusto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188064

    Saida_Celula_GerencialCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GerencialCusto:

    Saida_Celula_GerencialCusto = gErr
    
    Select Case gErr
    
        Case 188064
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188065)
        
    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    sModelo = Modelo.Text

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 36166
    
    Call Limpa_Tela_PadraoContab

    'carrega a combobox de modelos com os modelos referentes ao modulo/transacao em questão
    lErro = Carga_Combobox_Modelo(sModulo, iTransacao, iSubTipo)
    If lErro <> SUCESSO Then Error 36205

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 36166, 36205
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157876)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'grava os dados da tela

Dim lErro As Long
Dim objPadraoContab As New ClassPadraoContab
Dim colPadraoContabItem As New ClassColPadraoContabItem
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Modelo está preenchido. Se estiver é sinal que o modulo e a transacao tambem estao
    If Len(Modelo.Text) = 0 Then Error 36155
  
    objPadraoContab.sModulo = gcolModulo.Sigla(sModulo)
    objPadraoContab.iTransacao = iTransacao
    objPadraoContab.iSubTipo = iSubTipo
    objPadraoContab.sModelo = sModelo
    
    If Padrao.Value = vbChecked Then
        objPadraoContab.iPadrao = PADRAOCONTAB_MODELO_PADRAO
    Else
        objPadraoContab.iPadrao = PADRAOCONTAB_NAO_E_MODELO_PADRAO
    End If
    
    If Obrigatoriedade.Value = vbChecked Then
        objPadraoContab.iObrigatoriedade = CONTABILIZACAO_OBRIGATORIA
    Else
        objPadraoContab.iObrigatoriedade = CONTABILIZACAO_NAO_OBRIGATORIA
    End If
    
    'Move os dados do grid Lançamentos para a coleção colPadraoContabItem
    lErro = Grid_Lancamento(colPadraoContabItem)
    If lErro <> SUCESSO Then Error 36164
    
    'Move os dados do grid Custo para a coleção colPadraoContabItem
    lErro = Grid_Custo(colPadraoContabItem)
    If lErro <> SUCESSO Then Error 33600
    
'    If colPadraoContabItem.Count = 0 Then Error 36156
    
'    lErro = Trata_Alteracao(objPadraoContab, objPadraoContab.sModulo)
'    If lErro <> SUCESSO Then Error 32304

    
    'Grava o modelo padrão de contabilização em questão
    lErro = CF("PadraoContab_Grava", objPadraoContab, colPadraoContabItem)
    If lErro <> SUCESSO Then Error 36165
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 36155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", Err)
        
        Case 36156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_LANCAMENTOS_PADRAOCONTAB", Err)
        
        Case 33600, 36164, 36165, 32304
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157877)
            
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoExcluir_Click()
    
Dim lErro As Long
Dim objPadraoContab As New ClassPadraoContab
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Modelo está preenchido. Se estiver é sinal que o modulo e a transacao tambem estao
    If Len(Modelo.Text) = 0 Then Error 36185
    
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_PADRAOCONTAB")
    
    If vbMsgRes = vbYes Then
    
        objPadraoContab.sModulo = gcolModulo.Sigla(Modulo.Text)
        objPadraoContab.iTransacao = iTransacao
        objPadraoContab.iSubTipo = iSubTipo
        objPadraoContab.sModelo = Modelo.Text
    
        'exclui o modelo padrão de contabilização em questão
        lErro = CF("PadraoContab_Exclui", objPadraoContab)
        If lErro <> SUCESSO Then Error 36186
    
        Call Limpa_Tela_PadraoContab
        
        'carrega a combobox de modelos com os modelos referentes ao modulo/transacao em questão
        lErro = Carga_Combobox_Modelo(sModulo, iTransacao, iSubTipo)
        If lErro <> SUCESSO Then Error 36206
        
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 36185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", Err)
        
        Case 36186, 36206
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157878)
        
    End Select

    Exit Sub
    
End Sub

Function Limpa_Tela_PadraoContab() As Long

    Call Grid_Limpa(objGridCusto)
    Call Grid_Limpa(objGrid1)
    Modelo.Text = ""
    Padrao.Value = vbUnchecked
    sModelo = ""
    
    Limpa_Tela_PadraoContab = SUCESSO
    
End Function

Function Grid_Lancamento(colPadraoContabItem As ClassColPadraoContabItem) As Long
'move os dados do grid para a colecao colPadraoContabItem

Dim iIndice1 As Integer
Dim objPadraoContabItem As ClassPadraoContabItem
Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Grid_Lancamento

    For iIndice1 = 1 To objGrid1.ilinhasExistentes
        
        Set objPadraoContabItem = New ClassPadraoContabItem
        
        objPadraoContabItem.sModulo = gcolModulo.Sigla(sModulo)
        objPadraoContabItem.iTransacao = iTransacao
        objPadraoContabItem.iSubTipo = iSubTipo
        objPadraoContabItem.sModelo = sModelo
        objPadraoContabItem.iCusto = PADRAOCONTABITEM_NORMAL
        
        objPadraoContabItem.iItem = iIndice1
            
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridLancamentos.TextMatrix(iIndice1, iGrid_Conta_Col), TIPO_TEXTO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 36217
            
        objPadraoContabItem.sConta = GridLancamentos.TextMatrix(iIndice1, iGrid_Conta_Col)
        
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_Ccl_Col <> 999 Then
        
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridLancamentos.TextMatrix(iIndice1, iGrid_Ccl_Col), TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36218
        
            objPadraoContabItem.sCcl = GridLancamentos.TextMatrix(iIndice1, iGrid_Ccl_Col)
            
        End If
        
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col), TIPO_NUMERICO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 36219
    
        objPadraoContabItem.sCredito = GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col)
        
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col), TIPO_NUMERICO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 36220
            
        objPadraoContabItem.sDebito = GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col)
        
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridLancamentos.TextMatrix(iIndice1, iGrid_Historico_Col), TIPO_TEXTO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 36221
        
        objPadraoContabItem.sHistorico = GridLancamentos.TextMatrix(iIndice1, iGrid_Historico_Col)
    
        If GridLancamentos.TextMatrix(iIndice1, iGrid_Aglutina_Col) = "" Or GridLancamentos.TextMatrix(iIndice1, iGrid_Aglutina_Col) = "0" Then
            objPadraoContabItem.iAglutina = LANCAMENTO_NAO_AGLUTINADO
        Else
            objPadraoContabItem.iAglutina = LANCAMENTO_AGLUTINA
        End If
    
        If giContabGerencial = 1 Then
            If GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col) = "" Or GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col) = "0" Then
                objPadraoContabItem.iGerencial = LANCAMENTO_STATUS_NAO_GERENCIAL
            ElseIf GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col) = "1" Then
                objPadraoContabItem.iGerencial = LANCAMENTO_STATUS_GERENCIAL
            ElseIf GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col) = "2" Then
                objPadraoContabItem.iGerencial = LANCAMENTO_STATUS_AMBOS
            End If
        End If
        
        'Armazena o objeto objPadraoContabItem na coleção colPadraoContabItem
        colPadraoContabItem.Add objPadraoContabItem
        
    Next
    
    Grid_Lancamento = SUCESSO

    Exit Function

Erro_Grid_Lancamento:

    Grid_Lancamento = Err

    Select Case Err
    
        Case 36217
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Conta_Col
            TabStrip1.Tabs.Item(TAB_LANCAMENTOS).Selected = True
            GridLancamentos.SetFocus
            
        Case 36218
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Ccl_Col
            TabStrip1.Tabs.Item(TAB_LANCAMENTOS).Selected = True
            GridLancamentos.SetFocus
            
        Case 36219
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Credito_Col
            TabStrip1.Tabs.Item(TAB_LANCAMENTOS).Selected = True
            GridLancamentos.SetFocus
            
        Case 36220
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Debito_Col
            TabStrip1.Tabs.Item(TAB_LANCAMENTOS).Selected = True
            GridLancamentos.SetFocus
            
        Case 36221
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Historico_Col
            TabStrip1.Tabs.Item(TAB_LANCAMENTOS).Selected = True
            GridLancamentos.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157879)
            
    End Select
    
    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lDoc As Long
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 36167

    Call Limpa_Tela_PadraoContab
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 36167
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157880)
        
    End Select
    
End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub HistoricoCusto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub HistoricoCusto_GotFocus()

Dim iPos As Integer
    
    Call Grid_Campo_Recebe_Foco(objGridCusto)
    
    If Len(Historico.Text) > 0 Then
        iPos = InStr(HistoricoCusto.Text, CARACTER_HISTORICO_PARAM)
        If iPos > 0 Then
            HistoricoCusto.SelStart = iPos - 1
            HistoricoCusto.SelLength = 1
        End If
    End If


End Sub

Public Sub HistoricoCusto_KeyPress(KeyAscii As Integer)

Dim iInicio As Integer
Dim iPos As Integer
Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_HistoricoCusto_KeyPress

    'se digitou ENTER
    If KeyAscii = vbKeyReturn Then
        
        If Len(HistoricoCusto.Text) > 0 Then
        
            If Left(HistoricoCusto.Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(HistoricoCusto.Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44084
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44085
                
                If lErro = 5446 Then Error 44086
        
                HistoricoCusto.Text = Chr(34) + objHistPadrao.sDescHistPadrao + Chr(34)
                HistoricoCusto.SelStart = 0
                
            End If
    
            If HistoricoCusto.SelText = CARACTER_HISTORICO_PARAM Then
                iInicio = HistoricoCusto.SelStart + 2
            Else
                iInicio = HistoricoCusto.SelStart
            End If
        
            If iInicio = 0 Then iInicio = 1
        
            iPos = InStr(iInicio, HistoricoCusto.Text, CARACTER_HISTORICO_PARAM)
            If iPos > 0 Then
                HistoricoCusto.SelStart = iPos - 1
                HistoricoCusto.SelLength = 1
                Exit Sub
            End If
        End If
    End If

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCusto)

    Exit Sub
    
Erro_HistoricoCusto_KeyPress:

    Select Case Err
    
        Case 44084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
        
        Case 44085

        Case 44086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157881)
    
    End Select

    Exit Sub
    
End Sub

Public Sub HistoricoCusto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCusto.objControle = HistoricoCusto
    lErro = Grid_Campo_Libera_Foco(objGridCusto)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Mnemonicos_Click()

Dim iPos As Integer
Dim lErro As Long
Dim lPos As Long
Dim objMnemonico As New ClassMnemonico
Dim sMnemoncico As String
Dim sMnemonico As String

On Error GoTo Erro_Mnemonicos_Click
    
    If Len(Mnemonicos.Text) > 0 Then
    
        objMnemonico.sModulo = gcolModulo.Sigla(sModulo)
        objMnemonico.iTransacao = iTransacao
        objMnemonico.sMnemonicoCombo = Mnemonicos.Text
    
        'retorna os dados do mnemonico passado como parametro
        lErro = CF("MnemonicoCTB_Le_Mnemonico", objMnemonico)
        If lErro <> SUCESSO And lErro <> 36082 Then Error 36207
        
        If lErro = 36082 Then Error 36478
        
        Descricao.Text = objMnemonico.sMnemonicoDesc
        
        lPos = InStr(1, Mnemonicos.Text, "(")
        If lPos = 0 Then
            sMnemonico = Mnemonicos.Text
        Else
            sMnemonico = Mid(Mnemonicos.Text, 1, lPos)
        End If
        
        lErro = Mnemonicos1(sMnemonico)
        If lErro <> SUCESSO Then Error 33605
        
    End If
    
    Exit Sub
    
Erro_Mnemonicos_Click:

    Select Case Err
    
        Case 33605, 36207
    
        Case 36478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonicoCombo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157882)
            
    End Select
        
    Exit Sub
        
End Sub

Public Sub Modelo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Modelo_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Modelo_Click

    If Modelo.ListIndex = -1 Then Exit Sub

    If Modelo.Text = sModelo Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 36122
    
    'pega o valor do novo Modelo
    sModelo = Modelo.Text
    
    'se é o modelo padrão ==> ligar a checkbox
    If Modelo.ItemData(Modelo.ListIndex) = PADRAOCONTAB_MODELO_PADRAO Then
        Padrao.Value = vbChecked
    Else
        Padrao.Value = vbUnchecked
    End If
    
    'carrega o grid com os dados do modelo em questão
    lErro = Carga_Grid(sModulo, iTransacao, sModelo)
    If lErro <> SUCESSO Then Error 36123

    GridLancamentos.TopRow = 1
    GridCusto.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Modelo_Click:

    Select Case Err

        Case 36122
            For iIndice = 0 To Modelo.ListCount - 1
                If Modelo.List(iIndice) = sModelo Then
                    Modelo.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 36123
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157883)
            
    End Select
        
    Exit Sub

End Sub

Public Sub Modulo_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Modulo_Click

    If Modulo.ListIndex = -1 Then Exit Sub

    If Modulo.Text = sModulo Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 36101
    
    'pega o valor do novo Modulo
    sModulo = Modulo.Text

    'carrega a combobox de transações com as transações referentes ao modulo em questão
    lErro = Carga_Combobox_Transacao(sModulo)
    If lErro <> SUCESSO Then Error 36102

    GridLancamentos.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Modulo_Click:

    Select Case Err

        Case 36101
            For iIndice = 0 To Modulo.ListCount - 1
                If Modulo.List(iIndice) = sModulo Then
                    Modulo.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 36102
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157884)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1
    
    lErro = Traz_Ccl_Tela(objCcl)
    If lErro <> SUCESSO Then Error 33602
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case Err
    
        Case 33602
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157885)
        
    End Select

    Exit Sub

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta

On Error GoTo Erro_objEventoConta_evSelecao

    Set objPlanoConta = obj1
    
    lErro = Traz_Conta_Tela(objPlanoConta)
    If lErro <> SUCESSO Then Error 33601
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case Err
    
        Case 33601
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157886)
        
    End Select

    Exit Sub

End Sub

Private Sub objEventoHistorico_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
    
On Error GoTo Erro_objEventoHistorico_evSelecao
    
    Set objHistPadrao = obj1
    
    lErro = Traz_Historico_Tela(objHistPadrao)
    If lErro <> SUCESSO Then Error 33603
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoHistorico_evSelecao:

    Select Case Err
    
        Case 33603
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157887)
            
        End Select
        
    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then Error 36396

    Me.Show
    
    Exit Sub
    
Erro_objEventoProduto_evSelecao:

    Select Case Err
    
        Case 36396
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157888)
            
        End Select
        
    Exit Sub

End Sub

Public Sub Operadores_Click()

Dim iPos As Integer
Dim lErro As Long
Dim objFormulaOperador As New ClassFormulaOperador
Dim lPos As Integer

On Error GoTo Erro_Operadores_Click
    
    objFormulaOperador.sOperadorCombo = Operadores.Text
    
    'retorna os dados do operador passado como parametro
    lErro = CF("FormulaOperador_Le", objFormulaOperador)
    If lErro <> SUCESSO And lErro <> 36098 Then Error 36215
    
    Descricao.Text = objFormulaOperador.sOperadorDesc
    
    Call Operadores1
    
    Exit Sub
    
Erro_Operadores_Click:

    Select Case Err
    
        Case 36215
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157889)
            
    End Select
        
    Exit Sub

End Sub

Public Sub Padrao_Click()

Dim iIndice As Integer

    iAlterado = REGISTRO_ALTERADO
    
    If Padrao.Value = vbChecked Then
        For iIndice = 0 To Modelo.ListCount - 1
            If iIndice <> Modelo.ListIndex Then
                If Modelo.ItemData(iIndice) = PADRAOCONTAB_MODELO_PADRAO Then
                    Modelo.ItemData(iIndice) = PADRAOCONTAB_NAO_E_MODELO_PADRAO
                    Exit For
                End If
            End If
        Next
        If Modelo.ListIndex >= 0 Then Modelo.ItemData(Modelo.ListIndex) = PADRAOCONTAB_MODELO_PADRAO
    End If
    
End Sub

Public Sub Obrigatoriedade_Click()
'Verifica a Obrigatoriedade da trasacao

Dim iIndice As Integer

    iAlterado = REGISTRO_ALTERADO
    
    If Obrigatoriedade.Value = vbChecked Then
    
        For iIndice = 0 To Transacao.ListCount - 1
                    
            If iIndice <> Transacao.ListIndex Then
            
                'Verifica se a Contabilidade e Obrigatoria
                If Transacao.ItemData(iIndice) = CONTABILIZACAO_OBRIGATORIA Then
                    
                    Transacao.ItemData(iIndice) = CONTABILIZACAO_NAO_OBRIGATORIA
                    
                    Exit For
                
                End If
                
            End If
            
        Next
        
        If Transacao.ListIndex >= 0 Then Transacao.ItemData(Transacao.ListIndex) = CONTABILIZACAO_OBRIGATORIA
        
    End If
    
End Sub

Public Sub TabStrip1_Click()
    
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If
    
End Sub

Public Sub Transacao_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objTransacaoCTB As New ClassTransacaoCTB

On Error GoTo Erro_Transacao_Click

    If Transacao.ListIndex = -1 Then Exit Sub

    'verifica se existe a necessidade de salvar o modelo antigo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 36114
    
    objTransacaoCTB.sSigla = gcolModulo.Sigla(sModulo)
    objTransacaoCTB.sTransacaoTela = Transacao.Text
    
    lErro = CF("TransacaoCTB_Le_TransacaoTela", objTransacaoCTB)
    If lErro <> SUCESSO Then gError 71694
    
    'verifica se a contabilidade e obrigatoria
    If objTransacaoCTB.iObrigatoriedade = CONTABILIZACAO_OBRIGATORIA Then
        Obrigatoriedade = vbChecked
    Else
        Obrigatoriedade = vbUnchecked
    End If
        
    'pega o valor do novo Modulo
    iTransacao = objTransacaoCTB.iCodigo
    iSubTipo = objTransacaoCTB.iSubTipo

    'carrega a combobox de modelos com os modelos referentes ao modulo/transacao em questão
    lErro = Carga_Combobox_Modelo(sModulo, iTransacao, iSubTipo)
    If lErro <> SUCESSO Then gError 36115

    'carrega a combobox de mnemonicos com os mnemonicos referentes ao modulo/transacao em questão
    lErro = Carga_Combobox_Mnemonicos(sModulo, iTransacao)
    If lErro <> SUCESSO Then gError 36117

    GridLancamentos.TopRow = 1
    
    iAlterado = 0

    Exit Sub

Erro_Transacao_Click:

    Select Case gErr

        Case 36114
            For iIndice = 0 To Transacao.ListCount - 1
                If Transacao.ItemData(iIndice) = iTransacao Then
                    Transacao.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case 36115, 36117
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157890)
            
    End Select
        
    Exit Sub

End Sub

Private Function Carga_Combobox_Modelo(sModulo As String, iTransacao As Integer, iSubTipo As Integer) As Long
'carrega a combobox que contem os modelos disponiveis para o modulo/transacao selecionado.

Dim colPadraoContab As New Collection
Dim objPadraoContab As ClassPadraoContab
Dim lErro As Long
    
On Error GoTo Erro_Carga_Combobox_Modelo
        
    Modelo.Enabled = True
    Modelo.Clear
        
    'leitura dos modelos de contabilizacao referentes ao modulo/transacao em questao
    lErro = CF("PadraoContab_Le_Todos", gcolModulo.Sigla(sModulo), iTransacao, colPadraoContab, iSubTipo)
    If lErro <> SUCESSO Then Error 36116
    
    For Each objPadraoContab In colPadraoContab
        Modelo.AddItem objPadraoContab.sModelo
        Modelo.ItemData(Modelo.NewIndex) = objPadraoContab.iPadrao
    Next
    
    sModelo = ""
    
    Call Grid_Limpa(objGrid1)
    GridLancamentos.Enabled = True
    
    Call Grid_Limpa(objGridCusto)
    GridCusto.Enabled = True
    
    Padrao.Value = vbUnchecked
    
    Carga_Combobox_Modelo = SUCESSO

    Exit Function

Erro_Carga_Combobox_Modelo:

    Carga_Combobox_Modelo = Err

    Select Case Err

        Case 36116
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157891)

    End Select
    
    Exit Function

End Function
   
Function Carga_Grid(sModulo As String, iTransacao As Integer, sModelo As String) As Long
'carrega o grid com os dados do modelo em questão

Dim lErro As Long
Dim colPadraoContabItem As New ClassColPadraoContabItem
Dim objPadraoContabItem As ClassPadraoContabItem
Dim objPadraoContab As New ClassPadraoContab
    
On Error GoTo Erro_Carga_Grid
        
    Call Grid_Limpa(objGrid1)
    GridLancamentos.Enabled = True
    
    Call Grid_Limpa(objGridCusto)
    GridCusto.Enabled = True
        
    objPadraoContab.sModulo = gcolModulo.Sigla(sModulo)
    objPadraoContab.iTransacao = iTransacao
    objPadraoContab.iSubTipo = iSubTipo
    objPadraoContab.sModelo = sModelo
        
    'leitura dos lançamentos padrão do modelo em questão
    lErro = CF("PadraoContabItem_Le_Modelo", objPadraoContab, colPadraoContabItem)
    If lErro <> SUCESSO Then Error 36124
    
    For Each objPadraoContabItem In colPadraoContabItem
                
        If objPadraoContabItem.iCusto = PADRAOCONTABITEM_NORMAL Then
        
            'coloca os dados na tela
            GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Conta_Col) = objPadraoContabItem.sConta
            If iGrid_Ccl_Col <> 999 Then GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Ccl_Col) = objPadraoContabItem.sCcl
            GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Credito_Col) = objPadraoContabItem.sCredito
            GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Debito_Col) = objPadraoContabItem.sDebito
            GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Historico_Col) = objPadraoContabItem.sHistorico
            GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Aglutina_Col) = CStr(objPadraoContabItem.iAglutina)
            If giContabGerencial = 1 Then GridLancamentos.TextMatrix(objPadraoContabItem.iItem, iGrid_Gerencial_Col) = CStr(objPadraoContabItem.iGerencial)
        
            objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
            
        ElseIf objPadraoContabItem.iCusto = PADRAOCONTABITEM_CUSTO Then
                    
            'coloca os dados na tela
            GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_ContaCusto_Col) = objPadraoContabItem.sConta
            If iGrid_CclCusto_Col <> 999 Then GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_CclCusto_Col) = objPadraoContabItem.sCcl
            GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_CreditoCusto_Col) = objPadraoContabItem.sCredito
            GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_DebitoCusto_Col) = objPadraoContabItem.sDebito
            GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_HistoricoCusto_Col) = objPadraoContabItem.sHistorico
            GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_ProdutoCusto_Col) = objPadraoContabItem.sProduto
            GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_AglutinaCusto_Col) = CStr(objPadraoContabItem.iAglutina)
            If giContabGerencial = 1 Then GridCusto.TextMatrix(objPadraoContabItem.iItem, iGrid_GerencialCusto_Col) = CStr(objPadraoContabItem.iGerencial)
        
            objGridCusto.ilinhasExistentes = objGridCusto.ilinhasExistentes + 1
            
        End If
                
    Next
    
    Call Grid_Refresh_Checkbox(objGrid1)
    Call Grid_Refresh_Checkbox(objGridCusto)
    
    
    Carga_Grid = SUCESSO

    Exit Function

Erro_Carga_Grid:

    Carga_Grid = Err

    Select Case Err

        Case 36124
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157892)

    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Custo(objGridInt As AdmGrid) As Long
'Inicializa o grid de custo

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Custo
    
    'tela em questão
    Set objGridCusto.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("Centro de Custo/Lucro")
    objGridInt.colColuna.Add ("Débito (Quant.)")
    objGridInt.colColuna.Add ("Crédito (Quant.)")
    objGridInt.colColuna.Add ("Histórico")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Aglutina")
    If giContabGerencial = 1 Then objGridInt.colColuna.Add ("Status")
    
    'Campos de edição do grid
    objGridInt.colCampo.Add (ContaCusto.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (CclCusto.Name)
    objGridInt.colCampo.Add (DebitoCusto.Name)
    objGridInt.colCampo.Add (CreditoCusto.Name)
    objGridInt.colCampo.Add (HistoricoCusto.Name)
    objGridInt.colCampo.Add (ProdutoCusto.Name)
    objGridInt.colCampo.Add (AglutinaCusto.Name)
    If giContabGerencial = 1 Then objGridInt.colCampo.Add (GerencialCusto.Name)
    
    'Indica onde estão situadas as colunas do grid
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_ContaCusto_Col = 1
        iGrid_CclCusto_Col = 2
        iGrid_DebitoCusto_Col = 3
        iGrid_CreditoCusto_Col = 4
        iGrid_HistoricoCusto_Col = 5
        iGrid_ProdutoCusto_Col = 6
        iGrid_AglutinaCusto_Col = 7
    Else
        iGrid_ContaCusto_Col = 1
        '999 indica que não está sendo usado
        iGrid_CclCusto_Col = 999
        iGrid_DebitoCusto_Col = 2
        iGrid_CreditoCusto_Col = 3
        iGrid_HistoricoCusto_Col = 4
        iGrid_ProdutoCusto_Col = 5
        iGrid_AglutinaCusto_Col = 6
        CclCusto.Visible = False
    End If
    
    If giContabGerencial = 1 Then
        iGrid_GerencialCusto_Col = iGrid_AglutinaCusto_Col + 1
    Else
        GerencialCusto.Visible = False
    End If
    
    objGridInt.objGrid = GridCusto
    
    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21
    
    'Linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
        
    GridCusto.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Custo = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Custo:

    Inicializa_Grid_Custo = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157893)
        
    End Select

    Exit Function
        
End Function

Private Function Saida_Celula_ContaCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula conta custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_ContaCusto

    Set objGridInt.objControle = ContaCusto
    
    If Len(Trim(ContaCusto.Text)) > 0 Then

        If Checkbox_Verifica_Sintaxe.Value = 1 Then

            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, ContaCusto.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 33578
                
        End If
        
        If GridCusto.Row - GridCusto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33579

    Saida_Celula_ContaCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_ContaCusto:

    Saida_Celula_ContaCusto = Err
    
    Select Case Err
    
        Case 33578
            ContaCusto.SelStart = iInicio
            ContaCusto.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 33579
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157894)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_CclCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula ccl custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_CclCusto

    Set objGridInt.objControle = CclCusto
                
    If Len(Trim(CclCusto.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
        
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, CclCusto.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 33580
                
        End If
        
        If GridCusto.Row - GridCusto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33581

    Saida_Celula_CclCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_CclCusto:

    Saida_Celula_CclCusto = Err
    
    Select Case Err
    
        Case 33580
            CclCusto.SelStart = iInicio
            CclCusto.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 33581
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157895)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_CreditoCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula credito custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_CreditoCusto

    Set objGridInt.objControle = CreditoCusto

    If Len(Trim(CreditoCusto.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, CreditoCusto.Text, TIPO_NUMERICO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 33582
                
        End If
        
        If GridCusto.Row - GridCusto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33583
                
    Saida_Celula_CreditoCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_CreditoCusto:

    Saida_Celula_CreditoCusto = Err
    
    Select Case Err
    
        Case 33582
            CreditoCusto.SelStart = iInicio
            CreditoCusto.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 33583
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157896)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_DebitoCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula debito custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_DebitoCusto

    Set objGridInt.objControle = DebitoCusto

    If Len(Trim(DebitoCusto.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, DebitoCusto.Text, TIPO_NUMERICO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 33584
                
        End If
        
        If GridCusto.Row - GridCusto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33585

    Saida_Celula_DebitoCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DebitoCusto:

    Saida_Celula_DebitoCusto = Err
    
    Select Case Err
    
        Case 33584
            DebitoCusto.SelStart = iInicio
            DebitoCusto.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 33585
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157897)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_HistoricoCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula historico custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim sValor As String
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Saida_Celula_HistoricoCusto

    Set objGridInt.objControle = HistoricoCusto

    If Len(Trim(HistoricoCusto.Text)) > 0 Then
    
        If Left(HistoricoCusto.Text, 1) = CARACTER_HISTPADRAO Then
        
            sValor = Trim(Mid(HistoricoCusto.Text, 2))
            
            lErro = Valor_Inteiro_Critica(sValor)
            If lErro <> SUCESSO Then Error 44090
            
            objHistPadrao.iHistPadrao = CInt(sValor)
                    
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then Error 44091
            
            If lErro = 5446 Then Error 44092
    
            HistoricoCusto.Text = Chr(34) + objHistPadrao.sDescHistPadrao + Chr(34)
            HistoricoCusto.SelStart = 0
            
        End If
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, HistoricoCusto.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 33586
                
        End If
        
        
        If GridCusto.Row - GridCusto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 33587

    Saida_Celula_HistoricoCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_HistoricoCusto:

    Saida_Celula_HistoricoCusto = Err
    
    Select Case Err
    
        Case 33586
            HistoricoCusto.SelStart = iInicio
            HistoricoCusto.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 33587, 44091
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 44090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 44092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157898)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula produto custo do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_ProdutoCusto

    Set objGridInt.objControle = ProdutoCusto

    If Len(Trim(ProdutoCusto.Text)) > 0 Then
    
        If Checkbox_Verifica_Sintaxe.Value = 1 Then
    
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, ProdutoCusto.Text, TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 36397
                
        End If
        
        If GridCusto.Row - GridCusto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36398

    Saida_Celula_ProdutoCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_ProdutoCusto:

    Saida_Celula_ProdutoCusto = Err
    
    Select Case Err
    
        Case 36397
            ProdutoCusto.SelStart = iInicio
            ProdutoCusto.SelLength = iTamanho
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 36398
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157899)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_AglutinaCusto(objGridInt As AdmGrid) As Long
'Faz a critica da celula aglutinacusto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer

On Error GoTo Erro_Saida_Celula_AglutinaCusto

    Set objGridInt.objControle = AglutinaCusto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 36403

    Saida_Celula_AglutinaCusto = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_AglutinaCusto:

    Saida_Celula_AglutinaCusto = Err
    
    Select Case Err
    
        Case 36403
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 157900)
        
    End Select

    Exit Function

End Function

Function Grid_Custo(colPadraoContabItem As ClassColPadraoContabItem) As Long
'Move os dados do grid para a colecao colPadraoContabItem

Dim lErro As Long
Dim iInicio As Integer
Dim iTamanho As Integer
Dim iIndice1 As Integer
Dim objPadraoContabItem As ClassPadraoContabItem

On Error GoTo Erro_Grid_Custo

    For iIndice1 = 1 To objGridCusto.ilinhasExistentes
        
        Set objPadraoContabItem = New ClassPadraoContabItem
        
        objPadraoContabItem.sModulo = gcolModulo.Sigla(sModulo)
        objPadraoContabItem.iTransacao = iTransacao
        objPadraoContabItem.iSubTipo = iSubTipo
        objPadraoContabItem.sModelo = sModelo
        objPadraoContabItem.iCusto = PADRAOCONTABITEM_CUSTO
        
        objPadraoContabItem.iItem = iIndice1
            
        'se a conta não estiver preenchida ==> erro
'        If Len(GridCusto.TextMatrix(iIndice1, iGrid_ContaCusto_Col)) = 0 Then Error 36442
            
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridCusto.TextMatrix(iIndice1, iGrid_ContaCusto_Col), TIPO_TEXTO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 33595
            
        objPadraoContabItem.sConta = GridCusto.TextMatrix(iIndice1, iGrid_ContaCusto_Col)
        
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_CclCusto_Col <> 999 Then
        
            lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridCusto.TextMatrix(iIndice1, iGrid_CclCusto_Col), TIPO_TEXTO, iInicio, iTamanho)
            If lErro <> SUCESSO Then Error 33595
        
            objPadraoContabItem.sCcl = GridCusto.TextMatrix(iIndice1, iGrid_CclCusto_Col)
            
        End If
            
        'se as formulas de debito e credito não estiverem preenchidas ==> erro
'        If Len(GridCusto.TextMatrix(iIndice1, iGrid_CreditoCusto_Col)) = 0 And Len(GridCusto.TextMatrix(iIndice1, iGrid_DebitoCusto_Col)) = 0 Then Error 36443
        
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridCusto.TextMatrix(iIndice1, iGrid_CreditoCusto_Col), TIPO_NUMERICO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 33596
    
        objPadraoContabItem.sCredito = GridCusto.TextMatrix(iIndice1, iGrid_CreditoCusto_Col)
        
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridCusto.TextMatrix(iIndice1, iGrid_DebitoCusto_Col), TIPO_NUMERICO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 33597
            
        objPadraoContabItem.sDebito = GridCusto.TextMatrix(iIndice1, iGrid_DebitoCusto_Col)
        
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridCusto.TextMatrix(iIndice1, iGrid_HistoricoCusto_Col), TIPO_TEXTO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 33598
        
        objPadraoContabItem.sHistorico = GridCusto.TextMatrix(iIndice1, iGrid_HistoricoCusto_Col)
    
        'se o produto não estiver preenchido
'        If Len(GridCusto.TextMatrix(iIndice1, iGrid_ProdutoCusto_Col)) = 0 Then Error 36444
    
        lErro = CF("Valida_Formula", gcolModulo.Sigla(sModulo), iTransacao, GridCusto.TextMatrix(iIndice1, iGrid_ProdutoCusto_Col), TIPO_TEXTO, iInicio, iTamanho)
        If lErro <> SUCESSO Then Error 36399
            
        objPadraoContabItem.sProduto = GridCusto.TextMatrix(iIndice1, iGrid_ProdutoCusto_Col)
    
        If GridCusto.TextMatrix(iIndice1, iGrid_AglutinaCusto_Col) = "" Or GridCusto.TextMatrix(iIndice1, iGrid_AglutinaCusto_Col) = "0" Then
            objPadraoContabItem.iAglutina = LANCAMENTO_NAO_AGLUTINADO
        Else
            objPadraoContabItem.iAglutina = LANCAMENTO_AGLUTINA
        End If
    
        If giContabGerencial = 1 Then
            If GridCusto.TextMatrix(iIndice1, iGrid_GerencialCusto_Col) = "" Or GridCusto.TextMatrix(iIndice1, iGrid_GerencialCusto_Col) = "0" Then
                objPadraoContabItem.iGerencial = LANCAMENTO_STATUS_NAO_GERENCIAL
            ElseIf GridCusto.TextMatrix(iIndice1, iGrid_GerencialCusto_Col) = "1" Then
                objPadraoContabItem.iGerencial = LANCAMENTO_STATUS_GERENCIAL
            ElseIf GridCusto.TextMatrix(iIndice1, iGrid_GerencialCusto_Col) = "2" Then
                objPadraoContabItem.iGerencial = LANCAMENTO_STATUS_AMBOS
            End If
        End If
        
        'Armazena o objeto objPadraoContabItem na coleção colPadraoContabItem
        colPadraoContabItem.Add objPadraoContabItem
        
    Next
    
    Grid_Custo = SUCESSO

    Exit Function

Erro_Grid_Custo:

    Grid_Custo = Err

    Select Case Err
    
        Case 33594
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_ContaCusto_Col
            
        Case 33595
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_CclCusto_Col
            
        Case 33596
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_CreditoCusto_Col
            
        Case 33597
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_DebitoCusto_Col
            
        Case 33598
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_HistoricoCusto_Col
            
        Case 36399
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_ProdutoCusto_Col
            
        Case 36442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_CONTA_NAO_PREENCHIDA", Err, iIndice1)
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_ContaCusto_Col
        
        Case 36443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_DEBCRE_NAO_PREENCHIDA", Err, iIndice1)
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_CreditoCusto_Col
            
        Case 36444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMULA_PRODUTO_NAO_PREENCHIDA", Err, iIndice1)
            TabStrip1.Tabs.Item(TAB_CUSTOS).Selected = True
            GridCusto.SetFocus
            GridCusto.Row = iIndice1
            GridCusto.Col = iGrid_ProdutoCusto_Col
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157901)
            
    End Select
    
    Exit Function

End Function

Function Traz_Conta_Tela(objPlanoConta As ClassPlanoConta) As Long
'Mostra na tela a conta selecionada

Dim lErro As Long
Dim iPos As Integer
Dim sConta1 As String
Dim sContaMascarada As String
Dim iTamanho As Integer
Dim objControl As Control

On Error GoTo Erro_Traz_Conta_Tela
    
    sContaMascarada = String(STRING_CONTA, 0)
    
    lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
    If lErro <> SUCESSO Then Error 36475
    
    sConta1 = Chr(34) & sContaMascarada & Chr(34)
    
    If iFrameAtual = TAB_LANCAMENTOS Then
    
        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then

            Select Case GridLancamentos.Col

                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, sConta1)
                
                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, sConta1)
                    
                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, sConta1)
                    
                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, sConta1)
                    
                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, sConta1)

            End Select
            
        End If

    ElseIf iFrameAtual = TAB_CUSTOS Then

        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then

            Select Case GridCusto.Col

                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, sConta1)
                    
                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, sConta1)
                    
                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, sConta1)
                    
                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, sConta1)
                    
                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, sConta1)
                    
                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, sConta1)
                    
            End Select

        End If

    End If
    
    Traz_Conta_Tela = SUCESSO
        
    Exit Function
    
Erro_Traz_Conta_Tela:

    Traz_Conta_Tela = Err

    Select Case Err
    
        Case 36475
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararConta", Err, objPlanoConta.sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157902)
            
    End Select
        
    Exit Function
        
End Function

Private Sub Posiciona_Texto_Tela(objControl As Control, sTexto As String)
'posiciona o texto sTexto no controle objControl da tela

Dim iPos As Integer
Dim iTamanho As Integer
Dim objGrid As Object

    iPos = objControl.SelStart
    objControl.Text = Mid(objControl.Text, 1, iPos) & sTexto & Mid(objControl.Text, iPos + 1, Len(objControl.Text))
    objControl.SelStart = iPos + Len(sTexto)
    
    If Not (Me.ActiveControl Is objControl) Then
    
        If iFrameAtual = TAB_LANCAMENTOS Then
            Set objGrid = GridLancamentos
        Else
            Set objGrid = GridCusto
        End If
        
        If iPos >= Len(objGrid.TextMatrix(objGrid.Row, objGrid.Col)) Then
            iTamanho = 0
        Else
            iTamanho = Len(objGrid.TextMatrix(objGrid.Row, objGrid.Col)) - iPos
        End If
        objGrid.TextMatrix(objGrid.Row, objGrid.Col) = Mid(objGrid.TextMatrix(objGrid.Row, objGrid.Col), 1, iPos) & sTexto & Mid(objGrid.TextMatrix(objGrid.Row, objGrid.Col), iPos + 1, iTamanho)
        
        iAlterado = REGISTRO_ALTERADO
        
    End If

End Sub

Function Traz_Ccl_Tela(objCcl As ClassCcl) As Long
'Mostra na tela o Ccl selecionado
    
Dim lErro As Long
Dim iPos As Integer
Dim sCcl1 As String
Dim sCclMascarado As String

On Error GoTo Erro_Traz_Ccl_Tela

    sCclMascarado = String(STRING_CONTA, 0)
    
    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 36476

    sCcl1 = Chr(34) & sCclMascarado & Chr(34)
    
    
    If iFrameAtual = TAB_LANCAMENTOS Then
    
        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then

            Select Case GridLancamentos.Col

                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, sCcl1)

                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, sCcl1)
                    
                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, sCcl1)
                    
                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, sCcl1)
                    
                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, sCcl1)

            End Select

        End If

    ElseIf iFrameAtual = TAB_CUSTOS Then

        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then

            Select Case GridCusto.Col

                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, sCcl1)

                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, sCcl1)
                    
                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, sCcl1)
                    
                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, sCcl1)
                    
                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, sCcl1)
                    
                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, sCcl1)
                    
            End Select

        End If

    End If
    
    Traz_Ccl_Tela = SUCESSO
        
    Exit Function
    
Erro_Traz_Ccl_Tela:

    Traz_Ccl_Tela = Err

    Select Case Err
    
        Case 36476
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157903)
            
    End Select
        
    Exit Function
        
End Function

Function Traz_Historico_Tela(objHistPadrao As ClassHistPadrao) As Long
'Mostra na tela o Histórico Padrão selecionado
    
Dim lErro As Long
Dim iPos As Integer
Dim sHistorico1 As String

On Error GoTo Erro_Traz_Historico_Tela

    sHistorico1 = Chr(34) & objHistPadrao.sDescHistPadrao & Chr(34)
    
    If iFrameAtual = TAB_LANCAMENTOS Then
    
        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then

            Select Case GridLancamentos.Col

                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, sHistorico1)

                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, sHistorico1)
                    
                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, sHistorico1)
                    
                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, sHistorico1)
                    
                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, sHistorico1)

            End Select

        End If

    ElseIf iFrameAtual = TAB_CUSTOS Then

        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then

            Select Case GridCusto.Col

                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, sHistorico1)

                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, sHistorico1)
                    
                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, sHistorico1)
                    
                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, sHistorico1)
                    
                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, sHistorico1)
                    
                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, sHistorico1)
                    
            End Select

        End If

    End If
    
    Traz_Historico_Tela = SUCESSO
        
    Exit Function
    
Erro_Traz_Historico_Tela:

    Traz_Historico_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157904)
            
    End Select
        
    Exit Function
        
End Function

Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Mostra na tela o Produto selecionado
    
Dim lErro As Long
Dim iPos As Integer
Dim sProduto1 As String
Dim sProdutoMascarado As String

On Error GoTo Erro_Traz_Produto_Tela

    sProdutoMascarado = String(STRING_PRODUTO, 0)
    
    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then Error 36477

    sProduto1 = Chr(34) & sProdutoMascarado & Chr(34)
    
    If iFrameAtual = TAB_LANCAMENTOS Then
    
        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then

            Select Case GridLancamentos.Col

                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, sProduto1)

                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, sProduto1)
                    
                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, sProduto1)
                    
                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, sProduto1)
                    
                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, sProduto1)

            End Select

        End If

    ElseIf iFrameAtual = TAB_CUSTOS Then

        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then

            Select Case GridCusto.Col

                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, sProduto1)

                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, sProduto1)
                    
                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, sProduto1)
                    
                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, sProduto1)
                    
                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, sProduto1)
                    
                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, sProduto1)
                    
            End Select

        End If

    End If
    
    Traz_Produto_Tela = SUCESSO
        
    Exit Function
    
Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = Err

    Select Case Err
    
        Case 36477
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157905)
            
    End Select
        
    Exit Function
        
End Function

Function Funcoes1(sFuncao As String) As Long

Dim iPos As Integer

    If iFrameAtual = TAB_LANCAMENTOS Then
    
        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then
            
            Select Case GridLancamentos.Col
            
                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, Funcoes.Text)
                    
                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, Funcoes.Text)
                    
                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, Funcoes.Text)
                    
                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, Funcoes.Text)
                    
                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, Funcoes.Text)
                   
            End Select
            
        End If
        
    ElseIf iFrameAtual = TAB_CUSTOS Then
    
        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then
            
            Select Case GridCusto.Col
            
                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, Funcoes.Text)
                    
                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, Funcoes.Text)
                    
                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, Funcoes.Text)
                    
                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, Funcoes.Text)
                    
                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, Funcoes.Text)
                   
                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, Funcoes.Text)
                   
            End Select
            
        End If
    
    End If
    
    Funcoes1 = SUCESSO

End Function

Function Mnemonicos1(sMnemonico As String) As Long

Dim iPos As Integer

    If iFrameAtual = TAB_LANCAMENTOS Then
    
        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then
            
            Select Case GridLancamentos.Col
            
                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, Mnemonicos.Text)
                    
                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, Mnemonicos.Text)
                    
                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, Mnemonicos.Text)
                    
                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, Mnemonicos.Text)
                    
                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, Mnemonicos.Text)
                   
            End Select
            
            If GridLancamentos.Row - GridLancamentos.FixedRows = objGrid1.ilinhasExistentes Then
                objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
            End If
        
        End If
        
    ElseIf iFrameAtual = TAB_CUSTOS Then
        
        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then
            
            Select Case GridCusto.Col
            
                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, Mnemonicos.Text)
                    
                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, Mnemonicos.Text)
                    
                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, Mnemonicos.Text)
                    
                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, Mnemonicos.Text)
                    
                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, Mnemonicos.Text)
            
                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, Mnemonicos.Text)
            
            End Select
            
            If GridCusto.Row - GridCusto.FixedRows = objGridCusto.ilinhasExistentes Then
                objGridCusto.ilinhasExistentes = objGridCusto.ilinhasExistentes + 1
            End If
        
        End If
            
    End If

End Function

Function Operadores1() As Long

Dim iPos As Integer

    If iFrameAtual = TAB_LANCAMENTOS Then

        If GridLancamentos.Row > 0 And GridLancamentos.Row <= objGrid1.ilinhasExistentes + 1 And GridLancamentos.Col > 0 Then

            Select Case GridLancamentos.Col

                Case iGrid_Conta_Col
                    Call Posiciona_Texto_Tela(Conta, Operadores.Text)

                Case iGrid_Ccl_Col
                    Call Posiciona_Texto_Tela(Ccl, Operadores.Text)

                Case iGrid_Credito_Col
                    Call Posiciona_Texto_Tela(Credito, Operadores.Text)

                Case iGrid_Debito_Col
                    Call Posiciona_Texto_Tela(Debito, Operadores.Text)

                Case iGrid_Historico_Col
                    Call Posiciona_Texto_Tela(Historico, Operadores.Text)

            End Select

        End If

    ElseIf iFrameAtual = TAB_CUSTOS Then

        If GridCusto.Row > 0 And GridCusto.Row <= objGridCusto.ilinhasExistentes + 1 And GridCusto.Col > 0 Then

            Select Case GridCusto.Col

                Case iGrid_ContaCusto_Col
                    Call Posiciona_Texto_Tela(ContaCusto, Operadores.Text)

                Case iGrid_CclCusto_Col
                    Call Posiciona_Texto_Tela(CclCusto, Operadores.Text)

                Case iGrid_CreditoCusto_Col
                    Call Posiciona_Texto_Tela(CreditoCusto, Operadores.Text)

                Case iGrid_DebitoCusto_Col
                    Call Posiciona_Texto_Tela(DebitoCusto, Operadores.Text)

                Case iGrid_HistoricoCusto_Col
                    Call Posiciona_Texto_Tela(HistoricoCusto, Operadores.Text)

                Case iGrid_ProdutoCusto_Col
                    Call Posiciona_Texto_Tela(ProdutoCusto, Operadores.Text)

            End Select

        End If

    End If

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PADRAO_CONTABILIZACAO
    Set Form_Load_Ocx = Me
    Caption = "Padrão de Contabilização"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PadraoContab"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ContaCusto Or Me.ActiveControl Is Conta Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is CclCusto Or Me.ActiveControl Is Ccl Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is HistoricoCusto Or Me.ActiveControl Is Historico Then
            Call BotaoHistorico_Click
        ElseIf Me.ActiveControl Is ProdutoCusto Then
            Call BotaoProduto_Click
        End If
    
    ElseIf KeyCode = KEYCODE_VERIFICAR_SINTAXE Then
    
        If Checkbox_Verifica_Sintaxe.Value = MARCADO Then
            Checkbox_Verifica_Sintaxe.Value = DESMARCADO
        Else
            Checkbox_Verifica_Sintaxe.Value = MARCADO
        End If
            
    End If
    

End Sub




Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

