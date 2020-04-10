VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTLancamentosAt"
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

Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_SeqContraPartida_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim iGrid_Gerencial_Col As Integer

Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Private WithEvents objEventoDocAuto As AdmEvento
Attribute objEventoDocAuto.VB_VarHelpID = -1
Private WithEvents objEventoLancamento As AdmEvento
Attribute objEventoLancamento.VB_VarHelpID = -1

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoHist As AdmEvento
Attribute objEventoHist.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get BotaoEstornar() As Object
     Set BotaoEstornar = objUserControl.Controls("BotaoEstornar")
End Property

Public Property Get BotaoConsultar() As Object
     Set BotaoConsultar = objUserControl.Controls("BotaoConsultar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get ListHistoricos() As Object
     Set ListHistoricos = objUserControl.Controls("ListHistoricos")
End Property

Public Property Get ListDocAuto() As Object
     Set ListDocAuto = objUserControl.Controls("ListDocAuto")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get BotaoAplicar() As Object
     Set BotaoAplicar = objUserControl.Controls("BotaoAplicar")
End Property

Public Property Get DocAuto() As Object
     Set DocAuto = objUserControl.Controls("DocAuto")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get CclDescricao() As Object
     Set CclDescricao = objUserControl.Controls("CclDescricao")
End Property

Public Property Get ContaDescricao() As Object
     Set ContaDescricao = objUserControl.Controls("ContaDescricao")
End Property

Public Property Get CclLabel() As Object
     Set CclLabel = objUserControl.Controls("CclLabel")
End Property

Public Property Get SeqContraPartida() As Object
     Set SeqContraPartida = objUserControl.Controls("SeqContraPartida")
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

Public Property Get TvwContas() As Object
     Set TvwContas = objUserControl.Controls("TvwContas")
End Property

Public Property Get UpDown1() As Object
     Set UpDown1 = objUserControl.Controls("UpDown1")
End Property

Public Property Get GridLancamentos() As Object
     Set GridLancamentos = objUserControl.Controls("GridLancamentos")
End Property

Public Property Get TvwCcls() As Object
     Set TvwCcls = objUserControl.Controls("TvwCcls")
End Property

Public Property Get Documento() As Object
     Set Documento = objUserControl.Controls("Documento")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get LabelHistoricos() As Object
     Set LabelHistoricos = objUserControl.Controls("LabelHistoricos")
End Property

Public Property Get LabelDocAuto() As Object
     Set LabelDocAuto = objUserControl.Controls("LabelDocAuto")
End Property

Public Property Get TotalCredito() As Object
     Set TotalCredito = objUserControl.Controls("TotalCredito")
End Property

Public Property Get TotalDebito() As Object
     Set TotalDebito = objUserControl.Controls("TotalDebito")
End Property

Public Property Get LabelTotais() As Object
     Set LabelTotais = objUserControl.Controls("LabelTotais")
End Property

Public Property Get Line1() As Object
     Set Line1 = objUserControl.Controls("Line1")
End Property

Public Property Get LabelCCL() As Object
     Set LabelCCL = objUserControl.Controls("LabelCcl")
End Property

Public Property Get LabelContas() As Object
     Set LabelContas = objUserControl.Controls("LabelContas")
End Property

Public Property Get Origem() As Object
     Set Origem = objUserControl.Controls("Origem")
End Property

Public Property Get Periodo() As Object
     Set Periodo = objUserControl.Controls("Periodo")
End Property

Public Property Get Exercicio() As Object
     Set Exercicio = objUserControl.Controls("Exercicio")
End Property

Public Property Get DocumentoLabel() As Object
     Set DocumentoLabel = objUserControl.Controls("DocumentoLabel")
End Property

Public Property Get Gerencial() As Object
     Set Gerencial = objUserControl.Controls("Gerencial")
End Property


'--- fim dos properties get dos controles da tela

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lDoc As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo


On Error GoTo Erro_BotaoProxNum_Click

    If Len(Data.ClipText) = 0 Then Error 55715

    'Obtém Periodo e Exercicio correspondentes à data
    dtData = CDate(Data.Text)

    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 55714

    'Mostra número do próximo voucher(documento) disponível
    lErro = CF("Voucher_Automatico", giFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, MODULO_CONTABILIDADE, lDoc)
    If lErro <> SUCESSO Then Error 55702

    Documento.Text = CStr(lDoc)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 55702, 55714
        
        Case 55715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXNUM_DATA_NAO_PREENCHIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155990)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoAplicar_Click()

Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_BotaoAplicar_click

    If Len(DocAuto.ClipText) = 0 Then Error 11430
    
    lDoc = CLng(DocAuto.ClipText)
    
    lErro = Traz_DocAuto_Tela(lDoc)
    If lErro <> SUCESSO Then Error 11431

    Exit Sub

Erro_BotaoAplicar_click:

    Select Case Err

        Case 11430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
        
        Case 11431
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155991)
    
    End Select

    Exit Sub

End Sub

Public Sub Ccl_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Conta_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Conta_GotFocus()

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_Conta_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwContas.Visible = True
'    LabelContas.Visible = True
'    TvwCcls.Visible = False
'    LabelCCL.Visible = False
'    ListDocAuto.Visible = False
'    LabelDocAuto.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
    
    sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col)
    
    If Len(sConta) > 0 Then

        lErro = Conta_Exibe_Descricao(sConta)
        If lErro <> SUCESSO Then Error 11432
        
    Else

        ContaDescricao = ""

    End If
    
    Exit Sub
    
Erro_Conta_GotFocus:

    Select Case Err
    
        Case 11432
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155992)
    
    End Select
    
    Exit Sub
    
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

Dim sCcl As String
Dim lErro As Long

On Error GoTo Erro_Ccl_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    TvwCcls.Visible = True
'    LabelCCL.Visible = True
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListDocAuto.Visible = False
'    LabelDocAuto.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
   
    sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)
    
    'Coloca descricao de Ccl no panel
    If Len(sCcl) > 0 Then

        lErro = Ccl_Exibe_Descricao(sCcl)
        If lErro <> SUCESSO Then Error 11433

    Else
    
        CclDescricao = ""
        
    End If
    
    Exit Sub
    
Erro_Ccl_GotFocus:

    Select Case Err
    
        Case 11433
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155993)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub DocAuto_GotFocus()
    
'    ListDocAuto.Visible = True
'    LabelDocAuto.Visible = True
'    TvwCcls.Visible = False
'    LabelCCL.Visible = False
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListHistoricos.Visible = False
'    LabelHistoricos.Visible = False
    
    Call MaskEdBox_TrataGotFocus(DocAuto, iAlterado)

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

Public Sub Documento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Documento, iAlterado)

End Sub

Public Sub SeqContraPartida_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub SeqContraPartida_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid1)

End Sub

Public Sub SeqContraPartida_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid1)
    
End Sub

Public Sub SeqContraPartida_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = SeqContraPartida
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO
    
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

Public Sub DocAuto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Documento_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Exercicio_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
       
    Set objEventoDocAuto = Nothing
    Set objEventoLancamento = Nothing
    Set objGrid1 = Nothing
    
    Set objEventoConta = Nothing
    Set objEventoCcl = Nothing
    Set objEventoHist = Nothing
    
End Sub

Public Sub GridLancamentos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

'    If giTipoVersao = VERSAO_FULL Then
        'Verifica se foi o botao direito do mouse que foi pressionado
        If Button = vbRightButton Then
            Set PopupMenusCTB.objTela = Me
            objUserControl.Popup_Menu PopupMenusCTB.MenuGrid
            Set PopupMenusCTB.objTela = Nothing
        End If
'    End If
    
End Sub

Public Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Historico_GotFocus()
    
Dim iPos As Integer

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
'    ListHistoricos.Visible = True
'    LabelHistoricos.Visible = True
'    TvwCcls.Visible = False
'    LabelCCL.Visible = False
'    TvwContas.Visible = False
'    LabelContas.Visible = False
'    ListDocAuto.Visible = False
'    LabelDocAuto.Visible = False
    
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
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Historico_KeyPress

    'se digitou ENTER
    If KeyAscii = vbKeyReturn Then
        
        If Len(Historico.Text) > 0 Then
        
            If left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(Historico.Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44070
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44071
                
                If lErro = 5446 Then Error 44072
        
                Historico.Text = objHistPadrao.sDescHistPadrao
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
    
        Case 44070
            objGrid1.iExecutaSaidaCelula = 0
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            objGrid1.iExecutaSaidaCelula = 1
        
        Case 44071

        Case 44072
            objGrid1.iExecutaSaidaCelula = 0
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Chama_Tela("HistoricoPadrao", objHistPadrao)
            
            Else
                Historico.SetFocus
            End If
            
            objGrid1.iExecutaSaidaCelula = 1
            
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155994)
    
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

Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGrid1)
    If lErro <> SUCESSO Then Error 44066

    Call Trata_SeqContraPartida(GridLancamentos.Row)
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case Err
    
        Case 44066
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155995)
    
    End Select

    Exit Sub
    
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

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoDocAuto = New AdmEvento
    Set objEventoLancamento = New AdmEvento
            
    Set objEventoConta = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoHist = New AdmEvento
            
    TvwContas.Visible = False
    LabelContas.Visible = False
    TvwCcls.Visible = False
    LabelCCL.Visible = False
    ListDocAuto.Visible = False
    LabelDocAuto.Visible = False
    ListHistoricos.Visible = False
    LabelHistoricos.Visible = False
            
    Set objGrid1 = New AdmGrid
    
    lErro = Inicializa_Grid_Lancamento(objGrid1)
    If lErro <> SUCESSO Then Error 11436

'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then Error 11437
'
'    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
'
'        'Inicializa a Lista de Centros de Custo
'        lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
'        If lErro <> SUCESSO Then Error 11438
'
'    End If
'
'    'Inicializa a Lista de Historicos
'    lErro = Carga_Lista_Historico()
'    If lErro <> SUCESSO Then Error 11439
'
'    'Inicializa a Lista de Documentos Automaticos
'    lErro = Carga_Lista_DocAuto()
'    If lErro <> SUCESSO Then Error 11440

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        CclLabel.Visible = True
        CclDescricao.Visible = True
    End If
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 11436, 11437, 11438, 11439, 11440
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155996)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long
Dim lDoc As Long
Dim objDocAuto As ClassDocAuto
Dim objLancamento_Cabecalho As ClassLancamento_Cabecalho

On Error GoTo Erro_Trata_Parametros
    
    lErro = Traz_Cabecalho_Tela()
    If lErro <> SUCESSO Then Error 11441
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 11441
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155997)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 10500
    
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 10501

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 10502
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "S" & sCclPai
            
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
            'se o centro de custo/lucro não possui centro de custo/lucro "pai"
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 10500
        
        Case 10501
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case 10502
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155998)

    End Select
    
    Exit Function

End Function

Private Function Carga_Lista_Historico() As Long
'move os dados de historico padrão do banco de dados para a arvore colNodes.

Dim colHistPadrao As New Collection
Dim objHistPadrao As ClassHistPadrao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Lista_Historico
    
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 11448
    
    For Each objHistPadrao In colHistPadrao
        
        ListHistoricos.AddItem CStr(objHistPadrao.iHistPadrao) & SEPARADOR & objHistPadrao.sDescHistPadrao
        ListHistoricos.ItemData(ListHistoricos.NewIndex) = objHistPadrao.iHistPadrao
        
    Next
    
    Carga_Lista_Historico = SUCESSO

    Exit Function

Erro_Carga_Lista_Historico:

    Carga_Lista_Historico = Err

    Select Case Err

        Case 11448

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155999)

    End Select
    
    Exit Function

End Function

Private Function Carga_Lista_DocAuto() As Long
'move os dados do plano de contas do banco de dados para a arvore colNodes.

Dim colDocAuto As New Collection
Dim objDocAuto As ClassDocAuto
Dim lErro As Long
    
On Error GoTo Erro_Carga_Lista_DocAuto
    
    'leitura das contas no BD
    lErro = CF("DocAuto_Le_Todos", colDocAuto)
    If lErro <> SUCESSO Then Error 11449
    
    For Each objDocAuto In colDocAuto
        
        ListDocAuto.AddItem CStr(objDocAuto.lDoc) & SEPARADOR & objDocAuto.sDescricao
        
    Next
    
    Carga_Lista_DocAuto = SUCESSO

    Exit Function

Erro_Carga_Lista_DocAuto:

    Carga_Lista_DocAuto = Err

    Select Case Err

        Case 11449
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156000)

    End Select
    
    Exit Function

End Function

Private Function Traz_Cabecalho_Tela() As Long

Dim sDescricao As String
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
Dim iIndice As Integer
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lErro As Long

On Error GoTo Erro_Traz_Cabecalho_Tela

    'Inicializa Data
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Coloca o periodo relativo a data na tela
    lErro = CF("Periodo_Le", gdtDataAtual, objPeriodo)
    If lErro <> SUCESSO Then Error 11450
    
    Periodo.Caption = objPeriodo.sNomeExterno
    
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 11451
    
    'se o exercicio não está cadastrado ==> erro
    If lErro = 10083 Then Error 10092
    
    Exercicio.Caption = objExercicio.sNomeExterno

    Traz_Cabecalho_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Cabecalho_Tela:

    Traz_Cabecalho_Tela = Err

    Select Case Err
    
        Case 10092
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
        
        Case 11450, 11451
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156001)
    
    End Select
    
    Exit Function

End Function

Private Function Inicializa_Grid_Lancamento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_Lancamento
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Conta")
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colColuna.Add ("CCusto")
    objGridInt.colColuna.Add ("Débito")
    objGridInt.colColuna.Add ("Crédito")
    objGridInt.colColuna.Add ("C.P.")
    objGridInt.colColuna.Add ("Histórico")
    If giContabGerencial = 1 Then objGridInt.colColuna.Add ("Status")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Conta.Name)
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Debito.Name)
    objGridInt.colCampo.Add (Credito.Name)
    objGridInt.colCampo.Add (SeqContraPartida.Name)
    objGridInt.colCampo.Add (Historico.Name)
    If giContabGerencial = 1 Then objGridInt.colCampo.Add (Gerencial.Name)
    
    'indica onde estão situadas as colunas do grid
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        iGrid_Conta_Col = 1
        iGrid_Ccl_Col = 2
        iGrid_Debito_Col = 3
        iGrid_Credito_Col = 4
        iGrid_SeqContraPartida_Col = 5
        iGrid_Historico_Col = 6
    Else
        iGrid_Conta_Col = 1
        '999 indica que não está sendo usado
        iGrid_Ccl_Col = 999
        iGrid_Debito_Col = 2
        iGrid_Credito_Col = 3
        iGrid_SeqContraPartida_Col = 4
        iGrid_Historico_Col = 5
        Ccl.Visible = False
    End If
    
    If giContabGerencial = 1 Then
        iGrid_Gerencial_Col = iGrid_Historico_Col + 1
    Else
        Gerencial.Visible = False
    End If

    lErro = Inicializa_Mascaras()
    If lErro <> SUCESSO Then Error 11458

    objGridInt.objGrid = GridLancamentos
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = 400
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12 '7
        
    GridLancamentos.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalDebito.top = GridLancamentos.top + GridLancamentos.Height
    TotalDebito.left = GridLancamentos.left
    For iIndice = 0 To iGrid_Debito_Col - 1
        TotalDebito.left = TotalDebito.left + GridLancamentos.ColWidth(iIndice) + GridLancamentos.GridLineWidth + 20
    Next
    
    TotalDebito.Width = GridLancamentos.ColWidth(iGrid_Debito_Col)
    
    TotalCredito.top = TotalDebito.top
    TotalCredito.left = TotalDebito.left + TotalDebito.Width + GridLancamentos.GridLineWidth
    TotalCredito.Width = GridLancamentos.ColWidth(iGrid_Credito_Col)
    
    LabelTotais.top = TotalCredito.top + (TotalDebito.Height - LabelTotais.Height) / 2
    LabelTotais.left = TotalDebito.left - LabelTotais.Width

    Inicializa_Grid_Lancamento = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_Lancamento:

    Inicializa_Grid_Lancamento = Err
    
    Select Case Err
    
        Case 11458
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156002)
        
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
    If lErro <> SUCESSO Then Error 11459
    
    Conta.Mask = sMascaraConta
    
    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
        sMascaraCcl = String(STRING_CCL, 0)

        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 11460

        Ccl.Mask = sMascaraCcl
        
    End If
    
    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 11459, 11460
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156003)
        
    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridLancamentos.Col
    
            Case iGrid_Conta_Col
            
                lErro = Saida_Celula_Conta(objGridInt)
                If lErro <> SUCESSO Then gError 11461
                
            Case iGrid_Ccl_Col
            
                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then gError 11462
                
            Case iGrid_Credito_Col
            
                lErro = Saida_Celula_Credito(objGridInt)
                If lErro <> SUCESSO Then gError 11463
                
            Case iGrid_Debito_Col
            
                lErro = Saida_Celula_Debito(objGridInt)
                If lErro <> SUCESSO Then gError 11464

            Case iGrid_SeqContraPartida_Col
            
                lErro = Saida_Celula_SeqContraPartida(objGridInt)
                If lErro <> SUCESSO Then gError 20609

            Case iGrid_Historico_Col
            
                lErro = Saida_Celula_Historico(objGridInt)
                If lErro <> SUCESSO Then gError 11465
               

            Case iGrid_Gerencial_Col
            
                lErro = Saida_Celula_Gerencial(objGridInt)
                If lErro <> SUCESSO Then gError 188070

        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 11466
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = gErr
    
    Select Case gErr
    
        Case 11461, 11462, 11463, 11464, 11465, 20609, 188070
        
        Case 11466
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156004)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim lPosicaoSeparador As Long
Dim sContaFormatada As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim objContaCcl As New ClassContaCcl
Dim sCcl As String
Dim vbMsgRes As VbMsgBoxResult
Dim objPlanoConta As New ClassPlanoConta
Dim sContaMascarada As String

On Error GoTo Erro_Saida_Celula_Conta

    Set objGridInt.objControle = Conta
    
    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica", Conta.Text, Conta.ClipText, objPlanoConta)
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 20604

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 20642

        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True

        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            For iIndice = 0 To ListHistoricos.ListCount - 1
                If ListHistoricos.ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                    ListHistoricos.ListIndex = iIndice
                    lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
                    GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
                    Exit For
                End If
            Next
            
        End If

    'se não encontrou a conta simples
    ElseIf lErro = 44033 Or lErro = 44037 Then
    
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 11467
                
        'conta não cadastrada
        If lErro = 5700 Then Error 11468
        
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            For iIndice = 0 To ListHistoricos.ListCount - 1
                If ListHistoricos.ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                    ListHistoricos.ListIndex = iIndice
                    lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
                    GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
                    Exit For
                End If
            Next
            
        End If
                               
    End If
    
    'se a conta foi preenchida
    If Len(Trim(Conta.ClipText)) > 0 Then
    
        'se utiliza centro de custo extra-contabil
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            'se o centro de custo foi preenchido
            If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)) > 0 Then
            
                'verifica se a associação da conta com o centro de custo está cadastrado
                sCcl = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Ccl_Col)
        
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 11469
        
                objContaCcl.sConta = sContaFormatada
                objContaCcl.sCcl = sCclFormatada
        
                lErro = CF("ContaCcl_Le", objContaCcl)
                If lErro <> SUCESSO And lErro <> 5871 Then Error 11470
        
                'associação Conta x Centro de Custo/Lucro não cadastrada
                If lErro = 5871 Then Error 11471
                
            End If
            
        End If
                
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
                
        ContaDescricao.Caption = objPlanoConta.sDescConta
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11472

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = Err
    
    Select Case Err
    
        Case 11467, 11469, 11470, 11472, 20604
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 11468
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Conta.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
                
            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
            
        Case 11471
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", Conta.Text, sCcl)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case 20642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156005)
        
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
    If lErro <> SUCESSO And lErro <> 5703 Then Error 11473
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 11474
                
    'se o centro de custo foi preenchido
    If Len(Ccl.ClipText) > 0 Then
    
        'se a conta foi informada
        If Len(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)) > 0 Then
    
            'verificar se a associação da conta com o centro de custo em questão está cadastrada
            sConta = GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Conta_Col)
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 11475
        
            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada
        
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 11476
        
            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 11477
        
        End If
                        
        CclDescricao.Caption = objCcl.sDescCcl
             
'        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
'            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
'        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11478

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err
    
    Select Case Err
    
        Case 11473, 11475, 11476, 11478
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 11474
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
            
        Case 11477
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, Ccl.Text)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156006)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Credito(objGridInt As AdmGrid) As Long
'faz a critica da celula credito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Credito

    Set objGridInt.objControle = Credito
    
    If Len(Credito.Text) > 0 Then
    
        lErro = Valor_Critica(Credito.Text)
        If lErro <> SUCESSO Then Error 11479
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11480
                
    If Len(Credito.Text) > 0 Then
        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Debito_Col) = ""
              
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
    End If
        
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")

    Saida_Celula_Credito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Credito:

    Saida_Celula_Credito = Err
    
    Select Case Err
    
        Case 11479, 11480
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156007)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Debito(objGridInt As AdmGrid) As Long
'faz a critica da celula debito do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Debito

    Set objGridInt.objControle = Debito
    
    If Len(Debito.Text) > 0 Then
    
        lErro = Valor_Critica(Debito.Text)
        If lErro <> SUCESSO Then Error 11481
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11482
                
    If Len(Debito.Text) > 0 Then
        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Credito_Col) = ""
        
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
    End If
    
    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")

    Saida_Celula_Debito = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Debito:

    Saida_Celula_Debito = Err
    
    Select Case Err
    
        Case 11481, 11482
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156008)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_SeqContraPartida(objGridInt As AdmGrid) As Long
'faz a critica da celula sequencial de contra partida do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SeqContraPartida

    Set objGridInt.objControle = SeqContraPartida
    
    If Len(SeqContraPartida.Text) > 0 Then
    
        If GridLancamentos.Row = CInt(SeqContraPartida.Text) Then Error 20610
        
        If CInt(SeqContraPartida.Text) > objGridInt.ilinhasExistentes Then Error 20611
        
        If CInt(SeqContraPartida.Text) <= 0 Then Error 20612
    
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20613
                
    Saida_Celula_SeqContraPartida = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_SeqContraPartida:

    Saida_Celula_SeqContraPartida = Err
    
    Select Case Err
    
        Case 20610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_NAO_MESMO_LANCAMENTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 20611, 20612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_LANCAMENTO_INEXISTENTE", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 20613
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156009)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'faz a critica da celula historico do grid que está deixando de ser a corrente

Dim sValor As String
Dim lErro As Long
Dim objHistPadrao As ClassHistPadrao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Historico

    Set objHistPadrao = New ClassHistPadrao
    
    Set objGridInt.objControle = Historico
                
    If left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
    
        sValor = Trim(Mid(Historico.Text, 2))
        
        lErro = Valor_Inteiro_Critica(sValor)
        If lErro <> SUCESSO Then Error 11483
        
        objHistPadrao.iHistPadrao = CInt(sValor)
                
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then Error 11484
        
        If lErro = 5446 Then Error 11485

        Historico.Text = objHistPadrao.sDescHistPadrao
        
    End If
    
                
    If Len(Historico.Text) > 0 Then
        If GridLancamentos.Row - GridLancamentos.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11486

    Saida_Celula_Historico = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = Err
    
    Select Case Err
    
        Case 11483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 11484, 11486
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 11485
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("HistoricoPadrao", objHistPadrao)
                
            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156010)
        
    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double
    
Dim dAcumulador As Double
Dim iLinha As Integer
    
    dAcumulador = 0
    
    For iLinha = 1 To objGrid1.ilinhasExistentes
        If Len(GridLancamentos.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(GridLancamentos.TextMatrix(iLinha, iColuna))
        End If
    Next
    
    GridColuna_Soma = dAcumulador

End Function

Public Sub BotaoGravar_Click()

    Call Gravar_Registro
    
    iAlterado = 0
    
End Sub

Public Function Gravar_Registro() As Long
    
Dim lErro As Long
Dim lDoc As Long
Dim colLancamento_Detalhe As New Collection
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iIndice1 As Integer
Dim dSoma As Double
Dim iPeriodoDoc As Integer
Dim iExercicioDoc As Integer
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Data, determinação dos exercicio e periodo correspondentes
    If Len(Data.ClipText) = 0 Then Error 11490
    
    'Verifica a existencia de pelo menos um lançamento
    If objGrid1.ilinhasExistentes = 0 Then Error 11491
        
    'Documento
    If Len(Documento.ClipText) = 0 Then Error 11492
        
   'Preenche Objeto Lançamento_Cabeçalho
   objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
   objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Caption)
   objLancamento_Cabecalho.lDoc = CLng(Documento.ClipText)
   objLancamento_Cabecalho.dtData = CDate(Data.Text)
   
   'Preenche Objeto Lançamento_Detalhe
   lErro = Grid_Lancamento_Detalhe(colLancamento_Detalhe)
   If lErro <> SUCESSO Then Error 11493
       
   'Testa se soma dos débitos é igual a soma dos créditos
   dSoma = 0
   
   For Each objLancamento_Detalhe In colLancamento_Detalhe
       dSoma = dSoma + objLancamento_Detalhe.dValor
   Next
   
   dSoma = Format(dSoma, "Fixed")
   
   If dSoma <> 0 Then Error 11494

   lErro = CF("Lancamento_Atualiza", objLancamento_Cabecalho, colLancamento_Detalhe)
   If lErro <> SUCESSO Then Error 11496
    
    Call Limpa_Tela_Lancamentos
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault

    Select Case Err
    
        Case 11490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)
            Data.SetFocus
            
        Case 11491
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_LANCAMENTOS_GRAVAR", Err)
    
        Case 11492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
            Documento.SetFocus
        
        Case 11493, 11496
        
        Case 11494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_BALANCEADO", Err, objLancamento_Cabecalho.lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156011)
            
    End Select
    
    Exit Function
    
End Function

Private Function Limpa_Tela_Lancamentos() As Long
    
Dim lErro As Long
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    Call Grid_Limpa(objGrid1)
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""
    DocAuto.Text = ""
    Documento.Text = ""
    
    Limpa_Tela_Lancamentos = SUCESSO
    
End Function

Private Function Grid_Lancamento_Detalhe(colLancamento_Detalhe As Collection) As Long

Dim iIndice1 As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim dValorDebito As Double
Dim dValorCredito As Double
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colContraPartida As New Collection

On Error GoTo Erro_Grid_Lancamento_Detalhe

    For iIndice1 = 1 To objGrid1.ilinhasExistentes
        
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
            
        objLancamento_Detalhe.iSeq = iIndice1
            
        sConta = GridLancamentos.TextMatrix(iIndice1, iGrid_Conta_Col)
            
        If Len(Trim(sConta)) = 0 Then Error 11495
            
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 11499
        
        If iContaPreenchida = CONTA_VAZIA Then Error 11500
            
        objLancamento_Detalhe.sConta = sContaFormatada
    
        'Testa para ver se houve crédito ou débito
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col)) > 0 Then
            dValorCredito = CDbl(GridLancamentos.TextMatrix(iIndice1, iGrid_Credito_Col))
        Else
            dValorCredito = 0
        End If
            
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col)) > 0 Then
            dValorDebito = CDbl(GridLancamentos.TextMatrix(iIndice1, iGrid_Debito_Col))
        Else
            dValorDebito = 0
        End If
    
        'Armazena débito ou crédito
        If dValorDebito = 0 And dValorCredito = 0 Then Error 11501
            
        objLancamento_Detalhe.dValor = dValorCredito - dValorDebito
    
        'armazena o sequencial de contra partida, se estiver preenchido
        If Len(GridLancamentos.TextMatrix(iIndice1, iGrid_SeqContraPartida_Col)) > 0 Then
            objLancamento_Detalhe.iSeqContraPartida = CInt(GridLancamentos.TextMatrix(iIndice1, iGrid_SeqContraPartida_Col))
            
            lErro = Armazena_Contra_Partida(colContraPartida, objLancamento_Detalhe)
            If lErro <> SUCESSO Then Error 20614
            
        End If
    
        objLancamento_Detalhe.sProduto = ""
    
        'Armazena Histórico e Ccl
        objLancamento_Detalhe.sHistorico = GridLancamentos.TextMatrix(iIndice1, iGrid_Historico_Col)
            
        'verifica se o historico tem parametros que deveriam ter sido substituidos
        If InStr(objLancamento_Detalhe.sHistorico, CARACTER_HISTORICO_PARAM) <> 0 Then Error 20637
            
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_Ccl_Col <> 999 Then
                
            sCcl = GridLancamentos.TextMatrix(iIndice1, iGrid_Ccl_Col)
            
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 11502
            
            If iCclPreenchida = CCL_PREENCHIDA Then
                objLancamento_Detalhe.sCcl = sCclFormatada
            Else
                objLancamento_Detalhe.sCcl = ""
            End If
                
        End If
                
        objLancamento_Detalhe.iGerencial = GridLancamentos.TextMatrix(iIndice1, iGrid_Gerencial_Col)
                
        'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
        colLancamento_Detalhe.Add objLancamento_Detalhe
                
    Next
    
    lErro = Testa_Contra_Partida(colLancamento_Detalhe, colContraPartida)
    If lErro <> SUCESSO Then Error 20615
    
    Grid_Lancamento_Detalhe = SUCESSO

    Exit Function

Erro_Grid_Lancamento_Detalhe:

    Grid_Lancamento_Detalhe = Err

    Select Case Err
        
        Case 11495, 11500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_GRID_NAO_PREENCHIDA", Err, iIndice1)
            
        Case 11499, 11502, 20614, 20615
        
        Case 11501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_LANCAMENTO_NAO_PREENCHIDO", Err)
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Debito_Col
            GridLancamentos.SetFocus
            
        Case 20637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTORICO_PARAM", Err)
            GridLancamentos.Row = iIndice1
            GridLancamentos.Col = iGrid_Historico_Col
            GridLancamentos.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156012)
            
    End Select
    
    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 11503

    Call Limpa_Tela_Lancamentos
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 11503
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156013)
        
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub Label6_Click()

Dim objDocAuto As New ClassDocAuto
Dim colSelecao As Collection

    If Len(DocAuto.Text) = 0 Then
        objDocAuto.lDoc = 0
    Else
        objDocAuto.lDoc = CLng(DocAuto.ClipText)
    End If

    objDocAuto.iSeq = 0

    Call Chama_Tela("DocAutoLista", colSelecao, objDocAuto, objEventoDocAuto)

End Sub

Public Sub Data_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim lDoc As Long
Dim sNomeExterno As String
Dim vbMsgRes As VbMsgBoxResult
Dim colSelecao As Collection
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Data_Validate
    
    If Len(Data.ClipText) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 11507
    
        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)
    
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 11508
    
        'Verifica se Exercicio está fechado
        lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
        If lErro <> SUCESSO And lErro <> 10083 Then Error 11509
        
        'Se o exercicio não está cadastrado ==> erro
        If lErro = 10083 Then Error 10095
        
        If objExercicio.iStatus = EXERCICIO_FECHADO Then Error 11511
    
        objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
        objPeriodosFilial.iExercicio = objPeriodo.iExercicio
        objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
        objPeriodosFilial.sOrigem = MODULO_CONTABILIDADE
        
        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then Error 10169
        
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 10170
            
        If Periodo.Caption <> objPeriodo.sNomeExterno Or Exercicio.Caption <> objExercicio.sNomeExterno Then
            
            Documento.Text = ""
            
        End If
        
        'Preenche campo de periodo
        Periodo.Caption = objPeriodo.sNomeExterno
    
        Exercicio.Caption = objExercicio.sNomeExterno
    
    Else
    
        Periodo.Caption = ""
    
        Exercicio.Caption = ""
        
    End If
    
    Exit Sub

Erro_Data_Validate:
        
    Cancel = True
        
    Select Case Err
    
        Case 10095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)
    
        Case 10169, 11507, 11508, 11509
    
        Case 10170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, objPeriodosFilial.iPeriodo, objPeriodosFilial.iExercicio)
    
        Case 11511
            'Não é possível fazer lançamentos em exercício fechado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objPeriodo.iExercicio)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156014)
        
    End Select
    
    Exit Sub

End Sub

Public Sub menulimpar_Click()

    Call Grid_Limpa(objGrid1)
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""

End Sub

Public Sub menuRateio_Click()
'Quando a opção Rateio e Selecionada o menu ele chama a tela Rateio

Dim objRateioOn As New ClassRateioOn
Dim lErro As Long
Dim objConfirmaTela As New AdmConfirmaTela

On Error GoTo Erro_menuRateio_Click

    Call Chama_Tela_Modal("Rateio", objRateioOn, objConfirmaTela)
    
    If objConfirmaTela.iTelaOK = OK Then
    
        lErro = Traz_Rateio_Tela(objRateioOn)
        If lErro <> SUCESSO Then Error 9637
        
    End If
    
    Exit Sub

Erro_menuRateio_Click:

    Select Case Err
    
        Case 9637
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156015)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoDocAuto_evSelecao(obj1 As Object)
'Chama a funcao que vai trazer para a tela o Documento autmatico selecionado em DocAutoLista
    
Dim objDocAuto As ClassDocAuto
Dim lErro As Long
    
On Error GoTo Erro_objEventoDocAuto_evSelecao
    
    Set objDocAuto = obj1
    
    lErro = Traz_DocAuto_Tela(objDocAuto.lDoc)
    If lErro <> SUCESSO Then Error 11513
    
    iAlterado = 0
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoDocAuto_evSelecao:

    Select Case Err
    
        Case 11513
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156016)
            
    End Select
        
    Exit Sub

End Sub

Public Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sCcl As String
Dim sCclEnxuta As String
Dim lErro As Long
Dim lPosicaoSeparador As Long
Dim sCaracterInicial As String
    
On Error GoTo Erro_TvwCcls_NodeClick
    
    If GridLancamentos.Col = iGrid_Ccl_Col Then
    
        sCaracterInicial = left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sCcl = right(Node.Key, Len(Node.Key) - 1)
              
            sCclEnxuta = String(STRING_CCL, 0)
            
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 10503
            
            Ccl.PromptInclude = False
            Ccl.Text = sCclEnxuta
            Ccl.PromptInclude = True
              
            GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
                objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
            End If
        
            'Preenche a Descricao do centro de custo/lucro
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            CclDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    
        End If
    
    End If
    
    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err
    
        Case 10503
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156017)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub ListDocAuto_DblClick()

Dim lDoc As Long
Dim lErro As Long
Dim objDocAuto As New ClassDocAuto
Dim lPosicaoSeparador As Long

On Error GoTo Erro_ListDocAuto_DblClick
    
    'Guarda a posicao em que o separador se encontra
    lPosicaoSeparador = InStr(ListDocAuto.Text, SEPARADOR)
    
    'Pega o Numero do Documento Automático selecionado
    lDoc = CLng(left(ListDocAuto.Text, lPosicaoSeparador - 1))
    
    objDocAuto.lDoc = lDoc

    lErro = Traz_DocAuto_Tela(lDoc)
    If lErro <> SUCESSO Then Error 11516
    
    Exit Sub
    
Erro_ListDocAuto_DblClick:

    Select Case Err
    
        Case 11516
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156018)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub ListHistoricos_DblClick()

Dim lPosicaoSeparador As Long
    
    If GridLancamentos.Col = iGrid_Historico_Col Then
    
        lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
        Historico.Text = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
    
        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
            objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
        End If
    
    End If
    
End Sub

Public Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
        If lErro <> SUCESSO Then Error 44045
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 44045
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156019)
        
    End Select
        
    Exit Sub

End Sub

Public Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim lErro As Long
Dim iIndice As Integer
Dim lPosicaoSeparador As Long
Dim sConta As String
Dim sCaracterInicial As String
Dim sContaEnxuta As String
Dim objPlanoConta As New ClassPlanoConta
Dim sContaFormatada As String

On Error GoTo Erro_TvwContas_NodeClick
    
    If GridLancamentos.Col = iGrid_Conta_Col Then
    
        sCaracterInicial = left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sConta = right(Node.Key, Len(Node.Key) - 1)
            
            sContaEnxuta = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 11517
            
            Conta.PromptInclude = False
            Conta.Text = sContaEnxuta
            Conta.PromptInclude = True
        
            GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Conta.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
                objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
            End If
        
            'Preenche a Descricao da Conta
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            ContaDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
        
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 19135
                    
            'Conta não cadastrada
            If lErro = 5700 Then Error 11936
        
            'Se a Conta possui um Histórico Padrão "default" coloca na tela
            If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                            
                For iIndice = 0 To ListHistoricos.ListCount - 1
                    If ListHistoricos.ItemData(iIndice) = objPlanoConta.iHistPadrao Then
                        ListHistoricos.ListIndex = iIndice
                        lPosicaoSeparador = InStr(ListHistoricos.Text, SEPARADOR)
                        GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = Mid(ListHistoricos.Text, lPosicaoSeparador + 1)
                        Exit For
                    End If
                Next
        
            End If
        
        End If
        
    End If
        
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 11517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case 19135
                
        Case 11936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, Conta.Text)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156020)
            
    End Select
        
    Exit Sub
    
End Sub

Function Conta_Exibe_Descricao(sConta As String) As Long
'exibe a descrição da conta no campo ContaDescricao. A conta passada como parametro deve estar mascarada

Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Conta_Exibe_Descricao

    'Retorna conta formatada como no BD
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 11518
    
    lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 6030 Then Error 11519

    If lErro = 6030 Then Error 11520
    
    ContaDescricao.Caption = objPlanoConta.sDescConta
    
    Conta_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Conta_Exibe_Descricao:

    Conta_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11518, 11519
            ContaDescricao = ""
            
        Case 11520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            ContaDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156021)
            
    End Select
        
    Exit Function

End Function

Function Ccl_Exibe_Descricao(sCcl As String) As Long
'exibe a descrição do centro de custo/lucro no campo CclDescricao. O ccl passado como parametro deve estar mascarado

Dim sCclFormatada As String
Dim sCclArvore As String
Dim objNode As Node
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Exibe_Descricao

    'Retorna Ccl formatada como no BD
    lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then Error 11521
    
    objCcl.sCcl = sCclFormatada

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 11522
    
    If lErro = 5599 Then Error 11523
    
    CclDescricao.Caption = objCcl.sDescCcl
    
    Ccl_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Ccl_Exibe_Descricao:

    Ccl_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11521, 11522
            CclDescricao = ""
            
        Case 11523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
            CclDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156022)
            
    End Select
        
    Exit Function

End Function

Public Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text
        
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 11524
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_DownClick:
    
    Select Case Err
    
        Case 11524
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156023)
        
    End Select
    
    Exit Sub

End Sub

Public Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 11525
        
        Data.Text = sData
        
    End If
    
    Exit Sub
    
Erro_UpDown1_UpClick:
    
    Select Case Err
    
        Case 11525
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156024)
        
    End Select
    
    Exit Sub

End Sub

Private Function Traz_Rateio_Tela(objRateioOn1 As ClassRateioOn) As Long
'Traz os dados de um rateio, previamente escolhido, para a tela

Dim ColRateioOn As New Collection
Dim objRateioOn As New ClassRateioOn
Dim dValorFinal As Double
Dim dTotalDebito As Double
Dim dTotalCredito As Double
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim iUltimaLinha As Integer
Dim lErro As Long
   
On Error GoTo Erro_Traz_Rateio_Tela

    objRateioOn.iCodigo = objRateioOn1.iCodigo
    
    lErro = CF("RateioOn_Le_Doc", objRateioOn, ColRateioOn)
    If lErro <> SUCESSO And lErro <> 11136 Then Error 11526
    
    'se não encontrou o Rateio
    If lErro = 11136 Then Error 11527
        
    'Variarel que vai indicar apartir de onde deve ser inserido os lancamentos vindos na collection
    iUltimaLinha = objGrid1.ilinhasExistentes
        
    For Each objRateioOn In ColRateioOn

        If Len(objRateioOn.sConta) > 0 Then
        
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objRateioOn.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 11528
            
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
            
            'coloca a conta na tela
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Conta_Col) = Conta.Text
    
        Else
        
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Conta_Col) = ""
            
            
        End If
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(objRateioOn.sCcl) > 0 Then
        
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
            
                lErro = Mascara_RetornaCclEnxuta(objRateioOn.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 11529
            
                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True
            
                'coloca o centro de custo na tela
                GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Ccl_Col) = Ccl.Text
        
            Else
            
                GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Ccl_Col) = ""
            
            End If
            
        End If
        
        dValorFinal = objRateioOn1.dPercentual * objRateioOn.dPercentual
        
        'O campo dPercentual do ojbRateioon em "Traz_Rateio_Tela" já traz o valor final do rateio
        If dValorFinal > 0 Then
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Credito_Col) = Format(dValorFinal, "Standard")
        Else
            GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Debito_Col) = Format(-dValorFinal, "Standard")
        End If
            
        'coloca o historico na tela
        GridLancamentos.TextMatrix(objRateioOn.iSeq + iUltimaLinha, iGrid_Historico_Col) = objRateioOn.sHistorico
            
        objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
            
    Next

    dTotalCredito = GridColuna_Soma(iGrid_Credito_Col)
    dTotalDebito = GridColuna_Soma(iGrid_Debito_Col)
    
    TotalCredito.Caption = Format(dTotalCredito, "Standard")
    TotalDebito.Caption = Format(dTotalDebito, "Standard")
    
    Traz_Rateio_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Rateio_Tela:

    Traz_Rateio_Tela = Err
        
    Select Case Err
    
        Case 11526
        
        Case 11527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOON_NAO_CADASTRADO", Err, objRateioOn1.iCodigo)
    
        Case 11528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objRateioOn.sConta)
        
        Case 11529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, objRateioOn.sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156025)
        
    End Select
    
    Exit Function
    
End Function

Private Function Traz_DocAuto_Tela(lDoc As Long) As Long
'Coloca no GridLancamentos os Lancamentos do Documento Automático lDoc
    
Dim colDocAuto As New Collection
Dim objDocAuto As New ClassDocAuto
Dim objDocAuto1 As ClassDocAuto
Dim dTotalCredito As Double
Dim dTotalDebito As Double
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim sDescricao As String
Dim lErro As Long
Dim iLinha As Integer
   
On Error GoTo Erro_Traz_DocAuto_Tela
    
    Set objDocAuto = New ClassDocAuto
    
    objDocAuto.lDoc = lDoc
    
    'Le os dados do Documento Automatico passado como parâmetro
    lErro = CF("DocAuto_Le_Doc", objDocAuto, colDocAuto)
    If lErro <> SUCESSO And lErro <> 11017 Then Error 11534
    
    'se não encontrou o documento
    If lErro = 11017 Then Error 11535
    
    Call Limpa_Tela_Lancamentos

    'Transfere os dados para a Tela
    DocAuto.Text = CStr(lDoc)
    
    If colDocAuto.Count >= objGrid1.objGrid.Rows Then
        Call Refaz_Grid(objGrid1, colDocAuto.Count)
    End If
    
    For Each objDocAuto1 In colDocAuto

        iLinha = iLinha + 1

        If Len(objDocAuto1.sConta) > 0 Then

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objDocAuto1.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 11536
            
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
        
            'coloca a conta na tela
            GridLancamentos.TextMatrix(iLinha, iGrid_Conta_Col) = Conta.Text
        
        Else
        
            GridLancamentos.TextMatrix(iLinha, iGrid_Conta_Col) = ""
        
        End If
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(objDocAuto1.sCcl) > 0 Then
        
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
            
                lErro = Mascara_RetornaCclEnxuta(objDocAuto1.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 11537
            
                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True
            
                'coloca o centro de custo na tela
                GridLancamentos.TextMatrix(iLinha, iGrid_Ccl_Col) = Ccl.Text
            
            Else
            
                GridLancamentos.TextMatrix(iLinha, iGrid_Ccl_Col) = ""
                
            End If
            
        End If
        
        'coloca o valor na tela
        If objDocAuto1.dValor > 0 Then
            GridLancamentos.TextMatrix(iLinha, iGrid_Credito_Col) = Format(objDocAuto1.dValor, "Standard")
            dTotalCredito = dTotalCredito + objDocAuto1.dValor
        Else
            GridLancamentos.TextMatrix(iLinha, iGrid_Debito_Col) = Format(-objDocAuto1.dValor, "Standard")
            dTotalDebito = dTotalDebito - objDocAuto1.dValor
        End If
            
        'coloca o sequencial de contra-partida na tela
        If objDocAuto1.iSeqContraPartida <> 0 Then GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(objDocAuto1.iSeqContraPartida)
            
        'coloca o historico na tela
        GridLancamentos.TextMatrix(iLinha, iGrid_Historico_Col) = objDocAuto1.sHistorico
            
        objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
    
    Next

    TotalCredito.Caption = Format(dTotalCredito, "Standard")
    TotalDebito.Caption = Format(dTotalDebito, "Standard")
    
    Traz_DocAuto_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_DocAuto_Tela:

    Traz_DocAuto_Tela = Err
        
    Select Case Err
    
        Case 11534, 11535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCAUTO_NAO_CADASTRADO", Err)
            
        Case 11536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objDocAuto1.sConta)
        
        Case 11537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, objDocAuto1.sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156026)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Trata_SeqContraPartida(iLinhaExcluida As Integer)
'altera os indicadores de contra partida quando uma linha é excluida

Dim iLinha As Integer

    For iLinha = 1 To objGrid1.ilinhasExistentes
    
        If Len(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > 0 Then
    
            If CInt(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) = iLinhaExcluida Then
                GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = ""
            ElseIf CInt(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > iLinhaExcluida Then
                GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(CInt(GridLancamentos.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) - 1)
                
            End If
        End If
    
    Next

End Sub

Private Function Armazena_Contra_Partida(colContraPartida As Collection, objLancamento_Detalhe As ClassLancamento_Detalhe) As Long
'armazena os totais de contra partida para posteriormente checar se o total de contra partida bate com o lançamento oposto

Dim objContraPartida As ClassContraPartida
Dim iAchou As Integer
Dim lErro As Long

On Error GoTo Erro_Armazena_Contra_Partida

    For Each objContraPartida In colContraPartida
    
        If objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeqContraPartida Then
        
            objContraPartida.dValorContraPartida = objContraPartida.dValorContraPartida - objLancamento_Detalhe.dValor
            iAchou = 1
            Exit For
            
        End If
        
    Next
    
    If iAchou = 0 Then
            
        Set objContraPartida = New ClassContraPartida
        
        objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeqContraPartida
        objContraPartida.dValorContraPartida = -objLancamento_Detalhe.dValor
                
        colContraPartida.Add objContraPartida
    
    End If

    Armazena_Contra_Partida = SUCESSO
    
    Exit Function

Erro_Armazena_Contra_Partida:

    Armazena_Contra_Partida = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156027)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_Contra_Partida(colLancamento_Detalhe As Collection, colContraPartida As Collection) As Long
'checa se o total de contra partida bate com o lançamento correspondente

Dim objContraPartida As ClassContraPartida
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iAchou As Integer
Dim lErro As Long

On Error GoTo Erro_Testa_Contra_Partida

    For Each objContraPartida In colContraPartida
    
        For Each objLancamento_Detalhe In colLancamento_Detalhe
        
            If objContraPartida.iSeqContraPartida = objLancamento_Detalhe.iSeq Then
            
                iAchou = 1
                If objContraPartida.dValorContraPartida <> objLancamento_Detalhe.dValor Then Error 20616
                Exit For
                
            End If
            
        Next
        
        'se não achou o lancamento oposto da contra-partida
        If iAchou = 0 Then Error 20617
        
        iAchou = 0
        
    Next
        
    Testa_Contra_Partida = SUCESSO
    
    Exit Function

Erro_Testa_Contra_Partida:

    Testa_Contra_Partida = Err
    
    Select Case Err
    
        Case 20616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_CONTRA_PARTIDA_VALOR", Err, objLancamento_Detalhe.iSeq, Abs(objLancamento_Detalhe.dValor), Abs(objContraPartida.dValorContraPartida))
        
        Case 20617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_CONTRA_PARTIDA_INEXISTENTE", Err, objContraPartida.iSeqContraPartida)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156028)

    End Select
    
    Exit Function

End Function

Public Sub BotaoImprimir_Click()
'imprime um relatorio com os dados que estao na tela

Dim lErro As Long, objRelTela As New ClassRelTela, iIndice1 As Integer
Dim colTemp As Collection, objLancamento As New ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim sOrigem As String, sExercicio As String, sPeriodo As String
Dim lDoc As Long, dtData As Date
Dim sNomeTsk As String
Dim iIndice As Integer

On Error GoTo Erro_BotaoImprimir_Click
    
    lErro = objRelTela.Iniciar("REL_LANCAT_CTB")
    If lErro <> SUCESSO Then gError 47894
    
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        sNomeTsk = "LanAtCCL"
    Else
        sNomeTsk = "LancAt"
    End If

    'obter dados comuns a todas as linhas do grid
    sOrigem = gobjColOrigem.Origem(Origem.Caption)
    sExercicio = Exercicio.Caption
    sPeriodo = Periodo.Caption
    
    lDoc = StrParaLong(Documento.ClipText)
    dtData = StrParaDate(Data.Text)
    
    lErro = Grid_Lancamento_Detalhe(colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 47895
    
    For iIndice1 = 1 To colLancamento_Detalhe.Count
    
        Set objLancamento = colLancamento_Detalhe.Item(iIndice1)
        
        Set colTemp = New Collection
        
        'incluir os valores na mesma ordem da tabela RelTelaCampos no dicdados
        
        Call colTemp.Add(sOrigem)
        Call colTemp.Add(sExercicio)
        Call colTemp.Add(sPeriodo)
        Call colTemp.Add(lDoc)
        Call colTemp.Add(iIndice1)
        Call colTemp.Add(dtData)
        Call colTemp.Add(objLancamento.sConta)
        Call colTemp.Add(objLancamento.sCcl)
        Call colTemp.Add(objLancamento.sHistorico)
        Call colTemp.Add(objLancamento.dValor)
        Call colTemp.Add(objLancamento.iSeqContraPartida)

        lErro = objRelTela.IncluirRegistro(colTemp)
        If lErro <> SUCESSO Then gError 47896
    
    Next
    
    lErro = objRelTela.ExecutarRel(sNomeTsk)
    If lErro <> SUCESSO Then gError 47897
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
          
        Case 47894, 47895, 47896, 47897
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156029)
     
    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LANCAMENTO_AT
    Set Form_Load_Ocx = Me
    Caption = "Lançamentos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "LancamentosAt"
    
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
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is DocAuto Then
            Call Label6_Click
        ElseIf Me.ActiveControl Is DocumentoLabel Then
            Call DocumentoLabel_Click
        ElseIf Me.ActiveControl Is Conta Then
            Call BotaoConta_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcl_Click
        ElseIf Me.ActiveControl Is Historico Then
            Call BotaoHist_Click
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

Private Function Move_Tela_Memoria(objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long

Dim dtData As Date
Dim objPeriodo As New ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Obtém Período e Exercício correspondentes à data
    dtData = CDate(Data.Text)

    'Lê o Período
    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 61130

    'Preenche objLote
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    objLancamento_Cabecalho.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Cabecalho.lDoc = CLng(Documento.Text)
    objLancamento_Cabecalho.dtData = dtData

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 61130 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156030)

    End Select

    Exit Function

End Function

Public Sub BotaoEstornar_Click()

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim objBrowseConfigura As New AdmBrowseConfigura
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoEstornar_Click

    'Se Documento estiver vazio, Erro
    If Len(Trim(Documento.Text)) = 0 Then gError 61131

    'Se Data estiver vazia, Erro
    If Len(Trim(Data.Text)) = 0 Then gError 61132

    lErro = Move_Tela_Memoria(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then gError 61133

    'Lê algum lançamento contido no documento em questão
    lErro = CF("Lancamento_Le_Doc1", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 83863 And lErro <> 83864 And lErro <> 83865 Then gError 61134

    If lErro = 83863 Then gError 61135

    If lErro = 83864 Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ESTORNO_DOCUMENTO_ESTORNADO")
        If vbMsgRes <> vbYes Then gError 83869
    
    
    End If

    If lErro = 83865 Then

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ESTORNO_DOCUMENTO_ESTORNADOR")
        If vbMsgRes <> vbYes Then gError 83870

    End If

    'chama a tela LancamentoEstorno1 Modal
    lErro = Chama_Tela_Modal("LancamentoEstorno1", objLancamento_Cabecalho, objBrowseConfigura)
    If lErro <> SUCESSO Then gError 61136

    If objBrowseConfigura.iTelaOK = CANCELA Then gError 61137

    Call Limpa_Tela_Lancamentos

    Exit Sub

Erro_BotaoEstornar_Click:

    Select Case gErr

        Case 61131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Documento.SetFocus

        Case 61132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", gErr)
            Data.SetFocus

        Case 61133, 61134, 61136, 83869, 83870 'Tratado na rotina chamada

        Case 61135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", gErr, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case 61137
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_ESTORNO_LANCAMENTO_CANCELADO", objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156031)

    End Select

    Exit Sub

End Sub

Public Sub DocumentoLabel_Click()

Dim lErro As Long
Dim dtData As Date
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim objPeriodo As New ClassPeriodo
Dim colSelecao As New Collection

On Error GoTo Erro_DocumentoLabel_Click

    If Len(Data.ClipText) > 0 Then

        'Obtém Periodo e Exercicio correspondentes à data
        dtData = CDate(Data.Text)

        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then Error 61138

    Else

        objPeriodo.iExercicio = 0
        objPeriodo.iPeriodo = 0

    End If

    If Len(Documento.Text) = 0 Then
        objLancamento_Detalhe.lDoc = 0
    Else
        objLancamento_Detalhe.lDoc = CLng(Documento.ClipText)
    End If

    objLancamento_Detalhe.iFilialEmpresa = giFilialEmpresa
    objLancamento_Detalhe.sOrigem = gobjColOrigem.Origem(Origem.Caption)
    objLancamento_Detalhe.iExercicio = objPeriodo.iExercicio
    objLancamento_Detalhe.iPeriodoLan = objPeriodo.iPeriodo

    Call Chama_Tela("LancamentoLista", colSelecao, objLancamento_Detalhe, objEventoLancamento)

    Exit Sub

Erro_DocumentoLabel_Click:

    Select Case Err

        Case 61138

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156032)

    End Select

    Exit Sub

End Sub


Private Sub objEventoLancamento_evSelecao(obj1 As Object)
'Traz o lançamento selecionado para a tela

Dim lErro As Long
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_objEventoLancamento_evSelecao

    Set objLancamento_Detalhe = obj1

    objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Detalhe.iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = objLancamento_Detalhe.sOrigem
    objLancamento_Cabecalho.iExercicio = objLancamento_Detalhe.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objLancamento_Detalhe.iPeriodoLan
    objLancamento_Cabecalho.lDoc = objLancamento_Detalhe.lDoc

    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 61140

    Me.Show

    Exit Sub

Erro_objEventoLancamento_evSelecao:

    Select Case Err

        Case 61140 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156033)

    End Select

    Exit Sub

End Sub

Private Function Traz_Doc_Tela(objDoc As ClassLancamento_Cabecalho) As Long
'Traz os dados do documento para a tela

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim sDescricao As String
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim dColunaSoma As Double
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim iFilialEmpresaSalva As Integer
Dim iAchou As Integer
Dim iIndice1 As Integer

On Error GoTo Erro_Traz_Doc_Tela

    Call Limpa_Tela_Lancamentos

    iFilialEmpresaSalva = objDoc.iFilialEmpresa

    Do While objDoc.iFilialEmpresa > 0 And objDoc.iFilialEmpresa < 100

        'Lê os lançamentos contidos no documento em questão
        lErro = CF("Lancamentos_Le_Doc", objDoc, colLancamentos)
        If lErro <> SUCESSO And lErro <> 28700 Then Error 61141

        If lErro = SUCESSO Then iAchou = 1

        If giContabGerencial = 0 Then Exit Do

        objDoc.iFilialEmpresa = objDoc.iFilialEmpresa - giFilialAuxiliar

    Loop

    objDoc.iFilialEmpresa = iFilialEmpresaSalva

    'Se não encontrou o documento
    If iAchou = 0 Then Error 61142

    For iIndice = colLancamentos.Count To 1 Step -1
        
        For iIndice1 = iIndice - 1 To 1 Step -1
        
            If colLancamentos(iIndice).iSeq = colLancamentos(iIndice1).iSeq Then
                colLancamentos.Remove (iIndice)
                Exit For
            End If
            
        Next
    
    Next



    Documento.Text = CStr(objDoc.lDoc)

    Set objLanc = colLancamentos.Item(1)

    'Inicializa Data
    Data.Text = Format(objLanc.dtData, "dd/mm/yy")

    'Coloca o período relativo a data na tela
    lErro = CF("Periodo_Le", objLanc.dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 61143

    Periodo.Caption = objPeriodo.sNomeExterno

    'Coloca o exercício na tela
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then Error 61144

    'Se o exercício não está cadastrado
    If lErro = 10083 Then Error 61145

    Exercicio.Caption = objExercicio.sNomeExterno

    Origem.Caption = gobjColOrigem.Descricao(objDoc.sOrigem)
    
    If colLancamentos.Count > MAX_LANCAMENTOS_POR_DOC_CTB + 1 Then gError 197922
    
    If colLancamentos.Count >= objGrid1.objGrid.Rows Then
        Call Refaz_Grid(objGrid1, colLancamentos.Count)
    End If
    
    'Move os dados para a tela
    For Each objLanc In colLancamentos

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objLanc.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 61146

        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True

        'Coloca a conta na tela
        GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Conta_Col) = Conta.Text

        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then

            If objLanc.sCcl <> "" Then

                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)

                lErro = Mascara_RetornaCclEnxuta(objLanc.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 61147

                Ccl.PromptInclude = False
                Ccl.Text = sCclMascarado
                Ccl.PromptInclude = True

                'Coloca o centro de custo na tela
                GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Ccl_Col) = Ccl.Text

            Else
            
                'Coloca o centro de custo na tela
                GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Ccl_Col) = ""

            End If

        End If

        'Coloca o valor na tela
        If objLanc.dValor > 0 Then
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Credito_Col) = Format(objLanc.dValor, "Standard")
        Else
            GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Debito_Col) = Format(-objLanc.dValor, "Standard")
        End If

        'Coloca o histórico na tela
        GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Historico_Col) = objLanc.sHistorico

        If objLanc.iSeqContraPartida > 0 Then GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_SeqContraPartida_Col) = objLanc.iSeqContraPartida

        If giContabGerencial = 1 Then GridLancamentos.TextMatrix(objLanc.iSeq, iGrid_Gerencial_Col) = CStr(objLanc.iGerencial)

        objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1

    Next

    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito.Caption = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito.Caption = Format(dColunaSoma, "Standard")

    Call Grid_Refresh_Checkbox(objGrid1)

    Traz_Doc_Tela = SUCESSO

    Exit Function

Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = Err

    Select Case Err

        Case 61141, 61143, 61144 'Tratado nas rotinas chamadas

        Case 61142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_NAO_CADASTRADO", Err, objDoc.sOrigem, objDoc.iExercicio, objDoc.iPeriodoLan, objDoc.lDoc)

        Case 61145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objPeriodo.iExercicio)

        Case 61146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objLanc.sConta)

        Case 61147
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objLanc.sCcl)
            
        Case 197922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_LANC_MAIOR_LIMITE", gErr, colLancamentos.Count, MAX_LANCAMENTOS_POR_DOC_CTB)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156034)

    End Select

    Exit Function

End Function
 
Public Sub BotaoConsultar_Click()

Dim lErro As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho

On Error GoTo Erro_BotaoConsultar_Click

    'Se Origem estiver vazio, Erro
    If Len(Trim(Origem.Caption)) = 0 Then Error 61148

    'Se Documento estiver vazio, Erro
    If Len(Documento.ClipText) = 0 Then Error 61149

    'Se Data estiver vazio ==> Erro
    If Len(Data.ClipText) = 0 Then Error 61150

    lErro = Move_Tela_Memoria(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 61151

    lErro = Traz_Doc_Tela(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 61152

    Exit Sub

Erro_BotaoConsultar_Click:

    Select Case Err

        Case 61148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA", Err)
            Origem.SetFocus

        Case 61149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
            Documento.SetFocus

        Case 61150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA", Err)
            Data.SetFocus

        Case 61151, 61152 'Tratado nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156035)

    End Select

    Exit Sub

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

Private Function Saida_Celula_Gerencial(objGridInt As AdmGrid) As Long
'faz a critica da celula Gerencial do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Gerencial

    Set objGridInt.objControle = Gerencial

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188071

    Saida_Celula_Gerencial = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Gerencial:

    Saida_Celula_Gerencial = gErr
    
    Select Case gErr
    
        Case 188071
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188063)
        
    End Select

    Exit Function

End Function

Public Sub BotaoConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Conta.Text) > 0 Then objPlanoConta.sConta = Conta.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_objEventoConta_evSelecao
    
    If GridLancamentos.Col = iGrid_Conta_Col Then

        Set objPlanoConta = obj1
        
        sConta = objPlanoConta.sConta
        
        'le a conta
        lErro = CF("PlanoConta_Le_Conta1", sConta, objPlanoConta)
        If lErro <> SUCESSO And lErro <> 6030 Then gError 197910
        
        If objPlanoConta.iAtivo <> CONTA_ATIVA Then gError 197911
        
        If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then gError 197912
        
        sContaEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 197913

        Conta.PromptInclude = False
        Conta.Text = sContaEnxuta
        Conta.PromptInclude = True

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Conta.Text

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
            objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
        End If

        ContaDescricao.Caption = objPlanoConta.sDescConta
        
        'Se a Conta possui um Histórico Padrão "default" coloca na tela
        If Len(Trim(GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col))) = 0 And objPlanoConta.iHistPadrao <> 0 Then
                        
            objHistPadrao.iHistPadrao = objPlanoConta.iHistPadrao
                        
            'le os dados do historico
            lErro = CF("HistPadrao_Le", objHistPadrao)
            If lErro <> SUCESSO And lErro <> 5446 Then gError 197914
                                    
            If lErro = SUCESSO Then
            
                GridLancamentos.TextMatrix(GridLancamentos.Row, iGrid_Historico_Col) = objHistPadrao.sDescHistPadrao
                
            End If
                        
        End If

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr
    
        Case 197910, 197914
    
        Case 197911
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", gErr, sConta)
        
        Case 197912
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", gErr, sConta)
    
        Case 197913
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197915)
        
    End Select

    Exit Sub

End Sub

Public Sub BotaoCcl_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Ccl.Text) > 0 Then objCcl.sCcl = Ccl.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objCcl As ClassCcl
Dim sConta As String
Dim sCclEnxuta As String

On Error GoTo Erro_objEventoCcl_evSelecao
    
    If GridLancamentos.Col = iGrid_Ccl_Col Then

        Set objCcl = obj1

        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then gError 197916

        If objCcl.iTipoCcl <> CCL_ANALITICA Then gError 197917
        
        If objCcl.iAtivo = 0 Then gError 197918
        
        sCclEnxuta = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclEnxuta)
        If lErro <> SUCESSO Then gError 197919

        Ccl.PromptInclude = False
        Ccl.Text = sCclEnxuta
        Ccl.PromptInclude = True

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = Ccl.Text

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
            objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
        End If

        CclDescricao.Caption = objCcl.sDescCcl

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoCcl_evSelecao:

    Select Case gErr
    
        Case 197916

        Case 197917
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", gErr, objCcl.sCcl)
  
        Case 197918
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_INATIVO", gErr, objCcl.sCcl)

        Case 197919
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197920)
        
    End Select

    Exit Sub

End Sub

Public Sub BotaoHist_Click()

Dim colSelecao As Collection
Dim objHistPadrao As New ClassHistPadrao

    Call Chama_Tela("HistPadraoLista", colSelecao, objHistPadrao, objEventoHist)

End Sub

Private Sub objEventoHist_evSelecao(obj1 As Object)


Dim objHistPadrao As ClassHistPadrao

On Error GoTo Erro_objEventoHist_evSelecao

    If GridLancamentos.Col = iGrid_Historico_Col Then

        Set objHistPadrao = obj1

        GridLancamentos.TextMatrix(GridLancamentos.Row, GridLancamentos.Col) = objHistPadrao.sDescHistPadrao
        Historico.Text = objHistPadrao.sDescHistPadrao

        If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
            objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
        End If

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoHist_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197921)

    End Select

    Exit Sub

End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 10

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub


