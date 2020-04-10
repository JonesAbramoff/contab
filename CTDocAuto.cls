VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTDocAuto"
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
Dim iGrid_Gerencial_Col

Dim objGrid1 As AdmGrid
Dim iAlterado As Integer
Private WithEvents objEventoDocAuto As AdmEvento
Attribute objEventoDocAuto.VB_VarHelpID = -1

'--- inicio dos properties get dos controles da tela

Public Property Get Gerencial() As Object
     Set Gerencial = objUserControl.Controls("Gerencial")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
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

Public Property Get ListaHistorico() As Object
     Set ListaHistorico = objUserControl.Controls("ListaHistorico")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
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

Public Property Get Documento() As Object
     Set Documento = objUserControl.Controls("Documento")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get GridDocAuto() As Object
     Set GridDocAuto = objUserControl.Controls("GridDocAuto")
End Property

Public Property Get TvwCcls() As Object
     Set TvwCcls = objUserControl.Controls("TvwCcls")
End Property

Public Property Get TvwContas() As Object
     Set TvwContas = objUserControl.Controls("TvwContas")
End Property

Public Property Get LabelHistorico() As Object
     Set LabelHistorico = objUserControl.Controls("LabelHistorico")
End Property

Public Property Get LabelCCL() As Object
     Set LabelCCL = objUserControl.Controls("LabelCCL")
End Property

Public Property Get LabelPlanoConta() As Object
     Set LabelPlanoConta = objUserControl.Controls("LabelPlanoConta")
End Property

Public Property Get Line1() As Object
     Set Line1 = objUserControl.Controls("Line1")
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

'--- fim dos properties get dos controles da tela

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo voucher(documento) disponível
    lErro = CF("DocAuto_Automatico", lDoc)
    If lErro <> SUCESSO Then Error 57513

    Documento.Text = CStr(lDoc)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57513
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155668)
    
    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_Form_Load

    Set objEventoDocAuto = New AdmEvento
    
    Set objGrid1 = New AdmGrid
    
    'tela em questão
    Set objGrid1.objForm = Me
    
    lErro = Inicializa_Grid_DocAuto(objGrid1)
    If lErro <> SUCESSO Then Error 11000

    'Inicializa a Lista de Plano de Contas
    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
    If lErro <> SUCESSO Then Error 11001

    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
        'Inicializa a Lista de Centros de Custo
        lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
        If lErro <> SUCESSO Then Error 11002
        
    End If
    
    'Inicializa a Lista de Historicos
    lErro = Carga_Lista_Historico()
    If lErro <> SUCESSO Then Error 11003
    
    TvwContas.Visible = True
    LabelPlanoConta.Visible = True
    TvwCcls.Visible = False
    LabelCCL.Visible = False
    ListaHistorico.Visible = False
    LabelHistorico.Visible = False
    
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
    
        Case 11000, 11001, 11002, 11003
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155669)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_DocAuto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_DocAuto
    
    'titulos do grid
    objGridInt.colColuna.Add (" ")
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
    If lErro <> SUCESSO Then Error 11004
    
    objGridInt.objGrid = GridDocAuto
    
    'todas as linhas do grid
    objGridInt.objGrid.Rows = MAX_LANCAMENTOS_POR_DOC + 1
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
        
    GridDocAuto.ColWidth(0) = 400
    
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalDebito.top = GridDocAuto.top + GridDocAuto.Height
    TotalDebito.left = GridDocAuto.left
    For iIndice = 0 To iGrid_Debito_Col - 1
        TotalDebito.left = TotalDebito.left + GridDocAuto.ColWidth(iIndice) + GridDocAuto.GridLineWidth + 20
    Next
    
    TotalDebito.Width = GridDocAuto.ColWidth(iGrid_Debito_Col)
    
    TotalCredito.top = TotalDebito.top
    TotalCredito.left = TotalDebito.left + TotalDebito.Width + GridDocAuto.GridLineWidth
    TotalCredito.Width = GridDocAuto.ColWidth(iGrid_Credito_Col)
    
    LabelTotais.top = TotalDebito.top + (TotalDebito.Height - LabelTotais.Height) / 2
    LabelTotais.left = TotalDebito.left - LabelTotais.Width

    Inicializa_Grid_DocAuto = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_DocAuto:

    Inicializa_Grid_DocAuto = Err
    
    Select Case Err
    
        Case 11004
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155670)
        
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
    If lErro <> SUCESSO Then Error 11005
    
    Conta.Mask = sMascaraConta
    
    'Se usa centro de custo/lucro ==> inicializa mascara de centro de custo/lucro
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
    
        sMascaraCcl = String(STRING_CCL, 0)

        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then Error 11006

        Ccl.Mask = sMascaraCcl
        
    End If
    
    Inicializa_Mascaras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascaras:

    Inicializa_Mascaras = Err
    
    Select Case Err
    
        Case 11005, 11006
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155671)
        
    End Select

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
    If lErro <> SUCESSO Then Error 11010
    
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 11011

        If objCcl.iTipoCcl = CCL_ANALITICA Then
            sCcl = "A" & objCcl.sCcl
        Else
            sCcl = "S" & objCcl.sCcl
        End If

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" do centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 10489
        
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

        Case 10489
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case 11010

        Case 11011
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155672)

    End Select
    
    Exit Function

End Function

Private Function Carga_Lista_Historico() As Long
'move os dados de historico padrão do banco de dados para a arvore colNodes.

Dim objNode As Node
Dim colHistPadrao As New Collection
Dim objHistPadrao As ClassHistPadrao
Dim lErro As Long
    
On Error GoTo Erro_Carga_Lista_Historico
    
    lErro = CF("HistPadrao_Le_Todos", colHistPadrao)
    If lErro <> SUCESSO Then Error 11012
    
    For Each objHistPadrao In colHistPadrao
        
        ListaHistorico.AddItem CStr(objHistPadrao.iHistPadrao) & SEPARADOR & objHistPadrao.sDescHistPadrao
        
    Next
    
    Carga_Lista_Historico = SUCESSO

    Exit Function

Erro_Carga_Lista_Historico:

    Carga_Lista_Historico = Err

    Select Case Err

        Case 11012

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155673)

    End Select
    
    Exit Function

End Function

Private Function Traz_Doc_Tela(objDocAuto As ClassDocAuto) As Long
'traz os dados do voucher do banco de dados para a tela

Dim lErro As Long
Dim colDocAutos As New Collection
Dim objDocAuto1 As ClassDocAuto
Dim iLinha As Integer
Dim sContaMascarada As String
Dim sCclMascarado As String
Dim dTotalCredito As Double
Dim dTotalDebito As Double
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_Traz_Doc_Tela

    Call Limpa_Tela_DocAuto
    
    lErro = CF("DocAuto_Le_Doc", objDocAuto, colDocAutos)
    If lErro <> SUCESSO And lErro <> 11017 Then Error 11016
    
    'move os dados para a tela
    Documento.Text = CStr(objDocAuto.lDoc)
    
    If lErro = SUCESSO Then
    
        dTotalCredito = 0
        dTotalDebito = 0
    
        For Each objDocAuto1 In colDocAutos
        
            iLinha = iLinha + 1
        
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objDocAuto1.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 11014
            
            Conta.PromptInclude = False
            Conta.Text = sContaMascarada
            Conta.PromptInclude = True
            
            'coloca a conta na tela
            GridDocAuto.TextMatrix(iLinha, iGrid_Conta_Col) = Conta.Text
            
            If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
            
                'mascara o centro de custo
                sCclMascarado = String(STRING_CCL, 0)
                           
                If objDocAuto1.sCcl <> "" Then
                
                    lErro = Mascara_RetornaCclEnxuta(objDocAuto1.sCcl, sCclMascarado)
                    If lErro <> SUCESSO Then Error 11015
    
                    Ccl.PromptInclude = False
                    Ccl.Text = sCclMascarado
                    Ccl.PromptInclude = True
            
                    'coloca o centro de custo na tela
                    GridDocAuto.TextMatrix(iLinha, iGrid_Ccl_Col) = Ccl.Text
                
            
                Else
                    'coloca o centro de custo na tela
                    GridDocAuto.TextMatrix(iLinha, iGrid_Ccl_Col) = ""
                
                   
                End If
                
            End If
            
            'coloca o valor na tela
            If objDocAuto1.dValor > 0 Then
                GridDocAuto.TextMatrix(iLinha, iGrid_Credito_Col) = Format(objDocAuto1.dValor, "Standard")
                dTotalCredito = dTotalCredito + objDocAuto1.dValor
            Else
                If objDocAuto1.dValor < 0 Then
                    GridDocAuto.TextMatrix(iLinha, iGrid_Debito_Col) = Format(-objDocAuto1.dValor, "Standard")
                    dTotalDebito = dTotalDebito + objDocAuto1.dValor
                End If
            End If
                
            sDescricao = objDocAuto1.sDescricao
                
            If objDocAuto1.iSeqContraPartida <> 0 Then GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(objDocAuto1.iSeqContraPartida)
            
            'coloca o historico na tela
            GridDocAuto.TextMatrix(iLinha, iGrid_Historico_Col) = objDocAuto1.sHistorico
                
        Next
        
        objGrid1.ilinhasExistentes = colDocAutos.Count
        Descricao.Text = sDescricao
        
        TotalCredito.Caption = Format(dTotalCredito, "Standard")
        TotalDebito.Caption = Format(-dTotalDebito, "Standard")
       
    End If
            
    iAlterado = 0
            
    Traz_Doc_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Doc_Tela:

    Traz_Doc_Tela = Err

    Select Case Err
    
        Case 11014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objDocAuto1.sConta)
        
        Case 11015, 11016
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155674)
        
    End Select
    
    iAlterado = 0
    
    Exit Function
        
End Function

Public Function Trata_Parametros(Optional objDocAuto As ClassDocAuto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um documento(voucher) passado como parametro, exibir seus dados
    If Not (objDocAuto Is Nothing) Then
    
        lErro = Traz_Doc_Tela(objDocAuto)
        If lErro <> SUCESSO And lErro <> 11013 Then Error 11018
    
    Else
    
        Call Limpa_Tela_DocAuto
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
                    
        iAlterado = 0
        
    End If
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 11018
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155675)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Public Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Conta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Credito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Debito_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoDocAuto = Nothing
    
    Set objGrid1 = Nothing
    
End Sub

Public Sub Documento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Documento, iAlterado)

End Sub

Public Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GridDocAuto_LeaveCell()
    
    Call Saida_Celula(objGrid1)

End Sub

Public Sub GridDocAuto_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid1, iAlterado)
    
End Sub

Public Sub GridDocAuto_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If
    
End Sub

Public Sub GridDocAuto_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid1)

End Sub

Public Sub GridDocAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim dColunaSoma As Double
Dim lErro As Long

On Error GoTo Erro_GridLancamentos_KeyDown

    lErro = Grid_Trata_Tecla1(KeyCode, objGrid1)
    If lErro <> SUCESSO Then Error 44067

    Call Trata_SeqContraPartida(GridDocAuto.Row)

    dColunaSoma = GridColuna_Soma(iGrid_Credito_Col)
    TotalCredito = Format(dColunaSoma, "Standard")
    dColunaSoma = GridColuna_Soma(iGrid_Debito_Col)
    TotalDebito = Format(dColunaSoma, "Standard")
    
    Exit Sub
    
Erro_GridLancamentos_KeyDown:

    Select Case Err
    
        Case 44067
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155676)
    
    End Select

    Exit Sub
    
End Sub

Public Sub GridDocAuto_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid1, iAlterado)
    End If

End Sub

Public Sub GridDocAuto_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid1)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then
    
        Select Case GridDocAuto.Col
    
            Case iGrid_Conta_Col
            
                lErro = Saida_Celula_Conta(objGridInt)
                If lErro <> SUCESSO Then Error 11036
                
            Case iGrid_Ccl_Col
            
                lErro = Saida_Celula_Ccl(objGridInt)
                If lErro <> SUCESSO Then Error 11037
                
            Case iGrid_Credito_Col
            
                lErro = Saida_Celula_Credito(objGridInt)
                If lErro <> SUCESSO Then Error 11038
                
            Case iGrid_Debito_Col
            
                lErro = Saida_Celula_Debito(objGridInt)
                If lErro <> SUCESSO Then Error 11039

            Case iGrid_SeqContraPartida_Col
            
                lErro = Saida_Celula_SeqContraPartida(objGridInt)
                If lErro <> SUCESSO Then Error 20628

            Case iGrid_Historico_Col
            
                lErro = Saida_Celula_Historico(objGridInt)
                If lErro <> SUCESSO Then Error 11040
                
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 11041
        
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err
    
    Select Case Err
            
        Case 11036, 11037, 11038, 11039, 11040, 20628
    
        Case 11041
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155677)
        
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
    If lErro <> SUCESSO And lErro <> 5703 Then Error 11042
                
    'se o centro de custo/lucro não estiver cadastrado
    If lErro = 5703 Then Error 11043
                
    If Len(Ccl.ClipText) > 0 Then
    
        If Len(GridDocAuto.TextMatrix(GridDocAuto.Row, iGrid_Conta_Col)) > 0 Then
    
            sConta = GridDocAuto.TextMatrix(GridDocAuto.Row, iGrid_Conta_Col)
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 11044
        
            objContaCcl.sConta = sContaFormatada
            objContaCcl.sCcl = sCclFormatada
        
            
            lErro = CF("ContaCcl_Le", objContaCcl)
            If lErro <> SUCESSO And lErro <> 5871 Then Error 11045
        
            'associação Conta x Centro de Custo/Lucro não cadastrada
            If lErro = 5871 Then Error 11046
            
        End If
                
'        If GridDocAuto.Row - GridDocAuto.FixedRows = objGridInt.ilinhasExistentes Then
'            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
'        End If
                
        CclDescricao.Caption = objCcl.sDescCcl
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11047

    Saida_Celula_Ccl = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err
    
    Select Case Err
    
        Case 11042, 11047, 11044, 11045
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 11043
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)

            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
        
        Case 11046
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", sConta, Ccl.Text)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
            
                objContaCcl.sConta = sContaFormatada
                
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155678)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Conta(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim sContaFormatada As String
Dim sCclFormatada As String
Dim lErro As Long
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
    If lErro <> SUCESSO And lErro <> 44033 And lErro <> 44037 Then Error 20640
    
    'se é uma conta simples, coloca a conta normal no lugar da conta simples
    If lErro = SUCESSO Then
    
        sContaFormatada = objPlanoConta.sConta
        
        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 20641
        
        Conta.PromptInclude = False
        Conta.Text = sContaMascarada
        Conta.PromptInclude = True
        
    End If
    
    'se não encontrou a conta simples
    If lErro = 44033 Or lErro = 44037 Then
    
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", Conta.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 11048
                    
        'conta não cadastrada
        If lErro = 5700 Then Error 11056
                
    End If
    
    If Len(Conta.ClipText) > 0 Then
    
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            If Len(GridDocAuto.TextMatrix(GridDocAuto.Row, iGrid_Ccl_Col)) > 0 Then
            
                sCcl = GridDocAuto.TextMatrix(GridDocAuto.Row, iGrid_Ccl_Col)
        
                lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
                If lErro <> SUCESSO Then Error 11050
        
                objContaCcl.sConta = sContaFormatada
                objContaCcl.sCcl = sCclFormatada
        
                lErro = CF("ContaCcl_Le", objContaCcl)
                If lErro <> SUCESSO And lErro <> 5871 Then Error 11051
        
                'associação Conta x Centro de Custo/Lucro não cadastrada
                If lErro = 5871 Then Error 11055
                
            End If
            
        End If
                
        If GridDocAuto.Row - GridDocAuto.FixedRows = objGridInt.ilinhasExistentes Then
            objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
        End If
                
        ContaDescricao.Caption = objPlanoConta.sDescConta
        
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11049

    Saida_Celula_Conta = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Conta:

    Saida_Celula_Conta = Err
    
    Select Case Err
    
        Case 11048, 11049, 11050, 11051, 20640
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 11055
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTACCL_INEXISTENTE", Conta.Text, sCcl)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("ContaCcl", objContaCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
        
        Case 11056
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", Conta.Text)

            If vbMsgRes = vbYes Then
            
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("PlanoConta", objPlanoConta)
                
            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If
        
        Case 20641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155679)
    
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
        If lErro <> SUCESSO Then Error 11058
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11059
                
    If Len(Debito.Text) > 0 Then
        GridDocAuto.TextMatrix(GridDocAuto.Row, iGrid_Credito_Col) = ""
        If GridDocAuto.Row - GridDocAuto.FixedRows = objGridInt.ilinhasExistentes Then
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
    
        Case 11058, 11059
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155680)
        
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
        If lErro <> SUCESSO Then Error 11060
        
        objHistPadrao.iHistPadrao = CInt(sValor)
                
        lErro = CF("HistPadrao_Le", objHistPadrao)
        If lErro <> SUCESSO And lErro <> 5446 Then Error 11061
        
        If lErro = 5446 Then Error 11063

        Historico.Text = objHistPadrao.sDescHistPadrao

    End If
    
    If Len(Trim(Historico.Text)) > 0 Then
            If GridDocAuto.Row - GridDocAuto.FixedRows = objGridInt.ilinhasExistentes Then
                objGridInt.ilinhasExistentes = objGridInt.ilinhasExistentes + 1
            End If
    End If
      
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11062

    Saida_Celula_Historico = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = Err
    
    Select Case Err
    
        Case 11060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 11061, 11062
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 11063
             vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_HISTPADRAO_INEXISTENTE", objHistPadrao.iHistPadrao)

            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("HistoricoPadrao", objHistPadrao)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155681)
        
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
        If lErro <> SUCESSO Then Error 11064
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 11065
                
    If Len(Credito.Text) > 0 Then
        GridDocAuto.TextMatrix(GridDocAuto.Row, iGrid_Debito_Col) = ""
         If GridDocAuto.Row - GridDocAuto.FixedRows = objGridInt.ilinhasExistentes Then
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
    
        Case 11064, 11065
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155682)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_SeqContraPartida(objGridInt As AdmGrid) As Long
'faz a critica da celula sequencial de contra partida do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_SeqContraPartida

    Set objGridInt.objControle = SeqContraPartida
    If Len(SeqContraPartida.Text) > 0 Then
    
        If GridDocAuto.Row = CInt(SeqContraPartida.Text) Then Error 20629
        
        If CInt(SeqContraPartida.Text) > objGridInt.ilinhasExistentes Then Error 20630
        
        If CInt(SeqContraPartida.Text) <= 0 Then Error 20631
    
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 20632
                
    Saida_Celula_SeqContraPartida = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_SeqContraPartida:

    Saida_Celula_SeqContraPartida = Err
    
    Select Case Err
    
        Case 20629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_NAO_MESMO_LANCAMENTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 20630, 20631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTRAPARTIDA_LANCAMENTO_INEXISTENTE", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 20632
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155683)
        
    End Select

    Exit Function

End Function

Function GridColuna_Soma(iColuna As Integer) As Double
    
Dim dAcumulador As Double
Dim iLinha As Integer
    
    dAcumulador = 0
    
    For iLinha = 1 To objGrid1.ilinhasExistentes
        If Len(GridDocAuto.TextMatrix(iLinha, iColuna)) > 0 Then
            dAcumulador = dAcumulador + CDbl(GridDocAuto.TextMatrix(iLinha, iColuna))
        End If
    Next
    
    GridColuna_Soma = dAcumulador

End Function

Public Sub Label3_Click()

Dim objDocAuto As New ClassDocAuto
Dim colSelecao As Collection

    If Len(Documento.Text) = 0 Then
        objDocAuto.lDoc = 0
    Else
        objDocAuto.lDoc = CLng(Documento.ClipText)
    End If

    objDocAuto.iSeq = 0

    Call Chama_Tela("DocAutoLista", colSelecao, objDocAuto, objEventoDocAuto)
    
End Sub

Private Sub objEventoDocAuto_evSelecao(obj1 As Object)

Dim objDocAuto As ClassDocAuto
Dim lErro As Long
    
On Error GoTo Erro_objEventoDocAuto_evSelecao
    
    Set objDocAuto = obj1
    
    lErro = Traz_Doc_Tela(objDocAuto)
    If lErro <> SUCESSO Then Error 11120
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoDocAuto_evSelecao:

    Select Case Err
    
        Case 11120
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155684)
            
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
    
    If GridDocAuto.Col = iGrid_Ccl_Col Then
    
        sCaracterInicial = left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sCcl = right(Node.Key, Len(Node.Key) - 1)
              
            sCclEnxuta = String(STRING_CCL, 0)
            
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaCclEnxuta(sCcl, sCclEnxuta)
            If lErro <> SUCESSO Then Error 10490
            
            Ccl.PromptInclude = False
            Ccl.Text = sCclEnxuta
            Ccl.PromptInclude = True
              
            GridDocAuto.TextMatrix(GridDocAuto.Row, GridDocAuto.Col) = Ccl.Text
        
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
    
        Case 10490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", Err, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155685)
            
    End Select
        
    Exit Sub

End Sub

Public Sub ListaHistorico_DblClick()

Dim lPosicaoSeparador As Long
    
    If GridDocAuto.Col = iGrid_Historico_Col Then
    
        lPosicaoSeparador = InStr(ListaHistorico.Text, SEPARADOR)
        GridDocAuto.TextMatrix(GridDocAuto.Row, GridDocAuto.Col) = Mid(ListaHistorico.Text, lPosicaoSeparador + 1)
        Historico.Text = Mid(ListaHistorico.Text, lPosicaoSeparador + 1)
    
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
        If lErro <> SUCESSO Then Error 40818
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40818
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155686)
        
    End Select
        
    Exit Sub
    
End Sub

Public Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
    
Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick
    
    If GridDocAuto.Col = iGrid_Conta_Col Then
    
        sCaracterInicial = left(Node.Key, 1)
    
        If sCaracterInicial = "A" Then
    
            sConta = right(Node.Key, Len(Node.Key) - 1)
            
            sContaEnxuta = String(STRING_CONTA, 0)
                        
            'volta mascarado apenas os caracteres preenchidos
            lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
            If lErro <> SUCESSO Then Error 11067
            
            Conta.PromptInclude = False
            Conta.Text = sContaEnxuta
            Conta.PromptInclude = True
        
            GridDocAuto.TextMatrix(GridDocAuto.Row, GridDocAuto.Col) = Conta.Text
        
            If objGrid1.objGrid.Row - objGrid1.objGrid.FixedRows = objGrid1.ilinhasExistentes Then
                objGrid1.ilinhasExistentes = objGrid1.ilinhasExistentes + 1
            End If
        
            'Preenche a Descricao da Conta
            lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
            ContaDescricao.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
            
        End If
        
    End If
        
    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err
    
        Case 11067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155687)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 36797
    
    Call Limpa_Tela_DocAuto

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 36797
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155688)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
            
Dim lErro As Long
Dim lDoc As Long
Dim sDescricao As String
Dim colDocAutos As New Collection
Dim objDocAuto As New ClassDocAuto
Dim iIndice1 As Integer
Dim dSomaDebito As Double
Dim dSomaCredito As Double
Dim dSoma As Double

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Campo Documento esta preenchido
    If Len(Documento.ClipText) = 0 Then Error 11069
        
    'Verifica se pelo menos uma linha do Grid está preenchida
    If objGrid1.ilinhasExistentes = 0 Then Error 11070

    'Move-se os dados da tela para as variáveis
    lDoc = CLng(Documento.ClipText)

    If Len(Trim(Descricao.ClipText)) = 0 Then
       sDescricao = String(STRING_DOCAUTO_DESCRICAO, 0)
    Else
       sDescricao = Descricao.Text
    End If
    
    'Preenche a colDocAutos com as informacoes contidas no Grid
    lErro = Grid_DocAuto(colDocAutos)
    If lErro <> SUCESSO Then Error 11071
           
    lErro = Trata_Alteracao(objDocAuto, lDoc)
    If lErro <> SUCESSO Then Error 32301
                 
    lErro = CF("DocAuto_Grava", lDoc, MODULO_CONTABILIDADE, sDescricao, colDocAutos)
    If lErro <> SUCESSO Then Error 11073
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 11069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
        
        Case 11070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_DOCAUTO_GRAVAR", Err)
        
        Case 11071, 11073, 32301
             
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155689)
            
    End Select
    
    Exit Function
    
End Function

Function Grid_DocAuto(colDocAutos As Collection) As Long

Dim iIndice1 As Integer
Dim objDocAuto As ClassDocAuto
Dim sConta As String
Dim sContaFormatada As String
Dim sCcl As String
Dim sCclFormatada As String
Dim iContaPreenchida As Integer
Dim iCclPreenchida As Integer
Dim dValorCredito As Double
Dim dValorDebito As Double
Dim lErro As Long
Dim colContraPartida As New Collection

On Error GoTo Erro_Grid_DocAuto

    For iIndice1 = 1 To objGrid1.ilinhasExistentes
        
        Set objDocAuto = New ClassDocAuto
            
        objDocAuto.iSeq = iIndice1
  
        sConta = GridDocAuto.TextMatrix(iIndice1, iGrid_Conta_Col)
            
        If Len(Trim(sConta)) = 0 Then Error 11245
        
        lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 11075
            
        If iContaPreenchida = CONTA_VAZIA Then Error 55700
            
        objDocAuto.sConta = sContaFormatada
    
        'Testa para ver se houve crédito ou débito
        If Len(GridDocAuto.TextMatrix(iIndice1, iGrid_Credito_Col)) > 0 Then
            dValorCredito = CDbl(GridDocAuto.TextMatrix(iIndice1, iGrid_Credito_Col))
        Else
            dValorCredito = 0
        End If
            
        If Len(GridDocAuto.TextMatrix(iIndice1, iGrid_Debito_Col)) > 0 Then
            dValorDebito = CDbl(GridDocAuto.TextMatrix(iIndice1, iGrid_Debito_Col))
        Else
            dValorDebito = 0
        End If
                        
        objDocAuto.dValor = dValorCredito - dValorDebito
    
        'armazena o sequencial de contra partida, se estiver preenchido
        If Len(GridDocAuto.TextMatrix(iIndice1, iGrid_SeqContraPartida_Col)) > 0 Then objDocAuto.iSeqContraPartida = CInt(GridDocAuto.TextMatrix(iIndice1, iGrid_SeqContraPartida_Col))
        
        'Armazena Histórico e Ccl
        objDocAuto.sHistorico = GridDocAuto.TextMatrix(iIndice1, iGrid_Historico_Col)
            
        'Se está usando Centro de Custo/Lucro, armazena-o
        If iGrid_Ccl_Col <> 999 Then
                
            sCcl = GridDocAuto.TextMatrix(iIndice1, iGrid_Ccl_Col)
            
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 11076
            
            If iCclPreenchida = CCL_PREENCHIDA Then
                objDocAuto.sCcl = sCclFormatada
            Else
                objDocAuto.sCcl = ""
            End If
            
                
        End If
                
        'Armazena o objeto objDocAuto na coleção coldocautos
        colDocAutos.Add objDocAuto
        
    Next
    
    Grid_DocAuto = SUCESSO

    Exit Function

Erro_Grid_DocAuto:

    Grid_DocAuto = Err

    Select Case Err
    
        Case 11075, 11076
               
        Case 11245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
        
        Case 55700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_GRID_NAO_PREENCHIDA", Err, iIndice1)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155690)
            
    End Select
    
    Exit Function

End Function

Public Sub BotaoExcluir_Click()
'Exclui os lançamentos relativos ao Documento digitado na tela
    
Dim lErro As Long
Dim lDoc As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click
     
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o campo Documento está preenchido
    If Len(Documento.ClipText) = 0 Then Error 11078
     
    'Envia Mensagem pedindo confirmação da Exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_DOCUMENTO")
    
    If vbMsgRes = vbYes Then
    
        lDoc = CLng(Documento.ClipText)
        
        'Exclui todos os lancamentos daquele documento automático
        lErro = CF("DocAuto_Exclui", lDoc)
        If lErro <> SUCESSO Then Error 11079
    
        Call Limpa_Tela_DocAuto
        
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
            
        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
       Case 11078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO", Err)
            Documento.SetFocus
        
       Case 11079
      
       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155691)
        
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim objDocAuto As New ClassDocAuto

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 11083

    Call Limpa_Tela_DocAuto
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
        
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case Err
    
        Case 11083
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155692)
        
    End Select
    
    Exit Sub
    
End Sub

Public Sub Conta_GotFocus()

Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_Conta_GotFocus

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
    TvwContas.Visible = True
    LabelPlanoConta.Visible = True
    TvwCcls.Visible = False
    LabelCCL.Visible = False
    ListaHistorico.Visible = False
    LabelHistorico.Visible = False
    
    sConta = GridDocAuto.TextMatrix(GridDocAuto.Row, GridDocAuto.Col)
    
    If Len(sConta) > 0 Then

        lErro = Conta_Exibe_Descricao(sConta)
        If lErro <> SUCESSO Then Error 11085
        
    Else

        ContaDescricao = ""

    End If
    
    Exit Sub
    
Erro_Conta_GotFocus:

    Select Case Err
    
        Case 11085
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155693)
    
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
    
    TvwContas.Visible = False
    LabelPlanoConta.Visible = False
    TvwCcls.Visible = True
    LabelCCL.Visible = True
    ListaHistorico.Visible = False
    LabelHistorico.Visible = False
    
    sCcl = GridDocAuto.TextMatrix(GridDocAuto.Row, GridDocAuto.Col)
    
    'Coloca descricao de Ccl no panel
    If Len(sCcl) > 0 Then

        lErro = Ccl_Exibe_Descricao(sCcl)
        If lErro <> SUCESSO Then Error 11086

    Else
    
        CclDescricao = ""
        
    End If
    
    Exit Sub
    
Erro_Ccl_GotFocus:

    Select Case Err
    
        Case 11086
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155694)
    
    End Select
    
    Exit Sub
    
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

Public Sub Documento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Historico_GotFocus()
    
Dim iPos As Integer

    Call Grid_Campo_Recebe_Foco(objGrid1)
    
    TvwContas.Visible = False
    LabelPlanoConta.Visible = False
    TvwCcls.Visible = False
    LabelCCL.Visible = False
    ListaHistorico.Visible = True
    LabelHistorico.Visible = True
    
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
        
            If left(Historico.Text, 1) = CARACTER_HISTPADRAO Then
            
                sValor = Trim(Mid(Historico.Text, 2))
                
                lErro = Valor_Inteiro_Critica(sValor)
                If lErro <> SUCESSO Then Error 44093
                
                objHistPadrao.iHistPadrao = CInt(sValor)
                        
                lErro = CF("HistPadrao_Le", objHistPadrao)
                If lErro <> SUCESSO And lErro <> 5446 Then Error 44094
                
                If lErro = 5446 Then Error 44095
        
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
    
        Case 44093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_HISTPADRAO_INVALIDO", Err, sValor)
        
        Case 44094

        Case 44095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, objHistPadrao.iHistPadrao)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155695)
    
    End Select

    Exit Sub
    
End Sub

Public Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid1.objControle = Historico
    lErro = Grid_Campo_Libera_Foco(objGrid1)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridDocAuto_RowColChange()

    Call Grid_RowColChange(objGrid1)
       
End Sub

Public Sub GridDocAuto_Scroll()

    Call Grid_Scroll(objGrid1)
    
End Sub

Private Function Limpa_Tela_DocAuto() As Long

Dim lErro As Long

    Call Grid_Limpa(objGrid1)
    TotalDebito.Caption = ""
    TotalCredito.Caption = ""
    ContaDescricao.Caption = ""
    CclDescricao.Caption = ""
    Descricao.Text = ""
    Documento.Text = ""
    
    Limpa_Tela_DocAuto = SUCESSO
    
End Function

Function Conta_Exibe_Descricao(sConta As String) As Long
'exibe a descrição da conta no campo ContaDescricao. A conta passada como parametro deve estar mascarada

Dim sContaFormatada As String
Dim lErro As Long
Dim iContaPreenchida As Integer
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Conta_Exibe_Descricao

    'Retorna conta formatada como no BD
    lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
    If lErro <> SUCESSO Then Error 11089
    
    If iContaPreenchida = CONTA_PREENCHIDA Then
    
        lErro = CF("Conta_SelecionaUma", sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 6030 Then Error 11090
    
        If lErro = 6030 Then Error 11091
        
        ContaDescricao.Caption = objPlanoConta.sDescConta
    
    End If
    
    Conta_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Conta_Exibe_Descricao:

    Conta_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11089, 11090
            ContaDescricao = ""
            
        Case 11091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
            ContaDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155696)
            
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
    If lErro <> SUCESSO Then Error 11183
    
    objCcl.sCcl = sCclFormatada

    lErro = CF("Ccl_Le", objCcl)
    If lErro <> SUCESSO And lErro <> 5599 Then Error 11184
    
    If lErro = 5599 Then Error 11185
    
    CclDescricao.Caption = objCcl.sDescCcl
    
    Ccl_Exibe_Descricao = SUCESSO
    
    Exit Function

Erro_Ccl_Exibe_Descricao:

    Ccl_Exibe_Descricao = Err
    
    Select Case Err
    
        Case 11183, 11184
            CclDescricao = ""
            
        Case 11185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objCcl.sCcl)
            CclDescricao = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155697)
            
    End Select
        
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objDocAuto As New ClassDocAuto
Dim colDocAutos As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "DocumentosAuto"
    
    If Len(Trim(Documento.ClipText)) > 0 Then
        objDocAuto.lDoc = CLng(Documento.ClipText)
    Else
        objDocAuto.lDoc = 0
    End If
    
    If Len(Trim(Descricao.ClipText)) = 0 Then
       objDocAuto.sDescricao = String(STRING_DOCAUTO_DESCRICAO, 0)
    Else
       objDocAuto.sDescricao = Descricao.Text
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Doc", objDocAuto.lDoc, 0, "Doc"
    colCampoValor.Add "Descricao", objDocAuto.sDescricao, STRING_DOCAUTO_DESCRICAO, "Descricao"

    Exit Sub
    
Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155698)

    End Select
    
    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objDocAuto As New ClassDocAuto

On Error GoTo Erro_Tela_Preenche

    objDocAuto.lDoc = colCampoValor.Item("Doc").vValor

    If objDocAuto.lDoc <> 0 Then
    
        objDocAuto.sDescricao = colCampoValor.Item("Descricao").vValor
       
        lErro = Traz_Doc_Tela(objDocAuto)
        If lErro <> SUCESSO Then Error 14962

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 14962

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155699)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Trata_SeqContraPartida(iLinhaExcluida As Integer)
'altera os indicadores de contra partida quando uma linha é excluida

Dim iLinha As Integer

    For iLinha = 1 To objGrid1.ilinhasExistentes
    
        If Len(GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > 0 Then
    
            If CInt(GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) = iLinhaExcluida Then
                GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = ""
            ElseIf CInt(GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) > iLinhaExcluida Then
                GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col) = CStr(CInt(GridDocAuto.TextMatrix(iLinha, iGrid_SeqContraPartida_Col)) - 1)
                
            End If
        End If
    
    Next

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_DOCUMENTO_AUTOMATICO
    Set Form_Load_Ocx = Me
    Caption = "Documento Automático"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "DocAuto"
    
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
        
        If Me.ActiveControl Is Documento Then
            Call Label3_Click
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

