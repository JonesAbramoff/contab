VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTSaque"
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

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Private Const CONTACORRENTE1 As String = "Conta_Corrente"
Private Const SEQUENCIAL1 As String = "Sequencial"
Private Const VALOR1 As String = "Valor"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const HISTORICO1 As String = "Historico"
Private Const CTACONTABIL_DEBITO As String = "CtaContabil_Deb"
Private Const DOC_EXTERNO As String = "Doc_Externo"
Private Const CTACONTACORRENTE As String = "Cta_Conta_Corrente"
Private Const NUMERO_PAGAMENTO As String = "Numero_Pagamento"
'fim contabilidade

Dim iFrameAtual As Integer
Public iAlterado As Integer
Private WithEvents objEventoSaque As AdmEvento
Attribute objEventoSaque.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoFavorecido As AdmEvento
Attribute objEventoFavorecido.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2

'--- inicio dos properties get dos controles da tela
Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get CheckImprimirRecibo() As Object
     Set CheckImprimirRecibo = objUserControl.Controls("CheckImprimirRecibo")
End Property

'#############################################
'Inserido por Wagner
Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property
'#############################################

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

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
End Property

Public Property Get CTBCclLabel() As Object
     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
End Property

Public Property Get CTBLabel7() As Object
     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
End Property

Public Property Get CTBContaDescricao() As Object
     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
End Property

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
End Property

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get CTBConta() As Object
     Set CTBConta = objUserControl.Controls("CTBConta")
End Property

Public Property Get CTBDebito() As Object
     Set CTBDebito = objUserControl.Controls("CTBDebito")
End Property

Public Property Get CTBCredito() As Object
     Set CTBCredito = objUserControl.Controls("CTBCredito")
End Property

Public Property Get CTBCcl() As Object
     Set CTBCcl = objUserControl.Controls("CTBCcl")
End Property

Public Property Get CTBUpDown() As Object
     Set CTBUpDown = objUserControl.Controls("CTBUpDown")
End Property

Public Property Get CTBDataContabil() As Object
     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
End Property

Public Property Get CTBLote() As Object
     Set CTBLote = objUserControl.Controls("CTBLote")
End Property

Public Property Get CTBDocumento() As Object
     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
End Property

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
End Property

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
End Property

Public Property Get CTBLabel21() As Object
     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
End Property

Public Property Get CTBOrigem() As Object
     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
End Property

Public Property Get CTBLabel14() As Object
     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
End Property

Public Property Get CTBPeriodo() As Object
     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
End Property

Public Property Get CTBExercicio() As Object
     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
End Property

Public Property Get CTBLabel13() As Object
     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
End Property

Public Property Get CTBLabel5() As Object
     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
End Property

Public Property Get CTBLabelHistoricos() As Object
     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
End Property

Public Property Get CTBLabelContas() As Object
     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
End Property

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
End Property

Public Property Get CTBLabelTotais() As Object
     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
End Property

Public Property Get CTBTotalDebito() As Object
     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
End Property

Public Property Get CTBTotalCredito() As Object
     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
End Property

Public Property Get CTBLabel8() As Object
     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
End Property

Public Property Get CTBLabelDoc() As Object
     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
End Property

Public Property Get CTBLabelLote() As Object
     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get NumRefExterna() As Object
     Set NumRefExterna = objUserControl.Controls("NumRefExterna")
End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get ContaContabilLabel() As Object
     Set ContaContabilLabel = objUserControl.Controls("ContaContabilLabel")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

Public Property Get BotaoImprimirIC() As Object
     Set BotaoImprimirIC = objUserControl.Controls("BotaoImprimirIC")
End Property

Public Property Get Favorecido() As Object
     Set Favorecido = objUserControl.Controls("Favorecido")
End Property

Public Property Get TipoMeioPagto() As Object
     Set TipoMeioPagto = objUserControl.Controls("TipoMeioPagto")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get label1() As Object
     Set label1 = objUserControl.Controls("label1")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get CodContaCorrente() As Object
     Set CodContaCorrente = objUserControl.Controls("CodContaCorrente")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get SpinData() As Object
     Set SpinData = objUserControl.Controls("SpinData")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get LabelSequencial() As Object
     Set LabelSequencial = objUserControl.Controls("LabelSequencial")
End Property

Public Property Get LabelCtaCorrente() As Object
     Set LabelCtaCorrente = objUserControl.Controls("LabelCtaCorrente")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get Natureza() As Object
     Set Natureza = objUserControl.Controls("Natureza")
End Property

Public Property Get LabelNatureza() As Object
     Set LabelNatureza = objUserControl.Controls("LabelNatureza")
End Property

Public Property Get LabelNaturezaDesc() As Object
     Set LabelNaturezaDesc = objUserControl.Controls("LabelNaturezaDesc")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoImprimir_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objBanco As New ClassBanco
Dim sLayoutCheque As String
Dim objInfoChequePag As New ClassInfoChequePag
Dim dtDataEmissao As Date
Dim lNumImpressao As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(CodContaCorrente.Text)) = 0 Then gError 87006
    If Len(Trim(Sequencial.ClipText)) = 0 Then gError 87007
    If Len(Trim(Data.ClipText)) = 0 Then gError 87008
    If Len(Trim(Valor.ClipText)) = 0 Then gError 87009
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then gError 87010
    If Len(Trim(Favorecido.Text)) = 0 Then gError 87011
    
    'Retira o código da combo e passa para iCodigo
    iCodigo = Codigo_Extrai(CodContaCorrente.Text)
     
    'Le a Conta Corrente a partir de iCodigo passado como parâmetro
    lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 87012

    'Caso a Conta Corrente não tiver sido encontrada dispara erro
    If lErro = 11807 Then gError 87013

    'Caso a Conta Corrente não for bancária dispara erro
    If objContaCorrenteInt.iCodBanco = 0 Then gError 87014
    
    'Atribui o valor retornado de objContaCorrenteInt.iCodBanco a objBanco.iCodBanco
    objBanco.iCodBanco = objContaCorrenteInt.iCodBanco
          
    'Le o Banco a partir de objBanco.iCodBanco
    lErro = CF("Banco_Le", objBanco)
    If lErro <> SUCESSO And lErro <> 16091 Then gError 87015
        
    'Caso o banco não tiver sido encontrado dispara erro
    If lErro = 16091 Then gError 87016
        
    'Atribui retorno de objBanco.sLayoutCheque a variavel sLayoutCheque
    sLayoutCheque = objBanco.sLayoutCheque
                                                                                               
    'Recolhe os dados do cheque da tela para objInfoChequePag
    Call Move_tela_Cheque(objInfoChequePag, dtDataEmissao)

    'Chama a função que prepara a impressão do cheque
    lErro = CF("PreparaImpressao_Cheque", lNumImpressao, objInfoChequePag)
    If lErro <> SUCESSO Then gError 87017

    'Chama a função responsável pela impressão do cheque
    lErro = ImprimirCheques(lNumImpressao, sLayoutCheque, dtDataEmissao)
    If lErro <> SUCESSO Then gError 87018
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 87012, 87015, 87017, 87018

        Case 87006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)
                                    
        Case 87007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", gErr)
        
        Case 87008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_INFORMADA", gErr)

        Case 87009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)

        Case 87010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 87011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_NAO_PREENCHIDO", gErr)
        
        Case 87013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CORRENTE_NAO_ENCONTRADA", gErr, CodContaCorrente.Text)
        
        Case 87014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", gErr)

        Case 87016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, objBanco.iCodBanco)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158249)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iConta As Integer
Dim lSeq As Long

On Error GoTo Erro_BotaoProxNum_Click

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 57573

    'Pega no itemdata o codigo da conta selecionada
    iConta = Codigo_Extrai(CodContaCorrente.Text)
    
    'Retorna o proximo sequencial disponivel para a conta em questao
    lErro = CF("CtaCorrente_Sequencial_Transacao", iConta, lSeq)
    If lErro <> SUCESSO Then Error 57555

    Sequencial.Text = CStr(lSeq)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57555
        
        Case 57573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158250)
    
    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Saque()

Dim lErro As Long

    Call Limpa_Tela(Me)

    Sequencial.Text = ""
    CodContaCorrente.Text = ""
    TipoMeioPagto.Text = ""
    Favorecido.Text = ""
    Historico.Text = ""
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")

    ContaContabil.PromptInclude = False
    ContaContabil.Text = ""
    ContaContabil.PromptInclude = True
    
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
End Sub

Function Move_Tela_Memoria(objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long
Dim objProjetoInfo As Object
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim sCclFormatada As String, iCclPreenchida As Integer


On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objmovcontacorrente
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
    objMovContaCorrente.dtDataMovimento = CDate(Data.Text)
    objMovContaCorrente.dValor = CDbl(Valor.Text)
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then gError 18235
    objMovContaCorrente.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)
    If Len(Trim(Numero.Text)) > 0 Then objMovContaCorrente.lNumero = CLng(Numero.Text)
    If Len(Trim(Favorecido.Text)) > 0 Then objMovContaCorrente.iFavorecido = Codigo_Extrai(Favorecido.Text)
    objMovContaCorrente.sHistorico = Historico.Text
    objMovContaCorrente.sNumRefExterna = NumRefExterna.Text
    objMovContaCorrente.iFilialEmpresa = giFilialEmpresa
    objMovContaCorrente.iPortador = 0
    objMovContaCorrente.iConciliado = 0
    objMovContaCorrente.iExcluido = MOVCONTACORRENTE_NAO_EXCLUIDO
    objMovContaCorrente.lNumRefInterna = 0
    objMovContaCorrente.sObservacao = Observacao.Text 'Inserido por Wagner

    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_SAQUE)
    If lErro <> SUCESSO Then gError 181761
    
    Set objMovContaCorrente.objProjetoInfo = objProjetoInfo
    '##################################################
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError 122825
    
    objMovContaCorrente.sNatureza = sNaturezaFormatada
    
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError 122825

    If iCclPreenchida = CCL_PREENCHIDA Then
        objMovContaCorrente.sCcl = sCclFormatada
    Else
        objMovContaCorrente.sCcl = ""
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 18235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", gErr)
        
        Case 181761, 122825
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158251)
            
    End Select
    
    Exit Function

End Function

Public Sub CodContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub Favorecido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
    Set objEventoSaque = Nothing
    Set objEventoContaCorrenteInt = Nothing
    Set objEventoFavorecido = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objEventoNatureza = Nothing
    Set objEventoCcl = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

Public Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub LabelCtaCorrente_Click()
'chama browse de conta corrente

Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim colSelecao As New Collection

    If Len(Trim(CodContaCorrente.Text)) > 0 Then objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContasCorrentesInternas, objEventoContaCorrenteInt)

End Sub

Public Sub LabelSequencial_Click()

Dim objMovContaCorrente As New ClassMovContaCorrente
Dim colSelecao As New Collection

    If Len(Trim(CodContaCorrente.Text)) = 0 Then
        Set objMovContaCorrente = Nothing
        Call Chama_Tela("SaqueLista", colSelecao, objMovContaCorrente, objEventoSaque)
    
    Else
        colSelecao.Add Codigo_Extrai(CodContaCorrente.Text)
        
        If Len(Sequencial.Text) = 0 Then
            objMovContaCorrente.lSequencial = 0
        Else
            objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
        End If
        
        Call Chama_Tela("SaqueLista_Conta", colSelecao, objMovContaCorrente, objEventoSaque)
    End If
    
End Sub

Public Sub Numero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub objEventoSaque_evSelecao(obj1 As Object)

Dim objMovContaCorrente As ClassMovContaCorrente
Dim iIndex As Integer
Dim lErro As Long

On Error GoTo Erro_objEventoSaque_evSelecao

    Set objMovContaCorrente = obj1
    
    lErro = Traz_Saque_Tela(objMovContaCorrente)
    If lErro <> SUCESSO Then Error 18116
    
    Sequencial.SetFocus
    
    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoSaque_evSelecao:

    Select Case Err
    
        Case 18116
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158252)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Sequencial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Sequencial, iAlterado)

End Sub

Public Sub Sequencial_Validate(Cancel As Boolean)

On Error GoTo Erro_Sequencial_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Sequencial.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Sequencial.Text) Then Error 55953

        'Verifica se codigo é menor que um
        If CLng(Sequencial.Text) < 1 Then Error 55954

    End If

    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True

    Select Case Err

        Case 55953, 55954
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_INVALIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158253)

    End Select

    Exit Sub
    
End Sub

Public Sub SpinData_DownClick()
Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    Data.SetFocus

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        'Diminui a data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 19575

        Data.PromptInclude = False
        Data.Text = sData
        Data.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_DownClick:

    Select Case Err

        Case 19575

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158254)

    End Select

    Exit Sub

End Sub

Public Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    Data.SetFocus

    'verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text
        
        'Aumenta a data
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 19576

        Data.PromptInclude = False
        Data.Text = sData
        Data.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_UpClick:

    Select Case Err

        Case 19576

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158255)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValor As Double
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos essencias da tela foram preenchidos
    If Len(Trim(CodContaCorrente.Text)) = 0 Then gError 11912

    If Len(Trim(Sequencial.Text)) = 0 Then gError 11913

    If Len(Trim(Data.ClipText)) = 0 Then gError 11914

    If Len(Trim(Valor.Text)) = 0 Then gError 11915

    dValor = CDbl(Valor.Text)

    If dValor <= 0 Then gError 11916

    'Passa os dados da Tela para objMovcontacorrente
    lErro = Move_Tela_Memoria(objMovContaCorrente)
    If lErro <> SUCESSO Then gError 11917

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 20835

    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    '###############################################
    
    'Rotina encarregada de gravar o saque
    lErro = CF("MovCCI_Grava_Saque", objMovContaCorrente, objContabil)
    If lErro <> SUCESSO Then gError 11918
    
    If CheckImprimirRecibo.Value = vbChecked Then
        
        lErro = ImprimirRecibo(objMovContaCorrente)
        If lErro <> SUCESSO Then gError 184450
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 11912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 11913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

        Case 11914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 11915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_SAQUE_NAO_PREENCHIDO", gErr)

        Case 11916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NEGATIVO", gErr)

        Case 11917, 11918, 20835, 181786, 184450

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158256)

    End Select

    Exit Function

End Function

Public Sub TipoMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se há um valor digitado
    If Len(Trim(Valor.Text)) > 0 Then
    
        'Critiva o valor digitado
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 19577
     
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 19577

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158257)

    End Select

    Exit Sub

End Sub

Public Sub Sequencial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoMeioPagto_Click()
    
    iAlterado = REGISTRO_ALTERADO
    Call ValidaBotao_Cheque
    
End Sub

Public Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'verifica se foi preenchido o TipoMeioPagto
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub
    
    If iAlterado = REGISTRO_ALTERADO Then
    
        'Tenta selecionar o TipoMeioPagto com o codigo digitado
        lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 11903

        If lErro = 6730 Then

            objTipoMeioPagto.iTipo = iCodigo
    
            'Pesquisa no BD a existencia do tipo passado por parametro
            lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
            If lErro <> SUCESSO And lErro <> 11909 Then Error 11904
    
            'Se existir nao existir ----> Erro
            If lErro = 11909 Then Error 18044
            
            'se existir coloca na tela
            TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
        
        End If
    Else
        'Tenta selecionar o TipoMeioPagto com o codigo digitado
        lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 40759

        If lErro = 6730 Then

            objTipoMeioPagto.iTipo = iCodigo
    
            'Pesquisa no BD a existencia do tipo passado por parametro
            lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
            If lErro <> SUCESSO And lErro <> 11909 Then Error 40760
    
            'Se existir nao existir ----> Erro
            If lErro = 11909 Then Error 40761
            
            'se existir coloca na tela
            TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
        
        End If
        
        Call ValidaBotao_Cheque
        
        iAlterado = 0
        
    End If
    
    If lErro = 6731 Then Error 11902

    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 11902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, TipoMeioPagto.Text)
        
        Case 11903, 11904, 40759, 40760

        Case 18044, 40761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158258)

    End Select

    Exit Sub

End Sub

Public Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function Trata_Parametros(Optional objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum movimento foi passado por parametro
    If Not (objMovContaCorrente Is Nothing) Then
    
        'Traz os dados do movimento passado por parametro
        lErro = Traz_Saque_Tela(objMovContaCorrente)
        If lErro <> SUCESSO Then Error 11886
    Else
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 11886
              Call Limpa_Tela_Saque
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158259)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Private Function Traz_Saque_Tela(objMovContaCorrente As ClassMovContaCorrente) As Long
'Coloca na Tela os dados do Saque passado como parametro

Dim lErro As Long
Dim iIndice As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim objFavorecidos As New ClassFavorecidos
Dim sNaturezaEnxuta As String, sCclMascarado As String

On Error GoTo Erro_Traz_Saque_Tela

    '############################
    'Inserido por Wagner 04/01/2005
    Call Limpa_Tela_Saque
    '############################

    'Le o movimento passado como parametro
    lErro = CF("MovContaCorrente_Le", objMovContaCorrente)
    If lErro <> SUCESSO And lErro <> 11893 Then gError 11887

    If lErro = 11893 Then gError 18043
   
    If objMovContaCorrente.iExcluido = MOVCONTACORRENTE_EXCLUIDO Then gError 18252
    
    If objMovContaCorrente.iTipo <> MOVCCI_SAQUE Then gError 18066
    
    'passa os dados para a Tela
    Historico.Text = objMovContaCorrente.sHistorico
    NumRefExterna.Text = objMovContaCorrente.sNumRefExterna
    Numero.Text = CStr(objMovContaCorrente.lNumero)
    Data.Text = Format(objMovContaCorrente.dtDataMovimento, "dd/MM/yy")
    Valor.Text = (objMovContaCorrente.dValor)
    
    Observacao.Text = objMovContaCorrente.sObservacao 'Inserido por Wagner
    
    'verifica se a conta corrente existe
    lErro = CF("ContaCorrenteInt_Le", objMovContaCorrente.iCodConta, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 40755
    
    If lErro = 11807 Then gError 40756
    
    For iIndice = 0 To CodContaCorrente.ListCount - 1
        If CodContaCorrente.List(iIndice) = CStr(objMovContaCorrente.iCodConta) & SEPARADOR & objContaCorrenteInt.sNomeReduzido Then
            CodContaCorrente.ListIndex = iIndice
            Exit For
        End If
    Next
            
    Sequencial.Text = objMovContaCorrente.lSequencial
    
    'Verifica se o TiPoMeioPagot existe
    objTipoMeioPagto.iTipo = objMovContaCorrente.iTipoMeioPagto
    
    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then gError 40757
    
    If lErro = 11909 Then gError 40758
    
    TipoMeioPagto.Text = CStr(objMovContaCorrente.iTipoMeioPagto) & SEPARADOR & objTipoMeioPagto.sDescricao
   
    If objMovContaCorrente.iFavorecido <> 0 Then
    
        objFavorecidos.iCodigo = objMovContaCorrente.iFavorecido
    
        'Verifica se o favorecido existe
        lErro = CF("Favorecido_Le", objFavorecidos)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 18065
    
        If lErro = 11807 Then gError 18067
    
        Favorecido.Text = CStr(objFavorecidos.iCodigo) & SEPARADOR & objFavorecidos.sNome
        
    End If
    
    If Len(Trim(objMovContaCorrente.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objMovContaCorrente.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
    
    If Len(Trim(objMovContaCorrente.sCcl)) <> 0 Then
    
        sCclMascarado = String(STRING_CCL, 0)
        
        lErro = Mascara_RetornaCclEnxuta(objMovContaCorrente.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 122833
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    Else
        
        Ccl.PromptInclude = False
        Ccl.Text = ""
        Ccl.PromptInclude = True
    
    End If
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovContaCorrente.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39543
    
    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objMovContaCorrente.objProjetoInfo, objMovContaCorrente.lNumMovto, PRJ_CR_TIPO_SAQUE)
    If lErro <> SUCESSO Then gError 181676
    '###################################################
    
    Traz_Saque_Tela = SUCESSO

    Exit Function

Erro_Traz_Saque_Tela:

    Traz_Saque_Tela = gErr

    Select Case gErr

        Case 11887, 18065, 39543, 40755, 40757, 181676
        
        Case 18043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", gErr)
        
        Case 18066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_SAQUE", gErr, objMovContaCorrente.lSequencial)
        
        Case 18067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_INEXISTENTE", gErr, objMovContaCorrente.iFavorecido)
        
        Case 18252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVCONTACORRENTE_EXCLUIDO", gErr, objMovContaCorrente.iCodConta, objMovContaCorrente.lSequencial)
            
        Case 40756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objMovContaCorrente.iCodConta)
        
        Case 40758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", gErr, objMovContaCorrente.iTipoMeioPagto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158260)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoMeioPagto() As Long
'Carrega na Combo TipoMeioPagto os tipo de meio de pagamento ativos

Dim lErro As Long
Dim colTipoMeioPagto As New Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Le todos os tipo de pagamento
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 11874

    For Each objTipoMeioPagto In colTipoMeioPagto
        
        'Verifica se estao ativos
        If objTipoMeioPagto.iInativo = TIPOMEIOPAGTO_ATIVO Then

            'coloca na combo
            TipoMeioPagto.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
            TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo

        End If

    Next

    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = Err

    Select Case Err

        Case 11874

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158261)

    End Select

    Exit Function

End Function

Private Function Carrega_Favorecidos() As Long
'Carrega os favorecidos ativos na combo de Favorecidos

Dim lErro As Long
Dim objFavorecidos As ClassFavorecidos
Dim colFavorecidos As New Collection

On Error GoTo Erro_Carrega_Favorecidos

    'Le todos os favorecidos
    lErro = CF("Favorecidos_Le_Todos", colFavorecidos)
    If lErro <> SUCESSO Then Error 11880

    For Each objFavorecidos In colFavorecidos
        
        'Verifica se esta ativo
        If objFavorecidos.iInativo = FAVORECIDO_ATIVO Then

            'Insere na combo
            Favorecido.AddItem CStr(objFavorecidos.iCodigo) & SEPARADOR & objFavorecidos.sNome
            Favorecido.ItemData(Favorecido.NewIndex) = objFavorecidos.iCodigo

        End If
    Next

    Carrega_Favorecidos = SUCESSO

    Exit Function

Erro_Carrega_Favorecidos:

    Carrega_Favorecidos = Err

    Select Case Err

        Case 11880

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158262)

    End Select

    Exit Function

End Function

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente

    'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoDescricao)
    If lErro <> SUCESSO Then Error 11881

    For Each objCodigoNome In colCodigoDescricao

        'Insere na combo de contas correntes
        CodContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 11881

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158263)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de historicos com os historicos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'le o Codigo e a descricao de todos os historicos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 11883

    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de historicos
        Historico.AddItem objCodigoNome.sNome
        Historico.ItemData(Historico.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = Err

    Select Case Err

        Case 11883

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158264)

    End Select
    
    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro  As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a contaCorrente foi informada
    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 11994
    
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)

    'Verifica se o sequencial foi informado
    If Len(Trim(Sequencial.Text)) = 0 Then Error 11995
    
    objMovContaCorrente.lSequencial = CLng(Sequencial)
    
    'Pede a confirmacao da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_SAQUE", objMovContaCorrente.iCodConta, objMovContaCorrente.lSequencial)

    If vbMsgRes = vbYes Then

        objMovContaCorrente.iTipo = MOVCCI_SAQUE

        'Chama a rotina de exclusao
        lErro = CF("MovCCI_Exclui", objMovContaCorrente, objContabil)
        If lErro <> SUCESSO Then Error 11996

        Call Limpa_Tela_Saque

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 11994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
        
        Case 11995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", Err)
            
        Case 11996

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158265)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 18019

    Call Limpa_Tela_Saque

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 18019

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158266)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18018

    'Limpa a tela
    Call Limpa_Tela_Saque

    BotaoImprimir.Enabled = False
    BotaoImprimirIC.Enabled = False

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 18018

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158267)

    End Select

    Exit Sub
    
End Sub

Public Sub CodContaCorrente_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodContaCorrente_Validate

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If CodContaCorrente.Text = CodContaCorrente.List(CodContaCorrente.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 11896

    If lErro = 6730 Then
    
        objContaCorrenteInt.iCodigo = iCodigo
        
        'Lê a Conta Corrente
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 11897
    
        If lErro = 11807 Then Error 18045
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43536
        
        End If
        
        CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    
    End If

    If lErro = 6731 Then Error 11895

    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 11895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodContaCorrente.Text)
        
        Case 11896, 11897
        
        Case 18045
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If
        
        Case 43536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodContaCorrente.Text, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158268)

    End Select

    Exit Sub

End Sub

Public Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'verifica se a data final é válida
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 11899
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 11899

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158269)

    End Select

    Exit Sub

End Sub

Public Sub Favorecido_Click()

        iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Favorecido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFavorecido As New ClassFavorecidos
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_Favorecido_Validate

    'verifica se foi preenchido o Favorecido
    If Len(Trim(Favorecido.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox Favorecido
    If Favorecido.Text = Favorecido.List(Favorecido.ListIndex) Then Exit Sub
    
    If iAlterado = REGISTRO_ALTERADO Then
    
        lErro = Combo_Seleciona(Favorecido, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18118
        
        If lErro = 6730 Then
        
            objFavorecido.iCodigo = iCodigo
            
            lErro = CF("Favorecido_Le", objFavorecido)
            If lErro <> SUCESSO And lErro <> 17015 Then Error 18119
            
            If lErro = 17015 Then Error 18120
            
            Favorecido.Text = CStr(objFavorecido.iCodigo) & SEPARADOR & objFavorecido.sNome
    
        End If
    Else
        
        lErro = Combo_Seleciona(Favorecido, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 40762
        
        If lErro = 6730 Then
        
            objFavorecido.iCodigo = iCodigo
            
            lErro = CF("Favorecido_Le", objFavorecido)
            If lErro <> SUCESSO And lErro <> 17015 Then Error 40763
            
            If lErro = 17015 Then Error 40764
            
            Favorecido.Text = CStr(objFavorecido.iCodigo) & SEPARADOR & objFavorecido.sNome
            
        End If
        iAlterado = 0
    End If
    
    If lErro = 6731 Then Error 18117
    
    Exit Sub

Erro_Favorecido_Validate:

    Cancel = True


    Select Case Err

        Case 18117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAVORECIDO_INEXISTENTE1", Err, Favorecido.Text)
        
        Case 18118, 18119, 40762, 40763

        Case 18120, 40764
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FAVORECIDO_INEXISTENTE", objFavorecido.iCodigo)
        
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("Favorecidos", objFavorecido)
            Else
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158270)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Historico_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iTamanho As Integer
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sDescricao As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    'Verifica o tamanho do texto em historico
    iTamanho = Len(Trim(Historico.Text))

    If iTamanho = 0 Then Exit Sub
    
    'Verifica se é maior que o tamanho maximo
    If iTamanho > STRING_HISTORICOMOVCONTA Then Error 11910
    
    'Verifica se o que foi digitado é numerico
    If IsNumeric(Trim(Historico.Text)) Then
    
        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then Error 40709
           
        objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
                
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then Error 40736
    
        If lErro = 15011 Then Error 40743
                
        Historico.Text = objHistMovCta.sDescricao
            
    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err
    
        Case 11910
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
         
        Case 40709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
        
        Case 40736
        
        Case 40743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158271)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Tela_Extrai

    sTabela = "MovimentosContaCorrente"

    If Len(Trim(CodContaCorrente.Text)) > 0 Then
        objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    Else
        objMovContaCorrente.iCodConta = 0
    End If

    If Len(Trim(Sequencial.Text)) > 0 Then
        objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
    Else
        objMovContaCorrente.lSequencial = 0
    End If

    If Len(Trim(Data.ClipText)) > 0 Then objMovContaCorrente.dtDataMovimento = CDate(Data.Text)
    If Len(Trim(Valor.Text)) > 0 Then objMovContaCorrente.dValor = CDbl(Valor.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "CodConta", objMovContaCorrente.iCodConta, 0, "CodConta"
    colCampoValor.Add "Sequencial", objMovContaCorrente.lSequencial, 0, "Sequencial"
    colCampoValor.Add "NumMovto", 0, 0, "NumMovto"
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Tipo", OP_IGUAL, MOVCCI_SAQUE
    colSelecao.Add "Excluido", OP_IGUAL, MOVCONTACORRENTE_NAO_EXCLUIDO

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158272)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Tela_Preenche
    
    objMovContaCorrente.iCodConta = colCampoValor.Item("CodConta").vValor
    objMovContaCorrente.lSequencial = colCampoValor.Item("Sequencial").vValor
    
    If objMovContaCorrente.iCodConta > 0 And objMovContaCorrente.lSequencial > 0 Then
        
        'Carrega objMovContaCorrente com os dados passados em colCampoValor
        objMovContaCorrente.lNumMovto = colCampoValor.Item("NumMovto").vValor
        
        lErro = Traz_Saque_Tela(objMovContaCorrente)
        If lErro <> SUCESSO Then Error 49602
   
        iAlterado = 0
    
    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 49602

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158273)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim sMascaraConta As String
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        Opcao.Visible = False
    End If
    
    Set objEventoContaCorrenteInt = New AdmEvento
    Set objEventoFavorecido = New AdmEvento
    Set objEventoSaque = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoCcl = New AdmEvento

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    'Inicializa propriedade Mask de ContaContabil
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 64391

    ContaContabil.Mask = sMascaraConta
    
    'Inicializa Máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 64391

    Ccl.Mask = sMascaraCcl
    
    'Carrega a combo dos Tipos de meio de Pagamento
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then gError 11873

    'carrega a combo com os favorecidos ativos
    lErro = Carrega_Favorecidos()
    If lErro <> SUCESSO Then gError 11879

    'Carrega a combo com os codigoa e nomes das contass correntes
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then gError 11884

    'carrega a combo de historico
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then gError 11885
    
    iFrameAtual = 1
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_TESOURARIA)
    If lErro <> SUCESSO Then gError 39542
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then gError 122835
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 11873, 11879, 11884, 11885, 39542, 64391, 122835

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158274)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39545

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39546

    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
    
        Case 39545
        
        Case 39546
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
    End Select

    Exit Function

End Function

Public Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumRefExterna_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1

    CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo)

    CodContaCorrente.SetFocus

    Me.Show

    Exit Sub

End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_SAQUE_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_SAQUE_CONTABILIZACAO
                        
        End Select

    End If
    
End Sub

'Inclusão de Código 12/05/00 Raphael
Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_TESOURARIA)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 64370

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 64371

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_TESOURARIA)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 64372

        'Conta não cadastrada
        If lErro = 5700 Then gError 64373

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case gErr

        Case 64370, 64372
    
        Case 64371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 64373
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)
            
            'Pergunta se o usuário deseja cadastrar a conta
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            'Se sim
            If vbMsgRes = vbYes Then
                'Preenche o parâmetro que será passado para a tela
                objPlanoConta.sConta = sContaFormatada
                'Chama a tela de plano de contas
                Call Chama_Tela("PlanoConta", objPlanoConta)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158275)
    
    End Select

    Exit Sub
    
End Sub

Public Sub ContaContabilLabel_Click()
'BROWSE PLANO_CONTA :

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 64374

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaTESLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

        Case 64374
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158276)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaContabil.Text = ""
    Else
        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 64375

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 64375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158277)

    End Select

    Exit Sub

End Sub

'inicio contabilidade
Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click
    
'    If giTipoVersao = VERSAO_LIGHT Then
'        Call objContabil.Contabil_GridContabil_Consulta_Click
'    End If
    
End Sub

Public Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Public Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub

Public Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Public Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Public Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Public Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Public Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)
    
    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Public Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Public Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Public Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Public Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Public Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Public Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Public Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Public Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Public Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Public Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Public Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Public Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Public Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Public Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Public Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Public Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Public Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Public Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case DOC_EXTERNO
            If Len(Trim(NumRefExterna.Text)) > 0 Then
                objMnemonicoValor.colValor.Add NumRefExterna.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case CTACONTACORRENTE
            If Len(CodContaCorrente.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64406
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64407
                
                If objContasCorrentesInternas.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContasCorrentesInternas.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 64444
                    
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case CONTACORRENTE1
            If Len(CodContaCorrente.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64408
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64409
                
                objMnemonicoValor.colValor.Add objContasCorrentesInternas.sNomeReduzido
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case SEQUENCIAL1
            If Len(Sequencial.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Sequencial.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR1
            If Len(Valor.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Valor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
                    
        Case NUMERO_PAGAMENTO
            If Len(Numero.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Numero.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case FORMA1
            If Len(TipoMeioPagto.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                    
        Case CTACONTABIL_DEBITO
            If Len(ContaContabil.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add ContaContabil.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case Else
            gError 39544

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39544
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64406, 64408, 64444
        
        Case 64407, 64409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContasCorrentesInternas.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158278)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_SAQUE_ID
    Set Form_Load_Ocx = Me
    Caption = "Saque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Saque"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)

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
        
        If Me.ActiveControl Is Sequencial Then
            Call LabelSequencial_Click
        ElseIf Me.ActiveControl Is CodContaCorrente Then
            Call LabelCtaCorrente_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        '#########################################
        ElseIf Me.ActiveControl Is Favorecido Then
            Call LabelFavorecido_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        End If
    
    End If
    
End Sub

Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Function Move_tela_Cheque(objInfoChequePag As ClassInfoChequePag, dtDataEmissao As Date) As Long

    'Recolhe os dados do cheque
    objInfoChequePag.sFavorecido = Nome_Extrai(Favorecido.Text)
    objInfoChequePag.dValor = StrParaDbl(Valor.Text)
    objInfoChequePag.lNumRealCheque = StrParaLong(Numero.Text)
    dtDataEmissao = Data.Text
    
End Function

Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)
    
    Nome_Extrai = sString
    
    Exit Function

End Function

Function ImprimirCheques(lNumImpressao As Long, sLayoutCheques As String, dtDataEmissao As Date) As Long
'chama a impressao de cheques

Dim objRelatorio As New AdmRelatorio
Dim sNomeTsk As String
Dim lErro As Long, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ImprimirCheques

    'a cidade deve vir do endereco da filial que está emitindo, se entrar como EMPRESA_TODA pegar da matriz
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 19466
    
    lErro = objRelatorio.ExecutarDireto("Cheques", "", 0, sLayoutCheques, "NIMPRESSAO", CStr(lNumImpressao), "DEMISSAO", CStr(dtDataEmissao), "TCIDADE", objFilialEmpresa.objEndereco.sCidade)
    If lErro <> SUCESSO Then Error 7431

    ImprimirCheques = SUCESSO

    Exit Function

Erro_ImprimirCheques:

    ImprimirCheques = Err

    Select Case Err

        Case 7431, 19466

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158279)

    End Select

    Exit Function

End Function

Public Sub ValidaBotao_Cheque()
'Apenas verifica o conteúdo da combo TipoMeioPagto e
'se a condição for satisfeita habilita o botão Imprimir

Dim iCodigo As Integer

    'Atribui o valor retornado de Codigo_Extrai a variavel iCodigo
    iCodigo = Codigo_Extrai(TipoMeioPagto.Text)
    
    'Verifica se iCodigo é igual a Constante Cheque
    If iCodigo <> Cheque Then
        'Se for diferente desabilita o botão
        BotaoImprimir.Enabled = False
        BotaoImprimirIC.Enabled = False
    Else
        'Se for igual habilita o botão
        BotaoImprimir.Enabled = True
        BotaoImprimirIC.Enabled = True
    End If
    
    Exit Sub

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

'#####################################
'Inserido por Wagner 03/08/2006
Public Function BotaoProjetos_Click() As Long
        Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Function

Public Function LabelProjeto_Click() As Long
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Function

Public Function Projeto_GotFocus() As Long
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Function

Public Function Projeto_Change() As Long
    iAlterado = REGISTRO_ALTERADO
End Function

Public Function Projeto_Validate(Cancel As Boolean) As Long
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR
   
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "GR"
    objItensPRJCR.sDescricao = "Geral"

    objItensPRJCR.dValorOriginal = StrParaDbl(Valor.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_SAQUE
    
    colItensPRJ.Add objItensPRJCR

    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
End Function
'#####################################

Public Sub BotaoImprimirIC_Click()
    Call PortasCOM.BotaoImprimirIC_Click(Codigo_Extrai(CodContaCorrente.Text), 0, StrParaDate(Data.Text), Nome_Extrai(Favorecido.Text), 0, StrParaDbl(Valor.Text))
End Sub

Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
    
    End If

    Call Natureza_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Natureza_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza.ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_PAGAMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Public Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    colSelecao.Add NATUREZA_TIPO_PAGAMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

End Sub

Public Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Public Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Public Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Public Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Public Sub BotaoImprimirRecibo_Click()

Dim lErro As Long, lNumIntRel As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_BotaoImprimirRecibo_Click
    
    'Verifica se os campos essencias da tela foram preenchidos
    If Len(Trim(CodContaCorrente.Text)) = 0 Then gError 11912
    If Len(Trim(Sequencial.Text)) = 0 Then gError 11913

    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
    
    lErro = CF("MovContaCorrente_ObterNumMovto", objMovContaCorrente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184443
    If lErro <> SUCESSO Then gError 184444
    
    lErro = ImprimirRecibo(objMovContaCorrente)
    If lErro <> SUCESSO Then gError 184449
    
    Exit Sub
    
Erro_BotaoImprimirRecibo_Click:
    
    Select Case gErr

        Case 11912
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 11913
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)
    
        Case 184443, 184449
        
        Case 184444
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184445)

    End Select

    Exit Sub
    
End Sub

Private Function ImprimirRecibo(ByVal objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long, lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_ImprimirRecibo

    lErro = CF("RelReciboContabil_Prepara", objMovContaCorrente.lNumMovto, lNumIntRel)
    If lErro <> SUCESSO Then gError 184446
    
    lErro = objRelatorio.ExecutarDireto("Recibos Com Contabilização", "", 1, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 184447
    
    ImprimirRecibo = SUCESSO
    
    Exit Function
    
Erro_ImprimirRecibo:

    ImprimirRecibo = gErr

    Select Case gErr

        Case 184446, 184447
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184448)

    End Select
    
    Exit Function

End Function

Public Sub LabelFavorecido_Click()

Dim objFavorecido As New ClassFavorecidos
Dim colSelecao As New Collection

    If Len(Trim(Favorecido.Text)) > 0 Then
        objFavorecido.iCodigo = StrParaIntErr(Favorecido.Text)
        objFavorecido.sNome = Favorecido.Text
    End If

    Call Chama_Tela("FavorecidosLista", colSelecao, objFavorecido, objEventoFavorecido)
    
End Sub

Private Sub objEventoFavorecido_evSelecao(obj1 As Object)

Dim objFavorecido As ClassFavorecidos

    Set objFavorecido = obj1

    If objFavorecido.iCodigo <> 0 Then
        Favorecido.Text = CStr(objFavorecido.iCodigo) & SEPARADOR & objFavorecido.sNome
    End If

    Me.Show

    Exit Sub

End Sub

Public Sub CclLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
'Preenche Ccl

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 22930

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 22930
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175259)

    End Select

    Exit Sub

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(Ccl.Text)) = 0 Then Exit Sub

    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 31558

    If lErro = 5703 Then gError 31559

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True


    Select Case gErr

        Case 31558

        Case 31559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175260)

    End Select

    Exit Sub

End Sub


